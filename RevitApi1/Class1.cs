using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RevitApi1

{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        
        static AddInId addinId = new AddInId(new Guid("CD760254-A75C-48B1-9B14-786815D7F9F4"));

        static string pattern = @"\d+(-\d+)?";

        static Autodesk.Revit.DB.Document doc;

        static UIDocument uidoc;

        static string folderPath = "C:\\00_Проекты\\31_VSH4_PNR\\03_Рабочие файлы\\01_АР\\00_Обмен\\Планы поэтажные";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            //try
            //{
            //    string name = "";
            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("DWG exporting");
            //        ICollection<ElementId> selectedids = uidoc.Selection.GetElementIds();
            //        foreach (ElementId e in selectedids)
            //        {
            //            ElementId e = new ElementId(356);
            //            TaskDialog.Show("id", $"{e.IntegerValue}");
            //            setting up the dwg export options
            //            DWGExportOptions options = new DWGExportOptions();
            //            ExportDWGSettings dWGSettings = ExportDWGSettings.Create(doc, "VA export");
            //            options = dWGSettings.GetDWGExportOptions();
            //            options.Colors = ExportColorMode.TrueColorPerView;
            //            options.FileVersion = ACADVersion.R2010;
            //            options.MergedViews = true;
            //            Element f = doc.GetElement(e);
            //            ElementType ftype = doc.GetElement(f.GetTypeId()) as ElementType;
            //            List<ElementId> icollection = new List<ElementId>();
            //            icollection.Add(e);
            //            name = ftype.FamilyName + " " + f.Name;

            //            doc.Export(folderPath, name, icollection, options);
            //            ElementId dwgsettingid = dWGSettings.Id;
            //            doc.Delete(dwgsettingid);
            //        }
            //        tx.Commit();
            //    }
            //    TaskDialog.Show("Good result", $"Exported {name}");
            //    return Result.Succeeded;
            //}
            //catch (Exception e)
            //{
            //    TaskDialog.Show("Error", "Something went wrong, please contact your BIM manager");
            //    return Result.Failed;
            //}

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ImportInstance));
            List<ImportInstance> importInstances = collector.Cast<ImportInstance>().ToList();

            var instanceTable = ExtractNumberFromTitle(importInstances, pattern);

            int savedElemsCount = SaveInstances(instanceTable);
            if (savedElemsCount != 0)
            {
                TaskDialog.Show("Выгрузка DWG подложек",
                    $"Удачно выгружено {savedElemsCount} подложек," +
                    $"\nне удалось выгрузить {instanceTable.Count - savedElemsCount}");
                return Result.Succeeded;
            }
            else
            {
                TaskDialog.Show("Выгрузка DWG подложек",
                    $"Не удалось выгрузить подложки");
                return Result.Failed;

            }


        }

       
        static int SaveInstances(Dictionary<string, ImportInstance> instanceTable)
        {
            // В этой переменной сохраняем количество успешно сохраненных DWG
            int savedElemsCount = 0;

            foreach (var curElem in instanceTable)
            {
                Transaction trans = new Transaction(doc);
                trans.Start("Save DWG");

                try
                {
                    string fullFolderPath = Path.Combine(folderPath, curElem.Key);

                    if (!Directory.Exists(fullFolderPath))
                    {
                        Directory.CreateDirectory(fullFolderPath);
                    }

                    try
                    {
                        DWGExportOptions options = new DWGExportOptions
                        {
                            Colors = ExportColorMode.TrueColorPerView,
                            FileVersion = ACADVersion.R2010,
                            MergedViews = true
                        };
                        //изолируем двгшку
                        //ICollection<ElementId> toBeIsolated = new HashSet<ElementId> {
                        //    curElem.Value.Id
                        //};
                        //doc.ActiveView.IsolateElementsTemporary(toBeIsolated);


                        //добавляем viewplan(план этажа) с двгшкой в список
                        List<ElementId> elementsToExport = new List<ElementId>();
                        Level level = doc.GetElement(curElem.Value.LevelId) as Level;
                        ElementId viewPlanId = level.FindAssociatedPlanViewId();
                        elementsToExport.Add(viewPlanId);

                        //collector.OfClass(typeof(ImportInstance));
                        //List<ImportInstance> importInstances = collector.Cast<ImportInstance>().ToList();

                        ElementLevelFilter level1Filter = new ElementLevelFilter(level.Id);
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        List<Element> allElementsOnLevel = collector
                            
                            .WherePasses(level1Filter).Cast<Element>().ToList();

                        // Создайте список для элементов, которые нужно скрыть
                        List<ElementId> elementsToHide = new List<ElementId>();

                        foreach (Element element in allElementsOnLevel)
                        {
                            // Добавьте все элементы, кроме ImportInstance, в список элементов для скрытия
                            if (element.Id != curElem.Value.Id)
                            {
                                elementsToHide.Add(element.Id);
                            }
                        }

                        // Скройте все элементы на уровне, кроме ImportInstance
                        doc.ActiveView.HideElements(elementsToHide);

                        // Экспортируем DWG из ImportInstance в указанный путь с заданными параметрами
                        doc.Export(fullFolderPath, curElem.Value.Category.Name, elementsToExport, options);

                        doc.ActiveView.UnhideElements(elementsToHide);

                        //doc.ActiveView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);

                        // Фиксируем транзакцию, сохраняя изменения
                        trans.Commit();

                        savedElemsCount++; // Увеличиваем счетчик успешно сохраненных DWG
                    }
                    catch (Exception ex)
                    {
                        // Если произошла ошибка, откатываем транзакцию и выводим сообщение об ошибке
                        trans.RollBack();
                        TaskDialog.Show("Ошибка выгрузки файла", ex.Message + " " + ex.StackTrace + " " + ex.GetType());
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка создания папки", ex.Message);
                }
            }
            return savedElemsCount;
        }




        static Dictionary<string, ImportInstance> ExtractNumberFromTitle(List<ImportInstance> importInstances, string pattern)
        {
            var instanceTable = new Dictionary<string, ImportInstance>();
            foreach (ImportInstance instance in importInstances)
            {
                Match match = Regex.Match(instance.Category.Name, pattern);
                if (match.Success)
                {
                    //складываем строки с символом С и строку со значением этажа
                    instanceTable["C" + match.Value] = instance;
                }
                else
                {

                    continue;
                }
            }
            return instanceTable;
        }
    }

}
