using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic;
using System.Text;
using System.Data;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;

using Excel;

using PCF_Functions;
using BuildingCoder;

namespace PCF_Parameters
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class PopulateParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string msg, ElementSet elements)
        {
            return ExecuteMyCommand(data.Application, ref msg, elements);
        }

        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg, ElementSet elements)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Application app = doc.Application;

            //Two collectors are made because I couldn't figure out a way to obta
            FilteredElementCollector eCollector = new FilteredElementCollector(doc);
            eCollector.WherePasses(new LogicalOrFilter(new List<ElementFilter>
                    {
                        new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting),
                        new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory),
                    })).OfClass(typeof(FamilyInstance));

            FilteredElementCollector pCollector = new FilteredElementCollector(doc);
            pCollector.OfCategory(BuiltInCategory.OST_PipeCurves).OfClass(typeof(Pipe));
            
            string filename = InputVars.ExcelFilePath + InputVars.ExcelFileName;
            
            //Use ExcelDataReader to import data from the excel to a dataset
            FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet PCF_DATA_SOURCE = excelReader.AsDataSet();
            DataTable PCF_DATA = PCF_DATA_SOURCE.Tables[InputVars.ExcelSheet];

            //prepare input variables which are initialized when looping the elements
            string eFamilyType = null; string columnName = null;

            //query is using the variables in the loop to query the dataset
            var query = from value in PCF_DATA.AsEnumerable()
                        where value.Field<string>(0) == eFamilyType
                        select value.Field<string>(columnName);

            //Debugging
            //StringBuilder sbParameters = new StringBuilder();

            //Loop all elements pipes and fittings and accessories, setting parameters as defined in the dataset
            try
            {
                Transaction trans = new Transaction(doc, "Initialize PCF parameters");
                trans.Start();

                foreach (Element element in pCollector)
                {
                    eFamilyType = "Pipe Types: " + element.Name;
                    foreach (string parameterName in InputVars.parameterNames)
                    {
                        columnName = parameterName;
                        string parameterValue = query.First();
                        element.LookupParameter(parameterName).Set(parameterValue);
                    }

                    //sbParameters.Append(eFamilyType);
                    //sbParameters.AppendLine();
                }

                foreach (Element element in eCollector)
                {
                    FamilyInstance fInstance = element as FamilyInstance;
                    eFamilyType = fInstance.Symbol.FamilyName.ToString() + ": " + element.Name.ToString();
                    foreach (string parameterName in InputVars.parameterNames)
                    {
                        columnName = parameterName;
                        string parameterValue = query.First();
                        element.LookupParameter(parameterName).Set(parameterValue);
                    }

                    //sbParameters.Append(eFamilyType);
                    //sbParameters.AppendLine();
                }
                trans.Commit();
                excelReader.Close();

                //// Debugging
                //// Clear the output file
                //File.WriteAllBytes(InputVars.OutputDirectoryFilePath + "Parameters.pcf", new byte[0]);

                //// Write to output file
                //using (StreamWriter w = File.AppendText(InputVars.OutputDirectoryFilePath + "Parameters.pcf"))
                //{
                //    w.Write(sbParameters);
                //    w.Close();
                //}
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                msg = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;

        }
    }

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class CreateParameterBindings : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string msg, ElementSet elements)
        {
            return ExecuteMyCommand(data.Application, ref msg, elements);
        }

        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg, ElementSet elements)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Application app = doc.Application;
            Autodesk.Revit.Creation.Application ca = app.Create;
            string filename = app.SharedParametersFilename;

            Category pipeCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves);
            Category fittingCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting);
            Category accessoryCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory);

            CategorySet catSet = ca.NewCategorySet();
            catSet.Insert(pipeCat);
            catSet.Insert(fittingCat);
            catSet.Insert(accessoryCat);

            DefinitionFile file = app.OpenSharedParameterFile();
            if(null == file) Util.ErrorMsg("Error getting the shared prameters file.");

            DefinitionGroup group = file.Groups.get_Item(InputVars.PCF_GROUP_NAME);
            if (null == group) Util.ErrorMsg("Error getting the shared parameters group.");
            
            //Create parameter bindings
            try
            {
                Transaction trans = new Transaction(doc, "Bind PCF parameters");
                trans.Start();
                foreach (string name in InputVars.parameterAllNames)
                {
                    IEnumerable<ExternalDefinition> v = from ExternalDefinition d in @group.Definitions where d.Name == name select d;
                    if (v == null || v.Count() < 1) throw new Exception("Invalid Name Input!");
                    ExternalDefinition def = v.First();
                    Binding binding = ca.NewInstanceBinding(catSet);
                    BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
                    map.Insert(def, binding, InputVars.PCF_BUILTIN_GROUP_NAME);
                }
                trans.Commit();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                msg = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;

        }
    }

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class DeleteParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string msg, ElementSet elements)
        {
            return ExecuteMyCommand(data.Application, ref msg, elements);
        }

        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg, ElementSet elements)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            //Call the method to delete parameters
            try
            {
                Transaction trans = new Transaction(doc, "Delete PCF parameters");
                trans.Start();
                int i = 0;
                foreach (string name in InputVars.parameterAllNames)
                {
                    RemoveSharedParameterBinding(doc.Application, name, InputVars.parameterTypes[i]);
                    i++;
                }
                trans.Commit();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                msg = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;

        }

        //Method deletes parameters
        public static bool RemoveSharedParameterBinding(Application app, string name, ParameterType type)
        {
            BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();

            Definition def = null;
            while (it.MoveNext())
            {
                if (it.Key != null && it.Key.Name == name && type == it.Key.ParameterType)
                {
                    def = it.Key;
                    break;
                }
            }

            if (def != null)
            {
                map.Remove(def);
            }

            return false;
        }

    }
}