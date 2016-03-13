using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
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
using PCF_Exporter;

namespace PCF_Parameters
{
    //[TransactionAttribute(TransactionMode.Manual)]
    //[RegenerationAttribute(RegenerationOption.Manual)]

    public class PopulateParameters // : IExternalCommand
    {
        //public Result Execute(ExternalCommandData data, ref string msg, ElementSet elements)
        //{
        //    return Result.Succeeded;
        //    //return ExecuteMyCommand(data.Application, ref msg);
        //}

        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg, string path)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            string filename = path;
            StringBuilder sbFeedback = new StringBuilder();

            //Two collectors are made because I couldn't figure out a way to obta
            FilteredElementCollector eCollector = new FilteredElementCollector(doc);
            eCollector.WherePasses(new LogicalOrFilter(new List<ElementFilter>
                    {
                        new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting),
                        new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory),
                    })).OfClass(typeof(FamilyInstance));

            FilteredElementCollector pCollector = new FilteredElementCollector(doc);
            pCollector.OfCategory(BuiltInCategory.OST_PipeCurves).OfClass(typeof(Pipe));
            
            //string filename = InputVars.ExcelFilePath + InputVars.ExcelFileName; //Legacy code
            
            //Reading of excel moved to form class
            //Use ExcelDataReader to import data from the excel to a dataset
            //FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //excelReader.IsFirstRowAsColumnNames = true;
            //DataSet PCF_DATA_SOURCE = excelReader.AsDataSet();
            //DataTable PCF_DATA = PCF_DATA_SOURCE.Tables[InputVars.ExcelSheet];


            //prepare input variables which are initialized when looping the elements
            string eFamilyType = null; string columnName = null;

            //query is using the variables in the loop to query the dataset
            var query = from value in PCF_Exporter_form.DATA_TABLE.AsEnumerable()
                        where value.Field<string>(0) == eFamilyType
                        select value.Field<string>(columnName);

            
            //Debugging
            //StringBuilder sbParameters = new StringBuilder();

            //Loop all elements pipes and fittings and accessories, setting parameters as defined in the dataset
            try
            {
                Transaction trans = new Transaction(doc, "Initialize PCF parameters");
                trans.Start();

                //Reporting the number of different elements initialized
                int pNumber = 0, fNumber = 0, aNumber = 0;

                foreach (Element element in pCollector)
                {
                    //reporting
                    pNumber++;

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
                    //reporting
                    if (string.Equals(element.Category.Name.ToString(),"Pipe Fittings")) fNumber++;
                    if (string.Equals(element.Category.Name.ToString(), "Pipe Accessories")) aNumber++;

                    FamilyInstance fInstance = element as FamilyInstance;
                    eFamilyType = fInstance.Symbol.FamilyName + ": " + element.Name;
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
                sbFeedback.Append(pNumber + " Pipes initialized.\n"+fNumber + " Pipe fittings initialized.\n"+aNumber+" Pipe accessories initialized.");
                Util.InfoMsg(sbFeedback.ToString());
                //excelReader.Close();

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

    //[TransactionAttribute(TransactionMode.Manual)]
    //[RegenerationAttribute(RegenerationOption.Manual)]

    public class CreateParameterBindings //: IExternalCommand
    {
        //public Result Execute(ExternalCommandData data, ref string msg, ElementSet elements)
        //{
        //    return ExecuteMyCommand(data.Application, ref msg);
        //}

        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Application app = doc.Application;
            Autodesk.Revit.Creation.Application ca = app.Create;

            Category pipeCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves);
            Category fittingCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting);
            Category accessoryCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory);

            CategorySet catSet = ca.NewCategorySet();
            catSet.Insert(pipeCat);
            catSet.Insert(fittingCat);
            catSet.Insert(accessoryCat);

            string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string oriFile = app.SharedParametersFilename;
            string tempFile = ExecutingAssemblyPath + "Temp.txt";
            
            int i = 0;
            StringBuilder sbFeedback = new StringBuilder();
            
            //Create parameter bindings
            try
            {
                Transaction trans = new Transaction(doc, "Bind PCF parameters");
                trans.Start();
                foreach (string name in InputVars.parameterAllNames)
                {
                    using (File.Create(tempFile)) { }
                    app.SharedParametersFilename = tempFile;
                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(name, InputVars.parameterTypes[i]);
                    options.GUID = InputVars.ParameterGUID[i];
                    ExternalDefinition def = app.OpenSharedParameterFile().Groups.Create("TemporaryDefinitionGroup").Definitions.
                        Create(options) as ExternalDefinition;
                    i++;

                    BindingMap map = doc.ParameterBindings;
                    Binding binding = app.Create.NewInstanceBinding(catSet);

                    if (map.Contains(def)) sbFeedback.Append("Parameter " + name + " already exists.\n");
                    else
                    {
                        map.Insert(def, binding, InputVars.PCF_BUILTIN_GROUP_NAME);
                        if (map.Contains(def)) sbFeedback.Append("Parameter " + name + " added to project.\n");
                        else sbFeedback.Append("Creation of parameter " + name + " failed for some reason.\n");
                    }
                    File.Delete(tempFile);
                }
                trans.Commit();
                Util.InfoMsg(sbFeedback.ToString());
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException){return Result.Cancelled;}

            catch (Exception ex)
            {
                msg = ex.Message;
                return Result.Failed;
            }

            app.SharedParametersFilename = oriFile;

            return Result.Succeeded;
           
        }
    }

    //[TransactionAttribute(TransactionMode.Manual)]
    //[RegenerationAttribute(RegenerationOption.Manual)]

    public class DeleteParameters //: IExternalCommand
    {
        //public Result Execute(ExternalCommandData data, ref string msg, ElementSet elements)
        //{
        //    return ExecuteMyCommand(data.Application, ref msg);
        //}
        private StringBuilder sbFeedback = new StringBuilder();

        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg)
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
                Util.InfoMsg(sbFeedback.ToString());
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
        public void RemoveSharedParameterBinding(Application app, string name, ParameterType type)
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

            if (def == null) sbFeedback.Append("Parameter " + name + " does not exist.\n");
            else
            {
                map.Remove(def);
                if (map.Contains(def)) sbFeedback.Append("Failed to delete parameter " + name + " for some reason.\n");
                else sbFeedback.Append("Parameter " + name + " deleted.\n");
            }

            //if (def != null) map.Remove(def); //Legacy code
        }

    }
}