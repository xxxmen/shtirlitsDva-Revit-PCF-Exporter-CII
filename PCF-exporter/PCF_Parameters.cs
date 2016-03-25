using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;
using xel = Microsoft.Office.Interop.Excel;

using PCF_Functions;
using BuildingCoder;
using PCF_Exporter;
using pd = PCF_Functions.ParameterData;
using pdef = PCF_Functions.ParameterDefinition;

namespace PCF_Parameters
{
    //[TransactionAttribute(TransactionMode.Manual)]
    //[RegenerationAttribute(RegenerationOption.Manual)]
    public class ExportParameters
    {
        internal Result ExecuteMyCommand(UIApplication uiApp)
        {
            Document doc = uiApp.ActiveUIDocument.Document;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //Define a collector (Pipe OR FamInst) AND (Fitting OR Accessory OR Pipe).
            //This is to eliminate FamilySymbols from collector which would throw an exception later on.
            collector.WherePasses(new LogicalAndFilter(new List<ElementFilter>
                    {new LogicalOrFilter(new List<ElementFilter>
                        {
                            new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting),
                            new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory),
                            new ElementClassFilter(typeof (Pipe))
                        }),
                    new LogicalOrFilter(new List<ElementFilter>
                                {
                                    new ElementClassFilter(typeof(Pipe)),
                                    new ElementClassFilter(typeof(FamilyInstance))
                                })
                        }));

            //Group all elements by their Family and Type
            IEnumerable<IGrouping<string, Element>> elementGroups = from e in collector group e by e.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString();

            xel.Application excel = new xel.Application();
            if (null == excel) { Util.ErrorMsg("Failed to get or start Excel."); return Result.Failed;}
            excel.Visible = true;
            xel.Workbook workbook = excel.Workbooks.Add(Missing.Value);
            xel.Worksheet worksheet;
            worksheet = excel.ActiveSheet as xel.Worksheet;
            worksheet.Name = "PCF Export - elements";
            int i = 1;
            foreach (string s in pd.elemParametersExport)
            {
                worksheet.Cells[1, i] = s;
                i++;
            }
            worksheet.Range["A1", "Z1"].Font.Bold = true;
            worksheet.Columns.ColumnWidth = 20;
            int j = 2;
            foreach (IGrouping<string, Element> gp in elementGroups)
            {
                worksheet.Cells[j, 1] = gp.Key;
                j++;
            }
            
            collector.Dispose();
            return Result.Succeeded;
        }
    }

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
            EnumerableRowCollection<string> query = from value in PCF_Exporter_form.DATA_TABLE.AsEnumerable()
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
                    foreach (string parameterName in pd.parameterNames)
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
                    foreach (string parameterName in pd.parameterNames)
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

        internal Result CreateElementBindings(UIApplication uiApp, ref string msg)
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
            
            StringBuilder sbFeedback = new StringBuilder();

            //IList<pdef> elementParameters = new pdef().ElementParametersAll;
            
            //Create parameter bindings
            try
            {
                Transaction trans = new Transaction(doc, "Bind element PCF parameters");
                trans.Start();
                foreach (pdef parameter in new pdef().ElementParametersAll)
                {
                    using (File.Create(tempFile)) { }
                    app.SharedParametersFilename = tempFile;
                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(parameter.Name, parameter.Type);
                    options.GUID = parameter.Guid;
                    ExternalDefinition def = app.OpenSharedParameterFile().Groups.Create("TemporaryDefinitionGroup").Definitions.Create(options) as ExternalDefinition;

                    BindingMap map = doc.ParameterBindings;
                    Binding binding = app.Create.NewInstanceBinding(catSet);

                    if (map.Contains(def)) sbFeedback.Append("Parameter " + parameter.Name + " already exists.\n");
                    else
                    {
                        map.Insert(def, binding, InputVars.PCF_BUILTIN_GROUP_NAME);
                        if (map.Contains(def)) sbFeedback.Append("Parameter " + parameter.Name + " added to project.\n");
                        else sbFeedback.Append("Creation of parameter " + parameter.Name + " failed for some reason.\n");
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

        internal Result CreatePipelineBindings(UIApplication uiApp, ref string msg)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Application app = doc.Application;
            Autodesk.Revit.Creation.Application ca = app.Create;

            Category pipelineCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipingSystem);

            CategorySet catSet = ca.NewCategorySet();
            catSet.Insert(pipelineCat);

            string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string oriFile = app.SharedParametersFilename;
            string tempFile = ExecutingAssemblyPath + "Temp.txt";

            StringBuilder sbFeedback = new StringBuilder();

            //IList<pdef> pipelineParameters = new pdef().PipelineParametersAll;

            //Create parameter bindings
            try
            {
                Transaction trans = new Transaction(doc, "Bind PCF parameters");
                trans.Start();
                foreach (pdef parameter in new pdef().PipelineParametersAll)
                {
                    using (File.Create(tempFile)) { }
                    app.SharedParametersFilename = tempFile;
                    ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(parameter.Name, parameter.Type);
                    options.GUID = parameter.Guid;
                    ExternalDefinition def = app.OpenSharedParameterFile().Groups.Create("TemporaryDefinitionGroup").Definitions.Create(options) as ExternalDefinition;

                    BindingMap map = doc.ParameterBindings;
                    Binding binding = app.Create.NewTypeBinding(catSet);

                    if (map.Contains(def)) sbFeedback.Append("Parameter " + parameter.Name + " already exists.\n");
                    else
                    {
                        map.Insert(def, binding, InputVars.PCF_BUILTIN_GROUP_NAME);
                        if (map.Contains(def)) sbFeedback.Append("Parameter " + parameter.Name + " added to project.\n");
                        else sbFeedback.Append("Creation of parameter " + parameter.Name + " failed for some reason.\n");
                    }
                    File.Delete(tempFile);
                }
                trans.Commit();
                Util.InfoMsg(sbFeedback.ToString());
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

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

            //List<pdef> parametersAll = new pdef().ElementParametersAll.Concat(new pdef().PipelineParametersAll).ToList();

            //Call the method to delete parameters
            try
            {
                Transaction trans = new Transaction(doc, "Delete PCF parameters");
                trans.Start();
                foreach (pdef parameter in new pdef().ElementParametersAll.Concat(new pdef().PipelineParametersAll).ToList())
                    RemoveSharedParameterBinding(doc.Application, parameter.Name, parameter.Type);
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