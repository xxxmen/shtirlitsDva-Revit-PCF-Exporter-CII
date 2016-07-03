using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using PCF_Functions;
using pd = PCF_Functions.ParameterData;

namespace PCF_Exporter
{
    public class PCFExport
    {
        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            try
            {
                #region Declaration of variables
                // Instance a collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                //FilteredElementCollector pipeTypeCollector = new FilteredElementCollector(doc); //Obsolete???

                // Define a Filter instance to filter by System Abbreviation
                ElementParameterFilter sysAbbr = new Filter(InputVars.SysAbbr, InputVars.SysAbbrParam).epf;

                // Declare pipeline grouping object
                IEnumerable<IGrouping<string, Element>> pipelineGroups;

                // Instance a collecting stringbuilder
                StringBuilder sbCollect = new StringBuilder();
                #endregion

                #region Compose preamble
                //Compose preamble
                Composer composer = new Composer();
                
                StringBuilder sbPreamble = composer.PreambleComposer();
                
                //Append preamble
                sbCollect.Append(sbPreamble);
                #endregion

                #region Element collectors
                //If user chooses to export a single pipeline get only elements in that pipeline and create grouping.
                //Grouping is necessary even tho theres only one group to be able to process by the same code as the all pipelines case
                switch (InputVars.ExportAll)
                {
                    case false:
                        //Define a collector with multiple filters to collect PipeFittings OR PipeAccessories OR Pipes + filter by System Abbreviation
                        //System Abbreviation filter also filters FamilySymbols out.
                        collector.WherePasses(
                            new LogicalOrFilter(
                                new List<ElementFilter>
                                {
                                    new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting),
                                    new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory),
                                    new ElementClassFilter(typeof (Pipe))
                                })).WherePasses(sysAbbr);
                        break;

                    //If user chooses to export all pipelines get all elements and create grouping
                    case true:
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
                        break;
                }
                //Create a grouping of elements based on the Pipeline identifier (System Abbreviation)
                pipelineGroups = from e in collector
                                 group e by e.LookupParameter(InputVars.PipelineGroupParameterName).AsString();
                #endregion

                #region Initialize Material Data
                //Set the start number to count the COMPID instances and MAT groups.
                int elementIdentificationNumber = 0;
                int materialGroupIdentifier = 0;

                //Initialize material group numbers on the elements
                IEnumerable<IGrouping<string, Element>> materialGroups = from e in collector group e by e.LookupParameter(pd.PCF_MAT_DESCR).AsString();

                Transaction trans = new Transaction(doc, "Set PCF_ELEM_COMPID and PCF_MAT_ID");
                trans.Start();

                //Access groups
                foreach (IEnumerable<Element> group in materialGroups)
                {
                    materialGroupIdentifier++;
                    //Access parameters
                    foreach (Element element in group)
                    {
                        elementIdentificationNumber++;
                        element.LookupParameter(pd.PCF_ELEM_COMPID).Set(elementIdentificationNumber);
                        element.LookupParameter(pd.PCF_MAT_ID).Set(materialGroupIdentifier);
                    }
                }
                trans.Commit();

                #endregion

                #region Pipeline management
                foreach (IGrouping<string, Element> gp in pipelineGroups)
                {
                    IList<Element> pipeList = (from element in gp
                                   where element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves
                                   select element).ToList();
                    IList<Element> fittingList = (from element in gp
                                   where element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting
                                   select element).ToList();
                    IList<Element> accessoryList = (from element in gp
                                   where element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeAccessory
                                   select element).ToList();
                    
                    StringBuilder sbPipeline = new PCF_Pipeline.PCF_Pipeline_Export().Export(gp.Key, doc);
                    StringBuilder sbPipes = new PCF_Pipes.PCF_Pipes_Export().Export(gp.Key, pipeList, doc);
                    StringBuilder sbFittings = new PCF_Fittings.PCF_Fittings_Export().Export(gp.Key, fittingList, doc);
                    StringBuilder sbAccessories = new PCF_Accessories.PCF_Accessories_Export().Export(gp.Key, accessoryList, doc);

                    sbCollect.Append(sbPipeline); sbCollect.Append(sbPipes); sbCollect.Append(sbFittings); sbCollect.Append(sbAccessories);

                }
                #endregion

                #region Materials
                StringBuilder sbMaterials = composer.MaterialsSection(materialGroups);
                sbCollect.Append(sbMaterials);
                #endregion

                #region Output
                // Output the processed data

                PCF_Output.Output.OutputWriter(doc, sbCollect, InputVars.OutputDirectoryFilePath);
                #endregion

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
}