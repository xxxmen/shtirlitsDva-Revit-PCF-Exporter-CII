using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

using PCF_Functions;

namespace PCF_Exporter
{
    //[TransactionAttribute(TransactionMode.Manual)]
    //[RegenerationAttribute(RegenerationOption.Manual)]

    public class PCFExport //: IExternalCommand
    {
        //public Result Execute(ExternalCommandData data, ref string msg, ElementSet elements)
        //{
        //    return ExecuteMyCommand(data.Application, ref msg, elements);
        //}

        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            try
            {
                // Instance a collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                // Define a Filter instance to filter by System Abbreviation
                ElementParameterFilter sysAbbr = new Filter(InputVars.SysAbbr, InputVars.SysAbbrParam).epf;

                // Declare pipeline grouping object
                IEnumerable<IGrouping<string, Element>> pipelineGroups;

                //If user chooses to export a single pipeline get only elements in that pipeline and create grouping.
                //Grouping is necessary even tho theres only one group to be able to process by the same code as the all pipelines case
                if (InputVars.ExportAll == false)
                {
                    //Define a collector with multiple filters to collect PipeFittings OR PipeAccessories OR Pipes + filter by System Abbreviation
                    collector.WherePasses(
                        new LogicalOrFilter(
                            new List<ElementFilter>
                            {
                                new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting),
                                new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory),
                                new ElementClassFilter(typeof (Pipe))
                            })).WherePasses(sysAbbr);

                    //Create a grouping of elements based on the Pipeline identifier (System Abbreviation)
                    pipelineGroups = from e in collector
                                     group e by e.LookupParameter(InputVars.PipelineGroupParameterName).AsString();
                }

                //If user chooses to export all pipelines get all elements and create grouping
                if (InputVars.ExportAll == true)
                {
                    //Define a collector with multiple filters to collect PipeFittings OR PipeAccessories OR Pipes + filter by System Abbreviation
                    collector.WherePasses(
                        new LogicalOrFilter(
                            new List<ElementFilter>
                            {
                                new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting),
                                new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory),
                                new ElementClassFilter(typeof (Pipe))
                            }));

                    //Create a grouping of elements based on the Pipeline identifier (System Abbreviation)
                    pipelineGroups = from e in collector
                                     group e by e.LookupParameter(InputVars.PipelineGroupParameterName).AsString();
                }
                else pipelineGroups = null;

                #region Initialize Material Data
                //Set the start number to count the COMPID instances and MAT groups.
                int elementIdentificationNumber = 0;
                int materialGroupIdentifier = 0;

                //Initialize material group numbers on the elements
                IEnumerable<IGrouping<string, Element>> materialGroups = from e in collector group e by e.LookupParameter(InputVars.PCF_MAT_DESCR).AsString();

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
                        element.LookupParameter(InputVars.PCF_ELEM_COMPID).Set(elementIdentificationNumber);
                        element.LookupParameter(InputVars.PCF_MAT_ID).Set(materialGroupIdentifier);
                    }
                }
                trans.Commit();

                #endregion

                foreach (IGrouping<string, Element> gp in pipelineGroups)
                {
                    IList<Element> pipeList = (from element in gp
                                   where element.Category.Equals(BuiltInCategory.OST_PipeCurves)
                                   select element).ToList();
                    IList<Element> fittingList = (from element in gp
                                      where element.Category.Equals(BuiltInCategory.OST_PipeFitting)
                                      select element).ToList();
                    IList<Element> accessoryList = (from element in gp
                                      where element.Category.Equals(BuiltInCategory.OST_PipeAccessory)
                                      select element).ToList();

                    StringBuilder sbPipeline = PCF_Pipeline.PCF_Pipeline_Export.Export(gp.Key);
                    StringBuilder sbPipes = PCF_Pipes.PCF_Pipes_Export.Export(pipeList);
                    StringBuilder sbFittings = PCF_Fittings.PCF_Fittings_Export.Export(fittingList, doc);
                    StringBuilder sbAccessories = PCF_Accessories.PCF_Accessories_Export.Export(accessoryList, doc);

                }

                StringBuilder preamble = Composer.PreambleComposer();

                #region Pipes
                // Continue on to processing individual elements for export
                // Instance a collector
                // FilteredElementCollector collectorPipes = new FilteredElementCollector(doc);

                //Define a collector with multiple filters to collect Pipes + filter by System Abbreviation
                // IList<Element> pipeList = collectorPipes.OfClass(typeof(Pipe)).WherePasses(sysAbbr).ToElements();

                
                #endregion

                #region Fittings
                // Continue on to processing individual elements for export
                // Instance a collector
                //FilteredElementCollector collectorFittings = new FilteredElementCollector(doc);

                ////Define a collector with multiple filters to collect Pipes + filter by System Abbreviation
                //IEnumerable<Element> fittingsList = collectorFittings.OfCategory(BuiltInCategory.OST_PipeFitting).WherePasses(sysAbbr).ToElements();

                //StringBuilder sbFittings = PCF_Fittings.PCF_Fittings_Export.Export(fittingsList, doc);
                #endregion

                #region Accessories
                //// Continue on to processing individual pipe accessories for export
                //// Instance a collector
                //FilteredElementCollector collectorAccessories = new FilteredElementCollector(doc);

                ////Define a collector with multiple filters to collect Pipes + filter by System Abbreviation
                //IEnumerable<Element> accessoriesList = collectorAccessories.OfCategory(BuiltInCategory.OST_PipeAccessory).WherePasses(sysAbbr).ToElements();

                //StringBuilder sbAccessories = PCF_Accessories.PCF_Accessories_Export.Export(accessoriesList, doc);
                #endregion

                #region Materials
                StringBuilder materials = Composer.MaterialsSection(materialGroups);
                #endregion

                #region Output
                // Output the processed data

                PCF_Output.Output.OutputWriter(preamble, sbPipes, sbFittings, sbAccessories, materials, InputVars.OutputDirectoryFilePath);
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