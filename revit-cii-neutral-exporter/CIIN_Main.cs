using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MoreLinq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using CIINExporter.BuildingCoder;

using pd = CIINExporter.ParameterData;
using plst = CIINExporter.ParameterList;
using pdef = CIINExporter.ParameterDefinition;

namespace CIINExporter
{
    public class CIINExport
    {
        internal Result ExecuteMyCommand(UIApplication uiApp, ref string msg)
        {
            Document doc = uiApp.ActiveUIDocument.Document;

            try
            {
                #region Declaration of variables
                // Instance a collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                // Define a Filter instance to filter by System Abbreviation
                ElementParameterFilter sysAbbr = Filter.ParameterValueFilterStringEquals(InputVars.SysAbbr, InputVars.SysAbbrParam);

                // Declare pipeline grouping object
                IEnumerable<IGrouping<string, Element>> pipelineGroups;

                //Declare an object to hold collected elements from collector
                HashSet<Element> colElements = new HashSet<Element>();

                // Instance a collecting stringbuilder
                StringBuilder sbCollect = new StringBuilder();
                #endregion

                #region Compose preamble
                //Append VERSION
                sbCollect.Append(Composer.Section_VERSION());
                #endregion

                #region Element collectors
                //If user chooses to export a single pipeline get only elements in that pipeline and create grouping.
                //Grouping is necessary even tho theres only one group to be able to process by the same code as the all pipelines case

                //If user chooses to export all pipelines get all elements and create grouping
                if (InputVars.ExportAllOneFile)
                {
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

                    colElements = collector.ToElements().ToHashSet();

                }
                
                else if (InputVars.ExportAllSepFiles || InputVars.ExportSpecificPipeLine)
                {
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
                    colElements = collector.ToElements().ToHashSet();
                }

                else if (InputVars.ExportSelection)
                {
                    ICollection<ElementId> selection = uiApp.ActiveUIDocument.Selection.GetElementIds();
                    colElements = selection.Select(s => doc.GetElement(s)).ToHashSet();
                }

                try
                {
                    //DiameterLimit filter applied to ALL elements.
                    HashSet<Element> elements = (from element in colElements
                                                 where
                                                 //Diameter limit filter
                                                 new FilterDiameterLimit().FilterDL(element) &&
                                                 ////Filter out elements with empty PCF_ELEM_TYPE field (remember to !negate)
                                                 //!string.IsNullOrEmpty(element.get_Parameter(new plst().PCF_ELEM_TYPE.Guid).AsString()) &&
                                                 //Filter out EXCLUDED elements -> 0 means no checkmark
                                                 element.get_Parameter(new plst().CII_ELEM_EXCL.Guid).AsInteger() == 0
                                                 select element).ToHashSet();

                    //Create a grouping of elements based on the Pipeline identifier (System Abbreviation)
                    pipelineGroups = from e in elements
                                     group e by e.LookupParameter(InputVars.PipelineGroupParameterName).AsString();
                }
                catch (Exception ex)
                {
                    throw new Exception("Filtering in Main threw an exception:\n" + ex.Message +
                        "\nTo fix:\n" +
                        "1. See if parameter CII_ELEM_EXCL exists, if not, rerun parameter import.");
                }
                
                #endregion


                #region Pipeline management
                foreach (IGrouping<string, Element> gp in pipelineGroups)
                {
                    HashSet<Element> pipeList = (from element in gp
                                                 where element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves
                                                 select element).ToHashSet();
                    HashSet<Element> fittingList = (from element in gp
                                                    where element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting
                                                    select element).ToHashSet();
                    HashSet<Element> accessoryList = (from element in gp
                                                      where element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeAccessory
                                                      select element).ToHashSet();

                    //StringBuilder sbPipeline = new PCF_Pipeline.PCF_Pipeline_Export().Export(gp.Key, doc);
                    //StringBuilder sbPipes = new PCF_Pipes.PCF_Pipes_Export().Export(gp.Key, pipeList, doc);
                    //StringBuilder sbFittings = new PCF_Fittings.PCF_Fittings_Export().Export(gp.Key, fittingList, doc);
                    //StringBuilder sbAccessories = new PCF_Accessories.PCF_Accessories_Export().Export(gp.Key, accessoryList, doc);

                    //sbCollect.Append(sbPipeline); sbCollect.Append(sbPipes); sbCollect.Append(sbFittings); sbCollect.Append(sbAccessories);
                }
                #endregion


                #region Output
                // Output the processed data
                Output output = new Output();
                output.OutputWriter(doc, sbCollect, InputVars.OutputDirectoryFilePath);
                #endregion

            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return Result.Succeeded;
        }
    }
}