﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using BuildingCoder;
using MoreLinq;
using NTR_Functions;
using NTR_Output;
using PCF_Functions;
using iv = NTR_Functions.InputVars;

namespace NTR_Exporter
{
    class NTR_Exporter
    {
        StringBuilder outputBuilder = new StringBuilder();
        readonly ConfigurationData conf = new ConfigurationData();

        public NTR_Exporter()
        {
            //Clear data from previous runs
            outputBuilder.Clear();

            outputBuilder.Append(conf._01_GEN);
            outputBuilder.Append(conf._02_AUFT);
            outputBuilder.Append(conf._03_TEXT);
            outputBuilder.Append(conf._04_LAST);
            outputBuilder.Append(conf._05_DN);
            outputBuilder.Append(conf._06_ISO);
        }

        public Result ExportNtr(ExternalCommandData cData)
        {
            // UIApplication uiApp = commandData.Application;
            Document doc = cData.Application.ActiveUIDocument.Document;

            try
            {
                #region Declaration of variables
                // Instance a collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);

                // Define a Filter instance to filter by System Abbreviation
                ElementParameterFilter sysAbbr = Filter.ParameterValueFilterStringEquals(PCF_Functions.InputVars.SysAbbr, PCF_Functions.InputVars.SysAbbrParam);

                // Declare pipeline grouping object
                IEnumerable<IGrouping<string, Element>> pipelineGroups;

                //Declare an object to hold collected elements from collector
                HashSet<Element> colElements = new HashSet<Element>();
                #endregion

                #region Element collectors
                //If user chooses to export a single pipeline get only elements in that pipeline and create grouping.
                //Grouping is necessary even tho theres only one group to be able to process by the same code as the all pipelines case

                //If user chooses to export all pipelines get all elements and create grouping
                if (iv.ExportAllOneFile)
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

                if (iv.ExportAllSepFiles || iv.ExportSpecificPipeLine)
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

                if (iv.ExportSelection)
                {
                    ICollection<ElementId> selection = cData.Application.ActiveUIDocument.Selection.GetElementIds();
                    colElements = selection.Select(s => doc.GetElement(s)).ToHashSet();
                }

                //DiameterLimit filter applied to ALL elements.
                HashSet<Element> elements = (from element in colElements where NTR_Filter.FilterDiameterLimit(element) select element).ToHashSet();

                //Create a grouping of elements based on the Pipeline identifier (System Abbreviation)
                pipelineGroups = from e in elements
                                 group e by e.LookupParameter(PCF_Functions.InputVars.PipelineGroupParameterName).AsString();
                #endregion

                outputBuilder.AppendLine("C Element definitions");

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

                    //TODO: Change the splitting of elements to follow the worksheets
                    //There's now a mismatch of how many worksheets define the elements (ELEMENTS and SUPPORTS) and
                    //the division into fittings and accessories. Would be more concise to follow the excel configuration
                    //by worksheet

                    StringBuilder sbPipes = NTR_Pipes.Export(gp.Key, pipeList, conf, doc);
                    StringBuilder sbFittings = NTR_Fittings.Export(gp.Key, fittingList, conf, doc);
                    StringBuilder sbAccessories = NTR_Accessories.Export(gp.Key, accessoryList, conf, doc);

                    outputBuilder.Append(sbPipes);
                    outputBuilder.Append(sbFittings);
                    outputBuilder.Append(sbAccessories);
                }
                #endregion

                #region Output
                // Output the processed data
                Output output = new Output();
                output.OutputWriter(doc, outputBuilder, iv.OutputDirectoryFilePath);
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


