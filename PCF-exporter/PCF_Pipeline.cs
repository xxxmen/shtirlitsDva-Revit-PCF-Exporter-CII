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
using pdef = PCF_Functions.ParameterDefinition;

namespace PCF_Pipeline
{
    public static class PCF_Pipeline_Export
    {
        //static IList<Element> pipeList;
        private static StringBuilder sbPipeline;
        private static string key;

        public static StringBuilder Export(string pipeLineGroupingKey, Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof (PipingSystem));
            
            key = pipeLineGroupingKey;
            sbPipeline = new StringBuilder();

            var currentSys = (from e in collector
                where e.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString() == key
                select e).First();

            var query = from p in new pdef().ListParametersAll
                         where p.Domain == "PIPL"
                         select p;

            sbPipeline.Append("PIPELINE-REFERENCE ");
            sbPipeline.Append(key);
            sbPipeline.AppendLine();

            foreach (pdef p in query.ToList())
            {
                if (string.IsNullOrEmpty(currentSys.get_Parameter(p.Guid).AsString())) continue;
                sbPipeline.Append("    ");
                sbPipeline.Append(p.Keyword);
                sbPipeline.Append(" ");
                sbPipeline.Append(currentSys.get_Parameter(p.Guid).AsString());
                sbPipeline.AppendLine();
            }

            return sbPipeline;

            //Debug
            //// Clear the output file
            //System.IO.File.WriteAllBytes(InputVars.OutputDirectoryFilePath + "Pipes.pcf", new byte[0]);

            //// Write to output file
            //using (StreamWriter w = File.AppendText(InputVars.OutputDirectoryFilePath + "Pipes.pcf"))
            //{
            //    w.Write(sbPipes);
            //    w.Close();
            //}
        }
    }
}