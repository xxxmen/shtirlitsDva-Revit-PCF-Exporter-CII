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

namespace PCF_Pipeline
{
    public static class PCF_Pipeline_Export
    {
        //static IList<Element> pipeList;
        private static StringBuilder sbPipeline;
        private static string key;

        public static StringBuilder Export(string pipeLineGroupingKey)
        {
            key = pipeLineGroupingKey;
            sbPipeline = new StringBuilder();

            sbPipeline.Append("PIPELINE-REFERENCE ");
            sbPipeline.Append(key);
            sbPipeline.AppendLine();

            sbPipeline.Append("    PIPING-SPEC STD");
            sbPipeline.AppendLine();

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