using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

using PCF_Functions;

namespace PCF_Output
{
    public class Output
    {
        static StringBuilder preamble, pipes, fittings, accessories, materials;
        static string outputDir;

        public static void OutputWriter(StringBuilder Preamble, StringBuilder Pipes, StringBuilder Fittings, StringBuilder Accessories, StringBuilder Materials, string OutputDirectory)
        {
            preamble = Preamble; pipes = Pipes; fittings = Fittings; accessories = Accessories; materials = Materials; outputDir = OutputDirectory;

            //// Clear the output file
            System.IO.File.WriteAllBytes(outputDir + "PCF_Export.pcf", new byte[0]);

            //// Write to output file
            using (StreamWriter w = File.AppendText(outputDir + "PCF_Export.pcf"))
            {
                w.Write(preamble);
                w.Write(pipes);
                w.Write(fittings);
                w.Write(accessories);
                w.Write(materials);
                w.Close();
            }
        }
    }
}