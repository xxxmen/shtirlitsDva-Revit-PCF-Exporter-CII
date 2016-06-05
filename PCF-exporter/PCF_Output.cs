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
        private static StringBuilder _collect;
        private static string _outputDir;
        private static Document _document;

        public static void OutputWriter(Document doc, StringBuilder collect, string outputDirectory)
        {
            _collect = collect; _outputDir = outputDirectory;
            _document = doc;

            string docName = _document.ProjectInformation.Name;
            string dateAndTime = DateTime.Now.ToString();
            dateAndTime = dateAndTime.Replace(" ", "_");
            dateAndTime = dateAndTime.Replace(":", "-");
            //string filename = _outputDir+"\\" + docName + "_" + dateAndTime + ".pcf";
            string filename = _outputDir+"\\" + docName + ".pcf";

            //Clear the output file
            System.IO.File.WriteAllBytes(filename, new byte[0]);

            // Write to output file
            using (StreamWriter w = File.AppendText(filename))
            {
                w.Write(_collect);
                w.Close();
            }
        }
    }
}