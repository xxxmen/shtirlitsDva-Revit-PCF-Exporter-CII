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

namespace PCF_Pipes
{
    public static class PCF_Pipes_Export
    {
        static IList<Element> pipeList;
        public static StringBuilder sbPipes;

        public static StringBuilder Export(IList<Element> elements)
        {
            pipeList = elements;
            sbPipes = new StringBuilder();

            foreach (Element element in pipeList)
            {
                sbPipes.Append(element.LookupParameter(InputVars.PCF_ELEM_TYPE).AsString());
                sbPipes.AppendLine();
                sbPipes.Append("    COMPONENT-IDENTIFIER ");
                sbPipes.Append(element.LookupParameter(InputVars.PCF_ELEM_COMPID).AsInteger());
                sbPipes.AppendLine();
               
                Pipe pipe = (Pipe)element;
                //Get connector set for the pipes
                ConnectorSet connectorSet = pipe.ConnectorManager.Connectors;
                //Filter out non-end types of connectors
                IList<Connector> connectorEnd = (from Connector connector in connectorSet 
                                   where connector.ConnectorType.ToString() == "End"
                                   select connector).ToList();

                sbPipes.Append(EndWriter.WriteEP1(element, connectorEnd.First()));
                sbPipes.Append(EndWriter.WriteEP2(element, connectorEnd.Last()));

                sbPipes.Append("    MATERIAL-IDENTIFIER ");
                sbPipes.Append(element.LookupParameter(InputVars.PCF_MAT_ID).AsInteger());
                sbPipes.AppendLine();
                sbPipes.Append("    PIPING-SPEC ");
                sbPipes.Append(InputVars.PIPING_SPEC);
                sbPipes.AppendLine();
                sbPipes.Append("    UNIQUE-COMPONENT-IDENTIFIER ");
                sbPipes.Append(element.UniqueId);
                sbPipes.AppendLine();
            }

            return sbPipes;

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