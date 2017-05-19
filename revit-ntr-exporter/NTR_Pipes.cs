using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using NTR_Functions;
using PCF_Functions;

using dw = NTR_Functions.DataWriter;

namespace NTR_Exporter
{
    public class NTR_Pipes
    {
        public StringBuilder Export(string pipeLineGroupingKey, HashSet<Element> elements, ConfigurationData conf, Document doc)
        {
            var pipeList = elements;
            var sbPipes = new StringBuilder();
            var key = pipeLineGroupingKey;

            foreach (Element element in pipeList)
            {
                //Process RO
                sbPipes.Append("RO");

                //Process P1, P2, DN
                Pipe pipe = (Pipe)element;
                //Get connector set for the pipes
                ConnectorSet connectorSet = pipe.ConnectorManager.Connectors;
                //Filter out non-end types of connectors
                IList<Connector> connectorEnd = (from Connector connector in connectorSet
                                                 where connector.ConnectorType.ToString().Equals("End")
                                                 select connector).ToList();

                sbPipes.Append(dw.PointCoords("P1", connectorEnd.First()));
                sbPipes.Append(dw.PointCoords("P2", connectorEnd.Last()));
                sbPipes.Append(dw.DnWriter(element));
                sbPipes.Append(dw.ReadParameterFromDataTable(key, conf.Pipelines, "MAT"));
                sbPipes.Append(dw.ReadParameterFromDataTable(key, conf.Pipelines, "LAST"));
                sbPipes.Append(dw.WriteElementId(element, "REF"));
                sbPipes.Append(" LTG=" + key);

                //sbPipes.Append("    UNIQUE-COMPONENT-IDENTIFIER ");
                //sbPipes.Append(element.UniqueId);
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