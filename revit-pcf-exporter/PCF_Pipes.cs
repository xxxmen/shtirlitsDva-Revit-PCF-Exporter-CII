﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using PCF_Functions;
using pdw = PCF_Functions.ParameterDataWriter;
using pdef = PCF_Functions.ParameterDefinition;
using plst = PCF_Functions.ParameterList;

namespace PCF_Pipes
{
    public class PCF_Pipes_Export
    {
        public StringBuilder Export(string pipeLineGroupingKey, HashSet<Element> elements, Document doc)
        {
            var pipeList = elements;
            var sbPipes = new StringBuilder();
            var key = pipeLineGroupingKey;

            foreach (Element element in pipeList)
            {
                sbPipes.AppendLine(element.get_Parameter(new plst().PCF_ELEM_TYPE.Guid).AsString());
                sbPipes.AppendLine("    COMPONENT-IDENTIFIER " + element.get_Parameter(new plst().PCF_ELEM_COMPID.Guid).AsInteger());

                //Write Plant3DIso entries if turned on
                if (InputVars.ExportToPlant3DIso) sbPipes.Append(Composer.Plant3DIsoWriter(element, doc));

                Pipe pipe = (Pipe)element;
                //Get connector set for the pipes
                ConnectorSet connectorSet = pipe.ConnectorManager.Connectors;
                //Filter out non-end types of connectors
                IList<Connector> connectorEnd = (from Connector connector in connectorSet
                                                 where connector.ConnectorType.ToString().Equals("End")
                                                 select connector).ToList();

                sbPipes.Append(EndWriter.WriteEP1(element, connectorEnd.First()));
                sbPipes.Append(EndWriter.WriteEP2(element, connectorEnd.Last()));

                Composer elemParameterComposer = new Composer();
                sbPipes.Append(elemParameterComposer.ElemParameterWriter(element));

                #region CII export

                if (InputVars.ExportToCII) sbPipes.Append(Composer.CIIWriter(doc, key));
                #endregion

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