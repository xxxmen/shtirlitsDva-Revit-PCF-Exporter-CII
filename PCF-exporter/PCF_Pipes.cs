﻿using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using PCF_Functions;
using pd = PCF_Functions.ParameterData;
using pdef = PCF_Functions.ParameterDefinition;

namespace PCF_Pipes
{
    public class PCF_Pipes_Export
    {
        private IList<Element> pipeList;
        private StringBuilder sbPipes;
        private string key;

        public StringBuilder Export(string pipeLineGroupingKey, IList<Element> elements, Document doc)
        {
            pipeList = elements;
            sbPipes = new StringBuilder();
            key = pipeLineGroupingKey;

            foreach (Element element in pipeList)
            {
                sbPipes.Append(element.get_Parameter(new pdef().PCF_ELEM_TYPE.Guid).AsString());
                sbPipes.AppendLine();
                sbPipes.Append("    COMPONENT-IDENTIFIER ");
                sbPipes.Append(element.get_Parameter(new pdef().PCF_ELEM_COMPID.Guid).AsInteger());
                sbPipes.AppendLine();
               
                Pipe pipe = (Pipe)element;
                //Get connector set for the pipes
                ConnectorSet connectorSet = pipe.ConnectorManager.Connectors;
                //Filter out non-end types of connectors
                IList<Connector> connectorEnd = (from Connector connector in connectorSet 
                                   where connector.ConnectorType.ToString().Equals("End")
                                   select connector).ToList();

                sbPipes.Append(EndWriter.WriteEP1(element, connectorEnd.First()));
                sbPipes.Append(EndWriter.WriteEP2(element, connectorEnd.Last()));

                var pQuery = from p in new pdef().ListParametersAll where !string.IsNullOrEmpty(p.Keyword) && string.Equals(p.Domain, "ELEM") select p;

                foreach (pdef p in pQuery)
                {
                    //Check for parameter's storage type (can be Int for select few parameters)
                    int sT = (int)element.get_Parameter(p.Guid).StorageType;

                    if (sT == 1) //Integer
                    {
                        //Check if the parameter contains anything
                        if (string.IsNullOrEmpty(element.get_Parameter(p.Guid).AsInteger().ToString())) continue;
                        sbPipes.Append("    " + p.Keyword + " ");
                        sbPipes.Append(element.get_Parameter(p.Guid).AsInteger());
                    }
                    else if (sT == 3) //String
                    {
                        //Check if the parameter contains anything
                        if (string.IsNullOrEmpty(element.get_Parameter(p.Guid).AsString())) continue;
                        sbPipes.Append("    " + p.Keyword + " ");
                        sbPipes.Append(element.get_Parameter(p.Guid).AsString());
                    }
                    sbPipes.AppendLine();
                }

                #region CII export
                //Handle CII export parameters
                //Instantiate collector
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                //Get the elements
                collector.OfClass(typeof(PipingSystemType));
                //Select correct systemType
                PipingSystemType sQuery = (from PipingSystemType st in collector
                                           where string.Equals(st.Abbreviation, key)
                                           select st).FirstOrDefault();

                var query = from p in new pdef().ListParametersAll
                            where string.Equals(p.Domain, "PIPL") && string.Equals(p.ExportingTo, "CII")
                            select p;

                foreach (pdef p in query.ToList())
                {
                    if (string.IsNullOrEmpty(sQuery.get_Parameter(p.Guid).AsString())) continue;
                    sbPipes.Append("    ");
                    sbPipes.Append(p.Keyword);
                    sbPipes.Append(" ");
                    sbPipes.Append(sQuery.get_Parameter(p.Guid).AsString());
                    sbPipes.AppendLine();
                }
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