using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using NTR_Functions;
using PCF_Functions;

using dw = NTR_Functions.DataWriter;

namespace NTR_Exporter
{
    class NTR_Fittings
    {
        public StringBuilder Export(string key, HashSet<Element> elements, ConfigurationData conf, Document doc)
        {
            var sbFittings = new StringBuilder();
            NTR_Dictionary dict = new NTR_Dictionary();

            foreach (Element element in elements)
            {
                //Read the family and type of the element
                string famAndType = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString();

                //Read element kind
                string kind = dw.ReadParameterFromDataTable(famAndType, conf.Elements, "KIND");

                

                //Write element kind
                sbFittings.Append(kind);

                //Process P1, P2, DN
                Pipe pipe = (Pipe)element;
                //Get connector set for the pipes
                ConnectorSet connectorSet = pipe.ConnectorManager.Connectors;
                //Filter out non-end types of connectors
                IList<Connector> connectorEnd = (from Connector connector in connectorSet
                                                 where connector.ConnectorType.ToString().Equals("End")
                                                 select connector).ToList();

                sbFittings.Append(dw.PointCoords("P1", connectorEnd.First()));
                sbFittings.Append(dw.PointCoords("P2", connectorEnd.Last()));
                sbFittings.Append(dw.DnWriter(element));
                sbFittings.Append(dw.ReadParameterFromDataTable(key, conf.Pipelines, "MAT"));
                sbFittings.Append(dw.ReadParameterFromDataTable(key, conf.Pipelines, "LAST"));
                sbFittings.Append(dw.WriteElementId(element, "REF"));
                sbFittings.Append(" LTG=" + key);

                //sbPipes.Append("    UNIQUE-COMPONENT-IDENTIFIER ");
                //sbPipes.Append(element.UniqueId);
                sbFittings.AppendLine();

            }

            return sbFittings;
        }
    }
}
