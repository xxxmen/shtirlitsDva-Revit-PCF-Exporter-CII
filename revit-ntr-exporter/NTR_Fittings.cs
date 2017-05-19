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
using mu = PCF_Functions.MepUtils;

namespace NTR_Exporter
{
    class NTR_Fittings
    {
        public StringBuilder Export(string key, HashSet<Element> elements, ConfigurationData conf, Document doc)
        {
            var sbFittings = new StringBuilder();

            foreach (Element element in elements)
            {
                //Read the family and type of the element
                string famAndType = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString();

                //Read element kind
                string kind = dw.ReadParameterFromDataTable(famAndType, conf.Elements, "KIND");
                if (kind == null) continue;
                
                //Write element kind
                sbFittings.Append(kind);

                switch (kind)
                {
                    case "TEE":
                        var cons = NTR_Utils.GetConnectors(element);
                        sbFittings.Append(dw.PointCoords("PH1", cons.Primary));
                        sbFittings.Append(dw.PointCoords("PH2", cons.Secondary));
                        sbFittings.Append(dw.PointCoords("PA1", element));
                        sbFittings.Append(dw.PointCoords("PA2", cons.Tertiary));
                        sbFittings.Append(dw.DnWriter("DNH", cons.Primary));
                        sbFittings.Append(dw.DnWriter("DNA", cons.Tertiary));
                        sbFittings.Append(dw.ReadParameterFromDataTable(kind, conf.Elements, "TYP"));
                        sbFittings.Append(dw.ReadParameterFromDataTable(kind, conf.Elements, "NORM"));
                        break;
                    case "RED":

                        break;
                }

                sbFittings.Append(dw.ReadParameterFromDataTable(key, conf.Pipelines, "MAT"));
                sbFittings.Append(dw.ReadParameterFromDataTable(key, conf.Pipelines, "LAST"));
                sbFittings.Append(dw.WriteElementId(element, "REF"));
                sbFittings.Append(" LTG=" + key);
                sbFittings.AppendLine();

            }

            return sbFittings;
        }
    }
}
