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
    public class NTR_Dictionary
    {
        public Dictionary<string, Action<Element, Document>> ElementDictionary;

        public NTR_Dictionary()
        {
            ElementDictionary = new Dictionary<string, Action<Element, Document>>
            {
                {"TEE", TEE}
            };
        }

        public static void TEE(Element element, Document doc)
        {

        }
    }
}
