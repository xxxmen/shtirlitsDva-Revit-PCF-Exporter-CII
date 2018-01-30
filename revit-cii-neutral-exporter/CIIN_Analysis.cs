using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MoreLinq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using CIINExporter.BuildingCoder;

using static CIINExporter.MepUtils;
using static CIINExporter.Debugger;

namespace CIINExporter
{
    class CIIN_Analysis
    {
        public void AnalyzeSystem(Document doc, HashSet<Element> elements)
        {
            //Detect open ends
            var openEnds = detectOpenEnds(doc, elements);
            foreach (var item in openEnds)
            {
                PlaceMarker(doc, item.Origin);
            }
        }

        HashSet<Connector> detectOpenEnds(Document doc, HashSet<Element> elements)
        {
            var cons = GetALLConnectorsFromElements(elements);
            HashSet<Connector> singleConnectors = new HashSet<Connector>();
            foreach (Connector c1 in cons) if (!(1 < cons.Count(c => c.IsEqual(c1)))) singleConnectors.Add(c1);
            return singleConnectors;
        }
    }

    public static class Debugger
    {
        public static void PlaceMarker(Document doc, XYZ loc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var family = collector.OfClass(typeof(Family)).Where(e => e.Name == "Marker").Cast<Family>().FirstOrDefault();
            if (family == null) throw new Exception("No Marker family in project!");
            var famSymbolId = family.GetFamilySymbolIds().FirstOrDefault();
            if (famSymbolId == null) throw new Exception("Getting Marker familySymbol ID failed!");
            var famSymbol = (FamilySymbol)doc.GetElement(famSymbolId);
            if (famSymbol == null) throw new Exception("Getting Marker familySymbol failed!");

            

            using (Transaction trans = new Transaction(doc, "Place markers!"))
            {
                trans.Start();

                //The strange symbol activation thingie...
                //See: http://thebuildingcoder.typepad.com/blog/2014/08/activate-your-family-symbol-before-using-it.html
                if (!famSymbol.IsActive)
                {
                    famSymbol.Activate();
                    doc.Regenerate();
                }

                doc.Create.NewFamilyInstance(loc, famSymbol, StructuralType.NonStructural);

                trans.Commit();
            }
        }
    }
}
