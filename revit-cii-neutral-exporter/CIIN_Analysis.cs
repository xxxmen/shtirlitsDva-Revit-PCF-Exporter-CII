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
            //Instantiate the analytical model
            AnalyticModel aModel = new AnalyticModel();

            //Detect open ends
            var openEnds = detectOpenEnds(doc, elements);

            //Start analysis
            startAnalysis(doc, elements, openEnds, aModel);


            foreach (var item in openEnds) PlaceMarker(doc, item.Origin);
        }

        private void startAnalysis(Document doc, HashSet<Element> elements, List<Connector> openEnds, AnalyticModel aModel)
        {
            foreach (Connector openEnd in openEnds)
            {
                AnalyticSequence sequence = new AnalyticSequence();

                Element owner = openEnd.Owner;

                AnalyticElement ae = new AnalyticElement(owner);
                ae.From = new Node(openEnd);
                ae.To = new Node(DetermineNextConnector(doc, openEnd));

                sequence.Sequence.Push(ae);

                //int i = 10;

                //Node from = new Node();
                //from.Number = i;
                //from.Connector = openEnd;

                //Element owner = openEnd.Owner;
                //AnalyticElement aElement = new AnalyticElement();
                //aElement.Element = owner;
                //aElement.From = from;



                aModel.Sequences.Push(sequence);
            }
        }

        Connector DetermineNextConnector(Document doc, Connector c1)
        {
            Element owner = c1.Owner;

            Connector c2 = null;
            if (owner is Pipe pipe)
            {
                c2 = (from Connector c in pipe.ConnectorManager.Connectors //End of the host/dummy pipe
                      where c.Id != c1.Id && (int)c.ConnectorType == 1
                      select c).FirstOrDefault();
            }
            else if (owner is FamilyInstance fi)
            {
                var cons = GetConnectors(owner);
                var c1mep = c1.GetMEPConnectorInfo();
                if (c1mep.IsPrimary) c2 = cons.Secondary;
                else if (c1mep.IsSecondary) c2 = cons.Primary;
                else if (!c1mep.IsPrimary && !c1mep.IsSecondary) c2 = cons.Primary;
            }
            return c2;
        }

        List<Connector> detectOpenEnds(Document doc, HashSet<Element> elements)
        {
            var cons = GetALLConnectorsFromElements(elements);
            List<Connector> singleConnectors = new List<Connector>();
            foreach (Connector c1 in cons) if (!(1 < cons.Count(c => c.IsEqual(c1)))) singleConnectors.Add(c1);
            return singleConnectors;
        }
    }

    class Node
    {
        public Connector Connector { get; set; } = null;
        public int Number { get; set; } = 0;

        public Node(Connector connector) => Connector = connector;
    }

    class AnalyticElement
    {
        public Node From { get; set; } = null;
        public Node To { get; set; } = null;
        public Element Element { get; set; } = null;

        public AnalyticElement(Element element) => Element = element;
    }

    class AnalyticSequence
    {
        public Stack<AnalyticElement> Sequence { get; set; } = null;
    }

    class AnalyticModel
    {
        public Stack<AnalyticSequence> Sequences { get; set; } = new Stack<AnalyticSequence>();
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
