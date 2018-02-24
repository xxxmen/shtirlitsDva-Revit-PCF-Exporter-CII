using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MoreLinq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using CIINExporter.BuildingCoder;

using static CIINExporter.MepUtils;
using static CIINExporter.Debugger;

namespace CIINExporter
{
    class CIIN_Analysis
    {
        Document doc;
        AnalyticModel Model;

        public CIIN_Analysis(Document doc, HashSet<Element> elements)
        {
            Model = new AnalyticModel(elements);
            this.doc = doc;
        }

        public void AnalyzeSystem()
        {
            //Start analysis
            var openEnds = detectOpenEnds();

            Connector From = openEnds.FirstOrDefault();
            openEnds = openEnds.ExceptWhere(c => c.IsEqual(From)).ToList();
            Connector To = null;

            Node FromNode = new Node();
            FromNode.NextCon = From;
            Model.AllNodes.Add(FromNode);

            Node ToNode = new Node();

            Element curElem = null;
            curElem = From.Owner;
            AnalyticElement curAElem = new AnalyticElement(curElem);

            bool continueSequence = true;

            AnalyticSequence curSequence = new AnalyticSequence();

            for (int i = 0; i < Model.AllElements.Count; i++)
            {
                if (!continueSequence)
                {
                    From = openEnds.FirstOrDefault();
                    openEnds = openEnds.ExceptWhere(c => c.IsEqual(From)).ToList();
                    To = null;
                    FromNode = new Node();
                    FromNode.NextCon = From;
                    Model.AllNodes.Add(FromNode);

                    ToNode = new Node();

                    curElem = null;
                    curElem = From.Owner;
                    curAElem = new AnalyticElement(curElem);

                    continueSequence = true;

                    curSequence = new AnalyticSequence();
                }

                switch (curElem)
                {
                    case Pipe pipe:

                        To = (from Connector c in pipe.ConnectorManager.Connectors //End of the host/dummy pipe
                              where c.Id != From.Id && (int)c.ConnectorType == 1
                              select c).FirstOrDefault();

                        ToNode.PreviousCon = To;
                        Model.AllNodes.Add(ToNode);

                        curAElem.From = FromNode;
                        curAElem.To = ToNode;

                        curSequence.Sequence.Add(curAElem);

                        break;
                    case FamilyInstance fi:
                        Cons cons = GetConnectors(fi);
                        int cat = fi.Category.Id.IntegerValue;
                        switch (cat)
                        {
                            case (int)BuiltInCategory.OST_PipeFitting:
                                var mf = fi.MEPModel as MechanicalFitting;
                                var partType = mf.PartType;
                                switch (partType)
                                {
                                    case PartType.Elbow:
                                        //First Analytic Element
                                        XYZ elementLocation = ((LocationPoint)fi.Location).Point;
                                        ToNode.PreviousLoc = elementLocation;
                                        ToNode.NextLoc = elementLocation;
                                        ToNode.IsElbow = true;
                                        Model.AllNodes.Add(ToNode);

                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;
                                        curSequence.Sequence.Add(curAElem);

                                        //Second Analytic Element
                                        //First determine the To connector
                                        To = (from Connector c in GetALLConnectorsFromElements(curElem)
                                              where c.Id != From.Id
                                              select c).FirstOrDefault();

                                        FromNode = ToNode; //Switch to next element
                                        ToNode = new Node();
                                        Model.AllNodes.Add(ToNode);
                                        ToNode.PreviousCon = To;
                                        curAElem = new AnalyticElement(curElem);
                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;
                                        curSequence.Sequence.Add(curAElem);
                                        break;
                                    case PartType.Tee:
                                        break;
                                    case PartType.Transition:
                                        To = (from Connector c in GetALLConnectorsFromElements(curElem)
                                              where c.Id != From.Id
                                              select c).FirstOrDefault();
                                        ToNode.PreviousCon = To;
                                        Model.AllNodes.Add(ToNode);

                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        curSequence.Sequence.Add(curAElem);
                                        break;
                                    case PartType.Cap:
                                        To = (from Connector c in GetALLConnectorsFromElements(curElem)
                                              where c.Id != From.Id
                                              select c).FirstOrDefault();
                                        ToNode.PreviousCon = To;
                                        Model.AllNodes.Add(ToNode);

                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        curSequence.Sequence.Add(curAElem);
                                        break;
                                    case PartType.Union:
                                        //Unions are not implemeted
                                        continue;
                                    case PartType.SpudAdjustable:
                                        throw new NotImplementedException();
                                    default:
                                        continue;
                                }
                                break;
                            case (int)BuiltInCategory.OST_PipeAccessory:
                                if (From.GetMEPConnectorInfo().IsPrimary) To = cons.Secondary;
                                else if (From.GetMEPConnectorInfo().IsSecondary) To = cons.Primary;
                                else throw new Exception("Something went wrong with connectors of element " + curElem.Id.ToString());

                                ToNode.PreviousCon = To;
                                Model.AllNodes.Add(ToNode);

                                curAElem.From = FromNode;
                                curAElem.To = ToNode;

                                curSequence.Sequence.Add(curAElem);
                                break;
                            default:
                                continue;
                        }
                        break;
                    default:
                        continue;
                }

                //Prepare to restart iteration
                Model.AllConnectors = Model.AllConnectors.ExceptWhere(c => c.Owner.Id.IntegerValue == curElem.Id.IntegerValue).ToList();

                From = (from Connector c in Model.AllConnectors
                        where c.IsEqual(To)
                        select c).FirstOrDefault();

                if (From != null)
                {
                    curElem = From.Owner;

                    FromNode = ToNode;
                    FromNode.NextCon = From;
                    Model.AllNodes.Add(FromNode);

                    curAElem = new AnalyticElement(curElem);
                }
                else
                {
                    continueSequence = false;
                    Model.Sequences.Add(curSequence);

                    openEnds = openEnds.ExceptWhere(c => c.IsEqual(To)).ToList();
                }
                ToNode = new Node();

            }
        }

        public void NumberNodes()
        {
            int thCount = 0;
            int thousands = 0;

            foreach (AnalyticSequence ans in Model.Sequences)
            {
                int tens = 0;
                int tensCount = 0;

                foreach (AnalyticElement ae in ans.Sequence)
                {
                    tensCount++;
                    if (tensCount == 1)
                    {
                        ae.From.Number = thousands + 10;
                        ae.To.Number = thousands + 20;
                        tensCount = 2;
                        continue;
                    }
                    tens = tensCount * 10;
                    int nodeNumber = thousands + tens;
                    ae.To.Number = thousands + tens;
                }
                thCount++;
                thousands = thCount * 1000;
            }
        }

        public void PlaceTextNotesAtNodes()
        {
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create TextNotes");
                foreach (Node node in Model.AllNodes)
                {
                    XYZ location = null;
                    if (node.NextCon != null) location = node.NextCon.Origin;
                    else if (node.NextLoc != null) location = node.NextLoc;
                    else if (node.PreviousCon != null) location = node.PreviousCon.Origin;
                    else if (node.PreviousLoc != null) location = node.PreviousLoc;

                    TextNote.Create(doc, doc.ActiveView.Id, location, node.Number.ToString(), new ElementId(361));

                }
                t.Commit();
            }
        }

        List<Connector> detectOpenEnds()
        {
            List<Connector> singleConnectors = new List<Connector>();
            foreach (Connector c1 in Model.AllConnectors) if (!(1 < Model.AllConnectors.Count(c => c.IsEqual(c1)))) singleConnectors.Add(c1);
            return singleConnectors;
        }
    }

    class Node
    {
        public Connector PreviousCon { get; set; } = null;
        public Connector NextCon { get; set; } = null;
        public XYZ PreviousLoc { get; set; } = null;
        public XYZ NextLoc { get; set; } = null;
        public int Number { get; set; } = 0;
        public bool IsElbow { get; set; } = false;
        public bool IsJunction { get; set; } = false;

        //public Node(Connector connector) => ToConnector = connector;
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
        public List<AnalyticElement> Sequence { get; set; } = new List<AnalyticElement>();
    }

    class AnalyticModel
    {
        public List<AnalyticSequence> Sequences { get; } = new List<AnalyticSequence>();
        public List<Node> AllNodes { get; } = new List<Node>();
        public List<AnalyticElement> AllAnalyticElements { get; } = new List<AnalyticElement>();
        public List<Connector> AllConnectors { get; set; }
        public List<Element> AllElements { get; }

        public AnalyticModel(HashSet<Element> elements)
        {
            AllElements = elements.ToList();
            AllConnectors = GetALLConnectorsFromElements(elements).ToList();
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
