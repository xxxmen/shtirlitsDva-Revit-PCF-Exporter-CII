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

            //foreach (var item in openEnds) PlaceMarker(doc, item.Origin);

            Connector From = openEnds.FirstOrDefault();
            openEnds.Remove(From);
            Connector To = null;

            Node FromNode = new Node();
            FromNode.NextCon = From;
            Model.AllNodes.Add(FromNode);

            Node ToNode = new Node();

            Element curElem = null;
            curElem = From.Owner;
            AnalyticElement curAElem = null;

            bool continueSequence = true;

            AnalyticSequence curSequence = new AnalyticSequence();

            for (int i = 0; i < Model.AllElements.Count; i++)
            {
                switch (curElem)
                {
                    case Pipe pipe:

                        To = (from Connector c in pipe.ConnectorManager.Connectors //End of the host/dummy pipe
                              where c.Id != From.Id && (int)c.ConnectorType == 1
                              select c).FirstOrDefault();

                        ToNode.PreviousCon = To;

                        curAElem = new AnalyticElement(curElem);
                        curAElem.From = FromNode;
                        curAElem.To = ToNode;

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

                                        curAElem = new AnalyticElement(curElem);
                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;
                                        curSequence.Sequence.Add(curAElem);

                                        //Second Analytic Element
                                        //First determine the To connector
                                        To = (from Connector c in GetALLConnectorsFromElements(curElem)
                                              where c.Id != From.Id
                                              select c).FirstOrDefault();


                                        FromNode = ToNode; //Switch to next element
                                        ToNode = new Node(); //Added to allnodes later, see "ToNode" added to AllNodes
                                        ToNode.PreviousCon = To;
                                        curAElem = new AnalyticElement(curElem);
                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;
                                        curSequence.Sequence.Add(curAElem);

                                        break;
                                    #region Hide
                                    case PartType.Tee:
                                        break;
                                    case PartType.Transition:
                                        break;
                                    //case PartType.Cross:
                                    //    break;
                                    case PartType.Cap:
                                        break;
                                    case PartType.TapPerpendicular:
                                        break;
                                    case PartType.TapAdjustable:
                                        break;
                                    case PartType.Offset:
                                        break;
                                    case PartType.Union:
                                        break;
                                    case PartType.SpudPerpendicular:
                                        break;
                                    case PartType.SpudAdjustable:
                                        break;
                                    case PartType.InlineSensor:
                                        break;
                                    case PartType.Sensor:
                                        break;
                                    case PartType.EndCap:
                                        break;
                                    case PartType.PipeMechanicalCoupling:
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        break;
                        #endregion
                }

                //Prepare to restart iteration
                Model.AllConnectors.Remove(From);
                Model.AllConnectors.Remove(To);

                From = (from Connector c in Model.AllConnectors
                        where c.IsEqual(To)
                        select c).FirstOrDefault();

                if (From != null)
                {
                    curElem = From.Owner;

                    FromNode = ToNode;
                    FromNode.NextCon = From;
                    Model.AllNodes.Add(FromNode); //"ToNode" added to AllNodes
                }
                else
                {
                    continueSequence = false;
                    Model.Sequences.Add(curSequence);
                    curSequence = new AnalyticSequence();
                }
                ToNode = new Node();

            }
        }

        public void NumberNodes()
        {
            foreach (AnalyticSequence ans in Model.Sequences)
            {
                foreach (AnalyticElement ae in ans.Sequence)
                {

                }
            }
        }

        Connector DetermineToConnector(Document doc, Connector From)
        {
            Element owner = From.Owner;
            Connector

            Connector c2 = null;
            if (owner is Pipe pipe)
            {
                c2 = (from Connector c in pipe.ConnectorManager.Connectors //End of the host/dummy pipe
                      where c.Id != From.Id && (int)c.ConnectorType == 1
                      select c).FirstOrDefault();


            }
            else if (owner is FamilyInstance fi)
            {
                var cons = GetConnectors(owner);
                var c1mep = From.GetMEPConnectorInfo();
                if (c1mep.IsPrimary) c2 = cons.Secondary;
                else if (c1mep.IsSecondary) c2 = cons.Primary;
                else if (!c1mep.IsPrimary && !c1mep.IsSecondary) c2 = cons.Primary;
            }



            return c2;
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
        public List<AnalyticElement> Sequence { get; set; } = null;
    }

    class AnalyticModel
    {
        public List<AnalyticSequence> Sequences { get; } = new List<AnalyticSequence>();
        public List<Node> AllNodes { get; } = new List<Node>();
        public List<AnalyticElement> AllAnalyticElements { get; } = new List<AnalyticElement>();
        public List<Connector> AllConnectors { get; }
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
