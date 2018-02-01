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
            Connector To = null;

            Node FromNode = new Node();
            FromNode.FromConnector = From;
            FromNode.Number = 10; //TODO: Implement automatic numbering
            Model.AllNodes.Add(FromNode);

            Node ToNode = new Node();

            Element curElem = null;
            AnalyticElement curAElem = null;

            for (int i = 0; i < Model.AllElements.Count; i++)
            {
                curElem = From.Owner;

                switch (curElem)
                {
                    case Pipe pipe:

                        To = (from Connector c in pipe.ConnectorManager.Connectors //End of the host/dummy pipe
                              where c.Id != From.Id && (int)c.ConnectorType == 1
                              select c).FirstOrDefault();
                        curAElem = new AnalyticElement(curElem);
                        break;
                    case FamilyInstance fi:
                        int cat = fi.Category.Id.IntegerValue;
                        switch (cat)
                        {
                            case (int)BuiltInCategory.OST_PipeFitting:
                                var mf = fi.MEPModel as MechanicalFitting;
                                var partType = mf.PartType;
                                switch (partType)
                                {
                                    case PartType.Undefined:
                                        break;
                                    case PartType.Normal:
                                        break;
                                    case PartType.DuctMounted:
                                        break;
                                    case PartType.JunctionBox:
                                        break;
                                    case PartType.AttachesTo:
                                        break;
                                    case PartType.BreaksInto:
                                        break;
                                    case PartType.Elbow:
                                        break;
                                    case PartType.Tee:
                                        break;
                                    case PartType.Transition:
                                        break;
                                    case PartType.Cross:
                                        break;
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
                                    case PartType.PanelBoard:
                                        break;
                                    case PartType.Transformer:
                                        break;
                                    case PartType.SwitchBoard:
                                        break;
                                    case PartType.OtherPanel:
                                        break;
                                    case PartType.EquipmentSwitch:
                                        break;
                                    case PartType.Switch:
                                        break;
                                    case PartType.ValveBreaksInto:
                                        break;
                                    case PartType.SpudPerpendicular:
                                        break;
                                    case PartType.SpudAdjustable:
                                        break;
                                    case PartType.Damper:
                                        break;
                                    case PartType.Wye:
                                        break;
                                    case PartType.LateralTee:
                                        break;
                                    case PartType.LateralCross:
                                        break;
                                    case PartType.Pants:
                                        break;
                                    case PartType.MultiPort:
                                        break;
                                    case PartType.ValveNormal:
                                        break;
                                    case PartType.JunctionBoxTee:
                                        break;
                                    case PartType.JunctionBoxCross:
                                        break;
                                    case PartType.PipeFlange:
                                        break;
                                    case PartType.JunctionBoxElbow:
                                        break;
                                    case PartType.ChannelCableTrayElbow:
                                        break;
                                    case PartType.ChannelCableTrayVerticalElbow:
                                        break;
                                    case PartType.ChannelCableTrayCross:
                                        break;
                                    case PartType.ChannelCableTrayTee:
                                        break;
                                    case PartType.ChannelCableTrayTransition:
                                        break;
                                    case PartType.ChannelCableTrayUnion:
                                        break;
                                    case PartType.ChannelCableTrayOffset:
                                        break;
                                    case PartType.ChannelCableTrayMultiPort:
                                        break;
                                    case PartType.LadderCableTrayElbow:
                                        break;
                                    case PartType.LadderCableTrayVerticalElbow:
                                        break;
                                    case PartType.LadderCableTrayCross:
                                        break;
                                    case PartType.LadderCableTrayTee:
                                        break;
                                    case PartType.LadderCableTrayTransition:
                                        break;
                                    case PartType.LadderCableTrayUnion:
                                        break;
                                    case PartType.LadderCableTrayOffset:
                                        break;
                                    case PartType.LadderCableTrayMultiPort:
                                        break;
                                    case PartType.InlineSensor:
                                        break;
                                    case PartType.Sensor:
                                        break;
                                    case PartType.EndCap:
                                        break;
                                    case PartType.HandrailBracketHardware:
                                        break;
                                    case PartType.PanelBracketHardware:
                                        break;
                                    case PartType.TerminationHardware:
                                        break;
                                    case PartType.Rails:
                                        break;
                                    case PartType.Handrails:
                                        break;
                                    case PartType.TopRails:
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
                }

                ToNode.ToConnector = To;

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
        public Connector ToConnector { get; set; } = null;
        public Connector FromConnector { get; set; } = null;
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
