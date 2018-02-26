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
using static CIINExporter.Enums;

namespace CIINExporter
{
    class CIIN_Analysis
    {
        Document doc;
        public AnalyticModel Model;

        public CIIN_Analysis(Document doc, HashSet<Element> elements)
        {
            Model = new AnalyticModel(elements);
            this.doc = doc;
        }

        public void AnalyzeSystem()
        {
            //Start analysis
            var openEnds = detectOpenEnds();

            Connector To = null;
            Connector From = null;
            Node FromNode = null;
            Node ToNode = null;
            Element curElem = null;
            AnalyticElement curAElem = null;
            AnalyticSequence curSequence = null;

            bool continueSequence = false;

            IList<Connector> branchEnds = new List<Connector>();

            for (int i = 0; i < Model.AllElements.Count; i++)
            {
                if (!continueSequence)
                {
                    if (branchEnds.Count > 0)
                    {
                        From = branchEnds.FirstOrDefault();
                        branchEnds = branchEnds.ExceptWhere(c => c.IsEqual(From)).ToList();
                        FromNode = (from Node n in Model.AllNodes
                                    where n.PreviousCon != null && n.PreviousCon.IsEqual(From)
                                    select n).FirstOrDefault();
                        FromNode.NextCon = From;
                    }
                    else
                    {
                        From = openEnds.FirstOrDefault();
                        openEnds = openEnds.ExceptWhere(c => c.IsEqual(From)).ToList();
                        FromNode = new Node();
                        FromNode.NextCon = From;
                        Model.AllNodes.Add(FromNode);



                        curSequence = new AnalyticSequence();
                    }

                    ToNode = new Node();

                    curElem = From.Owner;
                    curAElem = new AnalyticElement(curElem);

                    continueSequence = true;


                }

                switch (curElem)
                {
                    case Pipe pipe:

                        To = (from Connector c in pipe.ConnectorManager.Connectors //End of the host/dummy pipe
                              where c.Id != From.Id && (int)c.ConnectorType == 1
                              select c).FirstOrDefault();

                        //Test if the ToNode already exists
                        //TODO: This test must be copied to all other elements
                        Node existingToNode = (from Node n in Model.AllNodes
                                               where
                                               (n.PreviousCon != null && n.PreviousCon.IsEqual(To)) ||
                                               (n.NextCon != null && n.NextCon.IsEqual(To))
                                               select n).FirstOrDefault();

                        if (existingToNode != null)
                        {
                            ToNode = existingToNode;
                            if (ToNode.PreviousCon == null) ToNode.PreviousCon = To;
                            else if (ToNode.NextCon == null) ToNode.NextCon = To;
                        }
                        else
                        {
                            ToNode.PreviousCon = To;
                            Model.AllNodes.Add(ToNode);
                        }

                        curAElem.From = FromNode;
                        curAElem.To = ToNode;

                        //Assign correct element type to analytic element
                        curAElem.Type = ElemType.Pipe;

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
                                        XYZ elbowLoc = ((LocationPoint)fi.Location).Point;
                                        ToNode.PreviousLoc = elbowLoc; //The node has only element Location point defining it -
                                        ToNode.NextLoc = elbowLoc; //and not two adjacent Connectors as element connection nodes
                                        ToNode.Type = ElemType.Elbow;
                                        Model.AllNodes.Add(ToNode);

                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        curAElem.Type = ElemType.Elbow;

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

                                        curAElem.Type = ElemType.Pipe;

                                        curSequence.Sequence.Add(curAElem);
                                        break;
                                    case PartType.Tee:
                                        //Junction logic
                                        Node primNode = null;
                                        Node secNode = null;
                                        Node tertNode = null;

                                        //First analytic element
                                        XYZ teeLoc = ((LocationPoint)fi.Location).Point;
                                        ToNode.PreviousLoc = teeLoc; //The node has only element Location point defining it -
                                        ToNode.NextLoc = teeLoc; //and not two adjacent Connectors as element connection nodes
                                        ToNode.Type = ElemType.Tee;
                                        Model.AllNodes.Add(ToNode);

                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        curAElem.Type = ElemType.Tee;

                                        curSequence.Sequence.Add(curAElem);

                                        //Logic to return correct node to next element
                                        if (From.GetMEPConnectorInfo().IsPrimary) primNode = FromNode;
                                        else if (From.GetMEPConnectorInfo().IsSecondary) secNode = FromNode;
                                        else tertNode = FromNode;

                                        //From node is common for next two elements
                                        FromNode = ToNode;

                                        //Second Analytic Element
                                        if (From.GetMEPConnectorInfo().IsPrimary) To = cons.Secondary;
                                        else if (From.GetMEPConnectorInfo().IsSecondary) To = cons.Primary;
                                        else To = cons.Primary;

                                        ToNode = new Node();
                                        Model.AllNodes.Add(ToNode);
                                        ToNode.PreviousCon = To;
                                        curAElem = new AnalyticElement(curElem);
                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        curAElem.Type = ElemType.Pipe;

                                        curSequence.Sequence.Add(curAElem);

                                        //Logic to return correct node to next element
                                        if (From.GetMEPConnectorInfo().IsPrimary) secNode = ToNode;
                                        else if (From.GetMEPConnectorInfo().IsSecondary) primNode = ToNode;
                                        else primNode = ToNode;

                                        //Third Analytic Element
                                        if (From.GetMEPConnectorInfo().IsPrimary) To = cons.Tertiary;
                                        else if (From.GetMEPConnectorInfo().IsSecondary) To = cons.Tertiary;
                                        else To = cons.Secondary;

                                        ToNode = new Node();
                                        Model.AllNodes.Add(ToNode);
                                        ToNode.PreviousCon = To;
                                        curAElem = new AnalyticElement(curElem);
                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        curAElem.Type = ElemType.Pipe;

                                        curSequence.Sequence.Add(curAElem);

                                        //Logic to return correct node to next element
                                        if (From.GetMEPConnectorInfo().IsPrimary) tertNode = ToNode;
                                        else if (From.GetMEPConnectorInfo().IsSecondary) tertNode = ToNode;
                                        else secNode = ToNode;

                                        //Continuation logic
                                        Connector candidate1;
                                        Connector candidate2;

                                        if (From.GetMEPConnectorInfo().IsPrimary)
                                        {
                                            candidate1 = (from Connector c in Model.AllConnectors
                                                          where c.IsEqual(cons.Secondary)
                                                          select c).FirstOrDefault();
                                            candidate2 = (from Connector c in Model.AllConnectors
                                                          where c.IsEqual(cons.Tertiary)
                                                          select c).FirstOrDefault();

                                            if (candidate1 != null)
                                            {
                                                To = cons.Secondary;
                                                ToNode = secNode;
                                                if (candidate2 != null) branchEnds.Add(candidate2);
                                            }
                                            else if (candidate2 != null)
                                            {
                                                To = cons.Tertiary;
                                                ToNode = tertNode;
                                            }
                                        }
                                        else if (From.GetMEPConnectorInfo().IsSecondary)
                                        {
                                            candidate1 = (from Connector c in Model.AllConnectors
                                                          where c.IsEqual(cons.Primary)
                                                          select c).FirstOrDefault();
                                            candidate2 = (from Connector c in Model.AllConnectors
                                                          where c.IsEqual(cons.Tertiary)
                                                          select c).FirstOrDefault();

                                            if (candidate1 != null)
                                            {
                                                To = cons.Primary;
                                                ToNode = primNode;
                                                if (candidate2 != null) branchEnds.Add(candidate2);
                                            }
                                            else if (candidate2 != null)
                                            {
                                                To = cons.Tertiary;
                                                ToNode = tertNode;
                                            }
                                        }
                                        else
                                        {
                                            candidate1 = (from Connector c in Model.AllConnectors
                                                          where c.IsEqual(cons.Primary)
                                                          select c).FirstOrDefault();
                                            candidate2 = (from Connector c in Model.AllConnectors
                                                          where c.IsEqual(cons.Secondary)
                                                          select c).FirstOrDefault();

                                            if (candidate1 != null)
                                            {
                                                To = cons.Primary;
                                                ToNode = primNode;
                                                if (candidate2 != null) branchEnds.Add(candidate2);
                                            }
                                            else if (candidate2 != null)
                                            {
                                                To = cons.Secondary;
                                                ToNode = secNode;
                                            }
                                        }
                                        break;
                                    case PartType.Transition:
                                        To = (from Connector c in GetALLConnectorsFromElements(curElem)
                                              where c.Id != From.Id
                                              select c).FirstOrDefault();
                                        ToNode.PreviousCon = To;
                                        Model.AllNodes.Add(ToNode);

                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        //Assign correct element type to analytic element
                                        curAElem.Type = ElemType.Transition;

                                        curSequence.Sequence.Add(curAElem);
                                        break;
                                    case PartType.Cap:
                                        //Handles flanges because of the workaround PartType.Cap for flanges
                                        //Real Caps are ignored for now
                                        To = (from Connector c in GetALLConnectorsFromElements(curElem)
                                              where c.Id != From.Id
                                              select c).FirstOrDefault();
                                        ToNode.PreviousCon = To;
                                        Model.AllNodes.Add(ToNode);

                                        curAElem.From = FromNode;
                                        curAElem.To = ToNode;

                                        //Assign correct element type to analytic element
                                        curAElem.Type = ElemType.Rigid;

                                        curSequence.Sequence.Add(curAElem);
                                        break;
                                    case PartType.Union:
                                        throw new NotImplementedException();
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

                                //Assign correct element type to analytic element
                                curAElem.Type = ElemType.Rigid;

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

                    ToNode = new Node();

                    curAElem = new AnalyticElement(curElem);
                }
                else
                {
                    continueSequence = false;
                    openEnds = openEnds.ExceptWhere(c => c.IsEqual(To)).ToList();

                    if (branchEnds.Count > 0 && To != null)
                    {
                        branchEnds = branchEnds.ExceptWhere(c => c.IsEqual(To)).ToList();
                    }

                    if (branchEnds.Count < 1) Model.Sequences.Add(curSequence);

                }
            }

            //Reference ALL analytic elements in the collecting pool
            foreach (var sequence in Model.Sequences)
            {
                foreach (var item in sequence.Sequence)
                {
                    Model.AllAnalyticElements.Add(item);
                }
            }

            Util.InfoMsg(Model.AllNodes.Count.ToString());
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
                    if (ae.To.Number == 0) tensCount++; //Possible fix for skipping a number after branches.
                    if (tensCount == 1)
                    {
                        if (ae.From.Number == 0) ae.From.Number = thousands + 10;
                        if (ae.To.Number == 0) ae.To.Number = thousands + 20;
                        tensCount = 2;
                        continue;
                    }
                    tens = tensCount * 10;
                    if (ae.To.Number == 0) ae.To.Number = thousands + tens;
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

    public class Node
    {
        public Connector PreviousCon { get; set; } = null;
        public Connector NextCon { get; set; } = null;
        public XYZ PreviousLoc { get; set; } = null;
        public XYZ NextLoc { get; set; } = null;
        public int Number { get; set; } = 0;
        public ElemType Type { get; set; } = 0;

        //public Node(Connector connector) => ToConnector = connector;
    }

    public class AnalyticElement
    {
        public Node From { get; set; } = null;
        public Node To { get; set; } = null;
        public Element Element { get; set; } = null;
        public Enums.ElemType Type { get; set; } = 0;

        public AnalyticElement(Element element) => Element = element;
    }

    public class AnalyticSequence
    {
        public List<AnalyticElement> Sequence { get; set; } = new List<AnalyticElement>();
    }

    public class AnalyticModel
    {
        public List<AnalyticSequence> Sequences { get; } = new List<AnalyticSequence>();
        public List<Node> AllNodes { get; } = new List<Node>();
        public List<AnalyticElement> AllAnalyticElements { get; } = new List<AnalyticElement>();
        public List<Connector> AllConnectors { get; set; }
        public List<Element> AllElements { get; }
        public ModelData Data { get; set; } = null;

        public AnalyticModel(HashSet<Element> elements)
        {
            AllElements = elements.ToList();
            AllConnectors = GetALLConnectorsFromElements(elements).ToList();
        }
    }

    public class ModelData
    {
        public StringBuilder _01_VERSION { get; set; } = new StringBuilder();
        public StringBuilder _02_CONTROL { get; set; } = new StringBuilder();
        public StringBuilder _03_ELEMENTS { get; set; } = new StringBuilder();
        public StringBuilder _04_AUXDATA { get; } = new StringBuilder("#$ AUX_DATA\n");
        public StringBuilder _05_NODENAME { get; set; } = new StringBuilder();
        public StringBuilder _06_BEND { get; set; } = new StringBuilder();
        public StringBuilder _07_RIGID { get; set; } = new StringBuilder();
        public StringBuilder _08_EXPJT { get; set; } = new StringBuilder();
        public StringBuilder _09_RESTRANT { get; set; } = new StringBuilder();
        public StringBuilder _10_DISPLMNT { get; set; } = new StringBuilder();
        public StringBuilder _11_FORCMNT { get; set; } = new StringBuilder();
        public StringBuilder _12_UNIFORM { get; set; } = new StringBuilder();
        public StringBuilder _13_WIND { get; set; } = new StringBuilder();
        public StringBuilder _14_OFFSETS { get; set; } = new StringBuilder();
        public StringBuilder _15_ALLOWBLS { get; set; } = new StringBuilder();
        public StringBuilder _16_SIFTEES { get; set; } = new StringBuilder();
        public StringBuilder _17_REDUCERS { get; set; } = new StringBuilder();
        public StringBuilder _18_FLANGES { get; set; } = new StringBuilder();
        public StringBuilder _19_EQUIPMNT { get; set; } = new StringBuilder();
        public StringBuilder _20_MISCEL_1 { get; set; } = new StringBuilder();
        public StringBuilder _21_UNITS { get; set; } = new StringBuilder();
        public StringBuilder _22_COORDS { get; set; } = new StringBuilder();

        public ModelData(AnalyticModel Model)
        {

        }
    }

    public static class Enums
    {
        public enum ElemType
        {
            NotAssigned = 0,
            Pipe = 1,
            Elbow = 2,
            Tee = 3,
            Transition = 4,
            Flange = 5,
            Rigid = 6
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
