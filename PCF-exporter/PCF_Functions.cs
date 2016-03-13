using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Microsoft.Office.Interop.Excel;

namespace PCF_Functions
{
    public class InputVars
    {
        //File I/O
        public static string DriveLetter = "E";
        public static string OutputDirectoryFilePath = DriveLetter + ":\\Dropbox\\Revit\\Dev\\Test_01\\TestProject\\Output\\";

        public static string ExcelFilePath = DriveLetter + ":\\Dropbox\\Revit\\Dev\\Test_01\\TestProject\\";
        public static string ExcelFileName = "PCF_DEVELOPEMENT_01.xlsx";
        public static string ExcelSheet = "COMP";

        //Filters
        public static string SysAbbr = "FVF";
        public static BuiltInParameter SysAbbrParam = BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM;

        //Shared parameter group
        public const string PCF_GROUP_NAME = "PCF";
        public const BuiltInParameterGroup PCF_BUILTIN_GROUP_NAME = BuiltInParameterGroup.PG_ANALYTICAL_MODEL;

        //Parameter definition, remember to add new parameters to the ParameterList in the next region
        public static string PCF_ELEM_BP1 = "PCF_ELEM_BP1";
        public static ParameterType PCF_ELEM_BP1_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_BP1_GUID = new Guid("89b1e62e-f9b8-48c3-ab3a-1861a772bda8");

        public static string PCF_ELEM_CATEGORY = "PCF_ELEM_CATEGORY";
        public static ParameterType PCF_ELEM_CATEGORY_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_CATEGORY_GUID = new Guid("35efc6ed-2f20-4aca-bf05-d81d3b79dce2");

        public static string PCF_ELEM_COMPID = "PCF_ELEM_COMPID";
        public static ParameterType PCF_ELEM_COMPID_parameterType = ParameterType.Integer;
        public static Guid PCF_ELEM_COMPID_GUID = new Guid("876d2334-f860-4b5a-8c24-507e2c545fc0");

        public static string PCF_MAT_DESCR = "PCF_MAT_DESCR";
        public static ParameterType PCF_MAT_DESCR_parameterType = ParameterType.Text;
        public static Guid PCF_MAT_DESCR_GUID = new Guid("d39418f2-fcb3-4dd1-b0be-3d647486ebe6");

        public static string PCF_MAT_ID = "PCF_MAT_ID";
        public static ParameterType PCF_MAT_ID_parameterType = ParameterType.Integer;
        public static Guid PCF_MAT_ID_GUID = new Guid("fc5d3b19-af5b-47f6-a269-149b701c9364");

        public static string PCF_ELEM_TYPE = "PCF_ELEM_TYPE";
        public static ParameterType PCF_ELEM_TYPE_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_TYPE_GUID = new Guid("bfc7b779-786d-47cd-9194-8574a5059ec8");

        public static string PCF_ELEM_SKEY = "PCF_ELEM_SKEY";
        public static ParameterType PCF_ELEM_SKEY_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_SKEY_GUID = new Guid("3feebd29-054c-4ce8-bc64-3cff75ed6121");

        public static string PCF_ELEM_END1 = "PCF_ELEM_END1";
        public static ParameterType PCF_ELEM_END1_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_END1_GUID = new Guid("cbc10825-c0a1-471e-9902-075a41533738");

        public static string PCF_ELEM_END2 = "PCF_ELEM_END2";
        public static ParameterType PCF_ELEM_END2_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_END2_GUID = new Guid("ecaf3f8a-c28b-4a89-8496-728af3863b09");

        public static string PCF_ELEM_TAP1 = "PCF_ELEM_TAP1";
        public static ParameterType PCF_ELEM_TAP1_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_TAP1_GUID = new Guid("5fda303c-5536-429b-9fcc-afb40d14c7b3");

        public static string PCF_ELEM_TAP2 = "PCF_ELEM_TAP2";
        public static ParameterType PCF_ELEM_TAP2_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_TAP2_GUID = new Guid("e1e9bc3b-ce75-4f3a-ae43-c270f4fde937");

        public static string PCF_ELEM_TAP3 = "PCF_ELEM_TAP3";
        public static ParameterType PCF_ELEM_TAP3_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_TAP3_GUID = new Guid("12693653-8029-4743-be6a-310b1fbc0620");

        #region ParameterList
        //parameterAllNames and parameterTypes must correspond to each other in element position
        public static IList<string> parameterAllNames = new List<string>()
        {
            PCF_ELEM_CATEGORY, PCF_ELEM_BP1, PCF_ELEM_COMPID, PCF_MAT_DESCR, PCF_MAT_ID, PCF_ELEM_TYPE, PCF_ELEM_SKEY, PCF_ELEM_END1,
            PCF_ELEM_END2, PCF_ELEM_TAP1, PCF_ELEM_TAP2, PCF_ELEM_TAP3
        };

        public static IList<Guid> ParameterGUID = new List<Guid>()
        {
            PCF_ELEM_CATEGORY_GUID, PCF_ELEM_BP1_GUID, PCF_ELEM_COMPID_GUID, PCF_MAT_DESCR_GUID, PCF_MAT_ID_GUID,
            PCF_ELEM_TYPE_GUID, PCF_ELEM_SKEY_GUID, PCF_ELEM_END1_GUID, PCF_ELEM_END2_GUID, PCF_ELEM_TAP1_GUID, PCF_ELEM_TAP2_GUID,
            PCF_ELEM_TAP3_GUID
        };

        public static IList<ParameterType> parameterTypes = new List<ParameterType>()
        {
            PCF_ELEM_CATEGORY_parameterType, PCF_ELEM_BP1_parameterType, PCF_ELEM_COMPID_parameterType, PCF_MAT_DESCR_parameterType, PCF_MAT_ID_parameterType,
            PCF_ELEM_TYPE_parameterType, PCF_ELEM_SKEY_parameterType, PCF_ELEM_END1_parameterType, PCF_ELEM_END2_parameterType, PCF_ELEM_TAP1_parameterType,
            PCF_ELEM_TAP2_parameterType, PCF_ELEM_TAP3_parameterType
        };

        public static IList<string> parameterNames = new List<string>()
        {
            PCF_ELEM_CATEGORY, PCF_ELEM_BP1, PCF_MAT_DESCR, PCF_ELEM_TYPE, PCF_ELEM_SKEY, PCF_ELEM_END1,
            PCF_ELEM_END2
        };

        #endregion

        //PCF specification
        public static String PIPING_SPEC = "STD";
    }

    public class Composer
    {
        //PCF Preamble composition
        static StringBuilder sbPreamble = new StringBuilder();
        public static StringBuilder PreambleComposer()
        {
            sbPreamble.Append("ISOGEN-FILES ISOGEN.FLS");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BORE MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-CO-ORDS MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-WEIGHT KGS");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BOLT-DIA MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BOLT-LENGTH MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-WEIGHT-LENGTH METRE");
            sbPreamble.AppendLine();
            sbPreamble.Append("PIPELINE-REFERENCE " + InputVars.SysAbbr);
            sbPreamble.AppendLine();
            sbPreamble.Append("    PIPING-SPEC " + InputVars.PIPING_SPEC);
            sbPreamble.AppendLine();
            return sbPreamble;
        }

        static StringBuilder sbMaterials = new StringBuilder();
        static IEnumerable<IGrouping<string, Element>> materialGroups = null;
        static int groupNumber = 0;
        public static StringBuilder MaterialsSection(IEnumerable<IGrouping<string, Element>> elementGroups)
        {
            materialGroups = elementGroups;
            sbMaterials.Append("MATERIALS");
            foreach (IGrouping<string, Element> group in elementGroups)
            {
                groupNumber++;
                sbMaterials.AppendLine();
                sbMaterials.Append("MATERIAL-IDENTIFIER " + groupNumber);
                sbMaterials.AppendLine();
                sbMaterials.Append("    DESCRIPTION ");
                sbMaterials.Append(group.Key);
            }
            return sbMaterials;
        }
    }

    public class Filter
    {
        BuiltInParameter testParam; ParameterValueProvider pvp; FilterStringRuleEvaluator str; FilterStringRule paramFr; public ElementParameterFilter epf;

        public Filter(string valueQualifier, BuiltInParameter parameterName)
        {
            testParam = parameterName;
            pvp = new ParameterValueProvider(new ElementId((int)testParam));
            str = new FilterStringContains();
            paramFr = new FilterStringRule(pvp, str, valueQualifier, false);
            epf = new ElementParameterFilter(paramFr);
        }
    }

    public class Conversion
    {
        const double _inch_to_mm = 25.4;
        const double _foot_to_mm = 12 * _inch_to_mm;

        /// <summary>
        /// Return a string for a real number formatted to two decimal places.
        /// </summary>
        public static string RealString(double a)
        {
            //return a.ToString("0.##");
            return (Math.Truncate(a * 100) / 100).ToString("0.00", CultureInfo.GetCultureInfo("en-GB"));
        }

        /// <summary>
        /// Return a string for an XYZ point or vector with its coordinates converted from feet to millimetres and formatted to two decimal places.
        /// </summary>
        public static string PointStringMm(XYZ p)
        {
            return string.Format("{0:0.00} {1:0.00} {2:0.00}",
              RealString(p.X * _foot_to_mm),
              RealString(p.Y * _foot_to_mm),
              RealString(p.Z * _foot_to_mm));
        }

        public static string PipeSizeToMm(double l)
        {
            return string.Format("{0}", Math.Round(l * 2 * _foot_to_mm));
        }

        public static string AngleToPCF(double l)
        {
            return string.Format("{0}", l);
        }

        public static double RadianToDegree(double angle)
        {
            return angle * (180.0 / Math.PI);
        }
    }

    public class EndWriter
    {
        public static StringBuilder WriteEP1 (Element element, Connector connector)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            XYZ connectorOrigin = connector.Origin;
            double connectorSize = connector.Radius;
            sbEndWriter.Append("    END-POINT ");
            sbEndWriter.Append(Conversion.PointStringMm(connectorOrigin));
            sbEndWriter.Append(" ");
            sbEndWriter.Append(Conversion.PipeSizeToMm(connectorSize));
            if (element.LookupParameter(InputVars.PCF_ELEM_END1).HasValue == true)
            {
                sbEndWriter.Append(" ");
                sbEndWriter.Append(element.LookupParameter(InputVars.PCF_ELEM_END1).AsString());
            }
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteEP2(Element element, Connector connector)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            XYZ connectorOrigin = connector.Origin;
            double connectorSize = connector.Radius;
            sbEndWriter.Append("    END-POINT ");
            sbEndWriter.Append(Conversion.PointStringMm(connectorOrigin));
            sbEndWriter.Append(" ");
            sbEndWriter.Append(Conversion.PipeSizeToMm(connectorSize));
            if (element.LookupParameter(InputVars.PCF_ELEM_END2).HasValue == true)
            {
                sbEndWriter.Append(" ");
                sbEndWriter.Append(element.LookupParameter(InputVars.PCF_ELEM_END2).AsString());
            }
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteBP1(Element element, Connector connector)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            XYZ connectorOrigin = connector.Origin;
            double connectorSize = connector.Radius;
            sbEndWriter.Append("    BRANCH1-POINT ");
            sbEndWriter.Append(Conversion.PointStringMm(connectorOrigin));
            sbEndWriter.Append(" ");
            sbEndWriter.Append(Conversion.PipeSizeToMm(connectorSize));
            if (element.LookupParameter(InputVars.PCF_ELEM_BP1).HasValue == true)
            {
                sbEndWriter.Append(" ");
                sbEndWriter.Append(element.LookupParameter(InputVars.PCF_ELEM_BP1).AsString());
            }
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteCP(FamilyInstance familyInstance)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            XYZ elementLocation = ((LocationPoint)familyInstance.Location).Point;
            sbEndWriter.Append("    CENTRE-POINT ");
            sbEndWriter.Append(Conversion.PointStringMm(elementLocation));
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteCP(XYZ point)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            sbEndWriter.Append("    CENTRE-POINT ");
            sbEndWriter.Append(Conversion.PointStringMm(point));
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteCO(XYZ point)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            sbEndWriter.Append("    CO-ORDS ");
            sbEndWriter.Append(Conversion.PointStringMm(point));
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

    }
}