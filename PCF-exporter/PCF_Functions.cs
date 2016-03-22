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

namespace PCF_Functions
{
    public class InputVars
    {
        #region Execution
        //File I/O
        public static string OutputDirectoryFilePath;
        public static string ExcelSheet = "COMP";

        //Execution control
        public static bool ExportAll = true;

        //PCF File Header (preamble) control
        public static string UNITS_BORE = "MM";
        public static bool UNITS_BORE_MM = true;
        public static bool UNITS_BORE_INCH = false;

        public static string UNITS_CO_ORDS = "MM";
        public static bool UNITS_CO_ORDS_MM = true;
        public static bool UNITS_CO_ORDS_INCH = false;

        public static string UNITS_WEIGHT = "KGS";
        public static bool UNITS_WEIGHT_KGS = true;
        public static bool UNITS_WEIGHT_LBS = false;

        public static string UNITS_WEIGHT_LENGTH = "METER";
        public static bool UNITS_WEIGHT_LENGTH_METER = true;
        public static bool UNITS_WEIGHT_LENGTH_INCH = false;
        public static bool UNITS_WEIGHT_LENGTH_FEET = false;
        #endregion Execution

        #region Filters
        //Filters
        public static string SysAbbr = "FVF";
        public static BuiltInParameter SysAbbrParam = BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM;
        public static string PipelineGroupParameterName = "System Abbreviation";
        #endregion Filters

        #region Element parameter definition
        //Shared parameter group
        public const string PCF_GROUP_NAME = "PCF";
        public const BuiltInParameterGroup PCF_BUILTIN_GROUP_NAME = BuiltInParameterGroup.PG_ANALYTICAL_MODEL;

        //Element parameter definition, remember to add new parameters to the ParameterList in the next region
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
        #endregion

        #region Pipeline parameters
        //Pipeline parameter definition, remember to add new parameters to the ParameterList in the next region
        public static string PCF_PIPL_SPEC = "PCF_PIPL_SPEC";
        public static ParameterType PCF_PIPL_SPEC_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_SPEC_GUID = new Guid("7b0c932b-2ebe-495f-9d2e-effc350e8a59");

        public static string PCF_PIPL_TRACING = "PCF_PIPL_TRACING";
        public static ParameterType PCF_PIPL_TRACING_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_TRACING_GUID = new Guid("9d463d11-c9e8-4160-ac55-578795d11b1d");

        public static string PCF_PIPL_INSUL = "PCF_PIPL_INSUL";
        public static ParameterType PCF_PIPL_INSUL_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_INSUL_GUID = new Guid("d0c429fe-71db-4adc-b54a-58ae2fb4e127");

        public static string PCF_PIPL_PAINT = "PCF_PIPL_PAINT";
        public static ParameterType PCF_PIPL_PAINT_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_PAINT_GUID = new Guid("e440ed45-ce29-4b42-9a48-238b62b7522e");

        public static string PCF_PIPL_MISC1 = "PCF_PIPL_MISC1";
        public static ParameterType PCF_PIPL_MISC1_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_MISC1_GUID = new Guid("22f1dbed-2978-4474-9a8a-26fd14bc6aac");

        public static string PCF_PIPL_MISC2 = "PCF_PIPL_MISC2";
        public static ParameterType PCF_PIPL_MISC2_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_MISC2_GUID = new Guid("6492e7d8-cbc3-42f8-86c0-0ba9000d65ca");

        public static string PCF_PIPL_MISC3 = "PCF_PIPL_MISC3";
        public static ParameterType PCF_PIPL_MISC3_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_MISC3_GUID = new Guid("680bac72-0a1c-44a9-806d-991401f71912");

        public static string PCF_PIPL_MISC4 = "PCF_PIPL_MISC4";
        public static ParameterType PCF_PIPL_MISC4_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_MISC4_GUID = new Guid("6f904559-568b-4eff-a016-9c81e3a6c3ab");

        public static string PCF_PIPL_MISC5 = "PCF_PIPL_MISC5";
        public static ParameterType PCF_PIPL_MISC5_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_MISC5_GUID = new Guid("c375351b-b585-4fb1-92f7-abcdc10fd53a");

        public static string PCF_PIPL_JACKET = "PCF_PIPL_JACKET";
        public static ParameterType PCF_PIPL_JACKET_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_JACKET_GUID = new Guid("a810b6b8-17da-4191-b408-e046c758b289");

        public static string PCF_PIPL_REV = "PCF_PIPL_REV";
        public static ParameterType PCF_PIPL_REV_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_REV_GUID = new Guid("fb1a5913-4c64-4bfe-b50a-a8243a5db89f");

        public static string PCF_PIPL_PROJID = "PCF_PIPL_PROJID";
        public static ParameterType PCF_PIPL_PROJID_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_PROJID_GUID = new Guid("50509d7f-1b99-45f9-9b24-0c423dff5078");

        public static string PCF_PIPL_AREA = "PCF_PIPL_AREA";
        public static ParameterType PCF_PIPL_AREA_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_AREA_GUID = new Guid("642e8ab1-f87d-4da6-894e-a007a4a186a6");

        public static string PCF_PIPL_DATE = "PCF_PIPL_DATE";
        public static ParameterType PCF_PIPL_DATE_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_DATE_GUID = new Guid("86dc9abf-80fa-4c87-8079-4a28824ff529");

        public static string PCF_PIPL_NOMCLASS = "PCF_PIPL_NOMCLASS";
        public static ParameterType PCF_PIPL_NOMCLASS_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_NOMCLASS_GUID = new Guid("998fa331-7f38-4129-9939-8495fcd6c3ae");

        public static string PCF_PIPL_TEMP = "PCF_PIPL_TEMP";
        public static ParameterType PCF_PIPL_TEMP_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_TEMP_GUID = new Guid("7efb37ee-b1a1-4766-bb5b-015f823f36e2");

        public static string PCF_PIPL_TYPE = "PCF_PIPL_TYPE";
        public static ParameterType PCF_PIPL_TYPE_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_TYPE_GUID = new Guid("af00ee7d-cfc0-4e1c-a2cf-1626e4bb7eb0");

        public static string PCF_PIPL_GRAV = "PCF_PIPL_GRAV";
        public static ParameterType PCF_PIPL_GRAV_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_GRAV_GUID = new Guid("a32c0713-a6a5-4e6c-9a6b-d96e82159611");

        public static string PCF_PIPL_PREFIX = "PCF_PIPL_PREFIX";
        public static ParameterType PCF_PIPL_PREFIX_parameterType = ParameterType.Text;
        public static Guid PCF_PIPL_PREFIX_GUID = new Guid("c7136bbc-4b0d-47c6-95d1-8623ad015e8f");

        #endregion

        #region ParameterList
        //Element: parameterAllNames and parameterTypes must correspond to each other in element position
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

        //Pipeline: the same rules as above
        public static IList<string> parameterPipelineAllNames = new List<string>()
        {
            PCF_PIPL_AREA, PCF_PIPL_DATE, PCF_PIPL_GRAV, PCF_PIPL_INSUL, PCF_PIPL_JACKET, PCF_PIPL_MISC1, PCF_PIPL_MISC2, PCF_PIPL_MISC3,
            PCF_PIPL_MISC4, PCF_PIPL_MISC5, PCF_PIPL_NOMCLASS, PCF_PIPL_PAINT, PCF_PIPL_PREFIX, PCF_PIPL_PROJID, PCF_PIPL_REV, PCF_PIPL_SPEC,
            PCF_PIPL_TEMP, PCF_PIPL_TRACING, PCF_PIPL_TYPE
        };

        public static IList<Guid> parameterGuidPipeline = new List<Guid>()
        {
            PCF_PIPL_AREA_GUID, PCF_PIPL_DATE_GUID, PCF_PIPL_GRAV_GUID, PCF_PIPL_INSUL_GUID, PCF_PIPL_JACKET_GUID, PCF_PIPL_MISC1_GUID,
            PCF_PIPL_MISC2_GUID, PCF_PIPL_MISC3_GUID, PCF_PIPL_MISC4_GUID, PCF_PIPL_MISC5_GUID, PCF_PIPL_NOMCLASS_GUID, PCF_PIPL_PAINT_GUID,
            PCF_PIPL_PREFIX_GUID, PCF_PIPL_PROJID_GUID, PCF_PIPL_REV_GUID, PCF_PIPL_SPEC_GUID, PCF_PIPL_TEMP_GUID, PCF_PIPL_TRACING_GUID,
            PCF_PIPL_TYPE_GUID
        };

        public static IList<ParameterType> parameterTypesPipeline = new List<ParameterType>()
        {
            PCF_PIPL_AREA_parameterType, PCF_PIPL_DATE_parameterType, PCF_PIPL_GRAV_parameterType, PCF_PIPL_INSUL_parameterType,
            PCF_PIPL_JACKET_parameterType, PCF_PIPL_MISC1_parameterType, PCF_PIPL_MISC2_parameterType, PCF_PIPL_MISC3_parameterType,
            PCF_PIPL_MISC4_parameterType, PCF_PIPL_MISC5_parameterType, PCF_PIPL_NOMCLASS_parameterType, PCF_PIPL_PAINT_parameterType,
            PCF_PIPL_PREFIX_parameterType, PCF_PIPL_PROJID_parameterType, PCF_PIPL_REV_parameterType, PCF_PIPL_SPEC_parameterType,
            PCF_PIPL_TEMP_parameterType, PCF_PIPL_TRACING_parameterType, PCF_PIPL_TYPE_parameterType
        };

        #endregion

        //PCF specification
        public static string PIPING_SPEC = "STD";
    }

    public class Composer
    {
        #region Preamble
        //PCF Preamble composition
        static StringBuilder sbPreamble = new StringBuilder();
        public static StringBuilder PreambleComposer()
        {
            sbPreamble.Append("ISOGEN-FILES ISOGEN.FLS");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BORE "+InputVars.UNITS_BORE);
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-CO-ORDS "+InputVars.UNITS_CO_ORDS);
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-WEIGHT "+InputVars.UNITS_WEIGHT);
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BOLT-DIA MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-BOLT-LENGTH MM");
            sbPreamble.AppendLine();
            sbPreamble.Append("UNITS-WEIGHT-LENGTH "+InputVars.UNITS_WEIGHT_LENGTH);
            sbPreamble.AppendLine();
            return sbPreamble;
        }
        #endregion

        #region Materials section
        static StringBuilder sbMaterials = new StringBuilder();
        static IEnumerable<IGrouping<string, Element>> materialGroups = null;
        static int groupNumber;
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
                sbMaterials.Append("    DESCRIPTION "+group.Key);
            }
            return sbMaterials;
        }
        #endregion
    }

    public class Filter
    {
        BuiltInParameter testParam; ParameterValueProvider pvp; FilterStringRuleEvaluator str;
        FilterStringRule paramFr; public ElementParameterFilter epf;

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
        const double _foot_to_inch = 12;

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

        public static string PointStringInch(XYZ p)
        {
            return string.Format("{0:0.00} {1:0.00} {2:0.00}",
              RealString(p.X * _foot_to_inch),
              RealString(p.Y * _foot_to_inch),
              RealString(p.Z * _foot_to_inch));
        }

        public static string PipeSizeToMm(double l)
        {
            return string.Format("{0}", Math.Round(l * 2 * _foot_to_mm));
        }

        public static string PipeSizeToInch(double l)
        {
            return string.Format("{0}", RealString(l*2*_foot_to_inch));
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
            if (InputVars.UNITS_CO_ORDS_MM) sbEndWriter.Append(Conversion.PointStringMm(connectorOrigin));
            if (InputVars.UNITS_CO_ORDS_INCH) sbEndWriter.Append(Conversion.PointStringInch(connectorOrigin));
            sbEndWriter.Append(" ");
            if (InputVars.UNITS_BORE_MM) sbEndWriter.Append(Conversion.PipeSizeToMm(connectorSize));
            if (InputVars.UNITS_BORE_INCH) sbEndWriter.Append(Conversion.PipeSizeToInch(connectorSize));
            if (string.IsNullOrEmpty(element.LookupParameter(InputVars.PCF_ELEM_END1).AsString()) == false)
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
            if (InputVars.UNITS_CO_ORDS_MM) sbEndWriter.Append(Conversion.PointStringMm(connectorOrigin));
            if (InputVars.UNITS_CO_ORDS_INCH) sbEndWriter.Append(Conversion.PointStringInch(connectorOrigin));
            sbEndWriter.Append(" ");
            if (InputVars.UNITS_BORE_MM) sbEndWriter.Append(Conversion.PipeSizeToMm(connectorSize));
            if (InputVars.UNITS_BORE_INCH) sbEndWriter.Append(Conversion.PipeSizeToInch(connectorSize));
            if (string.IsNullOrEmpty(element.LookupParameter(InputVars.PCF_ELEM_END2).AsString()) == false)
            {
                sbEndWriter.Append(" ");
                sbEndWriter.Append(element.LookupParameter(InputVars.PCF_ELEM_END2).AsString());
            }
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteEP2(Element element, XYZ connector, double size)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            XYZ connectorOrigin = connector;
            double connectorSize = size;
            sbEndWriter.Append("    END-POINT ");
            if (InputVars.UNITS_CO_ORDS_MM) sbEndWriter.Append(Conversion.PointStringMm(connectorOrigin));
            if (InputVars.UNITS_CO_ORDS_INCH) sbEndWriter.Append(Conversion.PointStringInch(connectorOrigin));
            sbEndWriter.Append(" ");
            if (InputVars.UNITS_BORE_MM) sbEndWriter.Append(Conversion.PipeSizeToMm(connectorSize));
            if (InputVars.UNITS_BORE_INCH) sbEndWriter.Append(Conversion.PipeSizeToInch(connectorSize));
            if (string.IsNullOrEmpty(element.LookupParameter(InputVars.PCF_ELEM_END2).AsString()) == false)
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
            if (InputVars.UNITS_CO_ORDS_MM) sbEndWriter.Append(Conversion.PointStringMm(connectorOrigin));
            if (InputVars.UNITS_CO_ORDS_INCH) sbEndWriter.Append(Conversion.PointStringInch(connectorOrigin));
            sbEndWriter.Append(" ");
            if (InputVars.UNITS_BORE_MM) sbEndWriter.Append(Conversion.PipeSizeToMm(connectorSize));
            if (InputVars.UNITS_BORE_INCH) sbEndWriter.Append(Conversion.PipeSizeToInch(connectorSize));
            if (string.IsNullOrEmpty(element.LookupParameter(InputVars.PCF_ELEM_BP1).AsString()) == false)
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
            if (InputVars.UNITS_CO_ORDS_MM) sbEndWriter.Append(Conversion.PointStringMm(elementLocation));
            if (InputVars.UNITS_CO_ORDS_INCH) sbEndWriter.Append(Conversion.PointStringInch(elementLocation));
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteCP(XYZ point)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            sbEndWriter.Append("    CENTRE-POINT ");
            if (InputVars.UNITS_CO_ORDS_MM) sbEndWriter.Append(Conversion.PointStringMm(point));
            if (InputVars.UNITS_CO_ORDS_INCH) sbEndWriter.Append(Conversion.PointStringInch(point));
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

        public static StringBuilder WriteCO(XYZ point)
        {
            StringBuilder sbEndWriter = new StringBuilder();
            sbEndWriter.Append("    CO-ORDS ");
            if (InputVars.UNITS_CO_ORDS_MM) sbEndWriter.Append(Conversion.PointStringMm(point));
            if (InputVars.UNITS_CO_ORDS_INCH) sbEndWriter.Append(Conversion.PointStringInch(point));
            sbEndWriter.AppendLine();
            return sbEndWriter;
        }

    }

    public class ScheduleCreator
    {
        private UIDocument _uiDoc;
        public ICollection<ViewSchedule> CreateAllItemsSchedule(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            Document doc = uiDoc.Document;

            Transaction t = new Transaction(doc, "Create all items schedules");
            t.Start();

            List<ViewSchedule> schedules = new List<ViewSchedule>();

            ViewSchedule schedule = ViewSchedule.CreateSchedule(doc,ElementId.InvalidElementId,ElementId.InvalidElementId);
            schedule.Name = "PCF ALL Elements";
            schedules.Add(schedule);

            foreach (SchedulableField schField in schedule.Definition.GetSchedulableFields())
            {
                
            }

           




            return schedules;
        }
    }
}