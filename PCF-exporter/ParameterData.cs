using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic;
using System.Text;
using System.Globalization;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

using BuildingCoder;
using iv = PCF_Functions.InputVars;
using pd = PCF_Functions.ParameterData;
using pdef = PCF_Functions.ParameterDefinition;

namespace PCF_Functions
{
    public class ParameterDefinition
    {
        public readonly IList<pdef> ListParametersAll = new List<pdef>();

        public ParameterDefinition()
        {
            //Populate the list with element parameters
            //User defined
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_TYPE, "ELEM", "U", pd.Text, pd.PCF_ELEM_TYPE_GUID, ""));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_SKEY, "ELEM", "U", pd.Text, pd.PCF_ELEM_SKEY_GUID, "SKEY"));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_SPEC, "ELEM", "U", pd.Text, pd.PCF_ELEM_SPEC_GUID, "PIPING-SPEC"));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_CATEGORY, "ELEM", "U", pd.Text, pd.PCF_ELEM_CATEGORY_GUID, "CATEGORY"));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_END1, "ELEM", "U", pd.Text, pd.PCF_ELEM_END1_GUID, ""));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_END2, "ELEM", "U", pd.Text, pd.PCF_ELEM_END2_GUID, ""));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_BP1, "ELEM", "U", pd.Text, pd.PCF_ELEM_BP1_GUID, ""));
            ListParametersAll.Add(new pdef("PCF_ELEM_STATUS", "ELEM", "U", pd.Text, new Guid("c16e4db2-15e8-41ac-9b8f-134e133df8a4"), "STATUS"));
            ListParametersAll.Add(new pdef("PCF_ELEM_TRACING_SPEC", "ELEM", "U", pd.Text, new Guid("8e1d43fb-9cd2-4591-a1f5-ba392f0a8708"), "TRACING-SPEC"));
            ListParametersAll.Add(new pdef("PCF_ELEM_INSUL_SPEC", "ELEM", "U", pd.Text, new Guid("d628605e-c0bf-43dc-9f05-e22dbae2022e"), "INSULATION-SPEC"));
            ListParametersAll.Add(new pdef("PCF_ELEM_PAINT_SPEC", "ELEM", "U", pd.Text, new Guid("b51db394-85ee-43af-9117-bb255ac0aaac"), "PAINTING-SPEC"));
            ListParametersAll.Add(new pdef("PCF_ELEM_MISC1", "ELEM", "U", pd.Text, new Guid("ea4315ce-e5f5-4538-a6e9-f548068c3c66"), "MISC-SPEC1"));
            ListParametersAll.Add(new pdef("PCF_ELEM_MISC2", "ELEM", "U", pd.Text, new Guid("cca78e21-5ed7-44bc-9dab-844997a1b965"), "MISC-SPEC2"));
            ListParametersAll.Add(new pdef("PCF_ELEM_MISC3", "ELEM", "U", pd.Text, new Guid("0e065f3e-83c8-44c8-a1cb-babaf20476b9"), "MISC-SPEC3"));
            ListParametersAll.Add(new pdef("PCF_ELEM_MISC4", "ELEM", "U", pd.Text, new Guid("3229c505-3802-416c-bf04-c109f41f3ab7"), "MISC-SPEC4"));
            ListParametersAll.Add(new pdef("PCF_ELEM_MISC5", "ELEM", "U", pd.Text, new Guid("692e2e97-3b9c-4616-8a03-dfd493b01762"), "MISC-SPEC5"));

            //Material
            ListParametersAll.Add(new pdef(pd.PCF_MAT_DESCR, "ELEM", "U", pd.Text, pd.PCF_MAT_DESCR_GUID, ""));

            //Programattically defined
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_TAP1, "ELEM", "P", pd.Text, pd.PCF_ELEM_TAP1_GUID, ""));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_TAP2, "ELEM", "P", pd.Text, pd.PCF_ELEM_TAP2_GUID, ""));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_TAP3, "ELEM", "P", pd.Text, pd.PCF_ELEM_TAP3_GUID, ""));
            ListParametersAll.Add(new pdef(pd.PCF_ELEM_COMPID, "ELEM", "P", pd.Integer, pd.PCF_ELEM_COMPID_GUID, ""));
            ListParametersAll.Add(new pdef(pd.PCF_MAT_ID, "ELEM", "P", pd.Integer, pd.PCF_MAT_ID_GUID, "MATERIAL-IDENTIFIER"));

            //Populate the list with pipeline parameters
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_AREA, "PIPL", "U", pd.Text, pd.PCF_PIPL_AREA_GUID, "AREA"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_DATE, "PIPL", "U", pd.Text, pd.PCF_PIPL_DATE_GUID, "DATE-DMY"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_GRAV, "PIPL", "U", pd.Text, pd.PCF_PIPL_GRAV_GUID, "SPECIFIC-GRAVITY"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_INSUL, "PIPL", "U", pd.Text, pd.PCF_PIPL_INSUL_GUID, "INSULATION-SPEC"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_JACKET, "PIPL", "U", pd.Text, pd.PCF_PIPL_JACKET_GUID, "JACKET-SPEC"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_MISC1, "PIPL", "U", pd.Text, pd.PCF_PIPL_MISC1_GUID, "MISC-SPEC1"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_MISC2, "PIPL", "U", pd.Text, pd.PCF_PIPL_MISC2_GUID, "MISC-SPEC2"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_MISC3, "PIPL", "U", pd.Text, pd.PCF_PIPL_MISC3_GUID, "MISC-SPEC3"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_MISC4, "PIPL", "U", pd.Text, pd.PCF_PIPL_MISC4_GUID, "MISC-SPEC4"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_MISC5, "PIPL", "U", pd.Text, pd.PCF_PIPL_MISC5_GUID, "MISC-SPEC5"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_NOMCLASS, "PIPL", "U", pd.Text, pd.PCF_PIPL_NOMCLASS_GUID, "NOMINAL-CLASS"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_PAINT, "PIPL", "U", pd.Text, pd.PCF_PIPL_PAINT_GUID, "PAINTING-SPEC"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_PREFIX, "PIPL", "U", pd.Text, pd.PCF_PIPL_PREFIX_GUID, "SPOOL-PREFIX"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_PROJID, "PIPL", "U", pd.Text, pd.PCF_PIPL_PROJID_GUID, "PROJECT-IDENTIFIER"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_REV, "PIPL", "U", pd.Text, pd.PCF_PIPL_REV_GUID, "REVISION"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_SPEC, "PIPL", "U", pd.Text, pd.PCF_PIPL_SPEC_GUID, "PIPING-SPEC"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_TEMP, "PIPL", "U", pd.Text, pd.PCF_PIPL_TEMP_GUID, "PIPELINE-TEMP"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_TRACING, "PIPL", "U", pd.Text, pd.PCF_PIPL_TRACING_GUID, "TRACING-SPEC"));
            ListParametersAll.Add(new pdef(pd.PCF_PIPL_TYPE, "PIPL", "U", pd.Text, pd.PCF_PIPL_TYPE_GUID, "PIPELINE-TYPE"));
        }

       public ParameterDefinition(string pName, string pDomain, string pUsage, ParameterType pType, Guid pGuid, string pKeyword)
        {
            Name = pName;
            Domain = pDomain;
            Usage = pUsage; //U = user, P = programmatic
            Type = pType;
            Guid = pGuid;
            Keyword = pKeyword;
        }

        public string Name { get; }
        public string Domain { get; } //PIPL = Pipeline, ELEM = Element
        public string Usage { get; } //U = user defined values, P = programatically defined values
        public ParameterType Type { get; }
        public Guid Guid { get; }
        public string Keyword { get; } //The keyword as defined in the PCF reference guide
    }

    public static class ParameterData
    {
        #region Parameter Data Entry

        //general values
        public const ParameterType Text = ParameterType.Text;
        public const ParameterType Integer = ParameterType.Integer;

        public static string PCF_ELEM_BP1 = "PCF_ELEM_BP1";
        public static Guid PCF_ELEM_BP1_GUID = new Guid("89b1e62e-f9b8-48c3-ab3a-1861a772bda8");

        public static string PCF_ELEM_CATEGORY = "PCF_ELEM_CATEGORY";
        public static Guid PCF_ELEM_CATEGORY_GUID = new Guid("35efc6ed-2f20-4aca-bf05-d81d3b79dce2");

        public static string PCF_ELEM_COMPID = "PCF_ELEM_COMPID";
        public static Guid PCF_ELEM_COMPID_GUID = new Guid("876d2334-f860-4b5a-8c24-507e2c545fc0");

        public static string PCF_ELEM_SPEC = "PCF_ELEM_SPEC";
        public static Guid PCF_ELEM_SPEC_GUID = new Guid("90be8246-25f7-487d-b352-554f810fcaa7");

        public static string PCF_MAT_DESCR = "PCF_MAT_DESCR";
        public static Guid PCF_MAT_DESCR_GUID = new Guid("d39418f2-fcb3-4dd1-b0be-3d647486ebe6");

        public static string PCF_MAT_ID = "PCF_MAT_ID";
        public static Guid PCF_MAT_ID_GUID = new Guid("fc5d3b19-af5b-47f6-a269-149b701c9364");

        public static string PCF_ELEM_TYPE = "PCF_ELEM_TYPE";
        public static Guid PCF_ELEM_TYPE_GUID = new Guid("bfc7b779-786d-47cd-9194-8574a5059ec8");

        public static string PCF_ELEM_SKEY = "PCF_ELEM_SKEY";
        public static Guid PCF_ELEM_SKEY_GUID = new Guid("3feebd29-054c-4ce8-bc64-3cff75ed6121");

        public static string PCF_ELEM_END1 = "PCF_ELEM_END1";
        public static Guid PCF_ELEM_END1_GUID = new Guid("cbc10825-c0a1-471e-9902-075a41533738");

        public static string PCF_ELEM_END2 = "PCF_ELEM_END2";
        public static Guid PCF_ELEM_END2_GUID = new Guid("ecaf3f8a-c28b-4a89-8496-728af3863b09");

        public static string PCF_ELEM_TAP1 = "PCF_ELEM_TAP1";
        public static Guid PCF_ELEM_TAP1_GUID = new Guid("5fda303c-5536-429b-9fcc-afb40d14c7b3");

        public static string PCF_ELEM_TAP2 = "PCF_ELEM_TAP2";
        public static Guid PCF_ELEM_TAP2_GUID = new Guid("e1e9bc3b-ce75-4f3a-ae43-c270f4fde937");

        public static string PCF_ELEM_TAP3 = "PCF_ELEM_TAP3";
        public static Guid PCF_ELEM_TAP3_GUID = new Guid("12693653-8029-4743-be6a-310b1fbc0620");

        public static string PCF_PIPL_SPEC = "PCF_PIPL_SPEC";
        public static Guid PCF_PIPL_SPEC_GUID = new Guid("7b0c932b-2ebe-495f-9d2e-effc350e8a59");

        public static string PCF_PIPL_TRACING = "PCF_PIPL_TRACING";
        public static Guid PCF_PIPL_TRACING_GUID = new Guid("9d463d11-c9e8-4160-ac55-578795d11b1d");

        public static string PCF_PIPL_INSUL = "PCF_PIPL_INSUL";
        public static Guid PCF_PIPL_INSUL_GUID = new Guid("d0c429fe-71db-4adc-b54a-58ae2fb4e127");

        public static string PCF_PIPL_PAINT = "PCF_PIPL_PAINT";
        public static Guid PCF_PIPL_PAINT_GUID = new Guid("e440ed45-ce29-4b42-9a48-238b62b7522e");

        public static string PCF_PIPL_MISC1 = "PCF_PIPL_MISC1";
        public static Guid PCF_PIPL_MISC1_GUID = new Guid("22f1dbed-2978-4474-9a8a-26fd14bc6aac");

        public static string PCF_PIPL_MISC2 = "PCF_PIPL_MISC2";
        public static Guid PCF_PIPL_MISC2_GUID = new Guid("6492e7d8-cbc3-42f8-86c0-0ba9000d65ca");

        public static string PCF_PIPL_MISC3 = "PCF_PIPL_MISC3";
        public static Guid PCF_PIPL_MISC3_GUID = new Guid("680bac72-0a1c-44a9-806d-991401f71912");

        public static string PCF_PIPL_MISC4 = "PCF_PIPL_MISC4";
        public static Guid PCF_PIPL_MISC4_GUID = new Guid("6f904559-568b-4eff-a016-9c81e3a6c3ab");

        public static string PCF_PIPL_MISC5 = "PCF_PIPL_MISC5";
        public static Guid PCF_PIPL_MISC5_GUID = new Guid("c375351b-b585-4fb1-92f7-abcdc10fd53a");

        public static string PCF_PIPL_JACKET = "PCF_PIPL_JACKET";
        public static Guid PCF_PIPL_JACKET_GUID = new Guid("a810b6b8-17da-4191-b408-e046c758b289");

        public static string PCF_PIPL_REV = "PCF_PIPL_REV";
        public static Guid PCF_PIPL_REV_GUID = new Guid("fb1a5913-4c64-4bfe-b50a-a8243a5db89f");

        public static string PCF_PIPL_PROJID = "PCF_PIPL_PROJID";
        public static Guid PCF_PIPL_PROJID_GUID = new Guid("50509d7f-1b99-45f9-9b24-0c423dff5078");

        public static string PCF_PIPL_AREA = "PCF_PIPL_AREA";
        public static Guid PCF_PIPL_AREA_GUID = new Guid("642e8ab1-f87d-4da6-894e-a007a4a186a6");

        public static string PCF_PIPL_DATE = "PCF_PIPL_DATE";
        public static Guid PCF_PIPL_DATE_GUID = new Guid("86dc9abf-80fa-4c87-8079-4a28824ff529");

        public static string PCF_PIPL_NOMCLASS = "PCF_PIPL_NOMCLASS";
        public static Guid PCF_PIPL_NOMCLASS_GUID = new Guid("998fa331-7f38-4129-9939-8495fcd6c3ae");

        public static string PCF_PIPL_TEMP = "PCF_PIPL_TEMP";
        public static Guid PCF_PIPL_TEMP_GUID = new Guid("7efb37ee-b1a1-4766-bb5b-015f823f36e2");

        public static string PCF_PIPL_TYPE = "PCF_PIPL_TYPE";
        public static Guid PCF_PIPL_TYPE_GUID = new Guid("af00ee7d-cfc0-4e1c-a2cf-1626e4bb7eb0");

        public static string PCF_PIPL_GRAV = "PCF_PIPL_GRAV";
        public static Guid PCF_PIPL_GRAV_GUID = new Guid("a32c0713-a6a5-4e6c-9a6b-d96e82159611");

        public static string PCF_PIPL_PREFIX = "PCF_PIPL_PREFIX";
        public static Guid PCF_PIPL_PREFIX_GUID = new Guid("c7136bbc-4b0d-47c6-95d1-8623ad015e8f");

        #endregion

        public static IList<string> parameterNames = new List<string>();
    }
}