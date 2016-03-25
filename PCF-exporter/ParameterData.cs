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

namespace PCF_Functions
{
    public class ParameterDefinition
    {
        public readonly IList<ParameterDefinition> ElementParametersAll = new List<ParameterDefinition>();
        public readonly IList<ParameterDefinition> PipelineParametersAll = new List<ParameterDefinition>();
        public ParameterDefinition()
        {
            //Populate the list with element parameters
            //User defined
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_TYPE, pd.PCF_ELEM_TYPE_parameterType, pd.PCF_ELEM_TYPE_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_SKEY, pd.PCF_ELEM_SKEY_parameterType, pd.PCF_ELEM_SKEY_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_SPEC, pd.PCF_ELEM_SPEC_parameterType, pd.PCF_ELEM_SPEC_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_CATEGORY, pd.PCF_ELEM_CATEGORY_parameterType, pd.PCF_ELEM_CATEGORY_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_END1, pd.PCF_ELEM_END1_parameterType, pd.PCF_ELEM_END1_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_END2, pd.PCF_ELEM_END2_parameterType, pd.PCF_ELEM_END2_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_BP1, pd.PCF_ELEM_BP1_parameterType, pd.PCF_ELEM_BP1_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_MAT_DESCR, pd.PCF_MAT_DESCR_parameterType, pd.PCF_MAT_DESCR_GUID));
            //Programattically defined
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_TAP1, pd.PCF_ELEM_TAP1_parameterType, pd.PCF_ELEM_TAP1_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_TAP2, pd.PCF_ELEM_TAP2_parameterType, pd.PCF_ELEM_TAP2_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_TAP3, pd.PCF_ELEM_TAP3_parameterType, pd.PCF_ELEM_TAP3_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_ELEM_COMPID, pd.PCF_ELEM_COMPID_parameterType, pd.PCF_ELEM_COMPID_GUID));
            ElementParametersAll.Add(new ParameterDefinition(pd.PCF_MAT_ID, pd.PCF_MAT_ID_parameterType, pd.PCF_MAT_ID_GUID));

            //Populate the list with pipeline parameters
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_AREA, pd.PCF_PIPL_AREA_parameterType, pd.PCF_PIPL_AREA_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_DATE, pd.PCF_PIPL_DATE_parameterType, pd.PCF_PIPL_DATE_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_GRAV, pd.PCF_PIPL_GRAV_parameterType, pd.PCF_PIPL_GRAV_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_INSUL, pd.PCF_PIPL_INSUL_parameterType, pd.PCF_PIPL_INSUL_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_JACKET, pd.PCF_PIPL_JACKET_parameterType, pd.PCF_PIPL_JACKET_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_MISC1, pd.PCF_PIPL_MISC1_parameterType, pd.PCF_PIPL_MISC1_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_MISC2, pd.PCF_PIPL_MISC2_parameterType, pd.PCF_PIPL_MISC2_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_MISC3, pd.PCF_PIPL_MISC3_parameterType, pd.PCF_PIPL_MISC3_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_MISC4, pd.PCF_PIPL_MISC4_parameterType, pd.PCF_PIPL_MISC4_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_MISC5, pd.PCF_PIPL_MISC5_parameterType, pd.PCF_PIPL_MISC5_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_NOMCLASS, pd.PCF_PIPL_NOMCLASS_parameterType, pd.PCF_PIPL_NOMCLASS_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_PAINT, pd.PCF_PIPL_PAINT_parameterType, pd.PCF_PIPL_PAINT_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_PREFIX, pd.PCF_PIPL_PREFIX_parameterType, pd.PCF_PIPL_PREFIX_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_PROJID, pd.PCF_PIPL_PROJID_parameterType, pd.PCF_PIPL_PROJID_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_REV, pd.PCF_PIPL_REV_parameterType, pd.PCF_PIPL_REV_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_SPEC, pd.PCF_PIPL_SPEC_parameterType, pd.PCF_PIPL_SPEC_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_TEMP, pd.PCF_PIPL_TEMP_parameterType, pd.PCF_PIPL_TEMP_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_TRACING, pd.PCF_PIPL_TRACING_parameterType, pd.PCF_PIPL_TRACING_GUID));
            PipelineParametersAll.Add(new ParameterDefinition(pd.PCF_PIPL_TYPE, pd.PCF_PIPL_TYPE_parameterType, pd.PCF_PIPL_TYPE_GUID));
        }

        public ParameterDefinition(string pName, ParameterType pType, Guid pGuid)
        {
            Name = pName;
            Type = pType;
            Guid = pGuid;
        }

        public string Name { get; set; }
        public ParameterType Type { get; set; }
        public Guid Guid { get; set; }
    }

    public static class ParameterData
    {
        #region Parameter Data Entry
        public static string PCF_ELEM_BP1 = "PCF_ELEM_BP1";
        public static ParameterType PCF_ELEM_BP1_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_BP1_GUID = new Guid("89b1e62e-f9b8-48c3-ab3a-1861a772bda8");

        public static string PCF_ELEM_CATEGORY = "PCF_ELEM_CATEGORY";
        public static ParameterType PCF_ELEM_CATEGORY_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_CATEGORY_GUID = new Guid("35efc6ed-2f20-4aca-bf05-d81d3b79dce2");

        public static string PCF_ELEM_COMPID = "PCF_ELEM_COMPID";
        public static ParameterType PCF_ELEM_COMPID_parameterType = ParameterType.Integer;
        public static Guid PCF_ELEM_COMPID_GUID = new Guid("876d2334-f860-4b5a-8c24-507e2c545fc0");

        public static string PCF_ELEM_SPEC = "PCF_ELEM_SPEC";
        public static ParameterType PCF_ELEM_SPEC_parameterType = ParameterType.Text;
        public static Guid PCF_ELEM_SPEC_GUID = new Guid("90be8246-25f7-487d-b352-554f810fcaa7");

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

        public static IList<string> parameterAllNames = new List<string>
        {
            PCF_ELEM_CATEGORY, PCF_ELEM_BP1, PCF_ELEM_COMPID, PCF_MAT_DESCR, PCF_MAT_ID, PCF_ELEM_TYPE, PCF_ELEM_SKEY, PCF_ELEM_END1,
            PCF_ELEM_END2, PCF_ELEM_TAP1, PCF_ELEM_TAP2, PCF_ELEM_TAP3
        };

        public static IList<string> parameterNames = new List<string>
        {
            PCF_ELEM_TYPE, PCF_ELEM_SKEY, PCF_ELEM_CATEGORY, PCF_ELEM_SPEC, PCF_ELEM_END1,
            PCF_ELEM_END2, PCF_ELEM_BP1, PCF_MAT_DESCR, PCF_ELEM_TAP1, PCF_ELEM_TAP2, PCF_ELEM_TAP3
        };

        //Add here user defined parameter names
        public static readonly IList<string> elemParametersExport = new List<string>
        {
            PCF_ELEM_TYPE, PCF_ELEM_SKEY, PCF_ELEM_CATEGORY, PCF_ELEM_SPEC, PCF_ELEM_END1,
            PCF_ELEM_END2, PCF_ELEM_BP1, PCF_MAT_DESCR
        };

        public static IList<string> parameterPipelineAllNames = new List<string>
        {
            PCF_PIPL_AREA, PCF_PIPL_DATE, PCF_PIPL_GRAV, PCF_PIPL_INSUL, PCF_PIPL_JACKET, PCF_PIPL_MISC1, PCF_PIPL_MISC2, PCF_PIPL_MISC3,
            PCF_PIPL_MISC4, PCF_PIPL_MISC5, PCF_PIPL_NOMCLASS, PCF_PIPL_PAINT, PCF_PIPL_PREFIX, PCF_PIPL_PROJID, PCF_PIPL_REV, PCF_PIPL_SPEC,
            PCF_PIPL_TEMP, PCF_PIPL_TRACING, PCF_PIPL_TYPE
        };

        //public static IList<Guid> ParameterGUID = new List<Guid>()
        //{
        //    PCF_ELEM_CATEGORY_GUID, PCF_ELEM_BP1_GUID, PCF_ELEM_COMPID_GUID, PCF_MAT_DESCR_GUID, PCF_MAT_ID_GUID,
        //    PCF_ELEM_TYPE_GUID, PCF_ELEM_SKEY_GUID, PCF_ELEM_END1_GUID, PCF_ELEM_END2_GUID, PCF_ELEM_TAP1_GUID, PCF_ELEM_TAP2_GUID,
        //    PCF_ELEM_TAP3_GUID
        //};

        //public static IList<ParameterType> parameterTypes = new List<ParameterType>()
        //{
        //    PCF_ELEM_CATEGORY_parameterType, PCF_ELEM_BP1_parameterType, PCF_ELEM_COMPID_parameterType, PCF_MAT_DESCR_parameterType, PCF_MAT_ID_parameterType,
        //    PCF_ELEM_TYPE_parameterType, PCF_ELEM_SKEY_parameterType, PCF_ELEM_END1_parameterType, PCF_ELEM_END2_parameterType, PCF_ELEM_TAP1_parameterType,
        //    PCF_ELEM_TAP2_parameterType, PCF_ELEM_TAP3_parameterType
        //};

        //public static IList<Guid> parameterGuidPipeline = new List<Guid>()
        //{
        //    PCF_PIPL_AREA_GUID, PCF_PIPL_DATE_GUID, PCF_PIPL_GRAV_GUID, PCF_PIPL_INSUL_GUID, PCF_PIPL_JACKET_GUID, PCF_PIPL_MISC1_GUID,
        //    PCF_PIPL_MISC2_GUID, PCF_PIPL_MISC3_GUID, PCF_PIPL_MISC4_GUID, PCF_PIPL_MISC5_GUID, PCF_PIPL_NOMCLASS_GUID, PCF_PIPL_PAINT_GUID,
        //    PCF_PIPL_PREFIX_GUID, PCF_PIPL_PROJID_GUID, PCF_PIPL_REV_GUID, PCF_PIPL_SPEC_GUID, PCF_PIPL_TEMP_GUID, PCF_PIPL_TRACING_GUID,
        //    PCF_PIPL_TYPE_GUID
        //};

        //public static IList<ParameterType> parameterTypesPipeline = new List<ParameterType>()
        //{
        //    PCF_PIPL_AREA_parameterType, PCF_PIPL_DATE_parameterType, PCF_PIPL_GRAV_parameterType, PCF_PIPL_INSUL_parameterType,
        //    PCF_PIPL_JACKET_parameterType, PCF_PIPL_MISC1_parameterType, PCF_PIPL_MISC2_parameterType, PCF_PIPL_MISC3_parameterType,
        //    PCF_PIPL_MISC4_parameterType, PCF_PIPL_MISC5_parameterType, PCF_PIPL_NOMCLASS_parameterType, PCF_PIPL_PAINT_parameterType,
        //    PCF_PIPL_PREFIX_parameterType, PCF_PIPL_PROJID_parameterType, PCF_PIPL_REV_parameterType, PCF_PIPL_SPEC_parameterType,
        //    PCF_PIPL_TEMP_parameterType, PCF_PIPL_TRACING_parameterType, PCF_PIPL_TYPE_parameterType
        //};
    }
}