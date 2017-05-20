using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms.ComponentModel.Com2Interop;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using BuildingCoder;
using MoreLinq;
using PCF_Functions;
using iv = NTR_Functions.InputVars;

namespace NTR_Functions
{
    public static class InputVars
    {
        //Scope control
        public static bool ExportAllOneFile = false;
        public static bool ExportAllSepFiles = false;
        public static bool ExportSpecificPipeLine = false;
        public static bool ExportSelection = false;
        public static double DiameterLimit = 0;

        //File control
        public static string OutputDirectoryFilePath = @"C:\";
        public static string ExcelPath = @"C:\";

        //Current SystemAbbreviation
        public static string SysAbbr = null;
    }

    public class ConfigurationData
    {
        public StringBuilder _01_GEN { get; }
        public StringBuilder _02_AUFT { get; }
        public StringBuilder _03_TEXT { get; }
        public StringBuilder _04_LAST { get; }
        public StringBuilder _05_DN { get; }
        public StringBuilder _06_ISO { get; }
        public DataTable Pipelines { get; }
        public DataTable Elements { get; }

        public ConfigurationData()
        {
            DataSet dataSet = DataHandler.ImportExcelToDataSet(iv.ExcelPath, "NO");

            DataTableCollection dataTableCollection = dataSet.Tables;

            _01_GEN = ReadNtrConfigurationData(dataTableCollection, "GEN", "C General settings");
            _02_AUFT = ReadNtrConfigurationData(dataTableCollection, "AUFT", "C Project description");
            _03_TEXT = ReadNtrConfigurationData(dataTableCollection, "TEXT", "C User text");
            _04_LAST = ReadNtrConfigurationData(dataTableCollection, "LAST", "C Loads definition");
            _05_DN = ReadNtrConfigurationData(dataTableCollection, "DN", "C Definition of pipe dimensions");
            _06_ISO = ReadNtrConfigurationData(dataTableCollection, "IS", "C Definition of insulation type");

            DataSet dataSetWithHeaders = DataHandler.ImportExcelToDataSet(iv.ExcelPath, "YES");
            Pipelines = ReadDataTable(dataSetWithHeaders.Tables, "PIPELINES");
            Elements = ReadDataTable(dataSetWithHeaders.Tables, "ELEMENTS");

            //http://stackoverflow.com/questions/10855/linq-query-on-a-datatable?rq=1
        }

        /// <summary>
        /// Selects a DataTable by name and creates a StringBuilder output to NTR format based on the data in table.
        /// </summary>
        /// <param name="dataTableCollection">A collection of datatables.</param>
        /// <param name="tableName">The name of the DataTable to process.</param>
        /// <returns>StringBuilder containing the output NTR data.</returns>
        private static StringBuilder ReadNtrConfigurationData(DataTableCollection dataTableCollection, string tableName, string description)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(description);

            var table = ReadDataTable(dataTableCollection, tableName);
            if (table == null)
            {
                sb.AppendLine("C " + tableName + " does not exist!");
                return sb;
            }

            int numberOfRows = table.Rows.Count;
            if (numberOfRows.IsOdd())
            {
                sb.AppendLine("C " + tableName + " is malformed, contains odd number of rows, must be even");
                return sb;
            }

            for (int i = 0; i < numberOfRows / 2; i++)
            {
                DataRow headerRow = table.Rows[i * 2];
                DataRow dataRow = table.Rows[i * 2 + 1];
                if (headerRow == null || dataRow == null)
                    throw new NullReferenceException(
                        tableName + " does not have two rows, check EXCEL configuration sheet!");

                sb.Append(tableName);

                for (int j = 0; j < headerRow.ItemArray.Length; j++)
                {
                    sb.Append(" ");
                    sb.Append(headerRow.Field<string>(j));
                    sb.Append("=");
                    sb.Append(dataRow.Field<string>(j));
                }

                sb.AppendLine();
            }

            return sb;
        }

        private static DataTable ReadDataTable(DataTableCollection dataTableCollection, string tableName)
        {
            var table = (from DataTable dtbl in dataTableCollection where dtbl.TableName == tableName select dtbl)
                .FirstOrDefault();
            return table;
        }
    }

    public static class DataWriter
    {
        public static string PointCoords(string p, Connector c)
        {
            return " " + p + "=" + NtrConversion.PointStringMm(c.Origin);
        }

        public static string PointCoords(string p, Element element)
        {
            return " " + p + "=" + NtrConversion.PointStringMm(((LocationPoint)element.Location).Point);
        }

        public static string DnWriter(Element element)
        {
            double dia = 0;

            if (element is Pipe pipe)
            {
                //Get connector set for the pipes
                ConnectorSet connectorSet = pipe.ConnectorManager.Connectors;
                //Filter out non-end types of connectors
                Connector con = (from Connector connector in connectorSet
                                 where connector.ConnectorType.ToString().Equals("End")
                                 select connector).FirstOrDefault();
                dia = con.Radius * 2;
            }
            else if (element is FamilyInstance fis)
            {
                //TODO: Fix FamilyInstance case, maybe not
                return "NTR_Functions DataWriter DnWriter FamilyInstance case";
            }
            return " DN=DN" + dia.FeetToMm().Round(0);
        }

        public static string DnWriter(string p, Connector con)
        {
            double dia = con.Radius * 2;
            return " " + p + "=DN" + dia.FeetToMm().Round(0);
        }

        public static string ReadParameterFromDataTable(string key, DataTable table, string parameter)
        {
            if (!(table.AsEnumerable().Any(row => row.Field<string>(0) == key))) return null;

            var query = from row in table.AsEnumerable()
                        where row.Field<string>(0) == key
                        select row.Field<string>(parameter);
            string value = query.FirstOrDefault();
            if (value == null)
                throw new Exception("There was no definition for " + parameter + " parameter for pipeline " + key);
            return " " + parameter + "=" + value;
        }

        public static string ReadElementTypeFromDataTable(string key, DataTable table, string parameter)
        {
            if (!(table.AsEnumerable().Any(row => row.Field<string>(0) == key))) return null;

            var query = from row in table.AsEnumerable()
                where row.Field<string>(0) == key
                select row.Field<string>(parameter);
            string value = query.FirstOrDefault();
            if (value == null)
                throw new Exception("There was no definition for " + parameter + " parameter for pipeline " + key);
            return value;
        }

        public static string WriteElementId(Element element, string parameter)
        {
            return " " + parameter + "=" + element.Id.IntegerValue;
        }
    }

    public class NtrConversion
    {
        const double _inch_to_mm = 25.4;
        const double _foot_to_mm = 12 * _inch_to_mm;
        const double _foot_to_inch = 12;

        /// <summary>
        /// Return a string for a real number.
        /// </summary>
        private static string RealString(double a)
        {
            //return a.ToString("0.##");
            //return (Math.Truncate(a * 100) / 100).ToString("0.00", CultureInfo.GetCultureInfo("en-GB"));
            return Math.Round(a, 1, MidpointRounding.AwayFromZero).ToString("0.0", CultureInfo.GetCultureInfo("en-GB"));
        }

        /// <summary>
        /// Return a string for an XYZ point or vector with its coordinates converted from feet to millimetres.
        /// </summary>
        public static string PointStringMm(XYZ p)
        {
            return string.Format("'{0:0.0}, {1:0.0}, {2:0.0}'",
                RealString(p.X * _foot_to_mm),
                RealString(p.Y * _foot_to_mm),
                RealString(p.Z * _foot_to_mm));
        }
    }

    public static class NTR_Filter
    {
        /// <summary>
        /// Tests the diameter of the pipe or primary connector of element against the diameter limit set in the interface.
        /// </summary>
        /// <param name="passedElement"></param>
        /// <returns>True if diameter is larger than limit and false if smaller.</returns>
        public static bool FilterDiameterLimit(Element element)
        {
            double diameterLimit = iv.DiameterLimit;
            bool diameterLimitBool = true;
            double testedDiameter = 0;
            switch (element)
            {
                case MEPCurve pipe:
                    testedDiameter = pipe.Diameter.FeetToMm().Round(0);
                    if (testedDiameter <= diameterLimit) diameterLimitBool = false;
                    break;

                case FamilyInstance inst:
                    //MEPModel of the elements is accessed
                    MEPModel mepmodel = inst.MEPModel;
                    //Get connector set for the element
                    ConnectorSet connectorSet = mepmodel.ConnectorManager.Connectors;
                    //Declare a variable for 
                    Connector testedConnector = null;

                    if (connectorSet.IsEmpty) break;
                    else if (connectorSet.Size == 1) foreach (Connector connector in connectorSet) testedConnector = connector;
                    else testedConnector = (from Connector connector in connectorSet
                                            where connector.GetMEPConnectorInfo().IsPrimary
                                            select connector).FirstOrDefault();

                    testedDiameter = (testedConnector.Radius * 2).FeetToMm().Round(0);

                    if (testedDiameter <= diameterLimit) diameterLimitBool = false;

                    break;
            }
            return diameterLimitBool;
        }
    }

    public static class NTR_Utils
    {
        public static (Connector Primary, Connector Secondary, Connector Tertiary) GetConnectors(Element element)
        {
            ConnectorManager cmgr = MepUtils.GetConnectorManager(element);
            //Sort connectors to primary, secondary and none
            Connector primCon = null; Connector secCon = null; Connector tertCon = null;

            foreach (Connector connector in cmgr.Connectors)
            {
                if (connector.GetMEPConnectorInfo().IsPrimary) primCon = connector;
                else if (connector.GetMEPConnectorInfo().IsSecondary) secCon = connector;
                else if ((connector.GetMEPConnectorInfo().IsPrimary == false) && (connector.GetMEPConnectorInfo().IsSecondary == false))
                    tertCon = connector;
            }

            return (primCon, secCon, tertCon);
        }
    }
}
