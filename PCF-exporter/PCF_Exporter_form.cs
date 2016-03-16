using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using BuildingCoder;
using Excel;
using Microsoft.Vbe.Interop;
using PCF_Parameters;
using PCF_Exporter;
using PCF_Functions;

namespace PCF_Exporter
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public partial class PCF_Exporter_form : System.Windows.Forms.Form
    {
        public static ExternalCommandData _commandData;
        public static UIApplication _uiapp;
        public static UIDocument _uidoc;
        public static Document _doc;
        public string _message;
        public string _excelPath = null;
        private IList<string> PCF_DATA_TABLE_NAMES = new List<string>();
        public DataSet DATA_SET = null;
        public static DataTable DATA_TABLE = null;


        public PCF_Exporter_form(ExternalCommandData cData, string message)
        {
            InitializeComponent();
            _commandData = cData;
            _uiapp = _commandData.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;
            _message = message;
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                _excelPath = openFileDialog1.FileName;
                FileStream stream = File.Open(_excelPath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                excelReader.IsFirstRowAsColumnNames = true;
                DATA_SET = excelReader.AsDataSet();
                DataTableCollection PCF_DATA_TABLES = DATA_SET.Tables;
                foreach (DataTable VARIABLE in PCF_DATA_TABLES)
                {
                    PCF_DATA_TABLE_NAMES.Add(VARIABLE.TableName);
                }
                excelReader.Close();
                comboBox1.DataSource = PCF_DATA_TABLE_NAMES;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreateParameterBindings CPB = new CreateParameterBindings();
            CPB.CreateElementBindings(_uiapp, ref _message);
            CPB.CreatePipelineBindings(_uiapp, ref _message);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DeleteParameters DP = new DeleteParameters();
            DP.ExecuteMyCommand(_uiapp, ref _message);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PopulateParameters PP = new PopulateParameters();
            PP.ExecuteMyCommand(_uiapp, ref _message, _excelPath);
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            InputVars.ExcelSheet = (string) comboBox1.SelectedItem;
            DATA_TABLE = DATA_SET.Tables[InputVars.ExcelSheet];
            InputVars.parameterNames = null;
            InputVars.parameterNames = (from dc in DATA_TABLE.Columns.Cast<DataColumn>() select dc.ColumnName).ToList();
            InputVars.parameterNames.RemoveAt(0);
            Util.InfoMsg("Following parameters will be initialized:\n"+ string.Join("\n",InputVars.parameterNames.ToArray()));
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            InputVars.SysAbbr = textBox3.Text;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true) InputVars.ExportAll = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true) InputVars.ExportAll = false;
            if (radioButton2.Checked == true)
            {
                textBox3.Visible = true;
                textBox4.Visible = true;
            }
            else
            {
                textBox3.Visible = false;
                textBox4.Visible = false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
            InputVars.OutputDirectoryFilePath = fbd.SelectedPath;
            textBox5.Text = InputVars.OutputDirectoryFilePath;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            PCFExport pcfExporter = new PCFExport();
            Result result = pcfExporter.ExecuteMyCommand(_uiapp, ref _message);
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true) InputVars.UNITS_BORE_MM = true;
            if (radioButton3.Checked == true) InputVars.UNITS_BORE_INCH = false;
            if (radioButton3.Checked == true) InputVars.UNITS_BORE = "MM";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true) InputVars.UNITS_BORE_MM = false;
            if (radioButton4.Checked == true) InputVars.UNITS_BORE_INCH = true;
            if (radioButton4.Checked == true) InputVars.UNITS_BORE = "INCH";
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked == true) InputVars.UNITS_CO_ORDS_MM = true;
            if (radioButton6.Checked == true) InputVars.UNITS_CO_ORDS_INCH = false;
            if (radioButton6.Checked == true) InputVars.UNITS_CO_ORDS = "MM";
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked == true) InputVars.UNITS_CO_ORDS_MM = false;
            if (radioButton5.Checked == true) InputVars.UNITS_CO_ORDS_INCH = true;
            if (radioButton5.Checked == true) InputVars.UNITS_CO_ORDS = "INCH";
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked == true) InputVars.UNITS_WEIGHT_KGS= true;
            if (radioButton8.Checked == true) InputVars.UNITS_WEIGHT_LBS = false;
            if (radioButton8.Checked == true) InputVars.UNITS_WEIGHT = "KGS";
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked == true) InputVars.UNITS_WEIGHT_KGS = false;
            if (radioButton7.Checked == true) InputVars.UNITS_WEIGHT_LBS = true;
            if (radioButton7.Checked == true) InputVars.UNITS_WEIGHT = "LBS";
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton11.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_METER = true;
            if (radioButton11.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_INCH = false;
            if (radioButton11.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_FEET = false;
            if (radioButton11.Checked == true) InputVars.UNITS_WEIGHT_LENGTH = "METER";
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton10.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_METER = false;
            if (radioButton10.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_INCH = true;
            if (radioButton10.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_FEET = false;
            if (radioButton10.Checked == true) InputVars.UNITS_WEIGHT_LENGTH = "INCH";
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton9.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_METER = false;
            if (radioButton9.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_INCH = false;
            if (radioButton9.Checked == true) InputVars.UNITS_WEIGHT_LENGTH_FEET = true;
            if (radioButton9.Checked == true) InputVars.UNITS_WEIGHT_LENGTH = "FEET";
        }

        
    }
}
