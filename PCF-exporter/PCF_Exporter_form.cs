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
            CPB.ExecuteMyCommand(_uiapp, ref _message);
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

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
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
    }
}
