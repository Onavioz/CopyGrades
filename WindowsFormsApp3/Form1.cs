using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private Dictionary<string, string> dict = new Dictionary<string, string>();
        private DataTable dt = new DataTable();
        private string targetName;
        private string targetTable;
        private string targetCodeCourse;
        private bool AnotherFlag;

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog()
                { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm", ValidateNames = true })
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        FileStream fileStream = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read);
                        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
                        DataSet result = excelDataReader.AsDataSet();
                        excelDataReader.Close();

                        dt = result.Tables["Worksheet"];
                        var count = dt.Rows.Count;
                        var too = dt.Select();
                        for (int i = 5; i < count; i++)
                        {
                            var id = too[i][0].ToString();
                            var grade = too[i][2].ToString();
                            dict.Add(id, grade);
                        }
                        button2.Enabled = true;
                        checkBox3.Checked = true;
                        button1.Enabled = false;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("ERRORs possible:\n1. incorrect file.\n2.file already opened. please close it before.");
                button2.Enabled = false;
                AnotherFlag = false;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog()
                { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm", ValidateNames = true })
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        targetName = dialog.FileName;
                        FileStream fileStream = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read);
                        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateBinaryReader(fileStream);
                        DataSet result = excelDataReader.AsDataSet();
                        excelDataReader.Close();
                        dt = result.Tables[targetTable];
                        if (dt is null)
                        {
                            MessageBox.Show("ERROR incorrect table name, please try again.");
                            button2.Enabled = false;
                            AnotherFlag = true;
                            textBox1.Clear();
                            checkBox2.Checked = false;
                        }
                        else
                        {
                            Cursor.Current = Cursors.WaitCursor;
                            HSSFWorkbook templateWorkbook;
                            HSSFSheet sheet;
                            HSSFRow dataRow;

                            var count = dt.Rows.Count;
                            var too = dt.Select();
                            using (var fs = new FileStream(targetName, FileMode.Open, FileAccess.Read))
                            {
                                templateWorkbook = new HSSFWorkbook(fs,true);
                                sheet = (HSSFSheet)templateWorkbook.GetSheet(targetTable);

                                for (int i = 0; i < count; i++)
                                {
                                    var id = "0" + too[i][26].ToString();
                                    dataRow = (HSSFRow)sheet.GetRow(i);
                                    if (dict.ContainsKey(id))
                                    {
                                        var val = dict.Where(pair => pair.Key == id)
                                        .Select(pair => pair.Value)
                                        .FirstOrDefault(); 
                                        dataRow.Cells[19].SetCellValue(val);
                                    }
                                    if (i ==12 && !dataRow.Cells[11].ToString().Contains(targetCodeCourse))
                                    {
                                        MessageBox.Show("ERROR The code of the course is not matched to the course in the table.\n" +
                                            "please check and write correct code of the course or correct name of table");
                                        targetTable = "";
                                        targetCodeCourse = "";
                                        textBox1.Clear();
                                        textBox2.Clear();
                                        button2.Enabled = false;
                                        checkBox1.Checked = false;
                                        checkBox2.Checked = false;
                                        AnotherFlag = true;
                                        return;
                                    }  
                                }
                            }
                            using (var fs = new FileStream(targetName, FileMode.Open, FileAccess.Write))
                            {
                                templateWorkbook.Write(fs);
                            }
                            Cursor.Current = Cursors.Default;
                            MessageBox.Show("Copy grades successfully!");
                            button2.Enabled = false;
                            targetTable = "";
                            targetCodeCourse = "";
                            checkBox1.Checked = false;
                            checkBox2.Checked = false;
                            AnotherFlag = true;
                            textBox1.Clear();
                            textBox2.Clear();
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("ERROR please close the excel file you're tring to work with");
            }
        }
        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(""))
                MessageBox.Show("ERROR please write the table name.");
            else if (textBox1.Text.Length == 1 && !AnotherFlag)
            {
                targetTable = textBox1.Text[0].ToString().ToUpper();
                button1.Enabled = true;
                checkBox2.Checked = true;
            }
            else if(!AnotherFlag)
            {
                targetTable = textBox1.Text[0].ToString().ToUpper() + textBox1.Text.Substring(1);
                button1.Enabled = true;
                checkBox2.Checked = true;
            }
            else if (AnotherFlag)
            {
                targetTable = textBox1.Text[0].ToString().ToUpper() + textBox1.Text.Substring(1);
                button2.Enabled = true;
                checkBox2.Checked = true;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (textBox2.Text.Equals(""))
                MessageBox.Show("ERROR please write the code of the course.");
            else
            {
                try
                {
                    targetCodeCourse =int.Parse(textBox2.Text).ToString();
                    checkBox1.Checked = true;
                }
                catch (Exception)
                {
                    MessageBox.Show("ERROR please write numbers only.");
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
        }
    }
}
