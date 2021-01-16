using ExcelProject.Model;
using Spire.Xls.Collections;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }


        public System.Data.DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                        fileName +
                        ";Extended Properties='Excel 12.0 XML;HDR=Yes;';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    SpireImplementations si = new SpireImplementations();
                    WorksheetsCollection allWorkeets = si.GetAllWorksheets(fileName);
                    string sheetName = allWorkeets[0].Name;

                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + sheetName + "$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return dtexcel;
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Are You Sure?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                lblFileName.Text = file.SafeFileName; 
                // Displaying currently running file
                //currentlyRunningFile = "File Name: " + file.SafeFileName + currentFileExt;

                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {

                        // dataGridView1
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView1.Visible = true;
                        
                        dataGridView1.DataSource = dtExcel;


                        //Creating Combo Boxes
                        MyDataGridView mdv = new MyDataGridView();
                        List<string> lstCoulumnNames = mdv.getColumnNames(dataGridView1, 0);
                        var singleColumnNames = mdv.getSingleColumnNames(dataGridView1);
                        var multiColumnNames = mdv.getMultiColumnNames(dataGridView1);

                        MyListBox lb = new MyListBox();
                        lb.createListBox(panelDropdown, dataGridView1, lstCoulumnNames);

                        // Assigning values to Working Column
                        lstCoulumnNames = mdv.getColumnNamesAfterWET(dataGridView1);
                        MyComboBox cb = new MyComboBox();
                        cb.assignHeaderNameAfterWETToComboBox(dataGridView1, cbWorkingColumn, lstCoulumnNames);

                        // Assign values to Dependent Column
                        lb.assignHeaderNameAfterWETToListBox(dataGridView1, lbDepCol, lstCoulumnNames);

                        // Assign values to Must Column
                        lb.assignHeaderNameAfterWETToListBox(dataGridView1, lbMustCol, lstCoulumnNames);

                        // Removing previous rows
                        var toBeDeleted = new List<DataGridViewRow>();
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            string value1 = row.Cells[0].Value.ToString();
                            if (value1 == "")
                            {
                                break;
                            }
                            if (value1.ToLower() == "p")
                            {
                                //processing data
                                toBeDeleted.Add(row);
                            }
                        }
                        toBeDeleted.ForEach(d => dataGridView1.Rows.Remove(d));

                        // dataGridView2
                        System.Data.DataTable dtExcel2 = new System.Data.DataTable();
                        dtExcel2 = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView2.DataSource = dtExcel2;
                        //// Removing current rows
                        toBeDeleted = new List<DataGridViewRow>();
                        foreach (DataGridViewRow row in dataGridView2.Rows)
                        {
                            string value1 = row.Cells[0].Value.ToString();
                            if (value1 == "")
                            {
                                break;
                            }
                            if (value1.ToLower() == "c")
                            {
                                //processing data
                                toBeDeleted.Add(row);
                            }
                        }
                        toBeDeleted.ForEach(d => dataGridView2.Rows.Remove(d));
                                               
                        // Setting Q1_1 combobox
                        //cb.setQ1_1ComboBox(dataGridView1, comboBoxQ1_1);
                    }
                    catch (Exception ex)
                    {
                        lblError.Text = ex.Message;
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        void testFunc()
        {
            int i = 1;
            // Removing previous rows
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string value1 = row.Cells[0].Value.ToString();
                if (value1.ToLower() == "p" || value1 == "")
                {
                    dataGridView1.Rows.Remove(row);
                }
            }

            // Removing current rows
            i = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                string value1 = row.Cells[0].Value.ToString();
                if (value1.ToLower() == "c" || value1 == "")
                {
                    dataGridView2.Rows.Remove(row);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            testFunc();
        }
    }
}
