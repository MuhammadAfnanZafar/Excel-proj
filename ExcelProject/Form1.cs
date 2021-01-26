using Microsoft.Office.Interop.Excel;
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
using Spire.Xls.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using Spire.Xls;
using System.Data.OleDb;
using ExcelProject.Model;
using System.Text.RegularExpressions;

namespace ExcelProject
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string currentlyRunningFile = "";
        string currentlyRunningFileName = "";

        string currentFile = "";
        string currentFileExt = "";
        string currentFileName = "";

        string recentFile = "";
        string recentFileExt = "";
        string recentFileName = "";


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

        public void filteration()
        {
            try
            {
                Panel childPanel = panelDropdown as Panel;
                int i = 0;
                string query = "";
                if (childPanel.Controls.Count > 0)
                {
                    foreach (Control c in childPanel.Controls)
                    {
                        if (c is ComboBox)
                        {
                            Control cc = this.Controls.Find(c.Name, true).First();
                            string text = cc.Text;
                            var NotEmpty = cc.Text.Trim();
                            var NotNone = cc.Text.ToString().Trim().ToLower();
                            var NotSelect = cc.Text.ToString().Trim().ToLower();
                            if (NotEmpty != "" && NotNone != "none" && NotSelect != "-- select --")
                            {
                                string headerText = dataGridView1.Columns[i].HeaderText;
                                //(dataGridView1.DataSource as System.Data.DataTable).DefaultView.RowFilter = string.Format("{0} = '{1}'", headerText, cc.Text);
                                //(dt.DataSource as System.Data.DataTable).DefaultView.RowFilter = string.Format("{0} = '{1}'", headerText, cc.Text);

                                // Setting query
                                query += " " + headerText + "='" + cc.Text + "' AND";
                            }
                            i++;
                        }
                    }
                    //If the last word is always AND or OR then it's pretty simple, just use a regex
                    Regex regex = new Regex("(\\s+(AND|OR)\\s*)$");
                    query = regex.Replace(query, "");
                    query = query.Trim();
                    System.Data.DataTable dtExcel = new System.Data.DataTable();
                    dtExcel = dataGridView1.DataSource as System.Data.DataTable; //read excel file  
                    dataGridView2.Visible = true;
                    dataGridView2.DataSource = dtExcel;
                    (dataGridView2.DataSource as System.Data.DataTable).DefaultView.RowFilter = query;

                    // Reading data from Excel File again for datagridview1
                    dtExcel = ReadExcel(currentlyRunningFile, currentFileExt); //read excel file  
                    dataGridView1.Visible = true;
                    dataGridView1.DataSource = dtExcel;
                    //dataGridView1.DataSource = dt.DataSource;
                }
            }
            catch (Exception ex)
            {
                label5.Show();
                label5.Text = "Error: " + ex.Message;
            }
        }

        private void btnCurrentFile_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                currentFile = filePath; // Set current File globally
                currentFileExt = fileExt;

                // Displaying currently running file
                //currentlyRunningFile = "File Name: " + file.SafeFileName + currentFileExt;
                currentFileName = file.SafeFileName;
                currentlyRunningFile = currentFile;
                currentlyRunningFileName = currentFileName;
                labelFileName.Text = currentFileName;

                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;

                        //Creating Combo Boxes
                        MyDataGridView mdv = new MyDataGridView();
                        List<string> lstCoulumnNames = mdv.getColumnNames(dataGridView1, 0);
                        MyComboBox cb = new MyComboBox();
                        cb.createComboBox(panelDropdown, dataGridView1, lstCoulumnNames);

                        // Setting Q1_1 combobox
                        //cb.setQ1_1ComboBox(dataGridView1, comboBoxQ1_1);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        private void btnRecentFile_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                recentFile = filePath; // Set current File globally
                recentFileExt = fileExt;

                // Displaying currently running file
                recentFileName = file.SafeFileName;
                currentlyRunningFile = recentFile;
                currentlyRunningFileName = recentFileName;
                labelFileName.Text = recentFileName;

                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;

                        //Creating Combo Boxes
                        MyDataGridView mdv = new MyDataGridView();
                        List<string> lstCoulumnNames = mdv.getColumnNames(dataGridView1, 0);
                        MyComboBox cb = new MyComboBox();
                        cb.createComboBox(panelDropdown, dataGridView1, lstCoulumnNames);

                        // Setting Q1_1 combobox
                        //cb.setQ1_1ComboBox(dataGridView1, comboBoxQ1_1);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        private void btnOpenCurrentFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentFile == "")
                {
                    MessageBox.Show("Kindly select file to proceed.");
                    return;
                }
                System.Data.DataTable dtExcel = new System.Data.DataTable();
                dtExcel = ReadExcel(currentFile, currentFileExt); //read excel file  
                dataGridView1.Visible = true;
                dataGridView1.DataSource = dtExcel;

                // Displaying currently running file
                currentlyRunningFile = currentFile;
                currentlyRunningFileName = currentFileName;
                labelFileName.Text = currentFileName;

                //Creating Combo Boxes
                MyDataGridView mdv = new MyDataGridView();
                List<string> lstCoulumnNames = mdv.getColumnNames(dataGridView1, 0);
                MyComboBox cb = new MyComboBox();
                cb.createComboBox(panelDropdown, dataGridView1, lstCoulumnNames);

                // Setting Q1_1 combobox
                //cb.setQ1_1ComboBox(dataGridView1, comboBoxQ1_1);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnOpenRecentFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (recentFile == "")
                {
                    MessageBox.Show("Kindly select file to proceed.");
                    return;
                }
                System.Data.DataTable dtExcel = new System.Data.DataTable();
                dtExcel = ReadExcel(recentFile, recentFileExt); //read excel file  
                dataGridView1.Visible = true;
                dataGridView1.DataSource = dtExcel;

                // Displaying currently running file
                currentlyRunningFile = recentFile;
                currentlyRunningFileName = recentFileName;
                labelFileName.Text = recentFileName;

                //Creating Combo Boxes
                MyDataGridView mdv = new MyDataGridView();
                List<string> lstCoulumnNames = mdv.getColumnNames(dataGridView1, 0);
                MyComboBox cb = new MyComboBox();
                cb.createComboBox(panelDropdown, dataGridView1, lstCoulumnNames);

                // Setting Q1_1 combobox
                //cb.setQ1_1ComboBox(dataGridView1, comboBoxQ1_1);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.panelDropdown.AutoScroll = true;
            textBox1.Hide();
            label1.Hide();
            label5.Hide();
            btnSaveChanges.Enabled = false;
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void btnSaveChanges_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                try
                {
                    //copyAlltoClipboard();
                    //Microsoft.Office.Interop.Excel.Application xlexcel;
                    //Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    //Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    //object misValue = System.Reflection.Missing.Value;
                    //xlexcel = new Excel.Application();
                    //xlexcel.Visible = true;
                    //xlWorkBook = xlexcel.Workbooks.Add(misValue);
                    //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    //Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                    //CR.Select();
                    //xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                    Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();

                    // Getting First Work Sheet
                    SpireImplementations si = new SpireImplementations();
                    WorksheetsCollection allWorkeets = si.GetAllWorksheets(currentlyRunningFile);
                    string sheetName = allWorkeets[0].Name;
                    Spire.Xls.Worksheet sheet = workbook.CreateEmptySheet(sheetName);
                    System.Data.DataTable dataTable = this.dataGridView1.DataSource as System.Data.DataTable;
                    sheet.InsertDataTable(dataTable, true, 1, 1);
                    workbook.SaveToFile(currentlyRunningFile);

                    MessageBox.Show("Saved Scuccessfully");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: ", ex.Message);
                }
            }
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.Message, "Error");
            }
            finally
            {
                GC.Collect();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Panel childPanel = panelDropdown as Panel;
                string query = "";
                int colCount = dataGridView1.Columns.Count;
                if (colCount > 0)
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        //DataGridViewRow selectedRow = dataGridView1.Rows[rowIndex];
                        //var item = selectedRow.Cells[i].Value.ToString();
                        //if (item.ToLower() == "wet")
                        //{
                        //    break;
                        //}
                        //lst.Add(item);
                        string headerText = dataGridView1.Columns[i].HeaderText;
                        if (headerText.ToLower() == "wet")
                        {
                            break;
                        }
                        query += " " + headerText + "='" + textBox1.Text + "' OR";
                    }
                    //If the last word is always AND or OR then it's pretty simple, just use a regex
                    Regex regex = new Regex("(\\s+(AND|OR)\\s*)$");
                    query = regex.Replace(query, "");
                    query = query.Trim();
                    (dataGridView1.DataSource as System.Data.DataTable).DefaultView.RowFilter = query;
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("All heading must be in CAMELCASE");
                label5.Show();
                label5.Text = "Error: " + ex.Message;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            filteration();
        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if(e.RowIndex > -1)
            //{
            //    MyDataGridView mdgv = new MyDataGridView();
            //    var getIndexesAfterWET = mdgv.getIndexesAfterWET(dataGridView1);
            //    if (getIndexesAfterWET.Count() > 0)
            //    {
            //        if (e.ColumnIndex == getIndexesAfterWET[0])
            //        {
            //            var val = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            //            textBoxValue.Text = val;

            //            var percentage = mdgv.calculatePercentage(dataGridView1, val);

            //            textBoxPercentage.Text = percentage;
            //        }
            //    }
            //}
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                var selectedVal = dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                if (selectedVal == null || selectedVal == "")
                {
                    return;
                }
                MyDataGridView mdgv = new MyDataGridView();
                var getIndexesAfterWET = mdgv.getIndexesAfterWET(dataGridView2);
                if (getIndexesAfterWET.Count() > 0)
                {
                    if (e.ColumnIndex == getIndexesAfterWET[0])
                    {
                        var val = selectedVal;
                        textBoxValue.Text = val;

                        var percentage = mdgv.calculatePercentage(dataGridView2, val);
                        SingleInstance.percentage = percentage;

                        textBoxPercentage.Text = percentage;
                    }
                }
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnChangePercentage_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBoxValue.Text == "")
                {
                    MessageBox.Show("Kindly select a value.");
                    return;
                }
                if (double.Parse(textBoxPercentage.Text) > 100)
                {
                    MessageBox.Show("Percentage must pe less then or equal to 100.");
                    return;
                }
                if (SingleInstance.percentage == textBoxPercentage.Text)
                {
                    return;
                }

                int i = 0;
                bool flag = false;
                MyDataGridView mdgv = new MyDataGridView();
                DataGridView tmpDataGridView = dataGridView2;
                var getIndexesAfterWET = mdgv.getIndexesAfterWET(tmpDataGridView);
                if (getIndexesAfterWET.Count() > 0)
                {
                    for (int rows = 0; rows < tmpDataGridView.Rows.Count - 1; rows++)
                    {
                        for (int col = 0; col < tmpDataGridView.Rows[rows].Cells.Count; col++)
                        {
                            if (i == getIndexesAfterWET[0])
                            {
                                string value = tmpDataGridView.Rows[rows].Cells[col].Value.ToString();
                                if (value != textBoxValue.Text)
                                {
                                    var currentValue = value;
                                    tmpDataGridView.Rows[rows].Cells[col].Value = textBoxValue.Text;
                                    var percentage = mdgv.calculatePercentage(tmpDataGridView, textBoxValue.Text);
                                    if (double.Parse(textBoxPercentage.Text) <= double.Parse(percentage))
                                    {
                                        flag = true;
                                        textBoxPercentage.Text = percentage;
                                        SingleInstance.percentage = percentage;
                                        break;
                                    }
                                    //else
                                    //{
                                    //    tmpDataGridView.Rows[rows].Cells[col].Value = currentValue;
                                    //}
                                }
                            }
                            i++;
                        }
                        if (flag)
                        {
                            break;
                        }
                        i = 0;
                    }
                }
                if (!flag)
                {
                    MessageBox.Show("Not Possible");
                }
                dataGridView2.DataSource = tmpDataGridView.DataSource;
            }
            catch (Exception ex)
            {
                label5.Show();
                label5.Text = "Error: " + ex.Message;
            }
        }

        private void btnSaveToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                Panel childPanel = panelDropdown as Panel;
                int i = 0;
                string filters = "";
                if (childPanel.Controls.Count > 0)
                {
                    foreach (Control c in childPanel.Controls)
                    {
                        if (c is ComboBox)
                        {
                            Control cc = this.Controls.Find(c.Name, true).First();
                            string text = cc.Text;
                            var NotEmpty = cc.Text.Trim();
                            var NotNone = cc.Text.ToString().Trim().ToLower();
                            var NotSelect = cc.Text.ToString().Trim().ToLower();
                            if (NotEmpty != "" && NotNone != "none" && NotSelect != "-- select --")
                            {
                                string headerText = dataGridView1.Columns[i].HeaderText;

                                // Setting query
                                filters += headerText + "+";
                            }
                            i++;
                        }
                    }

                    filters = filters.Trim('+');

                    // Getting Q1_1 header options
                    MyDataGridView mdgv = new MyDataGridView();
                    var getIndexesAfterWET = mdgv.getIndexesAfterWET(dataGridView1);
                    var q1_1_index = getIndexesAfterWET[0];
                    var colData = mdgv.getColumnData(dataGridView2, q1_1_index);
                    colData = colData.Distinct().ToList();

                    // Calculating Percentage of each option/data
                    List<SaveToExcel> lstPercentage = new List<SaveToExcel>();
                    if (colData.Count() > 0)
                    {
                        for (int j = 0; j < colData.Count(); j++)
                        {
                            var percentage = mdgv.calculatePercentage(dataGridView2, colData[j]);
                            SaveToExcel saveToExcel = new SaveToExcel();
                            saveToExcel.optionName = colData[j];
                            saveToExcel.percentage = percentage;
                            lstPercentage.Add(saveToExcel);
                        }
                    }

                    // Saving Data to excel
                    string filePath = string.Empty;
                    string fileExt = string.Empty;
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    saveDialog.FilterIndex = 2;
                    if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        filePath = saveDialog.FileName; //get the path of the file  
                        fileExt = Path.GetExtension(filePath);

                        //Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();

                        //// Getting First Work Sheet
                        //SpireImplementations si = new SpireImplementations();
                        //WorksheetsCollection allWorkeets = si.GetAllWorksheets(currentlyRunningFile);
                        //string sheetName = allWorkeets[0].Name;
                        //Spire.Xls.Worksheet sheet = workbook.CreateEmptySheet(sheetName);
                        //System.Data.DataTable dataTable = this.dataGridView1.DataSource as System.Data.DataTable;
                        //sheet.InsertDataTable(dataTable, true, 1, 1);
                        //workbook.SaveToFile(currentlyRunningFile);

                        //dataGridView2.DataSource = null;
                        //for (int k = 0; k < lstPercentage.Count - 1; k++)
                        //{
                        //    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        //    {
                        //        dataGridView2.Rows[k].Cells[j].Value = lstPercentage[k];
                        //    }
                        //}
                        
                        MessageBox.Show("Export Successful");
                    }
                }
            }
            catch (Exception ex)
            {
                label5.Show();
                label5.Text = "Error: " + ex.Message;
            }
        }
    }
}
