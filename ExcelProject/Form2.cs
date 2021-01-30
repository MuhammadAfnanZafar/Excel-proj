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
using System.Text.RegularExpressions;
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
        string currentFile = "";
        string currentFileExt = "";
        string currentFileName = "";


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
        public void WriteExcel()
        {
            try
            {
                ////Working Code of add data to excel file
                string filePath = @"C:\Users\MuhammadNauman\Desktop\abc.xlsx";
                string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                        filePath +
                        ";Extended Properties='Excel 12.0 XML;HDR=Yes;';"; //for above excel 2007  
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                string sql = null;
                MyConnection = new System.Data.OleDb.OleDbConnection(conn);
                MyConnection.Open();
                myCommand.Connection = MyConnection;
                sql = "Insert into [Sheet1$] values('5','e')";
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                MyConnection.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void filteration(string query, DataGridView dataGridView)
        {
            try
            {

                if (query != "")
                {
                    Panel childPanel = panelDropdown as Panel;
                    int i = 0;
                    if (childPanel.Controls.Count > 0)
                    {

                        //If the last word is always AND or OR then it's pretty simple, just use a regex
                        Regex regex = new Regex("(\\s+(AND|OR)\\s*)$");
                        query = regex.Replace(query, "");
                        query = query.Trim();
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = dataGridView.DataSource as System.Data.DataTable; //read excel file 
                        dataGridView3.Refresh();
                        dataGridView3.Visible = true;
                        dataGridView3.DataSource = dtExcel;
                        (dataGridView3.DataSource as System.Data.DataTable).DefaultView.RowFilter = query;

                        // Reading data from Excel File again for datagridview1
                        dtExcel = ReadExcel(currentFile, currentFileExt); //read excel file  
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;
                        //dataGridView1.DataSource = dt.DataSource;

                        i = 0;
                        // Removing previous rows
                        var toBeDeleted = new List<DataGridViewRow>();
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value == null || row.Cells[0].Value == DBNull.Value || String.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                            {
                                break;
                            }
                            string value1 = row.Cells[0].Value.ToString();
                            if (value1 == "")
                            {
                                break;
                            }
                            if (value1.ToLower() == "p")
                            {
                                i++;
                                //processing data
                                toBeDeleted.Add(row);
                            }
                        }
                        toBeDeleted.ForEach(d => dataGridView1.Rows.Remove(d));

                        // dataGridView2
                        System.Data.DataTable dtExcel2 = new System.Data.DataTable();
                        dtExcel2 = ReadExcel(currentFile, currentFileExt); //read excel file  
                        dataGridView2.DataSource = dtExcel2;
                        //// Removing current rows
                        toBeDeleted = new List<DataGridViewRow>();
                        foreach (DataGridViewRow row in dataGridView2.Rows)
                        {
                            if (row.Cells[0].Value == null || row.Cells[0].Value == DBNull.Value || String.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                            {
                                break;
                            }
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

                    }
                }
            }
            catch (Exception ex)
            {
                label5.Show();
                label5.Text = "Error: " + ex.Message;
                MessageBox.Show(ex.Message);
            }
        }


        void reportFormatDT(List<MyListBox> lst, DataGridView dataGridView3, DataGridView dataGridView, DataGridView reportDataGridView, ComboBox cbWorkingColumn, int indexInsertData, string reportType)
        {

            if (lst.Count() > 0)
            {

                dataGridView3.Refresh();
                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();
                dataGridView3.Columns.Clear();
                Percentages.percentagesList = new List<List<Percentages>>();

                // Getting all possible combinations
                Dictionary<int, List<string>> tags = new Dictionary<int, List<string>>();
                for (int j = 0; j < lst.Count(); j++)
                {
                    tags.Add(j, lst[j].Data);
                }
                NListBuilder nListBuilder = new NListBuilder(tags, "AND");
                var AllCombosList = nListBuilder.AllCombos;

                //string query = "";
                //int u = 0;
                foreach (var query in AllCombosList)
                {
                    List<Percentages> percentagesList = new List<Percentages>();
                    // Filtration
                    filteration(query, dataGridView);

                    var selectedVal = cbWorkingColumn.Text;
                    MyDataGridView mdgv = new MyDataGridView();
                    var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, selectedVal);  // Getting selected working column index
                    var workingColumnData = mdgv.getColumnPlusRowData(dataGridView3, searchColumnNameIndexAfterWET, Convert.ToInt32(tbDataChar.Text)); // Getting selected working column data
                    workingColumnData = workingColumnData.Distinct().ToList();
                    if (searchColumnNameIndexAfterWET == 0)
                    {
                        MessageBox.Show("Kindly add atleast one column after WET column.");
                        return;
                    }

                    decimal total = 0;
                    if (searchColumnNameIndexAfterWET > 0)
                    {
                        if (workingColumnData.Count() > 0)
                        {
                            foreach (var item in workingColumnData)
                            {
                                // Calulating Percentage
                                var percentage = mdgv.calculatePercentage(dataGridView3, item, searchColumnNameIndexAfterWET); // Getting percentage of each column in selected working column data
                                total = total + Convert.ToDecimal(percentage);
                                Percentages per = new Percentages()
                                {
                                    ColumnValue = item,
                                    Percentage = percentage,
                                    Query = query,
                                    Sum = total
                                };


                                percentagesList.Add(per);
                            }
                        }
                        else
                        {
                            Percentages per = new Percentages()
                            {
                                Query = query,
                            };
                            percentagesList.Add(per);
                        }
                    }
                    Percentages.percentagesList.Add(percentagesList);
                }

                //filteration(query);
                //filteration(query);

            }

            // Export all percentages to datagridView3
            if (Percentages.percentagesList.Count > 0)
            {
                var selectedVal = cbWorkingColumn.Text;
                MyDataGridView mdgv = new MyDataGridView();
                //var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView1, selectedVal);  // Getting selected working column index
                //var workingColumnData = mdgv.getColumnPlusRowData(dataGridView3, searchColumnNameIndexAfterWET, Convert.ToInt32(tbDataChar.Text)); // Getting selected working column data
                //workingColumnData = workingColumnData.Distinct().ToList();

                var combinePerCetagesList = new Percentages().combinePercentagesList(Percentages.percentagesList);
                combinePerCetagesList = combinePerCetagesList.Distinct().ToList();

                // Setting up structure// If rows already created
                DataTable dt = new DataTable();
                //Setting Columns
                if (indexInsertData == 0) // Checking if clumn already exists or in short PERVIOUS
                {
                    dt.Columns.Add(new DataColumn(selectedVal, typeof(string)));
                    for (int r = 0; r < Percentages.percentagesList.Count; r++)
                    {
                        var item = Percentages.percentagesList[r];
                        var queryAsColumn = item[0].Query;
                        dt.Columns.Add(new DataColumn(queryAsColumn, typeof(string)));
                    }
                }
                DataRow toInsert = dt.NewRow();
                //// Setting Rows
                for (int j = 0; j < combinePerCetagesList.Count; j++)
                {
                    toInsert = dt.NewRow();

                    toInsert[0] = combinePerCetagesList[j].ColumnValue;
                    dt.Rows.InsertAt(toInsert, 0);
                }


                //toInsert[0] = "Pervious";
                //dt.Rows.InsertAt(toInsert, 0);
                reportDataGridView.DataSource = dt;

                //
                for (int r = 0; r < Percentages.percentagesList.Count; r++)
                {
                    var item = Percentages.percentagesList[r]; // Item
                    for (int c = 0; c < item.Count; c++)
                    {
                        var colItem = item[c]; // Item
                        int searchedRowIndexOfQ1_X = 0;
                        var getQ1_X_Data = mdgv.getColumnData(reportDataGridView, 0);
                        if (colItem.ColumnValue != null)
                        {
                            searchedRowIndexOfQ1_X = getQ1_X_Data.FindIndex(x => x == colItem.ColumnValue);
                            reportDataGridView.Rows[searchedRowIndexOfQ1_X].Cells[r + 1].Value = colItem.Percentage;
                        }
                    }
                }


                // Save to Excel
                myExcel excel = new myExcel();
                string title = reportType + " Report";
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Documents (*.xls)|*.xls|(*.xlsx)|*.xlsx";
                sfd.FileName = reportType + ".xls";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    excel.ToCsV(reportDataGridView, reportType + " Report", "", "", title, sfd.FileName);
                    MessageBox.Show("Finish");
                }
            }

        }

        public void isRelational()
        {

            //try
            //{
                // Getting all list boxes from panel
                Panel childPanel = panelDropdown as Panel;
                List<MyListBox> lst = new List<MyListBox>();
                int i = 2;
                MyListBox obj = new MyListBox();
                if (childPanel.Controls.Count > 0)
                {
                    foreach (Control c in childPanel.Controls)
                    {
                        if (i % 2 == 0) // For label Name
                        {
                            obj = new MyListBox();
                        }
                        i++;

                        if (c is Label)
                        {
                            Control cc = this.Controls.Find(c.Name, true).First();
                            obj.ColumnName = cc.Text;
                        }
                        if (c is ListBox)
                        {
                            Control cc = this.Controls.Find(c.Name, true).First();
                            var tmpListBox = cc as ListBox;
                            var tmpListBoxItems = tmpListBox.SelectedItems;
                            if (tmpListBoxItems.Count > 0)
                            {
                                //MyListBox obj = new MyListBox();
                                foreach (var item in tmpListBoxItems)
                                {
                                    obj.Data.Add(string.Format("{0}='{1}'", obj.ColumnName, item.ToString()));
                                }
                                lst.Add(obj);
                            }
                        }
                    }
                }
                // Current Excel Data
                reportFormatDT(lst, dataGridView3, dataGridView1, dataGridView4, cbWorkingColumn, 0, "Current");
                reportFormatDT(lst, dataGridView3, dataGridView2, dataGridView5, cbWorkingColumn, 0, "Previous");

            //}
            //catch (Exception ex)
            //{
            //    label5.Show();
            //    label5.Text = "Error: " + ex.Message;
            //    MessageBox.Show(ex.Message);
            //}
        }
        
        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (rbRelational.Checked == true)
            {
                if (cbWorkingColumn.Text == "-- Select --" || tbDataChar.Text == null || tbDataChar.Text == "" || Convert.ToInt32(tbDataChar.Text) == 0)
                {
                    MessageBox.Show("Working column and Data Charater field must not be empty.");
                    return;
                }
                isRelational();
            }
            else if (rbNonRelational.Checked == true)
            {

            }
            else
            {

            }
            //MessageBox.Show("Are You Sure?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
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
                currentFile = filePath; // Set current File globally
                currentFileExt = fileExt;
                currentFileName = file.SafeFileName;
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
                        //var singleColumnNames = mdv.getSingleColumnNames(dataGridView1);
                        //var multiColumnNames = mdv.getMultiColumnNames(dataGridView1);

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

                        int i = 0;
                        // Removing previous rows
                        var toBeDeleted = new List<DataGridViewRow>();
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value == null || row.Cells[0].Value == DBNull.Value || String.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                            {
                                break;
                            }
                            string value1 = row.Cells[0].Value.ToString();
                            if (value1 == "")
                            {
                                break;
                            }
                            if (value1.ToLower() == "p")
                            {
                                i++;
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
                            if (row.Cells[0].Value == null || row.Cells[0].Value == DBNull.Value || String.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                            {
                                break;
                            }
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


        private void button5_Click(object sender, EventArgs e)
        {
            //WriteExcel();
            myExcel excel = new myExcel();
            string title = "Excel Report";
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls|(*.xlsx)|*.xlsx";
            sfd.FileName = "report.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                excel.ToCsV(dataGridView1, "Report", "Current", "Karachi", title, sfd.FileName);
                MessageBox.Show("Finish");
            }
        }

        private void button4_Click(object sender, EventArgs e)
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

                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView3.Visible = true;
                        dataGridView3.DataSource = dtExcel;

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

        private void tbDataChar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
    }
}
