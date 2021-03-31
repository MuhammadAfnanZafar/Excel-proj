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
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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

        void resetDatatable(DataGridView dgv)
        {
            dgv.DataSource = null;
            dgv.Rows.Clear();
            dgv.Columns.Clear();
            dgv.Refresh();
        }
        private string MyDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
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
        void UploadRangeFile()
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  


                // Displaying currently running file
                //currentlyRunningFile = "File Name: " + file.SafeFileName + currentFileExt;

                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dgvRange.Visible = true;
                        dgvRange.DataSource = dtExcel;

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
        public void filterationOnChangePercentage(string query, DataGridView dataGridView)
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
                        //System.Data.DataTable dtExcel = new System.Data.DataTable();
                        //dtExcel = ReadExcel(currentFile, currentFileExt); //read excel file 

                        dataGridView3.Refresh();
                        dataGridView3.Visible = true;
                        //dataGridView3.DataSource = dtExcel;
                        dataGridView3.DataSource = new MyDataGridView().CopyDataGridView(dataGridView1);

                        //
                        //var toBeDeleted = new List<DataGridViewRow>();
                        //foreach (DataGridViewRow row in dataGridView3.Rows)
                        //{
                        //    if (row.Cells[0].Value == null || row.Cells[0].Value == DBNull.Value || String.IsNullOrWhiteSpace(row.Cells[0].Value.ToString()))
                        //    {
                        //        break;
                        //    }
                        //    string value1 = row.Cells[0].Value.ToString();
                        //    if (value1 == "")
                        //    {
                        //        break;
                        //    }
                        //    if (value1.ToLower() == "p")
                        //    {
                        //        i++;
                        //        //processing data
                        //        toBeDeleted.Add(row);
                        //    }
                        //}
                        //toBeDeleted.ForEach(d => dataGridView3.Rows.Remove(d));
                        (dataGridView3.DataSource as System.Data.DataTable).DefaultView.RowFilter = query;

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

        public void filterationForLastCurrentFile(string query, DataGridView dataGridView)
        {
            try
            {

                if (query != "")
                {
                    int i = 0;
                    //If the last word is always AND or OR then it's pretty simple, just use a regex
                    Regex regex = new Regex("(\\s+(AND|OR)\\s*)$");
                    query = regex.Replace(query, "");
                    query = query.Trim();

                    //dataGridView3.Refresh();
                    //dataGridView3.Visible = true;
                    //dataGridView3.DataSource = dataGridView1.DataSource;
                    dataGridView3.Visible = true;
                    //dataGridView3.DataSource = dtExcel;
                    dataGridView3.DataSource = new MyDataGridView().CopyDataGridView(dataGridView1);

                    //
                    var toBeDeleted = new List<DataGridViewRow>();
                    foreach (DataGridViewRow row in dataGridView3.Rows)
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
                    toBeDeleted.ForEach(d => dataGridView3.Rows.Remove(d));
                    (dataGridView3.DataSource as System.Data.DataTable).DefaultView.RowFilter = query;

                }
            }
            catch (Exception ex)
            {
                label5.Show();
                label5.Text = "Error: " + ex.Message;
                MessageBox.Show(ex.Message);
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
                        dtExcel = dataGridView.DataSource as System.Data.DataTable;



                        dataGridView3.Refresh();
                        dataGridView3.Visible = true;
                        dataGridView3.DataSource = dtExcel;
                        (dataGridView3.DataSource as System.Data.DataTable).DefaultView.RowFilter = query;

                        // Reading data from Excel File again for datagridview1
                        dtExcel = ReadExcel(currentFile, currentFileExt); //read excel file  
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;

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
        void reportFormatDT(List<string> AllCombosList, DataGridView dataGridView3, DataGridView dataGridView, DataGridView reportDataGridView, ComboBox cbWorkingColumn, int indexInsertData, string reportType, bool isCurrentOnly)
        {
            if (AllCombosList.Count() > 0)
            {

                dataGridView3.Refresh();
                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();
                dataGridView3.Columns.Clear();
                Percentages.percentagesList = new List<List<Percentages>>();

                //string query = "";
                //int u = 0;
                foreach (var query in AllCombosList)
                {
                    List<Percentages> percentagesList = new List<Percentages>();
                    // Filtration
                    if (!isCurrentOnly)
                    {
                        filteration(query, dataGridView);
                    }
                    else
                    {
                        filterationForLastCurrentFile(query, dataGridView);
                    }
                    if (dataGridView3.Rows.Count > 1)
                    {

                        if (reportType.ToUpper() == "CURRENT")
                        {
                            myExcel objCountList = new myExcel();
                            objCountList.Count = dataGridView3.Rows.Count - 1;
                            objCountList.ColumnNameOrQuery = query;
                            myExcel.CountList.Add(objCountList);
                        }

                        var selectedVal = cbWorkingColumn.Text;
                        MyDataGridView mdgv = new MyDataGridView();
                        var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, selectedVal);  // Getting selected working column index
                        var workingColumnData = mdgv.getColumnPlusRowData(dataGridView1, searchColumnNameIndexAfterWET, Convert.ToInt32(tbDataChar.Text)); // Getting selected working column data
                        var workingColumnDataPrevious = mdgv.getColumnPlusRowData(dataGridView2, searchColumnNameIndexAfterWET, Convert.ToInt32(tbDataChar.Text));
                        workingColumnData.AddRange(workingColumnDataPrevious);
                        workingColumnData = workingColumnData.Distinct().ToList();
                        workingColumnData.Sort();
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
                        queryAsColumn = queryAsColumn.Replace("AND", ",");
                        queryAsColumn = queryAsColumn.Replace("=", " ");
                        //queryAsColumn = queryAsColumn.Replace("'", "");
                        dt.Columns.Add(new DataColumn(queryAsColumn, typeof(string)));
                    }
                }
                DataRow toInsert = dt.NewRow();
                //// Setting Rows
                for (int j = 0; j < combinePerCetagesList.Count; j++)
                {
                    toInsert = dt.NewRow();

                    toInsert[0] = combinePerCetagesList[j];
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


                // Assigning Zero(0) to empty values\

                for (int rows = 0; rows < reportDataGridView.Rows.Count - 1; rows++)
                {
                    for (int col = 1; col < reportDataGridView.Rows[rows].Cells.Count; col++)
                    {
                        var item = reportDataGridView.Rows[rows].Cells[col].Value.ToString();
                        if (item == null || String.IsNullOrWhiteSpace(item))
                        {
                            reportDataGridView.Rows[rows].Cells[col].Value = 0;
                        }
                    }
                }

                // Sorting datagridView
                reportDataGridView.Sort(reportDataGridView.Columns[0], ListSortDirection.Ascending);

                //// Save to Excel
                //myExcel excel = new myExcel();
                //string title = reportType + " Report";
                //SaveFileDialog sfd = new SaveFileDialog();
                //sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                //sfd.FileName = reportType + ".xlsx";
                //if (sfd.ShowDialog() == DialogResult.OK)
                //{
                //    excel.ToCsV(reportDataGridView, reportType + " Report", "", "", title, sfd.FileName);

                //}
            }

        }
        void reportFormatDTNonRelationsal(List<string> AllCombosList, DataGridView dataGridView3, DataGridView dataGridView, DataGridView reportDataGridView, ComboBox cbWorkingColumn, int indexInsertData, string reportType, bool isCurrentOnly)
        {

            if (AllCombosList.Count() > 0)
            {

                dataGridView3.Refresh();
                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();
                dataGridView3.Columns.Clear();
                Percentages.percentagesList = new List<List<Percentages>>();



                //string query = "";
                //int u = 0;
                foreach (var query in AllCombosList)
                {
                    List<Percentages> percentagesList = new List<Percentages>();
                    // Filtration
                    if (!isCurrentOnly)
                    {
                        filteration(query, dataGridView);
                    }
                    else
                    {
                        filterationForLastCurrentFile(query, dataGridView);
                    }

                    if (dataGridView3.Rows.Count > 1)
                    {
                        if (reportType.ToUpper() == "CURRENT")
                        {
                            myExcel objCountList = new myExcel();
                            objCountList.Count = dataGridView3.Rows.Count - 1;
                            objCountList.ColumnNameOrQuery = query;
                            myExcel.CountList.Add(objCountList);
                        }

                        var selectedVal = cbWorkingColumn.Text;
                        MyDataGridView mdgv = new MyDataGridView();
                        var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, selectedVal);  // Getting selected working column index

                        var workingColumnData = mdgv.getColumnPlusRowData(dataGridView1, searchColumnNameIndexAfterWET, Convert.ToInt32(tbDataChar.Text)); // Getting selected working column data
                        var workingColumnDataPrevious = mdgv.getColumnPlusRowData(dataGridView2, searchColumnNameIndexAfterWET, Convert.ToInt32(tbDataChar.Text));
                        workingColumnData.AddRange(workingColumnDataPrevious);
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
                        queryAsColumn = queryAsColumn.Replace("AND", ",");
                        queryAsColumn = queryAsColumn.Replace("=", " ");
                        //queryAsColumn = queryAsColumn.Replace("'", "");
                        dt.Columns.Add(new DataColumn(queryAsColumn, typeof(string)));
                    }
                }
                DataRow toInsert = dt.NewRow();
                //// Setting Rows
                for (int j = 0; j < combinePerCetagesList.Count; j++)
                {
                    toInsert = dt.NewRow();

                    toInsert[0] = combinePerCetagesList[j];
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


                // Assigning Zero(0) to empty values\

                for (int rows = 0; rows < reportDataGridView.Rows.Count - 1; rows++)
                {
                    for (int col = 1; col < reportDataGridView.Rows[rows].Cells.Count; col++)
                    {
                        var item = reportDataGridView.Rows[rows].Cells[col].Value.ToString();
                        if (item == null || String.IsNullOrWhiteSpace(item))
                        {
                            reportDataGridView.Rows[rows].Cells[col].Value = 0;
                        }
                    }
                }

                // Sorting datagridView
                reportDataGridView.Sort(reportDataGridView.Columns[0], ListSortDirection.Ascending);

                //// Save to Excel
                //myExcel excel = new myExcel();
                //string title = reportType + " Report";
                //SaveFileDialog sfd = new SaveFileDialog();
                //sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                //sfd.FileName = reportType + ".xlsx";
                //if (sfd.ShowDialog() == DialogResult.OK)
                //{
                //    excel.ToCsV(reportDataGridView, reportType + " Report", "", "", title, sfd.FileName);

                //}
            }

        }
        public void isRelational(bool isCurrentOnly, List<string> getAllQueries)
        {
            var AllCombosList = new List<string>();
            //try
            //{
            if (!isCurrentOnly)
            {
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

                //

                // Getting all possible combinations
                Dictionary<int, List<string>> tags = new Dictionary<int, List<string>>();
                for (int j = 0; j < lst.Count(); j++)
                {
                    tags.Add(j, lst[j].Data);
                }
                NListBuilder nListBuilder = new NListBuilder(tags, "AND");
                AllCombosList = nListBuilder.AllCombos;
            }
            else
            {

                AllCombosList = getAllQueries;
            }
            // Current Excel Data
            reportFormatDT(AllCombosList, dataGridView3, dataGridView1, dataGridView4, cbWorkingColumn, 0, "Current", isCurrentOnly);
            progressBar1.Value = 50;
            if (!isCurrentOnly)
            {
                reportFormatDT(AllCombosList, dataGridView3, dataGridView2, dataGridView5, cbWorkingColumn, 0, "Previous", isCurrentOnly);
            }
            progressBar1.Value = 70;

            //  }
            //catch (Exception ex)
            //{
            //    label5.Show();
            //    label5.Text = "Error: " + ex.Message;
            //    MessageBox.Show(ex.Message);
            //}
        }
        public void isNonRelational(bool isCurrentOnly, List<string> getAllQueries)
        {
            var AllCombosList = new List<string>();
            try
            {
                if (!isCurrentOnly)
                {
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

                    List<MyListBox> lst2 = new List<MyListBox>();
                    foreach (var item in lst)
                    {
                        foreach (var item2 in item.Data)
                        {
                            MyListBox myListBox = new MyListBox();
                            myListBox.ColumnName = item.ColumnName;
                            myListBox.Data.Add(item2);
                            lst2.Add(myListBox);
                        }
                    }

                    //

                    //List<string> AllCombosList = new List<string>();
                    foreach (var item in lst2)
                    {
                        AllCombosList.Add(item.Data[0]);
                    }
                }
                else
                {

                    AllCombosList = getAllQueries;
                }

                // Current Excel Data
                reportFormatDTNonRelationsal(AllCombosList, dataGridView3, dataGridView1, dataGridView4, cbWorkingColumn, 0, "Current", isCurrentOnly);
                progressBar1.Value = 50;
                if (!isCurrentOnly)
                {
                    reportFormatDTNonRelationsal(AllCombosList, dataGridView3, dataGridView2, dataGridView5, cbWorkingColumn, 0, "Previous", isCurrentOnly);
                }
                progressBar1.Value = 70;
            }
            catch (Exception ex)
            {
                label5.Show();
                label5.Text = "Error: " + ex.Message;
                MessageBox.Show(ex.Message);
            }
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            cbNatureOfDeptCol.Text = "AND";
            DateTime expDate = new DateTime(2021,04,01);

            var dateAndTime = DateTime.Now;
            //var date = dateAndTime.Date.ToString("dd/MM/yyyy");
            if (dateAndTime > expDate)
            {
                MessageBox.Show("your trial has expired. Kindly contact developer for further information.");
                Application.Exit();
            }
            //button5.Hide();
            //string applicationLocation = System.Reflection.Assembly.GetEntryAssembly().Location;
            //string applicationDirectory = Path.GetDirectoryName(applicationLocation);
            //applicationDirectory = Application.ExecutablePath;

            var filePath = MyDirectory() + @"\range.xlsx";
            var fileExt = Path.GetExtension(filePath);
            try
            {
                System.Data.DataTable dtExcel = new System.Data.DataTable();
                dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                dgvRange.Visible = true;
                dgvRange.DataSource = dtExcel;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        void processData()
        {
            if (rbRelational.Checked == true)
            {
                if (cbWorkingColumn.Text == "-- Select --" || cbWorkingColumn.Text == "" || tbDataChar.Text == null || tbDataChar.Text == "" || Convert.ToInt32(tbDataChar.Text) <= 0)
                {
                    MessageBox.Show("Working column and Data Charater field must not be empty.");
                    return;
                }
                isRelational(false, new List<string>());
                if (dataGridView2.Rows.Count > 1)
                {
                    generateTargetFile();
                }
                else
                {
                    MessageBox.Show("Error: Can't generate target and previous report because uploaded data file don't contain any Previous data.");
                }
                progressBar1.Value = 100;

                myExcel excel = new myExcel();
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                sfd.FileName = "report.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    excel.ToCsVCombineDGV(dataGridView4, dataGridView5, dataGridView6, sfd.FileName);

                }

                MessageBox.Show("Finish");
            }
            else if (rbNonRelational.Checked == true)
            {
                if (cbWorkingColumn.Text == "-- Select --" || cbWorkingColumn.Text == "" || tbDataChar.Text == null || tbDataChar.Text == "" || Convert.ToInt32(tbDataChar.Text) == 0)
                {
                    MessageBox.Show("Working column and Data Charater field must not be empty.");
                    return;
                }
                isNonRelational(false, new List<string>());
                generateTargetFile();
                progressBar1.Value = 100;
                myExcel excel = new myExcel();
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                sfd.FileName = "report.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    excel.ToCsVCombineDGV(dataGridView4, dataGridView5, dataGridView6, sfd.FileName);

                }
                MessageBox.Show("Finish");
            }
            else
            {
                MessageBox.Show("Kindly select relational or non relational");
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (dgvRange.Rows.Count<=0)
            {
                MessageBox.Show("Please Upload Range File");
                return;
            }
            try
            {
                lblError.Text = "";
                processData();
            }
            catch (Exception ex)
            {
                lblError.Visible = true;
                lblError.Text = ex.Message;
            }
            //MessageBox.Show("Are You Sure?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                lblError.Text = "";
                progressBar1.Value = 0;

                resetDatatable(dataGridView1);
                resetDatatable(dataGridView2);
                resetDatatable(dataGridView3);
                resetDatatable(dataGridView4);
                resetDatatable(dataGridView5);
                resetDatatable(dataGridView6);

                lbMustCol.Items.Clear();
                lbDepCol.Items.Clear();

                //
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
                            var lstDependentCoulumnNames = new List<string>();
                            for (int j = 0; j < lstCoulumnNames.Count; j++)
                            {
                                lstDependentCoulumnNames.Add(lstCoulumnNames[j] + "#Exist");
                            }
                            for (int j = 0; j < lstCoulumnNames.Count; j++)
                            {
                                lstDependentCoulumnNames.Add(lstCoulumnNames[j] + "#NotExist");
                            }
                            lb.assignHeaderNameAfterWETToListBox(dataGridView1, lbDepCol, lstDependentCoulumnNames);

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
            catch (Exception ex)
            {
                lblError.Visible = true;
                lblError.Text = ex.Message;
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            //WriteExcel();
            try
            {
                //myExcel excel = new myExcel();
                //string title = "Excel Report";
                //SaveFileDialog sfd = new SaveFileDialog();
                //sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                //sfd.FileName = "demo.xlsx";
                //if (sfd.ShowDialog() == DialogResult.OK)
                //{
                //    excel.ToCsVCombineDGV(dataGridView4, dataGridView5, dataGridView6, sfd.FileName);

                //}

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                lblError.Text = "";
                UploadRangeFile();
            }

            catch (Exception ex)
            {
                lblError.Visible = true;
                lblError.Text = ex.Message;
            }
        }

        private void tbDataChar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        void AssignZerosOnEmptySpaces(DataGridView dgv)
        {
            for (int rows = 0; rows < dgv.Rows.Count - 1; rows++)
            {
                for (int col = 1; col < dgv.Rows[rows].Cells.Count; col++)
                {
                    var item = dgv.Rows[rows].Cells[col].Value.ToString();
                    if (item == null || String.IsNullOrWhiteSpace(item))
                    {
                        dgv.Rows[rows].Cells[col].Value = 0;
                    }
                }
            }
        }

        void generateTargetFile()
        {

            try
            {
                MyDataGridView mdgv = new MyDataGridView();


                // Sorting datagridView
                dataGridView4.Sort(dataGridView4.Columns[0], ListSortDirection.Ascending);
                dataGridView5.Sort(dataGridView5.Columns[0], ListSortDirection.Ascending);
                // Cloning DatagridView
                //var combinePerCetagesList = new Percentages().combinePercentagesList(Percentages.percentagesList);
                //combinePerCetagesList = combinePerCetagesList.Distinct().ToList();
                var getColumnData_Current = mdgv.getColumnData(dataGridView4, 0);
                var getColumnData_Previous = mdgv.getColumnData(dataGridView5, 0);
                var Q1_X_List = new List<string>();
                Q1_X_List = getColumnData_Current;
                Q1_X_List.AddRange(getColumnData_Previous); // Combine 2 List
                Q1_X_List = Q1_X_List.Distinct().ToList();


                // Setting up structure// If rows already created
                DataTable dt = new DataTable();
                //Setting Columns
                dt.Columns.Add(new DataColumn(cbWorkingColumn.Text, typeof(string)));
                for (int r = 0; r < Percentages.percentagesList.Count; r++)
                {
                    var item = Percentages.percentagesList[r];
                    var queryAsColumn = item[0].Query;
                    //Query.quries.Add(queryAsColumn);
                    queryAsColumn = queryAsColumn.Replace("AND", ",");
                    queryAsColumn = queryAsColumn.Replace("=", " ");
                    //queryAsColumn = queryAsColumn.Replace("'", "");
                    dt.Columns.Add(new DataColumn(queryAsColumn, typeof(string)));
                }
                DataRow toInsert = dt.NewRow();
                //// Setting Rows
                for (int j = 0; j < Q1_X_List.Count; j++)
                {
                    toInsert = dt.NewRow();

                    toInsert[0] = Q1_X_List[j];
                    dt.Rows.InsertAt(toInsert, 0);
                }


                //toInsert[0] = "Pervious";
                //dt.Rows.InsertAt(toInsert, 0);
                dataGridView6.DataSource = dt;

                ////
                //for (int r = 0; r < Percentages.percentagesList.Count; r++)
                //{
                //    var item = Percentages.percentagesList[r]; // Item
                //    for (int c = 0; c < item.Count; c++)
                //    {
                //        var colItem = item[c]; // Item
                //        int searchedRowIndexOfQ1_X = 0;
                //        var getQ1_X_Data = mdgv.getColumnData(dataGridView6, 0);
                //        if (colItem.ColumnValue != null)
                //        {
                //            searchedRowIndexOfQ1_X = getQ1_X_Data.FindIndex(x => x == colItem.ColumnValue);
                //            dataGridView6.Rows[searchedRowIndexOfQ1_X].Cells[r + 1].Value = colItem.Percentage;
                //        }
                //    }
                //}


                // Assigning Zero(0) to empty values\

                AssignZerosOnEmptySpaces(dataGridView6);

                // Cloning DatagridView End


                // Caluculating and Assigning Count Of Q_X 
                //var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView1, cbWorkingColumn.Text);

                //var get_Q1_ColumnData = mdgv.getColumnData(dataGridView1, searchColumnNameIndexAfterWET);
                //var getCurrent_Q1_ColumnData = mdgv.getColumnData(dataGridView6, 0);

                //List<myExcel> countList = mdgv.CalculateCountOf_Q_X_Using_Range(getCurrent_Q1_ColumnData, get_Q1_ColumnData);

                List<myExcel> countList = myExcel.CountList;

                // Getting range =============
                var getRangeColumnData = mdgv.getColumnData(dgvRange, 0);
                var getRangePercentageColumnData = mdgv.getColumnData(dgvRange, 1);
                List<Range> ranges = mdgv.GettingRangesFromRangeFile(getRangeColumnData, getRangePercentageColumnData);

                // 
                for (int rows = 0; rows < dataGridView6.Rows.Count - 1; rows++)
                {
                    for (int col = 1; col < dataGridView6.Rows[rows].Cells.Count; col++)
                    {
                        var Q_X_Value = dataGridView6.Rows[rows].Cells[0].Value.ToString();
                        int Q_X_Count = countList[col - 1].Count;

                        var getQ1_X_Data_Current = mdgv.getColumnData(dataGridView4, 0);
                        var searchedRowIndexOfQ1_X_Current = getQ1_X_Data_Current.FindIndex(x => x == Q_X_Value);
                        var getQ1_X_Data_Previous = mdgv.getColumnData(dataGridView5, 0);
                        var searchedRowIndexOfQ1_X_Previous = getQ1_X_Data_Previous.FindIndex(x => x == Q_X_Value);

                        if (searchedRowIndexOfQ1_X_Current >= 0 && searchedRowIndexOfQ1_X_Previous >= 0)
                        {
                            string currentDgvValue = dataGridView4.Rows[searchedRowIndexOfQ1_X_Current].Cells[col].Value.ToString();
                            string perviousDgvValue = dataGridView5.Rows[searchedRowIndexOfQ1_X_Previous].Cells[col].Value.ToString();

                            // Calculating difference of current and previous values
                            double calculateDifference = Convert.ToDouble(currentDgvValue) - Convert.ToDouble(perviousDgvValue);

                            // Checking value of Q_X Count with Q_X_Value
                            //foreach (var item in countList)
                            //{
                            //    if (item.ColumnValue == Q_X_Value)
                            //    {
                            //        Q_X_Count = item.Count;
                            //        break;
                            //    }
                            //}

                            // finding Q1_X index
                            var getQ1_X_Data = mdgv.getColumnData(dataGridView6, 0);
                            var searchedRowIndexOfQ1_X = getQ1_X_Data.FindIndex(x => x == Q_X_Value);
                            //reportDataGridView.Rows[searchedRowIndexOfQ1_X].Cells[r + 1].Value = colItem.Percentage;

                            if (searchedRowIndexOfQ1_X >= 0)
                            {

                                // Checking range
                                for (int i = 0; i < ranges.Count; i++)
                                {
                                    var range = ranges[i];
                                    if (Q_X_Count >= range.Min && Q_X_Count <= range.Max) // checking if Q_X valuel lie with in range
                                    {

                                        var convert_calculatedDiff_Positive = calculateDifference.ToString().Trim('-');

                                        if (Convert.ToDouble(range.Percentage) > Convert.ToDouble(convert_calculatedDiff_Positive))
                                        {
                                            // Do Nothing 
                                            // Target value is same as current
                                            dataGridView6.Rows[searchedRowIndexOfQ1_X].Cells[col].Value = dataGridView4.Rows[searchedRowIndexOfQ1_X_Current].Cells[col].Value;
                                        }
                                        else if (Convert.ToDouble(range.Percentage) < Convert.ToDouble(convert_calculatedDiff_Positive))
                                        {
                                            if (calculateDifference > 0) // Positive answer
                                            {
                                                string value = dataGridView5.Rows[searchedRowIndexOfQ1_X_Previous].Cells[col].Value.ToString(); // Previous Value
                                                decimal sum = Convert.ToDecimal(value) + range.Percentage;
                                                dataGridView6.Rows[searchedRowIndexOfQ1_X].Cells[col].Value = sum;
                                            }
                                            else if (calculateDifference < 0) // Negative answer
                                            {
                                                string value = dataGridView5.Rows[searchedRowIndexOfQ1_X_Previous].Cells[col].Value.ToString(); // Previous Value
                                                decimal diff = Convert.ToDecimal(value) - range.Percentage;
                                                dataGridView6.Rows[searchedRowIndexOfQ1_X].Cells[col].Value = diff.ToString().Trim('-');
                                            }
                                            else
                                            {
                                                dataGridView6.Rows[searchedRowIndexOfQ1_X].Cells[col].Value = 0;
                                            }
                                        }
                                        break;
                                    }

                                }
                            }
                        }
                    }
                }

                dataGridView6.Sort(dataGridView6.Columns[0], ListSortDirection.Ascending);

                //myExcel excel1 = new myExcel();
                //string title1 = "Target Report";
                //SaveFileDialog sfd1 = new SaveFileDialog();
                //sfd1.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                //sfd1.FileName = "tagetReport.xlsx";
                //if (sfd1.ShowDialog() == DialogResult.OK)
                //{
                //    excel1.ToCsV(dataGridView6, "Target Report1", "", "", title1, sfd1.FileName);
                //    
                //}


                MyDataGridView myDataGridView = new MyDataGridView();
                myDataGridView.SetTargetwithValidation(dataGridView6, dataGridView4, cbWorkingColumn.Text);

                //myExcel excel = new myExcel();
                //string title = "Target Report";
                //SaveFileDialog sfd = new SaveFileDialog();
                //sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                //sfd.FileName = "tagetReport.xlsx";
                //if (sfd.ShowDialog() == DialogResult.OK)
                //{
                //    excel.ToCsV(dataGridView6, "Target Report", "", "", title, sfd.FileName);

                //}

            }
            catch (Exception ex)
            {
                lblError.Visible = true;
                lblError.Text = ex.Message;
            }
        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            generateTargetFile();
        }

        private void btnUploadNewTargetFile_Click(object sender, EventArgs e)
        {

            lblError.Text = "";
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  


                // Displaying currently running file
                //currentlyRunningFile = "File Name: " + file.SafeFileName + currentFileExt;

                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView6.Visible = true;
                        dataGridView6.DataSource = dtExcel;

                        AssignZerosOnEmptySpaces(dataGridView6);
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

        void changePercentagesMultipleColumn(List<string> lst, DataGridView dataGridView, DataGridView dataGridView3, ComboBox cbNatureOfDeptCol, double TargetPercentage_Formula_Value)
        {
            ///// find out working column is single or multiple save it in variable

            MyDataGridView dgvClass = new MyDataGridView();
            // string workingColumnNature = "";//dgvClass.getColumnNature();

            dataGridView3.Refresh();
            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            Percentages.percentagesList = new List<List<Percentages>>();

            var AllCombosList = lst;

            //string query = "";
            //int u = 0;
            int col = 1;

            MyDataGridView mdgv = new MyDataGridView();
            var getAllQueries = mdgv.getRegenratedQueries(dataGridView6);
            var getQ1_X_ColumnName = getAllQueries[0];

            List<IncreaseOrDecrease> lstIncOrDec = new List<IncreaseOrDecrease>();
            IncreaseOrDecrease.decreaseList = new List<IncreaseOrDecrease>();
            IncreaseOrDecrease.increaseList = new List<IncreaseOrDecrease>();
            List<IncreaseOrDecrease> increaseList = new List<IncreaseOrDecrease>();
            List<IncreaseOrDecrease> decreaseList = new List<IncreaseOrDecrease>();
            foreach (var query in AllCombosList)
            {
                int rows = 0;
                for (; rows < dataGridView4.Rows.Count - 1;)
                {
                    IncreaseOrDecrease increaseOrDecrease = new IncreaseOrDecrease();
                    var Q_X_Value = dataGridView4.Rows[rows].Cells[0].Value.ToString();
                    var Q_X_PercentageValue_OldTarget = dataGridView4.Rows[rows].Cells[col].Value.ToString();
                    var Q_X_PercentageValue_NewTarget = dataGridView6.Rows[rows].Cells[col].Value.ToString();
                    if (Convert.ToDouble(Q_X_PercentageValue_NewTarget) > Convert.ToDouble(Q_X_PercentageValue_OldTarget))
                    {
                        increaseOrDecrease = new IncreaseOrDecrease()
                        {
                            colIndex = col,
                            columnName = getQ1_X_ColumnName,
                            columnValue = Q_X_Value,
                            query = query,
                            rowIndex = rows,
                            status = "increase"
                        };
                        increaseList.Add(increaseOrDecrease);
                    }
                    else if (Convert.ToDouble(Q_X_PercentageValue_NewTarget) < Convert.ToDouble(Q_X_PercentageValue_OldTarget))
                    {
                        increaseOrDecrease = new IncreaseOrDecrease()
                        {
                            colIndex = col,
                            columnName = getQ1_X_ColumnName,
                            columnValue = Q_X_Value,
                            query = query,
                            rowIndex = rows,
                            status = "decrease"
                        };
                        decreaseList.Add(increaseOrDecrease);
                    }
                    else if (Convert.ToDouble(Q_X_PercentageValue_NewTarget) == Convert.ToDouble(Q_X_PercentageValue_OldTarget))
                    {
                        // Do nothing
                    }
                    rows++;
                }
                col++;
            }
            IncreaseOrDecrease.increaseList = increaseList;
            IncreaseOrDecrease.decreaseList = decreaseList;


            col = 1;
            foreach (var query in AllCombosList)
            {
                List<Percentages> percentagesList = new List<Percentages>();
                // Filtration
                filterationOnChangePercentage(query, dataGridView);
                int rows = 0;

                // if chcek for single or mulitple


                //if (workingColumnNature == "s")
                //{
                //    ///// single working
                //}
                //else if (workingColumnNature == "m")
                //{
                //    //// multiple working
                //}

                // multiple

                for (; rows < dataGridView4.Rows.Count - 1;)
                {
                    var Q_X_Value = dataGridView4.Rows[rows].Cells[0].Value.ToString();
                    var Q_X_PercentageValue_OldTarget = dataGridView4.Rows[rows].Cells[col].Value.ToString(); /// current file
                    var Q_X_PercentageValue_NewTarget = dataGridView6.Rows[rows].Cells[col].Value.ToString();
                    if (Convert.ToDouble(Q_X_PercentageValue_NewTarget) > Convert.ToDouble(Q_X_PercentageValue_OldTarget))
                    {
                        mdgv.increasePercentage(dataGridView3, lbDepCol, lbMustCol, getQ1_X_ColumnName, Q_X_Value, Q_X_PercentageValue_NewTarget, dataGridView1, dataGridView6, dgvRange, cbNatureOfDeptCol, TargetPercentage_Formula_Value);
                    }
                    else if (Convert.ToDouble(Q_X_PercentageValue_NewTarget) < Convert.ToDouble(Q_X_PercentageValue_OldTarget))
                    {
                        mdgv.decreasePercentage(dataGridView3, lbDepCol, lbMustCol, getQ1_X_ColumnName, Q_X_Value, Q_X_PercentageValue_NewTarget, dataGridView1, dataGridView6, dgvRange, cbNatureOfDeptCol, TargetPercentage_Formula_Value);
                    }
                    else if (Convert.ToDouble(Q_X_PercentageValue_NewTarget) == Convert.ToDouble(Q_X_PercentageValue_OldTarget))
                    {
                        // Do nothing
                    }
                    rows++;
                }
                col++;

                // Assign and replace values of current file i.e. (datagridview1)
                MyDataGridView mdgv2 = new MyDataGridView();
                mdgv2.AssignValuesToCurrentFile(dataGridView1, dataGridView3);
                //break;
            }
        }

        private void btnChangePercentages_Click(object sender, EventArgs e)
        {
            try
            {
                lblError.Text = "";
                progressBar1.Value = 0;
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Kindly upload current file.");
                    return;
                }

                if (tbTargetPerFormulaValue.Text == "" || Convert.ToDouble(tbTargetPerFormulaValue.Text) < 0.001 && Convert.ToDouble(tbTargetPerFormulaValue.Text) > 1)
                {
                    MessageBox.Show("Target Percentage Formula Value must not be empty and use a range (0.001 - 1.00)");
                    return;
                }

                if (cbNatureOfDeptCol.Text == null || cbNatureOfDeptCol.Text == "")
                {
                    MessageBox.Show("NATURE OF DEPENDENT COLUMN must not be empty.");
                    return;
                }

                if (dgvRange.Rows.Count == 0)
                {
                    MessageBox.Show("Kindly Upload Range File");
                    return;
                }
                if (dataGridView4.Rows.Count == 0)
                {
                    MessageBox.Show("Kindly Upload Current File");
                    return;
                }
                if (dataGridView6.Rows.Count == 0)
                {
                    MessageBox.Show("Kindly Upload Target File");
                    return;
                }
                // Temporary save current file
                var temp_currentPercentageReport = new MyDataGridView().CopyDataGridView(dataGridView4);
                DataRow recRow = temp_currentPercentageReport.Rows[temp_currentPercentageReport.Rows.Count - 1];
                recRow.Delete();
                temp_currentPercentageReport.AcceptChanges();

                // Re read datagridview1 or excel file
                int i = 0;
                dataGridView1.Refresh();
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = ReadExcel(currentFile, currentFileExt);
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

                // END Re read datagridview1 or excel file

                double TargetPercentage_Formula_Value = double.Parse(tbTargetPerFormulaValue.Text);

                MyDataGridView mdgv = new MyDataGridView();
                var getAllQueries = mdgv.getRegenratedQueries(dataGridView6);
                var getQ1_X_Value = getAllQueries[0];
                cbWorkingColumn.Text = getQ1_X_Value;
                getAllQueries.RemoveAt(0);
                if (getAllQueries.Count() > 0)
                {
                    string workingColumnNature = mdgv.getColumnNature(getQ1_X_Value);
                    if (workingColumnNature == "s")
                    {
                        changePercentageSingleColumn(dataGridView6, dataGridView4, getAllQueries, lbDepCol, lbMustCol, getQ1_X_Value, cbNatureOfDeptCol, TargetPercentage_Formula_Value);
                    }
                    else if (workingColumnNature == "m")
                    {
                        changePercentagesMultipleColumn(getAllQueries, dataGridView1, dataGridView3, cbNatureOfDeptCol, TargetPercentage_Formula_Value);
                    }
                    else
                    {
                        ///error
                        //changePercentagesMultipleColumn(getAllQueries, dataGridView1, dataGridView3, cbNatureOfDeptCol, TargetPercentage_Formula_Value);
                        MessageBox.Show("Error: Input file not in correct format.");
                        return;

                    }
                    progressBar1.Value = 40;


                }


                //
                myExcel excel = new myExcel();
                string title = "New Current Report";
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                sfd.FileName = "newCurrentFile.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    excel.ToCsV(dataGridView1, "New Current File", "", "", title, sfd.FileName);

                }

                // Generating Current file form new Current report
                if (rbRelational.Checked == true)
                {
                    isRelational(true, getAllQueries);

                    myExcel excel2 = new myExcel();
                    SaveFileDialog sfd2 = new SaveFileDialog();
                    sfd2.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                    sfd2.FileName = "newCurrentReport.xlsx";
                    if (sfd2.ShowDialog() == DialogResult.OK)
                    {
                        excel.ToCsV(dataGridView4, "New Current Report", "", "", title, sfd2.FileName);

                    }
                }
                else
                {
                    isNonRelational(true, getAllQueries);

                    myExcel excel2 = new myExcel();
                    SaveFileDialog sfd2 = new SaveFileDialog();
                    sfd2.Filter = "Excel Documents (*.xlsx)|*.xlsx";
                    sfd2.FileName = "newCurrentReport.xlsx";
                    if (sfd2.ShowDialog() == DialogResult.OK)
                    {
                        excel.ToCsV(dataGridView4, "New Current Report Percentage", "", "", title, sfd2.FileName);

                    }
                }

                // Reassign value to current datagridview from temp DataTable
                dataGridView4.DataSource = temp_currentPercentageReport;
                progressBar1.Value = 100;
            }
            catch (Exception ex)
            {
                lblError.Show();
                lblError.Text = "Error: " + ex.Message;
            }
        }

        private void btnUploadOladTargetFile_Click(object sender, EventArgs e)
        {

            lblError.Text = "";
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  


                // Displaying currently running file
                //currentlyRunningFile = "File Name: " + file.SafeFileName + currentFileExt;

                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                        dataGridView4.Visible = true;
                        dataGridView4.DataSource = dtExcel;

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
        void changePercentageSingleColumn(DataGridView target, DataGridView currentPercent, List<string> queries, ListBox dependentColumns, ListBox mustColumns, string workingColumn, ComboBox cbNatureOfDeptCol, double TargetPercentage_Formula_Value)
        {
            try
            {
                dataGridView3.Refresh();
                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();
                dataGridView3.Columns.Clear();
                MyDataGridView mdgv = new MyDataGridView();
                var brands = mdgv.getColumnData(target, 0);
                for (int i = 1; i < target.Columns.Count; i++)
                {
                    List<string> increase = new List<string>();
                    List<string> decrease = new List<string>();

                    var targetColumns = mdgv.getColumnData(target, i);
                    var currentColumns = mdgv.getColumnData(currentPercent, i);

                    for (int k = 0; k < targetColumns.Count; k++)
                    {
                        if (Convert.ToDecimal(targetColumns[k]) > Convert.ToDecimal(currentColumns[k]))
                        {
                            increase.Add(brands[k]);
                        }
                        else if (Convert.ToDecimal(targetColumns[k]) < Convert.ToDecimal(currentColumns[k]))
                        {
                            decrease.Add(brands[k]);
                        }
                    }

                    filterationOnChangePercentage(queries[i - 1], dataGridView1);
                    var workingColumnIndex = mdgv.searchColumnNameIndexAfterWET(dataGridView3, workingColumn);
                    var depColListBoxItems = dependentColumns.SelectedItems;
                    var mustColListBoxItems = mustColumns.SelectedItems;


                    for (int rows = 0; rows < dataGridView3.Rows.Count - 1; rows++)
                    {
                        var workingData = dataGridView3.Rows[rows].Cells[workingColumnIndex].Value.ToString();

                        if (decrease.Contains(workingData))
                        {
                            if (increase.Count() > 0)
                            {
                                bool isDependent = false;
                                int increaseIndex = 0;
                               
                                x:
                                    if (depColListBoxItems.Count > 0) // Optionl
                                    {
                                        for (int mustCheckIndex = increaseIndex; mustCheckIndex < increase.Count(); mustCheckIndex++)
                                        {
                                            if (cbNatureOfDeptCol.Text.Trim().ToUpper().ToUpper() == "AND")
                                            {
                                                isDependent = mdgv.isDependentCol_Satisfy_AND(dataGridView3, lbDepCol, increase[mustCheckIndex], rows);
                                            }
                                            else if (cbNatureOfDeptCol.Text.Trim().ToUpper().ToUpper() == "OR")
                                            {
                                                isDependent = mdgv.isDependentCol_Satisfy_OR(dataGridView3, lbDepCol, increase[mustCheckIndex], rows);
                                            }
                                            else
                                            {
                                                MessageBox.Show("Nature of dependent column not valid.");
                                                return;
                                            }
                                            if (isDependent)
                                            {
                                                increaseIndex = mustCheckIndex;
                                                break;
                                            }
                                        }
                                        
                                    }
                                    else
                                    {
                                        isDependent = true;
                                       
                                    }
                                
                               
                                if (isDependent)
                                {
                                    dataGridView3.Rows[rows].Cells[workingColumnIndex].Value = increase[increaseIndex];

                                    int filterDataCount = dataGridView3.Rows.Count - 1;



                                    var increaseDataPercentage = mdgv.calculatePercentage(dataGridView3, increase[increaseIndex], workingColumnIndex);
                                    var decreaseDataPercentage = mdgv.calculatePercentage(dataGridView3, workingData, workingColumnIndex);

                                    var decreaseData_Index = brands.IndexOf(workingData);
                                    var increaseData_Index = brands.IndexOf(increase[increaseIndex]);

                                    var arrMinMax_get_PercentageLimit_Target_increase = mdgv.calculatePercentageLimit(dataGridView3, dataGridView1, dataGridView6, workingColumnIndex, increase[increaseIndex], Convert.ToDouble(targetColumns[increaseData_Index]), dgvRange, TargetPercentage_Formula_Value, filterDataCount);
                                    var minPercentageValue_increase = arrMinMax_get_PercentageLimit_Target_increase[0];
                                    var maxPercentageValue_increase = arrMinMax_get_PercentageLimit_Target_increase[1];

                                    var arrMinMax_get_PercentageLimit_Target_decrease = mdgv.calculatePercentageLimit(dataGridView3, dataGridView1, dataGridView6, workingColumnIndex, workingData, Convert.ToDouble(targetColumns[decreaseData_Index]), dgvRange, TargetPercentage_Formula_Value, filterDataCount);
                                    var minPercentageValue_decrease = arrMinMax_get_PercentageLimit_Target_decrease[0];
                                    var maxPercentageValue_decrease = arrMinMax_get_PercentageLimit_Target_decrease[1];

                                    if ((double.Parse(decreaseDataPercentage) < minPercentageValue_decrease))
                                    {
                                        dataGridView3.Rows[rows].Cells[workingColumnIndex].Value = workingData;
                                        break;
                                    }
                                    else if (double.Parse(decreaseDataPercentage) >= minPercentageValue_decrease && double.Parse(decreaseDataPercentage) <= maxPercentageValue_decrease)
                                    {
                                        decrease.Remove(workingData);
                                    }

                                    if (double.Parse(increaseDataPercentage) > maxPercentageValue_increase)
                                    {
                                        if (increaseIndex+1 < increase.Count())
                                        {
                                            increaseIndex++;
                                            goto x;
                                        }
                                        dataGridView3.Rows[rows].Cells[workingColumnIndex].Value = workingData;
                                        //for (int k = 1; k < increase.Count; k++)
                                        //{
                                        //    dataGridView3.Rows[rows].Cells[workingColumnIndex].Value = increase[k];
                                        //    var temp_increaseDataPercentage = mdgv.calculatePercentage(dataGridView3, increase[k], workingColumnIndex);
                                        //    var temp_increaseData_Index = brands.IndexOf(increase[k]);
                                        //    var temp_arrMinMax_get_PercentageLimit_Target_increase = mdgv.calculatePercentageLimit(dataGridView3, dataGridView1, dataGridView6, workingColumnIndex, increase[k], Convert.ToDouble(targetColumns[temp_increaseData_Index]), dgvRange, TargetPercentage_Formula_Value, filterDataCount);
                                        //    var temp_minPercentageValue_increase = temp_arrMinMax_get_PercentageLimit_Target_increase[0];
                                        //    var temp_maxPercentageValue_increase = temp_arrMinMax_get_PercentageLimit_Target_increase[1];

                                        //    if (double.Parse(temp_increaseDataPercentage) >= temp_minPercentageValue_increase && double.Parse(temp_increaseDataPercentage) <= temp_maxPercentageValue_increase)
                                        //    {
                                        //        if (mustColListBoxItems.Count > 0) // Optionl
                                        //        {
                                        //            mdgv.assignValuesToMustColumn(lbMustCol, dataGridView3, increase[k], rows);
                                        //        }
                                        //        increase.RemoveAt(k);
                                        //        break;
                                        //    }
                                        //    else if(double.Parse(temp_increaseDataPercentage) < temp_minPercentageValue_increase)
                                        //    {
                                        //        if (mustColListBoxItems.Count > 0) // Optionl
                                        //        {
                                        //            mdgv.assignValuesToMustColumn(lbMustCol, dataGridView3, increase[k], rows);
                                        //        }
                                        //        break;
                                        //    }
                                        //}
                                    }
                                    else if (double.Parse(increaseDataPercentage) >= minPercentageValue_increase && double.Parse(increaseDataPercentage) <= maxPercentageValue_increase)
                                    {
                                        if (mustColListBoxItems.Count > 0) // Optionl
                                        {
                                            mdgv.assignValuesToMustColumn(lbMustCol, dataGridView3, increase[increaseIndex], rows);
                                        }
                                        increase.RemoveAt(increaseIndex);
                                    }
                                    else if (double.Parse(increaseDataPercentage) < minPercentageValue_increase)
                                    {
                                        mdgv.assignValuesToMustColumn(lbMustCol, dataGridView3, increase[increaseIndex], rows);
                                    }


                                }
                            }
                        }
                    }
                    mdgv.AssignValuesToCurrentFile(dataGridView1, dataGridView3);
                }
            }
            catch (Exception ex)
            {
                lblError.Visible = true;
                lblError.Text = ex.Message;
            }
        }
        private void tbTargetPerFormulaValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < 48 || e.KeyChar > 57) && e.KeyChar != 46
                                 && e.KeyChar != 8)
                e.Handled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void Label13_Click(object sender, EventArgs e)
        {

        }
    }
}
