using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject.Model
{
    /// <summary>
    /// datagrid view functions
    /// </summary>
    class MyDataGridView
    {
        int colIndex = 2;
        public List<string> getColumnNames(DataGridView dataGridView1, int colIndex)
        {
            List<string> lst = new List<string>();

            int colCount = dataGridView1.Columns.Count;
            for (int i = this.colIndex; i < colCount; i++)
            {
                //DataGridViewRow selectedRow = dataGridView1.Rows[rowIndex];
                //var item = selectedRow.Cells[i].Value.ToString();
                //if (item.ToLower() == "wet")
                //{
                //    break;
                //}
                //lst.Add(item);
                string text = dataGridView1.Columns[i].HeaderText;
                if (text.ToLower() == "wet")
                {
                    break;
                }
                lst.Add(text);
            }

            return lst;
        }
        public List<string> getAllColumnNames(DataGridView dataGridView1, int colIndex)
        {
            List<string> lst = new List<string>();

            int colCount = dataGridView1.Columns.Count;
            for (int i = colIndex; i < colCount; i++)
            {
                //DataGridViewRow selectedRow = dataGridView1.Rows[rowIndex];
                //var item = selectedRow.Cells[i].Value.ToString();
                //if (item.ToLower() == "wet")
                //{
                //    break;
                //}
                //lst.Add(item);
                string text = dataGridView1.Columns[i].HeaderText;
                if (text.ToLower() == "wet")
                {
                    break;
                }
                lst.Add(text);
            }

            return lst;
        }
        public List<string> getSingleColumnNames(DataGridView dataGridView1)
        {
            List<string> lstSingleCol = new List<string>();

            int colCount = dataGridView1.Columns.Count;
            for (int i = this.colIndex; i < colCount; i++)
            {
                string text = dataGridView1.Columns[i].HeaderText;
                if (text.Contains("_S"))
                {
                    lstSingleCol.Add(text);
                }
            }

            return lstSingleCol;
        }
        public List<string> getMultiColumnNames(DataGridView dataGridView1)
        {
            List<string> lstMultiCol = new List<string>();

            int colCount = dataGridView1.Columns.Count;
            for (int i = this.colIndex; i < colCount; i++)
            {
                string text = dataGridView1.Columns[i].HeaderText;
                if (text.Contains("_M"))
                {
                    lstMultiCol.Add(text);
                }
            }

            return lstMultiCol;
        }
        public List<string> getColumnNamesAfterWET(DataGridView dataGridView1)
        {
            List<string> lst = new List<string>();
            bool isWet = false;

            int colCount = dataGridView1.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                //DataGridViewRow selectedRow = dataGridView1.Rows[rowIndex];
                //var item = selectedRow.Cells[i].Value.ToString();
                //if (item.ToLower() == "wet")
                //{
                //    break;
                //}
                //lst.Add(item);
                string text = dataGridView1.Columns[i].HeaderText;
                if (isWet)
                {
                    lst.Add(text);
                }
                if (text.ToLower() == "wet")
                {
                    isWet = true;
                }
            }

            return lst;
        }
        public int searchColumnNameIndexAfterWET(DataGridView dataGridView1, string columnName)
        {
            List<string> lst = new List<string>();
            bool isWet = false;

            int colCount = dataGridView1.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                //DataGridViewRow selectedRow = dataGridView1.Rows[rowIndex];
                //var item = selectedRow.Cells[i].Value.ToString();
                //if (item.ToLower() == "wet")
                //{
                //    break;
                //}
                //lst.Add(item);
                string text = dataGridView1.Columns[i].HeaderText;
                if (isWet)
                {
                    if (text.Trim() == columnName.Trim())
                    {
                        return i;
                    }
                }
                if (text.ToLower() == "wet")
                {
                    isWet = true;
                }
            }

            return 0;
        }
        public List<string> getColumnData(DataGridView dataGridView1, int colIndex)
        {
            List<string> lst = new List<string>();

            int rowCount = dataGridView1.Rows.Count;
            for (int i = 0; i < rowCount - 1; i++)
            {
                DataGridViewRow selectedRow = dataGridView1.Rows[i];
                var item = selectedRow.Cells[colIndex].Value.ToString();
                lst.Add(item);
            }

            return lst;
        }
        public List<string> getColumnPlusRowData(DataGridView dataGridView1, int colIndex, int chunkSize)
        {
            List<string> lst = new List<string>();

            int rowCount = dataGridView1.Rows.Count;
            for (int i = 0; i < rowCount - 1; i++)
            {
                DataGridViewRow selectedRow = dataGridView1.Rows[i];
                var item = selectedRow.Cells[colIndex].Value.ToString();
                var rowDataList = Split(item, chunkSize);
                for (int j = 0; j < rowDataList.Count; j++)
                {
                    lst.Add(rowDataList[j]);
                }
            }
            lst = lst.Distinct().ToList();
            return lst;
        }
        public List<string> Split(string str, int chunkSize)
        {
            str = str.Replace(" ", String.Empty);
            //return (from Match m in Regex.Matches(str, @"\d{1," + chunkSize + "}")
            //        select m.Value).ToList();
            return Enumerable.Range(0, str.Length / chunkSize)
        .Select(i => str.Substring(i * chunkSize, chunkSize)).ToList();
        }

        public string getTrimed_QueryColName_AND(string queryAsColumn)
        {
            queryAsColumn = queryAsColumn.Replace(" AND ", ",");
            queryAsColumn = queryAsColumn.Replace("=", " ");
            return queryAsColumn;
        }

        public List<List<string>> getRowData(DataGridView dataGridView1, int Q_index)
        {
            List<List<string>> lst = new List<List<string>>();

            int rowCount = dataGridView1.Rows.Count;
            //int headerWETIndex = getWETIndex(dataGridView1) + 1;
            int headerWETIndex = Q_index;

            for (int i = 0; i < rowCount - 1; i++)
            {
                DataGridViewRow selectedRow = dataGridView1.Rows[i];

                List<string> lstColData = new List<string>();
                for (int j = 0; j <= headerWETIndex; j++)
                {
                    var item = selectedRow.Cells[j].Value.ToString();

                    lstColData.Add(item);
                }

                lst.Add(lstColData);
            }

            return lst;
        }
        public double getSumOfWET(DataGridView dataGridView1)
        {
            double sum = 0;
            int headerWETIndex = getWETIndex(dataGridView1);

            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                var val = dataGridView1.Rows[i].Cells[headerWETIndex].Value;
                if (val == null || val == DBNull.Value || String.IsNullOrWhiteSpace(val.ToString()))
                {
                    return sum;
                }
                string str = val.ToString().Replace('"', ' ').Trim();
                sum += Convert.ToDouble(str);
            }
            //label1.Text = sum.ToString();
            return sum;
        }
        public int getWETIndex(DataGridView dgv)
        {
            int headerWETIndex = 0;
            int colCount = dgv.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                string text = dgv.Columns[i].HeaderText;
                if (text.ToLower() == "wet")
                {
                    headerWETIndex = i;
                    break;
                }
            }

            return headerWETIndex;
        }
        public string calculatePercentage(DataGridView dgv, string val, int Q_index)
        {
            MyDataGridView mdgv = new MyDataGridView();
            double calulateWETSum = mdgv.getSumOfWET(dgv);

            var getRowData = mdgv.getRowData(dgv, Q_index);
            double brandSum = 0;
            var wetIndex = getWETIndex(dgv);
            for (int i = 0; i < getRowData.Count(); i++)
            {
                var item = getRowData[i];
                var q1_1_value = item[item.Count() - 1]; // getting last value
                var rowDataList = Split(q1_1_value, val.Length);
                //for (int j = 0; j < rowDataList.Count; j++)
                //{
                //    var item2 = rowDataList[j];
                //    if (item2.ToString() == val) // comparing last value with comboBox value i.e Q1_1
                //    {
                //        brandSum += double.Parse(item[wetIndex]);
                //    }
                //}
                if (rowDataList.Contains(val))
                {
                    brandSum += double.Parse(item[wetIndex]);
                }
            }
            if (calulateWETSum <= 0)
            {
                return 0.ToString();
            }

            if (calulateWETSum <= 0)
            {
                return 0.ToString();
            }

            var total = brandSum / calulateWETSum;
            //if (total<=0)
            //{
            //    return "0";
            //}
            var temp = (total * 100).ToString();
            //if (val == 16.ToString())
            //{

            //}

            return temp;
        }
        public List<int> getIndexesAfterWET(DataGridView dgv)
        {
            List<int> lst = new List<int>();
            bool IsHeaderWETIndex = false;
            int colCount = dgv.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                string text = dgv.Columns[i].HeaderText;
                if (text.ToLower() == "wet")
                {
                    IsHeaderWETIndex = true;
                }
                else if (IsHeaderWETIndex)
                {
                    lst.Add(i);
                }
            }

            return lst;
        }
        public DataTable CopyDataGridView(DataGridView dgv_org)
        {
            DataTable dt = new DataTable();
            //Setting Columns
            for (int c = 0; c < dgv_org.ColumnCount; c++)
            {
                var ColumnName = dgv_org.Columns[c].HeaderText;
                dt.Columns.Add(new DataColumn(ColumnName, typeof(string)));
            }
            //// Setting Rows
            foreach (DataGridViewRow row in dgv_org.Rows)
            {
                DataRow dRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dRow);
            }
            return dt;

        }
        public List<string> getRegenratedQueries(DataGridView dataGridView5)
        {
            List<string> quries = new List<string>();

            MyDataGridView mdgv = new MyDataGridView();
            var getAllColumnName = mdgv.getAllColumnNames(dataGridView5, 0);
            for (int j = 0; j < getAllColumnName.Count; j++)
            {
                var item = getAllColumnName[j];
                var arr = item.Split(',');
                var query = "";
                for (int i = 0; i < arr.Length; i++)
                {
                    var a = arr[i].Replace(" ", "=");
                    a = a.Trim('=');
                    query += a + " AND ";
                }
                Regex regex = new Regex("(\\s+(AND|OR)\\s*)$");
                query = regex.Replace(query, "");
                query = query.Trim();
                quries.Add(query);
            }

            return quries;
        }
        public void AssignValuesToCurrentFile(DataGridView dataGridView1, DataGridView dataGridView3)
        {
            int currentFile_rowIndex = -1;
            for (int rows = 0; rows < dataGridView3.Rows.Count - 1; rows++)
            {
                int formNoIndex = 1;
                var formNo = dataGridView3.Rows[rows].Cells[formNoIndex].Value.ToString(); //Form Number Value

                // Get filtered file row data
                var filteredFileRowData = getSpecificRowData(dataGridView3, rows);

                // Search Filtered File Row and Form number must be unique
                DataGridViewRow currentFile_row = dataGridView1.Rows
                    .Cast<DataGridViewRow>()
                    .Where(r => r.Cells[formNoIndex].Value.ToString().Equals(formNo))
                    .First();

                currentFile_rowIndex = currentFile_row.Index;

                // Replace Row
                replaceRowData(dataGridView1, filteredFileRowData, currentFile_rowIndex);
            }

        }
        public void replaceRowData(DataGridView dataGridView, List<string> lst, int rowIndex)
        {
            for (int i = 0; i < lst.Count; i++)
            {
                var item = lst[i];
                dataGridView.Rows[rowIndex].Cells[i + 1].Value = item;
            }
        }


        public List<string> getSpecificRowData(DataGridView dataGridView, int rows)
        {
            List<string> lst = new List<string>();
            for (int col = 1; col < dataGridView.Rows[rows].Cells.Count; col++)
            {
                var val = dataGridView.Rows[rows].Cells[col].Value.ToString();
                lst.Add(val);
            }
            return lst;
        }
        public List<myExcel> CalculateCountOf_Q_X_Using_Range(List<string> getCurrent_Q1_ColumnData, List<string> get_Q1_ColumnData)
        {
            List<myExcel> countList = new List<myExcel>();
            foreach (var item in getCurrent_Q1_ColumnData)
            {
                myExcel myExcel = new myExcel();
                myExcel.ColumnValue = item;
                foreach (var Q1_ColumnData in get_Q1_ColumnData)
                {
                    var rowDataList = Split(Q1_ColumnData, item.Length);
                    if (rowDataList.Contains(item))
                    {
                        myExcel.Count += 1;
                    }
                }
                countList.Add(myExcel);
            }
            return countList;
        }
        public List<Range> GettingRangesFromRangeFile(List<string> getRangeColumnData, List<string> getRangePercentageColumnData)
        {
            List<Range> ranges = new List<Range>();
            for (int i = 0; i < getRangeColumnData.Count; i++)
            {
                Range range = new Range();
                var item = getRangeColumnData[i];
                var splitRange = item.Split('-');
                var min = splitRange[0];
                var max = splitRange[1];
                var percentage = getRangePercentageColumnData[i].ToString();
                range.Min = Convert.ToInt32(min);
                range.Max = Convert.ToInt32(max);
                range.Percentage = Convert.ToDecimal(percentage);

                ranges.Add(range);
            }
            return ranges;
        }
        public void assignValuesToMustColumn(ListBox lbMustCol, DataGridView dataGridView3, string Q_X_Value, int rows)
        {
            MyDataGridView mdgv = new MyDataGridView();
            var result = "";
            var mustColListBoxItems = lbMustCol.SelectedItems;
            foreach (var item in mustColListBoxItems)
            {
                var Q_X_ColumnName = item.ToString();
                var checkNatureOfColumnName = getColumnNature(Q_X_ColumnName);
                var searchColumnNameIndexAfterWET_Q_X = mdgv.searchColumnNameIndexAfterWET(dataGridView3, Q_X_ColumnName);
                //var getColumnData = mdgv.getColumnData(dataGridView3, searchColumnNameIndexAfterWET_Q_X);
                var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value.ToString();
                var lst_GetAllCellData = Split(get_Q_X_ColumnData, Q_X_Value.Length);

                if (checkNatureOfColumnName.ToLower() == "s")
                {
                    // Not Exist
                    if (get_Q_X_ColumnData != Q_X_Value)
                    {
                        dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value = Q_X_Value;
                    }
                    else if (get_Q_X_ColumnData == Q_X_Value)// Exist
                    {
                        //dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value = result;
                    }
                }
                else
                {
                    // Not Exist
                    if (!lst_GetAllCellData.Contains(Q_X_Value))
                    {
                        lst_GetAllCellData.Add(Q_X_Value);
                        result = string.Join("", lst_GetAllCellData.ToArray());
                        dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value = result;
                    }
                    else if (lst_GetAllCellData.Contains(Q_X_Value))// Exist
                    {
                        result = string.Join("", lst_GetAllCellData.ToArray());
                        dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value = result;
                    }
                }
            }
        }
        public void increasePercentage(DataGridView dataGridView3, ListBox lbDepCol, ListBox lbMustCol, string getQ1_X_ColumnName, string Q_X_Value, string Q_X_PercentageValue_NewTarget, DataGridView dataGridView1, DataGridView dataGridView6, DataGridView dgvRange, ComboBox cbNatureOfDeptCol, double TargetPercentage_Formula_Value)
        {
            MyDataGridView mdgv = new MyDataGridView();
            var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, getQ1_X_ColumnName);

            int filterDataCount = dataGridView3.Rows.Count - 1;
            //Getting Minimum and maximum values
            var arrMinMax_get_PercentageLimit_Target = calculatePercentageLimit(dataGridView3, dataGridView1, dataGridView6, searchColumnNameIndexAfterWET, Q_X_Value, Convert.ToDouble(Q_X_PercentageValue_NewTarget), dgvRange, TargetPercentage_Formula_Value, filterDataCount);
            var minPercentageValue = arrMinMax_get_PercentageLimit_Target[0];
            var maxPercentageValue = arrMinMax_get_PercentageLimit_Target[1];

            //
            var depColListBoxItems = lbDepCol.SelectedItems;
            var mustColListBoxItems = lbMustCol.SelectedItems;


            bool flag = false;
            int i = 0;
            for (int rows = 0; rows < dataGridView3.Rows.Count - 1; rows++)
            {
                //for (int col = 0; col < dataGridView3.Rows[rows].Cells.Count; col++)
                //{
                //if (i == searchColumnNameIndexAfterWET)
                //{
                bool isDependentColSatisfied = false;
                if (depColListBoxItems.Count > 0) // Optional
                {
                    if (cbNatureOfDeptCol.Text.Trim().ToUpper().ToUpper() == "AND")
                    {
                        isDependentColSatisfied = isDependentCol_Satisfy_AND(dataGridView3, lbDepCol, Q_X_Value, rows);
                    }
                    else if (cbNatureOfDeptCol.Text.Trim().ToUpper().ToUpper() == "OR")
                    {
                        isDependentColSatisfied = isDependentCol_Satisfy_OR(dataGridView3, lbDepCol, Q_X_Value, rows);
                    }
                    else
                    {
                        MessageBox.Show("Nature of dependent column not valid.");
                        return;
                    }
                }
                else
                {
                    isDependentColSatisfied = true;
                }

                string percentage = "";
                if (isDependentColSatisfied) // Validation for dependent column IF ALL Dependent column VALIDATION SATISFIED THEN CHANGE COLUMN
                {
                    var currentVal = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value.ToString();
                    //if (currentVal != Q_X_Value) // if Q1 value already same do nothing
                    //{

                    string workingColumnNature = mdgv.getColumnNature(getQ1_X_ColumnName);
                    //if (workingColumnNature == "s")
                    //{
                    //    dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = Q_X_Value;
                    //}
                    // if (workingColumnNature == "m")
                    //{
                    var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value.ToString();
                    var lst_GetAllCellData = Split(get_Q_X_ColumnData, Q_X_Value.Length);
                    if (!lst_GetAllCellData.Contains(Q_X_Value))
                    {
                        lst_GetAllCellData.Add(Q_X_Value);
                        var result = string.Join("", lst_GetAllCellData.ToArray());
                        if (result != "")
                        {
                            dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = result;
                        }
                        else
                        {
                            dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = Q_X_Value;
                        }
                    }
                    //}
                    percentage = mdgv.calculatePercentage(dataGridView3, Q_X_Value, searchColumnNameIndexAfterWET);
                    if (double.Parse(percentage) > maxPercentageValue)
                    {
                        dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = get_Q_X_ColumnData;
                    }
                    else
                    {
                        if (mustColListBoxItems.Count > 0) // Optional
                        {
                            assignValuesToMustColumn(lbMustCol, dataGridView3, Q_X_Value, rows);
                        }
                    }

                    if (double.Parse(percentage) >= minPercentageValue && double.Parse(percentage) <= maxPercentageValue)
                    {
                        flag = true;
                    }
                    //}
                    //}
                    i++;
                }
                if (flag)
                {
                    break;
                }
                //i = 0;
                //}
            }
        }

        public void decreasePercentage(DataGridView dataGridView3, ListBox lbDepCol, ListBox lbMustCol, string getQ1_X_ColumnName, string Q_X_Value, string Q_X_PercentageValue_NewTarget, DataGridView dataGridView1, DataGridView dataGridView6, DataGridView dgvRange, ComboBox cbNatureOfDeptCol, double TargetPercentage_Formula_Value)
        {

            MyDataGridView mdgv = new MyDataGridView();
            var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, getQ1_X_ColumnName);

            int filterDataCount = dataGridView3.Rows.Count - 1;
            //Getting Minimum and maximum values
            var arrMinMax_get_PercentageLimit_Target = calculatePercentageLimit(dataGridView3, dataGridView1, dataGridView6, searchColumnNameIndexAfterWET, Q_X_Value, Convert.ToDouble(Q_X_PercentageValue_NewTarget), dgvRange, TargetPercentage_Formula_Value, filterDataCount);
            var minPercentageValue = arrMinMax_get_PercentageLimit_Target[0];
            var maxPercentageValue = arrMinMax_get_PercentageLimit_Target[1];

            bool flag = false;
            //  int i = 0;
            for (int rows = 0; rows < dataGridView3.Rows.Count - 1; rows++)
            {
                var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value.ToString();
                var lst_GetAllCellData = Split(get_Q_X_ColumnData, Q_X_Value.Length);
                if (lst_GetAllCellData.Contains(Q_X_Value))
                {
                    lst_GetAllCellData.RemoveAll(x => x == Q_X_Value);
                    var result = string.Join("", lst_GetAllCellData.ToArray());
                    if (result != "")
                    {
                        dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = result;
                    }
                    else
                    {
                        dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = Q_X_Value;
                    }
                    var percentage = mdgv.calculatePercentage(dataGridView3, Q_X_Value, searchColumnNameIndexAfterWET);
                    if (double.Parse(percentage) < minPercentageValue)
                    {
                        dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = result + Q_X_Value;
                    }
                    if (double.Parse(percentage) >= minPercentageValue && double.Parse(percentage) <= maxPercentageValue)
                    {
                        flag = true;
                    }
                    //if (double.Parse(Q_X_PercentageValue_NewTarget) >= double.Parse(percentage))
                    //{
                    //    flag = true;
                    //}
                }
                if (flag)
                {
                    break;
                }
            }
            // ================== OLD CODE ================
            //MyDataGridView mdgv = new MyDataGridView();
            //var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, getQ1_X_ColumnName);
            //var depColListBoxItems = lbDepCol.SelectedItems;
            //var mustColListBoxItems = lbMustCol.SelectedItems;


            //bool flag = false;
            //int i = 0;
            //for (int rows = 0; rows < dataGridView3.Rows.Count - 1; rows++)
            //{
            //    //for (int col = 0; col < dataGridView3.Rows[rows].Cells.Count; col++)
            //    //{
            //    //if (i == searchColumnNameIndexAfterWET)
            //    //{
            //    bool isDependentColSatisfied = true;

            //    // Dependent column validation
            //    List<bool> lst_CheckAllValidation_ExistNotExist = new List<bool>();
            //    foreach (var item in depColListBoxItems)
            //    {
            //        var arr_existOrNot = item.ToString().Split('#');
            //        var value_existorNotExist = arr_existOrNot[1];
            //        var columnName_Q_X_existorNotExist = arr_existOrNot[0];
            //        ;
            //        var searchColumnNameIndexAfterWET_Q_X_existorNotExist = mdgv.searchColumnNameIndexAfterWET(dataGridView3, columnName_Q_X_existorNotExist);
            //        var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X_existorNotExist].Value.ToString();


            //        int formNoIndex = 1;
            //        var formNo = dataGridView3.Rows[rows].Cells[formNoIndex].Value.ToString(); //Form Number Value

            //        var lst_GetAllCellData = Split(get_Q_X_ColumnData, Q_X_Value.Length);
            //        if (value_existorNotExist.ToLower() == "exist")
            //        {
            //            // Exist
            //            if (lst_GetAllCellData.Contains(Q_X_Value))
            //            {
            //                lst_CheckAllValidation_ExistNotExist.Add(true);
            //            }
            //            else
            //            {
            //                lst_CheckAllValidation_ExistNotExist.Add(false);
            //            }
            //        }
            //        else
            //        {
            //            //Not exist
            //            if (!lst_GetAllCellData.Contains(Q_X_Value))
            //            {
            //                lst_CheckAllValidation_ExistNotExist.Add(true);
            //            }
            //            else
            //            {
            //                lst_CheckAllValidation_ExistNotExist.Add(false);
            //            }
            //        }
            //    }

            //    // Checking Dependent column if all conditions are true witch means "AND"
            //    foreach (var item in lst_CheckAllValidation_ExistNotExist)
            //    {
            //        if (!item)
            //        {
            //            isDependentColSatisfied = false;
            //        }
            //    }

            //    if (isDependentColSatisfied) // Validation for dependent column IF ALL Dependent column VALIDATION SATISFIED THEN CHANGE COLUMN
            //    {
            //        var currentVal = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value.ToString();

            //        // assign Values To Must Column
            //        assignValuesToMustColumn(lbMustCol, dataGridView3, Q_X_Value, rows);

            //        if (currentVal != Q_X_Value) // if Q1 value already same do nothing
            //        {
            //            var findRandomVal = Q_X_Value;
            //            if (IncreaseOrDecrease.increaseList.Count > 0)
            //            {
            //                List<IncreaseOrDecrease> tmpList = new List<IncreaseOrDecrease>();
            //                tmpList = IncreaseOrDecrease.increaseList.Where(x => x.query == query).ToList();
            //                Random r = new Random(); 
            //                int index = r.Next(tmpList.Count);
            //                findRandomVal = tmpList[index].columnValue;
            //            }
            //            dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = findRandomVal;
            //            var percentage = mdgv.calculatePercentage(dataGridView3, Q_X_Value, searchColumnNameIndexAfterWET);
            //            if (double.Parse(Q_X_PercentageValue_NewTarget) >= double.Parse(percentage))
            //            {
            //                flag = true;
            //                break;
            //            }
            //        }
            //        //}
            //        i++;
            //    }
            //    if (flag)
            //    {
            //        break;
            //    }
            //    //i = 0;
            //}
            //  }
        }

        public double[] calculatePercentageLimit(DataGridView dataGridView3, DataGridView dataGridView1, DataGridView dataGridView6, int searchColumnNameIndexAfterWET, string Q_X_Value, double Q_X_PercentageValue_NewTarget, DataGridView dgvRange, double TargetPercentage_Formula_Value, int columnCount)
        {
            var arrMinMax = new double[2];

            // if Value is Zero
            if (Q_X_PercentageValue_NewTarget <= 0)
            {
                arrMinMax[0] = 0;
                arrMinMax[1] = 0;
                return arrMinMax;
            }

            // If Value not Zero
            var get_Q1_ColumnData = getColumnData(dataGridView1, searchColumnNameIndexAfterWET);
            var getCurrent_Q1_ColumnData = getColumnData(dataGridView6, 0); // getting data from target file
            List<myExcel> countList = CalculateCountOf_Q_X_Using_Range(getCurrent_Q1_ColumnData, get_Q1_ColumnData);

            var getRangeColumnData = getColumnData(dgvRange, 0);
            var getRangePercentageColumnData = getColumnData(dgvRange, 1);
            List<Range> ranges = GettingRangesFromRangeFile(getRangeColumnData, getRangePercentageColumnData);

            var Q_X_Value_Count = get_Q_X_Value_Count(countList, Q_X_Value);
            //double limit = get_Q_X_Value_Limit(ranges, Q_X_Value_Count);

            double limit = Convert.ToDouble(ranges.Where(x => x.Min <= columnCount && x.Max >= columnCount).FirstOrDefault().Percentage);
            //myExcel.CountList.Where(x=>x.Count == sampleSizeCount).FirstOrDefault()

            var result = TargetPercentage_Formula_Value * limit;
            var minValue = (Q_X_PercentageValue_NewTarget - result).ToString();
            var maxValue = (Q_X_PercentageValue_NewTarget + result).ToString();
            minValue = minValue.ToString().Trim('-');
            maxValue = maxValue.ToString().Trim('-');
            arrMinMax[0] = Convert.ToDouble(minValue);
            arrMinMax[1] = Convert.ToDouble(maxValue);
            return arrMinMax;
        }
        public bool isDependentCol_Satisfy_AND(DataGridView dataGridView3, ListBox lbDepCol, string Q_X_Value, int rows)
        {
            //   bool isDependentColSatisfied = true;
            var depColListBoxItems = lbDepCol.SelectedItems;
            // Dependent column validation
            List<bool> lst_CheckAllValidation_ExistNotExist = new List<bool>();
            foreach (var item in depColListBoxItems)
            {
                var arr_existOrNot = item.ToString().Split('#');
                var value_existorNotExist = arr_existOrNot[1];
                var columnName_Q_X_existorNotExist = arr_existOrNot[0];
                ;
                var searchColumnNameIndexAfterWET_Q_X_existorNotExist = searchColumnNameIndexAfterWET(dataGridView3, columnName_Q_X_existorNotExist);
                var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X_existorNotExist].Value.ToString();


                int formNoIndex = 1;
                var formNo = dataGridView3.Rows[rows].Cells[formNoIndex].Value.ToString(); //Form Number Value

                var lst_GetAllCellData = Split(get_Q_X_ColumnData, Q_X_Value.Length);
                if (value_existorNotExist.ToLower() == "exist")
                {
                    // Exist
                    if (lst_GetAllCellData.Contains(Q_X_Value))
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(true);
                    }
                    else
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(false);
                        return false;
                    }
                }
                else
                {
                    //Not exist
                    if (!lst_GetAllCellData.Contains(Q_X_Value))
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(true);
                    }
                    else
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(false);
                        return false;
                    }
                }
            }

            // Checking Dependent column if all conditions are true witch means "AND"
            //foreach (var item in lst_CheckAllValidation_ExistNotExist)
            //{
            //    if (!item)
            //    {
            //        isDependentColSatisfied = false;
            //        break;
            //    }
            //}

            return true;
        }
        public bool isDependentCol_Satisfy_OR(DataGridView dataGridView3, ListBox lbDepCol, string Q_X_Value, int rows)
        {
            var depColListBoxItems = lbDepCol.SelectedItems;
            // Dependent column validation
            List<bool> lst_CheckAllValidation_ExistNotExist = new List<bool>();
            foreach (var item in depColListBoxItems)
            {
                var arr_existOrNot = item.ToString().Split('#');
                var value_existorNotExist = arr_existOrNot[1];
                var columnName_Q_X_existorNotExist = arr_existOrNot[0];
                ;
                var searchColumnNameIndexAfterWET_Q_X_existorNotExist = searchColumnNameIndexAfterWET(dataGridView3, columnName_Q_X_existorNotExist);
                var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X_existorNotExist].Value.ToString();


                int formNoIndex = 1;
                var formNo = dataGridView3.Rows[rows].Cells[formNoIndex].Value.ToString(); //Form Number Value

                var lst_GetAllCellData = Split(get_Q_X_ColumnData, Q_X_Value.Length);
                if (value_existorNotExist.ToLower() == "exist")
                {
                    // Exist
                    if (lst_GetAllCellData.Contains(Q_X_Value))
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(true);
                        return true; // Or Condition Satified 
                    }
                    else
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(false);
                    }
                }
                else
                {
                    //Not exist
                    if (!lst_GetAllCellData.Contains(Q_X_Value))
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(true);
                        return true; // Or Condition Satified 
                    }
                    else
                    {
                        lst_CheckAllValidation_ExistNotExist.Add(false);
                    }
                }
            }

            return false;
        }

        public int get_Q_X_Value_Count(List<myExcel> countList, string Q_X_Value)
        {
            int Q_X_Count = 0;
            foreach (var item in countList)
            {
                if (item.ColumnValue == Q_X_Value)
                {
                    Q_X_Count = item.Count;
                    break;
                }
            }
            return Q_X_Count;
        }
        public double get_Q_X_Value_Limit(List<Range> ranges, int Q_X_Value_Count)
        {
            double limit = 0;
            for (int i = 0; i < ranges.Count; i++)
            {
                var range = ranges[i];
                if (Q_X_Value_Count >= range.Min && Q_X_Value_Count <= range.Max) // checking if Q_X valuel lie with in range
                {
                    limit = Convert.ToDouble(range.Percentage);
                }
            }

            return limit;
        }

        public void SetTargetwithValidation(DataGridView target, DataGridView current, string workingColumnName)
        {
            for (int i = 1; i < target.Columns.Count; i++)
            {
                var targetColumn = getColumnData(target, i);
                decimal columnSum = 0;
                foreach (var item in targetColumn)
                {
                    columnSum += Convert.ToDecimal(item);
                }
                var workingColumnNature = getColumnNature(workingColumnName);
                if (workingColumnNature == "s")
                {
                    if (columnSum < 100 || columnSum > 100)
                    {
                        decimal changedValueSum = 0;
                        decimal notChangedValueSum = 0;

                        var currentColumn = getColumnData(current, i);

                        for (int k = 0; k < currentColumn.Count(); k++)
                        {
                            if (currentColumn[k] == targetColumn[k])
                            {
                                notChangedValueSum += Convert.ToDecimal(targetColumn[k]);
                            }
                            else
                            {
                                changedValueSum += Convert.ToDecimal(targetColumn[k]);
                            }
                        }
                        decimal difference = 100 - notChangedValueSum;
                        decimal division = 0;
                        if (changedValueSum > 0)
                        {
                           division = difference / changedValueSum;
                        }



                        List<string> changedValues = new List<string>();
                        for (int k = 0; k < currentColumn.Count(); k++)
                        {
                            if (currentColumn[k] != targetColumn[k])
                            {
                                changedValues.Add((Convert.ToDecimal(targetColumn[k]) * division).ToString());
                            }
                            else
                            {
                                changedValues.Add(targetColumn[k].ToString());
                            }
                        }

                        //foreach (var item in targetColumn)
                        //{

                        //    changedValues.Add( (Convert.ToDecimal( item )* division).ToString());
                        //}

                        insertDataInColumn(i, changedValues, target);

                    }
                }
                else if (workingColumnNature == "m")
                {
                    if (columnSum < 100)
                    {
                        decimal changedValueSum = 0;
                        decimal notChangedValueSum = 0;

                        var currentColumn = getColumnData(current, i);

                        for (int k = 0; k < currentColumn.Count(); k++)
                        {
                            if (currentColumn[k] == targetColumn[k])
                            {
                                notChangedValueSum += Convert.ToDecimal(targetColumn[k]);
                            }
                            else
                            {
                                changedValueSum += Convert.ToDecimal(targetColumn[k]);
                            }
                        }
                        decimal difference = 100 - notChangedValueSum;
                        decimal division = difference / changedValueSum;
                        List<string> changedValues = new List<string>();
                        for (int k = 0; k < currentColumn.Count(); k++)
                        {
                            if (currentColumn[k] != targetColumn[k])
                            {
                                changedValues.Add((Convert.ToDecimal(targetColumn[k]) * division).ToString());
                            }
                            else
                            {
                                changedValues.Add(targetColumn[k].ToString());
                            }
                        }

                        //foreach (var item in targetColumn)
                        //{

                        //    changedValues.Add( (Convert.ToDecimal( item )* division).ToString());
                        //}

                        insertDataInColumn(i, changedValues, target);

                    }
                }

            }
        }

        public void insertDataInColumn(int index, List<string> data, DataGridView dataGridView)
        {
            for (int i = 0; i < data.Count(); i++)
            {
                dataGridView.Rows[i].Cells[index].Value = data[i];
            }

        }

        public string getColumnNature(string columnName)
        {
            string[] model = columnName.Split('_');
            return model[model.Length - 1].ToLower().ToString();

        }
        //=================================================== Form 2 ==================================

        public void changePercentages(List<string> lst, DataGridView dataGridView, DataGridView dataGridView3)
        {
            dataGridView3.Refresh();
            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            Percentages.percentagesList = new List<List<Percentages>>();

            var AllCombosList = lst;

            //string query = "";
            //int u = 0;
            foreach (var query in AllCombosList)
            {
                List<Percentages> percentagesList = new List<Percentages>();
                // Filtration
                new Form2().filteration(query, dataGridView);
            }
        }

    }
}
