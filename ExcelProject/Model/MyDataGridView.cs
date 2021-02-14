using System;
using System.Collections.Generic;
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
            return (from Match m in Regex.Matches(str, @"\d{1," + chunkSize + "}")
                    select m.Value).ToList();
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
                for (int j = 0; j < rowDataList.Count; j++)
                {
                    var item2 = rowDataList[j];
                    if (item2.ToString() == val) // comparing last value with comboBox value i.e Q1_1
                    {
                        brandSum += double.Parse(item[wetIndex]);
                    }
                }
                //var item = getRowData[i];
                //var q1_1_value = item[item.Count() - 1]; // getting last value

                //if (q1_1_value.ToString() == val) // comparing last value with comboBox value i.e Q1_1
                //{
                //    brandSum += double.Parse(item[item.Count() - 2]);
                //}
            }
            if (calulateWETSum <=0)
            {
                return 0.ToString();
            }

            var total = brandSum / calulateWETSum;
            //if (total<=0)
            //{
            //    return "0";
            //}
            var temp =  (total * 100).ToString();
            if (val == 16.ToString())
            {

            }

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
        public DataGridView CopyDataGridView(DataGridView dgv_org)
        {
            DataGridView dgv_copy = new DataGridView();
            try
            {
                if (dgv_copy.Columns.Count == 0)
                {
                    foreach (DataGridViewColumn dgvc in dgv_org.Columns)
                    {
                        dgv_copy.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                    }
                }

                DataGridViewRow row = new DataGridViewRow();

                for (int i = 0; i < dgv_org.Rows.Count; i++)
                {
                    row = (DataGridViewRow)dgv_org.Rows[i].Clone();
                    int intColIndex = 0;
                    foreach (DataGridViewCell cell in dgv_org.Rows[i].Cells)
                    {
                        row.Cells[intColIndex].Value = cell.Value;
                        intColIndex++;
                    }
                    dgv_copy.Rows.Add(row);
                }
                dgv_copy.AllowUserToAddRows = false;
                dgv_copy.Refresh();

            }
            catch (Exception ex)
            {
                //cf.ShowExceptionErrorMsg("Copy DataGridViw", ex);
                MessageBox.Show(ex.Message);
            }
            return dgv_copy;
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
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                int formNoIndex = 1;
                var formNo = row.Cells[formNoIndex].Value.ToString(); //Form Number Value

                // Get filtered file row data
                var filteredFileRowData = getSpecificRowData(dataGridView3, row.Index);

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
                dataGridView.Rows[rowIndex].Cells[i].Value = item;
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

        public void assignValuesToMustColumn(ListBox lbMustCol, DataGridView dataGridView3, string Q_X_Value, int rows)
        {
            MyDataGridView mdgv = new MyDataGridView();
            var result = "";
            var mustColListBoxItems = lbMustCol.SelectedItems;
            foreach (var item in mustColListBoxItems)
            {
                var Q_X_ColumnName = item.ToString();
                var searchColumnNameIndexAfterWET_Q_X = mdgv.searchColumnNameIndexAfterWET(dataGridView3, Q_X_ColumnName);
                //var getColumnData = mdgv.getColumnData(dataGridView3, searchColumnNameIndexAfterWET_Q_X);
                var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value.ToString();
                var lst_GetAllCellData = Split(get_Q_X_ColumnData, Q_X_Value.Length);
                // Not Exist
                if (!lst_GetAllCellData.Contains(Q_X_Value))
                {
                    lst_GetAllCellData.Add(Q_X_Value);
                    result = string.Join("", lst_GetAllCellData.ToArray());
                    dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value = result;
                    return;
                }
                result = string.Join("", lst_GetAllCellData.ToArray());
                dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X].Value = result;
            }
        }
        public void increasePercentage(DataGridView dataGridView3, ListBox lbDepCol, ListBox lbMustCol, string getQ1_X_ColumnName, string Q_X_Value, string Q_X_PercentageValue_NewTarget, DataGridView dataGridView1)
        {
            MyDataGridView mdgv = new MyDataGridView();
            var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, getQ1_X_ColumnName);
            var depColListBoxItems = lbDepCol.SelectedItems;
            var mustColListBoxItems = lbMustCol.SelectedItems;


            bool flag = false;
            int i = 0;
            for (int rows = 0; rows < dataGridView3.Rows.Count - 1; rows++)
            {
                for (int col = 0; col < dataGridView3.Rows[rows].Cells.Count; col++)
                {
                    //if (i == searchColumnNameIndexAfterWET)
                    //{
                    bool isDependentColSatisfied = true;

                    // Dependent column validation
                    List<bool> lst_CheckAllValidation_ExistNotExist = new List<bool>();
                    foreach (var item in depColListBoxItems)
                    {
                        var arr_existOrNot = item.ToString().Split('#');
                        var value_existorNotExist = arr_existOrNot[1];
                        var columnName_Q_X_existorNotExist = arr_existOrNot[0];
                        ;
                        var searchColumnNameIndexAfterWET_Q_X_existorNotExist = mdgv.searchColumnNameIndexAfterWET(dataGridView3, columnName_Q_X_existorNotExist);
                        var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X_existorNotExist].Value.ToString();
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
                            }
                        }
                    }

                    // Checking Dependent column if all conditions are true witch means "AND"
                    foreach (var item in lst_CheckAllValidation_ExistNotExist)
                    {
                        if (!item)
                        {
                            isDependentColSatisfied = false;
                        }
                    }

                    if (isDependentColSatisfied) // Validation for dependent column IF ALL Dependent column VALIDATION SATISFIED THEN CHANGE COLUMN
                    {
                        var currentVal = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value.ToString();

                        // assign Values To Must Column
                        assignValuesToMustColumn(lbMustCol, dataGridView3, Q_X_Value, rows);

                        if (currentVal != Q_X_Value) // if Q1 value already same do nothing
                        {
                            dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = Q_X_Value;
                            var percentage = mdgv.calculatePercentage(dataGridView3, Q_X_Value, searchColumnNameIndexAfterWET);
                            if (double.Parse(Q_X_PercentageValue_NewTarget) <= double.Parse(percentage))
                            {
                                flag = true;
                                break;
                            }
                        }
                        //}
                        i++;
                    }
                    if (flag)
                    {
                        // Assign and replace values of current file i.e. (datagridview1)
                        AssignValuesToCurrentFile(dataGridView1, dataGridView3);
                        break;
                    }
                    //i = 0;
                }
            }


        }

        public void decreasePercentage(DataGridView dataGridView3, ListBox lbDepCol, ListBox lbMustCol, string getQ1_X_ColumnName, string Q_X_Value, string Q_X_PercentageValue_NewTarget, DataGridView dataGridView1)
        {
            MyDataGridView mdgv = new MyDataGridView();
            var searchColumnNameIndexAfterWET = mdgv.searchColumnNameIndexAfterWET(dataGridView3, getQ1_X_ColumnName);
            var depColListBoxItems = lbDepCol.SelectedItems;
            var mustColListBoxItems = lbMustCol.SelectedItems;


            bool flag = false;
            int i = 0;
            for (int rows = 0; rows < dataGridView3.Rows.Count - 1; rows++)
            {
                for (int col = 0; col < dataGridView3.Rows[rows].Cells.Count; col++)
                {
                    //if (i == searchColumnNameIndexAfterWET)
                    //{
                    bool isDependentColSatisfied = true;

                    // Dependent column validation
                    List<bool> lst_CheckAllValidation_ExistNotExist = new List<bool>();
                    foreach (var item in depColListBoxItems)
                    {
                        var arr_existOrNot = item.ToString().Split('#');
                        var value_existorNotExist = arr_existOrNot[1];
                        var columnName_Q_X_existorNotExist = arr_existOrNot[0];
                        ;
                        var searchColumnNameIndexAfterWET_Q_X_existorNotExist = mdgv.searchColumnNameIndexAfterWET(dataGridView3, columnName_Q_X_existorNotExist);
                        var get_Q_X_ColumnData = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET_Q_X_existorNotExist].Value.ToString();
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
                            }
                        }
                    }

                    // Checking Dependent column if all conditions are true witch means "AND"
                    foreach (var item in lst_CheckAllValidation_ExistNotExist)
                    {
                        if (!item)
                        {
                            isDependentColSatisfied = false;
                        }
                    }

                    if (isDependentColSatisfied) // Validation for dependent column IF ALL Dependent column VALIDATION SATISFIED THEN CHANGE COLUMN
                    {
                        var currentVal = dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value.ToString();
                        // Is ka baad se kaam karna h

                        // assign Values To Must Column
                        assignValuesToMustColumn(lbMustCol, dataGridView3, Q_X_Value, rows);

                        if (currentVal != Q_X_Value) // if Q1 value already same do nothing
                        {
                            var getQ1_X_Data_Current = mdgv.getColumnData(dataGridView1, searchColumnNameIndexAfterWET);
                            dataGridView3.Rows[rows].Cells[searchColumnNameIndexAfterWET].Value = Q_X_Value;
                            var percentage = mdgv.calculatePercentage(dataGridView3, Q_X_Value, searchColumnNameIndexAfterWET);
                            if (double.Parse(Q_X_PercentageValue_NewTarget) >= double.Parse(percentage))
                            {
                                flag = true;
                                break;
                            }
                        }
                        //}
                        i++;
                    }
                    if (flag)
                    {
                        // Assign and replace values of current file i.e. (datagridview1)
                        AssignValuesToCurrentFile(dataGridView1, dataGridView3);
                        break;
                    }
                    //i = 0;
                }
            }


        }

        public void SetTargetwithValidation(DataGridView target,DataGridView current)
        {
            /////// single / multiple
            for (int i = 1; i < target.Columns.Count; i++)
            {
                var targetColumn = getColumnData(target, i);
                decimal columnSum = 0;
                foreach (var item in targetColumn)
                {
                    columnSum += Convert.ToDecimal(item);
                }
                
                if (columnSum < 100)
                {
                    decimal changedValueSum = 0;
                    decimal notChangedValueSum = 0;

                    var currentColumn = getColumnData(current,i);

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

        public void insertDataInColumn(int index,List<string> data,DataGridView dataGridView)
        {
            for (int i = 0; i < data.Count(); i++)
            {
                dataGridView.Rows[i].Cells[index].Value = data[i];
            }
            
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
