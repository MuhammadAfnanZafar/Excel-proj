using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject.Model
{
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

            var total = brandSum / calulateWETSum;
            return (total * 100).ToString();
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
    }
}
