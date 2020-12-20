﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject.Model
{
    class MyDataGridView
    {
        public List<string> getColumnNames(DataGridView dataGridView1, int rowIndex)
        {
            List<string> lst = new List<string>();

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
                if (text.ToLower() == "wet")
                {
                    break;
                }
                lst.Add(text);
            }

            return lst;
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
        public List<List<string>> getRowData(DataGridView dataGridView1)
        {
            List<List<string>> lst = new List<List<string>>();

            int rowCount = dataGridView1.Rows.Count;
            int headerWETIndex = getWETIndex(dataGridView1) + 1;

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
                sum += Convert.ToInt32(dataGridView1.Rows[i].Cells[headerWETIndex].Value);
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
        public string calculatePercentage(DataGridView dgv, string val)
        {
            MyDataGridView mdgv = new MyDataGridView();
            double calulateWETSum = mdgv.getSumOfWET(dgv);

            var getRowData = mdgv.getRowData(dgv);
            double brandSum = 0;
            for (int i = 0; i < getRowData.Count(); i++)
            {
                var item = getRowData[i];
                var q1_1_value = item[item.Count() - 1]; // getting last value
                if (q1_1_value.ToString() == val) // comparing last value with comboBox value i.e Q1_1
                {
                    brandSum += double.Parse(item[item.Count() - 2]);
                }
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

    }
}
