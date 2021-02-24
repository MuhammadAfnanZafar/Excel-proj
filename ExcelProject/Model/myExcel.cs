using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject.Model
{
    class myExcel
    {
        public string ColumnValue { get; set; }
        public string ColumnNameOrQuery { get; set; }
        public int Count { get; set; }
        public static List<myExcel> CountList = new List<myExcel>();


        public void ToCsV(DataGridView dgv, string name, string age, string address, string title, string filename)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            //Adding the Columns.
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            //Adding the Rows.
            int c = 0;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (dgv.Rows.Count-1 == c)
                {
                    break;
                }
                c++;
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                }
            }

            //Exporting to Excel.
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Customers");

                //Set the color of Header Row.
                //A resembles First Column while C resembles Third column.
                wb.Worksheet(1).Cells("A1:C1").Style.Fill.BackgroundColor = XLColor.DarkGreen;
                for (int i = 1; i <= dt.Rows.Count; i++)
                {
                    //A resembles First Column while C resembles Third column.
                    //Header row is at Position 1 and hence First row starts from Index 2.
                    string cellRange = string.Format("A{0}:C{0}", i + 1);
                    if (i % 2 != 0)
                    {
                        wb.Worksheet(1).Cells(cellRange).Style.Fill.BackgroundColor = XLColor.GreenYellow;
                    }
                    else
                    {
                        wb.Worksheet(1).Cells(cellRange).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }

                }
                //Adjust widths of Columns.
                wb.Worksheet(1).Columns().AdjustToContents();

                //Save the Excel file.
                wb.SaveAs(filename);
            }
        }
        private Worksheet FindSheet(Workbook workbook, string sheet_name)
        {
            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == sheet_name) return sheet;
            }
            return null;
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
