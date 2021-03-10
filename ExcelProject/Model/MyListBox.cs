using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject.Model
{
    class MyListBox
    {

        public string ColumnName { get; set; }
        public List<string> Data { get; set; }

        int colIndex = 2;


        public MyListBox()
        {
            Data = new List<string>();
        }

        ListBox[] lb;

        public void createListBox(Panel panelDropdown, DataGridView dataGridView1, List<string> lstCoulumnNames)
        {
            panelDropdown.Controls.Clear();

            int listBoxCount = lstCoulumnNames.Count;
            lb = new ListBox[listBoxCount];
            int width = 150;
            int height = 100;
            int spacing = 5;
            int colIndex = this.colIndex;
            for (int i = 0; i <= listBoxCount - 1; ++i)
            {
                Label mylab = new Label();
                mylab.Text = (lstCoulumnNames[i] != null || lstCoulumnNames[i] != "") ? lstCoulumnNames[i] : "-- Unknown --";
                mylab.Location = new System.Drawing.Point((i * (width + 2)) + spacing, 0);
                mylab.AutoSize = true;
                mylab.Name = "label" + string.Format("{0}_{1:N}", i, Guid.NewGuid());
                panelDropdown.Controls.Add(mylab);

                ListBox newBox = new ListBox();
                //newBox.Text = "-- Select --";
                newBox.SelectionMode = SelectionMode.MultiSimple;
                newBox.Size = new Size(width, height);
                newBox.Location = new System.Drawing.Point((i * (width + 2)) + spacing, 15);
                newBox.Name = "listBox" + string.Format("{0}_{1:N}", i, Guid.NewGuid());
                lb[i] = newBox;
                panelDropdown.Controls.Add(newBox);

                // Adding items in combo box
                assignItemsToListBox(dataGridView1, newBox, colIndex);
                colIndex++;
            }
        }
        public void assignHeaderNameAfterWETToListBox(DataGridView dataGridView1, ListBox listBox, List<string> Data)
        {
            MyDataGridView mdgv = new MyDataGridView();
            //columnData = columnData.Distinct().ToList();
            foreach (var item in Data)
            {
                listBox.Items.Add(item);
            }
        }

        public void assignItemsToListBox(DataGridView dataGridView1, ListBox listBox, int colIndex)
        {
            MyDataGridView mdgv = new MyDataGridView();
            List<string> columnData = mdgv.getColumnData(dataGridView1, colIndex);
            columnData = columnData.Distinct().ToList();
            columnData = columnData.OrderBy(x => x).ToList();
            columnData.RemoveAll(string.IsNullOrEmpty);
            foreach (var item in columnData)
            {
                listBox.Items.Add(item);
            }
        }
    }
}
