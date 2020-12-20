using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProject.Model
{
    class MyComboBox
    {
        ComboBox[] cmb;

        public void createComboBox(Panel panelDropdown, DataGridView dataGridView1, List<string> lstCoulumnNames)
        {
            panelDropdown.Controls.Clear();

            int comboBoxCount = lstCoulumnNames.Count;
            cmb = new ComboBox[comboBoxCount];
            int width = 150;
            int height = 35;
            int spacing = 5;
            for (int i = 0; i <= comboBoxCount - 1; ++i)
            {
                Label mylab = new Label();
                mylab.Text = (lstCoulumnNames[i] != null || lstCoulumnNames[i] != "") ? lstCoulumnNames[i] : "-- Unknown --";
                mylab.Location = new System.Drawing.Point((i * (width + 2)) + spacing, 0);
                mylab.AutoSize = true;
                panelDropdown.Controls.Add(mylab);

                ComboBox newBox = new ComboBox();
                newBox.Text = "-- Select --";
                newBox.Size = new Size(width, height);
                newBox.Location = new System.Drawing.Point((i * (width + 2)) + spacing, 15);
                newBox.Name = "comboBox" + string.Format("{0}_{1:N}", i, Guid.NewGuid());
                cmb[i] = newBox;
                panelDropdown.Controls.Add(newBox);
                
                // Adding items in combo box
                assignItemsToComboBox(dataGridView1, newBox, i);
            }
        }

        public void setQ1_1ComboBox(DataGridView dataGridView1, ComboBox comboBoxQ1_1)
        {
            comboBoxQ1_1.Items.Clear();
            int q1_1_Index = 0;
            int colCount = dataGridView1.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                string text = dataGridView1.Columns[i].HeaderText;
                if (text.ToLower() == "wet")
                {
                    q1_1_Index = i + 1;
                    break;
                }
            }

            // Adding items in combo box
            assignItemsToComboBox(dataGridView1, comboBoxQ1_1, q1_1_Index);
        }

        public void assignItemsToComboBox(DataGridView dataGridView1, ComboBox comboBox , int colIndex)
        {
            MyDataGridView mdgv = new MyDataGridView();
            List<string> columnData = mdgv.getColumnData(dataGridView1, colIndex);
            columnData = columnData.Distinct().ToList();
            comboBox.Items.Add("none");
            foreach (var item in columnData)
            {
                comboBox.Items.Add(item);
            }
        }
    }
}
