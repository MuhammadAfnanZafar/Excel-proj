﻿namespace ExcelProject
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.lblFileName = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.panelDropdown = new System.Windows.Forms.Panel();
            this.tbDataChar = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbWorkingColumn = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.rbRelational = new System.Windows.Forms.RadioButton();
            this.rbNonRelational = new System.Windows.Forms.RadioButton();
            this.btnProcessData = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lbDepCol = new System.Windows.Forms.ListBox();
            this.lbMustCol = new System.Windows.Forms.ListBox();
            this.lblError = new System.Windows.Forms.Label();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.label7 = new System.Windows.Forms.Label();
            this.dataGridView4 = new System.Windows.Forms.DataGridView();
            this.label8 = new System.Windows.Forms.Label();
            this.dataGridView5 = new System.Windows.Forms.DataGridView();
            this.label9 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.dataGridView6 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1079, 31);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(151, 28);
            this.button1.TabIndex = 0;
            this.button1.Text = "Upload Data File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblFileName
            // 
            this.lblFileName.AutoSize = true;
            this.lblFileName.Location = new System.Drawing.Point(13, 142);
            this.lblFileName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(71, 17);
            this.lblFileName.TabIndex = 1;
            this.lblFileName.Text = "File Name";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(1079, 67);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(151, 28);
            this.button2.TabIndex = 2;
            this.button2.Text = "Download Data File";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // panelDropdown
            // 
            this.panelDropdown.AutoScroll = true;
            this.panelDropdown.Location = new System.Drawing.Point(15, 161);
            this.panelDropdown.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panelDropdown.Name = "panelDropdown";
            this.panelDropdown.Size = new System.Drawing.Size(1217, 184);
            this.panelDropdown.TabIndex = 6;
            // 
            // tbDataChar
            // 
            this.tbDataChar.Location = new System.Drawing.Point(16, 54);
            this.tbDataChar.Margin = new System.Windows.Forms.Padding(4);
            this.tbDataChar.Name = "tbDataChar";
            this.tbDataChar.Size = new System.Drawing.Size(160, 22);
            this.tbDataChar.TabIndex = 7;
            this.tbDataChar.Text = "2";
            this.tbDataChar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tbDataChar_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 34);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(111, 17);
            this.label2.TabIndex = 8;
            this.label2.Text = "Data Characters";
            // 
            // cbWorkingColumn
            // 
            this.cbWorkingColumn.FormattingEnabled = true;
            this.cbWorkingColumn.Location = new System.Drawing.Point(16, 105);
            this.cbWorkingColumn.Margin = new System.Windows.Forms.Padding(4);
            this.cbWorkingColumn.Name = "cbWorkingColumn";
            this.cbWorkingColumn.Size = new System.Drawing.Size(160, 24);
            this.cbWorkingColumn.TabIndex = 9;
            this.cbWorkingColumn.Text = "-- Select --";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 88);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(111, 17);
            this.label3.TabIndex = 10;
            this.label3.Text = "Working Column";
            // 
            // rbRelational
            // 
            this.rbRelational.AutoSize = true;
            this.rbRelational.Checked = true;
            this.rbRelational.Location = new System.Drawing.Point(662, 54);
            this.rbRelational.Margin = new System.Windows.Forms.Padding(4);
            this.rbRelational.Name = "rbRelational";
            this.rbRelational.Size = new System.Drawing.Size(92, 21);
            this.rbRelational.TabIndex = 11;
            this.rbRelational.TabStop = true;
            this.rbRelational.Text = "Relational";
            this.rbRelational.UseVisualStyleBackColor = true;
            // 
            // rbNonRelational
            // 
            this.rbNonRelational.AutoSize = true;
            this.rbNonRelational.Location = new System.Drawing.Point(662, 83);
            this.rbNonRelational.Margin = new System.Windows.Forms.Padding(4);
            this.rbNonRelational.Name = "rbNonRelational";
            this.rbNonRelational.Size = new System.Drawing.Size(123, 21);
            this.rbNonRelational.TabIndex = 12;
            this.rbNonRelational.TabStop = true;
            this.rbNonRelational.Text = "Non-Relational";
            this.rbNonRelational.UseVisualStyleBackColor = true;
            // 
            // btnProcessData
            // 
            this.btnProcessData.Location = new System.Drawing.Point(1079, 102);
            this.btnProcessData.Margin = new System.Windows.Forms.Padding(4);
            this.btnProcessData.Name = "btnProcessData";
            this.btnProcessData.Size = new System.Drawing.Size(153, 28);
            this.btnProcessData.TabIndex = 13;
            this.btnProcessData.Text = "Process Data";
            this.btnProcessData.UseVisualStyleBackColor = true;
            this.btnProcessData.Click += new System.EventHandler(this.Button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(16, 351);
            this.button4.Margin = new System.Windows.Forms.Padding(4);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(151, 28);
            this.button4.TabIndex = 14;
            this.button4.Text = "Upload Range File";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(195, 34);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(129, 17);
            this.label5.TabIndex = 17;
            this.label5.Text = "Dependent Column";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(1081, 351);
            this.button5.Margin = new System.Windows.Forms.Padding(4);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(151, 28);
            this.button5.TabIndex = 18;
            this.button5.Text = "Save to Excel";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(420, 36);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 17);
            this.label6.TabIndex = 19;
            this.label6.Text = "Must Column";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(16, 415);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(217, 278);
            this.dataGridView1.TabIndex = 21;
            // 
            // lbDepCol
            // 
            this.lbDepCol.FormattingEnabled = true;
            this.lbDepCol.ItemHeight = 16;
            this.lbDepCol.Location = new System.Drawing.Point(198, 54);
            this.lbDepCol.Name = "lbDepCol";
            this.lbDepCol.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lbDepCol.Size = new System.Drawing.Size(219, 68);
            this.lbDepCol.TabIndex = 22;
            // 
            // lbMustCol
            // 
            this.lbMustCol.FormattingEnabled = true;
            this.lbMustCol.ItemHeight = 16;
            this.lbMustCol.Location = new System.Drawing.Point(423, 54);
            this.lbMustCol.Name = "lbMustCol";
            this.lbMustCol.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lbMustCol.Size = new System.Drawing.Size(219, 68);
            this.lbMustCol.TabIndex = 23;
            // 
            // lblError
            // 
            this.lblError.AutoSize = true;
            this.lblError.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblError.ForeColor = System.Drawing.Color.Red;
            this.lblError.Location = new System.Drawing.Point(13, 9);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(50, 18);
            this.lblError.TabIndex = 24;
            this.lblError.Text = "Error: ";
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(239, 415);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.Size = new System.Drawing.Size(217, 278);
            this.dataGridView2.TabIndex = 25;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 395);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 17);
            this.label1.TabIndex = 26;
            this.label1.Text = "Current";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(236, 395);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 17);
            this.label4.TabIndex = 27;
            this.label4.Text = "Previous";
            // 
            // dataGridView3
            // 
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Location = new System.Drawing.Point(462, 415);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.RowHeadersWidth = 51;
            this.dataGridView3.RowTemplate.Height = 24;
            this.dataGridView3.Size = new System.Drawing.Size(217, 278);
            this.dataGridView3.TabIndex = 28;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(459, 395);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 17);
            this.label7.TabIndex = 29;
            this.label7.Text = "Process";
            // 
            // dataGridView4
            // 
            this.dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView4.Location = new System.Drawing.Point(685, 415);
            this.dataGridView4.Name = "dataGridView4";
            this.dataGridView4.RowHeadersWidth = 51;
            this.dataGridView4.RowTemplate.Height = 24;
            this.dataGridView4.Size = new System.Drawing.Size(266, 278);
            this.dataGridView4.TabIndex = 30;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(682, 395);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(102, 17);
            this.label8.TabIndex = 31;
            this.label8.Text = "Report Current";
            // 
            // dataGridView5
            // 
            this.dataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView5.Location = new System.Drawing.Point(957, 415);
            this.dataGridView5.Name = "dataGridView5";
            this.dataGridView5.RowHeadersWidth = 51;
            this.dataGridView5.RowTemplate.Height = 24;
            this.dataGridView5.Size = new System.Drawing.Size(266, 278);
            this.dataGridView5.TabIndex = 32;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(954, 395);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(110, 17);
            this.label9.TabIndex = 33;
            this.label9.Text = "Report Previous";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(175, 351);
            this.button3.Margin = new System.Windows.Forms.Padding(4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(151, 28);
            this.button3.TabIndex = 34;
            this.button3.Text = "Generate Target File";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(1226, 395);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(50, 17);
            this.label10.TabIndex = 36;
            this.label10.Text = "Traget";
            // 
            // dataGridView6
            // 
            this.dataGridView6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView6.Location = new System.Drawing.Point(1229, 415);
            this.dataGridView6.Name = "dataGridView6";
            this.dataGridView6.RowHeadersWidth = 51;
            this.dataGridView6.RowTemplate.Height = 24;
            this.dataGridView6.Size = new System.Drawing.Size(266, 278);
            this.dataGridView6.TabIndex = 35;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1448, 705);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.dataGridView6);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.dataGridView5);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.dataGridView4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.dataGridView3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.lblError);
            this.Controls.Add(this.lbMustCol);
            this.Controls.Add(this.lbDepCol);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.btnProcessData);
            this.Controls.Add(this.rbNonRelational);
            this.Controls.Add(this.rbRelational);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbWorkingColumn);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbDataChar);
            this.Controls.Add(this.panelDropdown);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.lblFileName);
            this.Controls.Add(this.button1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panelDropdown;
        private System.Windows.Forms.TextBox tbDataChar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbWorkingColumn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RadioButton rbRelational;
        private System.Windows.Forms.RadioButton rbNonRelational;
        private System.Windows.Forms.Button btnProcessData;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ListBox lbDepCol;
        private System.Windows.Forms.ListBox lbMustCol;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DataGridView dataGridView4;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DataGridView dataGridView5;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.DataGridView dataGridView6;
    }
}