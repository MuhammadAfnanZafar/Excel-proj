namespace ExcelProject
{
    partial class Form1
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
            this.btnCurrentFile = new System.Windows.Forms.Button();
            this.btnRecentFile = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnOpenRecentFile = new System.Windows.Forms.Button();
            this.btnOpenCurrentFile = new System.Windows.Forms.Button();
            this.panelDropdown = new System.Windows.Forms.Panel();
            this.btnSaveChanges = new System.Windows.Forms.Button();
            this.labelFileName = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.btnChangePercentage = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSaveToExcel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxPercentage = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxValue = new System.Windows.Forms.TextBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCurrentFile
            // 
            this.btnCurrentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCurrentFile.Location = new System.Drawing.Point(9, 40);
            this.btnCurrentFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnCurrentFile.Name = "btnCurrentFile";
            this.btnCurrentFile.Size = new System.Drawing.Size(126, 28);
            this.btnCurrentFile.TabIndex = 0;
            this.btnCurrentFile.Text = "Upload Current File";
            this.btnCurrentFile.UseVisualStyleBackColor = true;
            this.btnCurrentFile.Click += new System.EventHandler(this.btnCurrentFile_Click);
            // 
            // btnRecentFile
            // 
            this.btnRecentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRecentFile.Location = new System.Drawing.Point(140, 40);
            this.btnRecentFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnRecentFile.Name = "btnRecentFile";
            this.btnRecentFile.Size = new System.Drawing.Size(130, 28);
            this.btnRecentFile.TabIndex = 1;
            this.btnRecentFile.Text = "Upload Recent File";
            this.btnRecentFile.UseVisualStyleBackColor = true;
            this.btnRecentFile.Click += new System.EventHandler(this.btnRecentFile_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(9, 193);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(529, 288);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // btnOpenRecentFile
            // 
            this.btnOpenRecentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenRecentFile.Location = new System.Drawing.Point(944, 40);
            this.btnOpenRecentFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnOpenRecentFile.Name = "btnOpenRecentFile";
            this.btnOpenRecentFile.Size = new System.Drawing.Size(128, 28);
            this.btnOpenRecentFile.TabIndex = 4;
            this.btnOpenRecentFile.Text = "Open Recent File";
            this.btnOpenRecentFile.UseVisualStyleBackColor = true;
            this.btnOpenRecentFile.Click += new System.EventHandler(this.btnOpenRecentFile_Click);
            // 
            // btnOpenCurrentFile
            // 
            this.btnOpenCurrentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenCurrentFile.Location = new System.Drawing.Point(818, 40);
            this.btnOpenCurrentFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnOpenCurrentFile.Name = "btnOpenCurrentFile";
            this.btnOpenCurrentFile.Size = new System.Drawing.Size(121, 28);
            this.btnOpenCurrentFile.TabIndex = 3;
            this.btnOpenCurrentFile.Text = "Open Current File";
            this.btnOpenCurrentFile.UseVisualStyleBackColor = true;
            this.btnOpenCurrentFile.Click += new System.EventHandler(this.btnOpenCurrentFile_Click);
            // 
            // panelDropdown
            // 
            this.panelDropdown.Location = new System.Drawing.Point(9, 73);
            this.panelDropdown.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.panelDropdown.Name = "panelDropdown";
            this.panelDropdown.Size = new System.Drawing.Size(1062, 67);
            this.panelDropdown.TabIndex = 5;
            // 
            // btnSaveChanges
            // 
            this.btnSaveChanges.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveChanges.Location = new System.Drawing.Point(56, 15);
            this.btnSaveChanges.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnSaveChanges.Name = "btnSaveChanges";
            this.btnSaveChanges.Size = new System.Drawing.Size(165, 28);
            this.btnSaveChanges.TabIndex = 6;
            this.btnSaveChanges.Text = "Save Current File Excel";
            this.btnSaveChanges.UseVisualStyleBackColor = true;
            this.btnSaveChanges.Click += new System.EventHandler(this.btnSaveChanges_Click);
            // 
            // labelFileName
            // 
            this.labelFileName.AutoSize = true;
            this.labelFileName.Location = new System.Drawing.Point(7, 168);
            this.labelFileName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelFileName.Name = "labelFileName";
            this.labelFileName.Size = new System.Drawing.Size(60, 13);
            this.labelFileName.TabIndex = 7;
            this.labelFileName.Text = "File Name: ";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(274, 50);
            this.textBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(208, 20);
            this.textBox1.TabIndex = 8;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(272, 34);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Search";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(976, 145);
            this.button1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(94, 28);
            this.button1.TabIndex = 10;
            this.button1.Text = "Filter";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnChangePercentage
            // 
            this.btnChangePercentage.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChangePercentage.Location = new System.Drawing.Point(778, 55);
            this.btnChangePercentage.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnChangePercentage.Name = "btnChangePercentage";
            this.btnChangePercentage.Size = new System.Drawing.Size(130, 28);
            this.btnChangePercentage.TabIndex = 11;
            this.btnChangePercentage.Text = "Change Percentage";
            this.btnChangePercentage.UseVisualStyleBackColor = true;
            this.btnChangePercentage.Click += new System.EventHandler(this.btnChangePercentage_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSaveToExcel);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBoxPercentage);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBoxValue);
            this.groupBox1.Controls.Add(this.btnChangePercentage);
            this.groupBox1.Controls.Add(this.btnSaveChanges);
            this.groupBox1.Location = new System.Drawing.Point(9, 486);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(1059, 93);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Actions";
            // 
            // btnSaveToExcel
            // 
            this.btnSaveToExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveToExcel.Location = new System.Drawing.Point(924, 15);
            this.btnSaveToExcel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnSaveToExcel.Name = "btnSaveToExcel";
            this.btnSaveToExcel.Size = new System.Drawing.Size(130, 28);
            this.btnSaveToExcel.TabIndex = 17;
            this.btnSaveToExcel.Text = "Save To Excel";
            this.btnSaveToExcel.UseVisualStyleBackColor = true;
            this.btnSaveToExcel.Click += new System.EventHandler(this.btnSaveToExcel_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(500, 29);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(81, 20);
            this.label4.TabIndex = 16;
            this.label4.Text = "Change: ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(727, 16);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "Percentage";
            // 
            // textBoxPercentage
            // 
            this.textBoxPercentage.Location = new System.Drawing.Point(729, 32);
            this.textBoxPercentage.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.textBoxPercentage.Name = "textBoxPercentage";
            this.textBoxPercentage.Size = new System.Drawing.Size(180, 20);
            this.textBoxPercentage.TabIndex = 14;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(579, 16);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Value";
            // 
            // textBoxValue
            // 
            this.textBoxValue.Location = new System.Drawing.Point(581, 32);
            this.textBoxValue.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.textBoxValue.Name = "textBoxValue";
            this.textBoxValue.ReadOnly = true;
            this.textBoxValue.Size = new System.Drawing.Size(134, 20);
            this.textBoxValue.TabIndex = 12;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(543, 193);
            this.dataGridView2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.Size = new System.Drawing.Size(525, 288);
            this.dataGridView2.TabIndex = 15;
            this.dataGridView2.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellClick);
            this.dataGridView2.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(9, 7);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(40, 15);
            this.label5.TabIndex = 16;
            this.label5.Text = "Error: ";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1111, 588);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.labelFileName);
            this.Controls.Add(this.panelDropdown);
            this.Controls.Add(this.btnOpenRecentFile);
            this.Controls.Add(this.btnOpenCurrentFile);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnRecentFile);
            this.Controls.Add(this.btnCurrentFile);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCurrentFile;
        private System.Windows.Forms.Button btnRecentFile;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnOpenRecentFile;
        private System.Windows.Forms.Button btnOpenCurrentFile;
        private System.Windows.Forms.Panel panelDropdown;
        private System.Windows.Forms.Button btnSaveChanges;
        private System.Windows.Forms.Label labelFileName;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnChangePercentage;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxPercentage;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxValue;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnSaveToExcel;
    }
}

