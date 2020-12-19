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
            this.btnCalculate = new System.Windows.Forms.Button();
            this.comboBoxQ1_1 = new System.Windows.Forms.ComboBox();
            this.labelTotal = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCurrentFile
            // 
            this.btnCurrentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCurrentFile.Location = new System.Drawing.Point(12, 49);
            this.btnCurrentFile.Name = "btnCurrentFile";
            this.btnCurrentFile.Size = new System.Drawing.Size(168, 35);
            this.btnCurrentFile.TabIndex = 0;
            this.btnCurrentFile.Text = "Upload Current File";
            this.btnCurrentFile.UseVisualStyleBackColor = true;
            this.btnCurrentFile.Click += new System.EventHandler(this.btnCurrentFile_Click);
            // 
            // btnRecentFile
            // 
            this.btnRecentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRecentFile.Location = new System.Drawing.Point(186, 49);
            this.btnRecentFile.Name = "btnRecentFile";
            this.btnRecentFile.Size = new System.Drawing.Size(174, 35);
            this.btnRecentFile.TabIndex = 1;
            this.btnRecentFile.Text = "Upload Recent File";
            this.btnRecentFile.UseVisualStyleBackColor = true;
            this.btnRecentFile.Click += new System.EventHandler(this.btnRecentFile_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 221);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1409, 394);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // btnOpenRecentFile
            // 
            this.btnOpenRecentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenRecentFile.Location = new System.Drawing.Point(1251, 49);
            this.btnOpenRecentFile.Name = "btnOpenRecentFile";
            this.btnOpenRecentFile.Size = new System.Drawing.Size(170, 35);
            this.btnOpenRecentFile.TabIndex = 4;
            this.btnOpenRecentFile.Text = "Open Recent File";
            this.btnOpenRecentFile.UseVisualStyleBackColor = true;
            this.btnOpenRecentFile.Click += new System.EventHandler(this.btnOpenRecentFile_Click);
            // 
            // btnOpenCurrentFile
            // 
            this.btnOpenCurrentFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenCurrentFile.Location = new System.Drawing.Point(1084, 49);
            this.btnOpenCurrentFile.Name = "btnOpenCurrentFile";
            this.btnOpenCurrentFile.Size = new System.Drawing.Size(161, 35);
            this.btnOpenCurrentFile.TabIndex = 3;
            this.btnOpenCurrentFile.Text = "Open Current File";
            this.btnOpenCurrentFile.UseVisualStyleBackColor = true;
            this.btnOpenCurrentFile.Click += new System.EventHandler(this.btnOpenCurrentFile_Click);
            // 
            // panelDropdown
            // 
            this.panelDropdown.Location = new System.Drawing.Point(12, 90);
            this.panelDropdown.Name = "panelDropdown";
            this.panelDropdown.Size = new System.Drawing.Size(1412, 84);
            this.panelDropdown.TabIndex = 5;
            // 
            // btnSaveChanges
            // 
            this.btnSaveChanges.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveChanges.Location = new System.Drawing.Point(1247, 621);
            this.btnSaveChanges.Name = "btnSaveChanges";
            this.btnSaveChanges.Size = new System.Drawing.Size(174, 35);
            this.btnSaveChanges.TabIndex = 6;
            this.btnSaveChanges.Text = "Save Changes";
            this.btnSaveChanges.UseVisualStyleBackColor = true;
            this.btnSaveChanges.Click += new System.EventHandler(this.btnSaveChanges_Click);
            // 
            // labelFileName
            // 
            this.labelFileName.AutoSize = true;
            this.labelFileName.Location = new System.Drawing.Point(12, 201);
            this.labelFileName.Name = "labelFileName";
            this.labelFileName.Size = new System.Drawing.Size(79, 17);
            this.labelFileName.TabIndex = 7;
            this.labelFileName.Text = "File Name: ";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(366, 62);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(276, 22);
            this.textBox1.TabIndex = 8;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(363, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 17);
            this.label1.TabIndex = 9;
            this.label1.Text = "Search";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(1302, 180);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(119, 35);
            this.button1.TabIndex = 10;
            this.button1.Text = "Filter";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnCalculate
            // 
            this.btnCalculate.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCalculate.Location = new System.Drawing.Point(237, 633);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(174, 35);
            this.btnCalculate.TabIndex = 11;
            this.btnCalculate.Text = "Calculate";
            this.btnCalculate.UseVisualStyleBackColor = true;
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);
            // 
            // comboBoxQ1_1
            // 
            this.comboBoxQ1_1.FormattingEnabled = true;
            this.comboBoxQ1_1.Location = new System.Drawing.Point(15, 639);
            this.comboBoxQ1_1.Name = "comboBoxQ1_1";
            this.comboBoxQ1_1.Size = new System.Drawing.Size(216, 24);
            this.comboBoxQ1_1.TabIndex = 12;
            this.comboBoxQ1_1.Text = "-- Select --";
            // 
            // labelTotal
            // 
            this.labelTotal.AutoSize = true;
            this.labelTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTotal.Location = new System.Drawing.Point(12, 666);
            this.labelTotal.Name = "labelTotal";
            this.labelTotal.Size = new System.Drawing.Size(31, 25);
            this.labelTotal.TabIndex = 13;
            this.labelTotal.Text = "%";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1433, 746);
            this.Controls.Add(this.labelTotal);
            this.Controls.Add(this.comboBoxQ1_1);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.labelFileName);
            this.Controls.Add(this.btnSaveChanges);
            this.Controls.Add(this.panelDropdown);
            this.Controls.Add(this.btnOpenRecentFile);
            this.Controls.Add(this.btnOpenCurrentFile);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnRecentFile);
            this.Controls.Add(this.btnCurrentFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
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
        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.ComboBox comboBoxQ1_1;
        private System.Windows.Forms.Label labelTotal;
    }
}

