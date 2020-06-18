namespace TableGenerater
{
    partial class TableGenerater
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
            this.labelProjectPath = new System.Windows.Forms.Label();
            this.labelExcelPath = new System.Windows.Forms.Label();
            this.comboBoxProject = new System.Windows.Forms.ComboBox();
            this.comboBoxExcel = new System.Windows.Forms.ComboBox();
            this.buttonFindProject = new System.Windows.Forms.Button();
            this.buttonFindExcel = new System.Windows.Forms.Button();
            this.buttonGenerater = new System.Windows.Forms.Button();
            this.textBox = new System.Windows.Forms.TextBox();
            this.checkBoxStreamingAssets = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // labelProjectPath
            // 
            this.labelProjectPath.AutoSize = true;
            this.labelProjectPath.Location = new System.Drawing.Point(19, 36);
            this.labelProjectPath.Name = "labelProjectPath";
            this.labelProjectPath.Size = new System.Drawing.Size(81, 12);
            this.labelProjectPath.TabIndex = 0;
            this.labelProjectPath.Text = "Project Path :";
            // 
            // labelExcelPath
            // 
            this.labelExcelPath.AutoSize = true;
            this.labelExcelPath.Location = new System.Drawing.Point(25, 81);
            this.labelExcelPath.Name = "labelExcelPath";
            this.labelExcelPath.Size = new System.Drawing.Size(74, 12);
            this.labelExcelPath.TabIndex = 1;
            this.labelExcelPath.Text = "Excel Path :";
            // 
            // comboBoxProject
            // 
            this.comboBoxProject.FormattingEnabled = true;
            this.comboBoxProject.Location = new System.Drawing.Point(106, 33);
            this.comboBoxProject.Name = "comboBoxProject";
            this.comboBoxProject.Size = new System.Drawing.Size(276, 20);
            this.comboBoxProject.TabIndex = 2;
            // 
            // comboBoxExcel
            // 
            this.comboBoxExcel.FormattingEnabled = true;
            this.comboBoxExcel.Location = new System.Drawing.Point(106, 78);
            this.comboBoxExcel.Name = "comboBoxExcel";
            this.comboBoxExcel.Size = new System.Drawing.Size(276, 20);
            this.comboBoxExcel.TabIndex = 3;
            // 
            // buttonFindProject
            // 
            this.buttonFindProject.Location = new System.Drawing.Point(388, 31);
            this.buttonFindProject.Name = "buttonFindProject";
            this.buttonFindProject.Size = new System.Drawing.Size(41, 23);
            this.buttonFindProject.TabIndex = 4;
            this.buttonFindProject.Text = "....";
            this.buttonFindProject.UseVisualStyleBackColor = true;
            this.buttonFindProject.Click += new System.EventHandler(this.buttonFindProject_Click);
            // 
            // buttonFindExcel
            // 
            this.buttonFindExcel.Location = new System.Drawing.Point(388, 76);
            this.buttonFindExcel.Name = "buttonFindExcel";
            this.buttonFindExcel.Size = new System.Drawing.Size(41, 23);
            this.buttonFindExcel.TabIndex = 5;
            this.buttonFindExcel.Text = "....";
            this.buttonFindExcel.UseVisualStyleBackColor = true;
            this.buttonFindExcel.Click += new System.EventHandler(this.buttonFindExcel_Click);
            // 
            // buttonGenerater
            // 
            this.buttonGenerater.Location = new System.Drawing.Point(354, 117);
            this.buttonGenerater.Name = "buttonGenerater";
            this.buttonGenerater.Size = new System.Drawing.Size(75, 23);
            this.buttonGenerater.TabIndex = 6;
            this.buttonGenerater.Text = "생성";
            this.buttonGenerater.UseVisualStyleBackColor = true;
            this.buttonGenerater.Click += new System.EventHandler(this.buttonGenerater_Click);
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(5, 146);
            this.textBox.Name = "textBox";
            this.textBox.ReadOnly = true;
            this.textBox.Size = new System.Drawing.Size(422, 21);
            this.textBox.TabIndex = 7;
            // 
            // checkBoxStreamingAssets
            // 
            this.checkBoxStreamingAssets.AutoSize = true;
            this.checkBoxStreamingAssets.Checked = true;
            this.checkBoxStreamingAssets.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxStreamingAssets.Location = new System.Drawing.Point(21, 117);
            this.checkBoxStreamingAssets.Name = "checkBoxStreamingAssets";
            this.checkBoxStreamingAssets.Size = new System.Drawing.Size(168, 16);
            this.checkBoxStreamingAssets.TabIndex = 8;
            this.checkBoxStreamingAssets.Text = "Copy to StreamingAssets";
            this.checkBoxStreamingAssets.UseVisualStyleBackColor = true;
            // 
            // TableGenerater
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(439, 170);
            this.Controls.Add(this.checkBoxStreamingAssets);
            this.Controls.Add(this.textBox);
            this.Controls.Add(this.buttonGenerater);
            this.Controls.Add(this.buttonFindExcel);
            this.Controls.Add(this.buttonFindProject);
            this.Controls.Add(this.comboBoxExcel);
            this.Controls.Add(this.comboBoxProject);
            this.Controls.Add(this.labelExcelPath);
            this.Controls.Add(this.labelProjectPath);
            this.Name = "TableGenerater";
            this.Text = "TableGenerater";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelProjectPath;
        private System.Windows.Forms.Label labelExcelPath;
        private System.Windows.Forms.ComboBox comboBoxProject;
        private System.Windows.Forms.ComboBox comboBoxExcel;
        private System.Windows.Forms.Button buttonFindProject;
        private System.Windows.Forms.Button buttonFindExcel;
        private System.Windows.Forms.Button buttonGenerater;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.CheckBox checkBoxStreamingAssets;
    }
}

