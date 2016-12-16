namespace Attempt1
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
            this.listBox_log = new System.Windows.Forms.ListBox();
            this.comboBoxCourse = new System.Windows.Forms.ComboBox();
            this.comboBoxBlks = new System.Windows.Forms.ComboBox();
            this.button_export = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_upperBound = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // listBox_log
            // 
            this.listBox_log.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.listBox_log.FormattingEnabled = true;
            this.listBox_log.HorizontalScrollbar = true;
            this.listBox_log.ItemHeight = 12;
            this.listBox_log.Location = new System.Drawing.Point(0, 137);
            this.listBox_log.Name = "listBox_log";
            this.listBox_log.Size = new System.Drawing.Size(313, 208);
            this.listBox_log.TabIndex = 0;
            // 
            // comboBoxCourse
            // 
            this.comboBoxCourse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxCourse.FormattingEnabled = true;
            this.comboBoxCourse.Location = new System.Drawing.Point(91, 12);
            this.comboBoxCourse.Name = "comboBoxCourse";
            this.comboBoxCourse.Size = new System.Drawing.Size(121, 20);
            this.comboBoxCourse.TabIndex = 1;
            this.comboBoxCourse.SelectedIndexChanged += new System.EventHandler(this.comboBoxCourse_SelectedIndexChanged);
            // 
            // comboBoxBlks
            // 
            this.comboBoxBlks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxBlks.FormattingEnabled = true;
            this.comboBoxBlks.Location = new System.Drawing.Point(91, 38);
            this.comboBoxBlks.Name = "comboBoxBlks";
            this.comboBoxBlks.Size = new System.Drawing.Size(121, 20);
            this.comboBoxBlks.TabIndex = 2;
            this.comboBoxBlks.SelectedIndexChanged += new System.EventHandler(this.comboBoxBlks_SelectedIndexChanged);
            // 
            // button_export
            // 
            this.button_export.Location = new System.Drawing.Point(91, 91);
            this.button_export.Name = "button_export";
            this.button_export.Size = new System.Drawing.Size(75, 23);
            this.button_export.TabIndex = 3;
            this.button_export.Text = "导出excel";
            this.button_export.UseVisualStyleBackColor = true;
            this.button_export.Click += new System.EventHandler(this.button_export_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "科目";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "题块";
            // 
            // textBox_upperBound
            // 
            this.textBox_upperBound.Location = new System.Drawing.Point(91, 64);
            this.textBox_upperBound.Name = "textBox_upperBound";
            this.textBox_upperBound.Size = new System.Drawing.Size(121, 21);
            this.textBox_upperBound.TabIndex = 6;
            this.textBox_upperBound.Text = "3";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 7;
            this.label3.Text = "<=分数上界";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(313, 345);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox_upperBound);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button_export);
            this.Controls.Add(this.comboBoxBlks);
            this.Controls.Add(this.comboBoxCourse);
            this.Controls.Add(this.listBox_log);
            this.Name = "Form1";
            this.Text = "低分仲裁卷信息导出";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBox_log;
        private System.Windows.Forms.ComboBox comboBoxCourse;
        private System.Windows.Forms.ComboBox comboBoxBlks;
        private System.Windows.Forms.Button button_export;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_upperBound;
        private System.Windows.Forms.Label label3;
    }
}

