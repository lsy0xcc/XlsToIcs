namespace XlsToIcs
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.XlsFilePath = new System.Windows.Forms.TextBox();
            this.XlsFileButton = new System.Windows.Forms.Button();
            this.ToIcsFileButton = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // XlsFilePath
            // 
            this.XlsFilePath.Location = new System.Drawing.Point(12, 33);
            this.XlsFilePath.Name = "XlsFilePath";
            this.XlsFilePath.Size = new System.Drawing.Size(408, 31);
            this.XlsFilePath.TabIndex = 0;
            // 
            // XlsFileButton
            // 
            this.XlsFileButton.Location = new System.Drawing.Point(426, 33);
            this.XlsFileButton.Name = "XlsFileButton";
            this.XlsFileButton.Size = new System.Drawing.Size(129, 31);
            this.XlsFileButton.TabIndex = 1;
            this.XlsFileButton.Text = "选择文件";
            this.XlsFileButton.UseVisualStyleBackColor = true;
            this.XlsFileButton.Click += new System.EventHandler(this.XlsFileButton_Click);
            // 
            // ToIcsFileButton
            // 
            this.ToIcsFileButton.Location = new System.Drawing.Point(12, 180);
            this.ToIcsFileButton.Name = "ToIcsFileButton";
            this.ToIcsFileButton.Size = new System.Drawing.Size(173, 31);
            this.ToIcsFileButton.TabIndex = 2;
            this.ToIcsFileButton.Text = "生成ics文件";
            this.ToIcsFileButton.UseVisualStyleBackColor = true;
            this.ToIcsFileButton.Click += new System.EventHandler(this.ToIcsFileButton_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(12, 116);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 31);
            this.dateTimePicker1.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(347, 21);
            this.label1.TabIndex = 4;
            this.label1.Text = "1.选择从教务系统导出的课程表文件";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(221, 21);
            this.label2.TabIndex = 5;
            this.label2.Text = "2.选择学期的开始日期";
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(573, 226);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.ToIcsFileButton);
            this.Controls.Add(this.XlsFileButton);
            this.Controls.Add(this.XlsFilePath);
            this.Name = "Form1";
            this.Text = "XlsToIcs";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox XlsFilePath;
        private System.Windows.Forms.Button XlsFileButton;
        private System.Windows.Forms.Button ToIcsFileButton;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

