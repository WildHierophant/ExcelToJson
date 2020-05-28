namespace ExcelToJson
{
    partial class FormMain
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxExcelURL = new System.Windows.Forms.TextBox();
            this.buttonStartingExport = new System.Windows.Forms.Button();
            this.textBoxJsonURL = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.listBox = new System.Windows.Forms.ListBox();
            this.buttonExcelURL = new System.Windows.Forms.Button();
            this.buttonJsonURL = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.buttonClearListBoxMessage = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(382, 364);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "Madeby:WildHierophant";
            // 
            // textBoxExcelURL
            // 
            this.textBoxExcelURL.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxExcelURL.Location = new System.Drawing.Point(81, 262);
            this.textBoxExcelURL.Name = "textBoxExcelURL";
            this.textBoxExcelURL.Size = new System.Drawing.Size(360, 21);
            this.textBoxExcelURL.TabIndex = 2;
            // 
            // buttonStartingExport
            // 
            this.buttonStartingExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonStartingExport.Location = new System.Drawing.Point(322, 329);
            this.buttonStartingExport.Name = "buttonStartingExport";
            this.buttonStartingExport.Size = new System.Drawing.Size(88, 23);
            this.buttonStartingExport.TabIndex = 3;
            this.buttonStartingExport.Text = "导出Json";
            this.buttonStartingExport.UseVisualStyleBackColor = true;
            this.buttonStartingExport.Click += new System.EventHandler(this.buttonStartingExport_Click);
            // 
            // textBoxJsonURL
            // 
            this.textBoxJsonURL.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxJsonURL.Location = new System.Drawing.Point(81, 296);
            this.textBoxJsonURL.Name = "textBoxJsonURL";
            this.textBoxJsonURL.Size = new System.Drawing.Size(360, 21);
            this.textBoxJsonURL.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 265);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "表格地址:";
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 299);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "导出地址:";
            // 
            // listBox
            // 
            this.listBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox.FormattingEnabled = true;
            this.listBox.ItemHeight = 12;
            this.listBox.Location = new System.Drawing.Point(18, 12);
            this.listBox.Name = "listBox";
            this.listBox.Size = new System.Drawing.Size(486, 232);
            this.listBox.TabIndex = 9;
            // 
            // buttonExcelURL
            // 
            this.buttonExcelURL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonExcelURL.Location = new System.Drawing.Point(457, 260);
            this.buttonExcelURL.Name = "buttonExcelURL";
            this.buttonExcelURL.Size = new System.Drawing.Size(47, 23);
            this.buttonExcelURL.TabIndex = 10;
            this.buttonExcelURL.Text = "...";
            this.buttonExcelURL.UseVisualStyleBackColor = true;
            this.buttonExcelURL.Click += new System.EventHandler(this.buttonExcelURL_Click);
            // 
            // buttonJsonURL
            // 
            this.buttonJsonURL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonJsonURL.Location = new System.Drawing.Point(457, 294);
            this.buttonJsonURL.Name = "buttonJsonURL";
            this.buttonJsonURL.Size = new System.Drawing.Size(47, 23);
            this.buttonJsonURL.TabIndex = 11;
            this.buttonJsonURL.Text = "...";
            this.buttonJsonURL.UseVisualStyleBackColor = true;
            this.buttonJsonURL.Click += new System.EventHandler(this.buttonJsonURL_Click);
            // 
            // buttonClearListBoxMessage
            // 
            this.buttonClearListBoxMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonClearListBoxMessage.Location = new System.Drawing.Point(416, 329);
            this.buttonClearListBoxMessage.Name = "buttonClearListBoxMessage";
            this.buttonClearListBoxMessage.Size = new System.Drawing.Size(88, 23);
            this.buttonClearListBoxMessage.TabIndex = 12;
            this.buttonClearListBoxMessage.Text = "清理消息";
            this.buttonClearListBoxMessage.UseVisualStyleBackColor = true;
            this.buttonClearListBoxMessage.Click += new System.EventHandler(this.buttonClearListBoxMessage_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(525, 385);
            this.Controls.Add(this.buttonClearListBoxMessage);
            this.Controls.Add(this.buttonJsonURL);
            this.Controls.Add(this.buttonExcelURL);
            this.Controls.Add(this.listBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxJsonURL);
            this.Controls.Add(this.buttonStartingExport);
            this.Controls.Add(this.textBoxExcelURL);
            this.Controls.Add(this.label1);
            this.Name = "FormMain";
            this.Text = "ExcelToJson工具";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormMain_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxExcelURL;
        private System.Windows.Forms.Button buttonStartingExport;
        private System.Windows.Forms.TextBox textBoxJsonURL;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox listBox;
        private System.Windows.Forms.Button buttonExcelURL;
        private System.Windows.Forms.Button buttonJsonURL;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button buttonClearListBoxMessage;
    }
}

