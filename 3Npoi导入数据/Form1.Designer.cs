namespace _03Npoi导入数据
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
            this.ExportData = new System.Windows.Forms.Button();
            this.InportDataBySpire = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ExportData
            // 
            this.ExportData.Location = new System.Drawing.Point(150, 92);
            this.ExportData.Name = "ExportData";
            this.ExportData.Size = new System.Drawing.Size(168, 40);
            this.ExportData.TabIndex = 0;
            this.ExportData.Text = "Npoi通过模板导出数据";
            this.ExportData.UseVisualStyleBackColor = true;
            this.ExportData.Click += new System.EventHandler(this.ExportData_Click);
            // 
            // InportDataBySpire
            // 
            this.InportDataBySpire.Location = new System.Drawing.Point(150, 27);
            this.InportDataBySpire.Name = "InportDataBySpire";
            this.InportDataBySpire.Size = new System.Drawing.Size(115, 40);
            this.InportDataBySpire.TabIndex = 1;
            this.InportDataBySpire.Text = "Spire导出数据";
            this.InportDataBySpire.UseVisualStyleBackColor = true;
            this.InportDataBySpire.Click += new System.EventHandler(this.InportDataBySpire_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(150, 161);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(231, 40);
            this.button1.TabIndex = 2;
            this.button1.Text = "数码视讯CRM系统npoi导出数据";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 380);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.InportDataBySpire);
            this.Controls.Add(this.ExportData);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ExportData;
        private System.Windows.Forms.Button InportDataBySpire;
        private System.Windows.Forms.Button button1;
    }
}

