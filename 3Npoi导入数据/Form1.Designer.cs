﻿namespace _03Npoi导入数据
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
            this.SuspendLayout();
            // 
            // ExportData
            // 
            this.ExportData.Location = new System.Drawing.Point(150, 93);
            this.ExportData.Name = "ExportData";
            this.ExportData.Size = new System.Drawing.Size(115, 40);
            this.ExportData.TabIndex = 0;
            this.ExportData.Text = "Npoi导入数据";
            this.ExportData.UseVisualStyleBackColor = true;
            this.ExportData.Click += new System.EventHandler(this.ExportData_Click);
            // 
            // InportDataBySpire
            // 
            this.InportDataBySpire.Location = new System.Drawing.Point(150, 163);
            this.InportDataBySpire.Name = "InportDataBySpire";
            this.InportDataBySpire.Size = new System.Drawing.Size(115, 40);
            this.InportDataBySpire.TabIndex = 1;
            this.InportDataBySpire.Text = "Spire导出数据";
            this.InportDataBySpire.UseVisualStyleBackColor = true;
            this.InportDataBySpire.Click += new System.EventHandler(this.InportDataBySpire_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 380);
            this.Controls.Add(this.InportDataBySpire);
            this.Controls.Add(this.ExportData);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ExportData;
        private System.Windows.Forms.Button InportDataBySpire;
    }
}

