using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _03Npoi导入数据
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region   npoi导入数据

        private void ExportData_Click(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 将excel表中的数据导入到datatable表中
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelSheetImportToDataTable(string fileName)
        {
            //这里注意  .xls和.xlsx两种类型的文件处理方式不一样   
            //.xls    
            //.xlsx   XSSFWorkBook
            return new DataTable();
        }



        #endregion

        #region Spire导出数据

        private void InportDataBySpire_Click(object sender, EventArgs e)
        {
            //创建工作簿对象
            Workbook workbook = new Workbook();
            //得到第一个Sheet页
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Range["A2"].Text = "你瞅啥？";
            //没有文件路径  就保存到应用程序路径下，即exe文件所在路径
            string fileName = "Test.pdf";
            workbook.SaveToPdf(fileName);

        }

        #endregion





    }
}
