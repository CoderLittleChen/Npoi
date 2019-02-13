using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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

        /// <summary>
        /// 将excel表中的数据导入到datatable表中
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelSheetImportToDataTable(string fileName)
        {
            return new DataTable();
        }


        /// <summary>
        /// 通过模板导出数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportData_Click(object sender, EventArgs e)
        {
            //这里注意  .xls和.xlsx两种类型的文件处理方式不一样   
            //.xls    
            //.xlsx   XSSFWorkBook

            ////1、点击按钮弹出文件选择框
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            ////设置文件打开的初始路径
            //openFileDialog.InitialDirectory = @"C:\Users\sdm\Desktop";
            ////设置找到的文件类型
            //openFileDialog.Filter = "Excel";
            ////设置默认显示的文件类型


            //文件路径
            string filePath = @"C:\Users\sdm\Desktop\模板.xlsx";
            //拿到文件扩展名
            string fileExtensionName = Path.GetExtension(filePath);
            //创建workBook对象
            IWorkbook workBook;
            try
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    if (fileExtensionName == ".xls")
                    {
                        workBook = new HSSFWorkbook(fileStream);
                    }
                    else
                    {
                        workBook = new XSSFWorkbook(fileStream);
                    }
                    ISheet sheet = workBook.GetSheet("明细");
                    for (int i = 0; i < 3; i++)
                    {
                        IRow rowField = sheet.CreateRow(i + 2);
                        ICell cell = rowField.CreateCell(0);
                        cell.SetCellValue("你瞅啥？");
                    }

                    //怎样给合并后的单元格 赋值？ 
                    IRow rowTtitle = sheet.GetRow(0);
                    ICell celltitle = rowTtitle.GetCell(0);
                    celltitle.SetCellValue("测试单元格重新赋值是否会改变样式");

                }
                using (FileStream fileStream = new FileStream(@"C:\Users\sdm\Desktop\读取模板写入数据.xls", FileMode.OpenOrCreate, FileAccess.Write))
                {
                    workBook.Write(fileStream);
                    fileStream.Flush();
                    fileStream.Close();
                }

            }
            catch (Exception)
            {
                throw;
            }


        }

        #endregion


        #region Spire导出数据

        private void InportDataBySpire_Click(object sender, EventArgs e)
        {
            //创建工作簿对象
            Workbook workbook = new Workbook();
            //得到第一个Sheet页
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Range["A1:C5"].Text = "你瞅啥？";
            //没有文件路径  就保存到应用程序路径下，即exe文件所在路径

            //保存为xlsx文件，需要在参数中选择Excel的版本

            sheet.Protect("123", SheetProtectionType.FormattingColumns);
            string fileName = "Test.xlsx";
            workbook.SaveToFile(fileName, ExcelVersion.Version2013);


            //将Excel保存为pdf    
            //string fileName = "Test.pdf";
            //workbook.SaveToPdf(fileName);

            //将sheet保存为图片   注意保存为sheet对象
            //string fileName = "Test.jpg";
            //sheet.SaveToImage(fileName, 1, 1, 5, 3);

        }


        #endregion


        /// <summary>
        /// 数码视讯CRM  通过npoi导出数据代码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {

            #region    未使用模板
            ///// <summary>
            ///// 开始（InitParas后执行）
            ///// </summary>
            //public override void Begin()
            //{
            //    PageHandle.Write(ExportExcel());
            //}


            ///// <summary>
            ///// 将生成excel的路径返回
            ///// </summary>
            ///// <returns></returns>
            //public string ExportExcel()
            //{
            //    string arrListStr = PageHandle.Request["arr"].ToString();
            //    arrListStr = arrListStr.Substring(1, arrListStr.Length - 2).Replace('\"', '\'');

            //    string firstDan = string.Empty;
            //    int index = arrListStr.IndexOf(",");
            //    if (index == -1)
            //    {
            //        firstDan = arrListStr;
            //    }
            //    else
            //    {
            //        firstDan = arrListStr.Substring(0, index);
            //    }
            //    if (!IsAuditFinish(firstDan.Substring(1, firstDan.Length - 2)))
            //    {
            //        return "0";
            //    }
            //    string departAlias = GetDepartNameByDan(firstDan);
            //    DataTable dt = GetDataTable(arrListStr);
            //    return GetWorkbook(dt, departAlias, 5184, firstDan.Substring(1, 6), firstDan.Substring(1, firstDan.Length - 2));
            //}


            ///// <summary>
            ///// 查询该条数据是否审核完成
            ///// </summary>
            ///// <param name="dan"></param>
            ///// <returns></returns>
            //public bool IsAuditFinish(string dan)
            //{
            //    string sql = string.Format(" select  count(*)  from  CUS_Table_5183  a  where  a.Status=3 and  a.Dan='{0}'  ", dan);
            //    int n = SqlHelper.ExecuteScalar(1, sql).ToInt32();
            //    if (n > 0)
            //    {
            //        return true;
            //    }
            //    return false;
            //}


            ///// <summary>
            ///// 获取需要导出的全部数据   
            ///// </summary>
            ///// <returns></returns>
            //public DataTable GetDataTable(string arrListStr)
            //{
            //    //接收参数     
            //    //De_29059Alias,De_29060,De_29061,De_31824,De_29062,De_31825,De_29063,
            //    //De_31827,De_29064,De_31828,De_31838,De_31839,De_29065,De_29066,
            //    //De_31829,De_29067,De_31830,De_29068,De_31831,De_29069,De_31826,De_31848,
            //    //De_31849,De_29071,Memo
            //    if (!string.IsNullOrEmpty(arrListStr))
            //    {
            //        string sqlQueryUserDetails = string.Format(" select " +
            //            " De_29060 as  '工号'  ,De_29059Alias  as  '姓名',De_29061  as '事假（H）' ,De_31824  as  '事假明细', " +
            //            " De_29062  as  '病假（H）'  ,De_31825   as   '病假明细'  ,De_29063   as   '工作日缓发（H）', " +
            //            " De_31827  as  '工作日缓发明细' ,De_29064  as  '工作日加班缓发（H）'  ,De_31828  as  '工作日加班缓发明细'  ," +
            //            " De_31838  as  '休息日和节假日加班缓发（H）' ,De_31839  as  '休息日和节假日加班缓发明细' ,De_29065   as   '迟到（次）'  ," +
            //            " De_32218  as  '迟到明细',De_29067 as  '休息日加班（H）',De_31830  as  '休息日加班明细', " +
            //            " De_29066  as  '工作日加班（H）',De_31829 as  '工作日加班明细' ," +
            //            " De_29068  as  '节假日加班（H）',De_31831 as  '节假日加班明细' ," +
            //            " De_29069  as  '年假（天）',De_31826  as  '年假明细',De_31848   as  '调休假总计（H）'  ," +
            //            " De_31849  as  '调休假明细',De_32219 as '当月累计调休（H）',De_32220  as   '当月累计调休明细',  " +
            //            " a.Memo  as  '备注' " +
            //            " from   CUS_Table_5183_Details  a   " +
            //            " inner  join  CUS_Table_5183  b   on  a.InnerId=b.Id " +
            //            " where  b.Dan  in  ({0})    order by    '工号'  asc   ; ", arrListStr);
            //        return SqlHelper.ExecuteDataset(1, sqlQueryUserDetails).Tables[0];
            //    }
            //    return new DataTable();
            //}


            ///// <summary>
            ///// 设置工作簿的格式
            ///// </summary>
            ///// <returns></returns>
            ///// <param name="currentAttdanceDate"></param>
            ///// <param name="departName"></param>
            ///// <param name="dt"></param>
            ///// <param name="moduleId"></param>
            ///// <param name="dan">主表的dan号</param>
            //public string GetWorkbook(DataTable dt, string departName, int moduleId, string currentAttdanceDate, string dan)
            //{
            //    //这里应该是直接通过npoi读取模板文件，格式已经调整好  
            //    IWorkbook workbook = new HSSFWorkbook();
            //    ISheet sheet = workbook.CreateSheet("明细");

            //    #region 创建样式   Font  Style
            //    IDataFormat dataFormat = workbook.CreateDataFormat();
            //    IFont fontTitle = workbook.CreateFont();
            //    fontTitle.FontName = "等线";
            //    fontTitle.FontHeight = 15;
            //    fontTitle.FontHeightInPoints = 15;

            //    IFont fontField = workbook.CreateFont();
            //    fontField.FontName = "等线";
            //    fontField.FontHeight = 10;
            //    fontField.FontHeightInPoints = 10;


            //    //创建style   四种   表名   字段名   字段间隔显示
            //    //表名的样式
            //    ICellStyle cellStyleTitle = workbook.CreateCellStyle();
            //    //居中
            //    cellStyleTitle.Alignment = HorizontalAlignment.Center;
            //    cellStyleTitle.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleTitle.BorderBottom = BorderStyle.Thin;
            //    cellStyleTitle.BorderLeft = BorderStyle.Thin;
            //    cellStyleTitle.BorderRight = BorderStyle.Thin;
            //    cellStyleTitle.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleTitle.SetFont(fontTitle);

            //    //字段名称样式
            //    ICellStyle cellStyleField = workbook.CreateCellStyle();
            //    cellStyleField.Alignment = HorizontalAlignment.Center;
            //    cellStyleField.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleField.BorderBottom = BorderStyle.Thin;
            //    cellStyleField.BorderLeft = BorderStyle.Thin;
            //    cellStyleField.BorderRight = BorderStyle.Thin;
            //    cellStyleField.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleField.SetFont(fontField);
            //    //前景色
            //    cellStyleField.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            //    cellStyleField.FillPattern = FillPattern.SolidForeground;
            //    cellStyleField.WrapText = true;
            //    //cellStyleField.IsLocked = false;

            //    //字段值的样式  单数
            //    ICellStyle cellStyleOddRow = workbook.CreateCellStyle();
            //    cellStyleOddRow.Alignment = HorizontalAlignment.Center;
            //    cellStyleOddRow.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleOddRow.BorderBottom = BorderStyle.Thin;
            //    cellStyleOddRow.BorderLeft = BorderStyle.Thin;
            //    cellStyleOddRow.BorderRight = BorderStyle.Thin;
            //    cellStyleOddRow.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleOddRow.SetFont(fontField);

            //    //字段值的样式  单数   int 类型
            //    ICellStyle cellStyleOddRowInt = workbook.CreateCellStyle();
            //    cellStyleOddRowInt.Alignment = HorizontalAlignment.Center;
            //    cellStyleOddRowInt.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleOddRowInt.BorderBottom = BorderStyle.Thin;
            //    cellStyleOddRowInt.BorderLeft = BorderStyle.Thin;
            //    cellStyleOddRowInt.BorderRight = BorderStyle.Thin;
            //    cellStyleOddRowInt.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleOddRowInt.SetFont(fontField);
            //    cellStyleOddRowInt.DataFormat = dataFormat.GetFormat("0");

            //    //字段值的样式  单数   double 类型
            //    ICellStyle cellStyleOddRowDouble = workbook.CreateCellStyle();
            //    cellStyleOddRowDouble.Alignment = HorizontalAlignment.Center;
            //    cellStyleOddRowDouble.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleOddRowDouble.BorderBottom = BorderStyle.Thin;
            //    cellStyleOddRowDouble.BorderLeft = BorderStyle.Thin;
            //    cellStyleOddRowDouble.BorderRight = BorderStyle.Thin;
            //    cellStyleOddRowDouble.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleOddRowDouble.SetFont(fontField);
            //    cellStyleOddRowDouble.DataFormat = dataFormat.GetFormat("0.0");


            //    //字段值的样式  双数
            //    ICellStyle cellStyleEvenRow = workbook.CreateCellStyle();
            //    cellStyleEvenRow.Alignment = HorizontalAlignment.Center;
            //    cellStyleEvenRow.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleEvenRow.BorderBottom = BorderStyle.Thin;
            //    cellStyleEvenRow.BorderLeft = BorderStyle.Thin;
            //    cellStyleEvenRow.BorderRight = BorderStyle.Thin;
            //    cellStyleEvenRow.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleEvenRow.SetFont(fontField);
            //    //前景色
            //    cellStyleEvenRow.FillForegroundColor = HSSFColor.LightGreen.Index;
            //    cellStyleEvenRow.FillPattern = FillPattern.SolidForeground;


            //    //字段值的样式  双数   int类型
            //    ICellStyle cellStyleEvenRowInt = workbook.CreateCellStyle();
            //    cellStyleEvenRowInt.Alignment = HorizontalAlignment.Center;
            //    cellStyleEvenRowInt.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleEvenRowInt.BorderBottom = BorderStyle.Thin;
            //    cellStyleEvenRowInt.BorderLeft = BorderStyle.Thin;
            //    cellStyleEvenRowInt.BorderRight = BorderStyle.Thin;
            //    cellStyleEvenRowInt.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleEvenRowInt.SetFont(fontField);
            //    //前景色
            //    cellStyleEvenRowInt.FillForegroundColor = HSSFColor.LightGreen.Index;
            //    cellStyleEvenRowInt.FillPattern = FillPattern.SolidForeground;
            //    cellStyleEvenRowInt.DataFormat = dataFormat.GetFormat("0");


            //    //字段值的样式  双数  double类型
            //    ICellStyle cellStyleEvenRowDouble = workbook.CreateCellStyle();
            //    cellStyleEvenRowDouble.Alignment = HorizontalAlignment.Center;
            //    cellStyleEvenRowDouble.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleEvenRowDouble.BorderBottom = BorderStyle.Thin;
            //    cellStyleEvenRowDouble.BorderLeft = BorderStyle.Thin;
            //    cellStyleEvenRowDouble.BorderRight = BorderStyle.Thin;
            //    cellStyleEvenRowDouble.BorderTop = BorderStyle.Thin;
            //    //字体
            //    cellStyleEvenRowDouble.SetFont(fontField);
            //    //前景色
            //    cellStyleEvenRowDouble.FillForegroundColor = HSSFColor.LightGreen.Index;
            //    cellStyleEvenRowDouble.FillPattern = FillPattern.SolidForeground;
            //    cellStyleEvenRowDouble.DataFormat = dataFormat.GetFormat("0.0");


            //    //操作员
            //    ICellStyle cellStyleOperator = workbook.CreateCellStyle();
            //    cellStyleOperator.Alignment = HorizontalAlignment.Left;
            //    cellStyleOperator.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleOperator.BorderBottom = BorderStyle.None;
            //    cellStyleOperator.BorderLeft = BorderStyle.None;
            //    cellStyleOperator.BorderRight = BorderStyle.None;
            //    cellStyleOperator.BorderTop = BorderStyle.None;
            //    //字体
            //    cellStyleOperator.SetFont(fontField);

            //    //审核人
            //    ICellStyle cellStyleAudit = workbook.CreateCellStyle();
            //    cellStyleAudit.Alignment = HorizontalAlignment.Left;
            //    cellStyleAudit.VerticalAlignment = VerticalAlignment.Center;
            //    //边框
            //    cellStyleAudit.BorderBottom = BorderStyle.None;
            //    cellStyleAudit.BorderLeft = BorderStyle.None;
            //    cellStyleAudit.BorderRight = BorderStyle.None;
            //    cellStyleAudit.BorderTop = BorderStyle.None;
            //    //字体
            //    cellStyleAudit.SetFont(fontField);
            //    #endregion

            //    //数值类型的字段集合
            //    List<string> list = new List<string>()
            //    {
            //        "事假（H）","病假（H）","工作日缓发（H）","工作日加班缓发（H）","休息日和节假日加班缓发（H）",
            //        "工作日加班（H）","休息日加班（H）","节假日加班（H）","年假（天）","调休假总计（H）","迟到（次）",
            //        "当月累计调休（H）"
            //    };


            //    //给单元格赋值
            //    //表的标题
            //    IRow rowTitle = sheet.CreateRow(0);
            //    ICell cellTitle = rowTitle.CreateCell(0);
            //    string fileName = GetModuleName(moduleId, departName, currentAttdanceDate);
            //    cellTitle.SetCellValue(fileName);
            //    cellTitle.CellStyle = cellStyleTitle;

            //    //字段名称
            //    IRow rowField = sheet.CreateRow(1);
            //    for (int i = 0; i < dt.Columns.Count; i++)
            //    {
            //        ICell cellField = rowField.CreateCell(i);
            //        cellField.SetCellValue(dt.Columns[i].ColumnName);
            //        cellField.CellStyle = cellStyleField;
            //    }

            //    if (dt != null && dt.Rows.Count != 0)
            //    {
            //        //字段值
            //        for (int i = 2; i < dt.Rows.Count + 2; i++)
            //        {
            //            //先创建行
            //            IRow rowValue = sheet.CreateRow(i);
            //            for (int k = 0; k < dt.Columns.Count; k++)
            //            {
            //                ICell cellValue = rowValue.CreateCell(k);
            //                if (i % 2 == 0)
            //                {
            //                    if (list.Contains(dt.Columns[k].ColumnName))
            //                    {
            //                        cellValue.SetCellType(CellType.Numeric);
            //                        cellValue.SetCellValue(dt.Rows[i - 2][k].ToDouble() == 0 ? string.Empty : dt.Rows[i - 2][k].ToDouble().ToString());
            //                        cellValue.CellStyle = cellStyleOddRowDouble;
            //                    }
            //                    else
            //                    {
            //                        cellValue.SetCellValue(dt.Rows[i - 2][k].ToString());
            //                        cellValue.CellStyle = cellStyleOddRow;
            //                    }

            //                }
            //                else
            //                {
            //                    if (list.Contains(dt.Columns[k].ColumnName))
            //                    {
            //                        cellValue.SetCellType(CellType.Numeric);
            //                        cellValue.SetCellValue(dt.Rows[i - 2][k].ToDouble() == 0 ? string.Empty : dt.Rows[i - 2][k].ToDouble().ToString());
            //                        cellValue.CellStyle = cellStyleEvenRowDouble;
            //                    }
            //                    else
            //                    {
            //                        cellValue.SetCellValue(dt.Rows[i - 2][k].ToString());
            //                        cellValue.CellStyle = cellStyleEvenRow;
            //                    }
            //                }
            //                sheet.SetColumnWidth(k, (Encoding.UTF8.GetBytes(dt.Rows[i - 2][k].ToString()).Length) * 256);

            //                ////另一种设置方法
            //                ////先拿到当前cell的宽度
            //                //int columnWidth = sheet.GetColumnWidth(k);
            //                ////拿到当前单元格中字符串的宽度
            //                //int length = Encoding.UTF8.GetBytes(dt.Rows[i - 2][k].ToString()).Length;
            //                //if (columnWidth < length + 1)
            //                //{
            //                //    columnWidth = length + 1;
            //                //}
            //                //sheet.SetColumnWidth(k, columnWidth * 256);
            //            }
            //        }

            //        //这里最后要加上 
            //        //操作员     导出日期        一级审核    审核日期
            //        //二级审核  审核日期        三级审核    审核日期
            //        //先创建行
            //        IRow rowOperatorInfo = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            //        ICell cellOperatorInfo = rowOperatorInfo.CreateCell(0);
            //        cellOperatorInfo.SetCellValue(string.Format("操作员：{0}    导出日期：{1}", PageHandle.profile.UserName, DateTime.Now.ToString("yyyy-MM-dd HH:mm")));
            //        cellOperatorInfo.CellStyle = cellStyleOperator;

            //        //审批流信息
            //        IRow rowAuditInfo = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            //        ICell cellAuditInfo = rowAuditInfo.CreateCell(0);
            //        Dictionary<string, string> dic = QueryAuditInfo(dan);
            //        if (dic.Count == 0)
            //        {
            //            cellAuditInfo.SetCellValue("审核人信息：无");
            //        }
            //        else
            //        {
            //            StringBuilder sb = new StringBuilder();
            //            foreach (KeyValuePair<string, string> item in dic)
            //            {
            //                sb.AppendFormat("{0}   审核日期：{1}   ", item.Key, item.Value);
            //            }
            //            cellAuditInfo.SetCellValue(sb.ToString());
            //        }
            //        cellAuditInfo.CellStyle = cellStyleAudit;
            //    }
            //    else
            //    {
            //        //没有查询到数据
            //        //应该显示暂无数据
            //        IRow rowContent = sheet.CreateRow(2);
            //        ICell cellContent = rowContent.CreateCell(0);
            //        cellContent.SetCellValue("暂无内容");
            //        cellContent.CellStyle = cellStyleOddRow;
            //        sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, dt.Columns.Count - 1));

            //    }
            //    //合并第一行  
            //    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dt.Columns.Count - 1));
            //    //合并倒数第二行
            //    sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 2, sheet.PhysicalNumberOfRows - 2, 0, dt.Columns.Count - 1));
            //    sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 1, sheet.PhysicalNumberOfRows - 1, 0, dt.Columns.Count - 1));
            //    //合并倒数第一行
            //    //给第一列增加筛选
            //    CellRangeAddress a2 = CellRangeAddress.ValueOf("A2");
            //    sheet.SetAutoFilter(a2);

            //    //自动调整列宽
            //    //ISheet sheetObj = workbook.GetSheetAt(0);
            //    //sheetObj.AutoSizeColumn(1);
            //    //sheetObj.AutoSizeColumn(6);

            //    //给Sheet表中锁定的区域设置密码
            //    //sheet.ProtectSheet("123");

            //    //先要在指定位置创建一个文件夹
            //    //这里拿到当前应用程序域的跟目录
            //    string rootPath = System.AppDomain.CurrentDomain.BaseDirectory;
            //    string newDirName = "TenantInfo\\ExportDetailsDataToExcel\\部门考勤汇总明细表\\";
            //    string newPath = Path.Combine(rootPath, newDirName);
            //    Directory.CreateDirectory(newPath);
            //    using (FileStream fsWrite = new FileStream(newPath + fileName + ".xls", FileMode.OpenOrCreate, FileAccess.Write))
            //    {
            //        //全部创建完成之后，通过写入流的方式将文件上传，或者写入磁盘
            //        workbook.Write(fsWrite);
            //        fsWrite.Flush();
            //        fsWrite.Close();
            //    }
            //    newDirName = newDirName.Replace("\\", "/");
            //    //这里注意的是只需要返回相对路径
            //    return string.Format("../../{0}{1}.xls", newDirName, fileName);
            //}


            ///// <summary>
            ///// 根据dan号查询审核信息
            ///// </summary>
            ///// <param name="dan"></param>
            ///// <returns></returns>
            //public Dictionary<string, string> QueryAuditInfo(string dan)
            //{
            //    Dictionary<string, string> dic = new Dictionary<string, string>();
            //    string sql = string.Format("  select  a.AuditBy,a.AuditDate,a.PhaseId  from   Base_Audit_PhaseAuditInfo a " +
            //                                          "  inner join  CUS_Table_5183  b   on  a.RecordId = b.Id" +
            //                                          "  where   a.RecordId = (select  a.Id  from CUS_Table_5183 a  where a.Dan = '{0}')" +
            //                                          "  and a.PhaseId != ''  order by  a.AuditDate; ", dan);
            //    DataSet dataSet = SqlHelper.ExecuteDataset(1, sql);
            //    if (dataSet != null && !dataSet.IsEmpty())
            //    {
            //        for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            //        {

            //            string phaseName = QueryDisplayNameByPhaseId(dataSet.Tables[0].Rows[i][2].ToString());
            //            string userName = QueryAliasByUserId(dataSet.Tables[0].Rows[i][0].ToString());
            //            string date = dataSet.Tables[0].Rows[i][1].ToDateTime().ToString("yyyy-MM-dd HH:mm");
            //            dic.Add(string.Format("{0}：{1}", phaseName, userName), date);
            //        }
            //    }
            //    return dic;
            //}


            ///// <summary>
            ///// 根据PhaseId来查询审核流的名称
            ///// </summary>
            ///// <param name="id"></param>
            ///// <returns></returns>
            //public string QueryDisplayNameByPhaseId(string id)
            //{
            //    string sql = string.Format(" select  a.DisplayName  from   Base_Audit_PhaseDefine  a   where  a.Id='{0}'   ", id);
            //    return SqlHelper.ExecuteScalar(1, sql).ToString();
            //}


            ///// <summary>
            ///// 根据UserId查询姓名
            ///// </summary>
            ///// <param name="userId"></param>
            ///// <returns></returns>
            //public string QueryAliasByUserId(string userId)
            //{
            //    string sql = string.Format("  select  a.Alias  from   Base_Users  a  where  a.Id='{0}'  ", userId);
            //    return SqlHelper.ExecuteScalar(1, sql).ToString();
            //}


            ///// <summary>
            ///// 模块id
            ///// </summary>
            ///// <param name="moduleId"></param>
            ///// <param name="departAlias"></param>
            ///// <param name="currentAttdanceDate">考勤月份</param>
            ///// <returns></returns>
            //public string GetModuleName(int moduleId, string departAlias, string currentAttdanceDate)
            //{
            //    string sql = string.Format("select  a.DisplayName  from   Base_MultiLang_Table_CN  a  where  a.ModuleId={0} ", moduleId);
            //    string name = SqlHelper.ExecuteScalar(1, sql).ToString();
            //    int index = name.LastIndexOf("@");
            //    return "CRM" + name.Substring(index + 1) + "-" + departAlias + "-" + currentAttdanceDate;
            //}


            ///// <summary>
            ///// 根据个人汇总表中的dan号来查询当前员工所在部门
            ///// </summary>
            ///// <param name="dan"></param>
            ///// <returns></returns>
            //public string GetDepartNameByDan(string dan)
            //{
            //    string sql = string.Format(" select   a.DepartAlias  from   CUS_Table_5183  a  where  a.Dan='{0}'   ", dan.Substring(1, dan.Length - 2));
            //    return SqlHelper.ExecuteScalar(1, sql).ToString();

            //}

            #endregion

            #region 使用定义号的模板导出数据

            //    /// <summary>
            //    /// 将数据返回给前端
            //    /// </summary>
            //public override void Begin()
            //{
            //    PageHandle.Write(ExportExcel());
            //}


            ///// <summary>
            ///// 将生成excel的路径返回
            ///// </summary>
            ///// <returns></returns>
            //public string ExportExcel()
            //    {
            //        string arrListStr = PageHandle.Request["arr"].ToString();
            //        arrListStr = arrListStr.Substring(1, arrListStr.Length - 2).Replace('\"', '\'');

            //        string firstDan = string.Empty;
            //        int index = arrListStr.IndexOf(",");
            //        if (index == -1)
            //        {
            //            firstDan = arrListStr;
            //        }
            //        else
            //        {
            //            firstDan = arrListStr.Substring(0, index);
            //        }
            //        if (!IsAuditFinish(firstDan.Substring(1, firstDan.Length - 2)))
            //        {
            //            //如果当前数据未审核完成
            //            return "0";
            //        }
            //        string departAlias = GetDepartNameByDan(firstDan);
            //        DataTable dt = GetDataTable(arrListStr);
            //        return GetWorkbook(dt, departAlias, 5184, firstDan.Substring(1, 6), firstDan.Substring(1, firstDan.Length - 2));
            //    }

            //    /// <summary>
            //    /// 获取需要导出的全部数据   
            //    /// </summary>
            //    /// <returns></returns>
            //    public DataTable GetDataTable(string arrListStr)
            //    {
            //        //接收参数     
            //        //De_29059Alias,De_29060,De_29061,De_31824,De_29062,De_31825,De_29063,
            //        //De_31827,De_29064,De_31828,De_31838,De_31839,De_29065,De_29066,
            //        //De_31829,De_29067,De_31830,De_29068,De_31831,De_29069,De_31826,De_31848,
            //        //De_31849,De_29071,Memo
            //        if (!string.IsNullOrEmpty(arrListStr))
            //        {
            //            string sqlQueryUserDetails = string.Format(" select " +
            //                " De_29060 as  '工号'  ,De_29059Alias  as  '姓名',De_29061  as '事假（H）' ,De_31824  as  '事假明细', " +
            //                " De_29062  as  '病假（H）'  ,De_31825   as   '病假明细'  ,De_29063   as   '工作日缓发（H）', " +
            //                " De_31827  as  '工作日缓发明细' ,De_29064  as  '工作日加班缓发（H）'  ,De_31828  as  '工作日加班缓发明细'  ," +
            //                " De_31838  as  '休息日和节假日加班缓发（H）' ,De_31839  as  '休息日和节假日加班缓发明细' ,De_29065   as   '迟到（次）'  ," +
            //                " De_32218  as  '迟到明细',De_29067 as  '休息日加班（H）',De_31830  as  '休息日加班明细', " +
            //                " De_29066  as  '工作日加班（H）',De_31829 as  '工作日加班明细' ," +
            //                " De_29068  as  '节假日加班（H）',De_31831 as  '节假日加班明细' ," +
            //                " De_29069  as  '年假（天）',De_31826  as  '年假明细',De_31848   as  '调休假总计（H）'  ," +
            //                " De_31849  as  '调休假明细',De_32219 as '当月累计调休（H）',De_32220  as   '当月累计调休明细',  " +
            //                " a.Memo  as  '备注' " +
            //                " from   CUS_Table_5183_Details  a   " +
            //                " inner  join  CUS_Table_5183  b   on  a.InnerId=b.Id " +
            //                " where  b.Dan  in  ({0})    order by    '工号'  asc   ; ", arrListStr);
            //            return SqlHelper.ExecuteDataset(1, sqlQueryUserDetails).Tables[0];
            //        }
            //        return new DataTable();
            //    }


            //    /// <summary>
            //    /// 设置工作簿的格式
            //    /// </summary>
            //    /// <returns></returns>
            //    /// <param name="currentAttdanceDate"></param>
            //    /// <param name="departName"></param>
            //    /// <param name="dt"></param>
            //    /// <param name="moduleId"></param>
            //    /// <param name="dan">主表的dan号</param>
            //    public string GetWorkbook(DataTable dt, string departName, int moduleId, string currentAttdanceDate, string dan)
            //    {
            //        //这里应该是直接通过npoi读取模板文件，格式已经调整好  
            //        //定义一个模板路径
            //        string templateFileName = "模板.xls";
            //        string templateFileExtensionName = Path.GetExtension(templateFileName);
            //        string templateFilePath = "TenantInfo\\ExportDetailsDataTemplate\\";
            //        string rootPath = System.AppDomain.CurrentDomain.BaseDirectory;
            //        string templateFileObsoletePath = Path.Combine(rootPath, templateFilePath);
            //        if (!Directory.Exists(templateFileObsoletePath))
            //        {
            //            //在指定路径下创建目录
            //            Directory.CreateDirectory(templateFileObsoletePath);
            //        }
            //        IWorkbook workbook;
            //        using (FileStream filestream = new FileStream(templateFileObsoletePath + templateFileName, FileMode.Open, FileAccess.Read))
            //        {
            //            if (templateFileExtensionName == ".xls")
            //            {
            //                workbook = new HSSFWorkbook(filestream);
            //            }
            //            else if (templateFileExtensionName == ".xlsx")
            //            {
            //                workbook = new XSSFWorkbook(filestream);
            //            }
            //            else
            //            {
            //                //其他类型的文件直接返回0
            //                return "0";
            //            }
            //        }
            //        ISheet sheet = workbook.GetSheetAt(0);

            //        #region 创建样式   Font  Style
            //        IDataFormat dataFormat = workbook.CreateDataFormat();
            //        IFont fontTitle = workbook.CreateFont();
            //        fontTitle.FontName = "等线";
            //        fontTitle.FontHeight = 15;
            //        fontTitle.FontHeightInPoints = 15;

            //        IFont fontField = workbook.CreateFont();
            //        fontField.FontName = "等线";
            //        fontField.FontHeight = 10;
            //        fontField.FontHeightInPoints = 10;


            //        //创建style   四种   表名   字段名   字段间隔显示
            //        //表名的样式
            //        ICellStyle cellStyleTitle = workbook.CreateCellStyle();
            //        //居中
            //        cellStyleTitle.Alignment = HorizontalAlignment.Center;
            //        cellStyleTitle.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleTitle.BorderBottom = BorderStyle.Thin;
            //        cellStyleTitle.BorderLeft = BorderStyle.Thin;
            //        cellStyleTitle.BorderRight = BorderStyle.Thin;
            //        cellStyleTitle.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleTitle.SetFont(fontTitle);

            //        //字段名称样式
            //        ICellStyle cellStyleField = workbook.CreateCellStyle();
            //        cellStyleField.Alignment = HorizontalAlignment.Center;
            //        cellStyleField.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleField.BorderBottom = BorderStyle.Thin;
            //        cellStyleField.BorderLeft = BorderStyle.Thin;
            //        cellStyleField.BorderRight = BorderStyle.Thin;
            //        cellStyleField.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleField.SetFont(fontField);
            //        //前景色
            //        cellStyleField.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            //        cellStyleField.FillPattern = FillPattern.SolidForeground;
            //        cellStyleField.WrapText = true;
            //        //cellStyleField.IsLocked = false;

            //        //字段值的样式  单数
            //        ICellStyle cellStyleOddRow = workbook.CreateCellStyle();
            //        cellStyleOddRow.Alignment = HorizontalAlignment.Center;
            //        cellStyleOddRow.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleOddRow.BorderBottom = BorderStyle.Thin;
            //        cellStyleOddRow.BorderLeft = BorderStyle.Thin;
            //        cellStyleOddRow.BorderRight = BorderStyle.Thin;
            //        cellStyleOddRow.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleOddRow.SetFont(fontField);
            //        //设置自动换行
            //        //这里是否需要设置换行   不确定
            //        //cellStyleOddRow.WrapText = true;

            //        //字段值的样式  单数   int 类型
            //        ICellStyle cellStyleOddRowInt = workbook.CreateCellStyle();
            //        cellStyleOddRowInt.Alignment = HorizontalAlignment.Center;
            //        cellStyleOddRowInt.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleOddRowInt.BorderBottom = BorderStyle.Thin;
            //        cellStyleOddRowInt.BorderLeft = BorderStyle.Thin;
            //        cellStyleOddRowInt.BorderRight = BorderStyle.Thin;
            //        cellStyleOddRowInt.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleOddRowInt.SetFont(fontField);
            //        cellStyleOddRowInt.DataFormat = dataFormat.GetFormat("0");

            //        //字段值的样式  单数   double 类型
            //        ICellStyle cellStyleOddRowDouble = workbook.CreateCellStyle();
            //        cellStyleOddRowDouble.Alignment = HorizontalAlignment.Center;
            //        cellStyleOddRowDouble.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleOddRowDouble.BorderBottom = BorderStyle.Thin;
            //        cellStyleOddRowDouble.BorderLeft = BorderStyle.Thin;
            //        cellStyleOddRowDouble.BorderRight = BorderStyle.Thin;
            //        cellStyleOddRowDouble.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleOddRowDouble.SetFont(fontField);
            //        cellStyleOddRowDouble.DataFormat = dataFormat.GetFormat("0.0");


            //        //字段值的样式  双数
            //        ICellStyle cellStyleEvenRow = workbook.CreateCellStyle();
            //        cellStyleEvenRow.Alignment = HorizontalAlignment.Center;
            //        cellStyleEvenRow.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleEvenRow.BorderBottom = BorderStyle.Thin;
            //        cellStyleEvenRow.BorderLeft = BorderStyle.Thin;
            //        cellStyleEvenRow.BorderRight = BorderStyle.Thin;
            //        cellStyleEvenRow.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleEvenRow.SetFont(fontField);
            //        //前景色
            //        cellStyleEvenRow.FillForegroundColor = HSSFColor.LightGreen.Index;
            //        cellStyleEvenRow.FillPattern = FillPattern.SolidForeground;


            //        //字段值的样式  双数   int类型
            //        ICellStyle cellStyleEvenRowInt = workbook.CreateCellStyle();
            //        cellStyleEvenRowInt.Alignment = HorizontalAlignment.Center;
            //        cellStyleEvenRowInt.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleEvenRowInt.BorderBottom = BorderStyle.Thin;
            //        cellStyleEvenRowInt.BorderLeft = BorderStyle.Thin;
            //        cellStyleEvenRowInt.BorderRight = BorderStyle.Thin;
            //        cellStyleEvenRowInt.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleEvenRowInt.SetFont(fontField);
            //        //前景色
            //        cellStyleEvenRowInt.FillForegroundColor = HSSFColor.LightGreen.Index;
            //        cellStyleEvenRowInt.FillPattern = FillPattern.SolidForeground;
            //        cellStyleEvenRowInt.DataFormat = dataFormat.GetFormat("0");


            //        //字段值的样式  双数  double类型
            //        ICellStyle cellStyleEvenRowDouble = workbook.CreateCellStyle();
            //        cellStyleEvenRowDouble.Alignment = HorizontalAlignment.Center;
            //        cellStyleEvenRowDouble.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleEvenRowDouble.BorderBottom = BorderStyle.Thin;
            //        cellStyleEvenRowDouble.BorderLeft = BorderStyle.Thin;
            //        cellStyleEvenRowDouble.BorderRight = BorderStyle.Thin;
            //        cellStyleEvenRowDouble.BorderTop = BorderStyle.Thin;
            //        //字体
            //        cellStyleEvenRowDouble.SetFont(fontField);
            //        //前景色
            //        cellStyleEvenRowDouble.FillForegroundColor = HSSFColor.LightGreen.Index;
            //        cellStyleEvenRowDouble.FillPattern = FillPattern.SolidForeground;
            //        cellStyleEvenRowDouble.DataFormat = dataFormat.GetFormat("0.0");


            //        //操作员
            //        ICellStyle cellStyleOperator = workbook.CreateCellStyle();
            //        cellStyleOperator.Alignment = HorizontalAlignment.Left;
            //        cellStyleOperator.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleOperator.BorderBottom = BorderStyle.None;
            //        cellStyleOperator.BorderLeft = BorderStyle.None;
            //        cellStyleOperator.BorderRight = BorderStyle.None;
            //        cellStyleOperator.BorderTop = BorderStyle.None;
            //        //字体
            //        cellStyleOperator.SetFont(fontField);

            //        //审核人
            //        ICellStyle cellStyleAudit = workbook.CreateCellStyle();
            //        cellStyleAudit.Alignment = HorizontalAlignment.Left;
            //        cellStyleAudit.VerticalAlignment = VerticalAlignment.Center;
            //        //边框
            //        cellStyleAudit.BorderBottom = BorderStyle.None;
            //        cellStyleAudit.BorderLeft = BorderStyle.None;
            //        cellStyleAudit.BorderRight = BorderStyle.None;
            //        cellStyleAudit.BorderTop = BorderStyle.None;
            //        //字体
            //        cellStyleAudit.SetFont(fontField);
            //        #endregion

            //        //数值类型的字段集合
            //        List<string> list = new List<string>()
            //    {
            //        "事假（H）","病假（H）","工作日缓发（H）","工作日加班缓发（H）","休息日和节假日加班缓发（H）",
            //        "工作日加班（H）","休息日加班（H）","节假日加班（H）","年假（天）","调休假总计（H）","迟到（次）",
            //        "当月累计调休（H）"
            //    };


            //        //给单元格赋值
            //        //表的标题
            //        IRow rowTitle = sheet.GetRow(0);
            //        ICell cellTitle = rowTitle.GetCell(0);
            //        string fileName = GetModuleName(moduleId, departName, currentAttdanceDate);
            //        cellTitle.SetCellValue(fileName);
            //        //cellTitle.CellStyle = cellStyleTitle;

            //        //字段名称
            //        IRow rowField = sheet.GetRow(1);
            //        for (int i = 0; i < dt.Columns.Count; i++)
            //        {
            //            ICell cellField = rowField.GetCell(i);
            //            cellField.SetCellValue(dt.Columns[i].ColumnName);
            //            //cellField.CellStyle = cellStyleField;
            //        }

            //        if (dt != null && dt.Rows.Count != 0)
            //        {
            //            //字段值
            //            for (int i = 2; i < dt.Rows.Count + 2; i++)
            //            {
            //                //先创建行
            //                IRow rowValue = sheet.CreateRow(i);
            //                for (int k = 0; k < dt.Columns.Count; k++)
            //                {
            //                    ICell cellValue = rowValue.CreateCell(k);
            //                    if (i % 2 == 0)
            //                    {
            //                        if (list.Contains(dt.Columns[k].ColumnName))
            //                        {
            //                            cellValue.SetCellType(CellType.Numeric);
            //                            cellValue.SetCellValue(dt.Rows[i - 2][k].ToDouble() == 0 ? string.Empty : dt.Rows[i - 2][k].ToDouble().ToString());
            //                            cellValue.CellStyle = cellStyleOddRowDouble;
            //                        }
            //                        else
            //                        {
            //                            cellValue.SetCellValue(dt.Rows[i - 2][k].ToString());
            //                            cellValue.CellStyle = cellStyleOddRow;
            //                        }

            //                    }
            //                    else
            //                    {
            //                        if (list.Contains(dt.Columns[k].ColumnName))
            //                        {
            //                            cellValue.SetCellType(CellType.Numeric);
            //                            cellValue.SetCellValue(dt.Rows[i - 2][k].ToDouble() == 0 ? string.Empty : dt.Rows[i - 2][k].ToDouble().ToString());
            //                            cellValue.CellStyle = cellStyleEvenRowDouble;
            //                        }
            //                        else
            //                        {
            //                            cellValue.SetCellValue(dt.Rows[i - 2][k].ToString());
            //                            cellValue.CellStyle = cellStyleEvenRow;
            //                        }
            //                    }
            //                    //删掉这句    导出的excel文件打开是乱码？  把xlsx改成xls就可以了
            //                    //sheet.SetColumnWidth(k, (Encoding.UTF8.GetBytes(dt.Rows[i - 2][k].ToString()).Length) * 256);
            //                    //sheet.SetColumnWidth(rowField.Cells[i].CellS)

            //                    ////另一种设置字段宽度的方法
            //                    ////先拿到当前cell的宽度
            //                    //int columnWidth = sheet.GetColumnWidth(k);
            //                    ////拿到当前单元格中字符串的宽度
            //                    //int length = Encoding.UTF8.GetBytes(dt.Rows[i - 2][k].ToString()).Length;
            //                    //if (columnWidth < length + 1)
            //                    //{
            //                    //    columnWidth = length + 1;
            //                    //}
            //                    //sheet.SetColumnWidth(k, columnWidth * 256);
            //                }
            //            }

            //            //这里最后要加上 
            //            //操作员     导出日期        一级审核    审核日期
            //            //二级审核  审核日期        三级审核    审核日期
            //            //先创建行
            //            IRow rowOperatorInfo = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            //            ICell cellOperatorInfo = rowOperatorInfo.CreateCell(0);
            //            cellOperatorInfo.SetCellValue(string.Format("操作员：{0}    导出日期：{1}", PageHandle.profile.UserName, DateTime.Now.ToString("yyyy-MM-dd HH:mm")));
            //            cellOperatorInfo.CellStyle = cellStyleOperator;

            //            //审批流信息
            //            IRow rowAuditInfo = sheet.CreateRow(sheet.PhysicalNumberOfRows);
            //            ICell cellAuditInfo = rowAuditInfo.CreateCell(0);
            //            Dictionary<string, string> dic = QueryAuditInfo(dan);
            //            if (dic.Count == 0)
            //            {
            //                cellAuditInfo.SetCellValue("审核人信息：无");
            //            }
            //            else
            //            {
            //                StringBuilder sb = new StringBuilder();
            //                foreach (KeyValuePair<string, string> item in dic)
            //                {
            //                    sb.AppendFormat("{0}   审核日期：{1}   ", item.Key, item.Value);
            //                }
            //                cellAuditInfo.SetCellValue(sb.ToString());
            //            }
            //            cellAuditInfo.CellStyle = cellStyleAudit;
            //        }
            //        else
            //        {
            //            //没有查询到数据
            //            //应该显示暂无数据
            //            IRow rowContent = sheet.CreateRow(2);
            //            ICell cellContent = rowContent.CreateCell(0);
            //            cellContent.SetCellValue("暂无内容");
            //            cellContent.CellStyle = cellStyleOddRow;
            //            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, dt.Columns.Count - 1));

            //        }
            //        //合并第一行  
            //        //sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dt.Columns.Count - 1));
            //        //合并倒数第二行
            //        sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 2, sheet.PhysicalNumberOfRows - 2, 0, dt.Columns.Count - 1));
            //        //合并倒数第一行
            //        sheet.AddMergedRegion(new CellRangeAddress(sheet.PhysicalNumberOfRows - 1, sheet.PhysicalNumberOfRows - 1, 0, dt.Columns.Count - 1));
            //        //给第一列增加筛选
            //        CellRangeAddress a2 = CellRangeAddress.ValueOf("A2");
            //        sheet.SetAutoFilter(a2);

            //        //自动调整列宽
            //        ISheet sheetObj = workbook.GetSheetAt(0);
            //        sheetObj.AutoSizeColumn(1);
            //        //sheetObj.AutoSizeColumn(6);

            //        //给Sheet表中锁定的区域设置密码
            //        //sheet.ProtectSheet("123");

            //        //先要在指定位置创建一个文件夹
            //        //这里拿到当前应用程序域的跟目录
            //        string newDirName = "TenantInfo\\ExportDetailsDataToExcel\\部门考勤汇总明细表\\";
            //        string newPath = Path.Combine(rootPath, newDirName);
            //        Directory.CreateDirectory(newPath);
            //        using (FileStream fsWrite = new FileStream(newPath + fileName + ".xls", FileMode.OpenOrCreate, FileAccess.Write))
            //        {
            //            //全部创建完成之后，通过写入流的方式将文件上传，或者写入磁盘
            //            workbook.Write(fsWrite);
            //            fsWrite.Flush();
            //            fsWrite.Close();
            //        }
            //        newDirName = newDirName.Replace("\\", "/");
            //        //这里注意的是只需要返回相对路径
            //        return string.Format("../../{0}{1}.xls", newDirName, fileName);
            //    }


            //    /// <summary>
            //    /// 查询该条数据是否审核完成
            //    /// </summary>
            //    /// <param name="dan"></param>
            //    /// <returns></returns>
            //    public bool IsAuditFinish(string dan)
            //    {
            //        string sql = string.Format(" select  count(*)  from  CUS_Table_5183  a  where  a.Status=3 and  a.Dan='{0}'  ", dan);
            //        int n = SqlHelper.ExecuteScalar(1, sql).ToInt32();
            //        if (n > 0)
            //        {
            //            return true;
            //        }
            //        return false;
            //    }


            //    /// <summary>
            //    /// 根据dan号查询审核信息
            //    /// </summary>
            //    /// <param name="dan"></param>
            //    /// <returns></returns>
            //    public Dictionary<string, string> QueryAuditInfo(string dan)
            //    {
            //        Dictionary<string, string> dic = new Dictionary<string, string>();
            //        string sql = string.Format("  select  a.AuditBy,a.AuditDate,a.PhaseId  from   Base_Audit_PhaseAuditInfo a " +
            //                                              "  inner join  CUS_Table_5183  b   on  a.RecordId = b.Id" +
            //                                              "  where   a.RecordId = (select  a.Id  from CUS_Table_5183 a  where a.Dan = '{0}')" +
            //                                              "  and a.PhaseId != ''  order by  a.AuditDate; ", dan);
            //        DataSet dataSet = SqlHelper.ExecuteDataset(1, sql);
            //        if (dataSet != null && !dataSet.IsEmpty())
            //        {
            //            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            //            {

            //                string phaseName = QueryDisplayNameByPhaseId(dataSet.Tables[0].Rows[i][2].ToString());
            //                string userName = QueryAliasByUserId(dataSet.Tables[0].Rows[i][0].ToString());
            //                string date = dataSet.Tables[0].Rows[i][1].ToDateTime().ToString("yyyy-MM-dd HH:mm");
            //                dic.Add(string.Format("{0}：{1}", phaseName, userName), date);
            //            }
            //        }
            //        return dic;
            //    }


            //    /// <summary>
            //    /// 根据PhaseId来查询审核流的名称
            //    /// </summary>
            //    /// <param name="id"></param>
            //    /// <returns></returns>
            //    public string QueryDisplayNameByPhaseId(string id)
            //    {
            //        string sql = string.Format(" select  a.DisplayName  from   Base_Audit_PhaseDefine  a   where  a.Id='{0}'   ", id);
            //        return SqlHelper.ExecuteScalar(1, sql).ToString();
            //    }


            //    /// <summary>
            //    /// 根据UserId查询姓名
            //    /// </summary>
            //    /// <param name="userId"></param>
            //    /// <returns></returns>
            //    public string QueryAliasByUserId(string userId)
            //    {
            //        string sql = string.Format("  select  a.Alias  from   Base_Users  a  where  a.Id='{0}'  ", userId);
            //        return SqlHelper.ExecuteScalar(1, sql).ToString();
            //    }


            //    /// <summary>
            //    /// 模块id
            //    /// </summary>
            //    /// <param name="moduleId"></param>
            //    /// <param name="departAlias"></param>
            //    /// <param name="currentAttdanceDate">考勤月份</param>
            //    /// <returns></returns>
            //    public string GetModuleName(int moduleId, string departAlias, string currentAttdanceDate)
            //    {
            //        string sql = string.Format("select  a.DisplayName  from   Base_MultiLang_Table_CN  a  where  a.ModuleId={0} ", moduleId);
            //        string name = SqlHelper.ExecuteScalar(1, sql).ToString();
            //        int index = name.LastIndexOf("@");
            //        return "CRM" + name.Substring(index + 1) + "-" + departAlias + "-" + currentAttdanceDate;
            //    }


            //    /// <summary>
            //    /// 根据个人汇总表中的dan号来查询当前员工所在部门
            //    /// </summary>
            //    /// <param name="dan"></param>  
            //    /// <returns></returns>
            //    public string GetDepartNameByDan(string dan)
            //    {
            //        string sql = string.Format(" select   a.DepartAlias  from   CUS_Table_5183  a  where  a.Dan='{0}'   ", dan.Substring(1, dan.Length - 2));
            //        return SqlHelper.ExecuteScalar(1, sql).ToString();
            //    }


            #endregion


        }


    }
}
