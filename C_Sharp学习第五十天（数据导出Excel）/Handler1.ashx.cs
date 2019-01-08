using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace C_Sharp学习第五十天_数据导出Excel_
{
    /// <summary>
    /// Handler1 的摘要说明
    /// </summary>
    public class Handler1 : IHttpHandler
    {
        HttpRequest request;
        HttpResponse response;

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            request = context.Request;
            response = context.Response;
            string action = request.QueryString["action"];
            //switch (action)
            //{
            //    case "ExportExcel":
            //        GetExcel();
            //        break;
            //    default:
            //        break;
            //}
            
            //context.Response.Write(name+"李四王五赵六");
        }

        // <summary>
        /// 将生成excel的路径返回
        /// </summary>
        /// <returns></returns>
        //public string ExportExcel()
        //{
        //    string arrListStr = Request["arr"].ToString();
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
        //    return GetWorkbook(dt, departAlias, 5184);
        //}

        /// <summary>
        /// 获取需要导出的全部数据   
        /// </summary>
        /// <returns></returns>
        public DataTable GetDataTable(string arrListStr)
        {
            //if (!string.IsNullOrEmpty(arrListStr))
            //{
            //    string sqlQueryUserDetails = string.Format(" select " +
            //        " De_29060 as  '工号'  ,De_29059Alias  as  '姓名',De_29061  as '事假（H）' ,De_31824  as  '事假明细', " +
            //        " De_29062  as  '病假（H）'  ,De_31825   as   '病假明细'  ,De_29063   as   '出勤薪资延缓发放时间（H）', " +
            //        " De_31827  as  '出勤薪资缓发明细' ,De_29064  as  '工作日加班薪资延缓发放时间（H）'  ,De_31828  as  '工作日加班薪资缓发明细'  ," +
            //        " De_31838  as  '休息日和节假日加班薪资延缓发放时间（H）' ,De_31839  as  '休息日和节假日薪资缓发明细' ,De_29065   as   '迟到（次）'  ," +
            //        " De_29066  as  '工作日加班（H）',De_31829 as  '工作日加班明细' ,De_29067 as  '休息日加班（H）'  ," +
            //        " De_31830  as  '休息日加班明细',De_29068  as  '节假日加班（H）'  ,De_31831   as   '节假日加班明细' ," +
            //        " De_29069  as  '年假（天）',De_31826  as  '年假明细',De_31848   as  '调休假总计（H）'  ," +
            //        " De_31849  as  '调休假明细',De_29071  as  '应出勤（天）'  ,b.Memo  as  '备注'  " +
            //        " from   CUS_Table_5183_Details  a   " +
            //        " inner  join  CUS_Table_5183  b   on  a.InnerId=b.Id " +
            //        " where  b.Dan  in  ({0}); ", arrListStr);
            //    return SqlHelper.ExecuteDataset(1, sqlQueryUserDetails).Tables[0];
            //}
            return new DataTable();
        }


        /// <summary>
        /// 设置工作簿的格式
        /// </summary>
        /// <returns></returns>
        public string GetWorkbook(DataTable dt, string departName, int moduleId)
        {
            #region  创建表及样式
            IWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("明细");
            IDataFormat dataFormat = workbook.CreateDataFormat();
            IFont fontTitle = workbook.CreateFont();
            fontTitle.FontName = "等线";
            fontTitle.FontHeight = 15;
            fontTitle.FontHeightInPoints = 15;

            IFont fontField = workbook.CreateFont();
            fontField.FontName = "等线";
            fontField.FontHeight = 10;
            fontField.FontHeightInPoints = 10;

            //创建style   四种   表名   字段名   字段间隔显示
            //表名的样式
            ICellStyle cellStyleTitle = workbook.CreateCellStyle();
            //居中
            cellStyleTitle.Alignment = HorizontalAlignment.Center;
            cellStyleTitle.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleTitle.BorderBottom = BorderStyle.None;
            cellStyleTitle.BorderLeft = BorderStyle.None;
            cellStyleTitle.BorderRight = BorderStyle.None;
            cellStyleTitle.BorderTop = BorderStyle.None;
            //字体
            cellStyleTitle.SetFont(fontTitle);

            //字段名称样式
            ICellStyle cellStyleField = workbook.CreateCellStyle();
            cellStyleField.Alignment = HorizontalAlignment.Center;
            cellStyleField.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleField.BorderBottom = BorderStyle.None;
            cellStyleField.BorderLeft = BorderStyle.None;
            cellStyleField.BorderRight = BorderStyle.None;
            cellStyleField.BorderTop = BorderStyle.None;
            //字体
            cellStyleField.SetFont(fontField);
            //前景色
            cellStyleField.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            cellStyleField.FillPattern = FillPattern.SolidForeground;

            //字段值的样式  单数
            ICellStyle cellStyleOddRow = workbook.CreateCellStyle();
            cellStyleOddRow.Alignment = HorizontalAlignment.Center;
            cellStyleOddRow.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleOddRow.BorderBottom = BorderStyle.None;
            cellStyleOddRow.BorderLeft = BorderStyle.None;
            cellStyleOddRow.BorderRight = BorderStyle.None;
            cellStyleOddRow.BorderTop = BorderStyle.None;
            //字体
            cellStyleOddRow.SetFont(fontField);

            //字段值的样式  单数   int 类型
            ICellStyle cellStyleOddRowInt = workbook.CreateCellStyle();
            cellStyleOddRowInt.Alignment = HorizontalAlignment.Center;
            cellStyleOddRowInt.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleOddRowInt.BorderBottom = BorderStyle.None;
            cellStyleOddRowInt.BorderLeft = BorderStyle.None;
            cellStyleOddRowInt.BorderRight = BorderStyle.None;
            cellStyleOddRowInt.BorderTop = BorderStyle.None;
            //字体
            cellStyleOddRowInt.SetFont(fontField);
            cellStyleOddRowInt.DataFormat = dataFormat.GetFormat("0");

            //字段值的样式  单数   double 类型
            ICellStyle cellStyleOddRowDouble = workbook.CreateCellStyle();
            cellStyleOddRowDouble.Alignment = HorizontalAlignment.Center;
            cellStyleOddRowDouble.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleOddRowDouble.BorderBottom = BorderStyle.None;
            cellStyleOddRowDouble.BorderLeft = BorderStyle.None;
            cellStyleOddRowDouble.BorderRight = BorderStyle.None;
            cellStyleOddRowDouble.BorderTop = BorderStyle.None;
            //字体
            cellStyleOddRowDouble.SetFont(fontField);
            cellStyleOddRowDouble.DataFormat = dataFormat.GetFormat("0.0");


            //字段值的样式  双数
            ICellStyle cellStyleEvenRow = workbook.CreateCellStyle();
            cellStyleEvenRow.Alignment = HorizontalAlignment.Center;
            cellStyleEvenRow.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleEvenRow.BorderBottom = BorderStyle.None;
            cellStyleEvenRow.BorderLeft = BorderStyle.None;
            cellStyleEvenRow.BorderRight = BorderStyle.None;
            cellStyleEvenRow.BorderTop = BorderStyle.None;
            //字体
            cellStyleEvenRow.SetFont(fontField);
            //前景色
            cellStyleEvenRow.FillForegroundColor = HSSFColor.LightGreen.Index;
            cellStyleEvenRow.FillPattern = FillPattern.SolidForeground;


            //字段值的样式  双数   int类型
            ICellStyle cellStyleEvenRowInt = workbook.CreateCellStyle();
            cellStyleEvenRowInt.Alignment = HorizontalAlignment.Center;
            cellStyleEvenRowInt.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleEvenRowInt.BorderBottom = BorderStyle.None;
            cellStyleEvenRowInt.BorderLeft = BorderStyle.None;
            cellStyleEvenRowInt.BorderRight = BorderStyle.None;
            cellStyleEvenRowInt.BorderTop = BorderStyle.None;
            //字体
            cellStyleEvenRowInt.SetFont(fontField);
            //前景色
            cellStyleEvenRowInt.FillForegroundColor = HSSFColor.LightGreen.Index;
            cellStyleEvenRowInt.FillPattern = FillPattern.SolidForeground;
            cellStyleEvenRowInt.DataFormat = dataFormat.GetFormat("0");


            //字段值的样式  双数  double类型
            ICellStyle cellStyleEvenRowDouble = workbook.CreateCellStyle();
            cellStyleEvenRowDouble.Alignment = HorizontalAlignment.Center;
            cellStyleEvenRowDouble.VerticalAlignment = VerticalAlignment.Center;
            //边框
            cellStyleEvenRowDouble.BorderBottom = BorderStyle.None;
            cellStyleEvenRowDouble.BorderLeft = BorderStyle.None;
            cellStyleEvenRowDouble.BorderRight = BorderStyle.None;
            cellStyleEvenRowDouble.BorderTop = BorderStyle.None;
            //字体
            cellStyleEvenRowDouble.SetFont(fontField);
            //前景色
            cellStyleEvenRowDouble.FillForegroundColor = HSSFColor.LightGreen.Index;
            cellStyleEvenRowDouble.FillPattern = FillPattern.SolidForeground;
            cellStyleEvenRowDouble.DataFormat = dataFormat.GetFormat("0.0");



            #endregion
            //数值类型的字段集合
            List<string> list = new List<string>()
            {
                "事假（H）","病假（H）","出勤薪资延缓发放时间（H）","工作日加班薪资延缓发放时间（H）","休息日和节假日加班薪资延缓发放时间（H）",
                "工作日加班（H）","休息日加班（H）","节假日加班（H）","年假（天）","调休假总计（H）"
            };


            //给单元格赋值
            //表的标题
            IRow rowTitle = sheet.CreateRow(0);
            ICell cellTitle = rowTitle.CreateCell(0);
            string fileName = GetModuleName(moduleId, departName);
            cellTitle.SetCellValue(fileName);
            cellTitle.CellStyle = cellStyleTitle;

            //字段名称
            IRow rowField = sheet.CreateRow(1);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cellField = rowField.CreateCell(i);
                cellField.SetCellValue(dt.Columns[i].ColumnName);
                cellField.CellStyle = cellStyleField;
            }

            if (dt != null && dt.Rows.Count != 0)
            {
                //字段值
                for (int i = 2; i < dt.Rows.Count + 2; i++)
                {
                    //先创建行
                    IRow rowValue = sheet.CreateRow(i);
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        ICell cellValue = rowValue.CreateCell(k);
                        if (i % 2 == 0)
                        {
                            if (list.Contains(dt.Columns[k].ColumnName))
                            {
                                cellValue.SetCellType(CellType.Numeric);
                                //cellValue.SetCellValue(dt.Rows[i - 2][k].ToDouble());
                                cellValue.CellStyle = cellStyleOddRowDouble;
                            }
                            else
                            {
                                cellValue.SetCellValue(dt.Rows[i - 2][k].ToString());
                                cellValue.CellStyle = cellStyleOddRow;
                            }

                        }
                        else
                        {
                            if (list.Contains(dt.Columns[k].ColumnName))
                            {
                                cellValue.SetCellType(CellType.Numeric);
                                //cellValue.SetCellValue(dt.Rows[i - 2][k].ToDouble());
                                cellValue.CellStyle = cellStyleEvenRowDouble;
                            }
                            else
                            {
                                cellValue.SetCellValue(dt.Rows[i - 2][k].ToString());
                                cellValue.CellStyle = cellStyleEvenRow;
                            }
                        }
                        sheet.SetColumnWidth(k, (Encoding.UTF8.GetBytes(dt.Rows[i - 2][k].ToString()).Length) * 256);

                        ////另一种设置方法
                        ////先拿到当前cell的宽度
                        //int columnWidth = sheet.GetColumnWidth(k);
                        ////拿到当前单元格中字符串的宽度
                        //int length = Encoding.UTF8.GetBytes(dt.Rows[i - 2][k].ToString()).Length;
                        //if (columnWidth < length + 1)
                        //{
                        //    columnWidth = length + 1;
                        //}
                        //sheet.SetColumnWidth(k, columnWidth * 256);
                    }
                }
            }
            else
            {
                //没有查询到数据
                //应该显示暂无数据
                IRow rowContent = sheet.CreateRow(2);
                ICell cellContent = rowContent.CreateCell(0);
                cellContent.SetCellValue("暂无内容");
                cellContent.CellStyle = cellStyleOddRow;
                sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, dt.Columns.Count - 1));

            }
            //合并第一行  
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dt.Columns.Count - 1));
            //给第一列增加筛选
            CellRangeAddress a2 = CellRangeAddress.ValueOf("A2");
            sheet.SetAutoFilter(a2);
            //自动调整列宽
            ISheet sheetObj = workbook.GetSheetAt(0);
            //sheetObj.AutoSizeColumn(1);
            //sheetObj.AutoSizeColumn(6);


            //先要在指定位置创建一个文件夹 
            //这里拿到当前应用程序域的跟目录     
            string rootPath = System.AppDomain.CurrentDomain.BaseDirectory;
            string newDirName = "TenantInfo\\ExportDetailsDataToExcel\\部门考勤汇总明细表\\";
            string newPath = Path.Combine(rootPath, newDirName);
            Directory.CreateDirectory(newPath);
            using (FileStream fsWrite = new FileStream(newPath + fileName + ".xls", FileMode.OpenOrCreate, FileAccess.Write))
            {
                //全部创建完成之后，通过写入流的方式将文件上传，或者写入磁盘
                workbook.Write(fsWrite);
                fsWrite.Flush();
                fsWrite.Close();
            }
            newDirName = newDirName.Replace("\\", "/");
            //这里注意的是只需要返回相对路径
            return string.Format("../../{0}{1}.xls", newDirName, fileName);
        }


        /// <summary>
        /// 模块id
        /// </summary>
        /// <param name="moduleId"></param>
        /// <param name="departAlias"></param>
        /// <returns></returns>
        public string GetModuleName(int moduleId, string departAlias)
        {
            string sql = string.Format("select  a.DisplayName  from   Base_MultiLang_Table_CN  a  where  a.ModuleId={0} ", moduleId);
            //string name = SqlHelper.ExecuteScalar(1, sql).ToString();
            string name = "张三的模块";
            int index = name.LastIndexOf("@");
            return "CRM" + name.Substring(index + 1) + "-" + departAlias;
        }


        /// <summary>
        /// 根据个人汇总表中的dan号来查询当前员工所在部门
        /// </summary>
        /// <param name="dan"></param>
        /// <returns></returns>
        public string GetDepartNameByDan(string dan)
        {
            string sql = string.Format(" select   a.DepartAlias  from   CUS_Table_5183  a  where  a.Dan='{0}'   ", dan.Substring(1, dan.Length - 2));
            //eturn SqlHelper.ExecuteScalar(1, sql).ToString();

            //这里是返回部门名称   
            return "部门名称";

        }




        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}