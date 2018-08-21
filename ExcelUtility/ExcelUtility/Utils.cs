using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    /// <summary>
    /// Utility class
    /// </summary>
    public static class Utils
    {

        /// <summary>
        /// Excel to DataTable
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="fileExt"></param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string excelFilePath,string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch(Exception ex) { }
            }
            return dtexcel;
        }


        /// <summary>
        /// DataTable to Excel
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool SaveDataTableExcel(DataTable dataTable ,string path)
        {
            try
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    dataTable.TableName = "Sheet1";
                    wb.Worksheets.Add(dataTable);
                    wb.SaveAs(path);
                    //Response.Clear();
                    //Response.Buffer = true;
                    //Response.Charset = "";
                    //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    //Response.AddHeader("content-disposition", "attachment;filename=GridView.xlsx");
                    //using (MemoryStream MyMemoryStream = new MemoryStream())
                    //{
                    //    wb.SaveAs(MyMemoryStream);
                    //    MyMemoryStream.WriteTo(Response.OutputStream);
                    //    Response.Flush();
                    //    Response.End();
                    //}
                }
            }
            catch (Exception ex)
            {

                return false;
            }
            return true;
        }
    }
}
