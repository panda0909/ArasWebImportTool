using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
namespace ArasWebImportTool
{
    public class ExcelLib
    {
        //============ Common Api ===============//
        /// <summary>
        /// ¦@¥ÎÅª¨úExcel to DataTable
        /// </summary>
        /// <param name="file_path"></param>
        /// <returns></returns>
        public static DataTable ReadExcelToDataTable(string file_path)
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (XLWorkbook workBook = new XLWorkbook(file_path))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                int ignore = 0,ignore_index=1;
                foreach (IXLRow row in workSheet.Rows())
                {
                    if(ignore_index <= ignore)
                    {
                        ignore_index++;
                        continue;
                    }
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        if (row.FirstCellUsed() == null || row.LastCellUsed() == null) continue;
                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            if (i >= dt.Columns.Count) break;
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dt;
            }
        }
        public static DataTable ReadExcelToDataTable(string file_path,string sheet_name)
        {
            try
            {
                using (XLWorkbook workBook = new XLWorkbook(file_path))
                {
                    //Read the first Sheet from Excel file.
                    IXLWorksheet workSheet = workBook.Worksheet(sheet_name);

                    //Create a new DataTable.
                    DataTable dt = new DataTable();

                    //Loop through the Worksheet rows.
                    bool firstRow = true;
                    int ignore = 0, ignore_index = 1;

                    foreach (IXLRow row in workSheet.Rows())
                    {
                        if (ignore_index <= ignore)
                        {
                            ignore_index++;
                            continue;
                        }
                        //Use the first row to add columns to DataTable.
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            //Add rows to DataTable.
                            dt.Rows.Add();
                            int i = 0;
                            if (row.FirstCellUsed() == null || row.LastCellUsed() == null) continue;
                            foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                            {
                                if (i >= dt.Columns.Count) break;
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }
                    }

                    return dt;
                }
            }
            catch(Exception ex)
            {
                return null;
            }
            
        }
        public static string SaveExcelFromDataTable(string file_path,DataTable dt)
        {
            XLWorkbook workBook = new XLWorkbook();

            // Add a DataTable as a worksheet
            dt.TableName = "Result";
            workBook.Worksheets.Add(dt);
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(file_path);
            string file_name = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            workBook.SaveAs(fileInfo.Directory+"\\"+ file_name+".xlsx");
            return fileInfo.Directory +"\\"+ file_name + ".xlsx";
        }
        public static string SaveExcelFromDataTable(string file_path, DataTable dt,string dt_name)
        {
            XLWorkbook workBook = new XLWorkbook();

            // Add a DataTable as a worksheet
            dt.TableName = dt_name;
            workBook.Worksheets.Add(dt);
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(file_path);
            string file_name = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
            workBook.SaveAs(fileInfo.Directory + "\\" + file_name + ".xlsx");
            return fileInfo.Directory + "\\" + file_name + ".xlsx";
        }
        public static List<string> GetExcelSheet(string file_path)
        {
            List<string> sheets = new List<string>();
            using (XLWorkbook workBook = new XLWorkbook(file_path))
            {
                foreach(IXLWorksheet sheet in workBook.Worksheets)
                {
                    sheets.Add(sheet.Name);
                }
            }
            return sheets;
        }
    }
}