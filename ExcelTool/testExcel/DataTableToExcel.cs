using System;
using CsvHelper;
using System.IO;
using System.Globalization;
using CsvHelper.Configuration;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using System.Linq;
using NPOI.SS.Util;
using Microsoft.Win32.SafeHandles;
using System.Text;
using NLog;
namespace testExcel
{
    class DataTableToExcel
    {
        
        public void ExportExcel(DataTable dt, Dictionary<string, object> dict)
        {
            //DeclareVariable
            var ExcelFile = (string)dict["ExcelFile"];
            var NameSheet = (string)dict["SheetName"];
            var FirstRowisHeader = (string)dict["FirstRowisHeader"];
            var CellStart = (string)dict["CellStart"];
            var CellEnd = (string)dict["CellEnd"];
            var serverName = (string)dict["DBServer"];
            var tableName = (string)dict["DBTable"];
            var ExportQueryFile = (string)dict["ExportQueryFile"];
            var dbName = (string)dict["DBName"];
            //IWorkbook xssfwb;
            var memoryStream = new MemoryStream();
            
            XSSFWorkbook workbook;
            XSSFSheet sheet;
            //CHECK FILE EXISTING
            if (File.Exists(ExcelFile))
            {
                FileStream file = new FileStream(ExcelFile, FileMode.Open, FileAccess.Read);
                workbook = new XSSFWorkbook(file);
                var sht = workbook.GetSheet(NameSheet);

                //CHECK IF SHEET EXISTING
                if (sht == null)
                {
                    sheet = workbook.CreateSheet(NameSheet) as XSSFSheet;
                }
                else
                {
                    //REMOVE EXISTING SHEET
                    //for (int i = 0; i <= workbook.NumberOfSheets; i++)
                    //{
                    //    if (NameSheet == workbook.GetSheetName(i))
                    //    {
                    //        workbook.RemoveSheetAt(i);
                    //    }
                    //}
                    //sheet = workbook.CreateSheet(NameSheet) as XSSFSheet;
                    sheet = (XSSFSheet)workbook.GetSheet(NameSheet);
                }
                file.Close();
            }
            else
            {
                workbook = new XSSFWorkbook();
                sheet = workbook.CreateSheet(NameSheet) as XSSFSheet;
            }
         
            using (var fs = new FileStream(ExcelFile, FileMode.Create, FileAccess.Write))
            {   
                //DATE FORMAT
                XSSFCellStyle dateStyle = workbook.CreateCellStyle() as XSSFCellStyle;
                XSSFDataFormat format = workbook.CreateDataFormat() as XSSFDataFormat;
                dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
                var cellAddr = new CellAddress(CellStart);
                var rowAddr = cellAddr.Row; 
                var colAddr = cellAddr.Column;

                //GET LENGTH OF CHAR
                int[] arrColWidth = new int[dt.Columns.Count];
                foreach (DataColumn item in dt.Columns)
                {
                    arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        int intTemp = Encoding.GetEncoding(936).GetBytes(dt.Rows[i][j].ToString()).Length;
                        if (intTemp > arrColWidth[j])
                        {
                            arrColWidth[j] = intTemp;
                        }
                    }
                }
                int rowIndex = 0;
               
                XSSFCellStyle headStyle = workbook.CreateCellStyle() as XSSFCellStyle;
                headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                XSSFFont font = workbook.CreateFont() as XSSFFont;
                font.FontHeightInPoints = 10;
                font.Boldweight = 700;
                headStyle.SetFont(font);

                foreach (DataRow row in dt.Rows)
                {
                    // CEK JIKA FIRSTROWISHEADER TERDAPAT KARAKTER * DI DEPAN, MAKA ROW ITU AKAN MENJADI HEADER
                    if (FirstRowisHeader.Equals("1"))
                    {
                        if (rowIndex == 0)
                        {
                           
                            //GIVE CELL STYLE
                            {
                                XSSFRow headerRow = sheet.CreateRow(rowAddr) as XSSFRow;
                                //headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                                //XSSFFont font = workbook.CreateFont() as XSSFFont;
                                //font.FontHeightInPoints = 10;
                                //font.Boldweight = 700;
                                //headStyle.SetFont(font);
                                foreach (DataColumn column in dt.Columns)
                                {
                                    
                                    headerRow.CreateCell(column.Ordinal + colAddr).SetCellValue(column.ColumnName);
                                    headerRow.GetCell(column.Ordinal + colAddr).CellStyle = headStyle;

                                    sheet.SetColumnWidth(column.Ordinal + colAddr, (arrColWidth[column.Ordinal] + 1) * 256);
                                }
                                //headerRow.Dispose();
                            }

                            rowIndex = rowAddr + 1;
                        }
                    }
                    
                   
                    //GIVE DATA TYPE
                    XSSFRow dataRow = sheet.CreateRow(rowIndex) as XSSFRow;
                    foreach (DataColumn column in dt.Columns)
                    {
                        
                        XSSFCell newCell = dataRow.CreateCell(column.Ordinal + colAddr) as XSSFCell;


                        string drValue = row[column].ToString();
                        //if (FirstRowisHeader.Equals("1"))
                        //{
                        //    XSSFRow headerRow = sheet.CreateRow(colAddr) as XSSFRow;
                        //}
                       
                    switch (column.DataType.ToString())
                        {
                            case "System.String":
                                double result;
                                if (double.TryParse(drValue, out result))
                                {
                                    
                                    double.TryParse(drValue, out result);
                                   
                                    newCell.SetCellValue(result);
                                    break;
                                }
                                else
                                {
                                  
                                    if (FirstRowisHeader.Equals("0"))
                                    {
                                        
                                        if (drValue != "")
                                        {
                                           
                                            string newDataValue = drValue.Substring(0, 1);
                                            if (newDataValue.Equals("*"))
                                            {
                                                
                                                drValue = drValue.Substring(1);
                                                newCell.CellStyle = headStyle;
                                                sheet.SetColumnWidth(column.Ordinal + colAddr, (arrColWidth[column.Ordinal] + 1) * 256);
                                            }
                                        }
                                    }
                                    newCell.SetCellValue(drValue);
                                    break;
                                }

                            case "System.DateTime":
                                DateTime dateV;
                                DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);

                                newCell.CellStyle = dateStyle;
                                break;
                            case "System.Boolean":
                                bool boolV = false;
                                bool.TryParse(drValue, out boolV);
                                newCell.SetCellValue(boolV);
                                break;
                            case "System.Int16":
                            case "System.Int32":
                            case "System.Int64":
                            case "System.Byte":
                                int intV = 0;
                                int.TryParse(drValue, out intV);
                                newCell.SetCellValue(intV);
                                break;
                            case "System.Decimal":
                            case "System.Double":
                                double doubV = 0;
                                double.TryParse(drValue, out doubV);
                                newCell.SetCellValue(doubV);
                                break;
                            case "System.DBNull":
                                newCell.SetCellValue("");
                                break;
                            default:
                                newCell.SetCellValue("");
                                break;
                        }

                    }
                    rowIndex++;
                    
                }
                workbook.Write(fs);
                fs.Close();


            }
        }

       
    }
}
