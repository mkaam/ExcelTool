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

using Newtonsoft.Json;
namespace testExcel
{
    class ExcelToDatatable
    {
        private CLogger MyLog;

        public ExcelToDatatable(CLogger logger)
        {
            MyLog = logger;
            TruncateTable = false;
            SkipBlankRow = false;
        }

        public bool TruncateTable { get; set; }

        public bool SkipBlankRow { get; set; }

        public IWorkbook GenerateWorkbookHelper(IWorkbook OriginalWorkbook, ISheet OriginalSheet, string range)
        {
            IWorkbook wb = OriginalWorkbook;
            ISheet sh = wb.CreateSheet("tmp");

            string[] cellStartStop = range.Split(':');

            CellReference cellRefStart = new CellReference(cellStartStop[0]);
            CellReference cellRefStop = new CellReference(cellStartStop[1]);

            for (var i = cellRefStart.Row; i < cellRefStop.Row; i++)
            {
                if (OriginalSheet.GetRow(i).Cells.All(d => d.CellType == CellType.Blank))
                {
                    break;
                }
                IRow rw = sh.CreateRow(i);
                for (var j = cellRefStart.Col; j <= cellRefStop.Col; j++)
                {                    
                    rw.Cells.Add(OriginalSheet.GetRow(i).GetCell(j));                    
                }

                
            }

            var aaa = sh.GetRowEnumerator();
            

            return wb;
        }

        public static ICell[,] GetRange(ISheet sheet, string range)
        {
            string[] cellStartStop = range.Split(':');

            CellReference cellRefStart = new CellReference(cellStartStop[0]);
            CellReference cellRefStop = new CellReference(cellStartStop[1]);

            ICell[,] cells = new ICell[cellRefStop.Row - cellRefStart.Row +1, cellRefStop.Col - cellRefStart.Col +1];


            for (int i = cellRefStart.Row; i < cellRefStop.Row ; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row.Cells.All(d => d.CellType == CellType.Blank))
                {                    
                    break;
                }
                if (i == 39)
                {
                    var asd = "";
                }
                for (int j = cellRefStart.Col; j < cellRefStop.Col + 1; j++)
                {
                    var aaa = row.GetCell(j).CellType;
                    cells[i - cellRefStart.Row, j - cellRefStart.Col] = row.GetCell(j);
                }
            }

            return cells;
        }

        public static T1[,] GetCellValues<T1>(ICell[,] cells)
        {
            T1[,] values = new T1[cells.GetLength(0), cells.GetLength(1)];

            for (int i = 0; i < values.GetLength(0); i++)
            {
                for (int j = 0; j < values.GetLength(1); j++)
                {
                    if (typeof(T1) == typeof(double) || typeof(T1) == typeof(int) ||


                        typeof(T1) == typeof(float) || typeof(T1) == typeof(long))
                    {
                        values[i, j] = (T1)Convert.ChangeType(cells[i, j].NumericCellValue, typeof(T1));


                    }
                    else if (typeof(T1) == typeof(DateTime))
                    {
                        values[i, j] = (T1)Convert.ChangeType(cells[i, j].DateCellValue, typeof(T1));


                    }
                    else if (typeof(T1) == typeof(string))
                    {
                        values[i, j] = (T1)Convert.ChangeType(cells[i, j].StringCellValue, typeof(T1));


                    }
                }
            }

            return values;
        }



        public object Conversion(Dictionary<string, object> dictionary, string[] columnsTypeToAdd)
        {
            //DeclareVariable
            var ExcelFile = ((string)dictionary["ExcelFile"]).ToLower();
            var NameSheet = (string)dictionary["SheetName"];
            var FirstRowisHeader = (string)dictionary["FirstRowisHeader"];
            var CellStart = (string)dictionary["CellStart"];
            var CellEnd = (string)dictionary["CellEnd"];
            var DBServer = (string)dictionary["DBServer"];
            var DBTable = (string)dictionary["DBTable"];
            var ExportQueryFile = (string)dictionary["ExportQueryFile"];
            var DBName = (string)dictionary["DBName"];
            string newstring = "";
            IWorkbook workbook;
            using (FileStream stream = new FileStream(ExcelFile, FileMode.Open, FileAccess.Read))
            {
                newstring = Path.GetExtension(ExcelFile); //ExcelFile.Substring(ExcelFile.Length - 4, 4);
             
                //checking excel type
                if (newstring.ToLower() == ".xlsx")
                {
                    MyLog.Debug("Excel Type : xlsx");                   
                    workbook = new XSSFWorkbook(stream);
                }
                else
                {
                    MyLog.Debug("Excel Type : xls");                    
                    workbook = new HSSFWorkbook(stream);
                }
            }            

            //Get Sheet Name
            ISheet sheet;
            

            var sht = workbook.GetSheet(NameSheet);
            if (sht == null)
            {
                sheet = workbook.GetSheetAt(0);
            }
            else
            {
                sheet = workbook.GetSheet(NameSheet);
            }
                       
            IRow headerRow;
            DataTable dt = new DataTable(sheet.SheetName);


            if (CellEnd != "")
            {
                //<<START WHEN CELLEND NOT EMPTY>>
                var range = "" + CellStart + ":" + CellEnd + "";
                var cellRange = CellRangeAddress.ValueOf(range);

                //var cells = GetCellValues<ICell>(GetRange(sheet, range));

                //IWorkbook wb = GenerateWorkbookHelper(workbook, sheet, range);
                
                //if (row.Cells.All(d => d.CellType == CellType.Blank)) RowIsBlank = true;
                //else RowIsBlank = false;


                var header = sheet.GetRow(cellRange.FirstRow);
                var headerIdx = 0;

                #region Generate Header
                for (var j = cellRange.FirstColumn; j <= cellRange.LastColumn; j++)
                {
                    if (FirstRowisHeader.Equals("1"))
                    {
                        if (header.GetCell(j) != null)
                        {
                            dt.Columns.Add(header.GetCell(j).ToString());
                        }
                        else
                        {
                            dt.Columns.Add("Column" + headerIdx);
                        }
                    }
                    else if (FirstRowisHeader.Equals("0"))
                    {
                        dt.Columns.Add("Column" + headerIdx.ToString());
                    }
                    headerIdx++;
                }
                #endregion

                //columnsTypeToAdd = new string[sheet.GetRow(cellRange.FirstRow).LastCellNum];

                columnsTypeToAdd = new string[headerIdx];
                // modified by Aam on 20201028 to skip first row if config FirstRowisHeader = 1                
                int firstRowAdd = FirstRowisHeader.Equals("1") ? 1 : 0;

                for (var i = cellRange.FirstRow + (firstRowAdd); i <= (cellRange.LastRow); i++)
                {
                    var row = sheet.GetRow(i);
                                      
                    DataRow dataRow = dt.NewRow();
                    bool[] isNull = new bool[dt.Columns.Count];

                    for (var j = cellRange.FirstColumn; j <= (cellRange.LastColumn); j++)
                    {


                        //if (i > cellRange.FirstRow)
                        //{
                        int indexDr = j - cellRange.FirstColumn;
                        if (newstring == ".xlsx")
                        {
                            IFormulaEvaluator fm = new XSSFFormulaEvaluator(workbook);

                            try
                            {
                                inputData(row.GetCell(j), indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);
                            }
                            catch
                            {

                                inputData(null, indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);

                            }


                        }
                        else
                        {
                            IFormulaEvaluator fm = new HSSFFormulaEvaluator(workbook);
                            try
                            {
                                inputData(row.GetCell(j), indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);
                            }
                            catch
                            {

                                inputData(null, indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);

                            }

                        }
                        //}
                    }

                    dt.Rows.Add(dataRow);
                    


                }
                //<<END WHEN CELLEND NOT EMPTY>>
             
            }
            else
            {
                //<<START CELLEND IS EMPTY>>
                //Get Specified Cell If null it will be start at index 0
                if (CellStart == "")
                {
                    CellStart = "A1";
                }
                var cr = new CellReference(CellStart);//for get by cellstart
                columnsTypeToAdd = new string[sheet.GetRow(cr.Row).LastCellNum];

                //WriteHeader
                headerRow = sheet.GetRow(cr.Row);
                var headerIdx = 0;
               
                foreach (var headerCell in headerRow)
                {
                    //if (FirstRowisHeader.Equals("1") && headerCell.ToString() != null && headerCell.ToString() != "")
                    if (FirstRowisHeader.Equals("1"))
                    {                        
                        dt.Columns.Add(headerCell.ToString());                        
                    }
                    else if (FirstRowisHeader.Equals("0"))
                    {
                        dt.Columns.Add(headerIdx.ToString());
                    }
                    headerIdx++;
                }
                // write the rest
                int i = 0; //for index datarow

                // modified by Aam on 20201028 to skip first row if config FirstRowisHeader = 1                
                int firstRowAdd = FirstRowisHeader.Equals("1") ? 1 : 0;

                for (var y = cr.Row + (firstRowAdd); y <= sheet.LastRowNum; y++)
                {
                    var row = sheet.GetRow(y);  // skip header row
             
                    DataRow dataRow = dt.NewRow();
                    //if (i > 0)
                    //{
                        bool[] isNull = new bool[dt.Columns.Count];
                        int z = 1; // for looping row in excel cz row.Cells.Count is start from 1
                        //for (var j = cr.Col; z <= row.Cells.Count; j++) //di rubah pertanggal 31-8-2020 karena selip antara header dengn detailnya
                        for (var j = cr.Col; z <= dt.Columns.Count; j++)
                        {
                            int indexDr = j - cr.Col;
                            if (newstring == ".xlsx")
                            {
                                IFormulaEvaluator fm = new XSSFFormulaEvaluator(workbook);

                                try //try if row.GetCell(j) is error because of bad formating template or offside index then make default to null
                                {
                                    inputData(row.GetCell(j), indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);
                               }
                               catch
                               {
                                    inputData(null, indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);
                               }

                            }
                            else
                            {
                                IFormulaEvaluator fm = new HSSFFormulaEvaluator(workbook);

                                try //try if row.GetCell(j) is error because of offside index then make default to null
                                {
                                    inputData(row.GetCell(j), indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);
                                }
                                catch
                                {
                                    inputData(null, indexDr, fm, workbook, dataRow, columnsTypeToAdd, isNull);
                                }
                            }
                            z++;
                        }
                        
                    //}

                    //if (i > 0)
                    //{
                        dt.Rows.Add(dataRow);
                    //}
                    //i++;
                }
                //<<END CELLEND IS EMPTY>>
            }
            if (SkipBlankRow) RemoveNullColumnFromDataTable(dt);
            DataTableToDatabase dt2db = new DataTableToDatabase(MyLog);
            columnsTypeToAdd = columnsTypeToAdd.Where(x => !string.IsNullOrEmpty(x)).ToArray();
            dt2db.IsTruncateTable = TruncateTable;
            dt2db.InputToDatabase(dt, DBServer, DBName, DBTable, columnsTypeToAdd);

            return dt2db;
            
        }

        public static void RemoveNullColumnFromDataTable(DataTable dt)
        {
            dt.AsEnumerable().Where(row => 
                row.ItemArray.All(field => field == null | field == DBNull.Value | field.Equals(""))).ToList()
                .ForEach(row => row.Delete());
            dt.AcceptChanges();
        }

        private void inputData(ICell cell, int i, IFormulaEvaluator formula, IWorkbook wb, DataRow dr, string[] columnsType, bool[] isNull)
        {
            formula.EvaluateInCell(cell);
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Unknown:
                        dr[i] = "'" + cell.StringCellValue + "'";
                        columnsType[i] = "nvarchar(MAX)";
                        break;
                    case CellType.Blank:
                        dr[i] = null;
                        columnsType[i] = "nvarchar(MAX)";
                        break;
                    case CellType.Boolean:
                        //if(cell.BooleanCellValue == true)
                        //{
                        //    dr[i] = Convert.ToByte(1);
                        //}
                        //else
                        //{
                        //    dr[i] = Convert.ToByte(0); ;
                        //}                   
                        dr[i] = cell.BooleanCellValue;
                        columnsType[i] = "Bit";
                        break;
                    case CellType.Numeric:

                        dr[i] = DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue.ToString() : cell.NumericCellValue.ToString(); ;
                        columnsType[i] = "numeric(18,8)";
                        break;
                    case CellType.String:
                        dr[i] = cell.StringCellValue;
                        columnsType[i] = "nvarchar(MAX)";
                        break;
                    case CellType.Error:
                        dr[i] = cell.ErrorCellValue;
                        columnsType[i] = "nvarchar(MAX)";
                        break;
                    case CellType.Formula:
                    default:
                        IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();
                        if (cell.CellType == CellType.Formula)
                        {
                            switch (eval.EvaluateFormulaCell(cell))
                            {
                                case CellType.Boolean:
                                    dr[i] = cell.BooleanCellValue;
                                    columnsType[i] = "varchar(255)";
                                    break;
                                case CellType.Numeric:
                                    dr[i] = cell.ToString();
                                    columnsType[i] = "numeric(18,8)";
                                    break;
                                case CellType.String:
                                    dr[i] = cell.StringCellValue;
                                    columnsType[i] = "nvarchar(MAX)";
                                    break;
                            }
                        }
                        dr[i] = "=" + cell.CellFormula;
                        break;
                }
            }
            catch
            {
               
                dr[i] = null;
                columnsType[i] = "nvarchar(MAX)";
               
            }
            
        }

       


    }
}
