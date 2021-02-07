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
using testExcel;
using Newtonsoft.Json;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace testExcel
{
    class GetMetaData
    {
        public string GetMeta(Dictionary<string, object> dictionary)
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
            var option = dictionary["Option"];
            string newstring = "";
            IWorkbook workbook;
            using (FileStream stream = new FileStream(ExcelFile, FileMode.Open, FileAccess.Read))
            {
                newstring = ExcelFile.Substring(ExcelFile.Length - 4, 4);
                
                //checking excel type
                if (newstring.ToLower() == "xlsx")
                {                    
                    workbook = new XSSFWorkbook(stream);

                }
                else
                {                    
                    workbook = new HSSFWorkbook(stream);

                }
            }

            var auth = GetAuthor(workbook);

            return auth;

        }

        public static String GetAuthor(IWorkbook workbook)
        {
            var Author = "";
            if (workbook is NPOI.XSSF.UserModel.XSSFWorkbook)
            {
                var xssfWorkbook = workbook as NPOI.XSSF.UserModel.XSSFWorkbook;
                var xmlProps = xssfWorkbook.GetProperties();
                var coreProps = xmlProps.CoreProperties;
                Author = coreProps.LastModifiedByUser;
            }

            if (workbook is NPOI.HSSF.UserModel.HSSFWorkbook)
            {
                var hssfWorkbook = workbook as NPOI.HSSF.UserModel.HSSFWorkbook;
                var summaryInfo = hssfWorkbook.SummaryInformation;

                Author = summaryInfo.LastAuthor;
            }

            return Author;
        }
    }
}
