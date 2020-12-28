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
using NLog;
using NLog.Config;
using NLog.Targets;

namespace testexcel
{
    class Program
    {

        private static Logger MyLog = LogManager.GetCurrentClassLogger();
        
        static void Main(string[] args)
        {
            configureLogger();
            string pathFile = args[0];
            try
            {
                //@"C:\Users\rdwianto1\Documents\Book1.csv"
                var dictionary = new Dictionary<string, object>();
                //Read CSV File
                using (var reader = new StreamReader(pathFile))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {   
                    while (csv.Read())
                    {
                        dictionary.Add(csv.GetField(0), csv.GetField(1));
                    
                    }

                }
                //Console.WriteLine("ini Option : " + option);
                var tipeCsv = dictionary["Import/Export"];
                string[] columnsTypeToAdd = { };
           
                if (!dictionary.ContainsKey("Option") || dictionary["Option"].Equals(""))
                {
                    if (tipeCsv.Equals("Import"))
                    {
                   
                            ExcelToDatatable e2dt = new ExcelToDatatable();
                            e2dt.Conversion(dictionary, columnsTypeToAdd);
                    
                    }
                    else
                    {
                        DatabaseToDataTable dt2d = new DatabaseToDataTable();
                        var dt = dt2d.ExportToDatabase(dictionary);

                        DataTableToExcel dt2e = new DataTableToExcel();
                        dt2e.ExportExcel(dt, dictionary);
                    }
                }
                else
                {
                    GetMetaData gm = new GetMetaData();
                    var auth = gm.GetMeta(dictionary);


                    Console.WriteLine(auth);

                }
            }
            catch (Exception ex)
            {
                MyLog.Info("Done");
                MyLog.Error(ex, "Message : " + ex.Message + ". StackTrace : " + ex.StackTrace);
                Console.WriteLine("Failure! Please read an error on 'logfile' folder");

                //Console.Read();
                //return;
                //string filePath = @"E:\ssisfiles\ExcelTool\Error.txt";


                //using (StreamWriter writer = new StreamWriter(filePath, true))
                //{
                //    writer.WriteLine("-----------------------------------------------------------------------------");
                //    writer.WriteLine("Date : " + DateTime.Now.ToString());
                //    writer.WriteLine();

                //    while (ex != null)
                //    {
                //        writer.WriteLine(ex.GetType().FullName);
                //        writer.WriteLine("Message : " + ex.Message);
                //        writer.WriteLine("StackTrace : " + ex.StackTrace);

                //        ex = ex.InnerException;
                //    }
                //}
            }
        }

        static void configureLogger()
        {
            var config = new LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new FileTarget("logfile") { FileName = "logfile/"+ DateTime.Now.ToString("yyyyMMdd") +".txt" };
            var logconsole = new ConsoleTarget("logconsole");

            // Rules for mapping loggers to targets            
            config.AddRule(LogLevel.Info, LogLevel.Info, logfile);
            config.AddRule(LogLevel.Error, LogLevel.Error, logfile);

            // Apply config           
            LogManager.Configuration = config;
        }



    }
}