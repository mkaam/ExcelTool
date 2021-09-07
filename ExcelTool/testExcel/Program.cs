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
using CommandLine;
using NLog.Layouts;
using System.Diagnostics;

namespace testexcel
{
    class Program
    {
        //private static readonly Logger MyLog = LogManager.GetCurrentClassLogger();
        private static testExcel.CLogger MyLog;
        private static Stopwatch _watch;
        private static string ExePath;
        private static string RootPath;
        private static string LogPath;
        private static bool ParserError = false;

        static void Main(string[] args)
        {
            ExePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            RootPath = ExePath;
            LogPath = Path.Combine(RootPath, "Logs");

            MyLog = new CLogger("log");
            
            var parser = new Parser(config => { 
                config.CaseSensitive = false;
                config.IgnoreUnknownArguments = false;                
                config.AutoHelp = true;
                config.AutoVersion = true;
                config.HelpWriter = Console.Error;
            });

            parser.ParseArguments<ConfigMode, ImportMode, ExportMode>(args)
                .WithParsed<ConfigMode>(s => RunConfigMode(s))
                .WithParsed<ImportMode>(s => RunImportMode(s))
                .WithParsed<ExportMode>(s => RunExportMode(s))
                .WithNotParsed(errors => HandleParseError(errors));

            if (!ParserError)
            {
                _watch.Stop();
                MyLog.Debug($"Application Finished. Elapsed time: {_watch.ElapsedMilliseconds}ms");
            }            

        }

        static void LoggerConfigure(Options opts)
        {
            if (opts.LogFile == "")
            {
                //opts.LogFile = "";
            }
            var config = new NLog.Config.LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new NLog.Targets.FileTarget("logfile");
            if (opts.LogFile != null && opts.LogFile != "")
            {
                if (Path.GetFileName(opts.LogFile) == opts.LogFile)
                    logfile.FileName = $"{Path.Combine(Path.Combine(RootPath, "Logs"), opts.LogFile)}";
                else
                    logfile.FileName = $"{opts.LogFile}";
            }
            else
                logfile.FileName = $"{Path.Combine(Path.Combine(RootPath, "Logs"), $"{DateTime.Now.ToString("yyyyMMdd")}.log")}";

            logfile.MaxArchiveFiles = 60;
            logfile.ArchiveAboveSize = 10240000;

            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");
            if (opts.Verbose)
                config.AddRule(LogLevel.Trace, LogLevel.Fatal, logconsole);
            else
                config.AddRule(LogLevel.Error, LogLevel.Fatal, logconsole);

            config.AddRule(LogLevel.Trace, LogLevel.Fatal, logfile);

            // design layout for file log rotation
            CsvLayout layout = new CsvLayout();
            layout.Delimiter = CsvColumnDelimiterMode.Comma;
            layout.Quoting = CsvQuotingMode.Auto;
            layout.Columns.Add(new CsvColumn("Start Time", "${longdate}"));
            layout.Columns.Add(new CsvColumn("Elapsed Time", "${elapsed-time}"));
            layout.Columns.Add(new CsvColumn("Machine Name", "${machinename}"));
            layout.Columns.Add(new CsvColumn("Login", "${windows-identity}"));
            layout.Columns.Add(new CsvColumn("Level", "${uppercase:${level}}"));
            layout.Columns.Add(new CsvColumn("Message", "${message}"));
            layout.Columns.Add(new CsvColumn("Exception", "${exception:format=toString}"));
            logfile.Layout = layout;

            // design layout for console log rotation
            SimpleLayout ConsoleLayout = new SimpleLayout("${longdate}:${message}\n${exception}");
            logconsole.Layout = ConsoleLayout;

            // Apply config           
            NLog.LogManager.Configuration = config;
        }

        static int RunExportMode(Options opts)
        {
            var exitCode = 0;
            LoggerConfigure(opts);

            _watch = new Stopwatch();
            _watch.Start();
            MyLog.Debug("Application Start");

            try
            {
                var FileOnly = Path.GetFileName(opts.ExcelFile);
                var DirectoryOnly = Path.GetDirectoryName(opts.ExcelFile);
                IEnumerable<string> InputFiles = Enumerable.Empty<string>();
                IEnumerable<string> FileEnum = Enumerable.Empty<string>();
  
                var dictionary = new Dictionary<string, object>();

                dictionary.Add("Import/Export", (opts is ExportMode ? "Export" : "Import"));
                dictionary.Add("ExcelFile", opts.ExcelFile );
                dictionary.Add("SheetName", opts.SheetName);
                dictionary.Add("FirstRowisHeader", opts.FirstRowIsHeader ? "1" : "0");
                dictionary.Add("CellStart", opts.CellStart ?? "A1");
                dictionary.Add("CellEnd", opts.CellEnd ?? "");
                dictionary.Add("DBServer", opts.DbServer);
                dictionary.Add("DBName", opts.DbName);
                dictionary.Add("DBTable", opts.DbTable);
                dictionary.Add("ExportQueryFile", opts.ExportQuery);                    

                var tipeCsv = dictionary["Import/Export"];

                string[] columnsTypeToAdd = { };

                if (!dictionary.ContainsKey("Option") || dictionary["Option"].Equals(""))
                {

                    MyLog.Debug("Export Mode");
                    MyLog.Debug($"From : [{dictionary["DBServer"]}].{dictionary["DBName"]} Into : {dictionary["ExcelFile"]} using Query : {dictionary["ExportQueryFile"]}");
                    DatabaseToDataTable dt2d = new DatabaseToDataTable(MyLog);

                    DataTable dt; 
                    if (opts.QueryParameter != null) 
                        dt = dt2d.ExportToDatabase(dictionary,opts.QueryParameter);
                    else
                        dt = dt2d.ExportToDatabase(dictionary);

                    DataTableToExcel dt2e = new DataTableToExcel();
                    dt2e.ExportExcel(dt, dictionary);                   
                }
                else
                {
                    GetMetaData gm = new GetMetaData();
                    var auth = gm.GetMeta(dictionary);

                    MyLog.Info(auth);

                }
                
            }
            catch (Exception ex)
            {
                exitCode = -1;
                MyLog.Error("Error!", ex);

            }

            return exitCode;
        }

        static int RunImportMode(Options opts)
        {
            var exitCode = 0;
            LoggerConfigure(opts);

            _watch = new Stopwatch();
            _watch.Start();
            MyLog.Debug("Application Start");

            try
            {
                var FileOnly = Path.GetFileName(opts.ExcelFile);
                var DirectoryOnly = Path.GetDirectoryName(opts.ExcelFile);
                IEnumerable<string> InputFiles = Enumerable.Empty<string>();
                IEnumerable<string> FileEnum = Enumerable.Empty<string>();

                if (FileOnly.Contains('*'))
                    InputFiles = GetFiles(DirectoryOnly, new string[] { FileOnly });
                else
                    InputFiles = InputFiles.Concat(new[] { opts.ExcelFile });

                foreach (string InputFile in InputFiles)
                {

                    var dictionary = new Dictionary<string, object>
                    {
                        { "Import/Export", (opts is ExportMode ? "Export" : "Import") },
                        { "ExcelFile", InputFile },
                        { "SheetName", opts.SheetName },
                        { "FirstRowisHeader", opts.FirstRowIsHeader ? "1" : "0" },
                        { "CellStart", opts.CellStart ?? "" },
                        { "CellEnd", opts.CellEnd ?? "" },
                        { "DBServer", opts.DbServer },
                        { "DBName", opts.DbName },
                        { "DBTable", opts.DbTable },
                        { "ExportQueryFile", opts.ExportQuery }
                    };

                    var tipeCsv = dictionary["Import/Export"];

                    string[] columnsTypeToAdd = { };

                    if (!dictionary.ContainsKey("Option") || dictionary["Option"].Equals(""))
                    {
                   
                        MyLog.Debug("Import Mode");

                        MyLog.Debug($"From : {dictionary["ExcelFile"]} Into [{dictionary["DBServer"]}].{dictionary["DBName"]}.dbo.{dictionary["DBTable"]}");
                        ExcelToDatatable e2dt = new ExcelToDatatable(MyLog);
                        if (opts.TruncateTable)
                        {
                            MyLog.Debug($"Truncate Table : {dictionary["DBTable"]}");
                            e2dt.TruncateTable = true;
                        }
                        if (opts.SkipBlankRow)
                        {
                            MyLog.Debug($"Skip Blank Row Enabled");
                            e2dt.SkipBlankRow = true;
                        }
                        e2dt.Conversion(dictionary, columnsTypeToAdd);
                   
                    }
                    else
                    {
                        GetMetaData gm = new GetMetaData();
                        var auth = gm.GetMeta(dictionary);

                        MyLog.Info(auth);

                    }

                    if (opts.BackupPath != null && Directory.Exists(Path.GetDirectoryName(opts.BackupPath)) )
                    {
                        
                        var FileNameOnly = Path.GetFileName(InputFile);
                        var DateStr = DateTime.Now.ToString("yyyyMMddHHmmss");
                        var BackupFile = $"{Path.Combine(Path.GetFullPath(opts.BackupPath),$"{FileNameOnly}_{DateStr}")}";

                        if (opts.BackupMove) { 
                            File.Move(InputFile, BackupFile);
                            MyLog.Debug($"Move file {InputFile} to {BackupFile}. Done");
                        }
                        if (!opts.BackupMove) {
                            File.Copy(InputFile, BackupFile);
                            MyLog.Debug($"Copy file {InputFile} to {BackupFile}. Done");
                        }

                    }
                    

                }



            }
            catch (Exception ex)
            {
                exitCode = -1;
                MyLog.Error("Error!", ex);

            }

            return exitCode;
        }

        static int RunConfigMode(Options opts)
        {
            var exitCode = 0;
            LoggerConfigure(opts);

            _watch = new Stopwatch();
            _watch.Start();
            MyLog.Debug("Application Start");

            var dictionary = new Dictionary<string, object>();

            if (opts is ConfigMode && opts.Props != null && opts.Props != "")
            {
                string pathFile = opts.Props;
                //Read CSV File
                using (var reader = new StreamReader(pathFile))
                {
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        while (csv.Read())
                        {
                            dictionary.Add(csv.GetField(0), csv.GetField(1));
                        }
                    }
                }
                var tipeCsv = dictionary["Import/Export"];
                if (tipeCsv.Equals("Import"))
                    RunImportMode(opts);
                else
                    RunExportMode(opts);
            }
  
            return exitCode;
        }

        static void HandleParseError(IEnumerable<Error> errs)
        {
            ParserError = true;

            if (errs.Any(x => x is HelpRequestedError || x is VersionRequestedError))
            {
            }
            else
                Console.WriteLine("Parameter unknown, please check the documentation or use parameter '--help' for more information");

        }

        public static IEnumerable<string> GetFiles(string path,
                            string[] searchPatterns,
                            SearchOption searchOption = SearchOption.TopDirectoryOnly)
        {

            if (searchPatterns.Length == 0)
            {
                searchPatterns = new string[] { "*" };
            }

            return searchPatterns.AsParallel()
                   .SelectMany(searchPattern =>
                          Directory.EnumerateFiles(path, searchPattern, searchOption));

        }
    }
}