using System;
using System.Collections.Generic;
using CommandLine;
using CommandLine.Text;

namespace testExcel
{
    public class Options
    {
        public virtual bool Verbose { get; set; }
        public virtual string Props { get; set; }        
        public virtual string ExcelFile { get; set; }        
        public virtual string SheetName { get; set; }     
        public virtual bool FirstRowIsHeader { get; set; }  
        public virtual string CellStart { get; set; }       
        public virtual string CellEnd { get; set; }       
        public virtual string DbServer { get; set; }      
        public virtual string DbName { get; set; }    
        public virtual string DbTable { get; set; }
        public virtual string ExportQuery { get; set; }
        public virtual string LogFile { get; set; }
        public virtual bool TruncateTable { get; set; }
        public virtual bool SkipBlankRow { get; set; }
        public virtual string BackupPath { get; set; }
        public virtual bool BackupMove { get; set; }
        public virtual IEnumerable<string> QueryParameter { get; set; }
    }

    [Verb("config", isDefault: true, HelpText = "run Excel Tool using configuration file")]
    class ConfigMode : Options 
    {
        [Option(HelpText = "Print process output to console")]
        public override bool Verbose { get; set; }

        [Value(0, Required = true, HelpText = "you can choose setup parameter using configuration file. ex: E:\\SSISFiles\\Config.csv")]
        public override string Props { get; set; }
    }

    [Verb("import", HelpText = "Import Mode, import data from excel and save into database")]
    class ImportMode : Options 
    {
        [Option('i', "excelfile", Required = true, HelpText = "[excelfile] name to be imported. you can use prefix for multiple file to be imported. ex: E:\\file.xlsx or with prefix E:\\*.xlsx")]
        public override string ExcelFile { get; set; }

        [Option('s', "sheetname", Required = false, HelpText = "data is taken from this sheet name during [import] mode \n " +
        "The query results from the [exportquery] will be stored in this [sheetname] during [export] mode", Default = "Sheet1")]
        public override string SheetName { get; set; }

        [Option(HelpText = "call this param as switch on, if first row is header")]
        public override bool FirstRowIsHeader { get; set; }

        [Option('x', "cellstart", Required = false, HelpText = " if not set using default value = A1")]
        public override string CellStart { get; set; }

        [Option('y', "cellend", Required = false, HelpText = "default value = Last cell before blank = get last column using first row data, and get last row using first column data.")]
        public override string CellEnd { get; set; }

        [Option('a', "dbserver", Required = false, HelpText = "importing data from [excelfile] into this [servername]")]
        public override string DbServer { get; set; }

        [Option('b', "dbname", Required = false, HelpText = "importing data from [excelfile] into this database")]
        public override string DbName { get; set; }

        [Option('c', "dbtable", Required = false, HelpText = "importing data from [excelfile] into this table")]
        public override string DbTable { get; set; }

        [Option(HelpText = "Full path Logging file or use filename only, by default App rootpath will be used. eg: E:\\Interface\\Sampoerna\\QTAv2\\Logs\\LogAMI.txt or LogAMI.txt")]
        public override string LogFile { get; set; }

        [Option(HelpText = "Print process output to console")]
        public override bool Verbose { get; set; }

        [Option(HelpText = "Truncate destination table before import")]
        public override bool TruncateTable { get; set; }

        [Option(HelpText = "Don't insert Blank row")]
        public override bool SkipBlankRow { get; set; }

        [Option(HelpText = "Backup [ExcelFile] to this path. Backup disable if this option not define. Ex : E:\\Backup")]
        public override string BackupPath { get; set; }

        [Option(HelpText = "Backup using Move File operation")]
        public override bool BackupMove { get; set; }

    }

    [Verb("export", HelpText = "Export Mode, Export query result into Excel File")]
    class ExportMode : Options 
    {
        [Option('i', "excelfile", Required = true, HelpText = "[excelfile] for saving query result. ex: E:\\file.xlsx")]
        public override string ExcelFile { get; set; }

        [Option('s', "sheetname", Required = false, HelpText = "data is taken from this sheet name during [import] mode \n " +
"The query results from the [exportquery] will be stored in this [sheetname] during [export] mode", Default = "Sheet1")]
        public override string SheetName { get; set; }

        [Option('x', "cellstart", Required = false, HelpText = " if not set using default value = A1")]
        public override string CellStart { get; set; }

        [Option('y', "cellend", Required = false, HelpText = "default value = Last cell before blank = get last column using first row data, and get last row using first column data.")]
        public override string CellEnd { get; set; }

        [Option('a', "dbserver", Required = true, HelpText = "importing data from [excelfile] into this [servername]")]
        public override string DbServer { get; set; }

        [Option('b', "dbname", Required = true, HelpText = "importing data from [excelfile] into this database")]
        public override string DbName { get; set; }

        [Option('q', "exportquery", Required = true, HelpText = "define queryfilename to be executed during export mode", SetName = "export")]
        public override string ExportQuery { get; set; }

        [Option(HelpText = "call this param as switch on, if first row is header")]
        public override bool FirstRowIsHeader { get; set; }

        [Option(HelpText = "Full path Logging file or use filename only, by default App rootpath will be used. eg: E:\\Interface\\Sampoerna\\QTAv2\\Logs\\LogAMI.txt or LogAMI.txt")]
        public override string LogFile { get; set; }

        [Option(HelpText = "Print process output to console")]
        public override bool Verbose { get; set; }

        [Option(HelpText = "Query Parameter, eg: paramCustomerCode=JKT0232301")]
        public override IEnumerable<string> QueryParameter { get; set; }



    }
}
