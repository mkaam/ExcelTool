using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using NPOI.SS.Formula;
using NLog;
using NLog.Config;
using NLog.Targets;
namespace testExcel
{
 
    class DataTableToDatabase
    {
        private static Logger MyLog = LogManager.GetCurrentClassLogger();
        private void Input(SqlConnection sqlcon, DataTable dtColumns, DataTable dt, String tableName)
        {
           
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlcon))
            {
                int j = 0;
                for (int i = 0; i <= dt.Columns.Count-1; i++)
                {
                    bulkCopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dtColumns.Rows[i][0].ToString());
                    j++;
                }
                //foreach (DataRow row in dtColumns.Rows)
                //{
                //    bulkCopy.ColumnMappings.Add(dt.Columns[i].ColumnName, row[0].ToString());
                //    i++;
                //}

                bulkCopy.DestinationTableName = tableName;
                try
                {
                    bulkCopy.WriteToServer(dt);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        public DataTable ReadDatabase(SqlCommand cmd, string tableName)
        {
            var table = new DataTable();
            cmd.CommandText = "SELECT sys.columns.name FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo." + tableName + "')";
            var rdr = cmd.ExecuteReader();
            table.Load(rdr);
            rdr.Close();
            //foreach (DataRow row in table.Rows)
            //{
            //    foreach (DataColumn dc in table.Columns)
            //    {
            //        Console.WriteLine("read database " + row.ItemArray[dc.Ordinal].ToString());
            //    }
            //}
            return table;
        }

        public bool CheckExistingTable(SqlCommand cmd, string tableName)
        {
            bool exists;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" + tableName + "'";
            cmd.ExecuteNonQuery();
            if (Convert.ToBoolean(cmd.ExecuteScalar()) == true)
                exists = true;
            else
                exists = false;
            return exists;
        }

        private void CreateTableOnSQL(SqlCommand cmd, DataTable dt, string[] columnsToAdd, string tableName)
        {
            Console.WriteLine("Create Table on SQL");
            string paramsToPut = "";
            for (int i = 0; i < columnsToAdd.Length; i++)
            {
                string trimmed = dt.Columns[i].ToString().Replace(" ", "");
                if (trimmed == "Group")
                    trimmed += "s";
                if (i == 0)
                    paramsToPut += "[" + trimmed + "] " + columnsToAdd[i];
                else
                    paramsToPut += ", [" + trimmed + "] " + columnsToAdd[i];
            }
            cmd.CommandText += "CREATE TABLE [dbo].[" + tableName + "] (" + paramsToPut + ");";
            Console.WriteLine("CREATE TABLE " + cmd.CommandText.ToString());
        }

        public void InputToDatabase(DataTable dt, string serverName, string dbName, string tableName, string[] columnsTypeToAdd)
        {
            configureLogger();
            SqlConnection sqlcon = new SqlConnection("Integrated Security=true; server=" + @"" + serverName + ";" +
                                       "Trusted_Connection=yes;" +
                                       "database=" + dbName + "; " +
                                       "connection timeout=0");
            try
            {
                if (sqlcon.State == ConnectionState.Closed)
                    sqlcon.Open();
                Console.WriteLine("Open Connection Succeess");
            }
            catch (Exception e)
            {
                

                Console.WriteLine("Error: " + e);
            }

            try
            {
                using (sqlcon)
                {
                    using (SqlCommand cmd = sqlcon.CreateCommand())
                    {
                        bool exists = CheckExistingTable(cmd, tableName);
                        Console.WriteLine("Exist? " + exists.ToString());
                        //DataTable dtColumns = ReadDatabase(cmd, tableName);
                        if (exists)
                        {
                            DataTable dtColumns = ReadDatabase(cmd, tableName);
                            Input(sqlcon, dtColumns, dt, tableName);
                        }
                        else
                        {
                            CreateTableOnSQL(cmd, dt, columnsTypeToAdd, tableName);
                            cmd.ExecuteNonQuery();
                            DataTable dtColumns = ReadDatabase(cmd, tableName);
                            Input(sqlcon, dtColumns, dt, tableName);
                        }
                        try
                        {                            
                            cmd.ExecuteNonQuery();
                            sqlcon.Close();
                        }
                        catch (Exception e)
                        {
                            MyLog.Info("Done");
                            MyLog.Error(e, "Message : " + e.Message + ". StackTrace : " + e.StackTrace);
                            Console.WriteLine("Failure! Please read an error on 'logfile' folder");
                            Console.WriteLine("ERROR " + e.Message);
                        }
                    }
                }               
            }
            catch (SqlException e)
            {
                Console.WriteLine("ERROR " + e.Message + ". Error Number " + e.Number);
            }

        }

        static void configureLogger()
        {
            var config = new LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new FileTarget("logfile") { FileName = "logfile/" + DateTime.Now.ToString("yyyyMMdd") + ".txt" };
            var logconsole = new ConsoleTarget("logconsole");

            // Rules for mapping loggers to targets            
            config.AddRule(LogLevel.Info, LogLevel.Info, logfile);
            config.AddRule(LogLevel.Error, LogLevel.Error, logfile);

            // Apply config           
            LogManager.Configuration = config;
        }
    }
}
