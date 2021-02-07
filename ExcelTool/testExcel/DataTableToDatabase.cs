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
        private CLogger MyLog;// = LogManager.GetCurrentClassLogger();
        public DataTableToDatabase(CLogger logger)
        {
            MyLog = logger;
        }

        public bool IsTruncateTable { get; set; }

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
                    MyLog.Error("Error!", ex);
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

        private void TruncateTable(SqlCommand cmd, string tableName)
        {            
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = $"truncate table [{tableName}]";
            cmd.ExecuteNonQuery();            
        }

        private void CreateTableOnSQL(SqlCommand cmd, DataTable dt, string[] columnsToAdd, string tableName)
        {
            MyLog.Debug("Create Table on SQL");            
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
            MyLog.Debug($"Execute : {cmd.CommandText.ToString()}");
        }

        public void InputToDatabase(DataTable dt, string serverName, string dbName, string tableName, string[] columnsTypeToAdd)
        {           
            SqlConnection sqlcon = new SqlConnection("Integrated Security=true; server=" + @"" + serverName + ";" +
                                       "Trusted_Connection=yes;" +
                                       "database=" + dbName + "; " +
                                       "connection timeout=0");
            try
            {
                if (sqlcon.State == ConnectionState.Closed)
                    sqlcon.Open();
                
            }
            catch (Exception e)
            {
                MyLog.Error("Error!", e);               
            }

            try
            {
                using (sqlcon)
                {
                    using (SqlCommand cmd = sqlcon.CreateCommand())
                    {
                        bool exists = CheckExistingTable(cmd, tableName);                        
                        //DataTable dtColumns = ReadDatabase(cmd, tableName);
                        if (exists)
                        {
                            if (IsTruncateTable)
                            {
                                TruncateTable(cmd, tableName);
                            }
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
                            MyLog.Error("Error!", e);
                            MyLog.Debug("Failure! Please read an error on 'logfile' folder");
                            
                        }
                    }
                }               
            }
            catch (SqlException e)
            {
                MyLog.Error("Error!", e);
            }

        }
        
  
    }
}
