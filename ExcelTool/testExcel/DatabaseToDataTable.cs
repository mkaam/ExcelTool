using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using System.IO;
namespace testExcel
{
    class DatabaseToDataTable
    {
        private static CLogger MyLog;
        public DatabaseToDataTable(CLogger logger)
        {
            MyLog = logger;
        }
        public DataTable ReadDatabase(SqlCommand cmd, string queryText)
        {
            var table = new DataTable();
            cmd.CommandText = queryText;
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

        public DataTable ExportToDatabase(Dictionary<string, object> dict, IEnumerable<string> QueryParameter = null)
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
            var dtColumns = new DataTable();
            SqlConnection sqlcon = new SqlConnection("server=" + @"" + serverName + ";" +
                                       "Trusted_Connection=yes;" +
                                       "database=" + dbName + "; Integrated Security=true; " +
                                       "connection timeout=0");
            try
            {
                if (sqlcon.State == ConnectionState.Closed)
                    sqlcon.Open();                
            }
            catch (Exception e)
            {
                MyLog.Error("ExportToDatabase Error!", e);                
            }

            string textFile = ExportQueryFile;

            string queryText = "";
            if (File.Exists(textFile))
            {
                // Read entire text file content in one string    
                queryText = File.ReadAllText(textFile);

                // added by aam 20210907 support for queryparameter, find and replace some text in query file
                if (QueryParameter != null)
                {
                    foreach (string qtr in QueryParameter)
                    {
                        //skip if wrong format 
                        if (qtr.Contains("="))
                        {
                            string lefttext = qtr.Split('=')[0].ToString();
                            string righttext = qtr.Split('=')[1].ToString();
                            if (lefttext.Length > 0 && righttext.Length > 0)
                                queryText = queryText.Replace(lefttext, righttext);
                        }
                    }
                }
                
            }
            else
            {
                MyLog.Warn("FILE DOESN'T EXIST");
            }

            try
            {
                using (sqlcon)
                {
                    using (SqlCommand cmd = sqlcon.CreateCommand())
                    {
                        bool exists = CheckExistingTable(cmd, tableName);                        
                    
                            dtColumns = ReadDatabase(cmd, queryText);
                                                
                            cmd.ExecuteNonQuery();
                            sqlcon.Close();
                    }
                }

            }
            catch (SqlException e)
            {
                MyLog.Error("ExportToDatabase Error!", e);                
            }
            return dtColumns;
        }

        public static implicit operator DatabaseToDataTable(GetMetaData v)
        {
            throw new NotImplementedException();
        }
    }
}
