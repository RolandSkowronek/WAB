using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace wab2018
{
    internal class cm
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public string con_str = ConfigurationManager.ConnectionStrings["wap"].ConnectionString;

        public string podajMiesiac(int numerMiesiaca)
        {
            switch (numerMiesiaca)
            {
                case 1: return "styczeń";
                case 2: return "luty";
                case 3: return "marzec";
                case 4: return "kwieceń";
                case 5: return "maj";
                case 6: return "czerwiec";
                case 7: return "lipiec";
                case 8: return "sierpień";
                case 9: return "wrzesień";
                case 10: return "październik";
                case 11: return "listopad";
                case 12: return "grudzień";
                default:
                    return "";
            }
        }

        public DataTable makeParameterTable()
        {
            DataTable parameters = new DataTable();
            parameters.Columns.Add("name", typeof(String));
            parameters.Columns.Add("value", typeof(String));
            return parameters;
        }

        public void runQuerry(string kwerenda, DataTable parameters)
        {
            log.Debug("start runQuerry");
            var conn = new SqlConnection(con_str);
            using (SqlCommand sqlCmd = new SqlCommand(kwerenda, conn))
            {
                try
                {
                    log.Debug("try to open DB connection");
                    conn.Open();
                    if (parameters != null)
                    {
                        foreach (DataRow row in parameters.Rows)
                        {
                            sqlCmd.Parameters.AddWithValue(row[0].ToString().Trim(), row[1].ToString().Trim());
                            log.Debug("add parameter -0 name: " + row[0].ToString().Trim() + " value: " + row[1].ToString().Trim());
                        }
                    }
                    log.Debug("executing querry");
                    sqlCmd.ExecuteScalar();
                    log.Debug("querry executed");
                    conn.Close();
                }
                catch (Exception ex)
                {
                    log.Error("runQuerry Error: " + ex.Message);
                    conn.Close();
                }
            } // end of using
        }

        public void runQuerry(string kwerenda, string connStr, DataTable parameters)
        {
            log.Debug("start runQuerry");
            var conn = new SqlConnection(connStr);
            using (SqlCommand sqlCmd = new SqlCommand(kwerenda, conn))
            {
                try
                {
                    log.Debug("try to open DB connection");
                    conn.Open();
                    if (parameters != null)
                    {
                        foreach (DataRow row in parameters.Rows)
                        {
                            sqlCmd.Parameters.AddWithValue(row[0].ToString().Trim(), row[1].ToString().Trim());
                            log.Debug("add parameter -0 name: " + row[0].ToString().Trim() + " value: " + row[1].ToString().Trim());
                        }
                    }
                    log.Debug("executing querry");
                    sqlCmd.ExecuteScalar();
                    log.Debug("querry executed");
                    conn.Close();
                }
                catch (Exception ex)
                {
                    log.Error("runQuerry Error: " + ex.Message);
                    conn.Close();
                }
            } // end of using
        }

        public string runQuerryWithResult(string kwerenda, string connStr, DataTable parameters)
        {
            log.Debug("start runQuerryWithResult");
            string result = string.Empty;
            var conn = new SqlConnection(connStr);
            using (SqlCommand sqlCmd = new SqlCommand(kwerenda, conn))
            {
                try
                {
                    log.Debug("try to open DB connection");
                    conn.Open();
                    if (parameters != null)
                    {
                        foreach (DataRow row in parameters.Rows)
                        {
                            sqlCmd.Parameters.AddWithValue(row[0].ToString().Trim(), row[1].ToString().Trim());
                            log.Debug("add parameter -0 name: " + row[0].ToString().Trim() + " value: " + row[1].ToString().Trim());
                        }
                    }
                    log.Debug("executing querry");
                    result = sqlCmd.ExecuteScalar().ToString().Trim();
                    log.Debug("querry executed with result");
                    conn.Close();
                }
                catch (Exception ex)
                {
                    log.Error("runQuerry Error: " + ex.Message);
                    conn.Close();
                }
                return result;
            } // end of using
        }

        public string runQuerryWithResult(string kwerenda, string connStr)
        {
            log.Debug("start runQuerryWithResult");
            string result = string.Empty;
            var conn = new SqlConnection(connStr);
            using (SqlCommand sqlCmd = new SqlCommand(kwerenda, conn))
            {
                try
                {
                    log.Debug("try to open DB connection");
                    conn.Open();
                    log.Debug("DB connection opened");
                    log.Debug("executing querry");
                    result = sqlCmd.ExecuteScalar().ToString().Trim();
                    log.Debug("querry executed with result");
                    conn.Close();
                }
                catch (Exception ex)
                {
                    log.Error("runQuerry Error: " + ex.Message);
                    conn.Close();
                }

                //    string.IsNullOrEmpty(result)  ? (log.Debug("return null value") ):  ( log.Debug("return value"));

                return result;
            } // end of using
        }

       /*

        public DataTable getDataTable(string Kwerenda, string connectionString)
        {
            
            SqlConnection conn = new SqlConnection(connectionString);
            DataSet lista = new DataSet();
            DataTable returntable = new DataTable();
            try
            {
                conn.Open();
                SqlDataAdapter daData = new SqlDataAdapter();
                daData.SelectCommand = new SqlCommand(Kwerenda, conn);
                daData.Fill(lista);
                conn.Close();
                returntable = lista.Tables[0];
            }
            catch (Exception ec)
            {
                conn.Close();
            }
            return returntable;
        }
        */
       
         public DataTable getDataTable(string kwerenda, string connStr, DataTable parameters)
         {
             log.Info("Start getDataTable");
             DataTable result = new DataTable();
             var conn = new SqlConnection(connStr);

             DataSet dsKwerendy = new DataSet();
             dsKwerendy = new DataSet();
             try
             {
                 log.Info("Open DB connection");
                 conn.Open();
                 log.Info("DB connection is open");
                 SqlDataAdapter daData = new SqlDataAdapter();
                 daData.SelectCommand = new SqlCommand(kwerenda, conn);
                 foreach (DataRow row in parameters.Rows)
                 {
                     daData.SelectCommand.Parameters.AddWithValue(row[0].ToString().Trim(), row[1].ToString().Trim());
                 }
                 log.Info("Executing querry");
                 daData.Fill(dsKwerendy);
                 log.Info("Querry is executed");

                 conn.Close();
                 log.Info("DB  is closed");

                 result = dsKwerendy.Tables[0];
             }
             catch (Exception ex)
             {
                 log.Error("Error : " + ex.Message);
                 conn.Close();
             }

             return result;
         } // end of getDataTable
        
        //==========================================================================================================================================
        //==========================================================================================================================================
        //==========================================================================================================================================
    }
}