using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExFormOfficeAddInDAL
{
    public class CommonSql
    {
        private SqlConnection objConn = null;

        public static string GetConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["dbConnectionString"].ConnectionString;
        }

        public SqlConnection GetConnection()
        {
            if (objConn != null && objConn.State == ConnectionState.Open)
            {
                return objConn;
            }

            else if (objConn != null && objConn.State == ConnectionState.Closed && !string.IsNullOrEmpty(objConn.ConnectionString))
            {
                objConn.Open();
                return objConn;
            }
            else
            {
                objConn = new SqlConnection();
                objConn.ConnectionString = GetConnectionString();
                objConn.Open();
                return objConn;
            }
        }

        public void CloseConnection()
        {
            if (objConn != null)
            {
                if (objConn.State == ConnectionState.Open)
                {
                    objConn.Close();
                }
            }
        }
    }
}
