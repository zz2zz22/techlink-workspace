using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace techlink_workspace.Data
{
    public class DatabaseSQLServerUtils
    {
        public static SqlConnection GetDBConnection(string datasource, string database, string username, string password)
        {
            string connString = @"Data Source=" + datasource + ";Initial Catalog=" + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;
            SqlConnection conn = new SqlConnection(connString);
            return conn;
        }
        //public static MySqlConnection GetMesDBConnection(string host, string user, string password, string database)
        //{
        //    string connectionString = String.Format("host={0};user={1};password={2}; database={3};", host, user, password, database);
        //    MySqlConnection con = new MySqlConnection(connectionString);
        //    return con;
        //}
    }
}
