using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace techlink_workspace.Data
{
    public class DatabaseUtils
    {
        public static SqlConnection GetDBConnection() //Main server của software
        {
            string datasource = "172.16.0.12";
            string database = "TLWA";
            string username = "ERPUSER";
            string password = "12345";
            return DatabaseSQLServerUtils.GetDBConnection(datasource, database, username, password);
        }
    }
}
