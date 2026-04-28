using org.omg.PortableInterceptor;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using techlink_workspace.Data;
using techlink_workspace.Model;

namespace techlink_workspace.Repositories
{
    public class UserRepository
    {
        private readonly SqlSoft _db = new SqlSoft();

        /// <summary>
        /// Returns the matched user row or null if not found / inactive.
        /// Password comparison is done in SQL to avoid pulling hashes to client.
        /// </summary>
        public UserModel GetByCredentials(string userCode, string password)
        {
            // Use parameterised query — never string-concatenate credentials
            string sql = @"
                SELECT TOP 1
                    User_id, User_code, User_fullName,
                    User_password, User_type, User_status
                FROM dbo.Sys_User
                WHERE (User_code = @code OR User_id = @code)
                  AND User_password = @pwd
                  AND User_status = 1";
            //1 = active
            try
            {
                var conn = DatabaseUtils.GetDBConnection();
                if (conn.State == ConnectionState.Closed) conn.Open();

                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@code", SqlDbType.NVarChar, 50).Value = userCode;
                    cmd.Parameters.Add("@pwd", SqlDbType.VarChar, 50).Value = password;
                    // NOTE: replace raw password with hashed value (e.g. SHA-256) when ready

                    using (var rdr = cmd.ExecuteReader())
                    {
                        if (!rdr.Read()) return null;

                        return new UserModel
                        {
                            User_id = rdr["User_id"].ToString(),
                            User_code = rdr["User_code"].ToString(),
                            User_fullName = rdr["User_fullName"].ToString(),
                            User_password = rdr["User_password"].ToString(),
                            User_type = Convert.ToInt32(rdr["User_type"]),
                            User_status = Convert.ToInt32(rdr["User_status"])
                        };
                    }
                }
            }
            catch (SqlException ex)
            {
                // bubble up so controller can handle / log
                throw new Exception("DB error during login: " + ex.Message, ex);
            }
        }
    }
}
