using System;
using System.Data;
using System.Data.SqlClient;
using techlink_workspace.Data;
using techlink_workspace.Model;

namespace techlink_workspace.Repositories
{
    public class UserRepository
    {
        /// <summary>
        /// Authenticates and returns the user including the new permissionLevel.
        /// User_type is now nvarchar(50) storing a UserType ID string.
        /// </summary>
        public UserModel GetByCredentials(string userCode, string password)
        {
            const string sql = @"
                SELECT TOP 1
                    u.User_id, u.User_code, u.User_fullName,
                    u.User_password, u.User_type, u.User_status,
                    u.User_permissionLevel
                FROM dbo.Sys_User u
                WHERE (u.User_code = @code OR u.User_id = @code)
                  AND u.User_password = @pwd
                  AND u.User_status = 1";

            try
            {
                var conn = DatabaseUtils.GetDBConnection();
                if (conn.State == ConnectionState.Closed) conn.Open();

                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@code", SqlDbType.NVarChar, 50).Value = userCode;
                    cmd.Parameters.Add("@pwd", SqlDbType.VarChar, 50).Value = password;

                    using (var rdr = cmd.ExecuteReader())
                    {
                        if (!rdr.Read()) return null;

                        return new UserModel
                        {
                            User_id = rdr["User_id"].ToString(),
                            User_code = rdr["User_code"].ToString(),
                            User_fullName = rdr["User_fullName"].ToString(),
                            User_password = rdr["User_password"].ToString(),
                            User_type = rdr["User_type"].ToString(),   // nvarchar(50)
                            User_status = Convert.ToInt32(rdr["User_status"]),
                            User_permissionLevel = rdr["User_permissionLevel"] == DBNull.Value
                                                   ? (int?)null
                                                   : Convert.ToInt32(rdr["User_permissionLevel"])
                        };
                    }
                }
            }
            catch (SqlException ex)
            {
                throw new Exception("DB error during login: " + ex.Message, ex);
            }
        }
    }
}