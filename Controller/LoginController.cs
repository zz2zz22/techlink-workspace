using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using techlink_workspace.Model;
using techlink_workspace.Repositories;
using techlink_workspace.Repositories.LogRepo;

namespace techlink_workspace.Controller
{
    public class LoginController
    {
        private readonly UserRepository _userRepo = new UserRepository();
        private readonly LogRepository _logRepo = new LogRepository();

        public enum LoginResult { Success, InvalidCredentials, DbError, Inactive }

        /// <summary>
        /// Returns a LoginResult and the matched user (null on failure).
        /// The View calls this and reacts to the enum — zero UI logic here.
        /// </summary>
        public LoginResult Authenticate(string userCode, string password, out UserModel user)
        {
            user = null;
            try
            {
                UserModel found = _userRepo.GetByCredentials(userCode, password);

                if (found == null)
                {
                    // Log failed attempt — use the typed code as employee id placeholder
                    _logRepo.WriteLog(
                        employeeId: userCode,
                        function: "LOGIN",
                        module: "AUTH",
                        result: 0,
                        variable: "WRONG_CREDENTIALS",
                        updateData: $"Attempted login at {DateTime.Now:yyyy-MM-dd HH:mm:ss}"
                    );
                    return LoginResult.InvalidCredentials;
                }

                // Success
                user = found;
                _logRepo.WriteLog(
                    employeeId: found.User_id,
                    function: "LOGIN",
                    module: "AUTH",
                    result: 1,
                    variable: "OK",
                    updateData: $"Login success at {DateTime.Now:yyyy-MM-dd HH:mm:ss}"
                );

                // Update last login timestamp
                UpdateLastLogin(found.User_id);

                return LoginResult.Success;
            }
            catch (Exception ex)
            {
                _logRepo.WriteLog(
                    employeeId: userCode,
                    function: "LOGIN",
                    module: "AUTH",
                    result: 0,
                    variable: "EXCEPTION",
                    updateData: ex.Message
                );
                return LoginResult.DbError;
            }
        }

        private void UpdateLastLogin(string userId)
        {
            string sql = $@"
                UPDATE dbo.Sys_User
                SET User_lastlogin = GETDATE()
                WHERE User_id = '{userId}'";
            // Safe here because userId comes from DB, not raw user input
            new SqlSoft().sqlExecuteNonQuery(sql, null, null);
        }
    }
}
