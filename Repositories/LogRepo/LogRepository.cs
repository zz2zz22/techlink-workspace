using System;
using System.Data;
using System.Data.SqlClient;
using techlink_workspace.Controller.Logic.IDGenerate;
using techlink_workspace.Data;

namespace techlink_workspace.Repositories.LogRepo
{
    public class LogRepository
    {
        /// <summary>
        /// Inserts one row into dbo.Sys_LogOperation.
        /// Call this from the Controller only — never from the View.
        /// </summary>
        public void WriteLog(
            string employeeId,
            string function,
            string module,
            int result,
            string variable = "",
            string oldData = "",
            string updateData = "",
            string sqlText = "")
        {
            string insertSql = @"
                INSERT INTO dbo.Sys_LogOperation
                    (Log_Id, Log_EmployeeId, Log_WriteDate,
                     Log_Function, Log_Module, Log_Result,
                     Log_Variable, Log_OldData, Log_UpdateData, Log_SQL)
                VALUES
                    (@id, @emp, @date,
                     @func, @mod, @res,
                     @var, @old, @upd, @sql)";

            try
            {
                // Use a brand-new connection string directly — avoids sharing
                // state with any connection already open in UserRepository
                string connStr = DatabaseUtils.GetDBConnection().ConnectionString;
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(insertSql, conn))
                    {
                        cmd.Parameters.Add("@id", SqlDbType.NVarChar, 50).Value = Guid.NewGuid().ToString();
                        cmd.Parameters.Add("@emp", SqlDbType.NVarChar, 50).Value = employeeId ?? "";
                        cmd.Parameters.Add("@date", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.Parameters.Add("@func", SqlDbType.NVarChar, 50).Value = function ?? "";
                        cmd.Parameters.Add("@mod", SqlDbType.NVarChar, 50).Value = module ?? "";
                        cmd.Parameters.Add("@res", SqlDbType.Int).Value = result;
                        cmd.Parameters.Add("@var", SqlDbType.NVarChar, 50).Value = variable ?? "";
                        cmd.Parameters.Add("@old", SqlDbType.NVarChar).Value = oldData ?? (object)DBNull.Value;
                        cmd.Parameters.Add("@upd", SqlDbType.NVarChar).Value = updateData ?? (object)DBNull.Value;
                        cmd.Parameters.Add("@sql", SqlDbType.NVarChar).Value = sqlText ?? (object)DBNull.Value;

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[LogRepository] WriteLog failed: " + ex.Message);
            }
        }
    }
}
