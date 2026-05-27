using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;
using techlink_workspace.Data;
using techlink_workspace.Model;
using techlink_workspace.Repositories.LogRepo;

namespace techlink_workspace.Repositories.FreightRepo
{
    public class ForwarderQuotationRepository
    {
        private readonly LogRepository _log = new LogRepository();

        // ── Helpers ──────────────────────────────────────────────────────
        private SqlConnection OpenConn()
        {
            var c = DatabaseUtils.GetDBConnection();
            if (c.State == ConnectionState.Closed) c.Open();
            return c;
        }

        // ── READ ─────────────────────────────────────────────────────────
        public DataTable GetAll()
        {
            const string sql = @"
                SELECT Forwarder_ID, Forwarder_name, Forwarder_portDelivery,
                       Forwarder_term, Forwarder_container, Forwarder_commodity,
                       Forwarder_hsCode, Forwarder_carrier,
                       Forwarder_total, Forwarder_of, Forwarder_localPol,
                       Forwarder_destCharge, Forwarder_delivery, Forwarder_otherCharge,
                       Forwarder_remark, Forwarder_volumn, Forwarder_validDate,
                       create_date, create_by, update_date, update_by
                FROM dbo.Logistic_ForwarderQuotation
                ORDER BY create_date DESC";
            var dt = new DataTable();
            using (var conn = OpenConn())
            using (var da = new SqlDataAdapter(sql, conn))
                da.Fill(dt);
            return dt;
        }

        // ── INSERT (single row, called from Excel upload) ─────────────────
        public void Insert(ForwarderQuotationModel m, string employeeCode)
        {
            const string sql = @"
                INSERT INTO dbo.Logistic_ForwarderQuotation
                    (Forwarder_ID, Forwarder_name, Forwarder_portDelivery,
                     Forwarder_term, Forwarder_container, Forwarder_commodity,
                     Forwarder_hsCode, Forwarder_carrier,
                     Forwarder_total, Forwarder_of, Forwarder_localPol,
                     Forwarder_destCharge, Forwarder_delivery, Forwarder_otherCharge,
                     Forwarder_remark, Forwarder_volumn, Forwarder_validDate,
                     create_date, create_by, update_date, update_by)
                VALUES
                    (@id,@name,@port,@term,@cont,@comm,@hs,@carrier,
                     @total,@of,@lpol,@dest,@del,@other,
                     @remark,@vol,@valid,
                     @cdate,@cby,@udate,@uby)";

            using (var conn = OpenConn())
            using (var cmd = new SqlCommand(sql, conn))
            {
                BindParams(cmd, m, employeeCode, isInsert: true);
                cmd.ExecuteNonQuery();
            }

            _log.WriteLog(employeeCode, "INSERT", "FREIGHT",
                result: 1,
                variable: m.Forwarder_ID,
                updateData: JsonConvert.SerializeObject(m),
                sqlText: sql);
        }

        // ── UPDATE ───────────────────────────────────────────────────────
        public void Update(ForwarderQuotationModel m, string employeeCode)
        {
            // Snapshot old data first
            var old = GetById(m.Forwarder_ID);

            const string sql = @"
                UPDATE dbo.Logistic_ForwarderQuotation SET
                    Forwarder_name=@name, Forwarder_portDelivery=@port,
                    Forwarder_term=@term, Forwarder_container=@cont,
                    Forwarder_commodity=@comm, Forwarder_hsCode=@hs,
                    Forwarder_carrier=@carrier,
                    Forwarder_total=@total, Forwarder_of=@of,
                    Forwarder_localPol=@lpol, Forwarder_destCharge=@dest,
                    Forwarder_delivery=@del, Forwarder_otherCharge=@other,
                    Forwarder_remark=@remark, Forwarder_volumn=@vol,
                    Forwarder_validDate=@valid,
                    update_date=@udate, update_by=@uby
                WHERE Forwarder_ID=@id";

            using (var conn = OpenConn())
            using (var cmd = new SqlCommand(sql, conn))
            {
                BindParams(cmd, m, employeeCode, isInsert: false);
                cmd.ExecuteNonQuery();
            }

            _log.WriteLog(employeeCode, "UPDATE", "FREIGHT",
                result: 1,
                variable: m.Forwarder_ID,
                oldData: old != null ? JsonConvert.SerializeObject(old) : "",
                updateData: JsonConvert.SerializeObject(m),
                sqlText: sql);
        }

        // ── DELETE ───────────────────────────────────────────────────────
        public void Delete(string forwarderId, string employeeCode)
        {
            var old = GetById(forwarderId);
            const string sql = "DELETE FROM dbo.Logistic_ForwarderQuotation WHERE Forwarder_ID=@id";

            using (var conn = OpenConn())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar, 50).Value = forwarderId;
                cmd.ExecuteNonQuery();
            }

            _log.WriteLog(employeeCode, "DELETE", "FREIGHT",
                result: 1,
                variable: forwarderId,
                oldData: old != null ? JsonConvert.SerializeObject(old) : "",
                sqlText: sql);
        }

        // ── GET BY ID ────────────────────────────────────────────────────
        public ForwarderQuotationModel GetById(string id)
        {
            const string fetchById = @"
                SELECT * FROM dbo.Logistic_ForwarderQuotation WHERE Forwarder_ID=@id";
            using (var conn = OpenConn())
            using (var cmd = new SqlCommand(fetchById, conn))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar, 50).Value = id;
                using (var rdr = cmd.ExecuteReader())
                {
                    if (!rdr.Read()) return null;
                    return MapRow(rdr);
                }
            }
        }

        // ── VERSION HISTORY (from Sys_LogOperation) ──────────────────────
        /// <summary>
        /// Returns log entries for FREIGHT module ordered newest first.
        /// When forwarderId is supplied, matches on Log_Variable (the ID stored at write time)
        /// OR on the Forwarder_ID embedded inside Log_OldData / Log_UpdateData JSON,
        /// so history survives even if Log_Variable was accidentally set to something else.
        /// </summary>
        public DataTable GetVersionHistory(string forwarderId = null)
        {
            string sql;
            if (string.IsNullOrEmpty(forwarderId))
            {
                sql = @"
                    SELECT Log_Id, Log_EmployeeId, Log_WriteDate,
                           Log_Function, Log_Result,
                           Log_Variable, Log_OldData, Log_UpdateData
                    FROM dbo.Sys_LogOperation
                    WHERE Log_Module = 'FREIGHT'
                    ORDER BY Log_WriteDate DESC";
            }
            else
            {
                // Match on Log_Variable directly, OR on the JSON payload containing the ID.
                // LIKE search on JSON is a safe fallback for SQL Server 2008+.
                sql = @"
                    SELECT Log_Id, Log_EmployeeId, Log_WriteDate,
                           Log_Function, Log_Result,
                           Log_Variable, Log_OldData, Log_UpdateData
                    FROM dbo.Sys_LogOperation
                    WHERE Log_Module = 'FREIGHT'
                      AND (
                            Log_Variable    = @var
                         OR Log_OldData    LIKE @varLike
                         OR Log_UpdateData LIKE @varLike
                      )
                    ORDER BY Log_WriteDate DESC";
            }

            var dt = new DataTable();
            using (var conn = OpenConn())
            using (var cmd = new SqlCommand(sql, conn))
            {
                if (!string.IsNullOrEmpty(forwarderId))
                {
                    cmd.Parameters.Add("@var", SqlDbType.NVarChar, 50).Value = forwarderId;
                    cmd.Parameters.Add("@varLike", SqlDbType.NVarChar).Value = $"%{forwarderId}%";
                }
                using (var da = new SqlDataAdapter(cmd))
                    da.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// Restores a row to the state stored in Log_OldData of the given log entry.
        /// </summary>
        public void Rollback(string logId, string employeeCode)
        {
            // Fetch the log entry
            const string fetchSql = @"
                SELECT Log_OldData, Log_Variable FROM dbo.Sys_LogOperation
                WHERE Log_Id = @lid";

            string oldJson, forwarderId;
            using (var conn = OpenConn())
            using (var cmd = new SqlCommand(fetchSql, conn))
            {
                cmd.Parameters.Add("@lid", SqlDbType.NVarChar, 50).Value = logId;
                using (var rdr = cmd.ExecuteReader())
                {
                    if (!rdr.Read()) throw new Exception("Log entry not found.");
                    oldJson = rdr["Log_OldData"]?.ToString();
                    forwarderId = rdr["Log_Variable"]?.ToString();
                }
            }

            if (string.IsNullOrWhiteSpace(oldJson))
                throw new Exception("No previous snapshot available for this log entry.");

            var restored = JsonConvert.DeserializeObject<ForwarderQuotationModel>(oldJson);
            if (restored == null) throw new Exception("Failed to deserialize snapshot.");

            // Upsert: delete then insert to guarantee clean state
            Delete(forwarderId, employeeCode);
            Insert(restored, employeeCode);

            _log.WriteLog(employeeCode, "ROLLBACK", "FREIGHT",
                result: 1,
                variable: forwarderId,
                updateData: $"Rolled back to log entry {logId}");
        }

        // ── Param binder ────────────────────────────────────────────────
        private static void BindParams(SqlCommand cmd, ForwarderQuotationModel m,
                                       string employeeCode, bool isInsert)
        {
            var now = DateTime.Now;
            cmd.Parameters.Add("@id", SqlDbType.NVarChar, 50).Value = m.Forwarder_ID ?? Guid.NewGuid().ToString();
            cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = m.Forwarder_name ?? (object)DBNull.Value;
            cmd.Parameters.Add("@port", SqlDbType.NVarChar, 50).Value = m.Forwarder_portDelivery ?? (object)DBNull.Value;
            cmd.Parameters.Add("@term", SqlDbType.NVarChar, 50).Value = m.Forwarder_term ?? (object)DBNull.Value;
            cmd.Parameters.Add("@cont", SqlDbType.NVarChar, 50).Value = m.Forwarder_container ?? (object)DBNull.Value;
            cmd.Parameters.Add("@comm", SqlDbType.NVarChar).Value = m.Forwarder_commodity ?? (object)DBNull.Value;
            cmd.Parameters.Add("@hs", SqlDbType.NVarChar, 50).Value = m.Forwarder_hsCode ?? (object)DBNull.Value;
            cmd.Parameters.Add("@carrier", SqlDbType.NVarChar, 50).Value = m.Forwarder_carrier ?? (object)DBNull.Value;
            cmd.Parameters.Add("@total", SqlDbType.Float).Value = (object)m.Forwarder_total ?? DBNull.Value;
            cmd.Parameters.Add("@of", SqlDbType.Float).Value = (object)m.Forwarder_of ?? DBNull.Value;
            cmd.Parameters.Add("@lpol", SqlDbType.Float).Value = (object)m.Forwarder_localPol ?? DBNull.Value;
            cmd.Parameters.Add("@dest", SqlDbType.Float).Value = (object)m.Forwarder_destCharge ?? DBNull.Value;
            cmd.Parameters.Add("@del", SqlDbType.Float).Value = (object)m.Forwarder_delivery ?? DBNull.Value;
            cmd.Parameters.Add("@other", SqlDbType.Float).Value = (object)m.Forwarder_otherCharge ?? DBNull.Value;
            cmd.Parameters.Add("@remark", SqlDbType.NVarChar).Value = m.Forwarder_remark ?? (object)DBNull.Value;
            cmd.Parameters.Add("@vol", SqlDbType.Int).Value = (object)m.Forwarder_volumn ?? DBNull.Value;
            cmd.Parameters.Add("@valid", SqlDbType.NVarChar).Value = m.Forwarder_validDate ?? (object)DBNull.Value;
            cmd.Parameters.Add("@udate", SqlDbType.DateTime).Value = now;
            cmd.Parameters.Add("@uby", SqlDbType.NVarChar, 50).Value = employeeCode;
            if (isInsert)
            {
                cmd.Parameters.Add("@cdate", SqlDbType.DateTime).Value = m.create_date ?? now;
                cmd.Parameters.Add("@cby", SqlDbType.NVarChar, 50).Value = employeeCode;
            }
        }

        // ── Row mapper ───────────────────────────────────────────────────
        private static ForwarderQuotationModel MapRow(SqlDataReader r)
        {
            decimal? N(string col) => r[col] == DBNull.Value ? (decimal?)null : Convert.ToDecimal(r[col]);
            int? I(string col) => r[col] == DBNull.Value ? (int?)null : Convert.ToInt32(r[col]);
            DateTime? D(string col) => r[col] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[col]);
            string S(string col) => r[col] == DBNull.Value ? null : r[col].ToString();

            return new ForwarderQuotationModel
            {
                Forwarder_ID = S("Forwarder_ID"),
                Forwarder_name = S("Forwarder_name"),
                Forwarder_portDelivery = S("Forwarder_portDelivery"),
                Forwarder_term = S("Forwarder_term"),
                Forwarder_container = S("Forwarder_container"),
                Forwarder_commodity = S("Forwarder_commodity"),
                Forwarder_hsCode = S("Forwarder_hsCode"),
                Forwarder_carrier = S("Forwarder_carrier"),
                Forwarder_total = N("Forwarder_total"),
                Forwarder_of = N("Forwarder_of"),
                Forwarder_localPol = N("Forwarder_localPol"),
                Forwarder_destCharge = N("Forwarder_destCharge"),
                Forwarder_delivery = N("Forwarder_delivery"),
                Forwarder_otherCharge = N("Forwarder_otherCharge"),
                Forwarder_remark = S("Forwarder_remark"),
                Forwarder_volumn = I("Forwarder_volumn"),
                Forwarder_validDate = S("Forwarder_validDate"),
                create_date = D("create_date"),
                create_by = S("create_by"),
                update_date = D("update_date"),
                update_by = S("update_by"),
            };
        }
    }
}