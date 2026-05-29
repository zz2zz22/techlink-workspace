using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using techlink_workspace.Data;
using techlink_workspace.Model;

namespace techlink_workspace.Repositories.PICRepo
{
    public class PICRepository
    {
        private const string TABLE = "dbo.Logistic_PICManagement";

        private static SqlConnection Open()
        {
            var c = DatabaseUtils.GetDBConnection();
            if (c.State == ConnectionState.Closed) c.Open();
            return c;
        }

        // ── READ ─────────────────────────────────────────────────────────────
        public List<PICModel> GetAll()
        {
            var list = new List<PICModel>();
            string sql = $"SELECT * FROM {TABLE} ORDER BY PIC_Code";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            using (var r = cmd.ExecuteReader())
                while (r.Read()) list.Add(Map(r));
            return list;
        }

        public PICModel GetByCode(string code)
        {
            string sql = $"SELECT * FROM {TABLE} WHERE PIC_Code = @code";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@code", code ?? "");
                using (var r = cmd.ExecuteReader())
                    return r.Read() ? Map(r) : null;
            }
        }

        /// <summary>
        /// Loads all PICs (typically few rows) and does exact match on each
        /// semicolon-separated customer code.
        /// </summary>
        public string FindPICByCustomerCode(string customerCode)
        {
            if (string.IsNullOrWhiteSpace(customerCode)) return null;
            foreach (var pic in GetAll())
            {
                if (string.IsNullOrWhiteSpace(pic.PIC_CustomerName)) continue;
                foreach (var part in pic.PIC_CustomerName.Split(';'))
                    if (part.Trim().Equals(customerCode.Trim(), StringComparison.OrdinalIgnoreCase))
                        return pic.PIC_Code;
            }
            return null;
        }

        // ── UPSERT ───────────────────────────────────────────────────────────
        public void Upsert(PICModel m, string byUser)
        {
            if (GetByCode(m.PIC_Code) == null) Insert(m, byUser);
            else Update(m, byUser);
        }

        private void Insert(PICModel m, string byUser)
        {
            string sql = $@"
                INSERT INTO {TABLE} (PIC_Code, PIC_Name, PIC_CustomerName,
                                     createdate, createby, updatedate, updateby)
                VALUES (@code, @name, @cust, @cd, @cb, @ud, @ub)";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            { Bind(cmd, m, byUser, true); cmd.ExecuteNonQuery(); }
        }

        private void Update(PICModel m, string byUser)
        {
            string sql = $@"
                UPDATE {TABLE}
                SET PIC_Name=@name, PIC_CustomerName=@cust, updatedate=@ud, updateby=@ub
                WHERE PIC_Code=@code";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            { Bind(cmd, m, byUser, false); cmd.ExecuteNonQuery(); }
        }

        // ── DELETE ───────────────────────────────────────────────────────────
        public void Delete(string code)
        {
            string sql = $"DELETE FROM {TABLE} WHERE PIC_Code = @code";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            { cmd.Parameters.AddWithValue("@code", code ?? ""); cmd.ExecuteNonQuery(); }
        }

        // ── Helpers ──────────────────────────────────────────────────────────
        private static void Bind(SqlCommand cmd, PICModel m, string byUser, bool isInsert)
        {
            var now = DateTime.Now;
            object N(object v) => v ?? DBNull.Value;
            cmd.Parameters.AddWithValue("@code", N(m.PIC_Code));
            cmd.Parameters.AddWithValue("@name", N(m.PIC_Name));
            cmd.Parameters.AddWithValue("@cust", N(m.PIC_CustomerName));
            cmd.Parameters.AddWithValue("@ud", now);
            cmd.Parameters.AddWithValue("@ub", byUser ?? (object)DBNull.Value);
            if (isInsert)
            {
                cmd.Parameters.AddWithValue("@cd", now);
                cmd.Parameters.AddWithValue("@cb", byUser ?? (object)DBNull.Value);
            }
        }

        private static PICModel Map(SqlDataReader r)
        {
            string S(string c) => r[c] == DBNull.Value ? null : r[c].ToString();
            return new PICModel
            {
                PIC_Code = S("PIC_Code"),
                PIC_Name = S("PIC_Name"),
                PIC_CustomerName = S("PIC_CustomerName"),
                createby = S("createby"),
                updateby = S("updateby"),
            };
        }
    }
}