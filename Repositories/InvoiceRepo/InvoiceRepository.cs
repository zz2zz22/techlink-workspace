using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;
using techlink_workspace.Data;
using techlink_workspace.Model;
using techlink_workspace.Repositories.LogRepo;

namespace techlink_workspace.Repositories.InvoiceRepo
{
    public class InvoiceRepository
    {
        private readonly LogRepository _log = new LogRepository();
        private const string TABLE = "dbo.Logistic_InvoiceData";

        // ── Connection helper ────────────────────────────────────────────
        private static SqlConnection Open()
        {
            var c = DatabaseUtils.GetDBConnection();
            if (c.State == ConnectionState.Closed) c.Open();
            return c;
        }

        // ════════════════════════════════════════════════════════════════
        // READ
        // ════════════════════════════════════════════════════════════════
        public DataTable GetAll()
        {
            string sql = $"SELECT * FROM {TABLE} ORDER BY createdate DESC";
            var dt = new DataTable();
            using (var conn = Open())
            using (var da = new SqlDataAdapter(sql, conn))
                da.Fill(dt);
            return dt;
        }

        public InvoiceModel GetById(string id)
        {
            string sql = $"SELECT * FROM {TABLE} WHERE Invoice_Id = @id";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar, 50).Value = id;
                using (var r = cmd.ExecuteReader())
                    return r.Read() ? Map(r) : null;
            }
        }

        public InvoiceModel GetByInvoiceNo(string invoiceNo)
        {
            string sql = $"SELECT * FROM {TABLE} WHERE Invoice_no = @no";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@no", SqlDbType.NVarChar, 50).Value = invoiceNo;
                using (var r = cmd.ExecuteReader())
                    return r.Read() ? Map(r) : null;
            }
        }

        // ════════════════════════════════════════════════════════════════
        // INSERT
        // ════════════════════════════════════════════════════════════════
        public void Insert(InvoiceModel m, string byUser)
        {
            using (var conn = Open())
            using (var cmd = new SqlCommand(InsertSql(), conn))
            {
                Bind(cmd, m, byUser, isInsert: true);
                cmd.ExecuteNonQuery();
            }
            _log.WriteLog(byUser, "INSERT", "INVOICE", 1,
                variable: m.Invoice_Id,
                updateData: JsonConvert.SerializeObject(m));
        }

        // ════════════════════════════════════════════════════════════════
        // BULK INSERT (Excel import) — upsert handled in form layer
        // ════════════════════════════════════════════════════════════════
        public (int inserted, int failed) BulkInsert(
            IEnumerable<InvoiceModel> records, string byUser)
        {
            int ok = 0, fail = 0;
            string sql = InsertSql();
            var errors = new System.Text.StringBuilder();

            using (var conn = Open())
            {
                foreach (var m in records)
                {
                    if (string.IsNullOrEmpty(m.Invoice_Id))
                        m.Invoice_Id = Guid.NewGuid().ToString();

                    // Guard: Invoice_no must not be empty (NOT NULL in DB)
                    if (string.IsNullOrWhiteSpace(m.Invoice_no))
                    { fail++; continue; }

                    try
                    {
                        using (var cmd = new SqlCommand(sql, conn))
                        {
                            Bind(cmd, m, byUser, isInsert: true);
                            cmd.ExecuteNonQuery();
                        }
                        ok++;
                    }
                    catch (SqlException ex)
                    {
                        fail++;
                        // Capture first 5 unique errors for diagnosis
                        if (fail <= 5)
                            errors.AppendLine($"Row {m.Invoice_no}: {ex.Message}");
                    }
                }
            }

            _log.WriteLog(byUser, "BULK_INSERT", "INVOICE", 1,
                updateData: $"{ok} inserted, {fail} failed" +
                            (errors.Length > 0 ? "\n" + errors : ""));

            // Surface errors to caller if any occurred
            if (fail > 0 && errors.Length > 0)
                throw new Exception(
                    $"{ok} inserted, {fail} failed.\n\nFirst errors:\n{errors}");

            return (ok, fail);
        }

        // ════════════════════════════════════════════════════════════════
        // UPDATE
        // ════════════════════════════════════════════════════════════════
        public void Update(InvoiceModel m, string byUser)
        {
            var old = GetById(m.Invoice_Id);
            using (var conn = Open())
            using (var cmd = new SqlCommand(UpdateSql(), conn))
            {
                Bind(cmd, m, byUser, isInsert: false);
                cmd.ExecuteNonQuery();
            }
            _log.WriteLog(byUser, "UPDATE", "INVOICE", 1,
                variable: m.Invoice_Id,
                oldData: old != null ? JsonConvert.SerializeObject(old) : "",
                updateData: JsonConvert.SerializeObject(m));
        }

        // ════════════════════════════════════════════════════════════════
        // DELETE
        // ════════════════════════════════════════════════════════════════
        public void Delete(string id, string byUser)
        {
            var old = GetById(id);
            string sql = $"DELETE FROM {TABLE} WHERE Invoice_Id = @id";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar, 50).Value = id;
                cmd.ExecuteNonQuery();
            }
            _log.WriteLog(byUser, "DELETE", "INVOICE", 1,
                variable: id,
                oldData: old != null ? JsonConvert.SerializeObject(old) : "");
        }

        // ════════════════════════════════════════════════════════════════
        // VERSION HISTORY
        // ════════════════════════════════════════════════════════════════
        public DataTable GetHistory(string invoiceId = null)
        {
            string sql = @"
                SELECT Log_Id, Log_EmployeeId, Log_WriteDate, Log_Function,
                       Log_Result, Log_Variable, Log_OldData, Log_UpdateData
                FROM dbo.Sys_LogOperation
                WHERE Log_Module = 'INVOICE'";

            if (!string.IsNullOrEmpty(invoiceId))
                sql += " AND (Log_Variable = @var OR Log_OldData LIKE @like OR Log_UpdateData LIKE @like)";

            sql += " ORDER BY Log_WriteDate DESC";

            var dt = new DataTable();
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                if (!string.IsNullOrEmpty(invoiceId))
                {
                    cmd.Parameters.Add("@var", SqlDbType.NVarChar, 50).Value = invoiceId;
                    cmd.Parameters.Add("@like", SqlDbType.NVarChar).Value = $"%{invoiceId}%";
                }
                using (var da = new SqlDataAdapter(cmd))
                    da.Fill(dt);
            }
            return dt;
        }

        // ════════════════════════════════════════════════════════════════
        // ROLLBACK
        // ════════════════════════════════════════════════════════════════
        public void Rollback(string logId, string byUser)
        {
            string fetchSql = @"
                SELECT Log_OldData, Log_Variable
                FROM dbo.Sys_LogOperation
                WHERE Log_Id = @lid";

            string oldJson, id;
            using (var conn = Open())
            using (var cmd = new SqlCommand(fetchSql, conn))
            {
                cmd.Parameters.Add("@lid", SqlDbType.NVarChar, 50).Value = logId;
                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) throw new Exception("Log entry not found.");
                    oldJson = r["Log_OldData"]?.ToString();
                    id = r["Log_Variable"]?.ToString();
                }
            }

            if (string.IsNullOrWhiteSpace(oldJson))
                throw new Exception("No snapshot available for this entry.");

            var m = JsonConvert.DeserializeObject<InvoiceModel>(oldJson);
            if (m == null) throw new Exception("Failed to deserialize snapshot.");

            Delete(id, byUser);
            Insert(m, byUser);

            _log.WriteLog(byUser, "ROLLBACK", "INVOICE", 1,
                variable: id,
                updateData: $"Rolled back to log entry {logId}");
        }

        // ════════════════════════════════════════════════════════════════
        // SQL strings
        // ════════════════════════════════════════════════════════════════
        private static string InsertSql() => $@"
            INSERT INTO {TABLE} (
                Invoice_Id, Invoice_no, Invoice_erpID, Invoice_erpInvoiceNo,
                Invoice_shippingTerm, Invoice_paymentTerm, Invoice_employee,
                Invoice_logisticRemark, Invoice_confirmDate, Invoice_fwdName,
                Invoice_bookingNo, Invoice_contType, Invoice_vgmCO, Invoice_cyCO,
                Invoice_etd, Invoice_eta, Invoice_billType, Invoice_billNo,
                Invoice_co, Invoice_coNo,
                Invoice_OF, Invoice_deliveryCharges, Invoice_taxes,
                Invoice_otherDestCharges, Invoice_thc, Invoice_blFee,
                Invoice_seal, Invoice_telexRelease, Invoice_cfs, Invoice_vgmFee,
                Invoice_ensebsams, Invoice_other, Invoice_totalVND,
                Invoice_subTotalOcean, Invoice_coFee, Invoice_feeStatus,
                Invoice_redInvoiceNo, Invoice_redInvoiceDate,
                Invoice_redInvoiceRecvDate, Invoice_transferAccountantDate,
                Invoice_trucking, Invoice_infrastructureFee,
                Invoice_customerClearance, Invoice_customFee, Invoice_otherCustomFee,
                Invoice_subTotalVNDCustom, Invoice_subTotalUSDCustom,
                Invoice_grandTotalVND, Invoice_grandTotalUSD,
                Invoice_cdsNo, Invoice_cdsDate, Invoice_line, Invoice_customType,
                createdate, createby, updatedate, updateby
            ) VALUES (
                @id, @no, @erpid, @erpinvno,
                @shipterm, @payterm, @emp,
                @logremark, @confirmdate, @fwdname,
                @bookno, @conttype, @vgmco, @cyco,
                @etd, @eta, @billtype, @billno,
                @co, @cono,
                @OF, @delchg, @taxes,
                @otherdest, @thc, @blfee,
                @seal, @telex, @cfs, @vgm,
                @ens, @other, @totalvnd,
                @subtotalocean, @cofee, @feestatus,
                @redinvno, @redinvdate,
                @redinvrecv, @transferacct,
                @truck, @infra,
                @custclear, @custfee, @othercustfee,
                @subtotalvndcust, @subtotalusdcust,
                @grandvnd, @grandusd,
                @cdsno, @cdsdate, @line, @customtype,
                @cdate, @cby, @udate, @uby
            )";

        private static string UpdateSql() => $@"
            UPDATE {TABLE} SET
                Invoice_no              = @no,
                Invoice_erpID           = @erpid,
                Invoice_erpInvoiceNo    = @erpinvno,
                Invoice_shippingTerm    = @shipterm,
                Invoice_paymentTerm     = @payterm,
                Invoice_employee        = @emp,
                Invoice_logisticRemark  = @logremark,
                Invoice_confirmDate     = @confirmdate,
                Invoice_fwdName         = @fwdname,
                Invoice_bookingNo       = @bookno,
                Invoice_contType        = @conttype,
                Invoice_vgmCO           = @vgmco,
                Invoice_cyCO            = @cyco,
                Invoice_etd             = @etd,
                Invoice_eta             = @eta,
                Invoice_billType        = @billtype,
                Invoice_billNo          = @billno,
                Invoice_co              = @co,
                Invoice_coNo            = @cono,
                Invoice_OF              = @OF,
                Invoice_deliveryCharges = @delchg,
                Invoice_taxes           = @taxes,
                Invoice_otherDestCharges= @otherdest,
                Invoice_thc             = @thc,
                Invoice_blFee           = @blfee,
                Invoice_seal            = @seal,
                Invoice_telexRelease    = @telex,
                Invoice_cfs             = @cfs,
                Invoice_vgmFee          = @vgm,
                Invoice_ensebsams       = @ens,
                Invoice_other           = @other,
                Invoice_totalVND        = @totalvnd,
                Invoice_subTotalOcean   = @subtotalocean,
                Invoice_coFee           = @cofee,
                Invoice_feeStatus       = @feestatus,
                Invoice_redInvoiceNo    = @redinvno,
                Invoice_redInvoiceDate  = @redinvdate,
                Invoice_redInvoiceRecvDate    = @redinvrecv,
                Invoice_transferAccountantDate = @transferacct,
                Invoice_trucking             = @truck,
                Invoice_infrastructureFee    = @infra,
                Invoice_customerClearance    = @custclear,
                Invoice_customFee            = @custfee,
                Invoice_otherCustomFee       = @othercustfee,
                Invoice_subTotalVNDCustom    = @subtotalvndcust,
                Invoice_subTotalUSDCustom    = @subtotalusdcust,
                Invoice_grandTotalVND        = @grandvnd,
                Invoice_grandTotalUSD        = @grandusd,
                Invoice_cdsNo           = @cdsno,
                Invoice_cdsDate         = @cdsdate,
                Invoice_line            = @line,
                Invoice_customType      = @customtype,
                updatedate              = @udate,
                updateby                = @uby
            WHERE Invoice_Id = @id";

        // ════════════════════════════════════════════════════════════════
        // Parameter binder — shared by Insert and Update
        // ════════════════════════════════════════════════════════════════
        private static void Bind(SqlCommand cmd, InvoiceModel m,
                                  string byUser, bool isInsert)
        {
            var now = DateTime.Now;

            // Use AddWithValue for all params — avoids size/type mismatch issues.
            // DBNull helper for nullable types.
            object N(object v) => v ?? (object)DBNull.Value;

            cmd.Parameters.AddWithValue("@id", N(m.Invoice_Id));
            cmd.Parameters.AddWithValue("@no", N(m.Invoice_no));
            cmd.Parameters.AddWithValue("@erpid", N(m.Invoice_erpID));
            cmd.Parameters.AddWithValue("@erpinvno", N(m.Invoice_erpInvoiceNo));
            cmd.Parameters.AddWithValue("@shipterm", N(m.Invoice_shippingTerm));
            cmd.Parameters.AddWithValue("@payterm", N(m.Invoice_paymentTerm));
            cmd.Parameters.AddWithValue("@emp", N(m.Invoice_employee));
            cmd.Parameters.AddWithValue("@logremark", N(m.Invoice_logisticRemark));
            cmd.Parameters.AddWithValue("@confirmdate", N(m.Invoice_confirmDate));
            cmd.Parameters.AddWithValue("@fwdname", N(m.Invoice_fwdName));
            cmd.Parameters.AddWithValue("@bookno", N(m.Invoice_bookingNo));
            cmd.Parameters.AddWithValue("@conttype", N(m.Invoice_contType));
            cmd.Parameters.AddWithValue("@vgmco", N(m.Invoice_vgmCO));
            cmd.Parameters.AddWithValue("@cyco", N(m.Invoice_cyCO));
            cmd.Parameters.AddWithValue("@etd", N(m.Invoice_etd));
            cmd.Parameters.AddWithValue("@eta", N(m.Invoice_eta));
            cmd.Parameters.AddWithValue("@billtype", N(m.Invoice_billType));
            cmd.Parameters.AddWithValue("@billno", N(m.Invoice_billNo));
            cmd.Parameters.AddWithValue("@co", N(m.Invoice_co));
            cmd.Parameters.AddWithValue("@cono", N(m.Invoice_coNo));
            cmd.Parameters.AddWithValue("@OF", N(m.Invoice_OF));
            cmd.Parameters.AddWithValue("@delchg", N(m.Invoice_deliveryCharges));
            cmd.Parameters.AddWithValue("@taxes", N(m.Invoice_taxes));
            cmd.Parameters.AddWithValue("@otherdest", N(m.Invoice_otherDestCharges));
            cmd.Parameters.AddWithValue("@thc", N(m.Invoice_thc));
            cmd.Parameters.AddWithValue("@blfee", N(m.Invoice_blFee));
            cmd.Parameters.AddWithValue("@seal", N(m.Invoice_seal));
            cmd.Parameters.AddWithValue("@telex", N(m.Invoice_telexRelease));
            cmd.Parameters.AddWithValue("@cfs", N(m.Invoice_cfs));
            cmd.Parameters.AddWithValue("@vgm", N(m.Invoice_vgmFee));
            cmd.Parameters.AddWithValue("@ens", N(m.Invoice_ensebsams));
            cmd.Parameters.AddWithValue("@other", N(m.Invoice_other));
            cmd.Parameters.AddWithValue("@totalvnd", N(m.Invoice_totalVND));
            cmd.Parameters.AddWithValue("@subtotalocean", N(m.Invoice_subTotalOcean));
            cmd.Parameters.AddWithValue("@cofee", N(m.Invoice_coFee));
            cmd.Parameters.AddWithValue("@feestatus", N(m.Invoice_feeStatus));
            cmd.Parameters.AddWithValue("@redinvno", N(m.Invoice_redInvoiceNo));
            cmd.Parameters.AddWithValue("@redinvdate", N(m.Invoice_redInvoiceDate));
            cmd.Parameters.AddWithValue("@redinvrecv", N(m.Invoice_redInvoiceRecvDate));
            cmd.Parameters.AddWithValue("@transferacct", N(m.Invoice_transferAccountantDate));
            cmd.Parameters.AddWithValue("@truck", N(m.Invoice_trucking));
            cmd.Parameters.AddWithValue("@infra", N(m.Invoice_infrastructureFee));
            cmd.Parameters.AddWithValue("@custclear", N(m.Invoice_customerClearance));
            cmd.Parameters.AddWithValue("@custfee", N(m.Invoice_customFee));
            cmd.Parameters.AddWithValue("@othercustfee", N(m.Invoice_otherCustomFee));
            cmd.Parameters.AddWithValue("@subtotalvndcust", N(m.Invoice_subTotalVNDCustom));
            cmd.Parameters.AddWithValue("@subtotalusdcust", N(m.Invoice_subTotalUSDCustom));
            cmd.Parameters.AddWithValue("@grandvnd", N(m.Invoice_grandTotalVND));
            cmd.Parameters.AddWithValue("@grandusd", N(m.Invoice_grandTotalUSD));
            cmd.Parameters.AddWithValue("@cdsno", N(m.Invoice_cdsNo));
            cmd.Parameters.AddWithValue("@cdsdate", N(m.Invoice_cdsDate));
            cmd.Parameters.AddWithValue("@line", N(m.Invoice_line));
            cmd.Parameters.AddWithValue("@customtype", N(m.Invoice_customType));
            cmd.Parameters.AddWithValue("@udate", now);
            cmd.Parameters.AddWithValue("@uby", byUser ?? (object)DBNull.Value);

            if (isInsert)
            {
                cmd.Parameters.AddWithValue("@cdate", (object)(m.createdate ?? now));
                cmd.Parameters.AddWithValue("@cby", byUser ?? (object)DBNull.Value);
            }
        }

        // ════════════════════════════════════════════════════════════════
        // Row mapper
        // ════════════════════════════════════════════════════════════════
        private static InvoiceModel Map(SqlDataReader r)
        {
            string S(string c) => r[c] == DBNull.Value ? null : r[c].ToString();
            double? F(string c) => r[c] == DBNull.Value ? (double?)null : Convert.ToDouble(r[c]);
            int? I(string c) => r[c] == DBNull.Value ? (int?)null : Convert.ToInt32(r[c]);
            DateTime? D(string c) => r[c] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[c]);

            return new InvoiceModel
            {
                Invoice_Id = S("Invoice_Id"),
                Invoice_no = S("Invoice_no"),
                Invoice_erpID = S("Invoice_erpID"),
                Invoice_erpInvoiceNo = S("Invoice_erpInvoiceNo"),
                Invoice_shippingTerm = S("Invoice_shippingTerm"),
                Invoice_paymentTerm = S("Invoice_paymentTerm"),
                Invoice_employee = S("Invoice_employee"),
                Invoice_logisticRemark = S("Invoice_logisticRemark"),
                Invoice_confirmDate = D("Invoice_confirmDate"),
                Invoice_fwdName = S("Invoice_fwdName"),
                Invoice_bookingNo = S("Invoice_bookingNo"),
                Invoice_contType = S("Invoice_contType"),
                Invoice_vgmCO = S("Invoice_vgmCO"),
                Invoice_cyCO = S("Invoice_cyCO"),
                Invoice_etd = D("Invoice_etd"),
                Invoice_eta = D("Invoice_eta"),
                Invoice_billType = S("Invoice_billType"),
                Invoice_billNo = S("Invoice_billNo"),
                Invoice_co = S("Invoice_co"),
                Invoice_coNo = S("Invoice_coNo"),
                Invoice_OF = F("Invoice_OF"),
                Invoice_deliveryCharges = F("Invoice_deliveryCharges"),
                Invoice_taxes = F("Invoice_taxes"),
                Invoice_otherDestCharges = F("Invoice_otherDestCharges"),
                Invoice_thc = F("Invoice_thc"),
                Invoice_blFee = F("Invoice_blFee"),
                Invoice_seal = F("Invoice_seal"),
                Invoice_telexRelease = F("Invoice_telexRelease"),
                Invoice_cfs = F("Invoice_cfs"),
                Invoice_vgmFee = F("Invoice_vgmFee"),
                Invoice_ensebsams = F("Invoice_ensebsams"),
                Invoice_other = F("Invoice_other"),
                Invoice_totalVND = F("Invoice_totalVND"),
                Invoice_subTotalOcean = F("Invoice_subTotalOcean"),
                Invoice_coFee = F("Invoice_coFee"),
                Invoice_feeStatus = I("Invoice_feeStatus"),
                Invoice_redInvoiceNo = S("Invoice_redInvoiceNo"),
                Invoice_redInvoiceDate = D("Invoice_redInvoiceDate"),
                Invoice_redInvoiceRecvDate = D("Invoice_redInvoiceRecvDate"),
                Invoice_transferAccountantDate = D("Invoice_transferAccountantDate"),
                Invoice_trucking = F("Invoice_trucking"),
                Invoice_infrastructureFee = F("Invoice_infrastructureFee"),
                Invoice_customerClearance = F("Invoice_customerClearance"),
                Invoice_customFee = F("Invoice_customFee"),
                Invoice_otherCustomFee = F("Invoice_otherCustomFee"),
                Invoice_subTotalVNDCustom = F("Invoice_subTotalVNDCustom"),
                Invoice_subTotalUSDCustom = F("Invoice_subTotalUSDCustom"),
                Invoice_grandTotalVND = F("Invoice_grandTotalVND"),
                Invoice_grandTotalUSD = F("Invoice_grandTotalUSD"),
                Invoice_cdsNo = S("Invoice_cdsNo"),
                Invoice_cdsDate = D("Invoice_cdsDate"),
                Invoice_line = S("Invoice_line"),
                Invoice_customType = S("Invoice_customType"),
                createdate = D("createdate"),
                createby = S("createby"),
                updatedate = D("updatedate"),
                updateby = S("updateby"),
            };
        }
    }
}