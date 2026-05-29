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

        /// <summary>Level-3 users only see their own invoices.</summary>
        public DataTable GetByPIC(string picCode)
        {
            string sql = $@"SELECT * FROM {TABLE}
                            WHERE Invoice_logisticPersonInCharge = @pic
                            ORDER BY createdate DESC";
            var dt = new DataTable();
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@pic", picCode ?? "");
                using (var da = new SqlDataAdapter(cmd))
                    da.Fill(dt);
            }
            return dt;
        }

        public InvoiceModel GetById(string id)
        {
            string sql = $"SELECT * FROM {TABLE} WHERE Invoice_Id = @id";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@id", id ?? "");
                using (var r = cmd.ExecuteReader())
                    return r.Read() ? Map(r) : null;
            }
        }

        /// <summary>Used for duplicate detection during ERP import.</summary>
        public InvoiceModel GetByPoNoAndItemCode(string poNo, string itemCode)
        {
            string sql = $@"SELECT * FROM {TABLE}
                            WHERE Invoice_poNo = @poNo AND Invoice_itemCode = @ic";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@poNo", poNo ?? "");
                cmd.Parameters.AddWithValue("@ic", itemCode ?? "");
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
            { Bind(cmd, m, byUser, true); cmd.ExecuteNonQuery(); }
            _log.WriteLog(byUser, "INSERT", "INVOICE", 1,
                variable: m.Invoice_Id, updateData: JsonConvert.SerializeObject(m));
        }

        // ════════════════════════════════════════════════════════════════
        // BULK INSERT (ERP Excel import)
        // ════════════════════════════════════════════════════════════════
        public (int inserted, int updated, int failed) BulkInsert(
            IEnumerable<InvoiceModel> records, string byUser)
        {
            int ok = 0, upd = 0, fail = 0;
            var errors = new System.Text.StringBuilder();

            using (var conn = Open())
            {
                foreach (var m in records)
                {
                    if (string.IsNullOrWhiteSpace(m.Invoice_poNo))
                    { fail++; continue; }

                    try
                    {
                        var existing = GetByPoNoAndItemCode(m.Invoice_poNo, m.Invoice_itemCode);
                        if (existing != null)
                        {
                            // Preserve logistics data, update only ERP columns
                            existing.Invoice_customerCode = m.Invoice_customerCode;
                            existing.Invoice_customerName = m.Invoice_customerName;
                            existing.Invoice_customerRequestDate = m.Invoice_customerRequestDate;
                            existing.Invoice_poNo = m.Invoice_poNo;
                            existing.Invoice_poDate = m.Invoice_poDate;
                            existing.Invoice_saleName = m.Invoice_saleName;
                            existing.Invoice_factoryNo = m.Invoice_factoryNo;
                            existing.Invoice_factoryName = m.Invoice_factoryName;
                            existing.Invoice_itemCode = m.Invoice_itemCode;
                            existing.Invoice_itemCodeCustomers = m.Invoice_itemCodeCustomers;
                            existing.Invoice_itemName = m.Invoice_itemName;
                            existing.Invoice_quantity = m.Invoice_quantity;
                            existing.Invoice_unit = m.Invoice_unit;
                            // Auto-update PIC if changed
                            if (!string.IsNullOrWhiteSpace(m.Invoice_logisticPersonInCharge))
                                existing.Invoice_logisticPersonInCharge = m.Invoice_logisticPersonInCharge;
                            using (var cmd = new SqlCommand(UpdateSql(), conn))
                            { Bind(cmd, existing, byUser, false); cmd.ExecuteNonQuery(); }
                            upd++;
                        }
                        else
                        {
                            using (var cmd = new SqlCommand(InsertSql(), conn))
                            { Bind(cmd, m, byUser, true); cmd.ExecuteNonQuery(); }
                            ok++;
                        }
                    }
                    catch (SqlException ex)
                    {
                        fail++;
                        if (fail <= 5) errors.AppendLine($"{m.Invoice_poNo}/{m.Invoice_itemCode}: {ex.Message}");
                    }
                }
            }

            _log.WriteLog(byUser, "BULK_INSERT", "INVOICE", 1,
                updateData: $"{ok} inserted, {upd} updated, {fail} failed" +
                            (errors.Length > 0 ? "\n" + errors : ""));

            if (fail > 0 && errors.Length > 0)
                throw new Exception($"{ok} inserted, {upd} updated, {fail} failed.\n\n{errors}");

            return (ok, upd, fail);
        }

        // ════════════════════════════════════════════════════════════════
        // UPDATE
        // ════════════════════════════════════════════════════════════════
        public void Update(InvoiceModel m, string byUser)
        {
            var old = GetById(m.Invoice_Id);
            using (var conn = Open())
            using (var cmd = new SqlCommand(UpdateSql(), conn))
            { Bind(cmd, m, byUser, false); cmd.ExecuteNonQuery(); }
            _log.WriteLog(byUser, "UPDATE", "INVOICE", 1,
                variable: m.Invoice_Id,
                oldData: old != null ? JsonConvert.SerializeObject(old) : "",
                updateData: JsonConvert.SerializeObject(m));
        }

        // ════════════════════════════════════════════════════════════════
        // DELETE (Level-1 only – enforced in the View layer)
        // ════════════════════════════════════════════════════════════════
        public void Delete(string id, string byUser)
        {
            var old = GetById(id);
            string sql = $"DELETE FROM {TABLE} WHERE Invoice_Id = @id";
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            { cmd.Parameters.AddWithValue("@id", id ?? ""); cmd.ExecuteNonQuery(); }
            _log.WriteLog(byUser, "DELETE", "INVOICE", 1,
                variable: id,
                oldData: old != null ? JsonConvert.SerializeObject(old) : "");
        }

        // ════════════════════════════════════════════════════════════════
        // HISTORY / ROLLBACK
        // ════════════════════════════════════════════════════════════════
        public DataTable GetHistory(string invoiceId = null)
        {
            string sql = @"SELECT Log_Id, Log_EmployeeId, Log_WriteDate, Log_Function,
                                  Log_Result, Log_Variable, Log_OldData, Log_UpdateData
                           FROM dbo.Sys_LogOperation WHERE Log_Module='INVOICE'";
            if (!string.IsNullOrEmpty(invoiceId))
                sql += " AND (Log_Variable=@var OR Log_OldData LIKE @like OR Log_UpdateData LIKE @like)";
            sql += " ORDER BY Log_WriteDate DESC";

            var dt = new DataTable();
            using (var conn = Open())
            using (var cmd = new SqlCommand(sql, conn))
            {
                if (!string.IsNullOrEmpty(invoiceId))
                {
                    cmd.Parameters.AddWithValue("@var", invoiceId);
                    cmd.Parameters.AddWithValue("@like", $"%{invoiceId}%");
                }
                using (var da = new SqlDataAdapter(cmd)) da.Fill(dt);
            }
            return dt;
        }

        public void Rollback(string logId, string byUser)
        {
            const string fetch = "SELECT Log_OldData, Log_Variable FROM dbo.Sys_LogOperation WHERE Log_Id=@lid";
            string oldJson, id;
            using (var conn = Open())
            using (var cmd = new SqlCommand(fetch, conn))
            {
                cmd.Parameters.AddWithValue("@lid", logId);
                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) throw new Exception("Log entry not found.");
                    oldJson = r["Log_OldData"]?.ToString();
                    id = r["Log_Variable"]?.ToString();
                }
            }
            if (string.IsNullOrWhiteSpace(oldJson)) throw new Exception("No snapshot available.");
            var m = JsonConvert.DeserializeObject<InvoiceModel>(oldJson)
                    ?? throw new Exception("Failed to deserialize snapshot.");
            Delete(id, byUser);
            Insert(m, byUser);
            _log.WriteLog(byUser, "ROLLBACK", "INVOICE", 1,
                variable: id, updateData: $"Rolled back to log entry {logId}");
        }

        // ════════════════════════════════════════════════════════════════
        // SQL strings
        // ════════════════════════════════════════════════════════════════
        private static string InsertSql() => $@"
            INSERT INTO {TABLE} (
                Invoice_Id, Invoice_no,
                Invoice_customerCode, Invoice_customerName, Invoice_customerRequestDate,
                Invoice_poNo, Invoice_poDate, Invoice_brand, Invoice_saleName,
                Invoice_factoryNo, Invoice_factoryName,
                Invoice_itemCode, Invoice_itemCodeCustomers, Invoice_itemName,
                Invoice_quantity, Invoice_unit,
                Invoice_shippingTerm, Invoice_paymentTerm, Invoice_logisticPersonInCharge,
                Invoice_logisticNote, Invoice_factoryConfirmDate, Invoice_shippingStatus,
                Invoice_fwdName, Invoice_bookingNo, Invoice_contType,
                Invoice_vgmCO, Invoice_cyCO, Invoice_etd, Invoice_eta,
                Invoice_billType, Invoice_billNo, Invoice_co, Invoice_coNo,
                Invoice_OF, Invoice_deliveryCharges, Invoice_taxes, Invoice_otherDestCharges,
                Invoice_thc, Invoice_blFee, Invoice_seal, Invoice_telexRelease,
                Invoice_cfs, Invoice_vgmFee, Invoice_ensebsams, Invoice_other,
                Invoice_coFee, Invoice_totalVND, Invoice_subTotalOcean, Invoice_feeStatus,
                Invoice_redInvoiceNo, Invoice_redInvoiceDate,
                Invoice_redInvoiceRecvDate, Invoice_transferAccountantDate,
                Invoice_trucking, Invoice_infrastructureFee, Invoice_customerClearance,
                Invoice_customFee, Invoice_otherCustomFee,
                Invoice_subTotalVNDCustom, Invoice_subTotalUSDCustom,
                Invoice_grandTotalVND, Invoice_grandTotalUSD,
                Invoice_cdsNo, Invoice_cdsDate, Invoice_cdsApproved,
                Invoice_line, Invoice_customType,
                createdate, createby, updatedate, updateby
            ) VALUES (
                @id, @no,
                @custCode, @custName, @custReqDate,
                @poNo, @poDate, @brand, @saleName,
                @factoryNo, @factoryName,
                @itemCode, @itemCodeCust, @itemName,
                @qty, @unit,
                @shipTerm, @payTerm, @pic,
                @note, @confirmDate, @status,
                @fwdName, @bookNo, @contType,
                @vgmCO, @cyCO, @etd, @eta,
                @billType, @billNo, @co, @coNo,
                @OF, @del, @taxes, @otherDest,
                @thc, @bl, @seal, @telex,
                @cfs, @vgm, @ens, @other,
                @coFee, @totalVND, @subOcean, @feeStatus,
                @redInvNo, @redInvDate,
                @redInvRecv, @transferAcct,
                @truck, @infra, @custClear,
                @custFee, @otherCust,
                @subVNDCust, @subUSDCust,
                @grandVND, @grandUSD,
                @cdsNo, @cdsDate, @cdsApproved,
                @line, @customType,
                @cdate, @cby, @udate, @uby
            )";

        private static string UpdateSql() => $@"
            UPDATE {TABLE} SET
                Invoice_no=@no,
                Invoice_customerCode=@custCode, Invoice_customerName=@custName,
                Invoice_customerRequestDate=@custReqDate,
                Invoice_poNo=@poNo, Invoice_poDate=@poDate, Invoice_brand=@brand, Invoice_saleName=@saleName,
                Invoice_factoryNo=@factoryNo, Invoice_factoryName=@factoryName,
                Invoice_itemCode=@itemCode, Invoice_itemCodeCustomers=@itemCodeCust,
                Invoice_itemName=@itemName, Invoice_quantity=@qty, Invoice_unit=@unit,
                Invoice_shippingTerm=@shipTerm, Invoice_paymentTerm=@payTerm,
                Invoice_logisticPersonInCharge=@pic,
                Invoice_logisticNote=@note, Invoice_factoryConfirmDate=@confirmDate,
                Invoice_shippingStatus=@status,
                Invoice_fwdName=@fwdName, Invoice_bookingNo=@bookNo,
                Invoice_contType=@contType, Invoice_vgmCO=@vgmCO, Invoice_cyCO=@cyCO,
                Invoice_etd=@etd, Invoice_eta=@eta,
                Invoice_billType=@billType, Invoice_billNo=@billNo,
                Invoice_co=@co, Invoice_coNo=@coNo,
                Invoice_OF=@OF, Invoice_deliveryCharges=@del,
                Invoice_taxes=@taxes, Invoice_otherDestCharges=@otherDest,
                Invoice_thc=@thc, Invoice_blFee=@bl, Invoice_seal=@seal,
                Invoice_telexRelease=@telex, Invoice_cfs=@cfs, Invoice_vgmFee=@vgm,
                Invoice_ensebsams=@ens, Invoice_other=@other,
                Invoice_coFee=@coFee, Invoice_totalVND=@totalVND,
                Invoice_subTotalOcean=@subOcean, Invoice_feeStatus=@feeStatus,
                Invoice_redInvoiceNo=@redInvNo, Invoice_redInvoiceDate=@redInvDate,
                Invoice_redInvoiceRecvDate=@redInvRecv,
                Invoice_transferAccountantDate=@transferAcct,
                Invoice_trucking=@truck, Invoice_infrastructureFee=@infra,
                Invoice_customerClearance=@custClear, Invoice_customFee=@custFee,
                Invoice_otherCustomFee=@otherCust,
                Invoice_subTotalVNDCustom=@subVNDCust, Invoice_subTotalUSDCustom=@subUSDCust,
                Invoice_grandTotalVND=@grandVND, Invoice_grandTotalUSD=@grandUSD,
                Invoice_cdsNo=@cdsNo, Invoice_cdsDate=@cdsDate,
                Invoice_cdsApproved=@cdsApproved,
                Invoice_line=@line, Invoice_customType=@customType,
                updatedate=@udate, updateby=@uby
            WHERE Invoice_Id=@id";

        // ════════════════════════════════════════════════════════════════
        // Param binder
        // ════════════════════════════════════════════════════════════════
        private static void Bind(SqlCommand cmd, InvoiceModel m, string byUser, bool isInsert)
        {
            var now = DateTime.Now;
            object N(object v) => v ?? DBNull.Value;

            cmd.Parameters.AddWithValue("@id", N(m.Invoice_Id));
            cmd.Parameters.AddWithValue("@no", N(m.Invoice_no));
            cmd.Parameters.AddWithValue("@custCode", N(m.Invoice_customerCode));
            cmd.Parameters.AddWithValue("@custName", N(m.Invoice_customerName));
            cmd.Parameters.AddWithValue("@custReqDate", N(m.Invoice_customerRequestDate));
            cmd.Parameters.AddWithValue("@poNo", N(m.Invoice_poNo));
            cmd.Parameters.AddWithValue("@poDate", N(m.Invoice_poDate));
            cmd.Parameters.AddWithValue("@brand", N(m.Invoice_brand));
            cmd.Parameters.AddWithValue("@saleName", N(m.Invoice_saleName));
            cmd.Parameters.AddWithValue("@factoryNo", N(m.Invoice_factoryNo));
            cmd.Parameters.AddWithValue("@factoryName", N(m.Invoice_factoryName));
            cmd.Parameters.AddWithValue("@itemCode", N(m.Invoice_itemCode));
            cmd.Parameters.AddWithValue("@itemCodeCust", N(m.Invoice_itemCodeCustomers));
            cmd.Parameters.AddWithValue("@itemName", N(m.Invoice_itemName));
            cmd.Parameters.AddWithValue("@qty", N(m.Invoice_quantity));
            cmd.Parameters.AddWithValue("@unit", N(m.Invoice_unit));
            cmd.Parameters.AddWithValue("@shipTerm", N(m.Invoice_shippingTerm));
            cmd.Parameters.AddWithValue("@payTerm", N(m.Invoice_paymentTerm));
            cmd.Parameters.AddWithValue("@pic", N(m.Invoice_logisticPersonInCharge));
            cmd.Parameters.AddWithValue("@note", N(m.Invoice_logisticNote));
            cmd.Parameters.AddWithValue("@confirmDate", N(m.Invoice_factoryConfirmDate));
            cmd.Parameters.AddWithValue("@status", N(m.Invoice_shippingStatus));
            cmd.Parameters.AddWithValue("@fwdName", N(m.Invoice_fwdName));
            cmd.Parameters.AddWithValue("@bookNo", N(m.Invoice_bookingNo));
            cmd.Parameters.AddWithValue("@contType", N(m.Invoice_contType));
            cmd.Parameters.AddWithValue("@vgmCO", N(m.Invoice_vgmCO));
            cmd.Parameters.AddWithValue("@cyCO", N(m.Invoice_cyCO));
            cmd.Parameters.AddWithValue("@etd", N(m.Invoice_etd));
            cmd.Parameters.AddWithValue("@eta", N(m.Invoice_eta));
            cmd.Parameters.AddWithValue("@billType", N(m.Invoice_billType));
            cmd.Parameters.AddWithValue("@billNo", N(m.Invoice_billNo));
            cmd.Parameters.AddWithValue("@co", N(m.Invoice_co));
            cmd.Parameters.AddWithValue("@coNo", N(m.Invoice_coNo));
            cmd.Parameters.AddWithValue("@OF", N(m.Invoice_OF));
            cmd.Parameters.AddWithValue("@del", N(m.Invoice_deliveryCharges));
            cmd.Parameters.AddWithValue("@taxes", N(m.Invoice_taxes));
            cmd.Parameters.AddWithValue("@otherDest", N(m.Invoice_otherDestCharges));
            cmd.Parameters.AddWithValue("@thc", N(m.Invoice_thc));
            cmd.Parameters.AddWithValue("@bl", N(m.Invoice_blFee));
            cmd.Parameters.AddWithValue("@seal", N(m.Invoice_seal));
            cmd.Parameters.AddWithValue("@telex", N(m.Invoice_telexRelease));
            cmd.Parameters.AddWithValue("@cfs", N(m.Invoice_cfs));
            cmd.Parameters.AddWithValue("@vgm", N(m.Invoice_vgmFee));
            cmd.Parameters.AddWithValue("@ens", N(m.Invoice_ensebsams));
            cmd.Parameters.AddWithValue("@other", N(m.Invoice_other));
            cmd.Parameters.AddWithValue("@coFee", N(m.Invoice_coFee));
            cmd.Parameters.AddWithValue("@totalVND", N(m.Invoice_totalVND));
            cmd.Parameters.AddWithValue("@subOcean", N(m.Invoice_subTotalOcean));
            cmd.Parameters.AddWithValue("@feeStatus", N(m.Invoice_feeStatus));
            cmd.Parameters.AddWithValue("@redInvNo", N(m.Invoice_redInvoiceNo));
            cmd.Parameters.AddWithValue("@redInvDate", N(m.Invoice_redInvoiceDate));
            cmd.Parameters.AddWithValue("@redInvRecv", N(m.Invoice_redInvoiceRecvDate));
            cmd.Parameters.AddWithValue("@transferAcct", N(m.Invoice_transferAccountantDate));
            cmd.Parameters.AddWithValue("@truck", N(m.Invoice_trucking));
            cmd.Parameters.AddWithValue("@infra", N(m.Invoice_infrastructureFee));
            cmd.Parameters.AddWithValue("@custClear", N(m.Invoice_customerClearance));
            cmd.Parameters.AddWithValue("@custFee", N(m.Invoice_customFee));
            cmd.Parameters.AddWithValue("@otherCust", N(m.Invoice_otherCustomFee));
            cmd.Parameters.AddWithValue("@subVNDCust", N(m.Invoice_subTotalVNDCustom));
            cmd.Parameters.AddWithValue("@subUSDCust", N(m.Invoice_subTotalUSDCustom));
            cmd.Parameters.AddWithValue("@grandVND", N(m.Invoice_grandTotalVND));
            cmd.Parameters.AddWithValue("@grandUSD", N(m.Invoice_grandTotalUSD));
            cmd.Parameters.AddWithValue("@cdsNo", N(m.Invoice_cdsNo));
            cmd.Parameters.AddWithValue("@cdsDate", N(m.Invoice_cdsDate));
            cmd.Parameters.AddWithValue("@cdsApproved", N(m.Invoice_cdsApproved));
            cmd.Parameters.AddWithValue("@line", N(m.Invoice_line));
            cmd.Parameters.AddWithValue("@customType", N(m.Invoice_customType));
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
            decimal? D(string c) => r[c] == DBNull.Value ? (decimal?)null : Convert.ToDecimal(r[c]);
            int? I(string c) => r[c] == DBNull.Value ? (int?)null : Convert.ToInt32(r[c]);
            DateTime? DT(string c) => r[c] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[c]);
            bool? B(string c) => r[c] == DBNull.Value ? (bool?)null : Convert.ToBoolean(r[c]);

            return new InvoiceModel
            {
                Invoice_Id = S("Invoice_Id"),
                Invoice_no = S("Invoice_no"),
                Invoice_customerCode = S("Invoice_customerCode"),
                Invoice_customerName = S("Invoice_customerName"),
                Invoice_customerRequestDate = DT("Invoice_customerRequestDate"),
                Invoice_poNo = S("Invoice_poNo"),
                Invoice_poDate = DT("Invoice_poDate"),
                Invoice_brand = S("Invoice_brand"),
                Invoice_saleName = S("Invoice_saleName"),
                Invoice_factoryNo = S("Invoice_factoryNo"),
                Invoice_factoryName = S("Invoice_factoryName"),
                Invoice_itemCode = S("Invoice_itemCode"),
                Invoice_itemCodeCustomers = S("Invoice_itemCodeCustomers"),
                Invoice_itemName = S("Invoice_itemName"),
                Invoice_quantity = D("Invoice_quantity"),
                Invoice_unit = S("Invoice_unit"),
                Invoice_shippingTerm = S("Invoice_shippingTerm"),
                Invoice_paymentTerm = S("Invoice_paymentTerm"),
                Invoice_logisticPersonInCharge = S("Invoice_logisticPersonInCharge"),
                Invoice_logisticNote = S("Invoice_logisticNote"),
                Invoice_factoryConfirmDate = DT("Invoice_factoryConfirmDate"),
                Invoice_shippingStatus = I("Invoice_shippingStatus"),
                Invoice_fwdName = S("Invoice_fwdName"),
                Invoice_bookingNo = S("Invoice_bookingNo"),
                Invoice_contType = S("Invoice_contType"),
                Invoice_vgmCO = S("Invoice_vgmCO"),
                Invoice_cyCO = S("Invoice_cyCO"),
                Invoice_etd = DT("Invoice_etd"),
                Invoice_eta = DT("Invoice_eta"),
                Invoice_billType = S("Invoice_billType"),
                Invoice_billNo = S("Invoice_billNo"),
                Invoice_co = B("Invoice_co"),
                Invoice_coNo = S("Invoice_coNo"),
                Invoice_OF = D("Invoice_OF"),
                Invoice_deliveryCharges = D("Invoice_deliveryCharges"),
                Invoice_taxes = D("Invoice_taxes"),
                Invoice_otherDestCharges = D("Invoice_otherDestCharges"),
                Invoice_thc = D("Invoice_thc"),
                Invoice_blFee = D("Invoice_blFee"),
                Invoice_seal = D("Invoice_seal"),
                Invoice_telexRelease = D("Invoice_telexRelease"),
                Invoice_cfs = D("Invoice_cfs"),
                Invoice_vgmFee = D("Invoice_vgmFee"),
                Invoice_ensebsams = D("Invoice_ensebsams"),
                Invoice_other = D("Invoice_other"),
                Invoice_coFee = D("Invoice_coFee"),
                Invoice_totalVND = D("Invoice_totalVND"),
                Invoice_subTotalOcean = D("Invoice_subTotalOcean"),
                Invoice_feeStatus = I("Invoice_feeStatus"),
                Invoice_redInvoiceNo = S("Invoice_redInvoiceNo"),
                Invoice_redInvoiceDate = DT("Invoice_redInvoiceDate"),
                Invoice_redInvoiceRecvDate = DT("Invoice_redInvoiceRecvDate"),
                Invoice_transferAccountantDate = DT("Invoice_transferAccountantDate"),
                Invoice_trucking = D("Invoice_trucking"),
                Invoice_infrastructureFee = D("Invoice_infrastructureFee"),
                Invoice_customerClearance = D("Invoice_customerClearance"),
                Invoice_customFee = D("Invoice_customFee"),
                Invoice_otherCustomFee = D("Invoice_otherCustomFee"),
                Invoice_subTotalVNDCustom = D("Invoice_subTotalVNDCustom"),
                Invoice_subTotalUSDCustom = D("Invoice_subTotalUSDCustom"),
                Invoice_grandTotalVND = D("Invoice_grandTotalVND"),
                Invoice_grandTotalUSD = D("Invoice_grandTotalUSD"),
                Invoice_cdsNo = S("Invoice_cdsNo"),
                Invoice_cdsDate = DT("Invoice_cdsDate"),
                Invoice_cdsApproved = B("Invoice_cdsApproved"),
                Invoice_line = S("Invoice_line"),
                Invoice_customType = S("Invoice_customType"),
                createdate = DT("createdate"),
                createby = S("createby"),
                updatedate = DT("updatedate"),
                updateby = S("updateby"),
            };
        }
    }
}