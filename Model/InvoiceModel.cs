using System;

namespace techlink_workspace.Model
{
    public class InvoiceModel
    {
        public string Invoice_Id { get; set; }
        public string Invoice_no { get; set; }

        // ── ERP (filled via ERP lookup button, not manual input) ─────────
        public string Invoice_erpID { get; set; }
        public string Invoice_erpInvoiceNo { get; set; }

        // ── Logistics Plan (Input) — editable by Logistic/Admin ──────────
        public string Invoice_shippingTerm { get; set; }
        public string Invoice_paymentTerm { get; set; }
        public string Invoice_employee { get; set; }
        public string Invoice_logisticRemark { get; set; }
        public DateTime? Invoice_confirmDate { get; set; }
        public string Invoice_fwdName { get; set; }
        public string Invoice_bookingNo { get; set; }
        public string Invoice_contType { get; set; }
        public string Invoice_vgmCO { get; set; }
        public string Invoice_cyCO { get; set; }
        public DateTime? Invoice_etd { get; set; }
        public DateTime? Invoice_eta { get; set; }
        public string Invoice_billType { get; set; }
        public string Invoice_billNo { get; set; }
        public string Invoice_co { get; set; }
        public string Invoice_coNo { get; set; }
        public double? Invoice_OF { get; set; }
        public double? Invoice_deliveryCharges { get; set; }
        public double? Invoice_taxes { get; set; }
        public double? Invoice_otherDestCharges { get; set; }
        public double? Invoice_thc { get; set; }
        public double? Invoice_blFee { get; set; }
        public double? Invoice_seal { get; set; }
        public double? Invoice_telexRelease { get; set; }
        public double? Invoice_cfs { get; set; }
        public double? Invoice_vgmFee { get; set; }
        public double? Invoice_ensebsams { get; set; }
        public double? Invoice_other { get; set; }
        public double? Invoice_totalVND { get; set; }
        public double? Invoice_subTotalOcean { get; set; }
        public double? Invoice_coFee { get; set; }
        public int? Invoice_feeStatus { get; set; }
        public string Invoice_redInvoiceNo { get; set; }
        public DateTime? Invoice_redInvoiceDate { get; set; }
        public DateTime? Invoice_redInvoiceRecvDate { get; set; }
        public DateTime? Invoice_transferAccountantDate { get; set; }

        // ── Logistics Customs — editable by Customs/Admin ────────────────
        public double? Invoice_trucking { get; set; }
        public double? Invoice_infrastructureFee { get; set; }
        public double? Invoice_customerClearance { get; set; }
        public double? Invoice_customFee { get; set; }
        public double? Invoice_otherCustomFee { get; set; }
        public double? Invoice_subTotalVNDCustom { get; set; }
        public double? Invoice_subTotalUSDCustom { get; set; }
        public double? Invoice_grandTotalVND { get; set; }
        public double? Invoice_grandTotalUSD { get; set; }
        public string Invoice_cdsNo { get; set; }
        public DateTime? Invoice_cdsDate { get; set; }
        public string Invoice_line { get; set; }
        public string Invoice_customType { get; set; }

        // ── Audit (auto-set, never shown in UI) ──────────────────────────
        public DateTime? createdate { get; set; }
        public string createby { get; set; }
        public DateTime? updatedate { get; set; }
        public string updateby { get; set; }
    }

    /// <summary>Which section a user type can edit.</summary>
    public enum InvoiceEditScope { None, LogisticPlan, CustomOnly, All }
}