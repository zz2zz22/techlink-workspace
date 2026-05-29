using System;

namespace techlink_workspace.Model
{
    public class InvoiceModel
    {
        public string Invoice_Id { get; set; }
        public string Invoice_no { get; set; }   // ERP invoice no – filled by logistics

        // ── ERP Section (locked after import – Level 1 only can edit) ────────
        public string Invoice_customerCode { get; set; }
        public string Invoice_customerName { get; set; }
        public DateTime? Invoice_customerRequestDate { get; set; }
        public string Invoice_poNo { get; set; }
        public DateTime? Invoice_poDate { get; set; }
        public string Invoice_brand { get; set; }
        public string Invoice_saleName { get; set; }
        public string Invoice_factoryNo { get; set; }
        public string Invoice_factoryName { get; set; }
        public string Invoice_itemCode { get; set; }
        public string Invoice_itemCodeCustomers { get; set; }
        public string Invoice_itemName { get; set; }
        public decimal? Invoice_quantity { get; set; }
        public string Invoice_unit { get; set; }

        // ── Logistics Plan ────────────────────────────────────────────────────
        public string Invoice_shippingTerm { get; set; }
        public string Invoice_paymentTerm { get; set; }
        public string Invoice_logisticPersonInCharge { get; set; }
        public string Invoice_logisticNote { get; set; }
        public DateTime? Invoice_factoryConfirmDate { get; set; }
        public int? Invoice_shippingStatus { get; set; }
        public string Invoice_fwdName { get; set; }
        public string Invoice_bookingNo { get; set; }
        public string Invoice_contType { get; set; }
        public string Invoice_vgmCO { get; set; }
        public string Invoice_cyCO { get; set; }
        public DateTime? Invoice_etd { get; set; }
        public DateTime? Invoice_eta { get; set; }
        public string Invoice_billType { get; set; }
        public string Invoice_billNo { get; set; }
        public bool? Invoice_co { get; set; }
        public string Invoice_coNo { get; set; }

        // ── Ocean Charges (decimal for accuracy) ─────────────────────────────
        public decimal? Invoice_OF { get; set; }
        public decimal? Invoice_deliveryCharges { get; set; }
        public decimal? Invoice_taxes { get; set; }
        public decimal? Invoice_otherDestCharges { get; set; }
        public decimal? Invoice_thc { get; set; }
        public decimal? Invoice_blFee { get; set; }
        public decimal? Invoice_seal { get; set; }
        public decimal? Invoice_telexRelease { get; set; }
        public decimal? Invoice_cfs { get; set; }
        public decimal? Invoice_vgmFee { get; set; }
        public decimal? Invoice_ensebsams { get; set; }
        public decimal? Invoice_other { get; set; }
        public decimal? Invoice_coFee { get; set; }
        public decimal? Invoice_totalVND { get; set; }
        public decimal? Invoice_subTotalOcean { get; set; }
        public int? Invoice_feeStatus { get; set; }
        public string Invoice_redInvoiceNo { get; set; }
        public DateTime? Invoice_redInvoiceDate { get; set; }
        public DateTime? Invoice_redInvoiceRecvDate { get; set; }
        public DateTime? Invoice_transferAccountantDate { get; set; }

        // ── Customs ───────────────────────────────────────────────────────────
        public decimal? Invoice_trucking { get; set; }
        public decimal? Invoice_infrastructureFee { get; set; }
        public decimal? Invoice_customerClearance { get; set; }
        public decimal? Invoice_customFee { get; set; }
        public decimal? Invoice_otherCustomFee { get; set; }
        public decimal? Invoice_subTotalVNDCustom { get; set; }
        public decimal? Invoice_subTotalUSDCustom { get; set; }
        public decimal? Invoice_grandTotalVND { get; set; }
        public decimal? Invoice_grandTotalUSD { get; set; }
        public string Invoice_cdsNo { get; set; }
        public DateTime? Invoice_cdsDate { get; set; }
        public bool? Invoice_cdsApproved { get; set; }
        public string Invoice_line { get; set; }
        public string Invoice_customType { get; set; }

        // ── Audit ─────────────────────────────────────────────────────────────
        public DateTime? createdate { get; set; }
        public string createby { get; set; }
        public DateTime? updatedate { get; set; }
        public string updateby { get; set; }
    }

    /// <summary>
    /// None     = read-only
    /// LevelThree = level-3 staff: ERP locked, prices lock once set, ETD+7d locks all
    /// ErpLocked  = level-2 manager: ERP locked, everything else editable
    /// All        = level-1 admin: unrestricted
    /// </summary>
    public enum InvoiceEditScope { None, LevelThree, ErpLocked, All }
}