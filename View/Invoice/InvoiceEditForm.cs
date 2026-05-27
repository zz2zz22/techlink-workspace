// ════════════════════════════════════════════════════════════════════════
// InvoiceEditForm.cs
// Compact two-column layout. Section permissions by user type name.
// ════════════════════════════════════════════════════════════════════════
using System;
using System.Drawing;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;

namespace techlink_workspace.View.Invoice
{
    public class InvoiceEditForm : Form
    {
        public InvoiceModel Result { get; private set; }

        private readonly string _byUser;
        private readonly bool _isNew;
        private readonly InvoiceEditScope _scope;   // what sections this user can edit
        private string _existingId;

        // ── ERP section (read-only, filled via ERP button) ────────────────
        private TextBox txtErpId, txtErpInvNo;
        private Button btnGetFromErp;

        // ── Basic / Plan fields ───────────────────────────────────────────
        private TextBox txtNo, txtShipTerm, txtPayTerm, txtEmployee,
                        txtRemark, txtFwdName, txtBookingNo,
                        txtContType, txtVgmCO, txtCyCO,
                        txtBillType, txtBillNo, txtCo, txtCoNo, txtRedInvNo;
        private DateTimePicker dtpConfirm, dtpEtd, dtpEta,
                               dtpRedInvDate, dtpRedInvRecv, dtpTransferAcct;
        private CheckBox chkConfirm, chkEtd, chkEta,
                         chkRedInvDate, chkRedInvRecv, chkTransferAcct;
        private NumericUpDown numOF, numDel, numTaxes, numOtherDest,
                              numThc, numBlFee, numSeal, numTelex, numCfs,
                              numVgm, numEns, numOther, numTotalVND,
                              numSubOcean, numCoFee;
        private ComboBox cmbFeeStatus;

        // ── Customs fields ────────────────────────────────────────────────
        private NumericUpDown numTruck, numInfra, numCustClear, numCustFee,
                              numOtherCust, numSubVNDCust, numSubUSDCust,
                              numGrandVND, numGrandUSD;
        private TextBox txtCdsNo, txtLine, txtCustomType;
        private DateTimePicker dtpCdsDate;
        private CheckBox chkCdsDate;

        private Button btnSave, btnCancel;
        private Label lblErr;

        // ── Layout constants ──────────────────────────────────────────────
        private const int LW = 170;  // label width
        private const int FW = 200;  // field width
        private const int RH = 26;   // row height
        private const int COL2 = 410;// x start of second column

        public InvoiceEditForm(InvoiceModel existing, string byUser,
                               bool isNew, InvoiceEditScope scope)
        {
            _byUser = byUser;
            _isNew = isNew;
            _scope = scope;
            _existingId = existing?.Invoice_Id;
            Build();
            if (existing != null) Populate(existing);
            ApplyScopePermissions();
        }

        // ════════════════════════════════════════════════════════════════
        private void Build()
        {
            Text = _isNew ? "Add Invoice" : "Edit Invoice";
            ClientSize = new Size(830, 680);
            MinimumSize = new Size(830, 500);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f);

            var scroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(10, 8, 10, 0)
            };

            int y = 0;

            // ── ERP section (button + read-only fields) ───────────────────
            SecHdr(scroll, "ERP Data", ref y);
            btnGetFromErp = new Button
            {
                Text = "🔗 Get from ERP",
                Location = new Point(0, y),
                Size = new Size(140, 26),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 8f)
            };
            btnGetFromErp.FlatAppearance.BorderSize = 0;
            btnGetFromErp.Click += (s, e) =>
                CTMessageBox.Show("ERP lookup not yet implemented.\nThis will link to the ERP table in a future release.",
                    "ERP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            scroll.Controls.Add(btnGetFromErp);
            txtErpId = ROField(scroll, "ERP ID", ref y, COL2);
            txtErpInvNo = ROField(scroll, "ERP Invoice No", ref y, COL2 + LW + FW / 2);
            y += RH + 4;

            // ── Logistics Plan section ────────────────────────────────────
            SecHdr(scroll, "Logistics Plan (Input)", ref y);

            TwoCol(scroll, ref y,
                "Invoice #", out txtNo,
                "Shipping Term", out txtShipTerm);
            TwoCol(scroll, ref y,
                "Payment Term", out txtPayTerm,
                "NV Logistics phụ trách", out txtEmployee);

            AddLabel(scroll, "Lưu ý đơn hàng", 0, y);
            txtRemark = new TextBox
            {
                Location = new Point(0, y + 18),
                Size = new Size(800, 46),
                Multiline = true,
                MaxLength = 2000
            };
            scroll.Controls.Add(txtRemark);
            y += 70;

            (chkConfirm, dtpConfirm) = DateRow(scroll, "Ngày xưởng confirm", ref y, 0);

            TwoCol(scroll, ref y,
                "FWD – Tên FWD", out txtFwdName,
                "Số Booking", out txtBookingNo);

            AddLabel(scroll, "Loại cont", 0, y);
            txtContType = new TextBox { Location = new Point(LW, y), Size = new Size(FW, RH - 2) };
            scroll.Controls.Add(txtContType);
            AddLabel(scroll, "Status paid/Acct/Not yet", COL2, y);
            cmbFeeStatus = new ComboBox
            {
                Location = new Point(COL2 + LW, y),
                Width = FW,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbFeeStatus.Items.AddRange(new object[] { "1 - Paid", "2 - Accounting", "3 - Not yet" });
            cmbFeeStatus.SelectedIndex = 0;
            scroll.Controls.Add(cmbFeeStatus);
            y += RH + 4;

            TwoCol(scroll, ref y,
                "SI & VGM cut-off", out txtVgmCO,
                "CY cut-off", out txtCyCO);

            (chkEtd, dtpEtd) = DateRow(scroll, "ETD", ref y, 0);
            (chkEta, dtpEta) = DateRow(scroll, "ETA", ref y, COL2);

            TwoCol(scroll, ref y,
                "Loại Bill", out txtBillType,
                "Số Bill", out txtBillNo);

            TwoCol(scroll, ref y,
                "CO", out txtCo,
                "Số CO", out txtCoNo);

            // ── Charges (3-per-row compact) ───────────────────────────────
            SecHdr(scroll, "Ocean Charges (USD/VND)", ref y);

            ThreeNum(scroll, ref y,
                "OF", out numOF,
                "Delivery charges", out numDel,
                "Duty/Taxes", out numTaxes);
            ThreeNum(scroll, ref y,
                "Other dest", out numOtherDest,
                "THC", out numThc,
                "b/l fee", out numBlFee);
            ThreeNum(scroll, ref y,
                "Seal", out numSeal,
                "Telex release", out numTelex,
                "CFS", out numCfs);
            ThreeNum(scroll, ref y,
                "VGM", out numVgm,
                "ENS/EBS/AMS", out numEns,
                "OTHERS", out numOther);
            ThreeNum(scroll, ref y,
                "TOTAL VND", out numTotalVND,
                "SUB-TOTAL OCEAN USD", out numSubOcean,
                "C/O fee", out numCoFee, dec0: true);

            // ── Red invoice / transfer ────────────────────────────────────
            SecHdr(scroll, "Invoice Status", ref y);
            AddLabel(scroll, "RED INVOICE #", 0, y);
            txtRedInvNo = new TextBox { Location = new Point(LW, y), Size = new Size(FW, RH - 2) };
            scroll.Controls.Add(txtRedInvNo); y += RH + 4;

            (chkRedInvDate, dtpRedInvDate) = DateRow(scroll, "RED INVOICE DATE", ref y, 0);
            (chkRedInvRecv, dtpRedInvRecv) = DateRow(scroll, "RED INV RECEIVED DATE", ref y, 0);
            (chkTransferAcct, dtpTransferAcct) = DateRow(scroll, "TRANSFER TO ACCOUNTANT", ref y, 0);

            // ── Customs section ───────────────────────────────────────────
            SecHdr(scroll, "Logistics Customs", ref y);
            ThreeNum(scroll, ref y,
                "Trucking", out numTruck,
                "Infra Fee", out numInfra,
                "Cust clearance", out numCustClear, dec0: true);
            ThreeNum(scroll, ref y,
                "Custom Fee", out numCustFee,
                "Other custom", out numOtherCust,
                "Sub VND Custom", out numSubVNDCust, dec0: true);
            ThreeNum(scroll, ref y,
                "Sub USD Custom", out numSubUSDCust,
                "GRAND TOTAL VND", out numGrandVND,
                "GRAND TOTAL USD", out numGrandUSD, dec0: true);

            TwoCol(scroll, ref y,
                "CDS NO", out txtCdsNo,
                "Luong", out txtLine);
            AddLabel(scroll, "Mã loại hình tk", 0, y);
            txtCustomType = new TextBox { Location = new Point(LW, y), Size = new Size(FW, RH - 2) };
            scroll.Controls.Add(txtCustomType);
            (chkCdsDate, dtpCdsDate) = DateRow(scroll, "CDS DATE", ref y, COL2);
            y += RH;

            // ── Error label ───────────────────────────────────────────────
            lblErr = new Label
            {
                Location = new Point(0, y + 4),
                Size = new Size(800, 20),
                ForeColor = Color.Crimson
            };
            scroll.Controls.Add(lblErr);

            // ── Bottom buttons ────────────────────────────────────────────
            var bot = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                BackColor = Color.WhiteSmoke
            };
            btnSave = new Button
            {
                Text = "Save",
                Size = new Size(110, 32),
                Location = new Point(12, 7),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Size = new Size(90, 32),
                Location = new Point(130, 7),
                BackColor = Color.FromArgb(200, 200, 200),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            bot.Controls.AddRange(new Control[] { btnSave, btnCancel });

            Controls.Add(scroll);
            Controls.Add(bot);
            AcceptButton = btnSave;
            CancelButton = btnCancel;
        }

        // ════════════════════════════════════════════════════════════════
        // Apply section permissions based on scope
        // ════════════════════════════════════════════════════════════════
        private void ApplyScopePermissions()
        {
            bool canPlan = _scope == InvoiceEditScope.LogisticPlan || _scope == InvoiceEditScope.All;
            bool canCustoms = _scope == InvoiceEditScope.CustomOnly || _scope == InvoiceEditScope.All;

            // Plan fields
            foreach (Control c in new Control[]
            {
                txtNo, txtShipTerm, txtPayTerm, txtEmployee, txtRemark,
                txtFwdName, txtBookingNo, txtContType, txtVgmCO, txtCyCO,
                txtBillType, txtBillNo, txtCo, txtCoNo, txtRedInvNo,
                dtpConfirm, dtpEtd, dtpEta, dtpRedInvDate, dtpRedInvRecv, dtpTransferAcct,
                chkConfirm, chkEtd, chkEta, chkRedInvDate, chkRedInvRecv, chkTransferAcct,
                numOF, numDel, numTaxes, numOtherDest, numThc, numBlFee,
                numSeal, numTelex, numCfs, numVgm, numEns, numOther,
                numTotalVND, numSubOcean, numCoFee, cmbFeeStatus
            })
                if (c != null) c.Enabled = canPlan;

            // Customs fields
            foreach (Control c in new Control[]
            {
                numTruck, numInfra, numCustClear, numCustFee, numOtherCust,
                numSubVNDCust, numSubUSDCust, numGrandVND, numGrandUSD,
                txtCdsNo, txtLine, txtCustomType,
                chkCdsDate, dtpCdsDate
            })
                if (c != null) c.Enabled = canCustoms;
        }

        // ════════════════════════════════════════════════════════════════
        private void Populate(InvoiceModel m)
        {
            txtErpId.Text = m.Invoice_erpID ?? "";
            txtErpInvNo.Text = m.Invoice_erpInvoiceNo ?? "";
            txtNo.Text = m.Invoice_no ?? "";
            txtShipTerm.Text = m.Invoice_shippingTerm ?? "";
            txtPayTerm.Text = m.Invoice_paymentTerm ?? "";
            txtEmployee.Text = m.Invoice_employee ?? "";
            txtRemark.Text = m.Invoice_logisticRemark ?? "";
            SetDate(chkConfirm, dtpConfirm, m.Invoice_confirmDate);
            txtFwdName.Text = m.Invoice_fwdName ?? "";
            txtBookingNo.Text = m.Invoice_bookingNo ?? "";
            txtContType.Text = m.Invoice_contType ?? "";
            txtVgmCO.Text = m.Invoice_vgmCO ?? "";
            txtCyCO.Text = m.Invoice_cyCO ?? "";
            SetDate(chkEtd, dtpEtd, m.Invoice_etd);
            SetDate(chkEta, dtpEta, m.Invoice_eta);
            txtBillType.Text = m.Invoice_billType ?? "";
            txtBillNo.Text = m.Invoice_billNo ?? "";
            txtCo.Text = m.Invoice_co ?? "";
            txtCoNo.Text = m.Invoice_coNo ?? "";
            SN(numOF, m.Invoice_OF);
            SN(numDel, m.Invoice_deliveryCharges);
            SN(numTaxes, m.Invoice_taxes);
            SN(numOtherDest, m.Invoice_otherDestCharges);
            SN(numThc, m.Invoice_thc);
            SN(numBlFee, m.Invoice_blFee);
            SN(numSeal, m.Invoice_seal);
            SN(numTelex, m.Invoice_telexRelease);
            SN(numCfs, m.Invoice_cfs);
            SN(numVgm, m.Invoice_vgmFee);
            SN(numEns, m.Invoice_ensebsams);
            SN(numOther, m.Invoice_other);
            SN(numTotalVND, m.Invoice_totalVND);
            SN(numSubOcean, m.Invoice_subTotalOcean);
            SN(numCoFee, m.Invoice_coFee);
            cmbFeeStatus.SelectedIndex = Math.Max(0, (m.Invoice_feeStatus ?? 3) - 1);
            txtRedInvNo.Text = m.Invoice_redInvoiceNo ?? "";
            SetDate(chkRedInvDate, dtpRedInvDate, m.Invoice_redInvoiceDate);
            SetDate(chkRedInvRecv, dtpRedInvRecv, m.Invoice_redInvoiceRecvDate);
            SetDate(chkTransferAcct, dtpTransferAcct, m.Invoice_transferAccountantDate);
            SN(numTruck, m.Invoice_trucking);
            SN(numInfra, m.Invoice_infrastructureFee);
            SN(numCustClear, m.Invoice_customerClearance);
            SN(numCustFee, m.Invoice_customFee);
            SN(numOtherCust, m.Invoice_otherCustomFee);
            SN(numSubVNDCust, m.Invoice_subTotalVNDCustom);
            SN(numSubUSDCust, m.Invoice_subTotalUSDCustom);
            SN(numGrandVND, m.Invoice_grandTotalVND);
            SN(numGrandUSD, m.Invoice_grandTotalUSD);
            txtCdsNo.Text = m.Invoice_cdsNo ?? "";
            SetDate(chkCdsDate, dtpCdsDate, m.Invoice_cdsDate);
            txtLine.Text = m.Invoice_line ?? "";
            txtCustomType.Text = m.Invoice_customType ?? "";
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            lblErr.Text = "";
            if (string.IsNullOrWhiteSpace(txtNo.Text))
            { lblErr.Text = "Invoice No is required."; return; }

            Result = new InvoiceModel
            {
                Invoice_Id = _isNew ? Guid.NewGuid().ToString() : _existingId,
                Invoice_no = txtNo.Text.Trim(),
                Invoice_erpID = txtErpId.Text.Trim(),
                Invoice_erpInvoiceNo = txtErpInvNo.Text.Trim(),
                Invoice_shippingTerm = txtShipTerm.Text.Trim(),
                Invoice_paymentTerm = txtPayTerm.Text.Trim(),
                Invoice_employee = txtEmployee.Text.Trim(),
                Invoice_logisticRemark = txtRemark.Text.Trim(),
                Invoice_confirmDate = chkConfirm.Checked ? dtpConfirm.Value : (DateTime?)null,
                Invoice_fwdName = txtFwdName.Text.Trim(),
                Invoice_bookingNo = txtBookingNo.Text.Trim(),
                Invoice_contType = txtContType.Text.Trim(),
                Invoice_vgmCO = txtVgmCO.Text.Trim(),
                Invoice_cyCO = txtCyCO.Text.Trim(),
                Invoice_etd = chkEtd.Checked ? dtpEtd.Value : (DateTime?)null,
                Invoice_eta = chkEta.Checked ? dtpEta.Value : (DateTime?)null,
                Invoice_billType = txtBillType.Text.Trim(),
                Invoice_billNo = txtBillNo.Text.Trim(),
                Invoice_co = txtCo.Text.Trim(),
                Invoice_coNo = txtCoNo.Text.Trim(),
                Invoice_OF = GN(numOF),
                Invoice_deliveryCharges = GN(numDel),
                Invoice_taxes = GN(numTaxes),
                Invoice_otherDestCharges = GN(numOtherDest),
                Invoice_thc = GN(numThc),
                Invoice_blFee = GN(numBlFee),
                Invoice_seal = GN(numSeal),
                Invoice_telexRelease = GN(numTelex),
                Invoice_cfs = GN(numCfs),
                Invoice_vgmFee = GN(numVgm),
                Invoice_ensebsams = GN(numEns),
                Invoice_other = GN(numOther),
                Invoice_totalVND = GN(numTotalVND),
                Invoice_subTotalOcean = GN(numSubOcean),
                Invoice_coFee = GN(numCoFee),
                Invoice_feeStatus = cmbFeeStatus.SelectedIndex + 1,
                Invoice_redInvoiceNo = txtRedInvNo.Text.Trim(),
                Invoice_redInvoiceDate = chkRedInvDate.Checked ? dtpRedInvDate.Value : (DateTime?)null,
                Invoice_redInvoiceRecvDate = chkRedInvRecv.Checked ? dtpRedInvRecv.Value : (DateTime?)null,
                Invoice_transferAccountantDate = chkTransferAcct.Checked ? dtpTransferAcct.Value : (DateTime?)null,
                Invoice_trucking = GN(numTruck),
                Invoice_infrastructureFee = GN(numInfra),
                Invoice_customerClearance = GN(numCustClear),
                Invoice_customFee = GN(numCustFee),
                Invoice_otherCustomFee = GN(numOtherCust),
                Invoice_subTotalVNDCustom = GN(numSubVNDCust),
                Invoice_subTotalUSDCustom = GN(numSubUSDCust),
                Invoice_grandTotalVND = GN(numGrandVND),
                Invoice_grandTotalUSD = GN(numGrandUSD),
                Invoice_cdsNo = txtCdsNo.Text.Trim(),
                Invoice_cdsDate = chkCdsDate.Checked ? dtpCdsDate.Value : (DateTime?)null,
                Invoice_line = txtLine.Text.Trim(),
                Invoice_customType = txtCustomType.Text.Trim(),
                createdate = DateTime.Now,
                createby = _byUser,
                updatedate = DateTime.Now,
                updateby = _byUser
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        // ════════════════════════════════════════════════════════════════
        // UI builder helpers — compact two-column layout
        // ════════════════════════════════════════════════════════════════
        private void SecHdr(Panel p, string text, ref int y)
        {
            p.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(0, y),
                Size = new Size(800, 22),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = Color.FromArgb(30, 40, 60),
                ForeColor = Color.White,
                Padding = new Padding(6, 0, 0, 0)
            });
            y += 26;
        }

        private void TwoCol(Panel p, ref int y,
            string lbl1, out TextBox tb1,
            string lbl2, out TextBox tb2)
        {
            AddLabel(p, lbl1, 0, y);
            tb1 = new TextBox { Location = new Point(LW, y), Size = new Size(FW, RH - 2) };
            p.Controls.Add(tb1);

            AddLabel(p, lbl2, COL2, y);
            tb2 = new TextBox { Location = new Point(COL2 + LW, y), Size = new Size(FW, RH - 2) };
            p.Controls.Add(tb2);
            y += RH + 4;
        }

        private (CheckBox, DateTimePicker) DateRow(Panel p, string lbl, ref int y, int xOff)
        {
            var chk = new CheckBox
            {
                Text = lbl,
                Location = new Point(xOff, y),
                Size = new Size(LW + FW / 2, RH),
                Font = new Font("Segoe UI", 8f)
            };
            var dtp = new DateTimePicker
            {
                Location = new Point(xOff + LW + FW / 2 + 4, y),
                Width = 120,
                Format = DateTimePickerFormat.Short,
                Enabled = false
            };
            chk.CheckedChanged += (s, e) => dtp.Enabled = chk.Checked;
            p.Controls.AddRange(new Control[] { chk, dtp });
            if (xOff == 0) y += RH + 2;
            return (chk, dtp);
        }

        private void ThreeNum(Panel p, ref int y,
            string l1, out NumericUpDown n1,
            string l2, out NumericUpDown n2,
            string l3, out NumericUpDown n3,
            bool dec0 = false)
        {
            int dec = dec0 ? 0 : 2;
            int cw = 260, nw = 90, lw2 = 160;
            AddLabel(p, l1, 0, y);
            n1 = NUD(p, lw2, y, nw, dec);
            AddLabel(p, l2, cw, y);
            n2 = NUD(p, cw + lw2, y, nw, dec);
            AddLabel(p, l3, cw * 2, y);
            n3 = NUD(p, cw * 2 + lw2, y, nw, dec);
            y += RH + 4;
        }

        private static NumericUpDown NUD(Panel p, int x, int y, int w, int dec)
        {
            var n = new NumericUpDown
            {
                Location = new Point(x, y),
                Size = new Size(w, RH - 2),
                Minimum = 0,
                Maximum = 999999999M,
                DecimalPlaces = dec
            };
            p.Controls.Add(n);
            return n;
        }

        private static TextBox ROField(Panel p, string lbl, ref int y, int xStart)
        {
            p.Controls.Add(new Label
            {
                Text = lbl,
                Location = new Point(xStart, y + 4),
                AutoSize = true,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 7.5f)
            });
            var tb = new TextBox
            {
                Location = new Point(xStart, y + 18),
                Size = new Size(FW, RH - 2),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.DimGray
            };
            p.Controls.Add(tb);
            return tb;
        }

        private static void AddLabel(Panel p, string text, int x, int y) =>
            p.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 4),
                AutoSize = true,
                ForeColor = Color.FromArgb(50, 50, 80),
                Font = new Font("Segoe UI", 7.5f)
            });

        private static void SetDate(CheckBox chk, DateTimePicker dtp, DateTime? val)
        {
            if (val.HasValue) { chk.Checked = true; dtp.Value = val.Value; }
            else chk.Checked = false;
        }

        private static void SN(NumericUpDown n, double? v)
        { if (v.HasValue) n.Value = Math.Min((decimal)v.Value, n.Maximum); }

        private static double? GN(NumericUpDown n)
        { return n.Value > 0 ? (double?)decimal.ToDouble(n.Value) : null; }
    }
}