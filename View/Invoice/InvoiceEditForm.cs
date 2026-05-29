using System;
using System.Drawing;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;

namespace techlink_workspace.View.Invoice
{
    /// <summary>
    /// Edit form using TabControl: ERP Data | Logistics Plan | Charges | Customs.
    ///
    /// Scope rules:
    ///   All        (Level 1) – every field editable
    ///   ErpLocked  (Level 2) – ERP tab read-only; rest editable
    ///   LevelThree (Level 3) – ERP locked; price fields locked if already set;
    ///                          everything locked if ETD > 7 days ago
    ///   None                 – fully read-only
    /// </summary>
    public class InvoiceEditForm : Form
    {
        public InvoiceModel Result { get; private set; }

        private readonly string _byUser;
        private readonly bool _isNew;
        private readonly InvoiceEditScope _scope;
        private readonly InvoiceModel _existing;
        private string _existingId;

        // ── Tab 1: ERP ──────────────────────────────────────────
        private TextBox txtCustCode, txtCustName, txtPoNo, txtBrand,
                        txtSaleName, txtFactoryNo, txtFactoryName,
                        txtItemCode, txtItemCodeCust, txtItemName, txtUnit;
        private DateTimePicker dtpCustReqDate, dtpPoDate;
        private CheckBox chkCustReqDate, chkPoDate;
        private NumericUpDown numQty;

        // ── Tab 2: Logistics Plan ───────────────────────────────
        private TextBox txtNo, txtShipTerm, txtPayTerm, txtPIC,
                        txtNote, txtFwdName, txtBookNo,
                        txtContType, txtVgmCO, txtCyCO,
                        txtBillType, txtBillNo, txtCoNo;
        private DateTimePicker dtpConfirm, dtpEtd, dtpEta;
        private CheckBox chkConfirm, chkEtd, chkEta, chkCo;
        private ComboBox cmbStatus, cmbFeeStatus;

        // ── Tab 3: Charges ──────────────────────────────────────
        private NumericUpDown numOF, numDel, numTaxes, numOtherDest, numThc, numBl,
                              numSeal, numTelex, numCfs, numVgm, numEns, numOther,
                              numCoFee, numTotalVND, numSubOcean;
        private TextBox txtRedInvNo;
        private DateTimePicker dtpRedInvDate, dtpRedInvRecv, dtpTransferAcct;
        private CheckBox chkRedInvDate, chkRedInvRecv, chkTransferAcct;

        // ── Tab 4: Customs ──────────────────────────────────────
        private NumericUpDown numTruck, numInfra, numCustClear, numCustFee, numOtherCust,
                              numSubVNDCust, numSubUSDCust, numGrandVND, numGrandUSD;
        private TextBox txtCdsNo, txtLine, txtCustomType;
        private DateTimePicker dtpCdsDate;
        private CheckBox chkCdsDate, chkCdsApproved;

        private Button btnSave, btnCancel;
        private Label lblErr;

        public InvoiceEditForm(InvoiceModel existing, string byUser, bool isNew, InvoiceEditScope scope)
        {
            _byUser = byUser; _isNew = isNew; _scope = scope;
            _existing = existing;
            _existingId = existing?.Invoice_Id;
            Build();
            if (existing != null) Populate(existing);
            ApplyPermissions();
        }

        // ════════════════════════════════════════════════════════
        private void Build()
        {
            Text = _isNew ? "New Invoice Record" : "Edit Invoice Record";
            ClientSize = new Size(820, 600);
            MinimumSize = new Size(700, 500);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f);

            var tabs = new TabControl { Dock = DockStyle.Fill };
            var tabErp = new TabPage("📋 ERP Data");
            var tabPlan = new TabPage("🚢 Logistics Plan");
            var tabCharges = new TabPage("💰 Charges");
            var tabCustoms = new TabPage("🏭 Customs");
            tabs.TabPages.AddRange(new[] { tabErp, tabPlan, tabCharges, tabCustoms });

            BuildErpTab(tabErp);
            BuildPlanTab(tabPlan);
            BuildChargesTab(tabCharges);
            BuildCustomsTab(tabCustoms);

            var bot = new Panel { Dock = DockStyle.Bottom, Height = 46, BackColor = Color.WhiteSmoke };
            btnSave = BtnCtrl("Save", Color.FromArgb(0, 122, 204), 12);
            btnSave.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            btnSave.Click += BtnSave_Click;
            btnCancel = BtnCtrl("Cancel", Color.FromArgb(200, 200, 200), 130);
            btnCancel.ForeColor = Color.FromArgb(33, 33, 33);
            btnCancel.DialogResult = DialogResult.Cancel;
            lblErr = new Label
            {
                Location = new Point(240, 14),
                AutoSize = true,
                ForeColor = Color.Crimson,
                Font = new Font("Segoe UI", 8.5f)
            };
            bot.Controls.AddRange(new Control[] { btnSave, btnCancel, lblErr });

            Controls.Add(tabs);
            Controls.Add(bot);
            AcceptButton = btnSave; CancelButton = btnCancel;
        }

        // ── Tab builders ──────────────────────────────────────────────────
        private void BuildErpTab(TabPage p)
        {
            var s = Scroll(p);
            int y = 0;

            TwoCol(s, ref y, "Mã KH (Customer Code)", out txtCustCode,
                             "KH Tên gọi tắt (Customer Name)", out txtCustName);

            DateRow2(s, ref y,
                "Ngày dự kiến giao hàng", out chkCustReqDate, out dtpCustReqDate,
                "Ngày ĐĐH (PO Date)", out chkPoDate, out dtpPoDate);

            TwoCol(s, ref y, "Mã ĐĐH của KH (PO No)", out txtPoNo,
                             "Bộ Phận (Brand)", out txtBrand);     // ← NEW row

            TwoCol(s, ref y, "NVBH (Sale Name)", out txtSaleName,
                             "Mã Xưởng (Factory No)", out txtFactoryNo);

            TwoCol(s, ref y, "Tên Xưởng (Factory Name)", out txtFactoryName,
                             "Mã SP (Item Code)", out txtItemCode);

            TwoCol(s, ref y, "Mã SP KH (Item Code Cust)", out txtItemCodeCust,
                             "Tên SP (Item Name)", out txtItemName);

            s.Controls.Add(Lbl("SL Đơn Đặt (Qty)", 0, y));
            numQty = new NumericUpDown
            {
                Location = new Point(160, y),
                Size = new Size(120, 22),
                Minimum = 0,
                Maximum = 9999999M,
                DecimalPlaces = 3
            };
            s.Controls.Add(numQty);
            s.Controls.Add(Lbl("ĐV (Unit)", 300, y));
            txtUnit = new TextBox { Location = new Point(370, y), Size = new Size(120, 22) };
            s.Controls.Add(txtUnit);
            y += 30;

            s.Controls.Add(new Label
            {
                Text = "⚠  ERP fields are locked after import. Only Level-1 Admin can modify.",
                Location = new Point(0, y + 4),
                AutoSize = true,
                ForeColor = Color.DarkOrange,
                Font = new Font("Segoe UI", 8f, FontStyle.Bold)
            });
        }

        private void BuildPlanTab(TabPage p)
        {
            var s = Scroll(p);
            int y = 0;
            TwoCol(s, ref y, "ERP Invoice No", out txtNo, "Shipping Term", out txtShipTerm);
            TwoCol(s, ref y, "Payment Term", out txtPayTerm, "PIC (Logistics Staff)", out txtPIC);
            s.Controls.Add(Lbl("Logistics Note", 0, y));
            txtNote = new TextBox
            {
                Location = new Point(160, y),
                Size = new Size(580, 54),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            s.Controls.Add(txtNote); y += 62;
            DateRow(s, ref y, "Factory Confirm Date", out chkConfirm, out dtpConfirm);

            s.Controls.Add(Lbl("Shipping Status", 0, y));
            cmbStatus = new ComboBox
            {
                Location = new Point(160, y),
                Width = 180,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbStatus.Items.AddRange(new object[] { "BLANK", "BOOKED", "CANCEL", "SHIPPED", "WAITING" });
            cmbStatus.SelectedIndex = 0;
            s.Controls.Add(cmbStatus); y += 30;

            TwoCol(s, ref y, "FWD Name", out txtFwdName, "Booking No", out txtBookNo);
            TwoCol(s, ref y, "Cont Type", out txtContType, "SI & VGM C/O", out txtVgmCO);
            TwoCol(s, ref y, "CY C/O", out txtCyCO, "Bill Type", out txtBillType);
            TwoCol(s, ref y, "Bill No", out txtBillNo, "C/O No", out txtCoNo);
            DateRow2(s, ref y, "ETD", out chkEtd, out dtpEtd, "ETA", out chkEta, out dtpEta);

            s.Controls.Add(Lbl("C/O", 0, y));
            chkCo = new CheckBox { Text = "Yes", Location = new Point(160, y), AutoSize = true };
            s.Controls.Add(chkCo);

            s.Controls.Add(Lbl("Fee Status", 300, y));
            cmbFeeStatus = new ComboBox
            {
                Location = new Point(460, y),
                Width = 160,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbFeeStatus.Items.AddRange(new object[] { "1 - Paid", "2 - Accounting", "3 - Not yet" });
            cmbFeeStatus.SelectedIndex = 0;
            s.Controls.Add(cmbFeeStatus); y += 30;
        }

        private void BuildChargesTab(TabPage p)
        {
            var s = Scroll(p);
            int y = 0;
            SectionHdr(s, "Ocean Charges (USD)", ref y);
            ThreeNum(s, ref y, "OF", out numOF, "Delivery Charges", out numDel, "Duty/Taxes", out numTaxes);
            ThreeNum(s, ref y, "Other Dest", out numOtherDest, "THC", out numThc, "b/l Fee", out numBl);
            ThreeNum(s, ref y, "Seal", out numSeal, "Telex Release", out numTelex, "CFS", out numCfs);
            ThreeNum(s, ref y, "VGM", out numVgm, "ENS/EBS/AMS", out numEns, "Others", out numOther);
            ThreeNum(s, ref y, "C/O Fee", out numCoFee, "TOTAL (VND)", out numTotalVND, "Sub-Total Ocean USD", out numSubOcean, dec0: true);

            SectionHdr(s, "Invoice Status", ref y);
            s.Controls.Add(Lbl("Red Invoice #", 0, y));
            txtRedInvNo = new TextBox { Location = new Point(160, y), Size = new Size(220, 22) };
            s.Controls.Add(txtRedInvNo); y += 30;
            DateRow(s, ref y, "Red Invoice Date", out chkRedInvDate, out dtpRedInvDate);
            DateRow(s, ref y, "Red Inv Received Date", out chkRedInvRecv, out dtpRedInvRecv);
            DateRow(s, ref y, "Transfer to Accountant", out chkTransferAcct, out dtpTransferAcct);

            var note = new Label
            {
                Text = "⚠  Level-3 staff: price fields are locked once a value has been saved.",
                Location = new Point(0, y + 4),
                AutoSize = true,
                ForeColor = Color.DarkOrange,
                Font = new Font("Segoe UI", 8f, FontStyle.Bold)
            };
            s.Controls.Add(note);
        }

        private void BuildCustomsTab(TabPage p)
        {
            var s = Scroll(p);
            int y = 0;

            SectionHdr(s, "Customs Charges", ref y);

            ThreeNum(s, ref y, "Trucking", out numTruck,
                               "Infrastructure Fee", out numInfra,
                               "Customs Clearance", out numCustClear, dec0: true);
            ThreeNum(s, ref y, "Custom Fee", out numCustFee,
                               "Other Custom", out numOtherCust,
                               "Sub-Total VND", out numSubVNDCust, dec0: true);
            ThreeNum(s, ref y, "Sub-Total USD", out numSubUSDCust,
                               "GRAND TOTAL VND", out numGrandVND,
                               "GRAND TOTAL USD", out numGrandUSD, dec0: true);

            // ── Row: CDS No | CDS Date ────────────────────────────────────
            s.Controls.Add(Lbl("CDS No", 0, y));
            txtCdsNo = new TextBox { Location = new Point(120, y), Size = new Size(160, 22) };
            s.Controls.Add(txtCdsNo);

            s.Controls.Add(Lbl("CDS Date", 300, y));
            chkCdsDate = new CheckBox
            { Text = "Set", Location = new Point(390, y), AutoSize = true };
            dtpCdsDate = new DateTimePicker
            {
                Location = new Point(430, y),
                Width = 130,
                Format = DateTimePickerFormat.Short,
                Enabled = false
            };
            chkCdsDate.CheckedChanged += (sv, ev) => dtpCdsDate.Enabled = chkCdsDate.Checked;
            s.Controls.AddRange(new Control[] { txtCdsNo, chkCdsDate, dtpCdsDate });
            y += 30;

            // ── Row: Luong/Line | Mã loại hình tk ────────────────────────
            TwoCol(s, ref y, "Luong/Line", out txtLine, "Mã loại hình tk", out txtCustomType);

            // ── Row: CDS Approve ─────────────────────────────────────────
            s.Controls.Add(Lbl("CDS Approve", 0, y));
            chkCdsApproved = new CheckBox
            { Text = "Approved", Location = new Point(120, y), AutoSize = true };
            s.Controls.Add(chkCdsApproved);
        }

        // ════════════════════════════════════════════════════════
        // Populate from existing model
        // ════════════════════════════════════════════════════════
        private void Populate(InvoiceModel m)
        {
            // ERP
            txtCustCode.Text = m.Invoice_customerCode ?? "";
            txtCustName.Text = m.Invoice_customerName ?? "";
            SetDateCtrl(chkCustReqDate, dtpCustReqDate, m.Invoice_customerRequestDate);
            txtPoNo.Text = m.Invoice_poNo ?? "";
            txtBrand.Text = m.Invoice_brand ?? "";
            SetDateCtrl(chkPoDate, dtpPoDate, m.Invoice_poDate);
            txtSaleName.Text = m.Invoice_saleName ?? "";
            txtFactoryNo.Text = m.Invoice_factoryNo ?? "";
            txtFactoryName.Text = m.Invoice_factoryName ?? "";
            txtItemCode.Text = m.Invoice_itemCode ?? "";
            txtItemCodeCust.Text = m.Invoice_itemCodeCustomers ?? "";
            txtItemName.Text = m.Invoice_itemName ?? "";
            if (m.Invoice_quantity.HasValue) numQty.Value = m.Invoice_quantity.Value;
            txtUnit.Text = m.Invoice_unit ?? "";

            // Plan
            txtNo.Text = m.Invoice_no ?? "";
            txtShipTerm.Text = m.Invoice_shippingTerm ?? "";
            txtPayTerm.Text = m.Invoice_paymentTerm ?? "";
            txtPIC.Text = m.Invoice_logisticPersonInCharge ?? "";
            txtNote.Text = m.Invoice_logisticNote ?? "";
            SetDateCtrl(chkConfirm, dtpConfirm, m.Invoice_factoryConfirmDate);
            cmbStatus.SelectedIndex = Math.Max(0, Math.Min(4, m.Invoice_shippingStatus ?? 0));
            txtFwdName.Text = m.Invoice_fwdName ?? "";
            txtBookNo.Text = m.Invoice_bookingNo ?? "";
            txtContType.Text = m.Invoice_contType ?? "";
            txtVgmCO.Text = m.Invoice_vgmCO ?? "";
            txtCyCO.Text = m.Invoice_cyCO ?? "";
            SetDateCtrl(chkEtd, dtpEtd, m.Invoice_etd);
            SetDateCtrl(chkEta, dtpEta, m.Invoice_eta);
            txtBillType.Text = m.Invoice_billType ?? "";
            txtBillNo.Text = m.Invoice_billNo ?? "";
            chkCo.Checked = m.Invoice_co ?? false;
            txtCoNo.Text = m.Invoice_coNo ?? "";
            cmbFeeStatus.SelectedIndex = Math.Max(0, (m.Invoice_feeStatus ?? 3) - 1);

            // Charges
            SD(numOF, m.Invoice_OF); SD(numDel, m.Invoice_deliveryCharges);
            SD(numTaxes, m.Invoice_taxes); SD(numOtherDest, m.Invoice_otherDestCharges);
            SD(numThc, m.Invoice_thc); SD(numBl, m.Invoice_blFee);
            SD(numSeal, m.Invoice_seal); SD(numTelex, m.Invoice_telexRelease);
            SD(numCfs, m.Invoice_cfs); SD(numVgm, m.Invoice_vgmFee);
            SD(numEns, m.Invoice_ensebsams); SD(numOther, m.Invoice_other);
            SD(numCoFee, m.Invoice_coFee); SD(numTotalVND, m.Invoice_totalVND);
            SD(numSubOcean, m.Invoice_subTotalOcean);
            txtRedInvNo.Text = m.Invoice_redInvoiceNo ?? "";
            SetDateCtrl(chkRedInvDate, dtpRedInvDate, m.Invoice_redInvoiceDate);
            SetDateCtrl(chkRedInvRecv, dtpRedInvRecv, m.Invoice_redInvoiceRecvDate);
            SetDateCtrl(chkTransferAcct, dtpTransferAcct, m.Invoice_transferAccountantDate);

            // Customs
            SD(numTruck, m.Invoice_trucking); SD(numInfra, m.Invoice_infrastructureFee);
            SD(numCustClear, m.Invoice_customerClearance); SD(numCustFee, m.Invoice_customFee);
            SD(numOtherCust, m.Invoice_otherCustomFee); SD(numSubVNDCust, m.Invoice_subTotalVNDCustom);
            SD(numSubUSDCust, m.Invoice_subTotalUSDCustom); SD(numGrandVND, m.Invoice_grandTotalVND);
            SD(numGrandUSD, m.Invoice_grandTotalUSD);
            txtCdsNo.Text = m.Invoice_cdsNo ?? "";
            SetDateCtrl(chkCdsDate, dtpCdsDate, m.Invoice_cdsDate);
            chkCdsApproved.Checked = m.Invoice_cdsApproved ?? false;
            txtLine.Text = m.Invoice_line ?? "";
            txtCustomType.Text = m.Invoice_customType ?? "";
        }

        // ════════════════════════════════════════════════════════
        // Apply scope permissions
        // ════════════════════════════════════════════════════════
        private void ApplyPermissions()
        {
            bool erpEditable = _scope == InvoiceEditScope.All;
            bool planEditable = _scope != InvoiceEditScope.None;
            bool etdExpired = _scope == InvoiceEditScope.LevelThree &&
                                  (_existing?.Invoice_etd.HasValue == true) &&
                                  _existing.Invoice_etd.Value.Date < DateTime.Now.AddDays(-7).Date;

            // ERP tab fields – only level 1
            foreach (Control c in new Control[]
            {
                txtCustCode, txtCustName, dtpCustReqDate, chkCustReqDate,
                txtPoNo, dtpPoDate, chkPoDate, txtSaleName,
                txtFactoryNo, txtFactoryName, txtItemCode, txtItemCodeCust,
                txtItemName, numQty, txtUnit
            })
                if (c != null) c.Enabled = erpEditable;

            // If ETD expired for level 3: lock everything
            if (etdExpired)
            {
                btnSave.Enabled = false;
                btnSave.Text = "⛔ Locked (ETD > 7 days)";
                return;
            }

            // Plan + customs fields
            foreach (Control c in new Control[]
            {
                txtNo, txtShipTerm, txtPayTerm, txtNote, txtFwdName, txtBookNo,
                txtContType, txtVgmCO, txtCyCO, txtBillType, txtBillNo, txtCoNo,
                chkCo, cmbStatus, cmbFeeStatus, dtpConfirm, chkConfirm,
                dtpEtd, chkEtd, dtpEta, chkEta,
                txtRedInvNo, dtpRedInvDate, chkRedInvDate, dtpRedInvRecv, chkRedInvRecv,
                dtpTransferAcct, chkTransferAcct,
                txtCdsNo, txtLine, txtCustomType, chkCdsDate, dtpCdsDate, chkCdsApproved
            })
                if (c != null) c.Enabled = planEditable;

            // PIC field: read-only for level 2 and 3 (auto-set on import)
            if (txtPIC != null) txtPIC.ReadOnly = _scope != InvoiceEditScope.All;

            // Price fields: for level 3 – lock if already has a non-zero value
            LockPriceIfSet(numOF, _existing?.Invoice_OF);
            LockPriceIfSet(numDel, _existing?.Invoice_deliveryCharges);
            LockPriceIfSet(numTaxes, _existing?.Invoice_taxes);
            LockPriceIfSet(numOtherDest, _existing?.Invoice_otherDestCharges);
            LockPriceIfSet(numThc, _existing?.Invoice_thc);
            LockPriceIfSet(numBl, _existing?.Invoice_blFee);
            LockPriceIfSet(numSeal, _existing?.Invoice_seal);
            LockPriceIfSet(numTelex, _existing?.Invoice_telexRelease);
            LockPriceIfSet(numCfs, _existing?.Invoice_cfs);
            LockPriceIfSet(numVgm, _existing?.Invoice_vgmFee);
            LockPriceIfSet(numEns, _existing?.Invoice_ensebsams);
            LockPriceIfSet(numOther, _existing?.Invoice_other);
            LockPriceIfSet(numCoFee, _existing?.Invoice_coFee);
            LockPriceIfSet(numTotalVND, _existing?.Invoice_totalVND);
            LockPriceIfSet(numSubOcean, _existing?.Invoice_subTotalOcean);
            LockPriceIfSet(numTruck, _existing?.Invoice_trucking);
            LockPriceIfSet(numInfra, _existing?.Invoice_infrastructureFee);
            LockPriceIfSet(numCustClear, _existing?.Invoice_customerClearance);
            LockPriceIfSet(numCustFee, _existing?.Invoice_customFee);
            LockPriceIfSet(numOtherCust, _existing?.Invoice_otherCustomFee);
            LockPriceIfSet(numSubVNDCust, _existing?.Invoice_subTotalVNDCustom);
            LockPriceIfSet(numSubUSDCust, _existing?.Invoice_subTotalUSDCustom);
            LockPriceIfSet(numGrandVND, _existing?.Invoice_grandTotalVND);
            LockPriceIfSet(numGrandUSD, _existing?.Invoice_grandTotalUSD);
        }

        private void LockPriceIfSet(NumericUpDown nud, decimal? existingVal)
        {
            if (nud == null) return;
            // For level 3 only: lock if already has a non-zero value in DB
            if (_scope == InvoiceEditScope.LevelThree && existingVal.HasValue && existingVal.Value != 0)
            {
                nud.Enabled = false;
                nud.BackColor = Color.FromArgb(240, 240, 240);
            }
        }

        // ════════════════════════════════════════════════════════
        // Save
        // ════════════════════════════════════════════════════════
        private void BtnSave_Click(object sender, EventArgs e)
        {
            lblErr.Text = "";
            if (string.IsNullOrWhiteSpace(txtPoNo.Text) && string.IsNullOrWhiteSpace(txtNo.Text))
            { lblErr.Text = "PO No or Invoice No is required."; return; }

            Result = new InvoiceModel
            {
                Invoice_Id = _isNew ? Guid.NewGuid().ToString() : _existingId,
                Invoice_no = txtNo.Text.Trim(),
                Invoice_customerCode = txtCustCode.Text.Trim(),
                Invoice_customerName = txtCustName.Text.Trim(),
                Invoice_customerRequestDate = chkCustReqDate.Checked ? dtpCustReqDate.Value : (DateTime?)null,
                Invoice_poNo = txtPoNo.Text.Trim(),
                Invoice_brand = txtBrand.Text.Trim(),
                Invoice_poDate = chkPoDate.Checked ? dtpPoDate.Value : (DateTime?)null,
                Invoice_saleName = txtSaleName.Text.Trim(),
                Invoice_factoryNo = txtFactoryNo.Text.Trim(),
                Invoice_factoryName = txtFactoryName.Text.Trim(),
                Invoice_itemCode = txtItemCode.Text.Trim(),
                Invoice_itemCodeCustomers = txtItemCodeCust.Text.Trim(),
                Invoice_itemName = txtItemName.Text.Trim(),
                Invoice_quantity = numQty.Value > 0 ? numQty.Value : (decimal?)null,
                Invoice_unit = txtUnit.Text.Trim(),
                Invoice_shippingTerm = txtShipTerm.Text.Trim(),
                Invoice_paymentTerm = txtPayTerm.Text.Trim(),
                Invoice_logisticPersonInCharge = txtPIC.Text.Trim(),
                Invoice_logisticNote = txtNote.Text.Trim(),
                Invoice_factoryConfirmDate = chkConfirm.Checked ? dtpConfirm.Value : (DateTime?)null,
                Invoice_shippingStatus = cmbStatus.SelectedIndex,
                Invoice_fwdName = txtFwdName.Text.Trim(),
                Invoice_bookingNo = txtBookNo.Text.Trim(),
                Invoice_contType = txtContType.Text.Trim(),
                Invoice_vgmCO = txtVgmCO.Text.Trim(),
                Invoice_cyCO = txtCyCO.Text.Trim(),
                Invoice_etd = chkEtd.Checked ? dtpEtd.Value : (DateTime?)null,
                Invoice_eta = chkEta.Checked ? dtpEta.Value : (DateTime?)null,
                Invoice_billType = txtBillType.Text.Trim(),
                Invoice_billNo = txtBillNo.Text.Trim(),
                Invoice_co = chkCo.Checked,
                Invoice_coNo = txtCoNo.Text.Trim(),
                Invoice_OF = GD(numOF),
                Invoice_deliveryCharges = GD(numDel),
                Invoice_taxes = GD(numTaxes),
                Invoice_otherDestCharges = GD(numOtherDest),
                Invoice_thc = GD(numThc),
                Invoice_blFee = GD(numBl),
                Invoice_seal = GD(numSeal),
                Invoice_telexRelease = GD(numTelex),
                Invoice_cfs = GD(numCfs),
                Invoice_vgmFee = GD(numVgm),
                Invoice_ensebsams = GD(numEns),
                Invoice_other = GD(numOther),
                Invoice_coFee = GD(numCoFee),
                Invoice_totalVND = GD(numTotalVND),
                Invoice_subTotalOcean = GD(numSubOcean),
                Invoice_feeStatus = cmbFeeStatus.SelectedIndex + 1,
                Invoice_redInvoiceNo = txtRedInvNo.Text.Trim(),
                Invoice_redInvoiceDate = chkRedInvDate.Checked ? dtpRedInvDate.Value : (DateTime?)null,
                Invoice_redInvoiceRecvDate = chkRedInvRecv.Checked ? dtpRedInvRecv.Value : (DateTime?)null,
                Invoice_transferAccountantDate = chkTransferAcct.Checked ? dtpTransferAcct.Value : (DateTime?)null,
                Invoice_trucking = GD(numTruck),
                Invoice_infrastructureFee = GD(numInfra),
                Invoice_customerClearance = GD(numCustClear),
                Invoice_customFee = GD(numCustFee),
                Invoice_otherCustomFee = GD(numOtherCust),
                Invoice_subTotalVNDCustom = GD(numSubVNDCust),
                Invoice_subTotalUSDCustom = GD(numSubUSDCust),
                Invoice_grandTotalVND = GD(numGrandVND),
                Invoice_grandTotalUSD = GD(numGrandUSD),
                Invoice_cdsNo = txtCdsNo.Text.Trim(),
                Invoice_cdsDate = chkCdsDate.Checked ? dtpCdsDate.Value : (DateTime?)null,
                Invoice_cdsApproved = chkCdsApproved.Checked,
                Invoice_line = txtLine.Text.Trim(),
                Invoice_customType = txtCustomType.Text.Trim(),
                createdate = _existing?.createdate ?? DateTime.Now,
                createby = _existing?.createby ?? _byUser,
                updatedate = DateTime.Now,
                updateby = _byUser
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        // ════════════════════════════════════════════════════════
        // UI builder helpers
        // ════════════════════════════════════════════════════════
        private static Panel Scroll(TabPage p)
        {
            var s = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10, 8, 10, 0) };
            p.Controls.Add(s);
            return s;
        }

        private static void SectionHdr(Panel p, string text, ref int y)
        {
            p.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(0, y),
                Size = new Size(760, 22),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = Color.FromArgb(30, 40, 60),
                ForeColor = Color.White,
                Padding = new Padding(6, 0, 0, 0)
            });
            y += 26;
        }

        private void TwoCol(Panel p, ref int y, string l1, out TextBox t1, string l2, out TextBox t2)
        {
            p.Controls.Add(Lbl(l1, 0, y));
            t1 = new TextBox { Location = new Point(160, y), Size = new Size(200, 22) };
            p.Controls.Add(t1);
            p.Controls.Add(Lbl(l2, 380, y));
            t2 = new TextBox { Location = new Point(540, y), Size = new Size(200, 22) };
            p.Controls.Add(t2);
            y += 30;
        }

        private void DateRow(Panel p, ref int y, string label,
    out CheckBox chk, out DateTimePicker dtp)
        {
            var c = new CheckBox { Text = "Set", Location = new Point(160, y), AutoSize = true };
            var d = new DateTimePicker
            {
                Location = new Point(200, y),
                Width = 120,
                Format = DateTimePickerFormat.Short,
                Enabled = false
            };
            c.CheckedChanged += (s, e) => d.Enabled = c.Checked;   // local capture, safe

            p.Controls.Add(Lbl(label, 0, y));
            p.Controls.AddRange(new Control[] { c, d });
            y += 30;

            chk = c; dtp = d;
        }

        private void DateRow2(Panel p, ref int y,
    string l1, out CheckBox chk1, out DateTimePicker dtp1,
    string l2, out CheckBox chk2, out DateTimePicker dtp2)
        {
            // Assign to locals first – lambdas cannot close over out/ref parameters
            var c1 = new CheckBox { Text = "Set", Location = new Point(160, y), AutoSize = true };
            var d1 = new DateTimePicker
            {
                Location = new Point(200, y),
                Width = 120,
                Format = DateTimePickerFormat.Short,
                Enabled = false
            };
            c1.CheckedChanged += (s, e) => d1.Enabled = c1.Checked;   // captures locals, not out params

            var c2 = new CheckBox { Text = "Set", Location = new Point(540, y), AutoSize = true };
            var d2 = new DateTimePicker
            {
                Location = new Point(580, y),
                Width = 120,
                Format = DateTimePickerFormat.Short,
                Enabled = false
            };
            c2.CheckedChanged += (s, e) => d2.Enabled = c2.Checked;

            p.Controls.Add(Lbl(l1, 0, y));
            p.Controls.Add(Lbl(l2, 380, y));
            p.Controls.AddRange(new Control[] { c1, d1, c2, d2 });
            y += 30;

            // Now assign out parameters from locals
            chk1 = c1; dtp1 = d1;
            chk2 = c2; dtp2 = d2;
        }

        private void ThreeNum(Panel p, ref int y,
            string l1, out NumericUpDown n1,
            string l2, out NumericUpDown n2,
            string l3, out NumericUpDown n3, bool dec0 = false)
        {
            int dec = dec0 ? 0 : 4;
            p.Controls.Add(Lbl(l1, 0, y)); n1 = NUD(p, 120, y, dec);
            p.Controls.Add(Lbl(l2, 260, y)); n2 = NUD(p, 380, y, dec);
            p.Controls.Add(Lbl(l3, 520, y)); n3 = NUD(p, 640, y, dec);
            y += 30;
        }

        private static NumericUpDown NUD(Panel p, int x, int y, int dec)
        {
            var n = new NumericUpDown
            {
                Location = new Point(x, y),
                Size = new Size(130, 22),
                Minimum = 0,
                Maximum = 999999999M,
                DecimalPlaces = dec
            };
            p.Controls.Add(n);
            return n;
        }

        private static Button BtnCtrl(string t, Color bg, int x)
        {
            var b = new Button
            {
                Text = t,
                Size = new Size(110, 32),
                Location = new Point(x, 7),
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private static Label Lbl(string t, int x, int y) =>
            new Label
            {
                Text = t,
                Location = new Point(x, y + 3),
                AutoSize = true,
                ForeColor = Color.FromArgb(50, 50, 80),
                Font = new Font("Segoe UI", 8f)
            };

        private static void SetDateCtrl(CheckBox chk, DateTimePicker dtp, DateTime? val)
        {
            if (val.HasValue) { chk.Checked = true; dtp.Value = val.Value; }
            else chk.Checked = false;
        }

        private static void SD(NumericUpDown n, decimal? v)
        { if (n != null && v.HasValue) n.Value = Math.Min(v.Value, n.Maximum); }

        private static decimal? GD(NumericUpDown n)
        { return n?.Value > 0 ? n.Value : (decimal?)null; }
    }
}