using System;
using System.Drawing;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories;
using techlink_workspace.View.Freight;
using techlink_workspace.View.Invoice;

namespace techlink_workspace.View
{
    public partial class Dashboard : Form
    {
        // ── Header ───────────────────────────────────────────────────────
        private Panel pnlHeader;
        private PictureBox picLogo;
        private Label lblAppTitle;
        private Panel pnlUserArea;
        private PictureBox picAvatar;
        private Label lblUserName;

        // Popup menu
        private Panel pnlPopup;
        private bool _popupVisible = false;

        // ── Layout ───────────────────────────────────────────────────────
        private Panel pnlSidebar;

        // ════════════════════════════════════════════════════════════════
        public Dashboard()
        {
            InitializeComponent();
            BuildHeader();
            BuildSidebar();
            BuildUserPopup();
            ApplySession();

            this.FormClosing += (s, e) =>
            {
                AppSession.Clear();
                new LoginForm().Show();
            };
        }

        // ════════════════════════════════════════════════════════════════
        // 1. Header
        // ════════════════════════════════════════════════════════════════
        private void BuildHeader()
        {
            this.Text = "Dashboard";
            this.MinimumSize = new Size(900, 600);
            this.ClientSize = new Size(1280, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 246, 250);

            pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 56,
                BackColor = Color.White,
                Padding = new Padding(16, 0, 16, 0)
            };
            pnlHeader.Paint += PnlHeader_Paint;

            picLogo = new PictureBox
            {
                Size = new Size(32, 32),
                Location = new Point(16, 12),
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand
            };
            try { picLogo.Image = Properties.Resources.CompanyLogo; } catch { }

            lblAppTitle = new Label
            {
                Text = "Tech-link Workspace",
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 33, 33),
                AutoSize = true,
                Location = new Point(56, 18),
                BackColor = Color.Transparent
            };

            // User area (right side)
            pnlUserArea = new Panel
            {
                Size = new Size(130, 36),
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand
            };
            picAvatar = new PictureBox
            {
                Size = new Size(32, 32),
                Location = new Point(0, 2),
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.FromArgb(38, 132, 255)
            };
            picAvatar.Region = System.Drawing.Region.FromHrgn(
                CreateRoundRectRgn(0, 0, 32, 32, 16, 16));

            lblUserName = new Label
            {
                Text = "User",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 33, 33),
                AutoSize = true,
                Location = new Point(38, 9),
                BackColor = Color.Transparent
            };

            pnlUserArea.Controls.AddRange(new Control[] { picAvatar, lblUserName });
            pnlUserArea.Click += (s, e) => TogglePopup();
            picAvatar.Click += (s, e) => TogglePopup();
            lblUserName.Click += (s, e) => TogglePopup();

            pnlHeader.Controls.AddRange(new Control[] { picLogo, lblAppTitle, pnlUserArea });
            this.Controls.Add(pnlHeader);

            this.Resize += (s, e) => PositionHeaderControls();
            pnlHeader.Resize += (s, e) => PositionHeaderControls();
            PositionHeaderControls();
        }

        private void PositionHeaderControls()
        {
            pnlUserArea.Location = new Point(pnlHeader.Width - 16 - pnlUserArea.Width, 10);
        }

        private void PnlHeader_Paint(object sender, PaintEventArgs e)
        {
            using (var pen = new Pen(Color.FromArgb(220, 220, 230), 1))
                e.Graphics.DrawLine(pen, 0, pnlHeader.Height - 1,
                                    pnlHeader.Width, pnlHeader.Height - 1);
        }

        // ════════════════════════════════════════════════════════════════
        // 2. Sidebar + Content area
        // ════════════════════════════════════════════════════════════════
        private void BuildSidebar()
        {
            pnlSidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 220,
                BackColor = Color.FromArgb(30, 40, 60)
            };

            // ── Sidebar items ─────────────────────────────────────────────
            // Each item: (label, icon-char, action)
            var items = new (string Label, string Icon, Action OnClick)[]
            {
                ("Dashboard",         "⊞", () => LoadContent(null)),
                ("Freight Quotation", "🚢", () => LoadContent(new FreightQuotationForm(AppSession.CurrentUser))),
                ("Invoice Input",     "🧾", () => LoadContent(new InvoiceForm(AppSession.CurrentUser))),
            };

            int y = 20;
            foreach (var item in items)
            {
                var btn = new Button
                {
                    Text = $"  {item.Icon}  {item.Label}",
                    Location = new Point(0, y),
                    Size = new Size(220, 44),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.Transparent,
                    ForeColor = Color.FromArgb(200, 210, 230),
                    Font = new Font("Segoe UI", 9.5f),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Cursor = Cursors.Hand,
                    Padding = new Padding(8, 0, 0, 0)
                };
                btn.FlatAppearance.BorderSize = 0;
                btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(50, 65, 90);
                var capturedAction = item.OnClick;
                btn.Click += (s, e) =>
                {
                    // Highlight active
                    foreach (Control c in pnlSidebar.Controls)
                        if (c is Button b)
                        {
                            b.BackColor = Color.Transparent;
                            b.ForeColor = Color.FromArgb(200, 210, 230);
                        }
                    btn.BackColor = Color.FromArgb(50, 65, 90);
                    btn.ForeColor = Color.White;
                    capturedAction();
                };
                pnlSidebar.Controls.Add(btn);
                y += 44;
            }

            // Add a bottom separator line on sidebar
            pnlSidebar.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(50, 65, 90), 1))
                    e.Graphics.DrawLine(pen, pnlSidebar.Width - 1, 0,
                                        pnlSidebar.Width - 1, pnlSidebar.Height);
            };

            // Controls — sidebar docks left, rest of window is empty dashboard area
            this.Controls.Add(pnlSidebar);
        }

        /// <summary>
        /// Opens a module form as a normal standalone window.
        /// Pass null to do nothing (Dashboard home — content area stays as-is).
        /// </summary>
        private void LoadContent(Form child)
        {
            if (child == null) return;   // Dashboard home — nothing to open
            child.StartPosition = FormStartPosition.CenterScreen;
            child.Show();
        }

        // ════════════════════════════════════════════════════════════════
        // 3. User popup menu (Profile / Change Password / Logout only)
        // ════════════════════════════════════════════════════════════════
        private void BuildUserPopup()
        {
            pnlPopup = new Panel
            {
                Size = new Size(200, 108),
                BackColor = Color.White,
                Visible = false,
                BorderStyle = BorderStyle.None
            };
            pnlPopup.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(220, 220, 230), 1))
                    e.Graphics.DrawRectangle(pen, 0, 0,
                        pnlPopup.Width - 1, pnlPopup.Height - 1);
            };

            string[] menuItems = { "  Profile", "  Change password", "  Logout" };
            int y = 8;
            for (int i = 0; i < menuItems.Length; i++)
            {
                var item = new Button
                {
                    Text = menuItems[i],
                    Size = new Size(200, 40),
                    Location = new Point(0, y),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(40, 40, 40),
                    Font = new Font("Segoe UI", 9.5f),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Cursor = Cursors.Hand,
                    Tag = i
                };
                item.FlatAppearance.BorderSize = 0;
                item.FlatAppearance.MouseOverBackColor = Color.FromArgb(245, 246, 250);
                item.Click += MenuItem_Click;
                pnlPopup.Controls.Add(item);
                y += 42;
            }

            // Separator before Logout
            pnlPopup.Controls.Add(new Panel
            {
                Location = new Point(12, 90),
                Size = new Size(176, 1),
                BackColor = Color.FromArgb(230, 230, 235)
            });

            this.Controls.Add(pnlPopup);
            pnlPopup.BringToFront();
            this.MouseClick += (s, e) => HidePopup();
        }

        private void TogglePopup()
        {
            _popupVisible = !_popupVisible;
            if (_popupVisible)
            {
                pnlPopup.Location = new Point(
                    pnlUserArea.Left + pnlHeader.Left - 70,
                    pnlHeader.Bottom + 4);
                pnlPopup.Visible = true;
                pnlPopup.BringToFront();
            }
            else pnlPopup.Visible = false;
        }

        private void HidePopup()
        {
            _popupVisible = false;
            pnlPopup.Visible = false;
        }

        private void MenuItem_Click(object sender, EventArgs e)
        {
            HidePopup();
            switch ((int)((Button)sender).Tag)
            {
                case 0: ShowProfile(); break;
                case 1: ShowChangePassword(); break;
                case 2: Logout(); break;
            }
        }

        // ════════════════════════════════════════════════════════════════
        // 4. Actions
        // ════════════════════════════════════════════════════════════════
        private void ShowProfile()
        {
            CTMessageBox.Show("Profile page coming soon.", "Profile",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowChangePassword()
        {
            using (var dlg = new Form_ChangePassword(AppSession.CurrentUser?.User_id))
                dlg.ShowDialog(this);
        }

        private void Logout()
        {
            if (CTMessageBox.Show("Are you sure you want to logout?",
                    "Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                == DialogResult.Yes)
            {
                AppSession.Clear();
                new LoginForm().Show();
                this.Close();
            }
        }

        private void ApplySession()
        {
            if (AppSession.CurrentUser != null)
                lblUserName.Text =
                    AppSession.CurrentUser.User_fullName?.Split(' ')[0] ?? "User";
        }

        // ── P/Invoke ─────────────────────────────────────────────────────
        [System.Runtime.InteropServices.DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);
    }
}