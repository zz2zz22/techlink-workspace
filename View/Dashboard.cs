using System;
using System.Drawing;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories;

namespace techlink_workspace.View
{
    public partial class Dashboard : Form
    {
        // ── Header controls ──────────────────────────────────
        private Panel pnlHeader;
        private PictureBox picLogo;
        private Label lblAppTitle;

        // Right-side header icons
        private Button btnSearch;
        private Button btnHelp;
        private Button btnNotification;
        private Label lblNotifBadge;
        private Panel pnlUserArea;
        private PictureBox picAvatar;
        private Label lblUserName;

        // Popup menu
        private Panel pnlPopup;
        private bool _popupVisible = false;

        public Dashboard()
        {
            InitializeComponent();
            BuildHeader();
            BuildUserPopup();
            ApplySession();
            this.FormClosing += (s, e) =>
            {
                AppSession.Clear();
                var login = new LoginForm();
                login.Show();
            };
        }

        // ════════════════════════════════════════════════════
        // 1. Header
        // ════════════════════════════════════════════════════
        private void BuildHeader()
        {
            this.Text = "Dashboard";
            this.MinimumSize = new Size(900, 600);
            this.ClientSize = new Size(1200, 700);
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

            // ── Logo / title ──────────────────────────────────
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

            // ── Right-side icon buttons ───────────────────────
            int iconY = 14;
            int rightStart = pnlHeader.Width - 16;   // anchored dynamically in Resize



            // Notification with badge
            btnNotification = MakeIconButton("&#128276;", "Notifications");
            lblNotifBadge = new Label
            {
                Text = "8",
                Font = new Font("Segoe UI", 7f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(220, 53, 69),
                Size = new Size(16, 16),
                TextAlign = ContentAlignment.MiddleCenter,
            };
            lblNotifBadge.Region = System.Drawing.Region.FromHrgn(
                CreateRoundRectRgn(0, 0, 16, 16, 8, 8));

            // User area (avatar + name)
            pnlUserArea = new Panel
            {
                Size = new Size(110, 36),
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
                Text = "John",
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

            pnlHeader.Controls.AddRange(new Control[]
            {
                picLogo, lblAppTitle,
                pnlUserArea
            });
            this.Controls.Add(pnlHeader);

            // Re-anchor right-side controls on resize
            this.Resize += (s, e) => PositionHeaderControls();
            pnlHeader.Resize += (s, e) => PositionHeaderControls();
            PositionHeaderControls();
        }

        private void PositionHeaderControls()
        {
            int right = pnlHeader.Width - 16;
            int iconY = 10;

            pnlUserArea.Location = new Point(right - pnlUserArea.Width, 10);
            right = pnlUserArea.Left - 8;
        }

        private Button MakeIconButton(string text, string tooltip)
        {
            var btn = new Button
            {
                Size = new Size(36, 36),
                Text = text,
                Font = new Font("Segoe UI", 13f),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(80, 80, 80),
                Cursor = Cursors.Hand,
                UseCompatibleTextRendering = true,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 241, 245);
            new ToolTip().SetToolTip(btn, tooltip);
            return btn;
        }

        private void PnlHeader_Paint(object sender, PaintEventArgs e)
        {
            // Bottom border line
            using (var pen = new Pen(Color.FromArgb(220, 220, 230), 1))
                e.Graphics.DrawLine(pen, 0, pnlHeader.Height - 1, pnlHeader.Width, pnlHeader.Height - 1);
        }

        // ════════════════════════════════════════════════════
        // 2. User popup menu
        // ════════════════════════════════════════════════════
        private void BuildUserPopup()
        {
            pnlPopup = new Panel
            {
                Size = new Size(200, 148),
                BackColor = Color.White,
                Visible = false,
                BorderStyle = BorderStyle.None,
                Padding = new Padding(0)
            };
            pnlPopup.Paint += PnlPopup_Paint;

            string[] menuItems = { "  Profile", "  Change password", "  Logout" };
            Color[] icons = {
                Color.FromArgb(38, 132, 255),
                Color.FromArgb(255, 152, 0),
                Color.FromArgb(220, 53, 69)
            };
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

            // Add separator before Logout
            var sep = new Panel
            {
                Location = new Point(12, 90),
                Size = new Size(176, 1),
                BackColor = Color.FromArgb(230, 230, 235)
            };
            pnlPopup.Controls.Add(sep);

            this.Controls.Add(pnlPopup);
            pnlPopup.BringToFront();

            // Close popup when clicking elsewhere
            this.MouseClick += (s, e) => HidePopup();
        }

        private void PnlPopup_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            using (var pen = new Pen(Color.FromArgb(220, 220, 230), 1))
                g.DrawRectangle(pen, 0, 0, pnlPopup.Width - 1, pnlPopup.Height - 1);
        }

        private void TogglePopup()
        {
            _popupVisible = !_popupVisible;
            if (_popupVisible)
            {
                int x = pnlUserArea.Left + pnlHeader.Left;
                int y = pnlHeader.Bottom;
                pnlPopup.Location = new Point(x - 90, y + 4);
                pnlPopup.Visible = true;
                pnlPopup.BringToFront();
            }
            else
            {
                pnlPopup.Visible = false;
            }
        }

        private void HidePopup()
        {
            _popupVisible = false;
            pnlPopup.Visible = false;
        }

        private void MenuItem_Click(object sender, EventArgs e)
        {
            HidePopup();
            int idx = (int)((Button)sender).Tag;
            switch (idx)
            {
                case 0: ShowProfile(); break;
                case 1: ShowChangePassword(); break;
                case 2: Logout(); break;
            }
        }

        // ════════════════════════════════════════════════════
        // 3. Actions
        // ════════════════════════════════════════════════════
        private void ShowProfile()
        {
            // Placeholder — replace with your Profile form
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
            var result = CTMessageBox.Show(
                "Are you sure you want to logout?",
                "Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                AppSession.Clear();
                var login = new LoginForm();
                login.Show();
                this.Close();
            }
        }

        private void ApplySession()
        {
            if (AppSession.CurrentUser != null)
                lblUserName.Text = AppSession.CurrentUser.User_fullName?.Split(' ')[0] ?? "User";
        }

        // ── P/Invoke for circular regions ────────────────────
        [System.Runtime.InteropServices.DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);
    }
}