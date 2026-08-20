using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    public partial class SayfaPlatformlar : UserControl
    {
        private TextBox txtYeniPlatform;
        private Button btnEkle;
        private FlowLayoutPanel flpPlatformlar;

        public SayfaPlatformlar()
        {
            InitializeComponent();
            this.Load += SayfaPlatformlar_Load;

            // Grafiksel yırtılmaları önler, pürüzsüz animasyon sağlar
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.ResizeRedraw, true);
            this.UpdateStyles();
        }

        private void SayfaPlatformlar_Load(object sender, EventArgs e)
        {
            VeritabaniniHazirla();
            ArayuzInsaEt();
            ListeyiGuncelle();
        }

        private void VeritabaniniHazirla()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"CREATE TABLE IF NOT EXISTS Platformlar (Ad TEXT UNIQUE);";
                    using (var cmd = new SqliteCommand(sql, conn)) { cmd.ExecuteNonQuery(); }

                    try { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN AktifMi INTEGER DEFAULT 0;", conn)) { cmd.ExecuteNonQuery(); } } catch { }
                    try { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN PlatformID TEXT;", conn)) { cmd.ExecuteNonQuery(); } } catch { }

                    // EFETÜRK buraya eklendi!
                    string[] adlar = { "TRENDYOL", "HEPSİBURADA", "N11", "PTT AVM", "TURKCELL PASAJ", "PAZARAMA", "ÇİÇEK SEPETİ", "TEKNOSA", "İDEFİX", "MEDİAMARKT", "AMAZON", "SAHİBİNDEN", "TEMU", "EFETÜRK" };
                    string[] idler = { "TY", "HB", "N11", "PTT", "PASAJ", "PZR", "CS", "TKN", "IDF", "MM", "AMZ", "SHB", "TMU", "EFE" };

                    for (int i = 0; i < adlar.Length; i++)
                    {
                        using (var cmdCheck = new SqliteCommand("SELECT COUNT(*) FROM Platformlar WHERE Ad = @ad", conn))
                        {
                            cmdCheck.Parameters.AddWithValue("@ad", adlar[i]);
                            if ((long)cmdCheck.ExecuteScalar() == 0)
                            {
                                using (var cmdIns = new SqliteCommand("INSERT INTO Platformlar (Ad, PlatformID, AktifMi) VALUES (@ad, @pid, 0)", conn))
                                {
                                    cmdIns.Parameters.AddWithValue("@ad", adlar[i]);
                                    cmdIns.Parameters.AddWithValue("@pid", idler[i]);
                                    cmdIns.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                using (var cmdUpd = new SqliteCommand("UPDATE Platformlar SET PlatformID = @pid WHERE Ad = @ad AND (PlatformID IS NULL OR PlatformID = '')", conn))
                                {
                                    cmdUpd.Parameters.AddWithValue("@ad", adlar[i]);
                                    cmdUpd.Parameters.AddWithValue("@pid", idler[i]);
                                    cmdUpd.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Veritabanı yapılandırma hatası: " + ex.Message, "Kritik Hata", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.Dock = DockStyle.Fill;
            // Çok derin, elit bir lacivert/antrasit tonu
            this.BackColor = Color.FromArgb(11, 15, 25);

            // --- 1. ÜST HEADER ---
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 100, BackColor = Color.FromArgb(11, 15, 25), Padding = new Padding(40, 30, 40, 0) };
            Label lblT = new Label
            {
                Text = "Sistem ve Platform Ayarları",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 18F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Label lblAltT = new Label
            {
                Text = "Satış kanallarının entegrasyon durumlarını tek bir merkezden yönetin.",
                Dock = DockStyle.Bottom,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(148, 163, 184),
                Height = 35,
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlHeader.Controls.Add(lblT);
            pnlHeader.Controls.Add(lblAltT);
            this.Controls.Add(pnlHeader);

            // --- 2. YENİ PLATFORM EKLEME (MİNİMALİST BAR) ---
            Panel pnlEkleWrapper = new Panel { Dock = DockStyle.Top, Height = 90, Padding = new Padding(40, 10, 40, 20) };

            Panel pnlEkleBox = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(26, 29, 41),
                Padding = new Padding(20, 10, 20, 10)
            };
            // Çok ince, lüks bir kenarlık (border)
            pnlEkleBox.Paint += (s, pe) => { ControlPaint.DrawBorder(pe.Graphics, pnlEkleBox.ClientRectangle, Color.FromArgb(45, 50, 65), ButtonBorderStyle.Solid); };

            Label lblEkle = new Label { Text = "Yeni Satış Kanalı:", Location = new Point(20, 19), AutoSize = true, Font = new Font("Segoe UI", 10.5F, FontStyle.Bold), ForeColor = Color.FromArgb(226, 232, 240) };

            Panel pnlInputBg = new Panel { Location = new Point(170, 15), Size = new Size(350, 32), BackColor = Color.FromArgb(18, 20, 28) };
            txtYeniPlatform = new TextBox
            {
                Location = new Point(10, 6),
                Width = 330,
                Font = new Font("Segoe UI", 11F),
                BackColor = Color.FromArgb(18, 20, 28),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.None
            };
            pnlInputBg.Controls.Add(txtYeniPlatform);
            pnlInputBg.Paint += (s, pe) => { ControlPaint.DrawBorder(pe.Graphics, pnlInputBg.ClientRectangle, Color.FromArgb(60, 65, 80), ButtonBorderStyle.Solid); };

            btnEkle = new Button
            {
                Text = "Sisteme Tanımla",
                Location = new Point(540, 14),
                Width = 150,
                Height = 34,
                BackColor = Color.FromArgb(37, 99, 235), // Güven veren kurumsal mavi
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnEkle.FlatAppearance.BorderSize = 0;
            btnEkle.Click += BtnEkle_Click;

            pnlEkleBox.Controls.Add(lblEkle);
            pnlEkleBox.Controls.Add(pnlInputBg);
            pnlEkleBox.Controls.Add(btnEkle);
            pnlEkleWrapper.Controls.Add(pnlEkleBox);
            this.Controls.Add(pnlEkleWrapper);

            // --- 3. ALT ALTA NİZAMLI LİSTE TASARIMI ---
            flpPlatformlar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(11, 15, 25),
                Padding = new Padding(40, 10, 40, 40),
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown, // SADECE ALT ALTA DİZİLİR
                WrapContents = false // Asla yan yana geçmez
            };

            // Satırların ekran genişliğine göre uzaması için
            flpPlatformlar.Resize += (s, e) => {
                flpPlatformlar.SuspendLayout();
                foreach (Control c in flpPlatformlar.Controls)
                {
                    // Scrollbar payını (25px) düşerek genişliği tam ekran yap
                    c.Width = flpPlatformlar.ClientSize.Width - 80 - 25;
                }
                flpPlatformlar.ResumeLayout();
            };

            this.Controls.Add(flpPlatformlar);
            flpPlatformlar.BringToFront();
            pnlEkleWrapper.SendToBack();
            pnlHeader.SendToBack();
        }

        private void ListeyiGuncelle()
        {
            try
            {
                flpPlatformlar.Controls.Clear();
                int rowWidth = flpPlatformlar.ClientSize.Width > 100 ? flpPlatformlar.ClientSize.Width - 80 - 25 : 800;

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT rowid, PlatformID, Ad, AktifMi FROM Platformlar ORDER BY Ad ASC", conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                long id = Convert.ToInt64(reader["rowid"]);
                                string pId = reader["PlatformID"].ToString();
                                string ad = reader["Ad"].ToString();
                                bool isAktif = Convert.ToInt32(reader["AktifMi"]) == 1;

                                SatirOlustur(id, pId, ad, isAktif, rowWidth);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Liste Hatası: " + ex.Message); }
        }

        // 🚀 MÜHÜR: TABLO YERİNE, WEB AYARLAR SAYFASI GİBİ ALT ALTA UZAYAN ZARİF SATIRLAR
        private void SatirOlustur(long id, string pid, string ad, bool isAktif, int w)
        {
            Color accentColor = isAktif ? Color.FromArgb(16, 185, 129) : Color.FromArgb(71, 85, 105);

            Panel pnlRow = new Panel
            {
                Width = w,
                Height = 70, // Daha kibar ve ince satır yapısı
                BackColor = Color.FromArgb(22, 25, 37),
                Margin = new Padding(0, 0, 0, 8) // Alt alta boşluk
            };

            // Sol tarafta duruma göre renklenen modern indikatör
            Panel pnlIndicator = new Panel
            {
                Location = new Point(0, 0),
                Width = 4,
                Height = 70,
                BackColor = accentColor
            };

            // Platform Adı
            Label lblName = new Label
            {
                Text = ad,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(25, 15),
                AutoSize = true
            };

            // Platform ID 
            Label lblId = new Label
            {
                Text = $"Sistem Kimliği: {pid}",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                ForeColor = Color.FromArgb(148, 163, 184),
                Location = new Point(27, 40),
                AutoSize = true
            };

            // Sağ Taraf Durum Yazısı
            Label lblStatusText = new Label
            {
                Text = isAktif ? "Sistemde Aktif" : "Devre Dışı",
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                ForeColor = accentColor,
                AutoSize = true,
                Anchor = AnchorStyles.Right | AnchorStyles.Top
            };

            // iPhone Stili Aç/Kapat Anahtarı
            ModernToggle toggle = new ModernToggle
            {
                Width = 48,
                Height = 24,
                Checked = isAktif,
                Tag = id,
                Anchor = AnchorStyles.Right | AnchorStyles.Top
            };

            // Anahtarı ve Durum Yazısını Sağ Tarafa Yaslama Matematiği
            toggle.Location = new Point(pnlRow.Width - toggle.Width - 20, (pnlRow.Height - toggle.Height) / 2);
            lblStatusText.Location = new Point(toggle.Left - 110, (pnlRow.Height - lblStatusText.Height) / 2);

            toggle.CheckedChanged += (s, ev) => {
                bool yeniDurum = toggle.Checked;
                DurumGuncelle(Convert.ToInt64(toggle.Tag), yeniDurum ? 1 : 0);

                // Animasyon hissi: Tıklandığı an sayfa yenilenmeden renkler değişir
                Color newColor = yeniDurum ? Color.FromArgb(16, 185, 129) : Color.FromArgb(71, 85, 105);
                pnlIndicator.BackColor = newColor;
                lblStatusText.Text = yeniDurum ? "Sistemde Aktif" : "Devre Dışı";
                lblStatusText.ForeColor = newColor;

                // Durum yazısını sola doğru dinamik kaydır
                lblStatusText.Left = toggle.Left - lblStatusText.Width - 15;
            };

            pnlRow.Controls.Add(pnlIndicator);
            pnlRow.Controls.Add(lblName);
            pnlRow.Controls.Add(lblId);
            pnlRow.Controls.Add(lblStatusText);
            pnlRow.Controls.Add(toggle);

            flpPlatformlar.Controls.Add(pnlRow);
        }

        private void DurumGuncelle(long id, int yeniDurum)
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("UPDATE Platformlar SET AktifMi=@aktif WHERE rowid=@id", conn))
                    {
                        cmd.Parameters.AddWithValue("@aktif", yeniDurum);
                        cmd.Parameters.AddWithValue("@id", id);
                        cmd.ExecuteNonQuery();
                    }
                }

                if (Form1.AnaForm != null)
                {
                    Form1.AnaForm.MuhasebeMenusunuGuncelle();
                }
            }
            catch (Exception ex) { MessageBox.Show("Sistem Güncelleme Hatası: " + ex.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private void BtnEkle_Click(object sender, EventArgs e)
        {
            string yeniAd = txtYeniPlatform.Text.Trim().ToUpper(CultureInfo.InvariantCulture);
            if (string.IsNullOrEmpty(yeniAd)) { MessageBox.Show("Lütfen geçerli bir platform adı giriniz.", "Eksik Veri", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            string kimlikUret = new string(yeniAd.Replace(" ", "").Where(c => !"AEIİIOÖUÜ".Contains(c)).ToArray());
            if (kimlikUret.Length > 3) kimlikUret = kimlikUret.Substring(0, 3);
            else if (kimlikUret.Length == 0) kimlikUret = yeniAd.Substring(0, Math.Min(yeniAd.Length, 3));

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmdCheck = new SqliteCommand("SELECT COUNT(*) FROM Platformlar WHERE Ad=@ad OR PlatformID=@pid", conn))
                    {
                        cmdCheck.Parameters.AddWithValue("@ad", yeniAd);
                        cmdCheck.Parameters.AddWithValue("@pid", kimlikUret);
                        if ((long)cmdCheck.ExecuteScalar() > 0)
                        {
                            MessageBox.Show("Sistemde bu ada veya kimliğe sahip bir platform zaten bulunuyor.", "İşlem Reddedildi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    using (var cmdIns = new SqliteCommand("INSERT INTO Platformlar (Ad, PlatformID, AktifMi) VALUES (@ad, @pid, 0)", conn))
                    {
                        cmdIns.Parameters.AddWithValue("@ad", yeniAd);
                        cmdIns.Parameters.AddWithValue("@pid", kimlikUret);
                        cmdIns.ExecuteNonQuery();
                    }
                }
                txtYeniPlatform.Clear();
                ListeyiGuncelle();
                if (Form1.AnaForm != null) Form1.AnaForm.MuhasebeMenusunuGuncelle();
            }
            catch (Exception ex) { MessageBox.Show("Veritabanı yazma hatası: " + ex.Message, "Kritik Hata", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }

    // =========================================================================================
    // ⚙️ ÖZEL ÇİZİM: IPHONE STİLİ MODERN TOGGLE (ANAHTAR) KONTROLÜ
    // =========================================================================================
    public class ModernToggle : Control
    {
        private bool _checked = false;
        public bool Checked
        {
            get => _checked;
            set { _checked = value; Invalidate(); }
        }

        public event EventHandler CheckedChanged;

        public ModernToggle()
        {
            this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            this.Cursor = Cursors.Hand;
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            Checked = !Checked;
            CheckedChanged?.Invoke(this, EventArgs.Empty);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.Clear(this.Parent.BackColor);

            Color onColor = Color.FromArgb(16, 185, 129); // Zümrüt Yeşili
            Color offColor = Color.FromArgb(71, 85, 105); // Slate Gray
            Color toggleColor = Checked ? onColor : offColor;

            // Arka Plan Kapsülü
            Rectangle rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
            int radius = this.Height - 1;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius, radius, 90, 180);
            path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 180);
            path.CloseFigure();

            e.Graphics.FillPath(new SolidBrush(toggleColor), path);

            // Hareketli Top
            int circleSize = this.Height - 6;
            int circleX = Checked ? this.Width - circleSize - 3 : 3;
            e.Graphics.FillEllipse(Brushes.White, circleX, 3, circleSize, circleSize);
        }
    }
}