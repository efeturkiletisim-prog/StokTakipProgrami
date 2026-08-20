using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // 🚚 ANLAŞMALI KARGO DESİ - FİYAT MATRİS YÖNETİMİ (INDEX-BASED GÜVENLİK VE PLATFORM KONTROLÜ)
    public partial class SayfaKargoMatrisi : UserControl
    {
        // Sol Panel Elemanları
        private ComboBox cmbPlatform, cmbKargoFirma;
        private TextBox txtYeniKargo;
        private Button btnYeniKargoEkle, btnTopluYukle, btnSil;

        // Sağ Panel Elemanları (Arama ve Filtreleme)
        private ComboBox cmbRightPlatform, cmbRightKargoFirma;
        private Button btnRightAra;
        private DataGridView dgvMatris;
        private ContextMenuStrip cms;

        // Tetikleyici Kalkanı
        private bool lYukleniyor = false;

        public SayfaKargoMatrisi()
        {
            InitializeComponent();
            this.Load += SayfaKargoMatrisi_Load;
        }

        private void SayfaKargoMatrisi_Load(object sender, EventArgs e)
        {
            TabloAltyapisiniGuvenceyeAl();
            ArayuzYarat();
            PlatformlariYukle();
            KargoFirmalariniYukle();
            MatrisiListele("", "");
        }

        private void TabloAltyapisiniGuvenceyeAl()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = @"CREATE TABLE IF NOT EXISTS KargoMatrisi (
                                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    PlatformAd TEXT, 
                                    KargoFirma TEXT, 
                                    Desi INTEGER, 
                                    Fiyat REAL,
                                    UNIQUE(PlatformAd, KargoFirma, Desi));";
                    using (var cmd = new SqliteCommand(sql, b)) { cmd.ExecuteNonQuery(); }
                }
            }
            catch { }
        }

        private void ArayuzYarat()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(33, 33, 36);

            // ==========================================
            // 1. ÜST BAŞLIK PANELİ (Z-INDEX KORUMALI)
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "🚚 ANLAŞMALI KARGO DESİ - FİYAT MATRİS YÖNETİMİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);

            // ==========================================
            // 2. ANA ORTA GÖVDE (MİZANPAJ IZGARASI)
            // ==========================================
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = Color.Transparent };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 330F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // SOL PANEL
            Panel pnlSol = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15) };
            Label lblVeriEkleBaslik = new Label { Text = "🟢 MATRİS TANIMLA", Location = new Point(15, 10), AutoSize = true, Font = new Font("Segoe UI Black", 14F, FontStyle.Bold), ForeColor = Color.White };
            pnlSol.Controls.Add(lblVeriEkleBaslik);

            pnlSol.Controls.Add(FormComboYarat("1. Platform / Pazaryeri Seçin:", out cmbPlatform, new Point(15, 50), 260));
            cmbPlatform.SelectedIndexChanged += CmbLeft_SelectedIndexChanged;

            pnlSol.Controls.Add(FormComboYarat("2. Anlaşmalı Kargo Firması:", out cmbKargoFirma, new Point(15, 115), 260));
            cmbKargoFirma.SelectedIndexChanged += CmbLeft_SelectedIndexChanged;

            Label lblYeniKrg = new Label { Text = "➕ Listede Olmayan Firma Ekle:", Location = new Point(15, 185), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            txtYeniKargo = new TextBox { Location = new Point(15, 205), Width = 150, Font = new Font("Segoe UI", 9.5F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            btnYeniKargoEkle = new Button { Text = "Firmayı Ekle", Location = new Point(175, 204), Width = 100, Height = 23, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnYeniKargoEkle.FlatAppearance.BorderSize = 0;
            btnYeniKargoEkle.Click += BtnYeniKargoEkle_Click;
            pnlSol.Controls.Add(lblYeniKrg); pnlSol.Controls.Add(txtYeniKargo); pnlSol.Controls.Add(btnYeniKargoEkle);

            Label lblTopluBaslik = new Label { Text = "⚡ TOPLU CSV AKTARIMI", Location = new Point(15, 255), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold | FontStyle.Underline), ForeColor = Color.FromArgb(56, 189, 248) };
            btnTopluYukle = new Button { Text = "📂 Toplu CSV Fiyat Listesi Yükle", Location = new Point(15, 280), Width = 260, Height = 40, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnTopluYukle.FlatAppearance.BorderSize = 0;
            btnTopluYukle.Click += BtnTopluYukle_Click;
            pnlSol.Controls.Add(lblTopluBaslik); pnlSol.Controls.Add(btnTopluYukle);

            btnSil = new Button { Text = "🗑️ Seçili Satırı Sistemden Sil", Location = new Point(15, 420), Width = 260, Height = 32, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnSil.FlatAppearance.BorderSize = 0;
            btnSil.Click += BtnSil_Click;
            pnlSol.Controls.Add(btnSil);

            mainLayout.Controls.Add(pnlSol, 0, 0);

            // SAĞ PANEL
            Panel pnlSagListe = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };

            Panel pnlRightFilter = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 28), Margin = new Padding(0, 0, 0, 10) };
            pnlRightFilter.Controls.Add(FormComboYarat("Platform Filtre:", out cmbRightPlatform, new Point(15, 5), 180));
            pnlRightFilter.Controls.Add(FormComboYarat("Kargo Firma Filtre:", out cmbRightKargoFirma, new Point(210, 5), 180));

            btnRightAra = new Button { Text = "🔍 FİLTRELE", Location = new Point(410, 20), Width = 120, Height = 26, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnRightAra.FlatAppearance.BorderSize = 0;
            btnRightAra.Click += (s, e) => MatrisiListele(cmbRightPlatform.Text, cmbRightKargoFirma.Text);
            pnlRightFilter.Controls.Add(btnRightAra);
            pnlSagListe.Controls.Add(pnlRightFilter);

            cms = new ContextMenuStrip();
            cms.Items.Add("✏️ Seçili Desi Fiyatını Düzenle", null, BtnBaremDuzenle_Click);
            cms.Items.Add("❌ Baremi Sil", null, BtnSil_Click);

            dgvMatris = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(24, 24, 27),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9.5F),
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 35,
                ContextMenuStrip = cms
            };
            dgvMatris.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgvMatris.ColumnHeadersDefaultCellStyle.ForeColor = Color.Orange;
            dgvMatris.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            dgvMatris.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgvMatris.DefaultCellStyle.ForeColor = Color.White;
            dgvMatris.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgvMatris.DataError += (s, e) => { e.Cancel = true; };

            pnlSagListe.Controls.Add(dgvMatris);
            dgvMatris.BringToFront();
            mainLayout.Controls.Add(pnlSagListe, 1, 0);

            this.Controls.Add(mainLayout);
            this.Controls.Add(pnlHeader);
            pnlHeader.SendToBack();
            mainLayout.BringToFront();
        }

        private void PlatformlariYukle()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("SELECT Ad FROM Platformlar ORDER BY Ad ASC", b))
                    using (var oku = cmd.ExecuteReader())
                    {
                        lYukleniyor = true;
                        cmbPlatform.Items.Clear(); cmbRightPlatform.Items.Clear();
                        cmbPlatform.Items.Add(""); cmbRightPlatform.Items.Add("");

                        while (oku.Read())
                        {
                            string ad = oku["Ad"].ToString();
                            cmbPlatform.Items.Add(ad);
                            cmbRightPlatform.Items.Add(ad);
                        }
                        cmbPlatform.SelectedIndex = 0;
                        cmbRightPlatform.SelectedIndex = 0;
                        lYukleniyor = false;
                    }
                }
            }
            catch { }
        }

        private void KargoFirmalariniYukle()
        {
            HashSet<string> firmalar = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ARAS KARGO", "YURTİÇİ KARGO", "SÜRAT KARGO", "DHL KARGO", "PTT KARGO", "KOLAY GELSİN"
            };
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("SELECT DISTINCT KargoFirma FROM KargoMatrisi WHERE KargoFirma IS NOT NULL AND KargoFirma != ''", b))
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read()) firmalar.Add(r["KargoFirma"].ToString().ToUpper().Trim());
                    }
                }
            }
            catch { }

            lYukleniyor = true;
            cmbKargoFirma.Items.Clear(); cmbRightKargoFirma.Items.Clear();
            cmbKargoFirma.Items.Add(""); cmbRightKargoFirma.Items.Add("");

            foreach (var f in firmalar)
            {
                cmbKargoFirma.Items.Add(f);
                cmbRightKargoFirma.Items.Add(f);
            }
            cmbKargoFirma.SelectedIndex = 0;
            cmbRightKargoFirma.SelectedIndex = 0;
            lYukleniyor = false;
        }

        private void BtnYeniKargoEkle_Click(object sender, EventArgs e)
        {
            if (cmbPlatform.SelectedIndex <= 0)
            {
                MessageBox.Show("⚠️ Lütfen yeni kargo firmasını eklemeden önce hangi platforma bağlanacağını sol taraftaki 1. Platform kutusundan seçin!", "Platform Seçimi Zorunlu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbPlatform.Focus();
                return;
            }

            string yeniFirma = txtYeniKargo.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(yeniFirma)) return;

            if (!cmbKargoFirma.Items.Contains(yeniFirma))
            {
                cmbKargoFirma.Items.Add(yeniFirma);
                cmbRightKargoFirma.Items.Add(yeniFirma);
            }

            cmbKargoFirma.SelectedItem = yeniFirma;
            txtYeniKargo.Clear();
        }

        private void CreateMatrisColumnsSafely(DataGridView dgv, DataTable dt)
        {
            dgv.DataSource = dt;
            // 🚀 MÜHÜR: CS0136 çökme hatasına neden olan dynamic string Column araması tamamen kaldırıldı. 
            // Sütun kilitleri ve formatlar %100 güvenli "Index-Based" [0] ve [1] mantığına geçirildi!
            if (dgv.Columns.Count >= 2)
            {
                dgv.Columns[0].ReadOnly = true; // Desi kolonu kilitli
                dgv.Columns[1].ReadOnly = false; // Fiyat kolonu açık
                dgv.Columns[1].DefaultCellStyle.Format = "N2";
                dgv.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void CmbLeft_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lYukleniyor) return;
            if (cmbPlatform.SelectedIndex > 0 && cmbKargoFirma.SelectedIndex > 0)
            {
                string plat = cmbPlatform.Text;
                string kargo = cmbKargoFirma.Text;

                AcDesiMatrisPenceresi(plat, kargo);

                lYukleniyor = true;
                cmbPlatform.SelectedIndex = 0;
                cmbKargoFirma.SelectedIndex = 0;
                lYukleniyor = false;
            }
        }

        private void AcDesiMatrisPenceresi(string plat, string kargo)
        {
            Form pop = new Form
            {
                Text = $"📦 DESİ FİYATLARI TANIMLAMA - {plat} [{kargo}]",
                Size = new Size(420, 560),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.FromArgb(33, 33, 36),
                ForeColor = Color.White,
                ShowIcon = false,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new Font("Segoe UI", 10F)
            };

            DataGridView dgvPop = new DataGridView
            {
                Location = new Point(15, 15),
                Width = 375,
                Height = 430,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                BackgroundColor = Color.FromArgb(24, 24, 27),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EnableHeadersVisualStyles = false
            };
            dgvPop.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgvPop.ColumnHeadersDefaultCellStyle.ForeColor = Color.Orange;
            dgvPop.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgvPop.DefaultCellStyle.ForeColor = Color.White;

            DataTable dt = new DataTable();
            dt.Columns.Add("Desi", typeof(int));
            dt.Columns.Add("Fiyat (TL)", typeof(double));

            Dictionary<int, double> baremler = new Dictionary<int, double>();
            for (int i = 1; i <= 30; i++) baremler[i] = 0.00;

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Desi, Fiyat FROM KargoMatrisi WHERE PlatformAd=@p AND KargoFirma=@k", conn))
                    {
                        cmd.Parameters.AddWithValue("@p", plat); cmd.Parameters.AddWithValue("@k", kargo);
                        using (var r = cmd.ExecuteReader()) { while (r.Read()) baremler[Convert.ToInt32(r["Desi"])] = Convert.ToDouble(r["Fiyat"]); }
                    }
                }
            }
            catch { }

            for (int i = 1; i <= 30; i++) dt.Rows.Add(i, baremler[i]);

            // Safe Method çağrısı ile çökme engellendi
            CreateMatrisColumnsSafely(dgvPop, dt);
            pop.Controls.Add(dgvPop);

            Button btnSave = new Button { Text = "💾 MATRİSİ VERİTABANINA MÜHÜRLE", Location = new Point(15, 460), Width = 375, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnSave.FlatAppearance.BorderSize = 0;

            btnSave.Click += (sX, eX) => {
                dgvPop.EndEdit();
                try
                {
                    using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        conn.Open();
                        using (var tx = conn.BeginTransaction())
                        {
                            using (var cmdDel = new SqliteCommand("DELETE FROM KargoMatrisi WHERE PlatformAd=@p AND KargoFirma=@k", conn, tx))
                            {
                                cmdDel.Parameters.AddWithValue("@p", plat); cmdDel.Parameters.AddWithValue("@k", kargo); cmdDel.ExecuteNonQuery();
                            }
                            foreach (DataRow row in dt.Rows)
                            {
                                int desi = Convert.ToInt32(row[0]);
                                double fiyat = row[1] == DBNull.Value ? 0.0 : Convert.ToDouble(row[1]);
                                using (var cmdIns = new SqliteCommand("INSERT INTO KargoMatrisi (PlatformAd, KargoFirma, Desi, Fiyat) VALUES (@p, @k, @d, @f)", conn, tx))
                                {
                                    cmdIns.Parameters.AddWithValue("@p", plat); cmdIns.Parameters.AddWithValue("@k", kargo);
                                    cmdIns.Parameters.AddWithValue("@d", desi); cmdIns.Parameters.AddWithValue("@f", fiyat);
                                    cmdIns.ExecuteNonQuery();
                                }
                            }
                            tx.Commit();
                        }
                    }
                    MatrisiListele(cmbRightPlatform.Text, cmbRightKargoFirma.Text);
                    pop.Close();
                    MessageBox.Show($"[{plat}] - [{kargo}] 30 Desilik Kargo fiyat matrisi başarıyla güncellendi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };
            pop.Controls.Add(btnSave);
            pop.ShowDialog();
        }

        private void MatrisiListele(string plat = "", string kargo = "")
        {
            try
            {
                DataTable dt = new DataTable();
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sorgu = "SELECT Id, PlatformAd as 'Platform', KargoFirma as 'Kargo Firması', Desi, Fiyat as 'Anlaşma Fiyatı (TL)' FROM KargoMatrisi WHERE 1=1";
                    if (!string.IsNullOrEmpty(plat)) sorgu += " AND PlatformAd = @p";
                    if (!string.IsNullOrEmpty(kargo)) sorgu += " AND KargoFirma = @k";
                    sorgu += " ORDER BY PlatformAd ASC, KargoFirma ASC, Desi ASC";

                    using (var cmd = new SqliteCommand(sorgu, b))
                    {
                        if (!string.IsNullOrEmpty(plat)) cmd.Parameters.AddWithValue("@p", plat);
                        if (!string.IsNullOrEmpty(kargo)) cmd.Parameters.AddWithValue("@k", kargo);
                        using (var oku = cmd.ExecuteReader()) { dt.Load(oku); }
                    }
                }
                dgvMatris.DataSource = dt;

                // 🚀 MÜHÜR: Ana Tablo Sütun atamaları da korumaya alındı
                if (dgvMatris.Columns.Count >= 5)
                {
                    dgvMatris.Columns[0].Visible = false; // Id gizle
                    dgvMatris.Columns[4].DefaultCellStyle.Format = "N2"; // Fiyat formatla
                }
            }
            catch { }
        }

        private void BtnBaremDuzenle_Click(object sender, EventArgs e)
        {
            if (dgvMatris.SelectedRows.Count == 0) return;
            var row = dgvMatris.SelectedRows[0];

            // 🚀 MÜHÜR: Sağ tık düzeltme hücreleri de Index'e bağlandı
            long id = Convert.ToInt64(row.Cells[0].Value);
            string pAd = row.Cells[1].Value.ToString();
            string kFirma = row.Cells[2].Value.ToString();
            string dMiktar = row.Cells[3].Value.ToString();
            double currentFiyat = Convert.ToDouble(row.Cells[4].Value);

            Form frm = new Form { Text = "✏️ Barem Fiyatı Düzenle", Size = new Size(320, 200), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MaximizeBox = false, MinimizeBox = false };
            TextBox txtNewFiyat;
            frm.Controls.Add(FormElemaniYarat($"{pAd} -> {kFirma} [{dMiktar} Desi] Fiyatı:", out txtNewFiyat, new Point(15, 20), 270, currentFiyat.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));

            Button btnGuncelle = new Button { Text = "💾 GÜNCELLE", Location = new Point(15, 95), Width = 270, Height = 38, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };
            btnGuncelle.FlatAppearance.BorderSize = 0;
            btnGuncelle.Click += (sX, eX) => {
                double.TryParse(txtNewFiyat.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nF);
                try
                {
                    using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        conn.Open();
                        using (var cmd = new SqliteCommand("UPDATE KargoMatrisi SET Fiyat=@f WHERE Id=@id", conn))
                        {
                            cmd.Parameters.AddWithValue("@f", nF); cmd.Parameters.AddWithValue("@id", id); cmd.ExecuteNonQuery();
                        }
                    }
                    MatrisiListele(cmbRightPlatform.Text, cmbRightKargoFirma.Text);
                    frm.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            frm.Controls.Add(btnGuncelle); frm.ShowDialog();
        }

        private void BtnTopluYukle_Click(object sender, EventArgs e)
        {
            if (cmbPlatform.SelectedIndex <= 0 || string.IsNullOrWhiteSpace(cmbKargoFirma.Text)) { MessageBox.Show("Lütfen önce listenin yükleneceği Platformu ve Kargo firmasını sol kutudan seçin!", "Seçim Eksik"); return; }
            OpenFileDialog ofd = new OpenFileDialog { Filter = "CSV Dosyası|*.csv", Title = "Toplu Kargo Listesini Seçin" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string[] satirlar = File.ReadAllLines(ofd.FileName, System.Text.Encoding.UTF8);
                    int sayac = 0;
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var tx = b.BeginTransaction())
                        {
                            for (int i = 0; i < satirlar.Length; i++)
                            {
                                if (string.IsNullOrWhiteSpace(satirlar[i])) continue;
                                string[] s = satirlar[i].Split(';', ','); if (s.Length < 2) continue;
                                if (int.TryParse(s[0].Trim(), out int desi) && double.TryParse(s[1].Trim().Replace('.', ','), out double fiyat))
                                {
                                    if (desi > 30) continue;
                                    using (var cmd = new SqliteCommand("INSERT INTO KargoMatrisi (PlatformAd, KargoFirma, Desi, Fiyat) VALUES (@p, @k, @d, @f) ON CONFLICT(PlatformAd, KargoFirma, Desi) DO UPDATE SET Fiyat = @f", b, tx))
                                    {
                                        cmd.Parameters.AddWithValue("@p", cmbPlatform.Text); cmd.Parameters.AddWithValue("@k", cmbKargoFirma.Text);
                                        cmd.Parameters.AddWithValue("@d", desi); cmd.Parameters.AddWithValue("@f", fiyat); cmd.ExecuteNonQuery();
                                    }
                                    sayac++;
                                }
                            }
                            tx.Commit();
                        }
                    }
                    MatrisiListele(cmbRightPlatform.Text, cmbRightKargoFirma.Text);
                    MessageBox.Show($"Toplam {sayac} barem başarıyla içe aktarıldı!", "Başarılı");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

        private void BtnSil_Click(object sender, EventArgs e)
        {
            if (dgvMatris.SelectedRows.Count == 0) return;
            int seciliId = Convert.ToInt32(dgvMatris.SelectedRows[0].Cells[0].Value);
            if (MessageBox.Show("Seçili baremi silmek istiyor musunuz?", "Onay", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open(); using (var cmd = new SqliteCommand("DELETE FROM KargoMatrisi WHERE Id = @id", b)) { cmd.Parameters.AddWithValue("@id", seciliId); cmd.ExecuteNonQuery(); }
                    }
                    MatrisiListele(cmbRightPlatform.Text, cmbRightKargoFirma.Text);
                }
                catch { }
            }
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan) { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) }; kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
        private Panel FormComboYarat(string baslik, out ComboBox kutu, Point konum, int genislik) { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) }; kutu = new ComboBox { Location = new Point(0, 18), Width = genislik, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F) }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
    }
}