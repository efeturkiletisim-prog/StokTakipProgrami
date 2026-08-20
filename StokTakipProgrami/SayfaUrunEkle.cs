using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    // 📇 ÜRÜN TANIM KARTI YÖNETİM MERKEZİ (MARKA 'X' HATASI KÖKTEN ÇÖZÜLDÜ)
    public partial class SayfaUrunEkle : UserControl
    {
        private TextBox ucKartTxtBarkod, ucKartTxtUrunAdi;
        private ComboBox ucKartCmbKategori, ucKartCmbMarka;
        private Button ucKartBtnKaydet;
        private DataGridView ucKartDgvListe;

        public SayfaUrunEkle()
        {
            InitializeComponent();
            TabloAltyapisiniGuvenceyeAl();
            ArayuzElemanlariniInsaEt();
            DinamikKategorileriYukle();
            DinamikMarkalariYukle();
        }

        private void SayfaUrunEkle_Load(object sender, EventArgs e)
        {
            MevcutUrunKartlariniListele();
            if (ucKartTxtBarkod != null) ucKartTxtBarkod.Focus();
        }

        private void TabloAltyapisiniGuvenceyeAl()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = "CREATE TABLE IF NOT EXISTS Urunler (Barkod TEXT PRIMARY KEY, UrunAdi TEXT, Kategori TEXT, Marka TEXT);";
                    using (var cmd = new SqliteCommand(sql, b)) { cmd.ExecuteNonQuery(); }

                    var sutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdCheck = new SqliteCommand("PRAGMA table_info('Urunler');", b))
                    using (var r = cmdCheck.ExecuteReader()) { while (r.Read()) sutunlar.Add(r["name"].ToString()); }
                    if (!sutunlar.Contains("Marka")) { using (var cmdAdd = new SqliteCommand("ALTER TABLE Urunler ADD COLUMN Marka TEXT DEFAULT 'DİĞER';", b)) cmdAdd.ExecuteNonQuery(); }
                }
            }
            catch { }
        }

        private void ArayuzElemanlariniInsaEt()
        {
            this.Controls.Clear(); this.Dock = DockStyle.Fill; this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "📇 ÜRÜN TANIM KARTI YÖNETİM MERKEZİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = Color.Transparent };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 360F)); // SOL FORM
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));  // SAĞ LİSTE

            // SOL FORM
            Panel pnlLeftForm = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15) };
            Label lblVeriEkleBaslik = new Label { Text = "🟢 VERİ EKLE", Location = new Point(15, 10), AutoSize = true, Font = new Font("Segoe UI Black", 14F, FontStyle.Bold), ForeColor = Color.White };
            pnlLeftForm.Controls.Add(lblVeriEkleBaslik);

            pnlLeftForm.Controls.Add(FormElemaniYarat("Ürün Barkodu (Zorunlu):", out ucKartTxtBarkod, new Point(15, 50), 320));
            pnlLeftForm.Controls.Add(FormElemaniYarat("Ürün Kart Adı:", out ucKartTxtUrunAdi, new Point(15, 115), 320));
            pnlLeftForm.Controls.Add(FormComboYarat("Ürün Kategorisi:", out ucKartCmbKategori, new Point(15, 180), 320));
            pnlLeftForm.Controls.Add(FormComboYarat("Ürün Markası:", out ucKartCmbMarka, new Point(15, 245), 320));

            ucKartBtnKaydet = new Button { Text = "💾 KAYDET", Location = new Point(15, 330), Width = 320, Height = 45, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
            ucKartBtnKaydet.FlatAppearance.BorderSize = 0; ucKartBtnKaydet.Click += UcKartBtnKaydet_Click;
            pnlLeftForm.Controls.Add(ucKartBtnKaydet);

            // SAĞ LİSTE
            Panel pnlRightList = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
            Label lblListTitle = new Label { Text = "📋 MEVCUT ÜRÜN KARTLARI (Sağ Tık: Düzenle / Sil)", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            pnlRightList.Controls.Add(lblListTitle);

            ContextMenuStrip cms = new ContextMenuStrip();
            cms.Items.Add("✏️ Kartı Düzenle", null, BtnKartDuzenle_Click);
            cms.Items.Add("❌ Kartı Sil", null, BtnKartSil_Click);

            ucKartDgvListe = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(26, 26, 28), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 10F), DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, SelectionBackColor = Color.FromArgb(0, 120, 215) }, EnableHeadersVisualStyles = false, ColumnHeadersHeight = 40, ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing, BorderStyle = BorderStyle.None, ContextMenuStrip = cms };
            ucKartDgvListe.DataError += (s, e) => { e.Cancel = true; };
            ucKartDgvListe.CellMouseDown += UcKartDgvListe_CellMouseDown;

            pnlRightList.Controls.Add(ucKartDgvListe);
            ucKartDgvListe.BringToFront();

            mainLayout.Controls.Add(pnlLeftForm, 0, 0); mainLayout.Controls.Add(pnlRightList, 1, 0);

            this.Controls.Add(mainLayout);
            this.Controls.Add(pnlHeader);
            pnlHeader.SendToBack();
            mainLayout.BringToFront();
        }

        private void UcKartDgvListe_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    ucKartDgvListe.ClearSelection();
                    ucKartDgvListe.Rows[e.RowIndex].Selected = true;
                    ucKartDgvListe.CurrentCell = ucKartDgvListe.Rows[e.RowIndex].Cells[e.ColumnIndex];
                }
            }
        }

        private void DinamikKategorileriYukle()
        {
            try { using (var c = new SqliteConnection(Veritabanı.BaglantiYolu)) { c.Open(); using (var cmd = new SqliteCommand("SELECT Ad FROM Kategoriler ORDER BY Ad ASC", c)) using (var r = cmd.ExecuteReader()) { ucKartCmbKategori.Items.Clear(); ucKartCmbKategori.Items.Add(""); while (r.Read()) ucKartCmbKategori.Items.Add(r["Ad"].ToString()); } } } catch { }
            if (ucKartCmbKategori.Items.Count <= 1) { ucKartCmbKategori.Items.Clear(); ucKartCmbKategori.Items.AddRange(new object[] { "", "GENEL", "TELEFON", "CEP TELEFONU", "AKSESUAR" }); }
            ucKartCmbKategori.SelectedIndex = 0;
        }

        private void DinamikMarkalariYukle()
        {
            try { using (var c = new SqliteConnection(Veritabanı.BaglantiYolu)) { c.Open(); using (var cmd = new SqliteCommand("SELECT Ad FROM Markalar ORDER BY Ad ASC", c)) using (var r = cmd.ExecuteReader()) { ucKartCmbMarka.Items.Clear(); ucKartCmbMarka.Items.Add(""); while (r.Read()) ucKartCmbMarka.Items.Add(r["Ad"].ToString()); } } } catch { }
            if (ucKartCmbMarka.Items.Count <= 1) { ucKartCmbMarka.Items.Clear(); ucKartCmbMarka.Items.AddRange(new object[] { "", "YOK/DİĞER", "APPLE", "SAMSUNG", "XIAOMI" }); }
            ucKartCmbMarka.SelectedIndex = 0;
        }

        // 🚀 MÜHÜR: X HATA SORUNU ÇÖZÜLDÜ, SQL YALINLAŞTIRILDI (Sıra: Barkod -> Marka -> Ürün Adı -> Kategori)
        private void MevcutUrunKartlariniListele()
        {
            try
            {
                using (var c = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    c.Open();
                    string sql = @"SELECT 
                                    Barkod as 'Barkod', 
                                    Marka as 'Marka', 
                                    UrunAdi as 'Ürün Adı',
                                    Kategori as 'Kategori' 
                                   FROM Urunler 
                                   ORDER BY Barkod ASC";
                    using (var cmd = new SqliteCommand(sql, c))
                    {
                        DataTable dt = new DataTable();
                        using (var reader = cmd.ExecuteReader()) { dt.Load(reader); }
                        ucKartDgvListe.DataSource = dt;
                    }
                }
                ucKartDgvListe.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
                ucKartDgvListe.ColumnHeadersDefaultCellStyle.ForeColor = Color.Orange;
                ucKartDgvListe.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
                ucKartDgvListe.ClearSelection();
            }
            catch { }
        }

        private void UcKartBtnKaydet_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ucKartTxtBarkod.Text) || string.IsNullOrEmpty(ucKartTxtUrunAdi.Text) || string.IsNullOrEmpty(ucKartCmbKategori.Text) || string.IsNullOrEmpty(ucKartCmbMarka.Text))
            {
                MessageBox.Show("Barkod, Ürün Adı, Kategori ve Marka alanlarının tümünü doldurmak/seçmek zorunludur!", "Eksik Veri", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 🚀 MÜHÜR: Boş değer hatası SQL'de değil, doğrudan burada engellendi.
            string guvenliMarka = ucKartCmbMarka.Text.Trim();
            if (string.IsNullOrEmpty(guvenliMarka)) guvenliMarka = "DİĞER";

            try
            {
                using (var c = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    c.Open();
                    using (var cmd = new SqliteCommand("INSERT OR REPLACE INTO Urunler (Barkod, UrunAdi, Kategori, Marka) VALUES (@b, @a, @k, @m)", c))
                    {
                        cmd.Parameters.AddWithValue("@b", ucKartTxtBarkod.Text.Trim());
                        cmd.Parameters.AddWithValue("@a", ucKartTxtUrunAdi.Text.Trim());
                        cmd.Parameters.AddWithValue("@k", ucKartCmbKategori.Text);
                        cmd.Parameters.AddWithValue("@m", guvenliMarka);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Ürün Kaydı Başarıyla Güncellendi/Eklendi!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                FormuSifirla();
                MevcutUrunKartlariniListele();
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }

        private void BtnKartDuzenle_Click(object sender, EventArgs e)
        {
            if (ucKartDgvListe.SelectedRows.Count == 0) return;
            var r = ucKartDgvListe.SelectedRows[0];
            string barkod = r.Cells["Barkod"].Value.ToString();
            string eskiUrunAdi = r.Cells["Ürün Adı"].Value.ToString();

            Form frmPopup = new Form
            {
                Text = "📝 Ürün Adı Güncelleme",
                Size = new Size(390, 210),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.FromArgb(33, 33, 36),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowIcon = false
            };

            Label lblB = new Label { Text = $"Barkod: {barkod}", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.Orange };
            Label lblA = new Label { Text = "Yeni Ürün Kart Adı:", Location = new Point(20, 52), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Regular), ForeColor = Color.DarkGray };
            TextBox txtYeniAd = new TextBox { Text = eskiUrunAdi, Location = new Point(20, 74), Width = 330, Font = new Font("Segoe UI", 10.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

            Button btnGuncelle = new Button { Text = "💾 DEĞİŞİKLİKLERİ MÜHÜRLE", Location = new Point(20, 115), Width = 330, Height = 38, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnGuncelle.FlatAppearance.BorderSize = 0;

            btnGuncelle.Click += (sX, eX) => {
                string yeniAd = txtYeniAd.Text.Trim();
                if (string.IsNullOrEmpty(yeniAd)) { MessageBox.Show("Ürün adı boş bırakılamaz patron!"); return; }
                try
                {
                    using (var c = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        c.Open();
                        using (var cmd = new SqliteCommand("UPDATE Urunler SET UrunAdi=@a WHERE Barkod=@b", c))
                        {
                            cmd.Parameters.AddWithValue("@a", yeniAd);
                            cmd.Parameters.AddWithValue("@b", barkod);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    frmPopup.Close();
                    MevcutUrunKartlariniListele();
                    FormuSifirla();
                    MessageBox.Show("Ürün kart adı saniyesinde güncellendi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };

            frmPopup.Controls.Add(lblB); frmPopup.Controls.Add(lblA);
            frmPopup.Controls.Add(txtYeniAd); frmPopup.Controls.Add(btnGuncelle);
            frmPopup.ShowDialog();
        }

        private void BtnKartSil_Click(object sender, EventArgs e)
        {
            if (ucKartDgvListe.SelectedRows.Count == 0) return;
            if (MessageBox.Show("Bu ürün kartını sistemden kalıcı olarak silmek istiyor musunuz?", "Kalıcı Kart Silme Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes)
            {
                try { using (var c = new SqliteConnection(Veritabanı.BaglantiYolu)) { c.Open(); using (var cmd = new SqliteCommand("DELETE FROM Urunler WHERE Barkod=@b", c)) { cmd.Parameters.AddWithValue("@b", ucKartDgvListe.SelectedRows[0].Cells["Barkod"].Value.ToString()); cmd.ExecuteNonQuery(); } } FormuSifirla(); MevcutUrunKartlariniListele(); } catch { }
            }
        }

        private void FormuSifirla()
        {
            ucKartTxtBarkod.Clear(); ucKartTxtBarkod.ReadOnly = false; ucKartTxtBarkod.BackColor = Color.FromArgb(41, 41, 44); ucKartTxtUrunAdi.Clear(); if (ucKartCmbKategori.Items.Count > 0) ucKartCmbKategori.SelectedIndex = 0; if (ucKartCmbMarka.Items.Count > 0) ucKartCmbMarka.SelectedIndex = 0; ucKartBtnKaydet.BackColor = Color.FromArgb(0, 120, 215); ucKartTxtBarkod.Focus();
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik) { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) }; kutu = new TextBox { Location = new Point(0, 18), Width = genislik, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
        private Panel FormComboYarat(string baslik, out ComboBox kutu, Point konum, int genislik) { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) }; kutu = new ComboBox { Location = new Point(0, 18), Width = genislik, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F) }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
    }
}