using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // 🔍 STOK SORGULAMA VE PROFESSIONAL ENVANTER YÖNETİM MERKEZİ (ALT PANE %70, ÜST %30 ORAN KORUMALI)
    public partial class SayfaStokSorgula : UserControl
    {
        private TextBox dynTxtStokAra;
        private Button dynBtnStokAra;
        private DataGridView dynDgvAramaSonuclari, dynDgvTumStokListesi;
        private Label dynLblToplamAdet, dynLblToplamTutar;
        private ContextMenuStrip cmsGridIslem;
        private SplitContainer splitKonteyner;

        // Arka planda kullanılacak dinamik liste depoları
        private List<string> dbKategoriler = new List<string>();
        private List<string> dbMarkalar = new List<string>();
        private List<string> dbFirmalar = new List<string>();

        public SayfaStokSorgula()
        {
            ArayuzElemanlariniInsaEt();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            // Önce geçmişten kalan veya sıfırlamadan kaçan hayalet veriler temizlenir!
            HayaletVerileriTemizle();

            DinamikVerileriHafizayaAl();
            OzetIstatistikleriHesapla();
            TumEnvanteriListele();
            StoklariSorgula("");

            // İlk yüklemede splitter'ı güvenli bir şekilde konumlandır
            DinamikSplitterHizala();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            // Ekran boyutu her değiştiğinde oranı güvenli bir şekilde koru
            DinamikSplitterHizala();
        }

        // 🚀 MÜHÜR: Çökmeleri Engelleyen Güvenli Splitter Hizalama Motoru
        private void DinamikSplitterHizala()
        {
            if (splitKonteyner != null && splitKonteyner.Height > 100)
            {
                try
                {
                    // Oranı nesnenin kendi yüksekliği üzerinden hesaplamak daha kararlıdır
                    int hesaplananMesafe = (int)(splitKonteyner.Height * 0.30);

                    // Hesaplanan değerin splitKonteyner sınırları içinde olup olmadığını doğrula (Çökme Engelleyici Kalkan)
                    if (hesaplananMesafe >= splitKonteyner.Panel1MinSize &&
                        hesaplananMesafe <= (splitKonteyner.Height - splitKonteyner.Panel2MinSize))
                    {
                        splitKonteyner.SplitterDistance = hesaplananMesafe;
                    }
                }
                catch { }
            }
        }

        private void HayaletVerileriTemizle()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = "DELETE FROM Imeiler WHERE Durum = 'STOKTA' AND Barkod NOT IN (SELECT Barkod FROM Stoklar WHERE StokMiktari > 0)";
                    using (var cmd = new SqliteCommand(sql, conn)) { cmd.ExecuteNonQuery(); }
                }
            }
            catch { }
        }

        private void DinamikVerileriHafizayaAl()
        {
            dbKategoriler.Clear(); dbMarkalar.Clear(); dbFirmalar.Clear();
            dbKategoriler.Add(""); dbMarkalar.Add(""); dbFirmalar.Add("");

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Ad FROM Kategoriler ORDER BY Ad ASC", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) dbKategoriler.Add(r["Ad"].ToString().ToUpper()); }

                    using (var cmd = new SqliteCommand("SELECT Ad FROM Markalar ORDER BY Ad ASC", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) dbMarkalar.Add(r["Ad"].ToString().ToUpper()); }

                    using (var cmd = new SqliteCommand("SELECT Ad FROM Firmalar ORDER BY Ad ASC", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) dbFirmalar.Add(r["Ad"].ToString().ToUpper()); }
                }
            }
            catch { }
        }

        private void ArayuzElemanlariniInsaEt()
        {
            this.Controls.Clear(); this.Dock = DockStyle.Fill; this.BackColor = Color.FromArgb(33, 33, 36); this.Padding = new Padding(15);

            cmsGridIslem = new ContextMenuStrip();
            cmsGridIslem.Items.Add("✏️ Seçili Kaydı Düzenle", null, BtnStokDuzenle_Click);
            cmsGridIslem.Items.Add("❌ Seçili Stoğu Sil", null, BtnStokSil_Click);

            Panel pnlOzetKonteyner = new Panel { Dock = DockStyle.Top, Height = 70, BackColor = Color.Transparent };
            Panel pnlAdetKart = new Panel { Location = new Point(5, 5), Width = 300, Height = 60, BackColor = Color.FromArgb(26, 26, 39), BorderStyle = BorderStyle.FixedSingle };
            dynLblToplamAdet = new Label { Text = "📊 TOPLAM STOK: 0 ADET", Dock = DockStyle.Fill, Font = new Font("Segoe UI Black", 11F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248), TextAlign = ContentAlignment.MiddleCenter };
            pnlAdetKart.Controls.Add(dynLblToplamAdet);

            Panel pnlTutarKart = new Panel { Location = new Point(320, 5), Width = 400, Height = 60, BackColor = Color.FromArgb(26, 26, 39), BorderStyle = BorderStyle.FixedSingle };
            dynLblToplamTutar = new Label { Text = "💰 TOPLAM MALİYET: 0,00 TL", Dock = DockStyle.Fill, Font = new Font("Segoe UI Black", 11F, FontStyle.Bold), ForeColor = Color.FromArgb(34, 197, 94), TextAlign = ContentAlignment.MiddleCenter };
            pnlTutarKart.Controls.Add(dynLblToplamTutar);

            pnlOzetKonteyner.Controls.Add(pnlAdetKart); pnlOzetKonteyner.Controls.Add(pnlTutarKart);
            this.Controls.Add(pnlOzetKonteyner);

            Panel pnlAramaBar = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.Transparent };
            Label lblAraBaslik = new Label { Text = "Ürün veya IMEI Ara:", Location = new Point(5, 15), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White };
            dynTxtStokAra = new TextBox { Location = new Point(175, 12), Width = 400, Font = new Font("Segoe UI", 12F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            dynTxtStokAra.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; StoklariSorgula(dynTxtStokAra.Text.Trim()); } };

            dynBtnStokAra = new Button { Text = "🔍 SORGULA", Location = new Point(585, 11), Width = 140, Height = 28, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnStokAra.FlatAppearance.BorderSize = 0; dynBtnStokAra.Click += (s, e) => StoklariSorgula(dynTxtStokAra.Text.Trim());

            pnlAramaBar.Controls.Add(lblAraBaslik); pnlAramaBar.Controls.Add(dynTxtStokAra); pnlAramaBar.Controls.Add(dynBtnStokAra);
            this.Controls.Add(pnlAramaBar);

            splitKonteyner = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterWidth = 6, BackColor = Color.FromArgb(33, 33, 36) };
            this.Controls.Add(splitKonteyner); splitKonteyner.BringToFront();

            Panel pnlUstGrid = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
            Label lblUstBaslik = new Label { Text = "🎯 SORGULAMA SONUÇLARI (Sağ Tık: Düzenle/Sil)", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange };
            dynDgvAramaSonuclari = GridYarat(Color.Orange);
            pnlUstGrid.Controls.Add(dynDgvAramaSonuclari); pnlUstGrid.Controls.Add(lblUstBaslik);
            dynDgvAramaSonuclari.BringToFront(); splitKonteyner.Panel1.Controls.Add(pnlUstGrid);

            Panel pnlAltGrid = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
            Label lblAltBaslik = new Label { Text = "📦 TÜM DEPO ENVANTER LİSTESİ (Sağ Tık: Düzenle/Sil)", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynDgvTumStokListesi = GridYarat(Color.FromArgb(56, 189, 248));
            pnlAltGrid.Controls.Add(dynDgvTumStokListesi); pnlAltGrid.Controls.Add(lblAltBaslik);
            dynDgvTumStokListesi.BringToFront(); splitKonteyner.Panel2.Controls.Add(pnlAltGrid);

            dynDgvAramaSonuclari.CellClick += (s, e) => { dynDgvTumStokListesi.ClearSelection(); };
            dynDgvTumStokListesi.CellClick += (s, e) => { dynDgvAramaSonuclari.ClearSelection(); };
        }

        private DataTable SabitTabloSablonuYarat()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Barkod No", typeof(string));
            dt.Columns.Add("Kategori", typeof(string));
            dt.Columns.Add("Marka", typeof(string));
            dt.Columns.Add("Ürün Adı", typeof(string));
            dt.Columns.Add("IMEI Numarası", typeof(string));
            dt.Columns.Add("Mevcut Stok", typeof(int));
            dt.Columns.Add("Maliyet (TL)", typeof(double));
            dt.Columns.Add("Tedarikçi", typeof(string));
            dt.Columns.Add("Giriş Tarihi", typeof(string));
            return dt;
        }

        private void OzetIstatistikleriHesapla()
        {
            try
            {
                int netToplamAdet = 0;
                double netToplamMaliyet = 0;

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sqlGenel = @"
                        SELECT IFNULL(SUM(s.StokMiktari), 0) as GenelAdet, 
                               IFNULL(SUM(s.StokMiktari * s.AlisFiyati), 0) as GenelMaliyet 
                        FROM Stoklar s
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod
                        WHERE u.Kategori NOT LIKE '%TELEFON%' AND u.Kategori NOT LIKE '%CEP%' AND s.StokMiktari > 0";

                    using (var cmd = new SqliteCommand(sqlGenel, conn))
                    {
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                netToplamAdet += Convert.ToInt32(r["GenelAdet"]);
                                netToplamMaliyet += Convert.ToDouble(r["GenelMaliyet"]);
                            }
                        }
                    }

                    string sqlTelefon = @"
                        SELECT COUNT(i.ImeiNo) as TelAdet,
                               IFNULL(SUM(s.AlisFiyati), 0) as TelMaliyet
                        FROM Imeiler i
                        INNER JOIN Stoklar s ON i.Barkod = s.Barkod
                        WHERE i.Durum = 'STOKTA' AND s.StokMiktari > 0";

                    using (var cmd = new SqliteCommand(sqlTelefon, conn))
                    {
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                netToplamAdet += Convert.ToInt32(r["TelAdet"]);
                                netToplamMaliyet += Convert.ToDouble(r["TelMaliyet"]);
                            }
                        }
                    }
                }

                dynLblToplamAdet.Text = $"📊 TOPLAM STOK: {netToplamAdet} ADET";
                dynLblToplamTutar.Text = $"💰 TOPLAM MALİYET: {netToplamMaliyet:N2} TL";
            }
            catch { }
        }

        private void StoklariSorgula(string aramaMetni)
        {
            try
            {
                DataTable dt = SabitTabloSablonuYarat();
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT s.Barkod, u.UrunAdi, u.Kategori, IFNULL(u.Marka, '-') as Marka, 
                               IFNULL(i.ImeiNo, '-') as Imei,
                               s.StokMiktari, s.AlisFiyati, s.Firma, s.Tarih 
                        FROM Stoklar s 
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod
                        LEFT JOIN Imeiler i ON s.Barkod = i.Barkod AND i.Durum = 'STOKTA'
                        WHERE s.StokMiktari > 0";

                    if (!string.IsNullOrEmpty(aramaMetni))
                    {
                        sql += " AND (s.Barkod LIKE @ara OR u.UrunAdi LIKE @ara OR u.Marka LIKE @ara OR i.ImeiNo LIKE @ara)";
                    }
                    else
                    {
                        sql += " AND 1=0";
                    }

                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        if (!string.IsNullOrEmpty(aramaMetni)) cmd.Parameters.AddWithValue("@ara", "%" + aramaMetni + "%");
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string kat = reader["Kategori"]?.ToString().ToUpper() ?? "";
                                string imeiNo = reader["Imei"]?.ToString() ?? "-";
                                int adet = reader["StokMiktari"] == DBNull.Value ? 0 : Convert.ToInt32(reader["StokMiktari"]);

                                if ((kat.Contains("TELEFON") || kat.Contains("CEP")) && imeiNo != "-")
                                {
                                    adet = 1;
                                }

                                dt.Rows.Add(
                                    reader["Barkod"]?.ToString(), reader["Kategori"]?.ToString(), reader["Marka"]?.ToString(), reader["UrunAdi"]?.ToString(),
                                    imeiNo, adet,
                                    reader["AlisFiyati"] == DBNull.Value ? 0.0 : Convert.ToDouble(reader["AlisFiyati"]),
                                    reader["Firma"]?.ToString(), reader["Tarih"]?.ToString()
                                );
                            }
                        }
                    }
                }

                dynDgvAramaSonuclari.DataSource = dt;

                if (dynDgvAramaSonuclari.Columns["Maliyet (TL)"] != null) dynDgvAramaSonuclari.Columns["Maliyet (TL)"].DefaultCellStyle.Format = "N2";

                if (!string.IsNullOrEmpty(aramaMetni) && dt.Rows.Count == 0)
                {
                    MessageBox.Show("⚠️ Aradığınız kriterlere uygun aktif bir stok bulunamadı veya tükenmiş!", "Envanter Bilgisi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dynTxtStokAra.Clear();
                    dynTxtStokAra.Focus();
                }
            }
            catch { }
        }

        private void TumEnvanteriListele()
        {
            try
            {
                DataTable dt = SabitTabloSablonuYarat();
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT s.Barkod, u.UrunAdi, u.Kategori, IFNULL(u.Marka, '-') as Marka, 
                               IFNULL(i.ImeiNo, '-') as Imei,
                               s.StokMiktari, s.AlisFiyati, s.Firma, s.Tarih 
                        FROM Stoklar s 
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod 
                        LEFT JOIN Imeiler i ON s.Barkod = i.Barkod AND i.Durum = 'STOKTA'
                        WHERE s.StokMiktari > 0
                        ORDER BY s.rowid DESC";

                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string kat = reader["Kategori"]?.ToString().ToUpper() ?? "";
                                string imeiNo = reader["Imei"]?.ToString() ?? "-";
                                int adet = reader["StokMiktari"] == DBNull.Value ? 0 : Convert.ToInt32(reader["StokMiktari"]);

                                if ((kat.Contains("TELEFON") || kat.Contains("CEP")) && imeiNo != "-")
                                {
                                    adet = 1;
                                }

                                dt.Rows.Add(
                                    reader["Barkod"]?.ToString(), reader["Kategori"]?.ToString(), reader["Marka"]?.ToString(), reader["UrunAdi"]?.ToString(),
                                    imeiNo, adet,
                                    reader["AlisFiyati"] == DBNull.Value ? 0.0 : Convert.ToDouble(reader["AlisFiyati"]),
                                    reader["Firma"]?.ToString(), reader["Tarih"]?.ToString()
                                );
                            }
                        }
                    }
                }
                dynDgvTumStokListesi.DataSource = dt;

                if (dynDgvTumStokListesi.Columns["Maliyet (TL)"] != null) dynDgvTumStokListesi.Columns["Maliyet (TL)"].DefaultCellStyle.Format = "N2";
            }
            catch { }
        }

        private DataGridView GridYarat(Color baslikRengi)
        {
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(24, 24, 27),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9.5F),
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 35,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ContextMenuStrip = cmsGridIslem
            };

            dgv.MouseDown += (s, e) => {
                if (e.Button == MouseButtons.Right)
                {
                    var h = dgv.HitTest(e.X, e.Y);
                    if (h.Type == DataGridViewHitTestType.Cell && h.RowIndex >= 0)
                    {
                        dgv.ClearSelection(); dgv.Rows[h.RowIndex].Selected = true;
                    }
                }
            };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = baslikRengi; dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgv.DefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgv.DataError += (s, e) => { e.Cancel = true; };
            return dgv;
        }

        private DataGridView AktifTabloyuBul()
        {
            if (dynDgvAramaSonuclari.SelectedRows.Count > 0) return dynDgvAramaSonuclari;
            if (dynDgvTumStokListesi.SelectedRows.Count > 0) return dynDgvTumStokListesi;
            return null;
        }

        private void BtnStokSil_Click(object sender, EventArgs e)
        {
            DataGridView aktifGrid = AktifTabloyuBul(); if (aktifGrid == null) return;

            var satir = aktifGrid.SelectedRows[0];
            string barkod = satir.Cells["Barkod No"].Value.ToString();
            string urunAdi = satir.Cells["Ürün Adı"].Value.ToString();
            string imei = satir.Cells["IMEI Numarası"].Value.ToString();

            bool isImei = (imei != "-");
            string soru = isImei
                ? $"[{urunAdi}] ürününün [{imei}] numaralı BİREYSEL TELEFON STOĞUNU silmek istiyor musunuz?\n(Envanter adedi 1 düşecektir)"
                : $"[{urunAdi}] ürününün GENEL STOĞUNU tamamen silmek istiyor musunuz?";

            if (MessageBox.Show(soru, "Kalıcı Silme Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        conn.Open();
                        if (isImei)
                        {
                            using (var cmd = new SqliteCommand("DELETE FROM Imeiler WHERE ImeiNo=@i AND Barkod=@b", conn)) { cmd.Parameters.AddWithValue("@i", imei); cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery(); }
                            using (var cmd = new SqliteCommand("UPDATE Stoklar SET Adet = MAX(0, Adet - 1), StokMiktari = MAX(0, StokMiktari - 1) WHERE Barkod=@b", conn)) { cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery(); }
                            using (var cmd = new SqliteCommand("DELETE FROM Stoklar WHERE Barkod=@b AND StokMiktari <= 0", conn)) { cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery(); }
                        }
                        else
                        {
                            using (var cmd = new SqliteCommand("DELETE FROM Stoklar WHERE Barkod=@b", conn)) { cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery(); }
                            using (var cmd = new SqliteCommand("DELETE FROM Imeiler WHERE Barkod=@b AND Durum='STOKTA'", conn)) { cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery(); }
                        }
                    }
                    TumEnvanteriListele(); StoklariSorgula(dynTxtStokAra.Text.Trim()); OzetIstatistikleriHesapla();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }
        }

        private void BtnStokDuzenle_Click(object sender, EventArgs e)
        {
            DataGridView aktifGrid = AktifTabloyuBul(); if (aktifGrid == null) return;

            var satir = aktifGrid.SelectedRows[0];
            string barkod = satir.Cells["Barkod No"].Value.ToString();
            string kategori = satir.Cells["Kategori"].Value.ToString();
            string marka = satir.Cells["Marka"].Value.ToString();
            string urunAdi = satir.Cells["Ürün Adı"].Value.ToString();
            string eskiImei = satir.Cells["IMEI Numarası"].Value.ToString();
            double maliyet = Convert.ToDouble(satir.Cells["Maliyet (TL)"].Value);
            string firma = satir.Cells["Tedarikçi"].Value.ToString();

            Form frmEdit = new Form { Text = "✏️ Envanter Kartı Düzenleme", Size = new Size(340, 500), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MaximizeBox = false, MinimizeBox = false };

            Label lblBkd = new Label { Text = $"Barkod: {barkod} (Kilitli)", Location = new Point(15, 15), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Orange };
            frmEdit.Controls.Add(lblBkd);

            Label lblKat = new Label { Text = "Ürün Kategorisi:", Location = new Point(15, 45), AutoSize = true, ForeColor = Color.DarkGray };
            ComboBox cmbKat = new ComboBox { Location = new Point(15, 65), Width = 290, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F) };
            cmbKat.Items.AddRange(dbKategoriler.ToArray());
            if (cmbKat.Items.Contains(kategori.ToUpper())) cmbKat.SelectedItem = kategori.ToUpper(); else cmbKat.SelectedIndex = 0;
            frmEdit.Controls.Add(lblKat); frmEdit.Controls.Add(cmbKat);

            Label lblMrk = new Label { Text = "Ürün Markası:", Location = new Point(15, 105), AutoSize = true, ForeColor = Color.DarkGray };
            ComboBox cmbMarka = new ComboBox { Location = new Point(15, 125), Width = 290, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F) };
            cmbMarka.Items.AddRange(dbMarkalar.ToArray());
            if (cmbMarka.Items.Contains(marka.ToUpper())) cmbMarka.SelectedItem = marka.ToUpper(); else cmbMarka.SelectedIndex = 0;
            frmEdit.Controls.Add(lblMrk); frmEdit.Controls.Add(cmbMarka);

            Label lblAd = new Label { Text = "Ürün Kart Adı (Model):", Location = new Point(15, 165), AutoSize = true, ForeColor = Color.DarkGray };
            TextBox txtAd = new TextBox { Location = new Point(15, 185), Width = 290, Text = urunAdi, BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) };
            frmEdit.Controls.Add(lblAd); frmEdit.Controls.Add(txtAd);

            Label lblMaliyet = new Label { Text = "Birim Alış Maliyet (TL):", Location = new Point(15, 225), AutoSize = true, ForeColor = Color.DarkGray };
            TextBox txtMaliyet = new TextBox { Location = new Point(15, 245), Width = 290, Text = maliyet.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) };

            txtMaliyet.Leave += (s, ev) => {
                string m = txtMaliyet.Text.Trim().Replace("₺", "");
                if (double.TryParse(m.Replace(".", "").Replace(",", "."), out double d))
                    txtMaliyet.Text = d.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));
            };
            frmEdit.Controls.Add(lblMaliyet); frmEdit.Controls.Add(txtMaliyet);

            Label lblFirma = new Label { Text = "Tedarikçi Toptancı Firma:", Location = new Point(15, 285), AutoSize = true, ForeColor = Color.DarkGray };
            ComboBox cmbFirma = new ComboBox { Location = new Point(15, 305), Width = 290, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F) };
            cmbFirma.Items.AddRange(dbFirmalar.ToArray());
            if (cmbFirma.Items.Contains(firma.ToUpper())) cmbFirma.SelectedItem = firma.ToUpper(); else cmbFirma.SelectedIndex = 0;
            frmEdit.Controls.Add(lblFirma); frmEdit.Controls.Add(cmbFirma);

            Label lblImei = new Label { Text = "Cihaz IMEI No:", Location = Point.Empty }; // Tasarım uyumluluğu için gizli referans
            lblImei = new Label { Text = "Cihaz IMEI No:", Location = new Point(15, 345), AutoSize = true, ForeColor = Color.DarkGray };
            TextBox txtImei = new TextBox { Location = new Point(15, 365), Width = 290, Text = (eskiImei == "-" ? "" : eskiImei), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) };
            if (eskiImei == "-") { txtImei.Enabled = false; txtImei.Text = "Bu üründe IMEI Yok"; txtImei.BackColor = Color.FromArgb(26, 26, 28); }
            frmEdit.Controls.Add(lblImei); frmEdit.Controls.Add(txtImei);

            Button btnKaydet = new Button { Text = "💾 DEĞİŞİKLİKLERİ GÜVENLE KAYDET", Location = new Point(15, 410), Width = 290, Height = 40, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };
            btnKaydet.FlatAppearance.BorderSize = 0; frmEdit.Controls.Add(btnKaydet);

            btnKaydet.Click += (sX, eX) => {
                double.TryParse(txtMaliyet.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double yeniMaliyet);
                try
                {
                    using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        conn.Open();
                        using (var tx = conn.BeginTransaction())
                        {
                            using (var cmd = new SqliteCommand("UPDATE Stoklar SET AlisFiyati=@a, Firma=@f WHERE Barkod=@b", conn, tx))
                            {
                                cmd.Parameters.AddWithValue("@a", yeniMaliyet);
                                cmd.Parameters.AddWithValue("@f", cmbFirma.Text.Trim());
                                cmd.Parameters.AddWithValue("@b", barkod);
                                cmd.ExecuteNonQuery();
                            }
                            using (var cmd = new SqliteCommand("UPDATE Urunler SET Kategori=@k, UrunAdi=@u, Marka=@m WHERE Barkod=@b", conn, tx))
                            {
                                cmd.Parameters.AddWithValue("@k", cmbKat.Text.Trim().ToUpper());
                                cmd.Parameters.AddWithValue("@u", txtAd.Text.Trim());
                                cmd.Parameters.AddWithValue("@m", cmbMarka.Text.Trim());
                                cmd.Parameters.AddWithValue("@b", barkod);
                                cmd.ExecuteNonQuery();
                            }
                            if (eskiImei != "-" && txtImei.Text.Trim() != "" && txtImei.Text.Trim() != eskiImei)
                            {
                                using (var cmd = new SqliteCommand("UPDATE Imeiler SET ImeiNo=@yeni WHERE ImeiNo=@eski AND Barkod=@b", conn, tx))
                                {
                                    cmd.Parameters.AddWithValue("@yeni", txtImei.Text.Trim()); cmd.Parameters.AddWithValue("@eski", eskiImei); cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery();
                                }
                            }
                            tx.Commit();
                        }
                    }
                    frmEdit.Close();
                    TumEnvanteriListele(); StoklariSorgula(dynTxtStokAra.Text.Trim()); OzetIstatistikleriHesapla();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            };

            frmEdit.ShowDialog();
        }
    }
}