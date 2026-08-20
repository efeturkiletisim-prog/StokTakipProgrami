using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // 🔍 SİPARİŞ SORGULAMA VE PROFESSIONAL FİNANS/STOK YÖNETİM MERKEZİ (İHTİLAF İPTALİ NÖTR 'BEKLEYEN' STATÜSÜNE BAĞLANDI)
    public partial class SayfaSiparisSorgula : UserControl
    {
        private TextBox ucSipTxtSiparisNoAra;
        private Button ucSipBtnSiparisAra;
        private ComboBox ucSipCmbDurumFiltre;
        private ComboBox ucSipCmbPlatformFiltre;
        private DataGridView ucSipDgvAramaSonuclari;
        private DataGridView ucSipDgvTumSiparisler;
        private ContextMenuStrip ucSipCmsSiparisIslem;
        private SplitContainer splitKonteyner;

        public SayfaSiparisSorgula()
        {
            InitializeComponent();
            ArayuzElemanlariniInsaEt();
        }

        private void SayfaSiparisSorgula_Load(object sender, EventArgs e)
        {
            PlatformFiltresiniDoldur();
            SiparisleriFiltreleVeListele();
            if (splitKonteyner != null) splitKonteyner.SplitterDistance = (int)(this.Height * 0.20);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            if (!this.DesignMode)
            {
                PlatformFiltresiniDoldur();
                SiparisleriFiltreleVeListele();
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (splitKonteyner != null && this.Height > 100)
            {
                splitKonteyner.SplitterDistance = (int)(this.Height * 0.20);
            }
        }

        private void PlatformFiltresiniDoldur()
        {
            if (ucSipCmbPlatformFiltre == null) return;

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("SELECT DISTINCT PlatformAd FROM Siparisler WHERE PlatformAd IS NOT NULL AND PlatformAd != '' ORDER BY PlatformAd ASC", b))
                    using (var r = cmd.ExecuteReader())
                    {
                        ucSipCmbPlatformFiltre.Items.Clear();
                        ucSipCmbPlatformFiltre.Items.Add("🌍 TÜM PLATFORMLAR");
                        while (r.Read()) ucSipCmbPlatformFiltre.Items.Add(r["PlatformAd"].ToString().ToUpper());
                    }
                }
            }
            catch { }

            if (ucSipCmbPlatformFiltre.Items.Count <= 1)
            {
                ucSipCmbPlatformFiltre.Items.AddRange(new string[] { "TRENDYOL", "HEPSİBURADA", "N11" });
            }
            ucSipCmbPlatformFiltre.SelectedIndex = 0;
        }

        private void SiparisleriFiltreleVeListele()
        {
            TumSiparisleriListele();
            SiparisNoIleSorgula(ucSipTxtSiparisNoAra.Text.Trim());
        }

        private void ArayuzElemanlariniInsaEt()
        {
            this.Controls.Clear();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(33, 33, 36);
            this.Padding = new Padding(15);

            ucSipCmsSiparisIslem = new ContextMenuStrip();
            var itemDuzenle = ucSipCmsSiparisIslem.Items.Add("✏️ Sipariş Bilgilerini Düzenle", null, BtnSiparisDuzenle_Click);
            var itemIade = ucSipCmsSiparisIslem.Items.Add("🔄 Siparişi İade Et (Maliyet Zararı)", null, BtnSiparisIade_Click);
            var itemIptal = ucSipCmsSiparisIslem.Items.Add("❌ Siparişi İptal Et (Stoka İade)", null, BtnSiparisIptal_Click);
            ucSipCmsSiparisIslem.Items.Add(new ToolStripSeparator());
            var itemIhtilaf = ucSipCmsSiparisIslem.Items.Add("⚠️ Siparişi İhtilaflı Olarak İşaretle", null, BtnSiparisIhtilafli_Click);
            var itemIhtilafKaldir = ucSipCmsSiparisIslem.Items.Add("✅ İhtilafı Kaldır (Normal Statüye Döndür)", null, BtnSiparisIhtilafIptal_Click);

            ucSipCmsSiparisIslem.Opening += (s, e) => {
                DataGridView grid = AktifGridiBul();
                if (grid == null || grid.SelectedRows.Count == 0) { e.Cancel = true; return; }

                string durum = grid.SelectedRows[0].Cells["Durum"].Value?.ToString() ?? "";

                itemIhtilaf.Visible = (durum != "İHTİLAFLI" && durum != "İPTAL" && durum != "İADE");
                itemIhtilafKaldir.Visible = (durum == "İHTİLAFLI");
                itemIade.Visible = (durum != "İADE" && durum != "İPTAL" && durum != "İHTİLAFLI");
                itemIptal.Visible = (durum != "İPTAL" && durum != "İADE" && durum != "İHTİLAFLI");
            };

            Panel pnlAramaBar = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.Transparent };

            Label lblAraBaslik = new Label { Text = "Sipariş No:", Location = new Point(5, 15), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White };
            ucSipTxtSiparisNoAra = new TextBox { Location = new Point(95, 12), Width = 150, Font = new Font("Segoe UI", 12F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            ucSipTxtSiparisNoAra.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; SiparisleriFiltreleVeListele(); } };

            Label lblDurumFiltre = new Label { Text = "Durum:", Location = new Point(260, 15), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White };
            ucSipCmbDurumFiltre = new ComboBox { Location = new Point(325, 13), Width = 175, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            ucSipCmbDurumFiltre.Items.AddRange(new string[] { "🌍 TÜMÜ", "🔄 SADECE İADELER", "❌ SADECE İPTALLER", "📦 AKTİF SİPARİŞLER", "⚠️ İHTİLAFLI SİPARİŞLER" });

            Label lblPlatFiltre = new Label { Text = "Platform:", Location = new Point(515, 15), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White };
            ucSipCmbPlatformFiltre = new ComboBox { Location = new Point(595, 13), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            ucSipCmbPlatformFiltre.Items.Add("🌍 TÜM PLATFORMLAR");

            ucSipBtnSiparisAra = new Button { Text = "🔍 SORGULA", Location = new Point(775, 11), Width = 120, Height = 28, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            ucSipBtnSiparisAra.FlatAppearance.BorderSize = 0;
            ucSipBtnSiparisAra.Click += (s, e) => SiparisleriFiltreleVeListele();

            pnlAramaBar.Controls.Add(lblAraBaslik);
            pnlAramaBar.Controls.Add(ucSipTxtSiparisNoAra);
            pnlAramaBar.Controls.Add(lblDurumFiltre);
            pnlAramaBar.Controls.Add(ucSipCmbDurumFiltre);
            pnlAramaBar.Controls.Add(lblPlatFiltre);
            pnlAramaBar.Controls.Add(ucSipCmbPlatformFiltre);
            pnlAramaBar.Controls.Add(ucSipBtnSiparisAra);
            this.Controls.Add(pnlAramaBar);

            splitKonteyner = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 6,
                BackColor = Color.FromArgb(33, 33, 36)
            };
            this.Controls.Add(splitKonteyner);
            splitKonteyner.BringToFront();

            Panel pnlUstGrid = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
            Label lblUstBaslik = new Label { Text = "🎯 SORGULAMA SONUÇLARI (Çift Tık: Detaylı Analiz Dökümü | Sağ Tık: Düzenle/İade/İptal/İhtilaf)", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange };
            ucSipDgvAramaSonuclari = GridYarat(Color.Orange);
            ucSipDgvAramaSonuclari.CellDoubleClick += UcSipDgv_CellDoubleClick;
            ucSipDgvAramaSonuclari.DataBindingComplete += Grid_DataBindingComplete;

            pnlUstGrid.Controls.Add(ucSipDgvAramaSonuclari); pnlUstGrid.Controls.Add(lblUstBaslik);
            splitKonteyner.Panel1.Controls.Add(pnlUstGrid);

            Panel pnlAltGrid = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
            Label lblAltBaslik = new Label { Text = "📋 SORGULANAN SİPARİŞ GEÇMİŞİ (Çift Tık: Detaylı Analiz Dökümü | Sağ Tık: Düzenle/İade/İptal/İhtilaf)", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) };
            ucSipDgvTumSiparisler = GridYarat(Color.FromArgb(56, 189, 248));
            ucSipDgvTumSiparisler.CellDoubleClick += UcSipDgv_CellDoubleClick;
            ucSipDgvTumSiparisler.DataBindingComplete += Grid_DataBindingComplete;

            pnlAltGrid.Controls.Add(ucSipDgvTumSiparisler); pnlAltGrid.Controls.Add(lblAltBaslik);
            splitKonteyner.Panel2.Controls.Add(pnlAltGrid);

            ucSipDgvAramaSonuclari.CellClick += (s, e) => ucSipDgvTumSiparisler.ClearSelection();
            ucSipDgvTumSiparisler.CellClick += (s, e) => ucSipDgvAramaSonuclari.ClearSelection();

            ucSipCmbDurumFiltre.SelectedIndex = 0;
            ucSipCmbDurumFiltre.SelectedIndexChanged += (s, e) => SiparisleriFiltreleVeListele();
            ucSipCmbPlatformFiltre.SelectedIndexChanged += (s, e) => SiparisleriFiltreleVeListele();
        }

        private void Grid_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv == null) return;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                string durum = row.Cells["Durum"].Value?.ToString().ToUpper();

                if (durum == "İPTAL")
                {
                    row.DefaultCellStyle.BackColor = Color.Crimson;
                    row.DefaultCellStyle.ForeColor = Color.White;
                }
                else if (durum == "İADE")
                {
                    row.DefaultCellStyle.BackColor = Color.Orange;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
                else if (durum == "İHTİLAFLI")
                {
                    row.DefaultCellStyle.BackColor = Color.Purple;
                    row.DefaultCellStyle.ForeColor = Color.White;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
                    row.DefaultCellStyle.ForeColor = Color.White;
                }
            }
        }

        private DataTable SiparisTabloSablonuYarat()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(long));
            dt.Columns.Add("Sipariş\nTarihi", typeof(string));
            dt.Columns.Add("Sipariş\nNo", typeof(string));
            dt.Columns.Add("Platform", typeof(string));
            dt.Columns.Add("Kategori", typeof(string));
            dt.Columns.Add("Barkod\nNo", typeof(string));
            dt.Columns.Add("Marka", typeof(string));
            dt.Columns.Add("Ürün\nAdı", typeof(string));
            dt.Columns.Add("Satılan\nIMEI", typeof(string));
            dt.Columns.Add("Adet", typeof(int));
            dt.Columns.Add("Birim\nSatış (TL)", typeof(double));
            dt.Columns.Add("İndirim\n(TL)", typeof(double));
            dt.Columns.Add("Komisyon\n(TL)", typeof(double));
            dt.Columns.Add("Kargo\nÜcreti", typeof(double));
            dt.Columns.Add("Hizmet\nBedeli", typeof(double));
            dt.Columns.Add("Stopaj\n(TL)", typeof(double));
            dt.Columns.Add("POS\nBedeli", typeof(double));
            dt.Columns.Add("Transfer\nTutarı", typeof(double));
            dt.Columns.Add("Maliyet\n(TL)", typeof(double));
            dt.Columns.Add("Net Kâr\n(TL)", typeof(double));
            dt.Columns.Add("Durum", typeof(string));
            return dt;
        }
        private void SiparisNoIleSorgula(string siparisNo)
        {
            try
            {
                DataTable dt = SiparisTabloSablonuYarat();

                bool durumTumu = (ucSipCmbDurumFiltre == null || ucSipCmbDurumFiltre.SelectedIndex == 0);
                bool platTumu = (ucSipCmbPlatformFiltre == null || ucSipCmbPlatformFiltre.SelectedIndex == 0);

                if (string.IsNullOrEmpty(siparisNo) && durumTumu && platTumu)
                {
                    ucSipDgvAramaSonuclari.DataSource = dt;
                    GizliIdSutununuGizle(ucSipDgvAramaSonuclari);
                    return;
                }

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT s.rowid as RowID, s.SiparisNo, s.PlatformAd, 
                               CASE WHEN UPPER(s.Barkod) = UPPER(IFNULL(u.UrunAdi, s.UrunAdi)) THEN '' ELSE s.Barkod END as Barkod,
                               IFNULL(u.UrunAdi, s.UrunAdi) as UrunAdi, IFNULL(u.Kategori, s.Kategori) as Kategori, IFNULL(u.Marka, s.Marka) as Marka,
                               IFNULL(s.Adet, 1) as Adet, s.SatisFiyati, s.PlatformIndirimi, s.Stopaj, s.PosBedeli, s.HizmetBedeli, s.KomisyonTL, s.KargoUcreti, 
                               s.SatilanImei, s.Durum, s.Tarih, IFNULL(s.AlisFiyati, 0) as AlisFiyati
                        FROM Siparisler s
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod
                        WHERE 1=1";

                    if (!string.IsNullOrEmpty(siparisNo)) sql += " AND s.SiparisNo = @sn";

                    int filtreIdx = ucSipCmbDurumFiltre != null ? ucSipCmbDurumFiltre.SelectedIndex : 0;
                    if (filtreIdx == 1) sql += " AND s.Durum = 'İADE'";
                    else if (filtreIdx == 2) sql += " AND s.Durum = 'İPTAL'";
                    else if (filtreIdx == 3) sql += " AND s.Durum NOT IN ('İADE', 'İPTAL', 'İHTİLAFLI')";
                    else if (filtreIdx == 4) sql += " AND s.Durum = 'İHTİLAFLI'";

                    if (!platTumu)
                    {
                        sql += " AND s.PlatformAd = @plat";
                    }

                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        if (!string.IsNullOrEmpty(siparisNo)) cmd.Parameters.AddWithValue("@sn", siparisNo);
                        if (!platTumu) cmd.Parameters.AddWithValue("@plat", ucSipCmbPlatformFiltre.Text);

                        using (var r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                double satis = Convert.ToDouble(r["SatisFiyati"] == DBNull.Value ? 0 : r["SatisFiyati"]);
                                double indirim = Convert.ToDouble(r["PlatformIndirimi"] == DBNull.Value ? 0 : r["PlatformIndirimi"]);
                                double alis = Convert.ToDouble(r["AlisFiyati"] == DBNull.Value ? 0 : r["AlisFiyati"]);

                                int adet = Convert.ToInt32(r["Adet"]);
                                if (adet == 0) adet = 1;

                                string durum = r["Durum"]?.ToString() ?? "";

                                double kesinti = Convert.ToDouble(r["Stopaj"] == DBNull.Value ? 0 : r["Stopaj"]) +
                                                 Convert.ToDouble(r["PosBedeli"] == DBNull.Value ? 0 : r["PosBedeli"]) +
                                                 Convert.ToDouble(r["HizmetBedeli"] == DBNull.Value ? 0 : r["HizmetBedeli"]) +
                                                 Convert.ToDouble(r["KomisyonTL"] == DBNull.Value ? 0 : r["KomisyonTL"]) +
                                                 Convert.ToDouble(r["KargoUcreti"] == DBNull.Value ? 0 : r["KargoUcreti"]);

                                double transferTutari = 0;
                                double toplamMaliyet = alis * adet;
                                double netKar = 0;

                                if (durum == "İADE")
                                {
                                    transferTutari = -kesinti; // İadede platform sadece kesinti eksiği yazar
                                    netKar = -kesinti;         // Maliyet depoya döndüğü için zarar sadece kesintilerdir
                                }
                                else if (durum != "İPTAL")
                                {
                                    transferTutari = (satis * adet) + indirim - kesinti;
                                    netKar = transferTutari - toplamMaliyet;
                                }

                                dt.Rows.Add(
                                    Convert.ToInt64(r["RowID"]),
                                    r["Tarih"]?.ToString(),
                                    r["SiparisNo"]?.ToString(),
                                    r["PlatformAd"]?.ToString(),
                                    r["Kategori"]?.ToString(),
                                    r["Barkod"]?.ToString(),
                                    r["Marka"]?.ToString(),
                                    r["UrunAdi"]?.ToString(),
                                    r["SatilanImei"]?.ToString(),
                                    adet,
                                    Math.Round(satis, 2),
                                    Math.Round(indirim, 2),
                                    Math.Round(Convert.ToDouble(r["KomisyonTL"]), 2),
                                    Math.Round(Convert.ToDouble(r["KargoUcreti"]), 2),
                                    Math.Round(Convert.ToDouble(r["HizmetBedeli"]), 2),
                                    Math.Round(Convert.ToDouble(r["Stopaj"]), 2),
                                    Math.Round(Convert.ToDouble(r["PosBedeli"]), 2),
                                    Math.Round(transferTutari, 2),
                                    Math.Round(toplamMaliyet, 2),
                                    Math.Round(netKar, 2),
                                    durum
                                );
                            }
                        }
                    }
                }
                ucSipDgvAramaSonuclari.DataSource = dt;
                GizliIdSutununuGizle(ucSipDgvAramaSonuclari);
            }
            catch { }
        }
        private void TumSiparisleriListele()
        {
            try
            {
                DataTable dt = SiparisTabloSablonuYarat();
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT s.rowid as RowID, s.SiparisNo, s.PlatformAd, 
                               CASE WHEN UPPER(s.Barkod) = UPPER(IFNULL(u.UrunAdi, s.UrunAdi)) THEN '' ELSE s.Barkod END as Barkod,
                               IFNULL(u.UrunAdi, s.UrunAdi) as UrunAdi, IFNULL(u.Kategori, s.Kategori) as Kategori, IFNULL(u.Marka, s.Marka) as Marka,
                               IFNULL(s.Adet, 1) as Adet, s.SatisFiyati, s.PlatformIndirimi, s.Stopaj, s.PosBedeli, s.HizmetBedeli, s.KomisyonTL, s.KargoUcreti, 
                               s.SatilanImei, s.Durum, s.Tarih, IFNULL(s.AlisFiyati, 0) as AlisFiyati
                        FROM Siparisler s
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod
                        WHERE 1=1";

                    int filtreIdx = ucSipCmbDurumFiltre != null ? ucSipCmbDurumFiltre.SelectedIndex : 0;
                    if (filtreIdx == 1) sql += " AND s.Durum = 'İADE'";
                    else if (filtreIdx == 2) sql += " AND s.Durum = 'İPTAL'";
                    else if (filtreIdx == 3) sql += " AND s.Durum NOT IN ('İADE', 'İPTAL', 'İHTİLAFLI')";
                    else if (filtreIdx == 4) sql += " AND s.Durum = 'İHTİLAFLI'";

                    bool platTumu = (ucSipCmbPlatformFiltre == null || ucSipCmbPlatformFiltre.SelectedIndex == 0);
                    if (!platTumu)
                    {
                        sql += " AND s.PlatformAd = @plat";
                    }

                    sql += " ORDER BY s.rowid DESC";

                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        if (!platTumu) cmd.Parameters.AddWithValue("@plat", ucSipCmbPlatformFiltre.Text);

                        using (var r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                double satis = Convert.ToDouble(r["SatisFiyati"] == DBNull.Value ? 0 : r["SatisFiyati"]);
                                double indirim = Convert.ToDouble(r["PlatformIndirimi"] == DBNull.Value ? 0 : r["PlatformIndirimi"]);
                                double alis = Convert.ToDouble(r["AlisFiyati"] == DBNull.Value ? 0 : r["AlisFiyati"]);
                                int adet = Convert.ToInt32(r["Adet"]);
                                if (adet == 0) adet = 1;
                                string durum = r["Durum"]?.ToString() ?? "";

                                double kesinti = Convert.ToDouble(r["Stopaj"] == DBNull.Value ? 0 : r["Stopaj"]) +
                                                 Convert.ToDouble(r["PosBedeli"] == DBNull.Value ? 0 : r["PosBedeli"]) +
                                                 Convert.ToDouble(r["HizmetBedeli"] == DBNull.Value ? 0 : r["HizmetBedeli"]) +
                                                 Convert.ToDouble(r["KomisyonTL"] == DBNull.Value ? 0 : r["KomisyonTL"]) +
                                                 Convert.ToDouble(r["KargoUcreti"] == DBNull.Value ? 0 : r["KargoUcreti"]);

                                double transferTutari = 0;
                                double toplamMaliyet = alis * adet;
                                double netKar = 0;

                                if (durum == "İADE")
                                {
                                    transferTutari = -kesinti;
                                    netKar = -kesinti;
                                }
                                else if (durum != "İPTAL")
                                {
                                    transferTutari = (satis * adet) + indirim - kesinti;
                                    netKar = transferTutari - toplamMaliyet;
                                }

                                dt.Rows.Add(
                                    Convert.ToInt64(r["RowID"]),
                                    r["Tarih"]?.ToString(),
                                    r["SiparisNo"]?.ToString(),
                                    r["PlatformAd"]?.ToString(),
                                    r["Kategori"]?.ToString(),
                                    r["Barkod"]?.ToString(),
                                    r["Marka"]?.ToString(),
                                    r["UrunAdi"]?.ToString(),
                                    r["SatilanImei"]?.ToString(),
                                    adet,
                                    Math.Round(satis, 2),
                                    Math.Round(indirim, 2),
                                    Math.Round(Convert.ToDouble(r["KomisyonTL"]), 2),
                                    Math.Round(Convert.ToDouble(r["KargoUcreti"]), 2),
                                    Math.Round(Convert.ToDouble(r["HizmetBedeli"]), 2),
                                    Math.Round(Convert.ToDouble(r["Stopaj"]), 2),
                                    Math.Round(Convert.ToDouble(r["PosBedeli"]), 2),
                                    Math.Round(transferTutari, 2),
                                    Math.Round(toplamMaliyet, 2),
                                    Math.Round(netKar, 2),
                                    durum
                                );
                            }
                        }
                    }
                }
                ucSipDgvTumSiparisler.DataSource = dt;
                GizliIdSutununuGizle(ucSipDgvTumSiparisler);
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
                // Sütun genişliklerini otomatik çok açmaması için ayarlandı
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                Font = new Font("Segoe UI", 9F),
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                // 🚀 Alt alta 2 satır başlığın sığması için yükseklik 48'e çıkarıldı
                ColumnHeadersHeight = 48,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ContextMenuStrip = ucSipCmsSiparisIslem
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

            // Başlıklarda hizalamanın düzgün olması için WrapMode eklendi
            dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = baslikRengi;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgv.DataError += (s, e) => { e.Cancel = true; };

            return dgv;
        }
        private void GizliIdSutununuGizle(DataGridView dgv)
        {
            if (dgv.Columns.Contains("ID")) dgv.Columns["ID"].Visible = false;

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.ValueType == typeof(double))
                {
                    col.DefaultCellStyle.Format = "N2";
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }

            dgv.ClearSelection();
        }

        private DataGridView AktifGridiBul()
        {
            if (ucSipDgvAramaSonuclari.SelectedRows.Count > 0) return ucSipDgvAramaSonuclari;
            if (ucSipDgvTumSiparisler.SelectedRows.Count > 0) return ucSipDgvTumSiparisler;
            return null;
        }

        private double SayiyiCozumle(string metin)
        {
            if (string.IsNullOrWhiteSpace(metin)) return 0;
            string s = metin.Replace("₺", "").Replace("TL", "").Replace(" ", "").Trim();
            int n = s.LastIndexOf('.'); int v = s.LastIndexOf(',');
            if (n > v) { if (s.Substring(n).Length <= 3) s = s.Substring(0, n).Replace(".", "").Replace(",", "") + "." + s.Substring(n + 1); else s = s.Replace(".", "").Replace(",", ""); }
            else if (v > n) { if (s.Substring(v).Length <= 3) s = s.Substring(0, v).Replace(".", "").Replace(",", "") + "." + s.Substring(v + 1); else s = s.Replace(".", "").Replace(",", ""); }
            else s = s.Replace(".", "").Replace(",", "");
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double rf)) return rf;
            return 0;
        }

        private void BtnSiparisIhtilafli_Click(object sender, EventArgs e)
        {
            DataGridView grid = AktifGridiBul(); if (grid == null || grid.SelectedRows.Count == 0) return;
            long rowId = Convert.ToInt64(grid.SelectedRows[0].Cells["ID"].Value);
            string sipNo = grid.SelectedRows[0].Cells["Sipariş No"].Value.ToString();
            string mevcutDurum = grid.SelectedRows[0].Cells["Durum"].Value.ToString();

            if (mevcutDurum == "İHTİLAFLI") { MessageBox.Show("Bu sipariş zaten ihtilaflı durumda kayıtlı patron!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            if (MessageBox.Show($"[{sipNo}] numaralı siparişi İHTİLAFLI duruma çekmek istiyor musunuz?\n\nBu işlem renk radarını güncelleyecek ve siparişi askıya alacaktır.", "Durum Güncellemesi", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='İHTİLAFLI' WHERE rowid=@id", b))
                        {
                            cmd.Parameters.AddWithValue("@id", rowId);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    SiparisleriFiltreleVeListele();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }
        }

        private void BtnSiparisIhtilafIptal_Click(object sender, EventArgs e)
        {
            DataGridView grid = AktifGridiBul(); if (grid == null || grid.SelectedRows.Count == 0) return;
            long rowId = Convert.ToInt64(grid.SelectedRows[0].Cells["ID"].Value);
            string sipNo = grid.SelectedRows[0].Cells["Sipariş No"].Value.ToString();

            if (MessageBox.Show($"[{sipNo}] numaralı siparişin İHTİLAFLI durumunu kaldırıp nötr 'BEKLEYEN' (Aktif) statüsüne döndürmek istiyor musunuz?", "İhtilafı Kaldır", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='BEKLEYEN' WHERE rowid=@id", b))
                        {
                            cmd.Parameters.AddWithValue("@id", rowId);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    SiparisleriFiltreleVeListele();
                    MessageBox.Show("Siparişin ihtilaf durumu kaldırıldı ve normal (bekleyen) statüye alındı!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }
        }

        private void AppSiparisIptalYonetimi(long rowId)
        {
            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
            {
                b.Open();
                using (var tx = b.BeginTransaction())
                {
                    string barkod = ""; int adet = 0; string imei = ""; string durum = "";
                    using (var cmd = new SqliteCommand("SELECT Barkod, Adet, SatilanImei, Durum FROM Siparisler WHERE rowid=@id", b, tx))
                    {
                        cmd.Parameters.AddWithValue("@id", rowId);
                        using (var r = cmd.ExecuteReader()) { if (r.Read()) { barkod = r["Barkod"].ToString(); adet = Convert.ToInt32(r["Adet"]); imei = r["SatilanImei"].ToString(); durum = r["Durum"].ToString(); } }
                    }

                    if (durum == "İPTAL") { MessageBox.Show("Bu sipariş zaten iptal edilmiş durumda!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                    if (durum == "İADE") { MessageBox.Show("İade edilmiş bir sipariş iptal edilemez patron!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                    using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='İPTAL' WHERE rowid=@id", b, tx)) { cmd.Parameters.AddWithValue("@id", rowId); cmd.ExecuteNonQuery(); }
                    using (var cmd = new SqliteCommand("UPDATE Stoklar SET Adet = Adet + @a, StokMiktari = StokMiktari + @a WHERE Barkod=@b", b, tx)) { cmd.Parameters.AddWithValue("@a", adet); cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery(); }

                    if (!string.IsNullOrEmpty(imei) && imei != "-")
                    {
                        using (var cmd = new SqliteCommand("UPDATE Imeiler SET Durum='STOKTA' WHERE ImeiNo=@i AND Barkod=@b", b, tx)) { cmd.Parameters.AddWithValue("@i", imei); cmd.Parameters.AddWithValue("@b", barkod); cmd.ExecuteNonQuery(); }
                    }
                    tx.Commit();
                }
            }
        }

        private void BtnSiparisIptal_Click(object sender, EventArgs e)
        {
            DataGridView grid = AktifGridiBul(); if (grid == null) return;
            long rowId = Convert.ToInt64(grid.SelectedRows[0].Cells["ID"].Value);
            string sipNo = grid.SelectedRows[0].Cells["Sipariş No"].Value.ToString();

            if (MessageBox.Show($"[{sipNo}] numaralı sipariş satırını iptal etmek ve ürünleri stoka geri yüklemek istiyor musunuz?", "Sipariş İptal Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    AppSiparisIptalYonetimi(rowId);
                    SiparisleriFiltreleVeListele();
                    MessageBox.Show("Sipariş başarıyla iptal edildi, envanter geri yüklendi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }
        }

        private void BtnSiparisIade_Click(object sender, EventArgs e)
        {
            DataGridView grid = AktifGridiBul(); if (grid == null) return;
            var satir = grid.SelectedRows[0];

            long rowId = Convert.ToInt64(satir.Cells["ID"].Value);
            string sipNo = satir.Cells["Sipariş No"].Value.ToString();
            string mevcutDurum = satir.Cells["Durum"].Value.ToString();

            if (mevcutDurum == "İADE") { MessageBox.Show("Bu sipariş zaten iade alınmış durumda patron!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (mevcutDurum == "İPTAL") { MessageBox.Show("İptal edilmiş bir sipariş iade alınamaz!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            Form frmIade = new Form { Text = "🔄 Müşteri İade Giriş Paneli", Size = new Size(360, 240), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MaximizeBox = false, MinimizeBox = false, FormBorderStyle = FormBorderStyle.FixedDialog };

            TextBox tIadeHizmet, tIadeKargo;
            frmIade.Controls.Add(FormElemaniYarat("İade Hizmet Bedeli Kesintisi (TL):", out tIadeHizmet, new Point(20, 15), 300, "0,00"));
            frmIade.Controls.Add(FormElemaniYarat("İade Lojistik Kargo Bedeli (TL):", out tIadeKargo, new Point(20, 75), 300, "0,00"));

            Button btnOnayla = new Button { Text = "✔ İADEYİ TAMAMLA VE STOKA İŞLE", Location = new Point(20, 145), Width = 300, Height = 40, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnOnayla.FlatAppearance.BorderSize = 0; frmIade.Controls.Add(btnOnayla);

            btnOnayla.Click += (sX, eX) => {
                double iadeHizmet = SayiyiCozumle(tIadeHizmet.Text);
                double iadeKargo = SayiyiCozumle(tIadeKargo.Text);

                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var tx = b.BeginTransaction())
                        {
                            string barkod = ""; int adet = 0; string imei = "";
                            using (var cmd = new SqliteCommand("SELECT Barkod, Adet, SatilanImei FROM Siparisler WHERE rowid=@id", b, tx))
                            {
                                cmd.Parameters.AddWithValue("@id", rowId);
                                using (var r = cmd.ExecuteReader()) { if (r.Read()) { barkod = r["Barkod"].ToString(); adet = Convert.ToInt32(r["Adet"]); imei = r["SatilanImei"].ToString(); } }
                            }

                            string upSql = "UPDATE Siparisler SET Durum='İADE', HizmetBedeli = HizmetBedeli + @ih, KargoUcreti = KargoUcreti + @ik WHERE rowid=@id";
                            using (var cmd = new SqliteCommand(upSql, b, tx))
                            {
                                cmd.Parameters.AddWithValue("@ih", iadeHizmet);
                                cmd.Parameters.AddWithValue("@ik", iadeKargo);
                                cmd.Parameters.AddWithValue("@id", rowId);
                                cmd.ExecuteNonQuery();
                            }

                            using (var cmd = new SqliteCommand("UPDATE Stoklar SET Adet = Adet + @a, StokMiktari = StokMiktari + @a WHERE Barkod=@b", b, tx))
                            {
                                cmd.Parameters.AddWithValue("@a", adet);
                                cmd.Parameters.AddWithValue("@b", barkod);
                                cmd.ExecuteNonQuery();
                            }

                            if (!string.IsNullOrEmpty(imei) && imei != "-")
                            {
                                using (var cmd = new SqliteCommand("UPDATE Imeiler SET Durum='STOKTA' WHERE ImeiNo=@i AND Barkod=@b", b, tx))
                                {
                                    cmd.Parameters.AddWithValue("@i", imei);
                                    cmd.Parameters.AddWithValue("@b", barkod);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            tx.Commit();
                        }
                    }
                    frmIade.Close();
                    SiparisleriFiltreleVeListele();
                    MessageBox.Show("Müşteri iadesi başarıyla envantere alındı ve lojistik zarar faturası sisteme mühürlendi patron!", "İade Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("İade İşlem Hatası: " + ex.Message); }
            };

            frmIade.ShowDialog();
        }

        private void BtnSiparisDuzenle_Click(object sender, EventArgs e)
        {
            DataGridView grid = AktifGridiBul(); if (grid == null) return;
            var satir = grid.SelectedRows[0];

            long rowId = Convert.ToInt64(satir.Cells["ID"].Value);
            string oldSipNo = satir.Cells["Sipariş No"].Value.ToString();
            string oldPlatform = satir.Cells["Platform"].Value.ToString();
            string oldBarkod = satir.Cells["Barkod No"].Value.ToString();
            int oldAdet = Convert.ToInt32(satir.Cells["Adet"].Value);
            double oldAlis = Convert.ToDouble(satir.Cells["Alış Fiyatı (TL)"].Value);
            double oldSatis = Convert.ToDouble(satir.Cells["Satış Fiyatı (TL)"].Value);
            double oldIndirim = Convert.ToDouble(satir.Cells["İndirim (TL)"].Value);
            double oldStopaj = Convert.ToDouble(satir.Cells["Stopaj (TL)"].Value);
            double oldPos = Convert.ToDouble(satir.Cells["POS Bedeli"].Value);
            double oldHizmet = Convert.ToDouble(satir.Cells["Hizmet Bedeli"].Value);
            double oldKomisyon = Convert.ToDouble(satir.Cells["Komisyon (TL)"].Value);
            double oldKargo = Convert.ToDouble(satir.Cells["Kargo Ücreti"].Value);
            string oldImei = satir.Cells["Satılan IMEI"].Value.ToString();
            string oldDurum = satir.Cells["Durum"].Value.ToString();
            string oldTarih = satir.Cells["Sipariş Tarihi"].Value.ToString();

            Form frm = new Form { Text = "✏️ Sipariş Düzenleme", Size = new Size(540, 520), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MaximizeBox = false, MinimizeBox = false };

            TextBox tSip; TextBox tPlat; ComboBox cDurum; TextBox tBar; TextBox tAdet; TextBox tImei;
            TextBox tSat; TextBox tInd; TextBox tKrg; TextBox tStp; TextBox tPos; TextBox tHzm; TextBox tKms; TextBox tTar; TextBox tAlis;

            frm.Controls.Add(FormElemaniYarat("Sipariş No:", out tSip, new Point(15, 15), 150, oldSipNo));
            frm.Controls.Add(FormElemaniYarat("Platform:", out tPlat, new Point(180, 15), 150, oldPlatform)); tPlat.ReadOnly = true;

            frmEditComboYarat(frm, "Sipariş Durumu:", out cDurum, new Point(345, 15), 150, new string[] { "BEKLEYEN", "OLUŞTURULDU", "TAMAMLANDI (TESLİM)", "KARGOYA VERİLECEK", "İADE", "İPTAL", "İHTİLAFLI" }, oldDurum);

            frm.Controls.Add(FormElemaniYarat("Barkod No (Değiştirilemez):", out tBar, new Point(15, 75), 160, oldBarkod)); tBar.ReadOnly = true; tBar.BackColor = Color.FromArgb(26, 26, 28);
            frm.Controls.Add(FormElemaniYarat("Adet Miktarı:", out tAdet, new Point(190, 75), 70, oldAdet.ToString()));
            frm.Controls.Add(FormElemaniYarat("Satılan IMEI Numarası:", out tImei, new Point(275, 75), 220, oldImei));
            if (oldImei == "-") tImei.ReadOnly = true;

            frm.Controls.Add(FormElemaniYarat("Alış Fiyatı (TL):", out tAlis, new Point(15, 145), 150, oldAlis.ToString("N2")));
            frm.Controls.Add(FormElemaniYarat("Satış Fiyatı (TL):", out tSat, new Point(180, 145), 150, oldSatis.ToString("N2")));
            frm.Controls.Add(FormElemaniYarat("Platform Desteği (TL):", out tInd, new Point(345, 145), 150, oldIndirim.ToString("N2")));

            frm.Controls.Add(FormElemaniYarat("Kargo Ücreti (TL):", out tKrg, new Point(15, 215), 150, oldKargo.ToString("N2")));
            frm.Controls.Add(FormElemaniYarat("Stopaj Kesintisi (TL):", out tStp, new Point(180, 215), 150, oldStopaj.ToString("N2")));
            frm.Controls.Add(FormElemaniYarat("POS Bedeli (TL):", out tPos, new Point(345, 215), 150, oldPos.ToString("N2")));

            frm.Controls.Add(FormElemaniYarat("Hizmet Bedeli (TL):", out tHzm, new Point(15, 285), 150, oldHizmet.ToString("N2")));
            frm.Controls.Add(FormElemaniYarat("Komisyon Kesintisi (TL):", out tKms, new Point(180, 285), 150, oldKomisyon.ToString("N2")));
            frm.Controls.Add(FormElemaniYarat("Sipariş Tarihi:", out tTar, new Point(345, 285), 150, oldTarih));

            Button btnKaydet = new Button { Text = "💾 HESAPLA VE SİPARİŞİ GÜNCELLE", Location = new Point(15, 370), Width = 485, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11F, FontStyle.Bold) };
            btnKaydet.FlatAppearance.BorderSize = 0; frm.Controls.Add(btnKaydet);

            btnKaydet.Click += (sX, eX) => {
                int.TryParse(tAdet.Text, out int newAdet); if (newAdet <= 0) newAdet = 1;

                double newAlis = SayiyiCozumle(tAlis.Text);
                double newSatis = SayiyiCozumle(tSat.Text);
                double newIndirim = SayiyiCozumle(tInd.Text);
                double newKargo = SayiyiCozumle(tKrg.Text);
                double newStopaj = SayiyiCozumle(tStp.Text);
                double newPos = SayiyiCozumle(tPos.Text);
                double newHizmet = SayiyiCozumle(tHzm.Text);
                double newKomisyon = SayiyiCozumle(tKms.Text);
                string newDurum = cDurum.Text;

                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var tx = b.BeginTransaction())
                        {
                            if (oldDurum != "İPTAL" && oldDurum != "İADE")
                            {
                                using (var cmd = new SqliteCommand("UPDATE Stoklar SET Adet=Adet+@a, StokMiktari=StokMiktari+@a WHERE Barkod=@b", b, tx)) { cmd.Parameters.AddWithValue("@a", oldAdet); cmd.Parameters.AddWithValue("@b", oldBarkod); cmd.ExecuteNonQuery(); }
                            }
                            if (newDurum != "İPTAL" && newDurum != "İADE")
                            {
                                using (var cmd = new SqliteCommand("UPDATE Stoklar SET Adet=MAX(0, Adet-@a), StokMiktari=MAX(0, StokMiktari-@a) WHERE Barkod=@b", b, tx)) { cmd.Parameters.AddWithValue("@a", newAdet); cmd.Parameters.AddWithValue("@b", oldBarkod); cmd.ExecuteNonQuery(); }
                            }

                            if (oldImei != "-" && oldImei != tImei.Text.Trim())
                            {
                                using (var cmd = new SqliteCommand("UPDATE Imeiler SET Durum='STOKTA' WHERE ImeiNo=@eski", b, tx)) { cmd.Parameters.AddWithValue("@eski", oldImei); cmd.ExecuteNonQuery(); }
                                using (var cmd = new SqliteCommand("UPDATE Imeiler SET Durum='SATILDI' WHERE ImeiNo=@yeni", b, tx)) { cmd.Parameters.AddWithValue("@yeni", tImei.Text.Trim()); cmd.ExecuteNonQuery(); }
                            }
                            if ((newDurum == "İPTAL" || newDurum == "İADE") && tImei.Text.Trim() != "-")
                            {
                                using (var cmd = new SqliteCommand("UPDATE Imeiler SET Durum='STOKTA' WHERE ImeiNo=@i", b, tx)) { cmd.Parameters.AddWithValue("@i", tImei.Text.Trim()); cmd.ExecuteNonQuery(); }
                            }

                            string upSql = @"
                                UPDATE Siparisler 
                                SET SiparisNo=@sn, Adet=@adet, AlisFiyati=@af, SatisFiyati=@sf, PlatformIndirimi=@pi, Stopaj=@st, 
                                    PosBedeli=@pb, HizmetBedeli=@hb, KomisyonTL=@km, KargoUcreti=@kr, SatilanImei=@im, 
                                    Durum=@dur, Tarih=@dt 
                                WHERE rowid=@id";

                            using (var cmd = new SqliteCommand(upSql, b, tx))
                            {
                                cmd.Parameters.AddWithValue("@sn", tSip.Text.Trim()); cmd.Parameters.AddWithValue("@adet", newAdet);
                                cmd.Parameters.AddWithValue("@af", newAlis);
                                cmd.Parameters.AddWithValue("@sf", newSatis); cmd.Parameters.AddWithValue("@pi", newIndirim);
                                cmd.Parameters.AddWithValue("@st", newStopaj); cmd.Parameters.AddWithValue("@pb", newPos);
                                cmd.Parameters.AddWithValue("@hb", newHizmet); cmd.Parameters.AddWithValue("@km", newKomisyon);
                                cmd.Parameters.AddWithValue("@kr", newKargo);
                                cmd.Parameters.AddWithValue("@im", tImei.Text.Trim());
                                cmd.Parameters.AddWithValue("@dur", newDurum); cmd.Parameters.AddWithValue("@dt", tTar.Text.Trim());
                                cmd.Parameters.AddWithValue("@id", rowId);
                                cmd.ExecuteNonQuery();
                            }
                            tx.Commit();
                        }
                    }
                    frm.Close();

                    // 🚀 MÜHÜR: Kaydetme sonrası aramanın silinmesi engellendi! PlatformFiltresiniDoldur metodu devredışı bırakıldı.
                    SiparisleriFiltreleVeListele();
                    MessageBox.Show("Sipariş finansalları ve stok balansı başarıyla güncellendi patron!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Hesaplama Hatası: " + ex.Message); }
            };

            frm.ShowDialog();
        }

        private void UcSipDgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            DataGridView dgv = sender as DataGridView;
            if (dgv == null) return;

            var satir = dgv.Rows[e.RowIndex];
            string sipNo = satir.Cells["Sipariş No"].Value.ToString();
            string platform = satir.Cells["Platform"].Value.ToString();
            string barkod = satir.Cells["Barkod No"].Value.ToString();
            string urunAdi = satir.Cells["Ürün Adı"].Value.ToString();
            string kat = satir.Cells["Kategori"]?.Value.ToString() ?? "-";
            string marka = satir.Cells["Marka"]?.Value.ToString() ?? "-";
            int adet = Convert.ToInt32(satir.Cells["Adet"].Value);

            double alis = Convert.ToDouble(satir.Cells["Alış Fiyatı (TL)"].Value);
            double satis = Convert.ToDouble(satir.Cells["Satış Fiyatı (TL)"].Value);
            double indirim = Convert.ToDouble(satir.Cells["İndirim (TL)"].Value);
            double stopaj = Convert.ToDouble(satir.Cells["Stopaj (TL)"].Value);
            double pos = Convert.ToDouble(satir.Cells["POS Bedeli"].Value);
            double hizmet = Convert.ToDouble(satir.Cells["Hizmet Bedeli"].Value);
            double komisyon = Convert.ToDouble(satir.Cells["Komisyon (TL)"].Value);
            double kargo = Convert.ToDouble(satir.Cells["Kargo Ücreti"].Value);
            string imei = satir.Cells["Satılan IMEI"].Value.ToString();
            double netKar = Convert.ToDouble(satir.Cells["Net Kâr (TL)"].Value);
            string durum = satir.Cells["Durum"].Value.ToString();
            string tarih = satir.Cells["Sipariş Tarihi"].Value.ToString();

            Form frmDetay = new Form
            {
                Text = $"📄 Detaylı Sipariş Analiz Raporu - No: {sipNo}",
                Size = new Size(720, 520),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.FromArgb(33, 33, 36),
                ForeColor = Color.White,
                ShowIcon = false,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new Font("Segoe UI", 10F)
            };

            GroupBox grpGenel = new GroupBox { Text = "📦 Genel ve Lojistik Bilgiler", Location = new Point(20, 15), Width = 330, Height = 360, ForeColor = Color.FromArgb(56, 189, 248) };
            grpGenel.Controls.Add(new Label { Text = $"Sipariş No:\n{sipNo}", Location = new Point(15, 25), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White });
            grpGenel.Controls.Add(new Label { Text = $"Satış Kanalı: {platform}", Location = new Point(15, 75), AutoSize = true, ForeColor = Color.LightGray });
            grpGenel.Controls.Add(new Label { Text = $"Sipariş Tarihi: {tarih}", Location = new Point(15, 100), AutoSize = true, ForeColor = Color.LightGray });

            Label lblDurum = new Label { Text = $"SİSTEM DURUMU: {durum}", Location = new Point(15, 135), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = (durum == "İPTAL" || durum == "İADE") ? Color.Crimson : (durum == "İHTİLAFLI" ? Color.MediumPurple : Color.Gold) };
            grpGenel.Controls.Add(lblDurum);
            grpGenel.Controls.Add(new Label { Text = "──────────────────────────", Location = new Point(15, 165), AutoSize = true, ForeColor = Color.FromArgb(50, 50, 50) });
            grpGenel.Controls.Add(new Label { Text = $"Ürün Kart Adı: {urunAdi}", Location = new Point(15, 190), Width = 300, Height = 40, ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) });
            grpGenel.Controls.Add(new Label { Text = $"Kategori: {kat}", Location = new Point(15, 240), AutoSize = true, ForeColor = Color.DarkGray });
            grpGenel.Controls.Add(new Label { Text = $"Marka: {marka}", Location = new Point(15, 265), AutoSize = true, ForeColor = Color.DarkGray });

            string temizBarkod = (string.Equals(barkod, urunAdi.ToUpper()) || barkod == "-") ? "-" : barkod;
            grpGenel.Controls.Add(new Label { Text = $"Barkod No: {temizBarkod}", Location = new Point(15, 290), AutoSize = true, ForeColor = Color.DarkGray });
            grpGenel.Controls.Add(new Label { Text = $"Cihaz IMEI: {imei}", Location = new Point(15, 315), AutoSize = true, ForeColor = imei == "-" ? Color.DarkGray : Color.Coral, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) });
            frmDetay.Controls.Add(grpGenel);

            GroupBox grpFinans = new GroupBox { Text = "💰 Finansal Hesap Dökümü", Location = new Point(365, 15), Width = 320, Height = 360, ForeColor = Color.FromArgb(34, 197, 94) };
            int y = 30;
            Action<string, double, Color> FinansSatiriEkle = (baslik, dgr, rnk) => {
                Label lblB = new Label { Text = baslik, Location = new Point(15, y), AutoSize = true, ForeColor = Color.LightGray };
                Label lblD = new Label { Text = dgr.ToString("N2") + " TL", Location = new Point(200, y), Width = 100, TextAlign = ContentAlignment.MiddleRight, ForeColor = rnk, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
                grpFinans.Controls.Add(lblB);
                grpFinans.Controls.Add(lblD);
                y += 28;
            };

            if (durum == "İADE")
            {
                FinansSatiriEkle("Brüt Satış (İADE EDİLDİ):", 0, Color.White);
                FinansSatiriEkle("Platform Desteği (İPTAL):", 0, Color.LimeGreen);
                FinansSatiriEkle($"Alış Maliyet Toplamı (İADE):", 0, Color.LightSkyBlue);
            }
            else
            {
                FinansSatiriEkle($"Satış Toplamı:", satis, Color.White);
                FinansSatiriEkle("Platform Desteği (Gelir) (+):", indirim, Color.LimeGreen);
                FinansSatiriEkle($"Alış Maliyet Toplamı (-):", alis, Color.LightSkyBlue);
            }

            grpFinans.Controls.Add(new Label { Text = "────────────────────────", Location = new Point(15, y - 5), AutoSize = true, ForeColor = Color.FromArgb(50, 50, 50) }); y += 15;

            FinansSatiriEkle("Pazaryeri Komisyonu:", komisyon, Color.FromArgb(239, 68, 68));
            FinansSatiriEkle("Lojistik Kargo Ücreti:", kargo, Color.FromArgb(239, 68, 68));
            FinansSatiriEkle("Güvence / POS Kesintisi:", pos, Color.FromArgb(239, 68, 68));
            FinansSatiriEkle("Sabit Hizmet Bedeli:", hizmet, Color.FromArgb(239, 68, 68));
            FinansSatiriEkle("Devlet Stopaj Vergisi:", stopaj, Color.FromArgb(239, 68, 68));
            frmDetay.Controls.Add(grpFinans);

            Panel pnlKarKutusu = new Panel { Location = new Point(20, 390), Width = 665, Height = 65, BackColor = Color.FromArgb(26, 26, 39), BorderStyle = BorderStyle.FixedSingle };
            Label lblKarYazi = new Label { Text = (durum == "İADE" || netKar < 0) ? "BU SİPARİŞ SATIRININ TOPLAM NET ZARARI:" : "BU SİPARİŞ SATIRININ NET KÂR / ZARAR DURUMU:", Location = new Point(15, 22), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.DarkGray };
            Label lblKarDeger = new Label { Text = netKar.ToString("N2") + " TL", Location = new Point(380, 12), Width = 260, Height = 40, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI Black", 16F, FontStyle.Bold), ForeColor = netKar < 0 ? Color.Crimson : Color.FromArgb(34, 197, 94) };

            pnlKarKutusu.Controls.Add(lblKarYazi); pnlKarKutusu.Controls.Add(lblKarDeger);
            frmDetay.Controls.Add(pnlKarKutusu);

            frmDetay.ShowDialog();
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 55 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            p.Controls.Add(lbl); p.Controls.Add(kutu);
            return p;
        }

        private void frmEditComboYarat(Form f, string baslik, out ComboBox combo, Point konum, int genislik, string[] ogeler, string varsayilan)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 55 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            combo = new ComboBox { Location = new Point(0, 18), Width = genislik, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };
            combo.Items.AddRange(ogeler);
            if (combo.Items.Contains(varsayilan)) combo.SelectedItem = varsayilan; else combo.SelectedIndex = 0;
            p.Controls.Add(lbl); p.Controls.Add(combo);
            f.Controls.Add(p);
        }
    }
}