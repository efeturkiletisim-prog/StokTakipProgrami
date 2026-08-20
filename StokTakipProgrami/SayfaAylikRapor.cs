using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // 📅 AYLIK RAPORLAMA - %100 KUSURSUZ GÖRSEL MATEMATİK, PAZARYERİ CEZA VE "TAHSİLAT" SİSTEMİ ENTEGRASYONU (RENK RADARI DAHİL)
    public partial class SayfaAylikRapor : UserControl
    {
        private DataGridView dynDgvAylik;
        private DataGridView dynDgvPlatform;
        private Label dynLblBuAyCiro, dynLblBuAyKar, dynLblBuAySiparis, dynLblIadeAdeti;
        private Label dynLblOdenenAdeti, dynLblBekleyenAdeti;

        private ComboBox cmbPlatformSecimi, cmbAySecimi, cmbYilSecimi;
        private Button btnRaporla;

        public SayfaAylikRapor()
        {
            InitializeComponent();
            this.Load += SayfaAylikRapor_Load;
        }

        private void SayfaAylikRapor_Load(object sender, EventArgs e)
        {
            ArayuzInsaEt();
            FiltreleriDoldur();
            AylikVerileriHesapla();
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "📅 AYLIK PERFORMANS VE FİNANSAL DAĞILIM ANALİZ MERKEZİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            Panel pnlFiltreBar = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(26, 26, 39) };

            cmbPlatformSecimi = ComboYarat(pnlFiltreBar, "Platform:", new Point(15, 8), 160);
            cmbAySecimi = ComboYarat(pnlFiltreBar, "Ay Seçimi:", new Point(190, 8), 130);
            cmbYilSecimi = ComboYarat(pnlFiltreBar, "Yıl:", new Point(335, 8), 90);

            btnRaporla = new Button { Text = "📊 RAPORU ANALİZ ET", Location = new Point(440, 24), Width = 160, Height = 28, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnRaporla.FlatAppearance.BorderSize = 0;
            btnRaporla.Click += (s, e) => AylikVerileriHesapla();
            pnlFiltreBar.Controls.Add(btnRaporla);
            this.Controls.Add(pnlFiltreBar);

            TableLayoutPanel pnlWidgets = new TableLayoutPanel { Dock = DockStyle.Top, Height = 85, BackColor = Color.Transparent, ColumnCount = 4, RowCount = 1, Padding = new Padding(5) };
            pnlWidgets.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            pnlWidgets.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            pnlWidgets.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            pnlWidgets.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            this.Controls.Add(pnlWidgets);

            Panel pSatis = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Margin = new Padding(8) };
            pSatis.Controls.Add(new Panel { Dock = DockStyle.Left, Width = 5, BackColor = Color.FromArgb(56, 189, 248) });
            pSatis.Controls.Add(new Label { Text = "🛍️ SATIŞ VE DURUM DAĞILIMI", Dock = DockStyle.Top, Height = 22, Padding = new Padding(8, 5, 0, 0), Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.DarkGray });

            dynLblBuAySiparis = new Label { Text = "0 Adet", Location = new Point(13, 25), AutoSize = true, Font = new Font("Segoe UI Black", 14F, FontStyle.Bold), ForeColor = Color.White };
            pSatis.Controls.Add(dynLblBuAySiparis);

            FlowLayoutPanel flpDurum = new FlowLayoutPanel { Location = new Point(13, 53), AutoSize = true, FlowDirection = FlowDirection.LeftToRight, BackColor = Color.Transparent };
            dynLblOdenenAdeti = new Label { Text = "✔ 0 Ödenen", AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(34, 197, 94), Margin = new Padding(0, 0, 10, 0) };
            dynLblBekleyenAdeti = new Label { Text = "⏳ 0 Bekleyen", AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.Orange, Margin = new Padding(0) };

            flpDurum.Controls.Add(dynLblOdenenAdeti);
            flpDurum.Controls.Add(dynLblBekleyenAdeti);
            pSatis.Controls.Add(flpDurum);

            pnlWidgets.Controls.Add(pSatis, 0, 0);

            dynLblIadeAdeti = WidgetYarat(pnlWidgets, "🔄 TOPLAM İADE ADETİ", Color.Crimson, 1);
            dynLblBuAyCiro = WidgetYarat(pnlWidgets, "🔥 TOPLAM BRÜT CİRO", Color.Orange, 2);
            dynLblBuAyKar = WidgetYarat(pnlWidgets, "👑 ELDE EDİLEN NET KÂR", Color.FromArgb(34, 197, 94), 3);

            Panel pnlGridTasiyici = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15, 10, 15, 15) };

            TableLayoutPanel tlpSplit = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            tlpSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55F));
            tlpSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));

            GroupBox grpSol = new GroupBox { Dock = DockStyle.Fill, Text = "📅 Dönemsel / Ürün Bazlı Döküm (Çift Tıkla Fatura Gör)", ForeColor = Color.Orange, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Padding = new Padding(8) };
            dynDgvAylik = GridYarat();
            dynDgvAylik.CellDoubleClick += DynDgv_CellDoubleClick;
            dynDgvAylik.DataBindingComplete += Grid_DataBindingComplete;
            grpSol.Controls.Add(dynDgvAylik);

            GroupBox grpSag = new GroupBox { Dock = DockStyle.Fill, Text = "🌐 Platform Satış ve Kârlılık Dağılımı", ForeColor = Color.FromArgb(56, 189, 248), Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Padding = new Padding(8) };
            dynDgvPlatform = GridYarat();
            grpSag.Controls.Add(dynDgvPlatform);

            tlpSplit.Controls.Add(grpSol, 0, 0);
            tlpSplit.Controls.Add(grpSag, 1, 0);
            pnlGridTasiyici.Controls.Add(tlpSplit);
            this.Controls.Add(pnlGridTasiyici);

            pnlGridTasiyici.BringToFront();
            pnlWidgets.SendToBack();
            pnlFiltreBar.SendToBack();
            pnlHeader.SendToBack();
        }

        private DataGridView GridYarat()
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
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 35 // 🚀 MÜHÜR: 152 HATAYA SEBEP OLAN O YABANCI KELİME BURADAN TAMAMEN TEMIZLENDI!
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgv.DataError += (s, e) => { e.Cancel = true; };
            return dgv;
        }

        private void Grid_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv == null || !dgv.Columns.Contains("Durum")) return;

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

        private void FiltreleriDoldur()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("SELECT DISTINCT PlatformAd FROM Siparisler WHERE PlatformAd IS NOT NULL ORDER BY PlatformAd ASC", b))
                    using (var r = cmd.ExecuteReader())
                    {
                        cmbPlatformSecimi.Items.Clear(); cmbPlatformSecimi.Items.Add("🌍 TÜM PLATFORMLAR");
                        while (r.Read()) cmbPlatformSecimi.Items.Add(r["PlatformAd"].ToString());
                    }
                }
            }
            catch { }
            if (cmbPlatformSecimi.Items.Count <= 1) cmbPlatformSecimi.Items.AddRange(new string[] { "EFETÜRK", "TRENDYOL", "HEPSİBURADA" });
            cmbPlatformSecimi.SelectedIndex = 0;

            cmbAySecimi.Items.Clear();
            cmbAySecimi.Items.AddRange(new string[] { "📅 TÜM AYLAR", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık" });
            cmbAySecimi.SelectedIndex = DateTime.Now.Month;

            cmbYilSecimi.Items.Clear();
            cmbYilSecimi.Items.Add("2026"); cmbYilSecimi.Items.Add("2025"); cmbYilSecimi.Items.Add("2024");
            cmbYilSecimi.SelectedIndex = 0;
        }

        private void AylikVerileriHesapla()
        {
            try
            {
                string seciliPlat = cmbPlatformSecimi.SelectedIndex > 0 ? cmbPlatformSecimi.Text : "";
                string seciliAyStr = cmbAySecimi.SelectedIndex > 0 ? cmbAySecimi.SelectedIndex.ToString("D2") : "";
                string seciliYil = cmbYilSecimi.Text;

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();

                    string sqlKpi = @"
                        SELECT 
                            IFNULL(SUM(CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END), 0) as ToplamSiparis,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END), 0) as IadeAdeti,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END), 0) as ToplamCiro,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE IFNULL(s.PlatformIndirimi, 0) END), 0) as ToplamIndirim,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.AlisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END), 0) as ToplamMaliyet,
                            IFNULL(SUM(IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0)), 0) as ToplamKesinti,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum) IN ('ÖDENDİ', 'TAHSİL EDİLDİ') THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END), 0) as OdenenAdet
                        FROM Siparisler s
                        WHERE UPPER(s.Durum) != 'İPTAL' AND SUBSTR(s.Tarih, 7, 4) = @yil";

                    string sqlKpiDinamik = sqlKpi;
                    if (!string.IsNullOrEmpty(seciliPlat)) sqlKpiDinamik += " AND s.PlatformAd = @plat";
                    if (!string.IsNullOrEmpty(seciliAyStr)) sqlKpiDinamik += " AND SUBSTR(s.Tarih, 4, 2) = @ay";

                    using (var cmdReal = new SqliteCommand(sqlKpiDinamik, conn))
                    {
                        cmdReal.Parameters.AddWithValue("@yil", seciliYil);
                        if (!string.IsNullOrEmpty(seciliPlat)) cmdReal.Parameters.AddWithValue("@plat", seciliPlat);
                        if (!string.IsNullOrEmpty(seciliAyStr)) cmdReal.Parameters.AddWithValue("@ay", seciliAyStr);

                        using (var r = cmdReal.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                int kpiSatisAdet = Convert.ToInt32(r["ToplamSiparis"]);
                                int kpiIadeAdet = Convert.ToInt32(r["IadeAdeti"]);
                                int kpiOdenen = Convert.ToInt32(r["OdenenAdet"]);

                                int kpiBekleyen = kpiSatisAdet - kpiOdenen - kpiIadeAdet;

                                double kpiCiro = Convert.ToDouble(r["ToplamCiro"]);
                                double kpiIndirim = Convert.ToDouble(r["ToplamIndirim"]);
                                double kpiMaliyet = Convert.ToDouble(r["ToplamMaliyet"]);
                                double kpiKesinti = Convert.ToDouble(r["ToplamKesinti"]);

                                // 🚀 MÜHÜR: NET KÂR MATEMATİĞİ TERTEMİZ EKSİ ÇİZGİSİNE ÇEKİLDİ
                                double kpiKar = (kpiCiro + kpiIndirim - kpiKesinti) - kpiMaliyet;

                                dynLblBuAySiparis.Text = kpiSatisAdet.ToString() + " Adet";
                                dynLblOdenenAdeti.Text = $"✔ {kpiOdenen} Ödenen";
                                dynLblBekleyenAdeti.Text = $"⏳ {kpiBekleyen} Bekleyen";

                                dynLblIadeAdeti.Text = kpiIadeAdet.ToString() + " Adet";
                                dynLblBuAyCiro.Text = kpiCiro.ToString("N2") + " TL";
                                dynLblBuAyKar.Text = kpiKar.ToString("N2") + " TL";
                                dynLblBuAyKar.ForeColor = kpiKar < 0 ? Color.Crimson : Color.FromArgb(34, 197, 94);
                            }
                        }
                    }

                    DataTable dtGrid = new DataTable();

                    if (string.IsNullOrEmpty(seciliAyStr))
                    {
                        dtGrid.Columns.Add("Dönem / Ay");
                        dtGrid.Columns.Add("Satış Adet", typeof(int));
                        dtGrid.Columns.Add("İade Adet", typeof(int));
                        dtGrid.Columns.Add("Brüt Ciro", typeof(double));
                        dtGrid.Columns.Add("Transfer Tutarı", typeof(double));
                        dtGrid.Columns.Add("Toplam Maliyet", typeof(double));
                        dtGrid.Columns.Add("Net Kâr", typeof(double));
                        string sqlGroup = @"
                            SELECT 
                                SUBSTR(s.Tarih, 4, 2) as AyNumara,
                                SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) END) as SatisAdeti,
                                SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END) as IadeAdeti,
                                SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END) as ToplamCiro,
                                SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE IFNULL(s.PlatformIndirimi, 0) END) as ToplamIndirim,
                                SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.AlisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END) as ToplamMaliyet,
                                SUM(IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0)) as ToplamKesinti
                            FROM Siparisler s
                            WHERE UPPER(s.Durum) != 'İPTAL' AND SUBSTR(s.Tarih, 7, 4) = @yil";

                        if (!string.IsNullOrEmpty(seciliPlat)) sqlGroup += " AND s.PlatformAd = @plat";
                        sqlGroup += " GROUP BY AyNumara ORDER BY SUBSTR(s.Tarih, 4, 2) ASC";

                        using (var cmdG = new SqliteCommand(sqlGroup, conn))
                        {
                            cmdG.Parameters.AddWithValue("@yil", seciliYil);
                            if (!string.IsNullOrEmpty(seciliPlat)) cmdG.Parameters.AddWithValue("@plat", seciliPlat);

                            string[] ayIsimleri = { "", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık" };
                            using (var rG = cmdG.ExecuteReader())
                            {
                                while (rG.Read())
                                {
                                    int aIndeks = Convert.ToInt32(rG["AyNumara"]);
                                    int sAdet = Convert.ToInt32(rG["SatisAdeti"]);
                                    int iAdet = Convert.ToInt32(rG["IadeAdeti"]);
                                    double gCiro = Convert.ToDouble(rG["ToplamCiro"]);
                                    double gIndirim = Convert.ToDouble(rG["ToplamIndirim"]);
                                    double gMaliyet = Convert.ToDouble(rG["ToplamMaliyet"]);
                                    double gKesinti = Convert.ToDouble(rG["ToplamKesinti"]);

                                    // 🚀 MÜHÜR: GRUP TABLOSUNDA NET KÂR DÜZELTİLDİ
                                    double gTransfer = gCiro + gIndirim - gKesinti;
                                    dtGrid.Rows.Add(ayIsimleri[aIndeks] + " " + seciliYil, sAdet, iAdet, gCiro, gTransfer, gMaliyet, gTransfer - gMaliyet);
                                }
                            }
                        }
                    }
                    else
                    {
                        dtGrid.Columns.Add("RowID", typeof(long));
                        dtGrid.Columns.Add("Tarih");
                        dtGrid.Columns.Add("Platform");
                        dtGrid.Columns.Add("Sipariş No");
                        dtGrid.Columns.Add("Ürün Modeli");
                        dtGrid.Columns.Add("Durum");
                        dtGrid.Columns.Add("Adet", typeof(int));
                        dtGrid.Columns.Add("Alış Fiyatı", typeof(double));
                        dtGrid.Columns.Add("İndirim", typeof(double));
                        dtGrid.Columns.Add("Satış Fiyatı", typeof(double));
                        dtGrid.Columns.Add("Net Kâr", typeof(double));

                        string sqlDeep = @"
                            SELECT 
                                s.rowid as RowID, s.Tarih, s.PlatformAd, s.SiparisNo, IFNULL(u.UrunAdi, s.UrunAdi) as GecerliUrunAdi, s.Durum, 
                                (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) as Adet,
                                IFNULL(s.AlisFiyati, 0) as AlisFiyati, IFNULL(s.PlatformIndirimi, 0) as PlatformIndirimi, IFNULL(s.SatisFiyati, 0) as SatisFiyati,
                                IFNULL(s.Stopaj, 0) as Stopaj, IFNULL(s.PosBedeli, 0) as PosBedeli, IFNULL(s.HizmetBedeli, 0) as HizmetBedeli, IFNULL(s.KomisyonTL, 0) as KomisyonTL, IFNULL(s.KargoUcreti, 0) as KargoUcreti,
                                IFNULL(s.CezaTutari, 0) as CezaTutari
                            FROM Siparisler s
                            LEFT JOIN Urunler u ON s.Barkod = u.Barkod
                            WHERE UPPER(s.Durum) != 'İPTAL' AND SUBSTR(s.Tarih, 7, 4) = @yil AND SUBSTR(s.Tarih, 4, 2) = @ay";

                        if (!string.IsNullOrEmpty(seciliPlat)) sqlDeep += " AND s.PlatformAd = @plat";
                        sqlDeep += " ORDER BY SUBSTR(s.Tarih,7,4) ASC, SUBSTR(s.Tarih,4,2) ASC, SUBSTR(s.Tarih,1,2) ASC, s.rowid ASC";

                        using (var cmdD = new SqliteCommand(sqlDeep, conn))
                        {
                            cmdD.Parameters.AddWithValue("@yil", seciliYil);
                            cmdD.Parameters.AddWithValue("@ay", seciliAyStr);
                            if (!string.IsNullOrEmpty(seciliPlat)) cmdD.Parameters.AddWithValue("@plat", seciliPlat);

                            using (var rD = cmdD.ExecuteReader())
                            {
                                while (rD.Read())
                                {
                                    double alis = Convert.ToDouble(rD["AlisFiyati"]);
                                    double satis = Convert.ToDouble(rD["SatisFiyati"]);
                                    double indirim = Convert.ToDouble(rD["PlatformIndirimi"]);
                                    string durum = rD["Durum"].ToString().ToUpper();
                                    int adet = Convert.ToInt32(rD["Adet"]);

                                    double kargo = Convert.ToDouble(rD["KargoUcreti"]);
                                    double komisyon = Convert.ToDouble(rD["KomisyonTL"]);
                                    double hizmet = Convert.ToDouble(rD["HizmetBedeli"]);
                                    double pos = Convert.ToDouble(rD["PosBedeli"]);
                                    double stopaj = Convert.ToDouble(rD["Stopaj"]);
                                    double ceza = Convert.ToDouble(rD["CezaTutari"]);

                                    // 🚀 MÜHÜR: DERİN LİSTEDE KÂR-ZARAR MATEMATİĞİ DÜZELTİLDİ
                                    double netKar = durum == "İADE" ? -(kargo + komisyon + hizmet + pos + stopaj + ceza) : ((satis * adet) + indirim - (alis * adet) - (kargo + komisyon + hizmet + pos + stopaj + ceza));

                                    dtGrid.Rows.Add(
                                        Convert.ToInt64(rD["RowID"]), rD["Tarih"].ToString(), rD["PlatformAd"].ToString(), rD["SiparisNo"].ToString(),
                                        rD["GecerliUrunAdi"]?.ToString() ?? "Bilinmeyen Ürün", durum, adet,
                                        alis, indirim, satis, netKar
                                    );
                                }
                            }
                        }
                    }

                    dynDgvAylik.DataSource = dtGrid;
                    if (dynDgvAylik.Columns.Contains("RowID")) dynDgvAylik.Columns["RowID"].Visible = false;

                    DataTable dtPlat = new DataTable();
                    dtPlat.Columns.Add("Platform");
                    dtPlat.Columns.Add("Satış Adet", typeof(int));
                    dtPlat.Columns.Add("Ödenen", typeof(int));
                    dtPlat.Columns.Add("Bekleyen", typeof(int));
                    dtPlat.Columns.Add("İade", typeof(int));
                    dtPlat.Columns.Add("Brüt Ciro", typeof(double));
                    dtPlat.Columns.Add("Transfer Tutarı", typeof(double));
                    dtPlat.Columns.Add("Toplam Maliyet", typeof(double));
                    dtPlat.Columns.Add("Net Kâr", typeof(double));

                    string sqlPlat = @"
                        SELECT 
                            s.PlatformAd, 
                            SUM(CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) as ToplamSiparis,
                            SUM(CASE WHEN UPPER(s.Durum) IN ('ÖDENDİ', 'TAHSİL EDİLDİ') THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END) as OdenenAdet,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END) as IadeAdeti,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END) as ToplamCiro,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE IFNULL(s.PlatformIndirimi, 0) END) as ToplamIndirim,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.AlisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END) as ToplamMaliyet,
                            SUM(IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0)) as ToplamKesinti
                        FROM Siparisler s
                        WHERE UPPER(s.Durum) != 'İPTAL' AND SUBSTR(s.Tarih, 7, 4) = @yil";

                    if (!string.IsNullOrEmpty(seciliAyStr)) sqlPlat += " AND SUBSTR(s.TAG, 4, 2) = @ay";
                    if (!string.IsNullOrEmpty(seciliAyStr)) sqlPlat = sqlPlat.Replace("s.TAG", "s.Tarih");
                    if (!string.IsNullOrEmpty(seciliPlat)) sqlPlat += " AND s.PlatformAd = @plat";
                    sqlPlat += " GROUP BY s.PlatformAd ORDER BY SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END) DESC";

                    using (var cmdX = new SqliteCommand(sqlPlat, conn))
                    {
                        cmdX.Parameters.AddWithValue("@yil", seciliYil);
                        if (!string.IsNullOrEmpty(seciliAyStr)) cmdX.Parameters.AddWithValue("@ay", seciliAyStr);
                        if (!string.IsNullOrEmpty(seciliPlat)) cmdX.Parameters.AddWithValue("@plat", seciliPlat);

                        using (var rP = cmdX.ExecuteReader())
                        {
                            while (rP.Read())
                            {
                                double gCiro = Convert.ToDouble(rP["ToplamCiro"]);
                                double gIndirim = Convert.ToDouble(rP["ToplamIndirim"]);
                                double gMaliyet = Convert.ToDouble(rP["ToplamMaliyet"]);
                                double gKesinti = Convert.ToDouble(rP["ToplamKesinti"]);

                                int topSiparis = Convert.ToInt32(rP["ToplamSiparis"]);
                                int oAdet = Convert.ToInt32(rP["OdenenAdet"]);
                                int iAdet = Convert.ToInt32(rP["IadeAdeti"]);

                                // 🚀 MÜHÜR: SAĞ TABLODA DA KÂRLILIK REEL SEVİYEDE
                                double pTransfer = gCiro + gIndirim - gKesinti;
                                dtPlat.Rows.Add(
                                    rP["PlatformAd"].ToString(), topSiparis, oAdet, (topSiparis - oAdet - iAdet), iAdet, gCiro, pTransfer, gMaliyet, pTransfer - gMaliyet
                                );
                            }
                        }
                    }

                    dynDgvPlatform.DataSource = dtPlat;

                    foreach (DataGridViewColumn col in dynDgvAylik.Columns) if (col.ValueType == typeof(double)) { col.DefaultCellStyle.Format = "N2"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; }
                    foreach (DataGridViewColumn col in dynDgvPlatform.Columns) if (col.ValueType == typeof(double)) { col.DefaultCellStyle.Format = "N2"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; }
                }
            }
            catch { }
        }
        private void DynDgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            DataGridView dgv = sender as DataGridView;
            if (dgv == null || !dgv.Columns.Contains("RowID")) return;

            long rowId = Convert.ToInt64(dgv.Rows[e.RowIndex].Cells["RowID"].Value);

            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
            {
                b.Open();
                string sqlPopup = @"
                    SELECT 
                        s.SiparisNo, s.PlatformAd, s.Barkod, s.Adet, 
                        s.SatisFiyati, s.PlatformIndirimi, s.Stopaj, s.PosBedeli, 
                        s.HizmetBedeli, s.KomisyonTL, s.KargoUcreti, s.SatilanImei, 
                        s.Durum, s.Tarih, s.AlisFiyati,
                        IFNULL(s.CezaTutari, 0) as CezaTutari, IFNULL(s.CezaSebebi, '') as CezaSebebi,
                        IFNULL(u.UrunAdi, s.UrunAdi) as GecerliUrunAdi, 
                        IFNULL(u.Kategori, s.Kategori) as GecerliKategori, 
                        IFNULL(u.Marka, s.Marka) as GecerliMarka 
                    FROM Siparisler s 
                    LEFT JOIN Urunler u ON s.Barkod = u.Barkod 
                    WHERE s.rowid = @id";

                using (var cmd = new SqliteCommand(sqlPopup, b))
                {
                    cmd.Parameters.AddWithValue("@id", rowId);
                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read())
                        {
                            string sipNo = r["SiparisNo"].ToString();
                            string platform = r["PlatformAd"].ToString();
                            string barkod = r["Barkod"]?.ToString() ?? "-";
                            string urunAdi = r["GecerliUrunAdi"]?.ToString() ?? "Bilinmeyen Ürün";
                            string kat = r["GecerliKategori"]?.ToString() ?? "-";
                            string marka = r["GecerliMarka"]?.ToString() ?? "-";

                            int adet = Convert.ToInt32(r["Adet"]);
                            if (adet == 0) adet = 1;

                            double alis = Convert.ToDouble(r["AlisFiyati"] == DBNull.Value ? 0 : r["AlisFiyati"]);
                            double satis = Convert.ToDouble(r["SatisFiyati"] == DBNull.Value ? 0 : r["SatisFiyati"]);
                            double indirim = Convert.ToDouble(r["PlatformIndirimi"] == DBNull.Value ? 0 : r["PlatformIndirimi"]);
                            double stopaj = Convert.ToDouble(r["Stopaj"] == DBNull.Value ? 0 : r["Stopaj"]);
                            double pos = Convert.ToDouble(r["PosBedeli"] == DBNull.Value ? 0 : r["PosBedeli"]);
                            double hizmet = Convert.ToDouble(r["HizmetBedeli"] == DBNull.Value ? 0 : r["HizmetBedeli"]);
                            double komisyon = Convert.ToDouble(r["KomisyonTL"] == DBNull.Value ? 0 : r["KomisyonTL"]);
                            double kargo = Convert.ToDouble(r["KargoUcreti"] == DBNull.Value ? 0 : r["KargoUcreti"]);
                            double cezaTutari = Convert.ToDouble(r["CezaTutari"] == DBNull.Value ? 0 : r["CezaTutari"]);
                            string cezaSebebi = r["CezaSebebi"]?.ToString() ?? "";
                            string imei = r["SatilanImei"].ToString();
                            string durum = r["Durum"].ToString().ToUpper();
                            string tarih = r["Tarih"].ToString();

                            // 🚀 MÜHÜR: DETAY POPUP FATURASINDA KÂRLILIK DÜZELTİLDİ
                            double kesintilerToplami = kargo + komisyon + hizmet + pos + stopaj + cezaTutari;
                            double netKar = durum == "İADE" ? -kesintilerToplami : ((satis * adet) + indirim - (alis * adet) - kesintilerToplami);

                            Form frmDetay = new Form { Text = $"📄 Detaylı Sipariş Analiz Raporu - No: {sipNo}", Size = new Size(720, 520), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false, MaximizeBox = false, MinimizeBox = false, Font = new Font("Segoe UI", 10F) };

                            GroupBox grpGenel = new GroupBox { Text = "📦 Genel ve Lojistik Bilgiler", Location = new Point(20, 15), Width = 330, Height = 360, ForeColor = Color.FromArgb(56, 189, 248) };
                            grpGenel.Controls.Add(new Label { Text = $"Sipariş No:\n{sipNo}", Location = new Point(15, 25), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White });
                            grpGenel.Controls.Add(new Label { Text = $"Satış Kanalı: {platform}", Location = new Point(15, 75), AutoSize = true, ForeColor = Color.LightGray });
                            grpGenel.Controls.Add(new Label { Text = $"Sipariş Tarihi: {tarih}", Location = new Point(15, 100), AutoSize = true, ForeColor = Color.LightGray });

                            Label lblDurum = new Label { Text = $"SİSTEM DURUMU: {durum}", Location = new Point(15, 135), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = (durum == "İPTAL" || durum == "İADE") ? Color.Crimson : ((durum == "TAHSİL EDİLDİ" || durum == "ÖDENDİ") ? Color.LimeGreen : (durum == "İHTİLAFLI" ? Color.MediumPurple : Color.Gold)) };
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
                                grpFinans.Controls.Add(lblB); grpFinans.Controls.Add(lblD); y += 28;
                            };

                            if (durum == "İADE") { FinansSatiriEkle("Brüt Satış (İADE EDİLDİ):", 0, Color.White); FinansSatiriEkle("Platform Desteği (İADE EDİLDİ):", 0, Color.LimeGreen); }
                            else
                            {
                                FinansSatiriEkle($"Satış Toplamı:", satis * adet, Color.White);
                                FinansSatiriEkle("Platform Desteği (Gelir) (+):", indirim, Color.LimeGreen);
                                FinansSatiriEkle($"Alış Maliyet Toplamı (-):", alis * adet, Color.LightSkyBlue);
                            }

                            grpFinans.Controls.Add(new Label { Text = "────────────────────────", Location = new Point(15, y - 5), AutoSize = true, ForeColor = Color.FromArgb(50, 50, 50) }); y += 15;
                            FinansSatiriEkle("Pazaryeri Komisyonu:", komisyon, Color.FromArgb(239, 68, 68));
                            FinansSatiriEkle("Lojistik Kargo Ücreti:", kargo, Color.FromArgb(239, 68, 68));
                            FinansSatiriEkle("Güvence / POS Kesintisi:", pos, Color.FromArgb(239, 68, 68));
                            FinansSatiriEkle("Sabit Hizmet Bedeli:", hizmet, Color.FromArgb(239, 68, 68));
                            FinansSatiriEkle("Devlet Stopaj Vergisi:", stopaj, Color.FromArgb(239, 68, 68));

                            if (cezaTutari > 0)
                            {
                                FinansSatiriEkle($"Pazaryeri Cezası ({cezaSebebi}):", cezaTutari, Color.Crimson);
                            }
                            frmDetay.Controls.Add(grpFinans);

                            Panel pnlKarKutusu = new Panel { Location = new Point(20, 390), Width = 665, Height = 65, BackColor = Color.FromArgb(26, 26, 39), BorderStyle = BorderStyle.FixedSingle };
                            Label lblKarYazi = new Label { Text = (durum == "İADE" || netKar < 0) ? "BU SİPARİŞ SATIRININ TOPLAM NET ZARARI:" : "BU SİPARİŞ SATIRININ NET KÂR / ZARAR DURUMU:", Location = new Point(15, 22), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.DarkGray };
                            Label lblKarDeger = new Label { Text = netKar.ToString("N2") + " TL", Location = new Point(380, 12), Width = 260, Height = 40, TextAlign = ContentAlignment.MiddleRight, Font = new Font("Segoe UI Black", 16F, FontStyle.Bold), ForeColor = netKar < 0 ? Color.Crimson : Color.FromArgb(34, 197, 94) };

                            pnlKarKutusu.Controls.Add(lblKarYazi); pnlKarKutusu.Controls.Add(lblKarDeger);
                            frmDetay.Controls.Add(pnlKarKutusu);

                            frmDetay.ShowDialog();
                        }
                    }
                }
            }
        }
        private ComboBox ComboYarat(Panel p, string labelText, Point loc, int width)
        {
            Label lbl = new Label { Text = labelText, Location = new Point(loc.X, loc.Y), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            ComboBox cmb = new ComboBox { Location = new Point(loc.X, loc.Y + 18), Width = width, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F) };
            p.Controls.Add(lbl); p.Controls.Add(cmb);
            return cmb;
        }

        private Label WidgetYarat(TableLayoutPanel tlp, string baslik, Color solRenk, int sutun)
        {
            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Margin = new Padding(8) };
            Panel pLine = new Panel { Dock = DockStyle.Left, Width = 5, BackColor = solRenk };
            p.Controls.Add(pLine);
            Label lblB = new Label { Text = baslik, Dock = DockStyle.Top, Height = 22, Padding = new Padding(8, 5, 0, 0), Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.DarkGray };
            Label lblV = new Label { Text = "0", Dock = DockStyle.Fill, Padding = new Padding(8, 5, 0, 0), Font = new Font("Segoe UI Black", 14F, FontStyle.Bold), ForeColor = Color.White };
            p.Controls.Add(lblV); p.Controls.Add(lblB);
            tlp.Controls.Add(p, sutun, 0);
            return lblV;
        }
    }
}