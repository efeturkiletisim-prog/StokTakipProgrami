using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    // 👑 ŞİRKET GENEL PERFORMANS VE MAKRO ENGELLENEMEZ BÜYÜME ANALİZİ (VİP KART TASARIMI VE AKILLI RENK RADARI ENTEGRE EDİLDİ)
    public partial class SayfaGenelRapor : UserControl
    {
        private DataGridView dynDgvTopUrunler, dynDgvPlatDagilim;
        private Label dynLblGenelCiro, dynLblGenelKar, dynLblToplamSiparis, dynLblIadeAdeti;
        private Label dynLblOdenenAdeti, dynLblBekleyenAdeti, dynLblBekleyenAlacak;

        // 🚀 MÜHÜR: Yeni Platform Özetleri Konteyneri
        private FlowLayoutPanel pnlPlatformOzetleri;

        public SayfaGenelRapor()
        {
            InitializeComponent();
            this.Load += SayfaGenelRapor_Load;
        }

        private void SayfaGenelRapor_Load(object sender, EventArgs e)
        {
            ArayuzInsaEt();
            GenelMetrikleriHesapla();
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "👑 ŞİRKET GENEL PERFORMANS VE MAKRO ENGELLENEMEZ BÜYÜME ANALİZİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            TableLayoutPanel pnlKPI = new TableLayoutPanel { Dock = DockStyle.Top, Height = 100, BackColor = Color.Transparent, ColumnCount = 7, RowCount = 1, Padding = new Padding(10, 15, 10, 10) };
            pnlKPI.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28F));
            pnlKPI.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28F));
            pnlKPI.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28F));
            pnlKPI.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28F));
            pnlKPI.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28F));
            pnlKPI.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28F));
            pnlKPI.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28F));
            this.Controls.Add(pnlKPI);

            // Kurumsal Tasarım Kartları (Sola Dayalı Yeni Mimari)
            dynLblToplamSiparis = WidgetYarat(pnlKPI, "🛍️ ÖMÜRLÜK SATIŞ", Color.FromArgb(56, 189, 248), 0);
            dynLblOdenenAdeti = WidgetYarat(pnlKPI, "✔ ÖDENEN ADET", Color.FromArgb(34, 197, 94), 1);
            dynLblBekleyenAdeti = WidgetYarat(pnlKPI, "⏳ BEKLEYEN ADET", Color.Orange, 2);
            dynLblIadeAdeti = WidgetYarat(pnlKPI, "🔄 İADE ADETİ", Color.Crimson, 3);
            dynLblGenelCiro = WidgetYarat(pnlKPI, "💎 BRÜT CİRO", Color.FromArgb(0, 120, 215), 4);
            dynLblGenelKar = WidgetYarat(pnlKPI, "💰 NET SAF KÂR", Color.FromArgb(34, 197, 94), 5);
            dynLblBekleyenAlacak = WidgetYarat(pnlKPI, "💸 NET BEKLEYEN ALACAK", Color.FromArgb(241, 196, 15), 6);

            // 🚀 MÜHÜR: Yeni VİP Platform Finans Paneli (Yükseklik 220px yapılarak ezilme sorunu çözüldü)
            Panel pnlPlatformWrapper = new Panel { Dock = DockStyle.Top, Height = 220, Padding = new Padding(10, 5, 10, 10), BackColor = Color.Transparent };
            Label lblPlatTitle = new Label { Text = "🌐 PLATFORM BAZLI FİNANSAL RÖNTGEN VE TAHSİLAT RADARI", Dock = DockStyle.Top, Height = 30, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.Gold, Padding = new Padding(0, 5, 0, 0) };
            pnlPlatformOzetleri = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, WrapContents = false, Padding = new Padding(0, 10, 0, 0) };
            pnlPlatformWrapper.Controls.Add(pnlPlatformOzetleri);
            pnlPlatformWrapper.Controls.Add(lblPlatTitle);
            this.Controls.Add(pnlPlatformWrapper);

            TableLayoutPanel mainGridTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(10) };
            mainGridTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));
            mainGridTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55F));
            this.Controls.Add(mainGridTable);

            GroupBox grpSol = new GroupBox { Dock = DockStyle.Fill, Text = "🔥 Mağazanın En Çok Satan Şampiyon 5 Ürünü", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(12, 25, 12, 12), Margin = new Padding(8) };
            dynDgvTopUrunler = GridYarat();
            dynDgvTopUrunler.DataBindingComplete += Grid_DataBindingComplete;
            grpSol.Controls.Add(dynDgvTopUrunler);
            mainGridTable.Controls.Add(grpSol, 0, 0);

            GroupBox grpSag = new GroupBox { Dock = DockStyle.Fill, Text = "🌐 Satış Kanalları (Platform) Dağılım Payı", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248), BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(12, 25, 12, 12), Margin = new Padding(8) };
            dynDgvPlatDagilim = GridYarat();
            dynDgvPlatDagilim.DataBindingComplete += Grid_DataBindingComplete;
            grpSag.Controls.Add(dynDgvPlatDagilim);
            mainGridTable.Controls.Add(grpSag, 1, 0);

            // Z-Order Dizilimi
            mainGridTable.BringToFront();
            pnlPlatformWrapper.SendToBack();
            pnlKPI.SendToBack();
            pnlHeader.SendToBack();
        }

        private void GenelMetrikleriHesapla()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();

                    string sqlMetrik = @"
                        SELECT 
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)!='İPTAL' THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END), 0) as ToplamSiparis,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END), 0) as iadeAdeti,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' OR UPPER(s.Durum)='İPTAL' THEN 0 ELSE (IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END), 0) as ciro,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' OR UPPER(s.Durum)='İPTAL' THEN 0 ELSE IFNULL(s.PlatformIndirimi, 0) END), 0) as platformIndirimi,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum)='İADE' OR UPPER(s.Durum)='İPTAL' THEN 0 ELSE (IFNULL(s.AlisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END), 0) as maliyet,
                            IFNULL(SUM(IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0)), 0) as kesintiler,
                            IFNULL(SUM(CASE WHEN UPPER(s.Durum) IN ('ÖDENDİ', 'TAHSİL EDİLDİ', 'TAMAMLANDI (TESLİM)') THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END), 0) as odenenAdet,
                            IFNULL(SUM(CASE 
                                WHEN UPPER(s.Durum) NOT IN ('ÖDENDİ', 'TAHSİL EDİLDİ', 'TAMAMLANDI (TESLİM)', 'İADE', 'İPTAL') 
                                THEN ((IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) + IFNULL(s.PlatformIndirimi, 0)) - (IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0))
                                ELSE 0 
                            END), 0) as bekleyenNetAlacak
                        FROM Siparisler s
                        WHERE UPPER(s.Durum) != 'İPTAL'";

                    using (var cmdM = new SqliteCommand(sqlMetrik, b))
                    using (var oku = cmdM.ExecuteReader())
                    {
                        if (oku.Read())
                        {
                            int sAdet = Convert.ToInt32(oku["ToplamSiparis"]);
                            int iAdet = Convert.ToInt32(oku["iadeAdeti"]);
                            int oAdet = Convert.ToInt32(oku["odenenAdet"]);

                            int bAdet = sAdet - oAdet - iAdet;

                            double ciro = Convert.ToDouble(oku["ciro"]);
                            double platformIndirimi = Convert.ToDouble(oku["platformIndirimi"]);
                            double kesintiler = Convert.ToDouble(oku["kesintiler"]);
                            double maliyet = Convert.ToDouble(oku["maliyet"]);

                            // Net Kâr: Ciroya platform desteği eklenir, fatura kesintileri ve ürün maliyeti düşülür.
                            double netKar = (ciro + platformIndirimi - kesintiler) - maliyet;

                            double bekleyenAlacakNet = Convert.ToDouble(oku["bekleyenNetAlacak"]);

                            dynLblToplamSiparis.Text = sAdet.ToString("N0") + " Adet";
                            dynLblOdenenAdeti.Text = oAdet.ToString("N0") + " Adet";
                            dynLblBekleyenAdeti.Text = bAdet.ToString("N0") + " Adet";
                            dynLblIadeAdeti.Text = iAdet.ToString("N0") + " Adet";

                            dynLblGenelCiro.Text = ciro.ToString("N2") + " TL";
                            dynLblGenelKar.Text = netKar.ToString("N2") + " TL";
                            dynLblGenelKar.ForeColor = netKar < 0 ? Color.Crimson : Color.FromArgb(34, 197, 94);

                            dynLblBekleyenAlacak.Text = bekleyenAlacakNet.ToString("N2") + " TL";
                            dynLblBekleyenAlacak.ForeColor = bekleyenAlacakNet < 0 ? Color.Crimson : Color.FromArgb(241, 196, 15);
                        }
                    }

                    pnlPlatformOzetleri.Controls.Clear();
                    string sqlKartlar = @"
                        SELECT 
                            IFNULL(s.PlatformAd, 'DİĞER') as Platform,
                            SUM(CASE WHEN UPPER(s.Durum) NOT IN ('İPTAL', 'İADE') THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END) as TopSiparis,
                            SUM(CASE WHEN UPPER(s.Durum) IN ('ÖDENDİ', 'TAHSİL EDİLDİ', 'TAMAMLANDI (TESLİM)') THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END) as OdenenAdet,
                            SUM(CASE WHEN UPPER(s.Durum) NOT IN ('İPTAL', 'İADE') THEN (IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) ELSE 0 END) as Ciro,
                            SUM(CASE WHEN UPPER(s.Durum) IN ('ÖDENDİ', 'TAHSİL EDİLDİ', 'TAMAMLANDI (TESLİM)') 
                                     THEN ((IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) + IFNULL(s.PlatformIndirimi, 0)) - (IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0)) ELSE 0 END) as Gelen,
                            SUM(CASE WHEN UPPER(s.Durum) NOT IN ('ÖDENDİ', 'TAHSİL EDİLDİ', 'TAMAMLANDI (TESLİM)', 'İPTAL', 'İADE') 
                                     THEN ((IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) + IFNULL(s.PlatformIndirimi, 0)) - (IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0)) ELSE 0 END) as Kalan,
                            
                            -- Net Kar hesaplayabilmek için ek sütunlar çekiliyor
                            SUM(CASE WHEN UPPER(s.Durum) NOT IN ('İPTAL', 'İADE') THEN IFNULL(s.PlatformIndirimi, 0) ELSE 0 END) as TopIndirim,
                            SUM(CASE WHEN UPPER(s.Durum) NOT IN ('İPTAL', 'İADE') THEN (IFNULL(s.AlisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) ELSE 0 END) as TopMaliyet,
                            SUM(CASE WHEN UPPER(s.Durum) NOT IN ('İPTAL', 'İADE') THEN IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0) ELSE 0 END) as TopKesinti

                        FROM Siparisler s
                        WHERE s.PlatformAd IS NOT NULL AND s.PlatformAd != ''
                        GROUP BY s.PlatformAd
                        ORDER BY Ciro DESC";

                    using (var cmdKart = new SqliteCommand(sqlKartlar, b))
                    using (var rK = cmdKart.ExecuteReader())
                    {
                        while (rK.Read())
                        {
                            string plat = rK["Platform"].ToString();
                            int tSip = Convert.ToInt32(rK["TopSiparis"]);
                            int oSip = Convert.ToInt32(rK["OdenenAdet"]);
                            int bSip = tSip - oSip;
                            double ciro = Convert.ToDouble(rK["Ciro"]);
                            double gelen = Convert.ToDouble(rK["Gelen"]);
                            double kalan = Convert.ToDouble(rK["Kalan"]);

                            double indirim = Convert.ToDouble(rK["TopIndirim"]);
                            double maliyet = Convert.ToDouble(rK["TopMaliyet"]);
                            double kesinti = Convert.ToDouble(rK["TopKesinti"]);

                            double netKar = (ciro + indirim - kesinti) - maliyet;

                            PlatformKartiOlustur(plat, tSip, bSip, ciro, gelen, kalan, netKar); // 🚀 netKar parametresi eklendi
                        }
                    }
                    string sqlTopUrun = @"
                        SELECT IFNULL(u.UrunAdi, s.UrunAdi) as 'Şampiyon Ürün Modeli', SUM(CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) as 'Satılan Adet'
                        FROM Siparisler s
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod
                        WHERE UPPER(s.Durum) != 'İPTAL' AND UPPER(s.Durum) != 'İADE'
                        GROUP BY s.Barkod, s.UrunAdi
                        ORDER BY SUM(CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) DESC LIMIT 5";

                    using (var cmdTop = new SqliteCommand(sqlTopUrun, b))
                    {
                        DataTable dtTop = new DataTable(); dtTop.Load(cmdTop.ExecuteReader());
                        dynDgvTopUrunler.DataSource = dtTop;
                    }

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
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) END) as SatisAdeti,
                            SUM(CASE WHEN UPPER(s.Durum) IN ('ÖDENDİ', 'TAHSİL EDİLDİ', 'TAMAMLANDI (TESLİM)') THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END) as OdenenAdet,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END) ELSE 0 END) as IadeAdeti,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.SatisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END) as ToplamCiro,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE IFNULL(s.PlatformIndirimi, 0) END) as ToplamIndirim,
                            SUM(IFNULL(s.Stopaj, 0) + IFNULL(s.PosBedeli, 0) + IFNULL(s.HizmetBedeli, 0) + IFNULL(s.KomisyonTL, 0) + IFNULL(s.KargoUcreti, 0) + IFNULL(s.CezaTutari, 0)) as ToplamKesinti,
                            SUM(CASE WHEN UPPER(s.Durum)='İADE' THEN 0 ELSE (IFNULL(s.AlisFiyati, 0) * (CASE WHEN IFNULL(s.Adet,0)=0 THEN 1 ELSE s.Adet END)) END) as ToplamMaliyet
                        FROM Siparisler s
                        WHERE UPPER(s.Durum) != 'İPTAL'
                        GROUP BY s.PlatformAd
                        ORDER BY ToplamCiro DESC";

                    using (var cmdPlat = new SqliteCommand(sqlPlat, b))
                    {
                        using (var rP = cmdPlat.ExecuteReader())
                        {
                            while (rP.Read())
                            {
                                int sAdeti = Convert.ToInt32(rP["SatisAdeti"]);
                                int oAdet = Convert.ToInt32(rP["OdenenAdet"]);
                                int iAdet = Convert.ToInt32(rP["IadeAdeti"]);

                                double gCiro = Convert.ToDouble(rP["ToplamCiro"]);
                                double gIndirim = Convert.ToDouble(rP["ToplamIndirim"]);
                                double gKesinti = Convert.ToDouble(rP["ToplamKesinti"]);
                                double gMaliyet = Convert.ToDouble(rP["ToplamMaliyet"]);

                                // Sağ grid dağılım payı kâr hesaplama kalibrasyonu
                                double gNetKar = (gCiro + gIndirim - gKesinti) - gMaliyet;

                                double gTransfer = gCiro + gIndirim - gKesinti;
                                dtPlat.Rows.Add(rP["PlatformAd"].ToString(), sAdeti, oAdet, (sAdeti - oAdet - iAdet), iAdet, gCiro, gTransfer, gMaliyet, gTransfer - gMaliyet);
                            }
                        }
                    }

                    dynDgvPlatDagilim.DataSource = dtPlat;

                    foreach (DataGridViewColumn col in dynDgvTopUrunler.Columns) if (col.ValueType == typeof(double)) { col.DefaultCellStyle.Format = "N2"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; }
                    foreach (DataGridViewColumn col in dynDgvPlatDagilim.Columns) if (col.ValueType == typeof(double)) { col.DefaultCellStyle.Format = "N2"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; }
                }
            }
            catch { }
        }

        private Color GetPlatformRenk(string plat)
        {
            string p = plat.ToUpper();
            if (p.Contains("TRENDYOL")) return Color.FromArgb(242, 122, 26);
            if (p.Contains("HEPSİBURADA") || p.Contains("HEPSIBURADA")) return Color.FromArgb(255, 96, 0);
            if (p.Contains("N11")) return Color.FromArgb(92, 62, 188);
            if (p.Contains("ÇİÇEK") || p.Contains("CICEK")) return Color.FromArgb(9, 115, 185);
            if (p.Contains("AMAZON")) return Color.FromArgb(255, 153, 0);
            if (p.Contains("TİKTOK") || p.Contains("TIKTOK")) return Color.FromArgb(255, 0, 80);
            if (p.Contains("DOLAP")) return Color.FromArgb(56, 189, 248);
            return Color.MediumSlateBlue;
        }

        private void PlatformKartiOlustur(string platform, int tSip, int bSip, double ciro, double gelen, double kalan, double netKar)
        {
            Color pColor = GetPlatformRenk(platform);
            Panel pnlOut = new Panel { Width = 350, Height = 140, Margin = new Padding(0, 0, 15, 0), BackColor = pColor, Padding = new Padding(1, 4, 1, 1) };
            Panel pnlIn = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28) };

            Label lblTitle = new Label { Text = platform.ToUpper(), Location = new Point(12, 10), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = pColor };
            pnlIn.Controls.Add(lblTitle);

            Panel sep = new Panel { Location = new Point(15, 33), Width = 320, Height = 1, BackColor = Color.FromArgb(45, 45, 48) };
            pnlIn.Controls.Add(sep);

            Label lblSip = new Label { Text = $"📦 Satılan: {tSip} Adet", Location = new Point(12, 45), AutoSize = true, Font = new Font("Segoe UI", 8.5F), ForeColor = Color.LightGray };
            Label lblBek = new Label { Text = $"⏳ Bekleyen: {bSip} Adet", Location = new Point(12, 67), AutoSize = true, Font = new Font("Segoe UI", 8.5F), ForeColor = Color.Orange };
            pnlIn.Controls.Add(lblSip); pnlIn.Controls.Add(lblBek);

            Label lblCiroT = new Label { Text = "Toplam Ciro:", Location = new Point(155, 42), AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.Gray };
            Label lblCiro = new Label { Text = $"{ciro:N2} ₺", Location = new Point(155, 58), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White };
            pnlIn.Controls.Add(lblCiroT); pnlIn.Controls.Add(lblCiro);

            // 🚀 MÜHÜR: Net Kâr sağ sütuna, Toplam Ciro ile yan yana konumlandırıldı
            Label lblKarT = new Label { Text = "Net Kâr:", Location = new Point(255, 42), AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.Gray };
            Label lblKar = new Label { Text = $"{netKar:N2} ₺", Location = new Point(255, 58), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = netKar < 0 ? Color.Crimson : Color.FromArgb(34, 197, 94) };
            pnlIn.Controls.Add(lblKarT); pnlIn.Controls.Add(lblKar);

            Label lblGelenT = new Label { Text = "Gelen Tutar:", Location = new Point(155, 88), AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.Gray };
            Label lblGelen = new Label { Text = $"{gelen:N2} ₺", Location = new Point(155, 104), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.LimeGreen };
            pnlIn.Controls.Add(lblGelenT); pnlIn.Controls.Add(lblGelen);

            Label lblKalanT = new Label { Text = "Kalan Alacak:", Location = new Point(255, 88), AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.Gray };
            Label lblKalan = new Label { Text = $"{kalan:N2} ₺", Location = new Point(255, 104), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) };
            pnlIn.Controls.Add(lblKalanT); pnlIn.Controls.Add(lblKalan);

            double totalHakedis = gelen + kalan;
            int pct = totalHakedis > 0 ? (int)((gelen / totalHakedis) * 100) : 0;
            if (pct > 100) pct = 100; if (pct < 0) pct = 0;

            Label lblPct = new Label { Text = $"Tahsilat: %{pct}", Location = new Point(12, 95), AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.Gray };
            pnlIn.Controls.Add(lblPct);

            Panel pnlBg = new Panel { Location = new Point(15, 115), Width = 120, Height = 5, BackColor = Color.FromArgb(45, 45, 48) };
            Panel pnlFg = new Panel { Location = new Point(0, 0), Width = (int)(120 * (pct / 100.0)), Height = 5, BackColor = Color.LimeGreen };
            pnlBg.Controls.Add(pnlFg);
            pnlIn.Controls.Add(pnlBg);

            pnlOut.Controls.Add(pnlIn);
            pnlPlatformOzetleri.Controls.Add(pnlOut);
        }
        private DataGridView GridYarat()
        {
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(26, 26, 28),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 35
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
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

        private Label WidgetYarat(TableLayoutPanel tlp, string baslik, Color temaRengi, int sutun)
        {
            Panel pOut = new Panel { Dock = DockStyle.Fill, Margin = new Padding(4), BackColor = Color.FromArgb(50, 50, 55), Padding = new Padding(1) };
            Panel pIn = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28) };
            Panel pTopLine = new Panel { Dock = DockStyle.Top, Height = 4, BackColor = temaRengi };

            Label lblBaslik = new Label
            {
                Text = baslik.ToUpper(),
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 5, 0, 0),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
                ForeColor = Color.DarkGray
            };

            Label lblDeger = new Label
            {
                Text = "0",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 10),
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                ForeColor = temaRengi
            };

            pIn.Controls.Add(lblDeger);
            pIn.Controls.Add(lblBaslik);
            pIn.Controls.Add(pTopLine);
            pOut.Controls.Add(pIn);

            tlp.Controls.Add(pOut, sutun, 0);
            return lblDeger;
        }
    }
}