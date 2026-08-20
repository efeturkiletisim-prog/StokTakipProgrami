using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using Microsoft.Data.Sqlite;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;

namespace StokTakipProgrami
{
    public partial class SayfaAnaEkran : UserControl
    {
        private Panel btnKartSiparis, btnKartStok, btnKartKritik, btnKartSorgula;

        public SayfaAnaEkran()
        {
            InitializeComponent();
            ArayuzInsaEt();
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(33, 33, 36);

            Label lblWelcome = new Label { Text = $"👋 HOŞ GELDİNİZ, {Veritabanı.AktifKullanici} | EFETURK.NET ERİŞİM PANELİ", Location = new Point(30, 25), AutoSize = true, Font = new Font("Segoe UI", 13F, FontStyle.Bold), ForeColor = Color.White };
            this.Controls.Add(lblWelcome);

            Label lblSub = new Label { Text = "Hızlı operasyon yönetimi için aşağıdaki 4 ana komuta butonunu kullanabilirsiniz:", Location = new Point(32, 58), AutoSize = true, Font = new Font("Segoe UI", 9.5F), ForeColor = Color.DarkGray };
            this.Controls.Add(lblSub);

            btnKartSiparis = HızlıButonYarat("🛍️  YENİ SİPARİŞ GİRİŞİ", "Gelen yeni siparişlerin sevkiyat kaydı.", new Point(30, 110), Color.FromArgb(46, 117, 89));
            btnKartStok = HızlıButonYarat("➕  YENİ ÜRÜN EKLE", "Envantere yeni ürün ekleme.", new Point(410, 110), Color.FromArgb(0, 120, 215));
            btnKartKritik = HızlıButonYarat("📋  STOK SORGULA", "Depoda mevcut olan ürünler.", new Point(30, 250), Color.FromArgb(230, 126, 34));
            btnKartSorgula = HızlıButonYarat("🔍  SİPARİŞ SORGULAMA", "Müşteri adı veya kargo kodu ile hızlı arama.", new Point(410, 250), Color.FromArgb(0, 158, 247));

            this.Controls.Add(btnKartSiparis); this.Controls.Add(btnKartStok); this.Controls.Add(btnKartKritik); this.Controls.Add(btnKartSorgula);

            Panel pnlOzets = new Panel { Location = new Point(30, 380), Size = new Size(730, 90), BackColor = Color.FromArgb(21, 21, 33) };
            Label lblGun = new Label { Text = "BUGÜN: ", Location = new Point(10, 10), AutoSize = true, Font = new Font("Segoe UI", 11F, FontStyle.Bold), ForeColor = Color.White };
            Label lblAdet = new Label { Name = "lblBugunAdet", Text = "Sipariş Adedi: 0", Location = new Point(10, 40), AutoSize = true, Font = new Font("Segoe UI", 10F), ForeColor = Color.LightGray };
            Label lblCiro = new Label { Name = "lblBugunCiro", Text = "Ciro: 0,00 TL", Location = new Point(260, 40), AutoSize = true, Font = new Font("Segoe UI", 10F), ForeColor = Color.LightGray };
            Label lblKar = new Label { Name = "lblBugunKar", Text = "Kâr: 0,00 TL", Location = new Point(480, 40), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.FromArgb(74, 222, 128) };
            pnlOzets.Controls.Add(lblGun); pnlOzets.Controls.Add(lblAdet); pnlOzets.Controls.Add(lblCiro); pnlOzets.Controls.Add(lblKar);
            this.Controls.Add(pnlOzets);

            BugununIstatistikleriniGoster();
        }

        private Panel HızlıButonYarat(string baslik, string aciklama, Point konum, Color solSeritRengi)
        {
            Panel p = new Panel { Location = konum, Size = new Size(350, 110), BackColor = Color.FromArgb(240, 242, 245), Cursor = Cursors.Hand };
            Panel pLine = new Panel { Dock = DockStyle.Left, Width = 6, BackColor = solSeritRengi }; p.Controls.Add(pLine);

            Label lblB = new Label { Text = baslik, Location = new Point(18, 18), AutoSize = true, Font = new Font("Segoe UI Black", 12F, FontStyle.Bold), ForeColor = Color.Black, Enabled = false };
            Label lblA = new Label { Text = aciklama, Location = new Point(18, 48), Size = new Size(310, 45), Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold), ForeColor = Color.FromArgb(40, 40, 40), Enabled = false };

            p.Controls.Add(lblB); p.Controls.Add(lblA);
            p.MouseEnter += (s, e) => { p.BackColor = Color.White; };
            p.MouseLeave += (s, e) => { p.BackColor = Color.FromArgb(240, 242, 245); };

            p.Click += (s, e) => {
                if (baslik.Contains("SİPARİŞ GİRİŞİ")) Form1.AnaForm?.PanelSayfaAc(new SayfaSiparisEkle());
                else if (baslik.Contains("ÜRÜN EKLE")) Form1.AnaForm?.PanelSayfaAc(new SayfaStoklar());
                else if (baslik.Contains("STOK SORGULA")) Form1.AnaForm?.PanelSayfaAc(new SayfaStokSorgula());
                else if (baslik.Contains("SİPARİŞ SORGULAMA")) Form1.AnaForm?.PanelSayfaAc(new SayfaSiparisSorgula());
            };
            return p;
        }

        private void BugununIstatistikleriniGoster()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string bugun = DateTime.Now.ToString("dd.MM.yyyy");

                    string stopaj = "0", pos = "0", hizmet = "0", kom = "0", kargo = "0", ceza = "0";
                    var sutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var pragma = new SqliteCommand("PRAGMA table_info('Siparisler');", b))
                    using (var reader = pragma.ExecuteReader()) { while (reader.Read()) sutunlar.Add(reader["name"].ToString()); }

                    if (sutunlar.Contains("Stopaj")) stopaj = "s.Stopaj";
                    if (sutunlar.Contains("PosBedeli")) pos = "s.PosBedeli";
                    if (sutunlar.Contains("HizmetBedeli")) hizmet = "s.HizmetBedeli";
                    if (sutunlar.Contains("KomisyonTL")) kom = "s.KomisyonTL";
                    if (sutunlar.Contains("KargoUcreti")) kargo = "s.KargoUcreti";
                    if (sutunlar.Contains("CezaTutari")) ceza = "s.CezaTutari";

                    string sql = $@"
                        SELECT 
                            COUNT(DISTINCT s.SiparisNo) as Adet, 
                            SUM(IFNULL(s.SatisFiyati, 0) * s.Adet) as Ciro, 
                            SUM( ((IFNULL(s.SatisFiyati, 0) * s.Adet) - (IFNULL(st.AlisFiyati, 0) * s.Adet)) - 
                                 (IFNULL({stopaj}, 0) + IFNULL({pos}, 0) + IFNULL({hizmet}, 0) + IFNULL({kom}, 0) + IFNULL({kargo}, 0) + IFNULL({ceza}, 0)) ) as Kar 
                        FROM Siparisler s 
                        LEFT JOIN Stoklar st ON s.Barkod = st.Barkod 
                        WHERE s.Tarih = @t AND s.Durum != 'İPTAL'";

                    using (var cmd = new SqliteCommand(sql, b))
                    {
                        cmd.Parameters.AddWithValue("@t", bugun);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                int adet = r["Adet"] == DBNull.Value ? 0 : Convert.ToInt32(r["Adet"]);
                                double ciro = r["Ciro"] == DBNull.Value ? 0.0 : Convert.ToDouble(r["Ciro"]);
                                double kar = r["Kar"] == DBNull.Value ? 0.0 : Convert.ToDouble(r["Kar"]);

                                Control[] cAdet = this.Controls.Find("lblBugunAdet", true); Control[] cCiro = this.Controls.Find("lblBugunCiro", true); Control[] cKar = this.Controls.Find("lblBugunKar", true);
                                if (cAdet.Length > 0) ((Label)cAdet[0]).Text = $"Sipariş Adedi: {adet} Adet";
                                // 🚀 MÜHÜR: Kaçak etiket tamamen temizlendi!
                                if (cCiro.Length > 0) ((Label)cCiro[0]).Text = $"Ciro: {ciro.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))} TL";
                                if (cKar.Length > 0)
                                {
                                    Label lK = (Label)cKar[0];
                                    lK.Text = $"Kâr: {kar.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))} TL";
                                    lK.ForeColor = kar < 0 ? Color.FromArgb(239, 68, 68) : Color.FromArgb(74, 222, 128);
                                }
                            }
                        }
                    }
                }
            }
            catch { }
        }

        // 🚀 MÜHÜR: Designer dosyasının patlamasını önleyen Paint metodu geri getirildi!
        public void panel1_Paint(object sender, PaintEventArgs e) { }
    }
}