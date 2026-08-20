using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public partial class SayfaKarZarar : UserControl
    {
        private DataGridView dynDgvRapor;
        private Label dynLblToplamCiro, dynLblToplamMaliyet, dynLblToplamKomisyon, dynLblNetKar;
        private Button dynBtnExcel;

        public SayfaKarZarar()
        {
            InitializeComponent();
            this.Load += SayfaKarZarar_Load;
        }

        private void SayfaKarZarar_Load(object sender, EventArgs e)
        {
            VeritabaniniOtoTamirEt(); // 🚀 SQL Hatalarını Önleyen Zeki Motor
            ArayuzInsaEt();
            NetFinansalAnalizYap();
        }

        // --- SQL SÜTUN HATALARINI OTOMATİK DÜZELTEN MOTOR ---
        private void VeritabaniniOtoTamirEt()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    // Siparisler tablosunda eksik olan sütunları otomatik onarır
                    string[] siparisSutunlari = { "MusteriAdSoyad TEXT", "Platform TEXT", "SatisFiyati REAL DEFAULT 0", "Barkod TEXT" };
                    foreach (var s in siparisSutunlari)
                    {
                        try { using (var cmd = new SqliteCommand($"ALTER TABLE Siparisler ADD COLUMN {s}", b)) cmd.ExecuteNonQuery(); } catch { }
                    }
                    // Stoklar tablosunda eksik olan sütunları otomatik onarır
                    string[] stokSutunlari = { "Barkod TEXT", "AlisFiyati REAL DEFAULT 0" };
                    foreach (var s in stokSutunlari)
                    {
                        try { using (var cmd = new SqliteCommand($"ALTER TABLE Stoklar ADD COLUMN {s}", b)) cmd.ExecuteNonQuery(); } catch { }
                    }
                }
            }
            catch { }
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "📈  KÂR - ZARAR VE KOMİSYON ANALİZİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            Panel pnlSkorlar = new Panel { Dock = DockStyle.Top, Height = 100, BackColor = Color.FromArgb(33, 33, 36) };

            dynLblToplamCiro = SkorKartYarat(pnlSkorlar, "BRÜT CİRO", "0,00 TL", new Point(25, 15), Color.FromArgb(0, 120, 215));
            dynLblToplamMaliyet = SkorKartYarat(pnlSkorlar, "ÜRÜN MALİYETİ", "0,00 TL", new Point(245, 15), Color.FromArgb(230, 126, 34));
            dynLblToplamKomisyon = SkorKartYarat(pnlSkorlar, "PLATFORM KESİNTİLERİ", "0,00 TL", new Point(465, 15), Color.FromArgb(155, 89, 182));
            dynLblNetKar = SkorKartYarat(pnlSkorlar, "NET KÂR / ZARAR", "0,00 TL", new Point(685, 15), Color.FromArgb(46, 117, 89));

            this.Controls.Add(pnlSkorlar);

            GroupBox grpTable = new GroupBox { Location = new Point(25, 175), Width = 960, Height = 420, Text = "📋 Sipariş Bazlı Net Kârlılık Analizi", Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Gray };

            dynDgvRapor = new DataGridView
            {
                Location = new Point(15, 30),
                Width = 930,
                Height = 320,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(26, 26, 28),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9.5F),
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, SelectionBackColor = Color.FromArgb(46, 117, 89) },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(21, 21, 33), ForeColor = Color.White, Font = new Font("Segoe UI", 9F, FontStyle.Bold) },
                EnableHeadersVisualStyles = false
            };
            grpTable.Controls.Add(dynDgvRapor);

            dynBtnExcel = new Button
            {
                Text = "📊 Mali Raporu Excel'e Aktar",
                Location = new Point(15, 365),
                Width = 220,
                Height = 40,
                BackColor = Color.FromArgb(39, 174, 96),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            dynBtnExcel.FlatAppearance.BorderSize = 0;
            dynBtnExcel.Click += (s, e) => ExcelMotoru.DataGridViewTablosunuDisaAktar(dynDgvRapor, "Net_Kar_Zarar_Mali_Rapor");
            grpTable.Controls.Add(dynBtnExcel);

            this.Controls.Add(grpTable);
        }

        private Label SkorKartYarat(Panel anaPanel, string baslik, string deger, Point konum, Color seritRengi)
        {
            Panel p = new Panel { Location = konum, Size = new Size(200, 70), BackColor = Color.FromArgb(26, 26, 28) };
            Panel pLine = new Panel { Dock = DockStyle.Left, Width = 4, BackColor = seritRengi };
            p.Controls.Add(pLine);

            Label lblB = new Label { Text = baslik, Location = new Point(12, 10), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.Gray };
            Label lblD = new Label { Text = deger, Location = new Point(12, 32), AutoSize = true, Font = new Font("Segoe UI", 13F, FontStyle.Bold), ForeColor = Color.White };

            p.Controls.Add(lblB); p.Controls.Add(lblD);
            anaPanel.Controls.Add(p);

            return lblD;
        }

        private void NetFinansalAnalizYap()
        {
            DataTable dtRapor = new DataTable();
            dtRapor.Columns.Add("Sipariş No");
            dtRapor.Columns.Add("Müşteri");
            dtRapor.Columns.Add("Platform");
            dtRapor.Columns.Add("Satış Tutar (TL)", typeof(double));
            dtRapor.Columns.Add("Ürün Maliyet (TL)", typeof(double));
            dtRapor.Columns.Add("Plat. Komisyon (TL)", typeof(double));
            dtRapor.Columns.Add("Plat. Sabit Kesinti (TL)", typeof(double));
            dtRapor.Columns.Add("NET KÂR (TL)", typeof(double));

            double toplamBrutCiro = 0; double toplamUrunMaliyeti = 0;
            double toplamKesinti = 0; double toplamNetKar = 0;

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    // IFNULL ile null gelen sütunları sıfıra çevirerek crash yemesini %100 önledik
                    string sql = @"SELECT s.SiparisNo, s.MusteriAdSoyad, s.Platform, 
                                          IFNULL(s.SatisFiyati, 0) as SatisFiyati, 
                                          IFNULL(st.AlisFiyati, 0) as UrunMaliyeti,
                                          IFNULL(p.KomisyonOrani, 0) as KomisyonOrani, 
                                          IFNULL(p.HizmetBedeli, 0) as HizmetBedeli
                                   FROM Siparisler s
                                   LEFT JOIN Stoklar st ON s.Barkod = st.Barkod
                                   LEFT JOIN Platformlar p ON s.Platform = p.Ad
                                   ORDER BY s.SiparisNo DESC";

                    using (var cmd = new SqliteCommand(sql, b))
                    using (var oku = cmd.ExecuteReader())
                    {
                        while (oku.Read())
                        {
                            string sipNo = oku["SiparisNo"].ToString();
                            string musteri = oku["MusteriAdSoyad"].ToString();
                            string platform = oku["Platform"].ToString();

                            double satisFiyati = Convert.ToDouble(oku["SatisFiyati"]);
                            double urunMaliyeti = Convert.ToDouble(oku["UrunMaliyeti"]);
                            double komisyonOrani = Convert.ToDouble(oku["KomisyonOrani"]);
                            double hizmetBedeli = Convert.ToDouble(oku["HizmetBedeli"]);

                            double hesaplananKomisyonTL = (satisFiyati * komisyonOrani) / 100.0;
                            double netKar = satisFiyati - urunMaliyeti - hesaplananKomisyonTL - hizmetBedeli;

                            dtRapor.Rows.Add(sipNo, musteri, platform, satisFiyati, urunMaliyeti, hesaplananKomisyonTL, hizmetBedeli, netKar);

                            toplamBrutCiro += satisFiyati;
                            toplamUrunMaliyeti += urunMaliyeti;
                            toplamKesinti += (hesaplananKomisyonTL + hizmetBedeli);
                            toplamNetKar += netKar;
                        }
                    }
                }

                dynDgvRapor.DataSource = dtRapor;
                FormatlaParaSutunlari();

                dynLblToplamCiro.Text = toplamBrutCiro.ToString("N2") + " TL";
                dynLblToplamMaliyet.Text = toplamUrunMaliyeti.ToString("N2") + " TL";
                dynLblToplamKomisyon.Text = toplamKesinti.ToString("N2") + " TL";
                dynLblNetKar.Text = toplamNetKar.ToString("N2") + " TL";

                if (toplamNetKar < 0) dynLblNetKar.ForeColor = Color.FromArgb(220, 53, 69);
                else dynLblNetKar.ForeColor = Color.FromArgb(46, 117, 89);

            }
            catch (Exception ex) { MessageBox.Show("Siparişler Tablonuz Boş Olabilir veya Finans Motoru Hatası: " + ex.Message); }
        }

        private void FormatlaParaSutunlari()
        {
            try
            {
                string[] paraSutunlari = { "Satış Tutar (TL)", "Ürün Maliyet (TL)", "Plat. Komisyon (TL)", "Plat. Sabit Kesinti (TL)", "NET KÂR (TL)" };
                foreach (var sutun in paraSutunlari)
                {
                    if (dynDgvRapor.Columns[sutun] != null) dynDgvRapor.Columns[sutun].DefaultCellStyle.Format = "N2";
                }
            }
            catch { }
        }
    }
}