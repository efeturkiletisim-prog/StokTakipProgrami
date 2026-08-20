using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // 💳 SİPARİŞ ÖDEME VE FİNANSAL TAKİP DENETİM MERKEZİ (DERLEME HATASI DÜZELTİLDİ)
    public partial class SayfaSiparisOdemeleri : UserControl
    {
        private DataGridView dynDgvOdemeler;
        private TextBox dynTxtSipNoAra;

        public SayfaSiparisOdemeleri()
        {
            InitializeComponent();
            this.Load += SayfaSiparisOdemeleri_Load;
        }

        private void SayfaSiparisOdemeleri_Load(object sender, EventArgs e)
        {
            ArayuzInsaEt();
            SiparisleriListele();
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(33, 33, 36);
            this.Padding = new Padding(15);

            // ==========================================
            // ÜST ARAMA BARI PANELİ
            // ==========================================
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15, 0, 15, 0) };

            Label lblT = new Label { Text = "💳 SİPARİŞ ÖDEME VE TAHSİLAT TAKİP PANELİ", Dock = DockStyle.Left, Width = 350, Font = new Font("Segoe UI", 11.5F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleLeft };

            dynTxtSipNoAra = new TextBox { Width = 220, Font = new Font("Segoe UI", 11F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            dynTxtSipNoAra.Location = new Point(pnlHeader.Width - 235, 17);
            dynTxtSipNoAra.Anchor = AnchorStyles.Right;
            dynTxtSipNoAra.TextChanged += (s, e) => SiparisleriListele();

            Label lblA = new Label { Text = "🔍 Sipariş No Ara: ", Width = 130, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray, TextAlign = ContentAlignment.MiddleRight };
            lblA.Location = new Point(pnlHeader.Width - 370, 17);
            lblA.Anchor = AnchorStyles.Right;

            pnlHeader.Controls.Add(lblT);
            pnlHeader.Controls.Add(dynTxtSipNoAra);
            pnlHeader.Controls.Add(lblA);
            this.Controls.Add(pnlHeader);

            // ==========================================
            // MERKEZİ MATRİS TABLOSU (DATA GRID VIEW)
            // ==========================================
            dynDgvOdemeler = new DataGridView
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
                ColumnHeadersHeight = 38
            };

            dynDgvOdemeler.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dynDgvOdemeler.ColumnHeadersDefaultCellStyle.ForeColor = Color.Orange;
            dynDgvOdemeler.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            dynDgvOdemeler.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dynDgvOdemeler.DefaultCellStyle.ForeColor = Color.White;
            dynDgvOdemeler.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dynDgvOdemeler.RowTemplate.Height = 30;

            dynDgvOdemeler.DataBindingComplete += DynDgvOdemeler_DataBindingComplete;
            dynDgvOdemeler.CellDoubleClick += DynDgvOdemeler_CellDoubleClick;

            this.Controls.Add(dynDgvOdemeler);

            dynDgvOdemeler.BringToFront();
            pnlHeader.SendToBack();
        }

        private void SiparisleriListele()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = @"
                        SELECT s.SiparisNo as 'Sipariş No', 
                               s.Tarih as 'Sipariş Tarihi', 
                               s.PlatformAd as 'Platform', 
                               s.Barkod as 'Barkod No', 
                               s.Adet as 'Satış Adedi', 
                               (s.SatisFiyati * s.Adet) as 'Beklenen Tahsilat (TL)', 
                               CASE 
                                  WHEN UPPER(s.Durum) = 'TAHSİL EDİLDİ' OR UPPER(s.Durum) = 'ÖDENDİ' THEN 'ÖDENDİ' 
                                  ELSE 'ÖDEME BEKLENEN' 
                               END as 'Ödeme Durumu'  
                        FROM Siparisler s 
                        WHERE s.SiparisNo LIKE @ara 
                        ORDER BY s.rowid DESC";

                    using (var cmd = new SqliteCommand(sql, b))
                    {
                        cmd.Parameters.AddWithValue("@ara", "%" + dynTxtSipNoAra.Text.Trim() + "%");
                        DataTable dt = new DataTable();
                        using (var reader = cmd.ExecuteReader()) { dt.Load(reader); }
                        dynDgvOdemeler.DataSource = dt;

                        if (dynDgvOdemeler.Columns["Beklenen Tahsilat (TL)"] != null)
                        {
                            dynDgvOdemeler.Columns["Beklenen Tahsilat (TL)"].DefaultCellStyle.Format = "N2";
                        }
                    }
                }
            }
            catch { }
        }

        private void DynDgvOdemeler_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow r in dynDgvOdemeler.Rows)
            {
                if (r.Cells["Ödeme Durumu"].Value != null)
                {
                    string st = r.Cells["Ödeme Durumu"].Value.ToString();
                    if (st == "ÖDEME BEKLENEN")
                    {
                        r.DefaultCellStyle.ForeColor = Color.Coral;
                    }
                    else if (st == "ÖDENDİ")
                    {
                        r.DefaultCellStyle.ForeColor = Color.FromArgb(34, 197, 94);
                    }
                }
            }
        }

        private void DynDgvOdemeler_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            try
            {
                string sipNo = dynDgvOdemeler.Rows[e.RowIndex].Cells["Sipariş No"].Value.ToString();
                string barkod = dynDgvOdemeler.Rows[e.RowIndex].Cells["Barkod No"].Value.ToString();

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT s.SiparisNo, s.PlatformAd, s.Barkod, u.UrunAdi, u.Kategori, IFNULL(u.Marka, '-') as Marka,
                               s.Adet, s.SatisFiyati, s.PlatformIndirimi, s.Stopaj, s.PosBedeli, s.HizmetBedeli, s.KomisyonTL, s.KargoUcreti, 
                               s.SatilanImei, s.Durum, s.Tarih, IFNULL(st.AlisFiyati, 0) as AlisFiyati
                        FROM Siparisler s
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod
                        LEFT JOIN Stoklar st ON s.Barkod = st.Barkod
                        WHERE s.SiparisNo = @sn AND s.Barkod = @b LIMIT 1";

                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@sn", sipNo);
                        cmd.Parameters.AddWithValue("@b", barkod);

                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                string platform = r["PlatformAd"].ToString();
                                string urunAdi = r["UrunAdi"]?.ToString() ?? "BİLİNMEYEN ÜRÜN";
                                string kat = r["Kategori"]?.ToString() ?? "-";
                                string marka = r["Marka"]?.ToString() ?? "-";
                                int adet = Convert.ToInt32(r["Adet"]);
                                double satis = Convert.ToDouble(r["SatisFiyati"]);
                                double indirim = Convert.ToDouble(r["PlatformIndirimi"]);
                                double alis = Convert.ToDouble(r["AlisFiyati"]);
                                double stopaj = Convert.ToDouble(r["Stopaj"]);
                                double pos = Convert.ToDouble(r["PosBedeli"]);
                                double hizmet = Convert.ToDouble(r["HizmetBedeli"]);
                                double komisyon = Convert.ToDouble(r["KomisyonTL"]);
                                double kargo = Convert.ToDouble(r["KargoUcreti"]);
                                string imei = r["SatilanImei"]?.ToString() ?? "-";
                                string durum = r["Durum"]?.ToString() ?? "OLUŞTURULDU";
                                string tarih = r["Tarih"]?.ToString() ?? "-";

                                double netKar = 0;
                                if (durum == "İADE")
                                {
                                    netKar = -(stopaj + pos + hizmet + komisyon + kargo);
                                }
                                else if (durum != "İPTAL")
                                {
                                    netKar = ((satis - indirim - alis) * adet) - stopaj - pos - hizmet - komisyon - kargo;
                                }

                                Form frmDetay = new Form
                                {
                                    Text = $"📄 Detaylı Finansal Sipariş Analiz Raporu - No: {sipNo}",
                                    Size = new Size(720, 520),
                                    StartPosition = FormStartPosition.CenterParent,
                                    BackColor = Color.FromArgb(33, 33, 36),
                                    ForeColor = Color.White,
                                    ShowIcon = false,
                                    MaximizeBox = false,
                                    MinimizeBox = false,
                                    FormBorderStyle = FormBorderStyle.FixedDialog,
                                    Font = new Font("Segoe UI", 10F)
                                };

                                GroupBox grpGenel = new GroupBox { Text = "📦 Genel ve Lojistik Bilgiler", Location = new Point(20, 15), Width = 330, Height = 360, ForeColor = Color.FromArgb(56, 189, 248) };
                                grpGenel.Controls.Add(new Label { Text = $"Sipariş No:\n{sipNo}", Location = new Point(15, 25), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White });
                                grpGenel.Controls.Add(new Label { Text = $"Satış Kanalı: {platform}", Location = new Point(15, 75), AutoSize = true, ForeColor = Color.LightGray });
                                grpGenel.Controls.Add(new Label { Text = $"Sipariş Tarihi: {tarih}", Location = new Point(15, 100), AutoSize = true, ForeColor = Color.LightGray });

                                Label lblDurum = new Label { Text = $"SİSTEM DURUMU: {durum}", Location = new Point(15, 135), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = (durum == "İPTAL" || durum == "İADE") ? Color.Crimson : Color.Gold };
                                grpGenel.Controls.Add(lblDurum);
                                grpGenel.Controls.Add(new Label { Text = "──────────────────────────", Location = new Point(15, 165), AutoSize = true, ForeColor = Color.FromArgb(50, 50, 50) });
                                grpGenel.Controls.Add(new Label { Text = $"Ürün Adı: {urunAdi}", Location = new Point(15, 190), Width = 300, Height = 40, ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) });
                                grpGenel.Controls.Add(new Label { Text = $"Kategori: {kat}", Location = new Point(15, 240), AutoSize = true, ForeColor = Color.DarkGray });
                                grpGenel.Controls.Add(new Label { Text = $"Marka: {marka}", Location = new Point(15, 265), AutoSize = true, ForeColor = Color.DarkGray });
                                grpGenel.Controls.Add(new Label { Text = $"Barkod No: {barkod}", Location = new Point(15, 290), AutoSize = true, ForeColor = Color.DarkGray });
                                grpGenel.Controls.Add(new Label { Text = $"Cihaz IMEI: {imei}", Location = new Point(15, 315), AutoSize = true, ForeColor = imei == "-" ? Color.DarkGray : Color.Coral, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) });
                                frmDetay.Controls.Add(grpGenel);

                                GroupBox grpFinans = new GroupBox { Text = "💰 Finansal Hesap Dökümü", Location = new Point(365, 15), Width = 320, Height = 360, ForeColor = Color.FromArgb(34, 197, 94) };
                                int y = 30;

                                // 🚀 İYİLEŞTİRME: SandyBrown hatasının temizlendiği güvenli satır üretici
                                Action<string, double, Color> FinansSatiriEkle = (baslik, deger, renk) => {
                                    Label lblB = new Label { Text = baslik, Location = new Point(15, y), AutoSize = true, ForeColor = Color.LightGray };
                                    Label lblD = new Label { Text = deger.ToString("N2") + " TL", Location = new Point(200, y), Width = 100, TextAlign = ContentAlignment.MiddleRight, ForeColor = renk, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
                                    grpFinans.Controls.Add(lblB);
                                    grpFinans.Controls.Add(lblD);
                                    y += 28;
                                };

                                if (durum == "İADE")
                                {
                                    FinansSatiriEkle("Brüt Satış (İADE EDİLDİ):", 0, Color.White);
                                    FinansSatiriEkle("Platform İndirimi (-):", 0, Color.LightCoral);
                                }
                                else
                                {
                                    FinansSatiriEkle($"Brüt Satış ({adet} Adet):", satis * adet, Color.White);
                                    FinansSatiriEkle("Platform İndirimi (-):", indirim * adet, Color.LightCoral);
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
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sipariş detay dökümü getirilirken hata oluştu:\n" + ex.Message, "Sistem Uyarısı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}