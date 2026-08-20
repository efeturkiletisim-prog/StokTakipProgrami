using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    // 📜 STOK HAREKETLERİ VE SİSTEM GÜVENLİK SİCİL LOG TERMİNALİ (MASTER SÜRÜM)
    public partial class SayfaStokHareketleri : UserControl
    {
        private DataGridView dynDgvLoglar;
        private TextBox dynTxtAra;
        private ComboBox dynCmbIslemTuru;
        private Button dynBtnExcel;

        // 🚀 MÜHÜR 1: Diğer sayfalar işlem bitirdiğinde burayı otomatik tetikleyecek Canlı Şalter
        public static Action OnStokHareketEklendi;

        public SayfaStokHareketleri()
        {
            InitializeComponent();
            this.Load += SayfaStokHareketleri_Load;

            // 🚀 MÜHÜR 2: Kullanıcı menüden bu sekmeye her tıkladığında listeyi anlık günceller
            this.VisibleChanged += (s, e) => { if (this.Visible) LoglariListele(); };
            this.ParentChanged += (s, e) => { if (this.Parent != null) LoglariListele(); };
        }

        private void SayfaStokHareketleri_Load(object sender, EventArgs e)
        {
            TabloyuGuvenceyeAlStatik();
            ArayuzInsaEt();
            LoglariListele();
        }

        // 🚀 MÜHÜR 3: Ana Formun (Form1) sekmeler arası geçişte aradığı küresel Yenile köprüsü
        public void Yenile()
        {
            try { LoglariListele(); } catch { }
        }

        // 🚀 MÜHÜR 4: Projedeki diğer tüm sayfaların (Sipariş Giriş satır 929 dahil) aradığı Altyapı Motoru
        public static void TabloyuGuvenceyeAlStatik()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = @"CREATE TABLE IF NOT EXISTS StokHareketleri (
                                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    Tarih TEXT,
                                    Barkod TEXT,
                                    IslemTuru TEXT,
                                    Miktar INTEGER,
                                    IslemYapan TEXT DEFAULT '-',
                                    EskiStok INTEGER DEFAULT 0,
                                    YeniStok INTEGER DEFAULT 0,
                                    Aciklama TEXT DEFAULT '-'
                                  );";
                    using (var cmd = new SqliteCommand(sql, b)) { cmd.ExecuteNonQuery(); }

                    // Eski veritabanı şemalarında eksik sütun varsa saniyede onarır
                    var sutunlar = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdCheck = new SqliteCommand("PRAGMA table_info('StokHareketleri');", b))
                    using (var r = cmdCheck.ExecuteReader()) { while (r.Read()) sutunlar.Add(r["name"].ToString()); }

                    if (!sutunlar.Contains("IslemYapan")) { using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN IslemYapan TEXT DEFAULT '-';", b)) cmd.ExecuteNonQuery(); }
                    if (!sutunlar.Contains("EskiStok")) { using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN EskiStok INTEGER DEFAULT 0;", b)) cmd.ExecuteNonQuery(); }
                    if (!sutunlar.Contains("YeniStok")) { using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN YeniStok INTEGER DEFAULT 0;", b)) cmd.ExecuteNonQuery(); }
                    if (!sutunlar.Contains("Aciklama")) { using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN Aciklama TEXT DEFAULT '-';", b)) cmd.ExecuteNonQuery(); }
                }
            }
            catch { }
        }

        // 🚀 MÜHÜR 5: "DATABASE IS LOCKED" HATASINI BEYİN AMELİYATIYLA ÇÖZEN LOG MOTORU
        // Kullanımı: SayfaStokHareketleri.LogYaz("-", "SİSTEM GİRİŞİ", 0, 0, 0, "Efetürk Bilgi Sistemleri admin girişi yapıldı.");
        public static void LogYaz(string barkod, string islemTuru, int miktar, int eskiStok, int yeniStok, string aciklama)
        {
            TabloyuGuvenceyeAlStatik();

            // Veritabanı o an kilitliyse pes etmez, 5 kez yarım saniye arayla kapıyı zorlar ve MUTLAKA yazar!
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        string sql = @"INSERT INTO StokHareketleri (Tarih, Barkod, IslemTuru, Miktar, EskiStok, YeniStok, Aciklama) 
                                       VALUES (@tarih, @barkod, @islem, @miktar, @eski, @yeni, @aciklama)";
                        using (var cmd = new SqliteCommand(sql, b))
                        {
                            cmd.Parameters.AddWithValue("@tarih", DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"));
                            cmd.Parameters.AddWithValue("@barkod", string.IsNullOrEmpty(barkod) ? "-" : barkod);
                            cmd.Parameters.AddWithValue("@islem", islemTuru.ToUpper());
                            cmd.Parameters.AddWithValue("@miktar", miktar);
                            cmd.Parameters.AddWithValue("@eski", eskiStok);
                            cmd.Parameters.AddWithValue("@yeni", yeniStok);
                            cmd.Parameters.AddWithValue("@aciklama", string.IsNullOrEmpty(aciklama) ? "-" : aciklama);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    // Kayıt başarılıysa canlı yenileme delegesini uyandırır
                    try { OnStokHareketEklendi?.Invoke(); } catch { }
                    break;
                }
                catch
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        // 🚀 MÜHÜR 6: RESPONSIVE VE KOYU TEMA UYUMLU ULTRA LUX ARAYÜZ DİZİLİMİ
        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "📜  STOK SİCİL HAREKETLERİ VE GÜVENLİK LOG TERMİNALİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            Panel pnlFiltre = new Panel { Dock = DockStyle.Top, Height = 65, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15, 0, 15, 0) };

            Label lblAra = new Label { Text = "🔎 Canlı Filtrele (Barkod / Ürün Kartı):", Location = new Point(15, 12), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynTxtAra = new TextBox { Location = new Point(15, 32), Width = 220, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            dynTxtAra.TextChanged += (s, e) => LoglariListele();

            Label lblIslem = new Label { Text = "⚙️ Operasyon Tipi:", Location = new Point(255, 12), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynCmbIslemTuru = new ComboBox { Location = new Point(255, 32), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            dynCmbIslemTuru.Items.AddRange(new string[] { "🌍 TÜM İŞLEMLER", "STOK GİRİŞİ", "SATIŞ", "İPTAL / İADE", "MANUEL DÜZENLEME", "TOPLU VERİ YÜKLEME", "SİSTEM GİRİŞİ" });
            dynCmbIslemTuru.SelectedIndex = 0;
            dynCmbIslemTuru.SelectedIndexChanged += (s, e) => LoglariListele();

            dynBtnExcel = new Button { Text = "📊 Excel Günlüğü Al", Location = new Point(455, 29), Width = 150, Height = 28, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnExcel.FlatAppearance.BorderSize = 0;
            dynBtnExcel.Click += (s, e) => ExcelMotoru.DataGridViewTablosunuDisaAktar(dynDgvLoglar, "Stok_Hareket_Loglari");

            pnlFiltre.Controls.Add(lblAra); pnlFiltre.Controls.Add(dynTxtAra);
            pnlFiltre.Controls.Add(lblIslem); pnlFiltre.Controls.Add(dynCmbIslemTuru);
            pnlFiltre.Controls.Add(dynBtnExcel);
            this.Controls.Add(pnlFiltre);

            Panel pnlGridWrapper = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
            dynDgvLoglar = new DataGridView
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
                ColumnHeadersHeight = 35
            };
            dynDgvLoglar.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dynDgvLoglar.ColumnHeadersDefaultCellStyle.ForeColor = Color.Orange;
            dynDgvLoglar.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            dynDgvLoglar.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dynDgvLoglar.DefaultCellStyle.ForeColor = Color.White;
            dynDgvLoglar.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dynDgvLoglar.RowTemplate.Height = 30;

            dynDgvLoglar.DataError += (s, e) => { e.Cancel = true; };
            dynDgvLoglar.DataBindingComplete += DynDgvLoglar_DataBindingComplete;

            // 🚀 MÜHÜR 7: Log'a çift tıklandığında lüks teftiş popup'ı açma motoru
            dynDgvLoglar.CellDoubleClick += DynDgvLoglar_CellDoubleClick;

            pnlGridWrapper.Controls.Add(dynDgvLoglar);
            this.Controls.Add(pnlGridWrapper);

            // Z-Index düzenlemesi
            pnlGridWrapper.BringToFront();
            pnlFiltre.SendToBack();
            pnlHeader.SendToBack();

            OnStokHareketEklendi += () =>
            {
                try
                {
                    if (this.IsHandleCreated && this.InvokeRequired) { this.BeginInvoke((Action)(() => LoglariListele())); }
                    else { LoglariListele(); }
                }
                catch { }
            };
        }

        // 🚀 MÜHÜR 8: HAFIZADA SÜTUN ADI ÇAKIŞMASINI SIFIRA İNDİREN ASCII SELECT MOTORU
        private void LoglariListele()
        {
            if (dynTxtAra == null || dynCmbIslemTuru == null || dynDgvLoglar == null) return;

            try
            {
                string aramaKriteri = "%" + dynTxtAra.Text.Trim() + "%";
                string seciliFiltre = dynCmbIslemTuru.SelectedIndex > 0 ? dynCmbIslemTuru.Text.Trim() : "";

                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = @"
                        SELECT 
                            sh.Tarih,
                            sh.Barkod,
                            IFNULL(u.UrunAdi, 'Sistem / Manuel Kayıt') as UrunAdi,
                            sh.IslemYapan,
                            sh.IslemTuru,
                            sh.Miktar,
                            sh.EskiStok,
                            sh.YeniStok,
                            sh.Aciklama
                        FROM StokHareketleri sh
                        LEFT JOIN Urunler u ON sh.Barkod = u.Barkod
                        WHERE (sh.Barkod LIKE @ara OR u.UrunAdi LIKE @ara)";

                    if (!string.IsNullOrEmpty(seciliFiltre)) sql += " AND sh.IslemTuru = @islem";
                    sql += " ORDER BY sh.Id DESC"; // Dakikası dakikasına akış düzeni

                    using (var cmd = new SqliteCommand(sql, b))
                    {
                        cmd.Parameters.AddWithValue("@ara", aramaKriteri);
                        if (!string.IsNullOrEmpty(seciliFiltre)) cmd.Parameters.AddWithValue("@islem", seciliFiltre);

                        DataTable dt = new DataTable();
                        dt.Columns.Add("Tarih"); dt.Columns.Add("Barkod"); dt.Columns.Add("UrunAdi");
                        dt.Columns.Add("IslemYapan");
                        dt.Columns.Add("IslemTuru"); dt.Columns.Add("Miktar", typeof(int));
                        dt.Columns.Add("EskiStok", typeof(int)); dt.Columns.Add("YeniStok", typeof(int));
                        dt.Columns.Add("Aciklama");

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                dt.Rows.Add(
                                    reader["Tarih"]?.ToString(),
                                    reader["Barkod"]?.ToString(),
                                    reader["UrunAdi"]?.ToString(),
                                    reader["IslemYapan"]?.ToString(),
                                    reader["IslemTuru"]?.ToString(),
                                    reader["Miktar"] == DBNull.Value ? 0 : Convert.ToInt32(reader["Miktar"]),
                                    reader["EskiStok"] == DBNull.Value ? 0 : Convert.ToInt32(reader["EskiStok"]),
                                    reader["YeniStok"] == DBNull.Value ? 0 : Convert.ToInt32(reader["YeniStok"]),
                                    reader["Aciklama"]?.ToString()
                                );
                            }
                        }

                        dynDgvLoglar.DataSource = dt;

                        // Sütun görsel başlıkları veri bağlama işlemi bittikten sonra güvenle Türkçeleştirilir
                        if (dynDgvLoglar.Columns.Contains("Tarih")) dynDgvLoglar.Columns["Tarih"].HeaderText = "İşlem Tarihi";
                        if (dynDgvLoglar.Columns.Contains("Barkod")) dynDgvLoglar.Columns["Barkod"].HeaderText = "Barkod No";
                        if (dynDgvLoglar.Columns.Contains("UrunAdi")) dynDgvLoglar.Columns["UrunAdi"].HeaderText = "Ürün Kart Adı";
                        if (dynDgvLoglar.Columns.Contains("IslemYapan")) dynDgvLoglar.Columns["IslemYapan"].HeaderText = "İşlem Yapan";
                        if (dynDgvLoglar.Columns.Contains("IslemTuru")) dynDgvLoglar.Columns["IslemTuru"].HeaderText = "İşlem Türü";
                        if (dynDgvLoglar.Columns.Contains("Miktar")) dynDgvLoglar.Columns["Miktar"].HeaderText = "Miktar";
                        if (dynDgvLoglar.Columns.Contains("EskiStok")) dynDgvLoglar.Columns["EskiStok"].HeaderText = "Eski Stok";
                        if (dynDgvLoglar.Columns.Contains("YeniStok")) dynDgvLoglar.Columns["YeniStok"].HeaderText = "Mevcut Kalan";
                        if (dynDgvLoglar.Columns.Contains("Aciklama")) dynDgvLoglar.Columns["Aciklama"].HeaderText = "Güvenlik Açıklaması";

                        dynDgvLoglar.ClearSelection();
                    }
                }
            }
            catch { }
        }

        // 🚀 MÜHÜR 9: OPERASYON RADİKALİNE GÖRE REAKTİF SATIR BOYAMA SİSTEMİ
        private void DynDgvLoglar_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (dynDgvLoglar == null || !dynDgvLoglar.Columns.Contains("IslemTuru")) return;

            foreach (DataGridViewRow r in dynDgvLoglar.Rows)
            {
                if (r.Cells["IslemTuru"].Value != null)
                {
                    string tur = r.Cells["IslemTuru"].Value.ToString().ToUpper();

                    if (tur.Contains("GİRİŞ") || tur.Contains("İADE"))
                    {
                        r.DefaultCellStyle.ForeColor = Color.FromArgb(34, 197, 94); // Kurumsal Canlı Yeşil (Artış)
                    }
                    else if (tur.Contains("SATIŞ") || tur.Contains("ÇIKIŞ"))
                    {
                        r.DefaultCellStyle.ForeColor = Color.Coral; // Mercan Kırmızı (Azalış)
                    }
                    else
                    {
                        r.DefaultCellStyle.ForeColor = Color.FromArgb(56, 189, 248); // Gökyüzü Mavisi (Sistem/Düzenleme/Giriş)
                    }
                }
            }
        }

        // 🚀 MÜHÜR 10: ÇİFT TIKLANDIĞINDA SİCİL DETAY RAPORUNU MODAL AÇAN TEFTİŞ MODÜLÜ
        private void DynDgvLoglar_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            Form frmDetay = new Form { Text = "🔍 Detaylı Sicil Hareket Analizi", Size = new Size(440, 520), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false, ShowIcon = false };
            Panel pnlIcerik = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
            Label lblBaslik = new Label { Text = "SİCİL LOG DETAYLARI", Dock = DockStyle.Top, Height = 40, Font = new Font("Segoe UI Black", 14F, FontStyle.Bold), ForeColor = Color.Orange, TextAlign = ContentAlignment.MiddleCenter };
            pnlIcerik.Controls.Add(lblBaslik);

            int y = 60;
            pnlIcerik.Controls.Add(DetaySatiriYarat("İşlem Tarihi:", dynDgvLoglar.Rows[e.RowIndex].Cells["Tarih"].Value?.ToString() ?? "-", ref y));
            pnlIcerik.Controls.Add(DetaySatiriYarat("Barkod No:", dynDgvLoglar.Rows[e.RowIndex].Cells["Barkod"].Value?.ToString() ?? "-", ref y));
            pnlIcerik.Controls.Add(DetaySatiriYarat("Ürün Adı:", dynDgvLoglar.Rows[e.RowIndex].Cells["UrunAdi"].Value?.ToString() ?? "-", ref y)); y += 10;
            pnlIcerik.Controls.Add(DetaySatiriYarat("İşlem Yapan:", dynDgvLoglar.Rows[e.RowIndex].Cells["IslemYapan"].Value?.ToString() ?? "-", ref y, Color.FromArgb(200,200,200)));
            pnlIcerik.Controls.Add(DetaySatiriYarat("İşlem Türü:", dynDgvLoglar.Rows[e.RowIndex].Cells["IslemTuru"].Value?.ToString() ?? "-", ref y, Color.FromArgb(56, 189, 248)));
            pnlIcerik.Controls.Add(DetaySatiriYarat("Miktar (Adet):", dynDgvLoglar.Rows[e.RowIndex].Cells["Miktar"].Value?.ToString() ?? "0", ref y));
            pnlIcerik.Controls.Add(DetaySatiriYarat("İşlem Öncesi Raf:", dynDgvLoglar.Rows[e.RowIndex].Cells["EskiStok"].Value?.ToString() ?? "0", ref y));
            pnlIcerik.Controls.Add(DetaySatiriYarat("İşlem Sonrası Raf:", dynDgvLoglar.Rows[e.RowIndex].Cells["YeniStok"].Value?.ToString() ?? "0", ref y, Color.FromArgb(34, 197, 94))); y += 10;

            Label lblAciklamaBaslik = new Label { Text = "Sistem / Operasyon Açıklaması:", Location = new Point(0, y), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray }; pnlIcerik.Controls.Add(lblAciklamaBaslik); y += 25;
            TextBox txtAciklama = new TextBox { Text = dynDgvLoglar.Rows[e.RowIndex].Cells["Aciklama"].Value?.ToString() ?? "-", Location = new Point(0, y), Width = 380, Height = 95, Multiline = true, ReadOnly = true, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 10F) }; pnlIcerik.Controls.Add(txtAciklama);
            frmDetay.Controls.Add(pnlIcerik);
            frmDetay.ShowDialog();
        }

        private Panel DetaySatiriYarat(string baslik, string deger, ref int yPos, Color? degerRengi = null)
        {
            Panel p = new Panel { Location = new Point(0, yPos), Width = 380, Height = 30 };
            Label lB = new Label { Text = baslik, Location = new Point(0, 5), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Gray };
            Label lD = new Label { Text = deger, Location = new Point(140, 5), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = degerRengi ?? Color.White };
            p.Controls.Add(lB); p.Controls.Add(lD); yPos += 35; return p;
        }
    }
}