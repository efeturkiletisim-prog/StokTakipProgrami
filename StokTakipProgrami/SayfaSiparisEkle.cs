using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;
using System.Linq;

namespace StokTakipProgrami
{
    // 🛒 SİPARİŞ GİRİŞ MERKEZİ (PAZARYERİ CEZA KESİNTİ SİSTEMİ VE YENİ STRATEJİK DİZİLİM)
    public partial class SayfaSiparisEkle : UserControl
    {
        private List<(string Kaynak, string Ad, string Tip, double Deger)> platformKesintiler = new List<(string, string, string, double)>();
        private DataGridView dynDgvPlatformKesintiler;

        private TextBox dynTxtBarkod, dynTxtAdet, dynTxtSatisFiyati, dynTxtPlatformIndirimi, dynTxtImei, dynTxtSiparisNo, dynTxtKargoUcreti;
        private ComboBox dynCmbPlatform, dynCmbKargoFirmasi, dynCmbDesi;

        private Label dynLblUrunAdi, dynLblKategori, dynLblMarka, dynLblMevcutStok, dynLblSepetToplam;
        private Button dynBtnAra, dynBtnSepeteEkle, dynBtnSatisiTamamla, dynBtnSepetiTemizle;
        private DateTimePicker dynDtSiparisTarihi;
        private Label dynLblCanliStopaj, dynLblCanliPos, dynLblCanliHizmet, dynLblCanliNetKar;
        private DataGridView dynDgvSepet;
        private DataTable sepetTablosu;

        private ContextMenuStrip cmsSepet;

        private string aktifKategori = "";
        private double seciliUrunAlisFiyati = 0;
        private int seciliUrunMevcutStok = 0;

        private double dbHizmetBedeli = 0;
        private double dbPosOrani = 0;
        private string dbPosModu = "ANA_TUTAR";

        private bool dbKomisyonMuaf = false;
        private bool dbStopajMuaf = false;

        private List<string> aktifImeiListesi = new List<string>();

        public SayfaSiparisEkle()
        {
            InitializeComponent();
            this.Load += SayfaSiparisEkle_Load;
        }

        private void SayfaSiparisEkle_Load(object sender, EventArgs e)
        {
            TabloAltyapisiniGuvenceyeAl();
            ArayuzInsaEt();
            SepetAltyapisiniKur();
            PlatformlariYukle();
            if (dynTxtBarkod != null) dynTxtBarkod.Focus();
        }

        private void TabloAltyapisiniGuvenceyeAl()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();

                    var pltSutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var cmd = new SqliteCommand("PRAGMA table_info('Platformlar');", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) pltSutunlar.Add(r["name"].ToString()); }
                    if (!pltSutunlar.Contains("KomisyonMuaf")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN KomisyonMuaf INTEGER DEFAULT 0;", conn)) cmd.ExecuteNonQuery(); }
                    if (!pltSutunlar.Contains("StopajMuaf")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN StopajMuaf INTEGER DEFAULT 0;", conn)) cmd.ExecuteNonQuery(); }

                    var mevcutSiparisSutunlari = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var cmd = new SqliteCommand("PRAGMA table_info('Siparisler');", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) mevcutSiparisSutunlari.Add(r["name"].ToString()); }

                    var gerekliSiparisKolonlar = new Dictionary<string, string> {
                        { "PlatformIndirimi", "REAL DEFAULT 0" },
                        { "Stopaj", "REAL DEFAULT 0" },
                        { "PosBedeli", "REAL DEFAULT 0" },
                        { "HizmetBedeli", "REAL DEFAULT 0" },
                        { "KomisyonTL", "REAL DEFAULT 0" },
                        { "KargoUcreti", "REAL DEFAULT 0" },
                        { "Desi", "TEXT" },
                        { "KargoFirmasi", "TEXT" },
                        { "ImeiNo", "TEXT" },
                        { "SatisFiyati", "REAL DEFAULT 0" },
                        { "AlisFiyati", "REAL DEFAULT 0" } // EKLENDİ!
                    };

                    foreach (var kv in gerekliSiparisKolonlar)
                    {
                        if (!mevcutSiparisSutunlari.Contains(kv.Key))
                        {
                            try { using (var cmd2 = new SqliteCommand($"ALTER TABLE Siparisler ADD COLUMN {kv.Key} {kv.Value};", conn)) cmd2.ExecuteNonQuery(); } catch { }
                        }
                    }

                    string sqlLog = @"CREATE TABLE IF NOT EXISTS StokHareketleri (Id INTEGER PRIMARY KEY AUTOINCREMENT, Tarih TEXT, Barkod TEXT, IslemTuru TEXT, Miktar INTEGER, EskiStok INTEGER DEFAULT 0, YeniStok INTEGER DEFAULT 0, Aciklama TEXT DEFAULT '-');";
                    using (var cmdLog = new SqliteCommand(sqlLog, conn)) { cmdLog.ExecuteNonQuery(); }

                    var logSutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var cmdCheck = new SqliteCommand("PRAGMA table_info('StokHareketleri');", conn))
                    using (var r = cmdCheck.ExecuteReader()) { while (r.Read()) logSutunlar.Add(r["name"].ToString()); }

                    if (!logSutunlar.Contains("EskiStok")) { using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN EskiStok INTEGER DEFAULT 0;", conn)) cmd.ExecuteNonQuery(); }
                    if (!logSutunlar.Contains("YeniStok")) { using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN YeniStok INTEGER DEFAULT 0;", conn)) cmd.ExecuteNonQuery(); }
                    if (!logSutunlar.Contains("Aciklama")) { using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN Aciklama TEXT DEFAULT '-';", conn)) cmd.ExecuteNonQuery(); }
                }
            }
            catch { }
        }

        private void PlatformlariYukle()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Ad FROM Platformlar ORDER BY Ad ASC", conn))
                    using (var oku = cmd.ExecuteReader())
                    {
                        dynCmbPlatform.Items.Clear();
                        dynCmbPlatform.Items.Add("");
                        bool efeturkVarMi = false;
                        while (oku.Read())
                        {
                            string pAd = oku["Ad"].ToString();
                            dynCmbPlatform.Items.Add(pAd);
                            if (pAd.ToUpper() == "EFETÜRK") efeturkVarMi = true;
                        }
                        if (!efeturkVarMi) dynCmbPlatform.Items.Add("EFETÜRK");
                        dynCmbPlatform.SelectedIndex = 0;
                    }
                }
            }
            catch { }
        }

        private void SepetAltyapisiniKur()
        {
            sepetTablosu = new DataTable();
            sepetTablosu.Columns.Add("Barkod");
            sepetTablosu.Columns.Add("Sipariş No");
            sepetTablosu.Columns.Add("Platform");
            sepetTablosu.Columns.Add("Ürün Adı");
            sepetTablosu.Columns.Add("Adet", typeof(int));
            sepetTablosu.Columns.Add("Satış (TL)", typeof(double));
            sepetTablosu.Columns.Add("İndirim (TL)", typeof(double));
            sepetTablosu.Columns.Add("Stopaj (TL)", typeof(double));
            sepetTablosu.Columns.Add("POS (TL)", typeof(double));
            sepetTablosu.Columns.Add("Hizmet (TL)", typeof(double));
            sepetTablosu.Columns.Add("Komisyon (TL)", typeof(double));
            sepetTablosu.Columns.Add("Kargo (TL)", typeof(double));
            sepetTablosu.Columns.Add("Transfer Tutarı (TL)", typeof(double)); // EKLENDİ (Platformdan Gelecek Para)
            sepetTablosu.Columns.Add("Toplam Maliyet (TL)", typeof(double));  // EKLENDİ (Ürün Alış Maliyeti)
            sepetTablosu.Columns.Add("IMEI");
            sepetTablosu.Columns.Add("Net Kâr (TL)", typeof(double));         // (Transfer Tutarı - Maliyet)
            sepetTablosu.Columns.Add("Kargo Firması");
            sepetTablosu.Columns.Add("Desi");
            sepetTablosu.Columns.Add("Alış Fiyatı", typeof(double));
            sepetTablosu.Columns.Add("Kategori");

            if (dynDgvSepet != null)
            {
                dynDgvSepet.DataSource = sepetTablosu;
                dynDgvSepet.Columns["Kargo Firması"].Visible = false;
                dynDgvSepet.Columns["Desi"].Visible = false;
                dynDgvSepet.Columns["Alış Fiyatı"].Visible = false;
                dynDgvSepet.Columns["Kategori"].Visible = false;
            }
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(33, 33, 36);
            this.Dock = DockStyle.Fill;

            cmsSepet = new ContextMenuStrip();
            cmsSepet.Items.Add("✏️ Düzenle", null, SepetDuzenle_Click);
            cmsSepet.Items.Add("❌ Sil", null, SepetSil_Click);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "🛒 YENİ SİPARİŞ KAYIT MERKEZİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleLeft };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 480F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.Controls.Add(mainLayout);
            mainLayout.BringToFront();

            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15), AutoScroll = true };

            Label lblB = new Label { Text = "Ürün Barkodu Okutun:", Location = new Point(15, 15), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynTxtBarkod = new TextBox { Location = new Point(15, 35), Width = 300, Font = new Font("Segoe UI", 12F, FontStyle.Bold), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.FromArgb(56, 189, 248), BorderStyle = BorderStyle.FixedSingle };
            dynTxtBarkod.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; UrunuSatisIcinCagir(); } };

            dynBtnAra = new Button { Text = "🔍 Ara", Location = new Point(325, 34), Width = 120, Height = 28, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnAra.FlatAppearance.BorderSize = 0; dynBtnAra.Click += (s, e) => UrunuSatisIcinCagir();
            pnlLeft.Controls.Add(lblB); pnlLeft.Controls.Add(dynTxtBarkod); pnlLeft.Controls.Add(dynBtnAra);

            Panel pnlBilgi = new Panel { Location = new Point(15, 75), Width = 430, Height = 80, BackColor = Color.FromArgb(33, 33, 36), BorderStyle = BorderStyle.FixedSingle };
            dynLblUrunAdi = new Label { Text = "Ürün bekleniyor...", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White };
            dynLblKategori = new Label { Text = "Kategori: -", Location = new Point(10, 30), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.Orange };
            dynLblMarka = new Label { Text = "Marka: -", Location = new Point(180, 30), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.Gold };
            dynLblMevcutStok = new Label { Text = "Depo Stoğu: -", Location = new Point(10, 55), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) };
            pnlBilgi.Controls.Add(dynLblUrunAdi); pnlBilgi.Controls.Add(dynLblKategori); pnlBilgi.Controls.Add(dynLblMarka); pnlBilgi.Controls.Add(dynLblMevcutStok);
            pnlLeft.Controls.Add(pnlBilgi);

            pnlLeft.Controls.Add(FormElemaniYarat("Sipariş No (Pazar Yeri Kodu):", out dynTxtSiparisNo, new Point(15, 165), 200, ""));
            pnlLeft.Controls.Add(FormComboYarat("Satış Kanalı / Platform:", out dynCmbPlatform, new Point(230, 165), 215));
            dynCmbPlatform.SelectedIndexChanged += DynCmbPlatform_SelectedIndexChanged;

            pnlLeft.Controls.Add(FormElemaniYarat("Adet:", out dynTxtAdet, new Point(15, 230), 70, "1"));
            pnlLeft.Controls.Add(FormElemaniYarat("Satış Fiyatı (TL):", out dynTxtSatisFiyati, new Point(100, 230), 165, "0,00"));
            pnlLeft.Controls.Add(FormElemaniYarat("Platform İndirimi (TL):", out dynTxtPlatformIndirimi, new Point(280, 230), 165, "0,00"));

            pnlLeft.Controls.Add(FormComboYarat("Kargo Firması:", out dynCmbKargoFirmasi, new Point(15, 295), 160));
            dynCmbKargoFirmasi.Items.Add("");
            dynCmbKargoFirmasi.SelectedIndex = 0;

            pnlLeft.Controls.Add(FormComboYarat("Desi:", out dynCmbDesi, new Point(185, 295), 80));
            for (int i = 1; i <= 30; i++) dynCmbDesi.Items.Add(i.ToString());
            dynCmbDesi.SelectedIndex = 0;
            dynCmbDesi.SelectedIndexChanged += (s, e) => { KargoUcretiniGetir(); };
            dynCmbKargoFirmasi.SelectedIndexChanged += (s, e) => KargoUcretiniGetir();



            pnlLeft.Controls.Add(FormElemaniYarat("Kargo Ücreti (TL):", out dynTxtKargoUcreti, new Point(275, 295), 170, "0,00"));
            // UX Formüllerini Devreye Al
            ParaKutusuFormatla(dynTxtSatisFiyati);
            ParaKutusuFormatla(dynTxtPlatformIndirimi);
            ParaKutusuFormatla(dynTxtKargoUcreti);
            pnlLeft.Controls.Add(FormElemaniYarat("IMEI Numarası (Telefon İçin):", out dynTxtImei, new Point(15, 360), 430, ""));
            dynTxtImei.Enabled = false;

            dynTxtAdet.TextChanged += (s, e) => {
                CanliFinansalAnalizYap();
                int adet = 0; int.TryParse(dynTxtAdet.Text, out adet);
                if (aktifKategori.Contains("TELEFON") || aktifKategori.Contains("CEP"))
                {
                    if (adet > 1)
                    {
                        dynTxtImei.ReadOnly = true;
                        dynTxtImei.Text = "Çoklu IMEI satışı için tıklayın...";
                    }
                    else
                    {
                        dynTxtImei.ReadOnly = false;
                        if (dynTxtImei.Text.Contains("tıklayın") || dynTxtImei.Text.Contains("GİRİLDİ")) dynTxtImei.Clear();
                        aktifImeiListesi.Clear();
                    }
                }
            };

            dynTxtImei.Click += (s, e) => {
                int adet = 0; int.TryParse(dynTxtAdet.Text, out adet);
                if (adet > 1 && (aktifKategori.Contains("TELEFON") || aktifKategori.Contains("CEP")))
                {
                    CokluImeiSatisPenceresiAc(adet);
                }
            };

            Label lblTarihSec = new Label { Text = "Sipariş Tarihi:", Location = new Point(15, 425), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynDtSiparisTarihi = new DateTimePicker { Location = new Point(15, 443), Width = 430, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Format = DateTimePickerFormat.Short };
            pnlLeft.Controls.Add(lblTarihSec); pnlLeft.Controls.Add(dynDtSiparisTarihi);

            dynBtnSepeteEkle = new Button { Text = "📥 SEPETE EKLE", Location = new Point(15, 500), Width = 430, Height = 50, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSepeteEkle.FlatAppearance.BorderSize = 0; dynBtnSepeteEkle.Click += DynBtnSepeteEkle_Click;
            pnlLeft.Controls.Add(dynBtnSepeteEkle);

            mainLayout.Controls.Add(pnlLeft, 0, 0);

            dynTxtSatisFiyati.TextChanged += (s, e) => CanliFinansalAnalizYap();
            dynTxtPlatformIndirimi.TextChanged += (s, e) => CanliFinansalAnalizYap();
            dynTxtKargoUcreti.TextChanged += (s, e) => CanliFinansalAnalizYap();

            // ==========================================
            // SAĞ PANEL
            // ==========================================
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Margin = new Padding(15) };

            Panel pnlFatura = new Panel { Dock = DockStyle.Bottom, Height = 80, BackColor = Color.FromArgb(26, 26, 28) };
            dynLblSepetToplam = new Label { Text = "0,00 TL / Kâr: 0,00 TL", AutoSize = true, Location = new Point(15, 25), Font = new Font("Segoe UI", 16F, FontStyle.Bold), ForeColor = Color.White };
            pnlFatura.Controls.Add(dynLblSepetToplam);

            dynBtnSatisiTamamla = new Button { Text = "🏁 SİPARİŞİ KAYDET", Width = 180, Height = 45, Anchor = AnchorStyles.Top | AnchorStyles.Right, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSatisiTamamla.Location = new Point(pnlFatura.Width - 195, 18);
            dynBtnSatisiTamamla.Click += DynBtnSatisiTamamla_Click;
            pnlFatura.Controls.Add(dynBtnSatisiTamamla);

            dynBtnSepetiTemizle = new Button { Text = "🗑️ Sepeti Boşalt", Width = 150, Height = 45, Anchor = AnchorStyles.Top | AnchorStyles.Right, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSepetiTemizle.Location = new Point(pnlFatura.Width - 355, 18);
            dynBtnSepetiTemizle.Click += (s, e) => { SiparisEkraniniTemizle(); }; pnlFatura.Controls.Add(dynBtnSepetiTemizle);

            pnlRight.Controls.Add(pnlFatura);

            Label lblRightTitle = new Label { Name = "lblRightTitle", Text = "SİPARİŞ SEPETİ VE FİNANSAL ÖZET", Dock = DockStyle.Top, Height = 45, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI Black", 15F, FontStyle.Bold), ForeColor = Color.Orange };
            pnlRight.Controls.Add(lblRightTitle);

            TableLayoutPanel tlpIndikator = new TableLayoutPanel { Dock = DockStyle.Top, Height = 80, BackColor = Color.FromArgb(26, 26, 28), ColumnCount = 4, RowCount = 1 };
            tlpIndikator.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpIndikator.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpIndikator.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            tlpIndikator.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));

            dynLblCanliStopaj = GöstergeEtiketiYarat(tlpIndikator, "STOPAJ KESİNTİSİ", "0,00 TL", 0, Color.DarkGray);
            dynLblCanliPos = GöstergeEtiketiYarat(tlpIndikator, "POS KESİNTİSİ", "0,00 TL", 1, Color.DarkGray);
            dynLblCanliHizmet = GöstergeEtiketiYarat(tlpIndikator, "HİZMET BEDELİ", "0,00 TL", 2, Color.DarkGray);
            dynLblCanliNetKar = GöstergeEtiketiYarat(tlpIndikator, "ANLIK KÂR", "0,00 TL", 3, Color.FromArgb(56, 189, 248));
            pnlRight.Controls.Add(tlpIndikator);

            Panel wrapKesinti = new Panel { Dock = DockStyle.Top, Height = 130, Padding = new Padding(0, 10, 0, 10) };
            dynDgvPlatformKesintiler = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                BackgroundColor = Color.FromArgb(26, 26, 28),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9F),
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(21, 21, 33), ForeColor = Color.White, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) },
                EnableHeadersVisualStyles = false
            };
            wrapKesinti.Controls.Add(dynDgvPlatformKesintiler);
            pnlRight.Controls.Add(wrapKesinti);

            dynDgvSepet = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(26, 26, 28),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9.5F),
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, SelectionBackColor = Color.FromArgb(230, 126, 34) },
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 42,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, Font = new Font("Segoe UI", 9F, FontStyle.Bold) },
                ContextMenuStrip = cmsSepet
            };

            dynDgvSepet.MouseDown += (s, e) => {
                if (e.Button == MouseButtons.Right)
                {
                    var h = dynDgvSepet.HitTest(e.X, e.Y);
                    if (h.Type == DataGridViewHitTestType.Cell && h.RowIndex >= 0)
                    {
                        dynDgvSepet.ClearSelection();
                        dynDgvSepet.Rows[h.RowIndex].Selected = true;
                    }
                }
            };

            dynDgvSepet.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dynDgvSepet.RowTemplate.Height = 32;
            dynDgvSepet.AllowUserToResizeRows = false;
            pnlRight.Controls.Add(dynDgvSepet);

            dynDgvSepet.BringToFront();
            wrapKesinti.SendToBack();
            tlpIndikator.SendToBack();
            lblRightTitle.SendToBack();

            mainLayout.Controls.Add(pnlRight, 1, 0);

            this.Controls.Add(mainLayout);
            this.Controls.Add(pnlHeader);
            pnlHeader.SendToBack();
            mainLayout.BringToFront();
        }

        private void CokluImeiSatisPenceresiAc(int adet)
        {
            Form frmImei = new Form { Text = "📱  Çoklu Satış IMEI Seçimi", Size = new Size(380, 450), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MinimizeBox = false, MaximizeBox = false };

            Panel pnlScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(15) };
            Button btnOnay = new Button { Text = "✔  IMEI NUMARALARINI ONAYLA", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnOnay.FlatAppearance.BorderSize = 0;

            Label lblBilgi = new Label { Text = $"Satılacak {adet} cihaz için IMEI giriniz:", Dock = DockStyle.Top, Height = 30, ForeColor = Color.Gold, Font = new Font("Segoe UI", 10F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
            pnlScroll.Controls.Add(lblBilgi);

            List<TextBox> dinamikKutular = new List<TextBox>();
            int yPos = 40;
            for (int i = 0; i < adet; i++)
            {
                Label l = new Label { Text = $"{i + 1}. Cihaz IMEI:", Location = new Point(20, yPos), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
                TextBox t = new TextBox { Location = new Point(20, yPos + 20), Width = 300, Font = new Font("Segoe UI", 11F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

                if (i < aktifImeiListesi.Count) t.Text = aktifImeiListesi[i];

                pnlScroll.Controls.Add(l); pnlScroll.Controls.Add(t);
                dinamikKutular.Add(t);
                yPos += 65;
            }

            btnOnay.Click += (sX, eX) => {
                aktifImeiListesi.Clear();
                foreach (var kutu in dinamikKutular)
                {
                    if (!string.IsNullOrWhiteSpace(kutu.Text)) aktifImeiListesi.Add(kutu.Text.Trim());
                }
                if (aktifImeiListesi.Count < adet) { MessageBox.Show("Lütfen tüm cihazların IMEI'sini eksiksiz girin!", "Eksik Veri", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                dynTxtImei.Text = $"[{aktifImeiListesi.Count} ADET IMEI GİRİLDİ]";
                frmImei.Close();
            };

            frmImei.Controls.Add(pnlScroll);
            frmImei.Controls.Add(btnOnay);
            frmImei.ShowDialog();
        }

        private void DynCmbPlatform_SelectedIndexChanged(object sender, EventArgs e)
        {
            dbHizmetBedeli = 0; dbPosOrani = 0; dbPosModu = "ANA_TUTAR";
            dbKomisyonMuaf = false; dbStopajMuaf = false;
            // Platform değiştikçe kargo firmalarını dinamik filtrele
            KargoFirmalariniPlatformaGoreGetir(dynCmbPlatform.Text);

            if (dynCmbPlatform.Text.ToUpper() == "EFETÜRK")
            {
                EFETurkSiparisNoUret();
            }
            else
            {
                dynTxtSiparisNo.ReadOnly = false;
                dynTxtSiparisNo.BackColor = Color.FromArgb(41, 41, 44);
                if (dynTxtSiparisNo.Text.StartsWith("2026")) dynTxtSiparisNo.Clear();
            }

            var lblTitle = this.Controls.Find("lblRightTitle", true).FirstOrDefault() as Label;
            if (lblTitle != null)
            {
                lblTitle.Text = string.IsNullOrWhiteSpace(dynCmbPlatform.Text) ? "SİPARİŞ SEPETİ VE FİNANSAL ÖZET" : $"🚀  {dynCmbPlatform.Text.ToUpper()} PLATFORMU SİPARİŞ SEPETİ";
            }

            if (dynCmbPlatform.SelectedIndex <= 0) { CanliFinansalAnalizYap(); return; }

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT HizmetBedeli, PosOrani, PosModu, KomisyonMuaf, StopajMuaf FROM Platformlar WHERE Ad = @ad", conn))
                    {
                        cmd.Parameters.AddWithValue("@ad", dynCmbPlatform.Text);
                        using (var oku = cmd.ExecuteReader())
                        {
                            if (oku.Read())
                            {
                                dbHizmetBedeli = Convert.ToDouble(oku["HizmetBedeli"]);
                                dbPosOrani = Convert.ToDouble(oku["PosOrani"]);
                                dbPosModu = oku["PosModu"]?.ToString() ?? "ANA_TUTAR";
                                dbKomisyonMuaf = Convert.ToInt32(oku["KomisyonMuaf"]) == 1;
                                dbStopajMuaf = Convert.ToInt32(oku["StopajMuaf"]) == 1;
                            }
                        }
                    }
                }
            }
            catch { }

            LoadPlatformKesintileri(dynCmbPlatform.Text);
            CanliFinansalAnalizYap();
        }

        private void LoadPlatformKesintileri(string platform)
        {
            platformKesintiler.Clear();
            DataTable dt = new DataTable();
            dt.Columns.Add("Kaynak"); dt.Columns.Add("Ad"); dt.Columns.Add("Tip"); dt.Columns.Add("Deger", typeof(double));
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT PosOrani, HizmetBedeli, PosModu FROM Platformlar WHERE Ad=@p", conn))
                    {
                        cmd.Parameters.AddWithValue("@p", platform);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                double pos = Convert.ToDouble(r["PosOrani"]);
                                double hizm = Convert.ToDouble(r["HizmetBedeli"]);
                                string mod = r["PosModu"]?.ToString() ?? "ANA_TUTAR";
                                string modStr = mod == "INDIRIMLI" ? " (İndirimli Tutar)" : " (Ana Tutar)";

                                if (hizm > 0) { platformKesintiler.Add(("SABIT_HIZM", "Sabit Hizmet Bedeli", "TL", hizm)); dt.Rows.Add("SABIT_HIZM", "Sabit Hizmet Bedeli", "TL", hizm); }
                                if (pos > 0) { platformKesintiler.Add(("SABIT_POS", "POS Oranı" + modStr, "YÜZDE", pos)); dt.Rows.Add("SABIT_POS", "POS Oranı" + modStr, "YÜZDE", pos); }
                            }
                        }
                    }
                    using (var cmd = new SqliteCommand("SELECT KategoriAd, Tip, Deger FROM PlatformKategoriKesintileri WHERE PlatformAd=@p", conn))
                    {
                        cmd.Parameters.AddWithValue("@p", platform);
                        using (var r = cmd.ExecuteReader()) { while (r.Read()) { string kat = r["KategoriAd"].ToString(); string tip = r["Tip"].ToString(); double val = Convert.ToDouble(r["Deger"]); platformKesintiler.Add(("KAT", kat, tip, val)); dt.Rows.Add("KAT", kat, tip, val); } }
                    }
                    using (var cmd = new SqliteCommand("SELECT UcretAdi, Tip, Deger FROM PlatformEkUcretler WHERE PlatformAd=@p", conn))
                    {
                        cmd.Parameters.AddWithValue("@p", platform);
                        using (var r = cmd.ExecuteReader()) { while (r.Read()) { string ad = r["UcretAdi"].ToString(); string tip = r["Tip"].ToString(); double val = Convert.ToDouble(r["Deger"]); platformKesintiler.Add(("GEN", ad, tip, val)); dt.Rows.Add("GEN", ad, tip, val); } }
                    }
                }
            }
            catch { }

            dynDgvPlatformKesintiler.DataSource = dt;
            if (dynDgvPlatformKesintiler.Columns.Contains("Deger")) dynDgvPlatformKesintiler.Columns["Deger"].DefaultCellStyle.Format = "N2";
        }

        private void EFETurkSiparisNoUret()
        {
            dynTxtSiparisNo.ReadOnly = true; dynTxtSiparisNo.BackColor = Color.FromArgb(21, 33, 21);
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT MAX(CAST(SiparisNo AS INTEGER)) FROM Siparisler WHERE PlatformAd = 'EFETÜRK'", conn))
                    {
                        var res = cmd.ExecuteScalar();
                        if (res != DBNull.Value && res != null && long.TryParse(res.ToString(), out long maxNo)) dynTxtSiparisNo.Text = (maxNo + 1).ToString();
                        else dynTxtSiparisNo.Text = "20261000";
                    }
                }
            }
            catch { dynTxtSiparisNo.Text = "20261000"; }
        }

        private bool SiparisNoVeritabanindaVarMi(string sipNo)
        {
            if (string.IsNullOrEmpty(sipNo)) return false;
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT COUNT(*) FROM Siparisler WHERE SiparisNo = @sn", conn))
                    {
                        cmd.Parameters.AddWithValue("@sn", sipNo);
                        long count = Convert.ToInt64(cmd.ExecuteScalar());
                        return count > 0;
                    }
                }
            }
            catch { return false; }
        }

        private void UrunuSatisIcinCagir()
        {
            string bkd = dynTxtBarkod.Text.Trim(); if (string.IsNullOrEmpty(bkd)) return;
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT u.UrunAdi, u.Kategori, IFNULL(u.Marka, '-') as Marka, IFNULL(s.SatisFiyati, 0) as SatisFiyati, 
                               IFNULL(s.AlisFiyati, 0) as AlisFiyati, IFNULL(s.Adet, 0) as Adet 
                        FROM Urunler u 
                        LEFT JOIN Stoklar s ON u.Barkod = s.Barkod 
                        WHERE u.Barkod = @b";

                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@b", bkd);
                        using (var oku = cmd.ExecuteReader())
                        {
                            if (oku.Read())
                            {
                                dynLblUrunAdi.Text = "Ürün: " + oku["UrunAdi"].ToString();
                                aktifKategori = oku["Kategori"].ToString().ToUpper();
                                dynLblKategori.Text = "Kategori: " + aktifKategori;
                                dynLblMarka.Text = "Marka: " + oku["Marka"].ToString();

                                int.TryParse(oku["Adet"].ToString(), out seciliUrunMevcutStok);
                                dynLblMevcutStok.Text = $"Depo Stoğu: {seciliUrunMevcutStok} Adet";
                                dynLblMevcutStok.ForeColor = seciliUrunMevcutStok > 0 ? Color.FromArgb(56, 189, 248) : Color.Crimson;

                                seciliUrunAlisFiyati = oku["AlisFiyati"] == DBNull.Value ? 0 : Convert.ToDouble(oku["AlisFiyati"]);
                                double sf = oku["SatisFiyati"] == DBNull.Value ? 0 : Convert.ToDouble(oku["SatisFiyati"]);
                                dynTxtSatisFiyati.Text = sf.ToString("N2", new CultureInfo("tr-TR"));

                                if (aktifKategori.Contains("TELEFON") || aktifKategori.Contains("CEP"))
                                {
                                    dynTxtImei.Enabled = true; dynTxtImei.BackColor = Color.FromArgb(41, 24, 24);
                                    dynTxtAdet.Text = "1"; dynTxtImei.Focus();
                                }
                                else
                                {
                                    dynTxtImei.Clear(); dynTxtImei.Enabled = false; dynTxtImei.BackColor = Color.FromArgb(41, 41, 44);
                                    dynTxtAdet.Enabled = true; dynTxtAdet.Focus();
                                }
                            }
                            else
                            {
                                dynLblUrunAdi.Text = "⚠️  BARKOD BULUNAMADI!";
                                dynLblKategori.Text = "Kategori: -";
                                dynLblMarka.Text = "Marka: -";
                                dynLblMevcutStok.Text = "Depo Stoğu: -";
                                dynTxtSatisFiyati.Text = "0,00";
                                seciliUrunMevcutStok = 0;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void DynBtnSepeteEkle_Click(object sender, EventArgs e)
        {
            string barkod = dynTxtBarkod.Text.Trim(); string sipNo = dynTxtSiparisNo.Text.Trim();
            if (string.IsNullOrEmpty(barkod) || string.IsNullOrEmpty(sipNo) || aktifKategori == "") { MessageBox.Show("Geçerli bir barkod okutun veya sipariş no girin!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            if (SiparisNoVeritabanindaVarMi(sipNo))
            {
                MessageBox.Show($"⚠️  MÜKERRER SİPARİŞ HATASI! \n\n[{sipNo}] numaralı sipariş kodu daha önce sisteme kaydedilmiş patron! Her sipariş numarası benzersiz olmalıdır.", "Sistem Güvenlik Duvarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double.TryParse(dynTxtSatisFiyati.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double satisFiyati);

            if (satisFiyati <= 0 || dynCmbKargoFirmasi.SelectedIndex <= 0) { MessageBox.Show("Zorunlu alanları (Satış Fiyatı ve Kargo Firması) kontrol edin!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            int adet = 0; int.TryParse(dynTxtAdet.Text, out adet); if (adet <= 0) adet = 1;

            int sepettekiAyniUrunAdedi = 0;
            foreach (DataRow r in sepetTablosu.Rows)
            {
                if (r["Barkod"].ToString() == barkod) sepettekiAyniUrunAdedi += Convert.ToInt32(r["Adet"]);
            }

            if ((adet + sepettekiAyniUrunAdedi) > seciliUrunMevcutStok)
            {
                MessageBox.Show($"Stok Yetersiz! \n\nDepoda bulunan: {seciliUrunMevcutStok} Adet\nSepete Eklenmiş Olan: {sepettekiAyniUrunAdedi} Adet\nSatılmak İstenen: {adet} Adet", "Stok Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double.TryParse(dynTxtPlatformIndirimi.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double platIndirimi);
            double.TryParse(dynTxtKargoUcreti.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double kargoUcreti);

            string imeiSistemi = "";
            if (dynTxtImei.Enabled)
            {
                if (adet == 1)
                {
                    if (string.IsNullOrWhiteSpace(dynTxtImei.Text) || dynTxtImei.Text.Contains("tıklayın"))
                    {
                        MessageBox.Show("Telefon satışlarında IMEI girmek zorunludur!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                    }
                    imeiSistemi = dynTxtImei.Text.Trim();
                }
                else
                {
                    if (aktifImeiListesi.Count != adet)
                    {
                        MessageBox.Show($"Eksik IMEI girişi! Lütfen kutuya tıklayıp {adet} adet cihazın IMEI'sini eksiksiz girin.", "IMEI Adet Uyuşmazlığı", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
                    }
                    imeiSistemi = string.Join(", ", aktifImeiListesi);
                }
            }

            // 1. İşlem: POS ve Komisyon için platform ayarlarına göre taban (matrah) belirliyoruz
            double islemTabani = dbPosModu == "INDIRIMLI" ? (satisFiyati - platIndirimi) : satisFiyati;
            if (islemTabani < 0) islemTabani = 0;

            // 2. İşlem: Stopaj HER DURUMDA doğrudan ana Satış Fiyatı üzerinden (KDV'siz) hesaplanır
            double kdvsizAnaSatis = satisFiyati / 1.20;
            double stopaj = dbStopajMuaf ? 0 : (kdvsizAnaSatis * 0.01);

            // 3. İşlem: POS Bedeli platform ayarlarına göre belirlenen tabandan hesaplanır
            double posBedeli = islemTabani * (dbPosOrani / 100.0);

            double ekstraTutarTL = 0; double ekstraYuzde = 0;
            foreach (var k in platformKesintiler)
            {
                if (k.Kaynak == "KAT")
                {
                    if (!string.IsNullOrEmpty(aktifKategori) && (k.Ad.ToUpper() == "TÜM KATEGORİLER" || aktifKategori.Contains(k.Ad.ToUpper())))
                    {
                        if (k.Tip == "TL") ekstraTutarTL += k.Deger; else ekstraYuzde += k.Deger;
                    }
                }
                else
                {
                    if (k.Kaynak == "GEN")
                    {
                        if (k.Tip == "TL") ekstraTutarTL += k.Deger; else ekstraYuzde += k.Deger;
                    }
                }
            }

            double ekstraYuzdeTutar = (satisFiyati * ekstraYuzde / 100.0);
            double toplamEkstra = ekstraTutarTL + ekstraYuzdeTutar;

            double geciciKomisyon = 0;

            // Transfer Tutarı: Satıştan sadece kesintilerin çıkmış hali (Platformun sana ödeyeceği para)
            double transferTutari = ((satisFiyati) - (kargoUcreti + stopaj + posBedeli + dbHizmetBedeli + geciciKomisyon + toplamEkstra)) * adet;
            double toplamMaliyet = seciliUrunAlisFiyati * adet;
            double netKar = transferTutari - toplamMaliyet; // Gerçek cebine kalan kâr

            sepetTablosu.Rows.Add(barkod, sipNo, dynCmbPlatform.Text, dynLblUrunAdi.Text.Replace("Ürün: ", ""), adet, satisFiyati, platIndirimi, stopaj * adet, posBedeli * adet, dbHizmetBedeli * adet, geciciKomisyon, kargoUcreti, transferTutari, toplamMaliyet, imeiSistemi, netKar, dynCmbKargoFirmasi.Text, dynCmbDesi.Text, seciliUrunAlisFiyati, aktifKategori);

            dynTxtBarkod.Clear(); dynTxtImei.Clear(); dynTxtAdet.Text = "1"; dynTxtSatisFiyati.Text = "0,00"; dynTxtPlatformIndirimi.Text = "0,00"; dynTxtKargoUcreti.Text = "0,00";
            dynCmbKargoFirmasi.SelectedIndex = 0;
            dynLblUrunAdi.Text = "Ürün bekleniyor..."; dynLblKategori.Text = "Kategori: -"; dynLblMarka.Text = "Marka: -"; dynLblMevcutStok.Text = "Depo Stoğu: -";
            seciliUrunAlisFiyati = 0; seciliUrunMevcutStok = 0; aktifKategori = "";
            aktifImeiListesi.Clear();

            CanliFinansalAnalizYap();
            SepetToplaminiHesapla();
            dynTxtBarkod.Focus();

            if (dynCmbPlatform.Text.ToUpper() == "EFETÜRK" && long.TryParse(sipNo, out long lNo)) dynTxtSiparisNo.Text = (lNo + 1).ToString();
        }

        private void SepetToplaminiHesapla()
        {
            double toplamSatis = 0; double toplamKar = 0; double toplamStopaj = 0; double toplamPos = 0; double toplamHizmet = 0;
            foreach (DataRow r in sepetTablosu.Rows)
            {
                toplamSatis += Convert.ToDouble(r["Satış (TL)"]) * Convert.ToInt32(r["Adet"]);
                toplamKar += Convert.ToDouble(r["Net Kâr (TL)"]);
                toplamStopaj += Convert.ToDouble(r["Stopaj (TL)"]);
                toplamPos += Convert.ToDouble(r["POS (TL)"]);
                toplamHizmet += Convert.ToDouble(r["Hizmet (TL)"]);
            }

            dynLblSepetToplam.Text = toplamSatis.ToString("N2") + " TL / Kâr: " + toplamKar.ToString("N2") + " TL";

            dynLblCanliStopaj.Text = (toplamStopaj).ToString("N2") + " TL" + (dbStopajMuaf ? " (Muaf)" : "");
            dynLblCanliPos.Text = toplamPos.ToString("N2") + " TL";
            dynLblCanliHizmet.Text = toplamHizmet.ToString("N2") + " TL";
            dynLblCanliNetKar.Text = toplamKar.ToString("N2") + " TL";
            dynLblCanliNetKar.ForeColor = toplamKar < 0 ? Color.Crimson : Color.FromArgb(46, 117, 89);

            if (dynDgvSepet.DataSource != null)
            {
                dynDgvSepet.EnableHeadersVisualStyles = false;
                dynDgvSepet.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
                dynDgvSepet.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dynDgvSepet.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                foreach (DataGridViewColumn c in dynDgvSepet.Columns)
                {
                    if (c.ValueType == typeof(double) || c.ValueType == typeof(decimal)) c.DefaultCellStyle.Format = "N2";
                    if (c.Name == "Adet") c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
        }

        private void SepetSil_Click(object sender, EventArgs e)
        {
            if (dynDgvSepet.SelectedRows.Count > 0)
            {
                dynDgvSepet.Rows.RemoveAt(dynDgvSepet.SelectedRows[0].Index);
                SepetToplaminiHesapla();
            }
        }

        private void SepetDuzenle_Click(object sender, EventArgs e)
        {
            if (dynDgvSepet.SelectedRows.Count == 0) return;
            DataGridViewRow row = dynDgvSepet.SelectedRows[0];

            string barkod = row.Cells["Barkod"].Value.ToString();
            int currentAdet = Convert.ToInt32(row.Cells["Adet"].Value);
            double currentSatis = Convert.ToDouble(row.Cells["Satış (TL)"].Value);
            double currentIndirim = Convert.ToDouble(row.Cells["İndirim (TL)"].Value);
            string currentKargoFirma = row.Cells["Kargo Firması"].Value.ToString();
            string currentDesi = row.Cells["Desi"].Value.ToString();
            string currentImei = row.Cells["IMEI"].Value.ToString();
            string kat = row.Cells["Kategori"].Value.ToString();

            int gercekStokDB = 0;
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Adet FROM Stoklar WHERE Barkod=@b", conn))
                    {
                        cmd.Parameters.AddWithValue("@b", barkod);
                        gercekStokDB = Convert.ToInt32(cmd.ExecuteScalar());
                    }
                }
            }
            catch { }

            int baskaSepettekiler = 0;
            foreach (DataGridViewRow r in dynDgvSepet.Rows)
            {
                if (r.Index != row.Index && r.Cells["Barkod"].Value.ToString() == barkod)
                {
                    baskaSepettekiler += Convert.ToInt32(r.Cells["Adet"].Value);
                }
            }
            int izinVerilenMaksAdet = gercekStokDB - baskaSepettekiler;

            Form frm = new Form { Text = "✏️  Sepet Ürününü Düzenle", Size = new Size(340, 360), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MaximizeBox = false, MinimizeBox = false };

            TextBox tAdet, tSatis, tIndirim, tImei;
            ComboBox cKargo, cDesi;

            frm.Controls.Add(FormElemaniYarat("Yeni Adet:", out tAdet, new Point(15, 15), 290, currentAdet.ToString()));
            frm.Controls.Add(FormElemaniYarat("Satış Fiyatı (TL):", out tSatis, new Point(15, 75), 140, currentSatis.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));
            frm.Controls.Add(FormElemaniYarat("İndirim (TL):", out tIndirim, new Point(165, 75), 140, currentIndirim.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));

            frm.Controls.Add(FormComboYarat("Kargo Firması:", out cKargo, new Point(15, 135), 180));
            cKargo.Items.AddRange(new object[] { "", "YURTİÇİ", "ARAS", "MNG", "SÜRAT", "PTT", "HEPSİJET", "TRENDYOL EXPRESS" });
            if (cKargo.Items.Contains(currentKargoFirma)) cKargo.SelectedItem = currentKargoFirma; else cKargo.SelectedIndex = 0;

            frm.Controls.Add(FormComboYarat("Desi:", out cDesi, new Point(205, 135), 100));
            for (int i = 1; i <= 30; i++) cDesi.Items.Add(i.ToString());
            if (cDesi.Items.Contains(currentDesi)) cDesi.SelectedItem = currentDesi; else cDesi.SelectedIndex = 0;

            frm.Controls.Add(FormElemaniYarat("IMEI Numarası (Varsa):", out tImei, new Point(15, 195), 290, currentImei));
            if (!kat.Contains("TELEFON") && !kat.Contains("CEP")) { tImei.Enabled = false; tImei.BackColor = Color.FromArgb(41, 41, 44); } else { tImei.BackColor = Color.FromArgb(41, 24, 24); }

            Button btnGuncelle = new Button { Text = "💾  GÜNCELLE", Location = new Point(15, 265), Width = 290, Height = 40, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnGuncelle.FlatAppearance.BorderSize = 0;
            btnGuncelle.Click += (sX, eX) => {

                int nAdet = 0; int.TryParse(tAdet.Text, out nAdet);
                if (nAdet <= 0) { MessageBox.Show("Adet en az 1 olmalıdır!", "Hata"); return; }
                if (nAdet > izinVerilenMaksAdet) { MessageBox.Show($"Stok Yetersiz! Ekleyebileceğiniz maksimum adet: {izinVerilenMaksAdet}", "Hata"); return; }

                double.TryParse(tSatis.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nSatis);
                double.TryParse(tIndirim.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nIndirim);
                string nImei = tImei.Text.Trim();

                if (nSatis <= 0 || cKargo.SelectedIndex <= 0) { MessageBox.Show("Satış Fiyatı ve Kargo Firması zorunludur!"); return; }
                if (tImei.Enabled && string.IsNullOrEmpty(nImei)) { MessageBox.Show("IMEI Numarası girmek zorunludur!"); return; }

                double nKargoUcreti = 0;
                if (cKargo.SelectedIndex > 0)
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var cmd = new SqliteCommand("SELECT Fiyat FROM KargoMatrisi WHERE UPPER(KargoFirma) LIKE '%' || @k || '%' AND Desi <= @d ORDER BY Desi DESC LIMIT 1", b))
                        {
                            cmd.Parameters.AddWithValue("@k", cKargo.Text.ToUpper());
                            cmd.Parameters.AddWithValue("@d", cDesi.Text);
                            var res = cmd.ExecuteScalar();
                            if (res != null) nKargoUcreti = Convert.ToDouble(res);
                        }
                    }
                }

                double alisFiyat = Convert.ToDouble(row.Cells["Alış Fiyatı"].Value);

                // POS için taban
                double islemTabani = dbPosModu == "INDIRIMLI" ? (nSatis - nIndirim) : nSatis;
                if (islemTabani < 0) islemTabani = 0;

                // Stopaj HER DURUMDA doğrudan nSatis üzerinden (KDV'siz) hesaplanır
                double kdvsizAnaSatis = nSatis / 1.20;
                double stopaj = dbStopajMuaf ? 0 : (kdvsizAnaSatis * 0.01);

                // POS Bedeli
                double posBedeli = islemTabani * (dbPosOrani / 100.0);

                double ekstraTutarTL = 0; double ekstraYuzde = 0;
                foreach (var k in platformKesintiler)
                {
                    if (k.Kaynak == "KAT")
                    {
                        if (!string.IsNullOrEmpty(kat) && (k.Ad.ToUpper() == "TÜM KATEGORİLER" || kat.Contains(k.Ad.ToUpper())))
                        {
                            if (k.Tip == "TL") ekstraTutarTL += k.Deger; else ekstraYuzde += k.Deger;
                        }
                    }
                    else if (k.Kaynak == "GEN")
                    {
                        if (k.Tip == "TL") ekstraTutarTL += k.Deger; else ekstraYuzde += k.Deger;
                    }
                }
                double ekstraYuzdeTutar = (nSatis * ekstraYuzde / 100.0);
                double toplamEkstra = ekstraTutarTL + ekstraYuzdeTutar;

                double nKomisyon = 0;

                double transferTutari = ((nSatis) - (nKargoUcreti + stopaj + posBedeli + dbHizmetBedeli + nKomisyon + toplamEkstra)) * nAdet;
                double toplamMaliyet = alisFiyat * nAdet;
                double netKar = transferTutari - toplamMaliyet;

                row.Cells["Adet"].Value = nAdet;
                row.Cells["Satış (TL)"].Value = nSatis;
                row.Cells["İndirim (TL)"].Value = nIndirim;
                row.Cells["Stopaj (TL)"].Value = stopaj * nAdet;
                row.Cells["POS (TL)"].Value = posBedeli * nAdet;
                row.Cells["Hizmet (TL)"].Value = dbHizmetBedeli * nAdet;
                row.Cells["Komisyon (TL)"].Value = nKomisyon * nAdet;
                row.Cells["Kargo (TL)"].Value = nKargoUcreti;
                row.Cells["Transfer Tutarı (TL)"].Value = transferTutari; // EKLENDİ
                row.Cells["Toplam Maliyet (TL)"].Value = toplamMaliyet;   // EKLENDİ
                row.Cells["Kargo Firması"].Value = cKargo.Text;
                row.Cells["Desi"].Value = cDesi.Text;
                row.Cells["IMEI"].Value = nImei;
                row.Cells["Net Kâr (TL)"].Value = netKar;

                SepetToplaminiHesapla();
                frm.Close();
            };
            frm.Controls.Add(btnGuncelle);
            frm.ShowDialog();
        }

        private void DynBtnSatisiTamamla_Click(object sender, EventArgs e)
        {
            if (sepetTablosu.Rows.Count == 0) return;

            Form frmKomisyon = new Form { Text = "💰  Sipariş Komisyonu", Size = new Size(350, 210), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MaximizeBox = false, MinimizeBox = false, FormBorderStyle = FormBorderStyle.FixedDialog };

            Label lblK = new Label { Text = "Bu siparişten kesilen TOPLAM Komisyon (TL):", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Orange };
            TextBox txtKomisyon = new TextBox { Location = new Point(20, 45), Width = 290, Font = new Font("Segoe UI", 12F, FontStyle.Bold), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Text = "0,00" };

            txtKomisyon.Leave += (sX, eX) => {
                string m = txtKomisyon.Text.Trim().Replace("₺", "");
                if (double.TryParse(m.Replace(".", "").Replace(",", "."), out double d))
                    txtKomisyon.Text = d.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));
            };

            Button btnOnay = new Button { Text = "SİPARİŞİ MÜHÜRLE", Location = new Point(20, 100), Width = 290, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnOnay.FlatAppearance.BorderSize = 0;

            btnOnay.Click += (sX, eX) => {
                double.TryParse(txtKomisyon.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double toplamKomisyon);

                if (!dbKomisyonMuaf && toplamKomisyon <= 0)
                {
                    MessageBox.Show("Bu platform komisyondan muaf değil! Lütfen sıfırdan büyük bir komisyon tutarı girin.", "Eksik Veri", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                double totalSales = 0;
                foreach (DataRow r in sepetTablosu.Rows)
                {
                    totalSales += Convert.ToDouble(r["Satış (TL)"]) * Convert.ToInt32(r["Adet"]);
                }

                foreach (DataRow r in sepetTablosu.Rows)
                {
                    double rowSales = Convert.ToDouble(r["Satış (TL)"]) * Convert.ToInt32(r["Adet"]);
                    double rowKomisyon = totalSales > 0 ? (rowSales / totalSales) * toplamKomisyon : (toplamKomisyon / sepetTablosu.Rows.Count);

                    r["Komisyon (TL)"] = rowKomisyon;
                    // Komisyon hem Transfer Tutarından hem de Net Kar'dan düşer
                    r["Transfer Tutarı (TL)"] = Convert.ToDouble(r["Transfer Tutarı (TL)"]) - rowKomisyon;
                    r["Net Kâr (TL)"] = Convert.ToDouble(r["Net Kâr (TL)"]) - rowKomisyon;
                }

                frmKomisyon.DialogResult = DialogResult.OK;
                frmKomisyon.Close();
            };

            frmKomisyon.Controls.Add(lblK); frmKomisyon.Controls.Add(txtKomisyon); frmKomisyon.Controls.Add(btnOnay);

            if (frmKomisyon.ShowDialog() != DialogResult.OK) return;

            try
            {
                SayfaStokHareketleri.TabloyuGuvenceyeAlStatik();

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var tx = conn.BeginTransaction())
                    {
                        foreach (DataRow r in sepetTablosu.Rows)
                        {
                            string sql = @"INSERT INTO Siparisler (SiparisNo, PlatformAd, Barkod, Adet, SatisFiyati, PlatformIndirimi, Stopaj, PosBedeli, HizmetBedeli, KomisyonTL, KargoUcreti, Desi, KargoFirmasi, SatilanImei, Durum, Tarih, AlisFiyati) 
                                           VALUES (@sn, @p, @b, @adet, @sf, @pi, @st, @pb, @hb, @km, @kr, @d, @kf, @im, 'OLUŞTURULDU', @t, @af)";
                            using (var cmd = new SqliteCommand(sql, conn, tx))
                            {
                                cmd.Parameters.AddWithValue("@sn", r["Sipariş No"]); cmd.Parameters.AddWithValue("@p", r["Platform"]);
                                cmd.Parameters.AddWithValue("@b", r["Barkod"]); cmd.Parameters.AddWithValue("@adet", r["Adet"]);
                                cmd.Parameters.AddWithValue("@sf", r["Satış (TL)"]); cmd.Parameters.AddWithValue("@pi", r["İndirim (TL)"]);
                                cmd.Parameters.AddWithValue("@st", r["Stopaj (TL)"]); cmd.Parameters.AddWithValue("@pb", r["POS (TL)"]);
                                cmd.Parameters.AddWithValue("@hb", r["Hizmet (TL)"]); cmd.Parameters.AddWithValue("@km", r["Komisyon (TL)"]);
                                cmd.Parameters.AddWithValue("@kr", r["Kargo (TL)"]);
                                cmd.Parameters.AddWithValue("@d", r["Desi"]);
                                cmd.Parameters.AddWithValue("@kf", r["Kargo Firması"]);
                                cmd.Parameters.AddWithValue("@im", r["IMEI"]);
                                cmd.Parameters.AddWithValue("@t", dynDtSiparisTarihi.Value.ToString("dd.MM.yyyy"));
                                cmd.Parameters.AddWithValue("@af", r["Alış Fiyatı"]); // EKLENDİ! Artık maliyet veritabanına yazılıyor.
                                cmd.ExecuteNonQuery();
                            }

                            using (var cmd2 = new SqliteCommand("UPDATE Stoklar SET Adet = MAX(0, Adet - @a), StokMiktari = MAX(0, StokMiktari - @a) WHERE Barkod=@b", conn, tx))
                            {
                                cmd2.Parameters.AddWithValue("@a", r["Adet"]); cmd2.Parameters.AddWithValue("@b", r["Barkod"]); cmd2.ExecuteNonQuery();
                            }

                            string logSql = @"INSERT INTO StokHareketleri (Tarih, Barkod, IslemTuru, Miktar, EskiStok, YeniStok, Aciklama) 
                                              SELECT @tarih, @b, 'SATIŞ', @a, StokMiktari + @a, StokMiktari, @ack 
                                              FROM Stoklar WHERE Barkod=@b";
                            using (var cmdLog = new SqliteCommand(logSql, conn, tx))
                            {
                                cmdLog.Parameters.AddWithValue("@tarih", DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"));
                                cmdLog.Parameters.AddWithValue("@b", r["Barkod"]);
                                cmdLog.Parameters.AddWithValue("@a", r["Adet"]);
                                cmdLog.Parameters.AddWithValue("@ack", "Sipariş No: " + r["Sipariş No"] + " | Platform: " + r["Platform"]);
                                cmdLog.ExecuteNonQuery();
                            }

                            string satilanImeiler = r["IMEI"].ToString();
                            if (!string.IsNullOrEmpty(satilanImeiler) && satilanImeiler != "-")
                            {
                                var imeiDizi = satilanImeiler.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                foreach (var im in imeiDizi)
                                {
                                    using (var cmd3 = new SqliteCommand("UPDATE Imeiler SET Durum='SATILDI' WHERE ImeiNo=@i AND Barkod=@b", conn, tx))
                                    {
                                        cmd3.Parameters.AddWithValue("@i", im);
                                        cmd3.Parameters.AddWithValue("@b", r["Barkod"]);
                                        cmd3.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                        tx.Commit(); MessageBox.Show("Sipariş Başarıyla Kaydedildi!", "Kayıt Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        SiparisEkraniniTemizle(); // TÜM EKRAN BURADA SIFIRLANIYOR
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Kayıt Hatası: " + ex.Message); }
        }
        private void KargoUcretiniGetir() { try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT Fiyat FROM KargoMatrisi WHERE UPPER(KargoFirma) LIKE '%' || @k || '%' AND PlatformAd = @p AND Desi <= @d ORDER BY Desi DESC LIMIT 1", conn)) { cmd.Parameters.AddWithValue("@k", dynCmbKargoFirmasi.Text.ToUpper()); cmd.Parameters.AddWithValue("@p", dynCmbPlatform.Text); cmd.Parameters.AddWithValue("@d", dynCmbDesi.Text); var res = cmd.ExecuteScalar(); dynTxtKargoUcreti.Text = res != null ? Convert.ToDouble(res).ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) : "0,00"; } } } catch { } }
        private void SiparisEkraniniTemizle()
        {
            sepetTablosu.Clear();
            SepetToplaminiHesapla();

            // Platformu ve Sipariş Numarasını Sıfırla
            if (dynCmbPlatform.Items.Count > 0) dynCmbPlatform.SelectedIndex = 0;
            dynTxtSiparisNo.Clear();

            dynTxtBarkod.Clear();
            dynTxtImei.Clear();
            dynTxtAdet.Text = "1";
            dynTxtSatisFiyati.Text = "0,00";
            dynTxtPlatformIndirimi.Text = "0,00";
            dynTxtKargoUcreti.Text = "0,00";

            if (dynCmbKargoFirmasi.Items.Count > 0) dynCmbKargoFirmasi.SelectedIndex = 0;
            if (dynCmbDesi.Items.Count > 0) dynCmbDesi.SelectedIndex = 0;

            dynLblUrunAdi.Text = "Ürün bekleniyor...";
            dynLblKategori.Text = "Kategori: -";
            dynLblMarka.Text = "Marka: -";
            dynLblMevcutStok.Text = "Depo Stoğu: -";

            seciliUrunAlisFiyati = 0;
            seciliUrunMevcutStok = 0;
            aktifKategori = "";
            aktifImeiListesi.Clear();

            CanliFinansalAnalizYap();
            dynTxtBarkod.Focus();
        }
        private void ParaKutusuFormatla(TextBox txt)
        {
            // Kutuya tıklanınca içindeki "0,00" değerini komple seçer, doğrudan rakam yazmaya başlarsın
            txt.GotFocus += (s, e) => { txt.SelectAll(); };
            txt.Click += (s, e) => { txt.SelectAll(); };

            // Klavyedeki numpad noktasına basıldığında otomatik olarak virgüle çevirir (Kuruş ayırıcı)
            txt.KeyPress += (s, e) => {
                if (e.KeyChar == '.') e.KeyChar = ',';
            };

            // Kutudan çıkıldığında girilen değeri profesyonel muhasebe formatına dönüştürür (Örn: 1.500,50)
            txt.Leave += (s, e) => {
                if (string.IsNullOrWhiteSpace(txt.Text))
                {
                    txt.Text = "0,00";
                    return;
                }

                if (double.TryParse(txt.Text, NumberStyles.Any, CultureInfo.GetCultureInfo("tr-TR"), out double val))
                    txt.Text = val.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));
                else
                    txt.Text = "0,00";
            };
        }
        private void KargoFirmalariniPlatformaGoreGetir(string seciliPlatform)
        {
            dynCmbKargoFirmasi.Items.Clear();
            dynCmbKargoFirmasi.Items.Add("");
            dynTxtKargoUcreti.Text = "0,00"; // Platform değişince ücreti de sıfırla

            if (string.IsNullOrWhiteSpace(seciliPlatform))
            {
                dynCmbKargoFirmasi.SelectedIndex = 0;
                return;
            }

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    // Sadece seçili platforma ait kayıtlı kargo firmalarını getirir
                    using (var cmd = new SqliteCommand("SELECT DISTINCT KargoFirma FROM KargoMatrisi WHERE PlatformAd = @p", conn))
                    {
                        cmd.Parameters.AddWithValue("@p", seciliPlatform);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                dynCmbKargoFirmasi.Items.Add(reader["KargoFirma"].ToString());
                            }
                        }
                    }
                }
            }
            catch { }

            dynCmbKargoFirmasi.SelectedIndex = 0;
        }
        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan) { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray }; kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
        private Panel FormComboYarat(string baslik, out ComboBox kutu, Point konum, int genislik) { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray }; kutu = new ComboBox { Location = new Point(0, 18), Width = genislik, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
        private Label GöstergeEtiketiYarat(TableLayoutPanel tlp, string baslik, string deger, int sutun, Color renk) { Panel p = new Panel { Dock = DockStyle.Fill }; Label lblB = new Label { Text = baslik, Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.Gray, TextAlign = ContentAlignment.MiddleCenter }; Label lblD = new Label { Text = deger, Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = renk, TextAlign = ContentAlignment.MiddleCenter }; p.Controls.Add(lblD); p.Controls.Add(lblB); tlp.Controls.Add(p, sutun, 0); return lblD; }

        private void CanliFinansalAnalizYap()
        {
            double satisFiyati = 0; double.TryParse(dynTxtSatisFiyati.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out satisFiyati);
            double platIndirimi = 0; double.TryParse(dynTxtPlatformIndirimi.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out platIndirimi);
            double kargoUcreti = 0; double.TryParse(dynTxtKargoUcreti.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out kargoUcreti);
            int adet = 0; int.TryParse(dynTxtAdet.Text, out adet); if (adet <= 0) adet = 1;

            // POS için taban
            double islemTabani = dbPosModu == "INDIRIMLI" ? (satisFiyati - platIndirimi) : satisFiyati;
            if (islemTabani < 0) islemTabani = 0;

            // Stopaj HER DURUMDA doğrudan ana Satış Fiyatı üzerinden (KDV'siz) hesaplanır
            double kdvsizAnaSatis = satisFiyati / 1.20;
            double stopaj = dbStopajMuaf ? 0 : (kdvsizAnaSatis * 0.01);

            // POS Bedeli
            double posBedeli = islemTabani * (dbPosOrani / 100.0);

            double ekstraTutarTL = 0; double ekstraYuzde = 0;
            foreach (var k in platformKesintiler)
            {
                if (k.Kaynak == "KAT")
                {
                    if (!string.IsNullOrEmpty(aktifKategori) && (k.Ad.ToUpper() == "TÜM KATEGORİLER" || aktifKategori.Contains(k.Ad.ToUpper())))
                    {
                        if (k.Tip == "TL") ekstraTutarTL += k.Deger; else ekstraYuzde += k.Deger;
                    }
                }
                else if (k.Kaynak == "GEN")
                {
                    if (k.Tip == "TL") ekstraTutarTL += k.Deger; else ekstraYuzde += k.Deger;
                }
            }
            double ekstraYuzdeTutar = satisFiyati * ekstraYuzde / 100.0;

            double geciciKomisyon = 0;
            double kesintiler = kargoUcreti + stopaj + posBedeli + dbHizmetBedeli + geciciKomisyon + ekstraTutarTL + ekstraYuzdeTutar;

            double transferTutari = (satisFiyati - kesintiler) * adet; // Platformdan gelecek para
            double kalanKar = transferTutari - (seciliUrunAlisFiyati * adet); // Net kâr

            dynLblCanliStopaj.Text = (stopaj * adet).ToString("N2") + " TL" + (dbStopajMuaf ? " (Muaf)" : "");
            dynLblCanliPos.Text = (posBedeli * adet).ToString("N2") + " TL";
            dynLblCanliHizmet.Text = (dbHizmetBedeli * adet).ToString("N2") + " TL";

            dynLblCanliNetKar.Text = kalanKar.ToString("N2") + " TL";
            dynLblCanliNetKar.ForeColor = kalanKar < 0 ? Color.Crimson : Color.FromArgb(46, 117, 89);
        }
    }
}