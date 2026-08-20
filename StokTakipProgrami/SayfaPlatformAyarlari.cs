using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // ⚙️ PLATFORM FİNANSAL AYARLAR MERKEZİ (KARGO İNDİRİMİ SİSTEMİ ENTEGRE EDİLDİ)
    public partial class SayfaPlatformAyarlari : UserControl
    {
        private ComboBox ucPltCmbPlatformlar;
        private TextBox ucPltTxtYeniPlatform;

        private TableLayoutPanel pnlRightMain;
        private Label ucPltLblPanelBaslik;

        private TextBox ucPltTxtHizmetBedeli;
        private TextBox ucPltTxtPosOrani;
        private ComboBox ucPltCmbPosModu;
        private ComboBox ucPltCmbHedefKat, ucPltCmbKatTip;
        private TextBox ucPltTxtKatDeger;
        private TextBox ucPltTxtGenelAd, ucPltTxtGenelDeger;
        private ComboBox ucPltCmbGenelTip;

        private CheckBox ucPltChkKomisyonMuaf;
        private CheckBox ucPltChkStopajMuaf;

        // 🚀 MÜHÜR: Kargo İndirimi Kontrolleri
        private ComboBox ucPltCmbKargoFirma, ucPltCmbKargoIndTip;
        private TextBox ucPltTxtKargoIndDeger;

        private DataGridView ucPltDgvTumKesintiler;

        public SayfaPlatformAyarlari()
        {
            InitializeComponent();
            ArayuzElemanlariniInsaEt();
        }

        private void SayfaPlatformAyarlari_Load(object sender, EventArgs e)
        {
            TabloAltyapisiniGuvenceyeAl();
            PlatformlariYukle();
            pnlRightMain.Visible = false;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            TabloAltyapisiniGuvenceyeAl();
            PlatformlariYukle();
        }

        private void TabloAltyapisiniGuvenceyeAl()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("CREATE TABLE IF NOT EXISTS Platformlar (Ad TEXT PRIMARY KEY, PosOrani REAL DEFAULT 0, HizmetBedeli REAL DEFAULT 0, PosModu TEXT DEFAULT 'ANA_TUTAR');", b)) { cmd.ExecuteNonQuery(); }

                    var sutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var cmd = new SqliteCommand("PRAGMA table_info('Platformlar');", b))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) sutunlar.Add(r["name"].ToString()); }

                    if (!sutunlar.Contains("PosModu")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN PosModu TEXT DEFAULT 'ANA_TUTAR';", b)) cmd.ExecuteNonQuery(); }
                    if (!sutunlar.Contains("KomisyonMuaf")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN KomisyonMuaf INTEGER DEFAULT 0;", b)) cmd.ExecuteNonQuery(); }
                    if (!sutunlar.Contains("StopajMuaf")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN StopajMuaf INTEGER DEFAULT 0;", b)) cmd.ExecuteNonQuery(); }

                    // 🚀 MÜHÜR: Kargo indirimi için gerekli sütunlar eklendi
                    if (!sutunlar.Contains("KargoIndirimTip")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN KargoIndirimTip TEXT DEFAULT 'TL';", b)) cmd.ExecuteNonQuery(); }
                    if (!sutunlar.Contains("KargoIndirimDeger")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN KargoIndirimDeger REAL DEFAULT 0;", b)) cmd.ExecuteNonQuery(); }
                    if (!sutunlar.Contains("KargoIndirimFirma")) { using (var cmd = new SqliteCommand("ALTER TABLE Platformlar ADD COLUMN KargoIndirimFirma TEXT DEFAULT 'TÜMÜ';", b)) cmd.ExecuteNonQuery(); }

                    using (var cmd = new SqliteCommand("CREATE TABLE IF NOT EXISTS PlatformEkUcretler (PlatformAd TEXT, UcretAdi TEXT, Tip TEXT, Deger REAL);", b)) { cmd.ExecuteNonQuery(); }
                    using (var cmd = new SqliteCommand("CREATE TABLE IF NOT EXISTS PlatformKategoriKesintileri (PlatformAd TEXT, KategoriAd TEXT, Tip TEXT, Deger REAL);", b)) { cmd.ExecuteNonQuery(); }
                }
            }
            catch { }
        }

        private void ArayuzElemanlariniInsaEt()
        {
            this.Controls.Clear(); this.Dock = DockStyle.Fill; this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 90, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Name = "lblAnaBaslik", Text = "⚙️ PLATFORM FİNANSAL HESAP VE KESİNTİ YÖNETİM MERKEZİ", Dock = DockStyle.Top, Height = 55, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            Label lblSeciliPlatform = new Label { Name = "lblSeciliPlatform", Text = "", Dock = DockStyle.Top, Height = 32, Font = new Font("Segoe UI", 13F, FontStyle.Bold), ForeColor = Color.Orange, TextAlign = ContentAlignment.MiddleCenter, Visible = false };
            pnlHeader.Controls.Add(lblSeciliPlatform);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = Color.Transparent };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 420F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Margin = new Padding(15) };
            Label lblSecim = new Label { Text = "🔍 PLATFORM SEÇİMİ", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.Orange };

            Label lblPlt = new Label { Text = "Mevcut Platformlar:", Location = new Point(20, 60), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 9F) };
            ucPltCmbPlatformlar = new ComboBox { Location = new Point(20, 82), Width = 350, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 11F, FontStyle.Bold), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            ucPltCmbPlatformlar.SelectedIndexChanged += UcPltCmbPlatformlar_SelectedIndexChanged;

            Panel pnlAyrac = new Panel { Location = new Point(20, 135), Width = 350, Height = 1, BackColor = Color.FromArgb(50, 50, 55) };
            Label lblYeni = new Label { Text = "➕ YENİ PLATFORM YARAT", Location = new Point(20, 160), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) };
            Label lblYeniAd = new Label { Text = "Platform Adı:", Location = new Point(20, 195), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 9F) };
            ucPltTxtYeniPlatform = new TextBox { Location = new Point(20, 215), Width = 350, Font = new Font("Segoe UI", 11F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

            Button btnYeniPlt = new Button { Text = "KAYDET VE LİSTEYE EKLE", Location = new Point(20, 260), Width = 350, Height = 42, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnYeniPlt.FlatAppearance.BorderSize = 0; btnYeniPlt.Click += UcPltBtnYeniPlatformEkle_Click;

            pnlLeft.Controls.Add(lblSecim); pnlLeft.Controls.Add(lblPlt); pnlLeft.Controls.Add(ucPltCmbPlatformlar);
            pnlLeft.Controls.Add(pnlAyrac); pnlLeft.Controls.Add(lblYeni); pnlLeft.Controls.Add(lblYeniAd);
            pnlLeft.Controls.Add(ucPltTxtYeniPlatform); pnlLeft.Controls.Add(btnYeniPlt);

            // Sağ panel yüksekliği 6 satıra çıkarıldığı için RowCount=4 yapıldı
            pnlRightMain = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = new Padding(15) };
            pnlRightMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 0F));
            pnlRightMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 480F)); // Yükseklik artırıldı
            pnlRightMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Panel pnlBaslik = new Panel { Dock = DockStyle.Top, Height = 0 };
            ucPltLblPanelBaslik = new Label { Text = "📝 PLATFORM SEÇİLMEDİ", Location = new Point(0, 5), AutoSize = true, Font = new Font("Segoe UI Black", 16F, FontStyle.Bold), ForeColor = Color.Orange, Visible = false };
            pnlBaslik.Controls.Add(ucPltLblPanelBaslik);

            FlowLayoutPanel flpGirisler = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };

            // 1- Sabit Hizmet
            Panel pS1 = new Panel { Width = 880, Height = 70, BackColor = Color.FromArgb(26, 26, 28), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 8) };
            Label lS1 = new Label { Text = "1- SABİT HİZMET BEDELİ (TL):", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.White };
            pS1.Controls.Add(lS1); pS1.Controls.Add(FormElemaniYarat("Tutar (TL):", out ucPltTxtHizmetBedeli, new Point(10, 25), 180, "0,00"));
            Button bS1 = new Button { Text = "💾 HİZMET BEDELİNİ KAYDET", Location = new Point(640, 22), Width = 220, Height = 32, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            bS1.FlatAppearance.BorderSize = 0; bS1.Click += (s, e) => KaydetHizmetBedeli(); pS1.Controls.Add(bS1);

            // 2- POS Bedeli
            Panel pS2 = new Panel { Width = 880, Height = 70, BackColor = Color.FromArgb(26, 26, 28), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 8) };
            Label lS2 = new Label { Text = "2- GÜVENCE / POS BEDELİ (% Oran):", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.Orange };
            pS2.Controls.Add(lS2); pS2.Controls.Add(FormElemaniYarat("Komisyon Oranı (%):", out ucPltTxtPosOrani, new Point(10, 25), 180, "0,00"));
            pS2.Controls.Add(FormComboYarat("POS Hesaplama Tabanı / Modu:", out ucPltCmbPosModu, new Point(210, 25), 250, new string[] { "1- ANA TUTAR (SATIŞ ÜZERİNDEN)", "2- İNDİRİMLİ TUTAR ÜZERİNDEN" }));
            Button bS2 = new Button { Text = "💾 POS BEDELİNİ KAYDET", Location = new Point(640, 22), Width = 220, Height = 32, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            bS2.FlatAppearance.BorderSize = 0; bS2.Click += (s, e) => KaydetPosBedeli(); pS2.Controls.Add(bS2);

            // 3- Kategori Kesintisi
            Panel pS3 = new Panel { Width = 880, Height = 70, BackColor = Color.FromArgb(26, 26, 28), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 8) };
            Label lS3 = new Label { Text = "3- KATEGORİYE ÖZEL KESİNTİ / HİZMET TANIMLA:", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) };
            pS3.Controls.Add(lS3); pS3.Controls.Add(FormComboYarat("Hedef Kategori:", out ucPltCmbHedefKat, new Point(10, 25), 180, new string[] { "TÜM KATEGORİLER" }));
            pPltCmbHedefKatDoldur(); pS3.Controls.Add(FormComboYarat("Kesinti Tipi:", out ucPltCmbKatTip, new Point(210, 25), 150, new string[] { "YÜZDE (% Oran)", "TL (Sabit Tutar)" }));
            pS3.Controls.Add(FormElemaniYarat("Oran/Değer:", out ucPltTxtKatDeger, new Point(380, 25), 120, "0,00"));
            Button bS3 = new Button { Text = "➕ KATEGORİ KESİNTİSİ EKLE", Location = new Point(640, 22), Width = 220, Height = 32, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            bS3.FlatAppearance.BorderSize = 0; bS3.Click += (s, e) => EkleKategoriKesinti(); pS3.Controls.Add(bS3);

            // 4- Genel Kesinti
            Panel pS4 = new Panel { Width = 880, Height = 70, BackColor = Color.FromArgb(26, 26, 28), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 8) };
            Label lS4 = new Label { Text = "4- GENEL EK KESİNTİ / HİZMET TANIMLA:", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.Gold };
            pS4.Controls.Add(lS4); pS4.Controls.Add(FormElemaniYarat("Kesinti Adı (Örn: Kargo Payı):", out ucPltTxtGenelAd, new Point(10, 25), 180, ""));
            pS4.Controls.Add(FormComboYarat("Kesinti Tipi:", out ucPltCmbGenelTip, new Point(210, 25), 150, new string[] { "TL (Sabit Tutar)", "YÜZDE (% Oran)" }));
            pS4.Controls.Add(FormElemaniYarat("Oran/Değer:", out ucPltTxtGenelDeger, new Point(380, 25), 120, "0,00"));
            Button bS4 = new Button { Text = "➕ GENEL KESİNTİ EKLE", Location = new Point(640, 22), Width = 220, Height = 32, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            bS4.FlatAppearance.BorderSize = 0; bS4.Click += (s, e) => EkleGenelKesinti(); pS4.Controls.Add(bS4);

            // 5- Muafiyetler
            Panel pS5 = new Panel { Width = 880, Height = 70, BackColor = Color.FromArgb(26, 26, 28), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 8) };
            Label lS5 = new Label { Text = "5- PLATFORM MUAFİYET VE ÖZEL DURUMLAR:", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.LightGreen };
            pS5.Controls.Add(lS5);
            ucPltChkKomisyonMuaf = new CheckBox { Text = "Siparişlerde Komisyon Zorunlu OLMASIN", Location = new Point(10, 32), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 9F), Cursor = Cursors.Hand };
            ucPltChkStopajMuaf = new CheckBox { Text = "Bu Platform Stopajdan Muaftır (%1 Kesilmez)", Location = new Point(330, 32), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 9F), Cursor = Cursors.Hand };
            Button bS5 = new Button { Text = "💾 MUAFİYETLERİ KAYDET", Location = new Point(640, 22), Width = 220, Height = 32, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            bS5.FlatAppearance.BorderSize = 0; bS5.Click += (s, e) => KaydetMuafiyet();
            pS5.Controls.Add(ucPltChkKomisyonMuaf); pS5.Controls.Add(ucPltChkStopajMuaf); pS5.Controls.Add(bS5);

            // 🚀 MÜHÜR: 6- Kargo İndirim Matrisi Paneli
            Panel pS6 = new Panel { Width = 880, Height = 70, BackColor = Color.FromArgb(26, 26, 28), BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 8) };
            Label lS6 = new Label { Text = "6- KARGO İNDİRİMİ VE LOJİSTİK ANLAŞMASI:", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.Crimson };
            pS6.Controls.Add(lS6);
            pS6.Controls.Add(FormComboYarat("Kargo Firması:", out ucPltCmbKargoFirma, new Point(10, 25), 180, new string[] { "TÜMÜ", "YURTİÇİ", "ARAS", "MNG", "SÜRAT", "PTT", "HEPSİJET", "TRENDYOL EXPRESS" }));
            pS6.Controls.Add(FormComboYarat("İndirim Tipi:", out ucPltCmbKargoIndTip, new Point(210, 25), 150, new string[] { "YÜZDE (% Oran)", "TL (Sabit Tutar)" }));
            pS6.Controls.Add(FormElemaniYarat("Oran/Değer:", out ucPltTxtKargoIndDeger, new Point(380, 25), 120, "0,00"));
            Button bS6 = new Button { Text = "💾 KARGO İNDİRİMİNİ KAYDET", Location = new Point(640, 22), Width = 220, Height = 32, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            bS6.FlatAppearance.BorderSize = 0; bS6.Click += (s, e) => KaydetKargoIndirimi(); pS6.Controls.Add(bS6);

            flpGirisler.Controls.Add(pS1); flpGirisler.Controls.Add(pS2); flpGirisler.Controls.Add(pS3);
            flpGirisler.Controls.Add(pS4); flpGirisler.Controls.Add(pS5); flpGirisler.Controls.Add(pS6);

            pnlRightMain.Controls.Add(pnlBaslik, 0, 0);
            pnlRightMain.Controls.Add(flpGirisler, 0, 1);

            Panel pnlTabloTasiyici = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 10, 0, 0) };
            Label lblTabloBaslik = new Label { Text = "📋 BU PLATFORMA AİT TANIMLI TÜM KESİNTİLER LİSTESİ", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.White };

            ucPltDgvTumKesintiler = GridYarat(Color.White);
            ucPltDgvTumKesintiler.CellContentClick += UcPltDgvTumKesintiler_CellContentClick;

            pnlTabloTasiyici.Controls.Add(ucPltDgvTumKesintiler);
            pnlTabloTasiyici.Controls.Add(lblTabloBaslik);
            ucPltDgvTumKesintiler.BringToFront();

            pnlRightMain.Controls.Add(pnlTabloTasiyici, 0, 2);

            mainLayout.Controls.Add(pnlLeft, 0, 0);
            mainLayout.Controls.Add(pnlRightMain, 1, 0);

            this.Controls.Add(mainLayout);
            this.Controls.Add(pnlHeader);
            pnlHeader.SendToBack();
            mainLayout.BringToFront();
        }

        // 🚀 MÜHÜR: Kargo İndirimi Kaydetme Motoru
        private void KaydetKargoIndirimi()
        {
            string secili = ucPltCmbPlatformlar.Text; if (string.IsNullOrEmpty(secili)) return;
            double.TryParse(ucPltTxtKargoIndDeger.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double indDeger);
            string firma = ucPltCmbKargoFirma.Text;
            string tip = ucPltCmbKargoIndTip.SelectedIndex == 0 ? "YÜZDE" : "TL";

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("UPDATE Platformlar SET KargoIndirimFirma=@f, KargoIndirimTip=@t, KargoIndirimDeger=@d WHERE Ad=@ad", b))
                    {
                        cmd.Parameters.AddWithValue("@f", firma);
                        cmd.Parameters.AddWithValue("@t", tip);
                        cmd.Parameters.AddWithValue("@d", indDeger);
                        cmd.Parameters.AddWithValue("@ad", secili);
                        cmd.ExecuteNonQuery();
                    }
                }
                TabloyuDoldur(secili);
                MessageBox.Show("Kargo lojistik indirimi platforma başarıyla tanımlandı!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void PlatformlariYukle()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open(); using (var cmd = new SqliteCommand("SELECT Ad FROM Platformlar ORDER BY Ad ASC", b))
                    using (var r = cmd.ExecuteReader()) { ucPltCmbPlatformlar.Items.Clear(); while (r.Read()) ucPltCmbPlatformlar.Items.Add(r["Ad"].ToString()); }
                }
            }
            catch { }
            ucPltCmbPlatformlar.SelectedIndex = -1;
        }

        private void pPltCmbHedefKatDoldur()
        {
            if (ucPltCmbHedefKat == null) return;
            ucPltCmbHedefKat.Items.Clear(); ucPltCmbHedefKat.Items.Add("TÜM KATEGORİLER");
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open(); using (var cmd = new SqliteCommand("SELECT Ad FROM Kategoriler ORDER BY Ad ASC", b))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) ucPltCmbHedefKat.Items.Add(r["Ad"].ToString().ToUpper()); }
                }
            }
            catch { }
            if (ucPltCmbHedefKat.Items.Count <= 1) ucPltCmbHedefKat.Items.AddRange(new string[] { "CEP TELEFONU", "AKSESUAR" });
            ucPltCmbHedefKat.SelectedIndex = 0;
        }

        private void TemizleGirisKutulari()
        {
            ucPltTxtHizmetBedeli.Text = "0,00";
            ucPltTxtPosOrani.Text = "0,00";
            ucPltTxtKatDeger.Text = "0,00";
            ucPltTxtGenelAd.Clear();
            ucPltTxtGenelDeger.Text = "0,00";
            ucPltTxtKargoIndDeger.Text = "0,00";
            if (ucPltCmbPosModu.Items.Count > 0) ucPltCmbPosModu.SelectedIndex = 0;
            if (ucPltCmbHedefKat.Items.Count > 0) ucPltCmbHedefKat.SelectedIndex = 0;
            if (ucPltCmbKatTip.Items.Count > 0) ucPltCmbKatTip.SelectedIndex = 0;
            if (ucPltCmbGenelTip.Items.Count > 0) ucPltCmbGenelTip.SelectedIndex = 0;
            if (ucPltCmbKargoFirma.Items.Count > 0) ucPltCmbKargoFirma.SelectedIndex = 0;
            if (ucPltCmbKargoIndTip.Items.Count > 0) ucPltCmbKargoIndTip.SelectedIndex = 0;
            ucPltChkKomisyonMuaf.Checked = false;
            ucPltChkStopajMuaf.Checked = false;
        }

        private void UcPltCmbPlatformlar_SelectedIndexChanged(object sender, EventArgs e)
        {
            string secili = ucPltCmbPlatformlar.Text;
            if (string.IsNullOrEmpty(secili))
            {
                pnlRightMain.Visible = false;
                var lblSecili = this.Controls["lblSeciliPlatform"] as Label;
                if (lblSecili != null) { lblSecili.Text = ""; lblSecili.Visible = false; }
                return;
            }

            pnlRightMain.Visible = true;
            var lbl = this.Controls.Find("lblSeciliPlatform", true);
            if (lbl.Length > 0 && lbl[0] is Label l)
            {
                l.Text = $"SEÇİLİ PLATFORM: {secili}";
                l.Visible = true;
            }
            pPltCmbHedefKatDoldur();
            TemizleGirisKutulari();

            TabloyuDoldur(secili);
        }

        private void TabloyuDoldur(string platformAd)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("KaynakTablo", typeof(string));
            dt.Columns.Add("Kesinti Türü", typeof(string));
            dt.Columns.Add("Hedef / Adı", typeof(string));
            dt.Columns.Add("Tip", typeof(string));
            dt.Columns.Add("Oran / Değer", typeof(double));

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("SELECT PosOrani, HizmetBedeli, PosModu, KomisyonMuaf, StopajMuaf, KargoIndirimFirma, KargoIndirimTip, KargoIndirimDeger FROM Platformlar WHERE Ad=@p", b))
                    {
                        cmd.Parameters.AddWithValue("@p", platformAd);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                double pos = Convert.ToDouble(r["PosOrani"]);
                                double hizm = Convert.ToDouble(r["HizmetBedeli"]);
                                string mod = r["PosModu"]?.ToString() ?? "ANA_TUTAR";
                                string modStr = mod == "INDIRIMLI" ? " (İndirimli Tutar Üzerinden)" : " (Ana Satış Tutarı Üzerinden)";

                                double kargoIndDeger = r["KargoIndirimDeger"] == DBNull.Value ? 0 : Convert.ToDouble(r["KargoIndirimDeger"]);
                                string kargoFirma = r["KargoIndirimFirma"]?.ToString() ?? "TÜMÜ";
                                string kargoTip = r["KargoIndirimTip"]?.ToString() ?? "TL";

                                ucPltChkKomisyonMuaf.Checked = Convert.ToInt32(r["KomisyonMuaf"]) == 1;
                                ucPltChkStopajMuaf.Checked = Convert.ToInt32(r["StopajMuaf"]) == 1;

                                if (hizm > 0) dt.Rows.Add("HIZM", "SABIT_HIZM", "Sabit Kesinti", "Sabit Hizmet Bedeli", "TL", hizm);
                                if (pos > 0) dt.Rows.Add("POS", "SABIT_POS", "Sabit Kesinti", "Güvence / POS Bedeli" + modStr, "YÜZDE", pos);

                                // Kargo İndirimini Tabloda Göster
                                if (kargoIndDeger > 0)
                                {
                                    dt.Rows.Add("KARGO", "SABIT_KARGO", "Kargo İndirimi", $"Firma: {kargoFirma}", kargoTip, kargoIndDeger);

                                    // Arayüz kutularını doldur
                                    if (ucPltCmbKargoFirma.Items.Contains(kargoFirma)) ucPltCmbKargoFirma.SelectedItem = kargoFirma;
                                    ucPltCmbKargoIndTip.SelectedIndex = kargoTip == "YÜZDE" ? 0 : 1;
                                    ucPltTxtKargoIndDeger.Text = kargoIndDeger.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));
                                }
                            }
                        }
                    }
                    using (var cmd = new SqliteCommand("SELECT rowid, KategoriAd, Tip, Deger FROM PlatformKategoriKesintileri WHERE PlatformAd=@p", b))
                    {
                        cmd.Parameters.AddWithValue("@p", platformAd);
                        using (var r = cmd.ExecuteReader()) { while (r.Read()) dt.Rows.Add(r["rowid"].ToString(), "KAT", "Kategoriye Özel", r["KategoriAd"].ToString(), r["Tip"].ToString(), Convert.ToDouble(r["Deger"])); }
                    }
                    using (var cmd = new SqliteCommand("SELECT rowid, UcretAdi, Tip, Deger FROM PlatformEkUcretler WHERE PlatformAd=@p", b))
                    {
                        cmd.Parameters.AddWithValue("@p", platformAd);
                        using (var r = cmd.ExecuteReader()) { while (r.Read()) dt.Rows.Add(r["rowid"].ToString(), "GEN", "Genel Ek Kesinti", r["UcretAdi"].ToString(), r["Tip"].ToString(), Convert.ToDouble(r["Deger"])); }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Tablo Doldurma Hatası: " + ex.Message); }

            ucPltDgvTumKesintiler.DataSource = null;
            ucPltDgvTumKesintiler.Columns.Clear();
            ucPltDgvTumKesintiler.DataSource = dt;

            if (ucPltDgvTumKesintiler.Columns.Contains("ID")) ucPltDgvTumKesintiler.Columns["ID"].Visible = false;
            if (ucPltDgvTumKesintiler.Columns.Contains("KaynakTablo")) ucPltDgvTumKesintiler.Columns["KaynakTablo"].Visible = false;
            if (ucPltDgvTumKesintiler.Columns.Contains("Oran / Değer")) ucPltDgvTumKesintiler.Columns["Oran / Değer"].DefaultCellStyle.Format = "N2";

            GridButonlariniEkle(ucPltDgvTumKesintiler);
            ucPltDgvTumKesintiler.ClearSelection();
        }

        private void GridButonlariniEkle(DataGridView dgv)
        {
            if (!dgv.Columns.Contains("btnEdit"))
            {
                dgv.Columns.Add(new DataGridViewButtonColumn { Name = "btnEdit", HeaderText = "", Text = "✏️ Düzenle", UseColumnTextForButtonValue = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White } });
            }
            if (!dgv.Columns.Contains("btnDel"))
            {
                dgv.Columns.Add(new DataGridViewButtonColumn { Name = "btnDel", HeaderText = "", Text = "❌ Sil", UseColumnTextForButtonValue = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Crimson, ForeColor = Color.White } });
            }
        }

        private void UcPltBtnYeniPlatformEkle_Click(object sender, EventArgs e)
        {
            string ad = ucPltTxtYeniPlatform.Text.Trim().ToUpper(); if (string.IsNullOrEmpty(ad)) return;
            if (ucPltCmbPlatformlar.Items.Contains(ad)) { MessageBox.Show("Bu platform var!", "Uyarı"); return; }
            try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("INSERT INTO Platformlar (Ad, PosOrani, HizmetBedeli, PosModu, KomisyonMuaf, StopajMuaf, KargoIndirimTip, KargoIndirimDeger, KargoIndirimFirma) VALUES (@a, 0, 0, 'ANA_TUTAR', 0, 0, 'TL', 0, 'TÜMÜ')", b)) { cmd.Parameters.AddWithValue("@a", ad); cmd.ExecuteNonQuery(); } } PlatformlariYukle(); ucPltCmbPlatformlar.SelectedItem = ad; ucPltTxtYeniPlatform.Clear(); } catch { }
        }

        private void KaydetMuafiyet()
        {
            string secili = ucPltCmbPlatformlar.Text; if (string.IsNullOrEmpty(secili)) return;
            int kM = ucPltChkKomisyonMuaf.Checked ? 1 : 0;
            int sM = ucPltChkStopajMuaf.Checked ? 1 : 0;
            try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Platformlar SET KomisyonMuaf=@k, StopajMuaf=@s WHERE Ad=@ad", b)) { cmd.Parameters.AddWithValue("@k", kM); cmd.Parameters.AddWithValue("@s", sM); cmd.Parameters.AddWithValue("@ad", secili); cmd.ExecuteNonQuery(); } } MessageBox.Show("Platform Muafiyet Ayarları Kilitlendi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void KaydetHizmetBedeli()
        {
            string secili = ucPltCmbPlatformlar.Text; if (string.IsNullOrEmpty(secili)) return;
            double.TryParse(ucPltTxtHizmetBedeli.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double hizm);
            try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Platformlar SET HizmetBedeli=@h WHERE Ad=@ad", b)) { cmd.Parameters.AddWithValue("@h", hizm); cmd.Parameters.AddWithValue("@ad", secili); cmd.ExecuteNonQuery(); } } TabloyuDoldur(secili); MessageBox.Show("Sabit Hizmet Bedeli Güncellendi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void KaydetPosBedeli()
        {
            string secili = ucPltCmbPlatformlar.Text; if (string.IsNullOrEmpty(secili)) return;
            double.TryParse(ucPltTxtPosOrani.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double pos);
            string mod = ucPltCmbPosModu.SelectedIndex == 1 ? "INDIRIMLI" : "ANA_TUTAR";
            try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Platformlar SET PosOrani=@p, PosModu=@m WHERE Ad=@ad", b)) { cmd.Parameters.AddWithValue("@p", pos); cmd.Parameters.AddWithValue("@m", mod); cmd.Parameters.AddWithValue("@ad", secili); cmd.ExecuteNonQuery(); } } TabloyuDoldur(secili); MessageBox.Show("Güvence/POS Bedeli Güncellendi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void EkleKategoriKesinti()
        {
            string secili = ucPltCmbPlatformlar.Text; if (string.IsNullOrEmpty(secili)) return;
            double.TryParse(ucPltTxtKatDeger.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double dgr);
            if (dgr <= 0) { MessageBox.Show("Lütfen 0'dan büyük bir değer girin!"); return; }
            string tip = ucPltCmbKatTip.SelectedIndex == 0 ? "YÜZDE" : "TL"; string kat = ucPltCmbHedefKat.Text;
            try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("INSERT INTO PlatformKategoriKesintileri (PlatformAd, KategoriAd, Tip, Deger) VALUES (@p, @k, @t, @d)", b)) { cmd.Parameters.AddWithValue("@p", secili); cmd.Parameters.AddWithValue("@k", kat); cmd.Parameters.AddWithValue("@t", tip); cmd.Parameters.AddWithValue("@d", dgr); cmd.ExecuteNonQuery(); } } TabloyuDoldur(secili); } catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void EkleGenelKesinti()
        {
            string secili = ucPltCmbPlatformlar.Text; if (string.IsNullOrEmpty(secili)) return;
            string ad = ucPltTxtGenelAd.Text.Trim(); if (string.IsNullOrEmpty(ad)) { MessageBox.Show("Kesinti adını yazın!"); return; }
            double.TryParse(ucPltTxtGenelDeger.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double dgr);
            if (dgr <= 0) { MessageBox.Show("Lütfen 0'dan büyük bir değer girin!"); return; }
            string tip = ucPltCmbGenelTip.SelectedIndex == 0 ? "TL" : "YÜZDE";
            try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("INSERT INTO PlatformEkUcretler (PlatformAd, UcretAdi, Tip, Deger) VALUES (@p, @u, @t, @d)", b)) { cmd.Parameters.AddWithValue("@p", secili); cmd.Parameters.AddWithValue("@u", ad); cmd.Parameters.AddWithValue("@t", tip); cmd.Parameters.AddWithValue("@d", dgr); cmd.ExecuteNonQuery(); } } TabloyuDoldur(secili); } catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void UcPltDgvTumKesintiler_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = ucPltDgvTumKesintiler.Columns[e.ColumnIndex].Name;
            string platformAd = ucPltCmbPlatformlar.Text;

            if (colName == "btnDel")
            {
                string id = ucPltDgvTumKesintiler.Rows[e.RowIndex].Cells["ID"].Value.ToString();
                string kaynak = ucPltDgvTumKesintiler.Rows[e.RowIndex].Cells["KaynakTablo"].Value.ToString();

                if (MessageBox.Show("Bu veriyi silmek/sıfırlamak istiyor musunuz?", "Onay", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    try
                    {
                        using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                        {
                            b.Open();
                            string sql = "";
                            if (kaynak == "SABIT_POS") sql = "UPDATE Platformlar SET PosOrani=0 WHERE Ad=@id";
                            else if (kaynak == "SABIT_HIZM") sql = "UPDATE Platformlar SET HizmetBedeli=0 WHERE Ad=@id";
                            else if (kaynak == "SABIT_KARGO") sql = "UPDATE Platformlar SET KargoIndirimDeger=0 WHERE Ad=@id";
                            else if (kaynak == "KAT") sql = "DELETE FROM PlatformKategoriKesintileri WHERE rowid=@id";
                            else if (kaynak == "GEN") sql = "DELETE FROM PlatformEkUcretler WHERE rowid=@id";

                            using (var cmd = new SqliteCommand(sql, b))
                            {
                                cmd.Parameters.AddWithValue("@id", (kaynak.StartsWith("SABIT") ? platformAd : id));
                                cmd.ExecuteNonQuery();
                            }
                        }
                        TabloyuDoldur(platformAd);
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                }
            }
            else if (colName == "btnEdit")
            {
                string id = ucPltDgvTumKesintiler.Rows[e.RowIndex].Cells["ID"].Value.ToString();
                string kaynak = ucPltDgvTumKesintiler.Rows[e.RowIndex].Cells["KaynakTablo"].Value.ToString();
                string ad = ucPltDgvTumKesintiler.Rows[e.RowIndex].Cells["Hedef / Adı"].Value.ToString();
                string tip = ucPltDgvTumKesintiler.Rows[e.RowIndex].Cells["Tip"].Value.ToString();
                double val = Convert.ToDouble(ucPltDgvTumKesintiler.Rows[e.RowIndex].Cells["Oran / Değer"].Value);

                Form frm = new Form { Text = "✏️ Düzenle", Size = new Size(320, 300), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MaximizeBox = false, MinimizeBox = false };

                TextBox tA = null; ComboBox cT = null; TextBox tV = null; ComboBox cMod = null;

                if (kaynak == "SABIT_POS")
                {
                    frm.Controls.Add(FormComboYarat("Mod:", out cMod, new Point(15, 10), 270, new string[] { "ANA_TUTAR", "INDIRIMLI" }));
                    cMod.SelectedIndex = ad.Contains("İndirimli") ? 1 : 0;
                    frm.Controls.Add(FormElemaniYarat("Oran Değeri (%):", out tV, new Point(15, 70), 270, val.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));
                }
                else if (kaynak == "SABIT_HIZM")
                {
                    frm.Controls.Add(FormElemaniYarat("Hizmet Bedeli (TL):", out tV, new Point(15, 10), 270, val.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));
                }
                else if (kaynak == "SABIT_KARGO")
                {
                    frm.Controls.Add(FormComboYarat("İndirim Tipi:", out cT, new Point(15, 10), 270, new string[] { "TL", "YÜZDE" })); cT.SelectedItem = tip;
                    frm.Controls.Add(FormElemaniYarat("İndirim Değeri:", out tV, new Point(15, 70), 270, val.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));
                }
                else if (kaynak == "GEN")
                {
                    frm.Controls.Add(FormElemaniYarat("Kesinti Adı:", out tA, new Point(15, 10), 270, ad));
                    frm.Controls.Add(FormComboYarat("Tip:", out cT, new Point(15, 70), 270, new string[] { "TL", "YÜZDE" })); cT.SelectedItem = tip;
                    frm.Controls.Add(FormElemaniYarat("Değer:", out tV, new Point(15, 130), 270, val.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));
                }
                else
                {
                    Label l = new Label { Text = $"Kategori: {ad}", Location = new Point(15, 15), AutoSize = true, ForeColor = Color.Orange, Font = new Font("Segoe UI", 9F, FontStyle.Bold) }; frm.Controls.Add(l);
                    frm.Controls.Add(FormComboYarat("Tip:", out cT, new Point(15, 50), 270, new string[] { "TL", "YÜZDE" })); cT.SelectedItem = tip;
                    frm.Controls.Add(FormElemaniYarat("Değer:", out tV, new Point(15, 110), 270, val.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"))));
                }

                Button b = new Button { Text = "💾 GÜNCELLE", Location = new Point(15, 205), Width = 270, Height = 40, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
                b.FlatAppearance.BorderSize = 0;
                b.Click += (sX, eX) => {
                    double.TryParse(tV.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nV);
                    try
                    {
                        using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                        {
                            conn.Open();
                            if (kaynak == "SABIT_POS")
                            {
                                using (var cmd = new SqliteCommand("UPDATE Platformlar SET PosOrani=@d, PosModu=@m WHERE Ad=@id", conn)) { cmd.Parameters.AddWithValue("@d", nV); cmd.Parameters.AddWithValue("@m", cMod.Text); cmd.Parameters.AddWithValue("@id", platformAd); cmd.ExecuteNonQuery(); }
                            }
                            else if (kaynak == "SABIT_HIZM")
                            {
                                using (var cmd = new SqliteCommand("UPDATE Platformlar SET HizmetBedeli=@d WHERE Ad=@id", conn)) { cmd.Parameters.AddWithValue("@d", nV); cmd.Parameters.AddWithValue("@id", platformAd); cmd.ExecuteNonQuery(); }
                            }
                            else if (kaynak == "SABIT_KARGO")
                            {
                                using (var cmd = new SqliteCommand("UPDATE Platformlar SET KargoIndirimDeger=@d, KargoIndirimTip=@t WHERE Ad=@id", conn)) { cmd.Parameters.AddWithValue("@d", nV); cmd.Parameters.AddWithValue("@t", cT.Text); cmd.Parameters.AddWithValue("@id", platformAd); cmd.ExecuteNonQuery(); }
                            }
                            else if (kaynak == "KAT")
                            {
                                using (var cmd = new SqliteCommand("UPDATE PlatformKategoriKesintileri SET Tip=@t, Deger=@d WHERE rowid=@id", conn)) { cmd.Parameters.AddWithValue("@t", cT.Text); cmd.Parameters.AddWithValue("@d", nV); cmd.Parameters.AddWithValue("@id", id); cmd.ExecuteNonQuery(); }
                            }
                            else
                            {
                                using (var cmd = new SqliteCommand("UPDATE PlatformEkUcretler SET UcretAdi=@u, Tip=@t, Deger=@d WHERE rowid=@id", conn)) { cmd.Parameters.AddWithValue("@u", tA.Text); cmd.Parameters.AddWithValue("@t", cT.Text); cmd.Parameters.AddWithValue("@d", nV); cmd.Parameters.AddWithValue("@id", id); cmd.ExecuteNonQuery(); }
                            }
                        }
                        TabloyuDoldur(platformAd); frm.Close();
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                }; frm.Controls.Add(b); frm.ShowDialog();
            }
        }

        private DataGridView GridYarat(Color baslikRengi) { DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9.5F), BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false, ColumnHeadersHeight = 35, ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing }; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = baslikRengi; dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold); dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgv.DefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DataError += (s, e) => { e.Cancel = true; }; return dgv; }

        // 🚀 MÜHÜR: Genişlik parametresi ile oluşturulan kutunun içine yerleşen combobox genişliği eşitlendi!
        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan) { Panel p = new Panel { Location = konum, Width = genislik, Height = 45 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 3), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.DarkGray }; kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
        private Panel FormComboYarat(string baslik, out ComboBox kutu, Point konum, int genislik, string[] ogeler) { Panel p = new Panel { Location = konum, Width = genislik, Height = 45 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 3), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.DarkGray }; kutu = new ComboBox { Location = new Point(0, 18), Width = genislik, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat }; kutu.Items.AddRange(ogeler); if (kutu.Items.Count > 0) kutu.SelectedIndex = 0; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
    }
}