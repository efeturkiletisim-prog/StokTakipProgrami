using Microsoft.Data.Sqlite;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Globalization;

namespace StokTakipProgrami
{
    public partial class Form1 : Form
    {
        public static Form1 AnaForm;
        private Button btnCikis;
        private Panel pnlBrand;
        private Panel pnlAyarlar;
        private Button btnAyarlarAna;

        // --- HİYERARŞİ KONTROLLERİ ---
        private Panel pnlSubPlatformlar;
        private bool platMenuAcik = false;

        // --- PROFESYONEL REFERANS HAFIZASI ---
        private Button cacheBtnSipOde;
        private Button cacheBtnPlatOde;
        private Button cacheBtnCezaAlt;

        public Form1()
        {
            InitializeComponent();
            AnaForm = this;
            
            this.Load += Form1_Load;

            if (button1 != null) { button1.Click -= button1_Click; button1.Click += button1_Click; }
            if (btnSiparislerAna != null) { btnSiparislerAna.Click -= btnSiparislerAna_Click; btnSiparislerAna.Click += btnSiparislerAna_Click; }
            if (btnPlatformlarAna != null) { btnPlatformlarAna.Click -= btnPlatformlarAna_Click; btnPlatformlarAna.Click += btnPlatformlarAna_Click; }
            if (btnHesaplarAna != null) { btnHesaplarAna.Click -= btnHesaplarAna_Click; btnHesaplarAna.Click += btnHesaplarAna_Click; }
            if (btnRaporlarAna != null) { btnRaporlarAna.Click -= btnRaporlarAna_Click; btnRaporlarAna.Click += btnRaporlarAna_Click; }
            if (btnAnaSayfa != null) { btnAnaSayfa.Click -= AnaSayfaAc; btnAnaSayfa.Click += AnaSayfaAc; }
        }

        private Control[] ControlsFind(string name, bool searchAllChildren) { return this.Controls.Find(name, searchAllChildren); }
        private Button ButonBul(string ad) { Control[] bulunanlar = ControlsFind(ad, true); if (bulunanlar.Length > 0 && bulunanlar[0] is Button btn) return btn; return null; }

        private void Form1_Load(object sender, EventArgs e)
        {
            AnaFormuYukle();
        }

        public void AnaFormuYukle()
        {
            AnaMenuButonStiliniGiy(btnAnaSayfa, "🏠  Ana Sayfa");
            AnaMenuButonStiliniGiy(button1, "📦  Ürün/Stok Yönetimi");
            AnaMenuButonStiliniGiy(btnSiparislerAna, "🛒  Siparişler");
            AnaMenuButonStiliniGiy(btnPlatformlarAna, "🌐  Pazaryeri");
            AnaMenuButonStiliniGiy(btnHesaplarAna, "💰  Muhasebe");
            AnaMenuButonStiliniGiy(btnRaporlarAna, "📊  Rapor");

            // --- ÜRÜN & STOK YÖNETİMİ ALT MENÜSÜ ---
            if (button3 != null) button3.Text = "📋  Ürün Ekle";
            if (button4 != null) button4.Text = "⚡  Toplu Veri Ekle";
            if (button2 != null) button2.Text = "➕  Veri Ekle";

            Button btnStokSorgula = ButonBul("btnStokSorgula") ?? new Button { Name = "btnStokSorgula" };
            btnStokSorgula.Text = "🔎  Stok Sorgula";
            btnStokSorgula.Click -= (s, ev) => PanelSayfaAc(new SayfaStokSorgula());
            btnStokSorgula.Click += (s, ev) => PanelSayfaAc(new SayfaStokSorgula());

            Button btnStokDefteriAlt = ButonBul("btnStokDefteriAlt") ?? new Button { Name = "btnStokDefteriAlt" };
            btnStokDefteriAlt.Text = "📜  Stok Hareketleri";
            btnStokDefteriAlt.Click -= (s, ev) => PanelSayfaAc(new SayfaStokHareketleri());
            btnStokDefteriAlt.Click += (s, ev) => PanelSayfaAc(new SayfaStokHareketleri());

            Button btnKatMarkaAlt = ButonBul("btnKatMarkaAlt") ?? new Button { Name = "btnKatMarkaAlt" };
            btnKatMarkaAlt.Text = "🏷️  Kategori/Marka";
            btnKatMarkaAlt.Click -= (s, ev) => PanelSayfaAc(new SayfaKategoriMarka());
            btnKatMarkaAlt.Click += (s, ev) => PanelSayfaAc(new SayfaKategoriMarka());

            AltMenuDuzenle(pnlUrunler, button3, btnStokSorgula, btnStokDefteriAlt, btnKatMarkaAlt, button4, button2);

            // --- SİPARİŞLER ALT MENÜSÜ ---
            if (btnSiparisEkle != null) btnSiparisEkle.Text = "🛍️  Sipariş Ekle";
            if (btnSiparisSorgula != null) btnSiparisSorgula.Text = "🔍  Sipariş Sorgulama";
            AltMenuDuzenle(pnlSiparisler, btnSiparisEkle, btnSiparisSorgula);

            // --- PLATFORM (PAZARYERİ) ALT MENÜSÜ ---
            if (btnPlatformEkleAlt != null) { btnPlatformEkleAlt.Visible = false; }
            if (btnPlatformAyarlariAlt != null) { btnPlatformAyarlariAlt.Text = "⚙️  Platform Ayarları"; }
            if (btnPlatformAyarlariAlt != null) { btnPlatformAyarlariAlt.Click -= (s, ev) => PanelSayfaAc(new SayfaPlatformAyarlari()); btnPlatformAyarlariAlt.Click += (s, ev) => PanelSayfaAc(new SayfaPlatformAyarlari()); }

            Button btnKargo = ButonBul("btnKargo") ?? new Button { Name = "btnKargo" };
            btnKargo.Text = "🚚  Kargo Ayarları";
            btnKargo.Click -= (s, ev) => PanelSayfaAc(new SayfaKargoMatrisi());
            btnKargo.Click += (s, ev) => PanelSayfaAc(new SayfaKargoMatrisi());

            Button btnPlatformlarYonetimi = ButonBul("btnPlatformlarYonetimi") ?? new Button { Name = "btnPlatformlarYonetimi" };
            btnPlatformlarYonetimi.Text = "🔗  Platformlar";
            btnPlatformlarYonetimi.Click -= (s, ev) => PanelSayfaAc(new SayfaPlatformlar());
            btnPlatformlarYonetimi.Click += (s, ev) => PanelSayfaAc(new SayfaPlatformlar());

            AltMenuDuzenle(pnlPlatformlar, btnPlatformAyarlariAlt, btnKargo, btnPlatformlarYonetimi);

            // --- MUHASEBE ALT MENÜSÜ ---
            MuhasebeMenusunuGuncelle();

            // --- RAPORLAR ALT MENÜSÜ ---
            Button bAylik = ButonBul("btnAylikRaporAlt") ?? new Button { Name = "btnAylikRaporAlt" }; bAylik.Text = "📅  Aylık Satış Raporu";
            Button bYillik = ButonBul("btnYillikRaporAlt") ?? new Button { Name = "btnYillikRaporAlt" }; bYillik.Text = "📆  Yıllık Satış Raporu";
            Button bGenel = ButonBul("btnGenelRaporAlt") ?? new Button { Name = "btnGenelRaporAlt" }; bGenel.Text = "👑  Genel Satış Raporu";
            bAylik.Click -= (s, ev) => PanelSayfaAc(new SayfaAylikRapor()); bAylik.Click += (s, ev) => PanelSayfaAc(new SayfaAylikRapor());
            bYillik.Click -= (s, ev) => PanelSayfaAc(new SayfaYillikRapor()); bYillik.Click += (s, ev) => PanelSayfaAc(new SayfaYillikRapor());
            bGenel.Click -= (s, ev) => PanelSayfaAc(new SayfaGenelRapor()); bGenel.Click += (s, ev) => PanelSayfaAc(new SayfaGenelRapor());
            AltMenuDuzenle(pnlRaporlar, bAylik, bYillik, bGenel);

            // --- AYARLAR ALT MENÜSÜ ---
            pnlAyarlar = new Panel { Name = "pnlAyarlar" };
            btnAyarlarAna = new Button { Name = "btnAyarlarAna" };
            AnaMenuButonStiliniGiy(btnAyarlarAna, "⚙️  Sistem Ayarları");
            btnAyarlarAna.Click -= (s, ev) => MenüKapatAc(pnlAyarlar); btnAyarlarAna.Click += (s, ev) => MenüKapatAc(pnlAyarlar);
            Button bKullaniciAyar = ButonBul("btnKullaniciAyarAlt") ?? new Button { Name = "btnKullaniciAyarAlt", Text = "👥  Kullanıcı & Personel Yetki" };
            bKullaniciAyar.Click -= (s, ev) => PanelSayfaAc(new SayfaKullanicilar()); bKullaniciAyar.Click += (s, ev) => PanelSayfaAc(new SayfaKullanicilar());
            AltMenuDuzenle(pnlAyarlar, bKullaniciAyar);

            // --- ÇIKIŞ BUTONU ve MARKA PANELİ ---
            btnCikis = new Button { Name = "btnCikis", Text = "🚪 Çıkış Yap", Height = 50, Dock = DockStyle.Bottom, FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(26, 26, 39), ForeColor = Color.FromArgb(241, 65, 108), Font = new Font("Segoe UI", 10F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand };
            btnCikis.FlatAppearance.BorderSize = 0; btnCikis.Click += BtnCikis_Click;
            btnCikis.MouseEnter += (s, ev) => { btnCikis.BackColor = Color.FromArgb(241, 65, 108); btnCikis.ForeColor = Color.White; };
            btnCikis.MouseLeave += (s, ev) => { btnCikis.BackColor = Color.FromArgb(26, 26, 39); btnCikis.ForeColor = Color.FromArgb(241, 65, 108); };

            pnlBrand = new Panel { Dock = DockStyle.Top, Height = 90, BackColor = Color.FromArgb(21, 21, 33) };
            Panel pnlCizgi = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = Color.FromArgb(15, 15, 23) }; pnlBrand.Controls.Add(pnlCizgi);
            Label lblEfe = new Label { Text = "EFETURK", Font = new Font("Segoe UI Black", 16F, FontStyle.Bold), ForeColor = Color.White, AutoSize = true, Location = new Point(15, 25) };
            Label lblNet = new Label { Text = ".NET", Font = new Font("Segoe UI Black", 16F, FontStyle.Bold), ForeColor = Color.FromArgb(0, 158, 247), AutoSize = true };
            Label lblSub = new Label { Text = "BİLGİ SİSTEMLERİ", Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.FromArgb(100, 100, 115), Location = new Point(19, 57), AutoSize = true };
            pnlBrand.Controls.Add(lblEfe); pnlBrand.Controls.Add(lblNet); pnlBrand.Controls.Add(lblSub);
            pnlBrand.Layout += (s, ev) => { lblNet.Location = new Point(lblEfe.Right - 5, 25); };

            Control solMenüPaneli = btnAnaSayfa?.Parent;
            if (solMenüPaneli != null)
            {
                solMenüPaneli.BackColor = Color.FromArgb(26, 26, 39); solMenüPaneli.Controls.Clear();
                solMenüPaneli.Controls.Add(btnCikis); solMenüPaneli.Controls.Add(pnlAyarlar); solMenüPaneli.Controls.Add(btnAyarlarAna);
                solMenüPaneli.Controls.Add(pnlRaporlar); solMenüPaneli.Controls.Add(btnRaporlarAna);
                solMenüPaneli.Controls.Add(pnlHesaplar); solMenüPaneli.Controls.Add(btnHesaplarAna);
                solMenüPaneli.Controls.Add(pnlPlatformlar); solMenüPaneli.Controls.Add(btnPlatformlarAna);
                solMenüPaneli.Controls.Add(pnlSiparisler); solMenüPaneli.Controls.Add(btnSiparislerAna);
                solMenüPaneli.Controls.Add(pnlUrunler); solMenüPaneli.Controls.Add(button1);
                solMenüPaneli.Controls.Add(btnAnaSayfa); solMenüPaneli.Controls.Add(pnlBrand);
            }

            // 🚀 MÜHÜR 1: İLK AÇILIŞTA TÜM SİSTEMİ KAPATARAK TERTEMİZ BAŞLATAN FABRİKA AYAR KALKANI
            if (pnlUrunler != null) pnlUrunler.Visible = false;
            if (pnlSiparisler != null) pnlSiparisler.Visible = false;
            if (pnlPlatformlar != null) pnlPlatformlar.Visible = false;
            if (pnlHesaplar != null) pnlHesaplar.Visible = false;
            if (pnlRaporlar != null) pnlRaporlar.Visible = false;
            if (pnlAyarlar != null) pnlAyarlar.Visible = false;
            platMenuAcik = false;
            MuhasebeMenusunuGuncelle();

            AnaEkraniYukle();
        }

        public void MuhasebeMenusunuGuncelle()
        {
            if (cacheBtnSipOde == null) cacheBtnSipOde = btnSiparisOdemeleriAlt ?? ButonBul("btnSiparisOdemeleriAlt");
            if (cacheBtnPlatOde == null) cacheBtnPlatOde = btnPlatformOdemeleriAlt ?? ButonBul("btnPlatformOdemeleriAlt") ?? ButonBul("btnPlatformAlt");
            if (cacheBtnCezaAlt == null) cacheBtnCezaAlt = ButonBul("btnSiparisCezalariAlt") ?? new Button { Name = "btnSiparisCezalariAlt" };

            Button btnSipOde = cacheBtnSipOde;
            Button btnPlatOdeAna = cacheBtnPlatOde;
            Button btnCezaAlt = cacheBtnCezaAlt;

            if (pnlSubPlatformlar == null)
            {
                pnlSubPlatformlar = new Panel { Name = "pnlSubPlatformlar", Dock = DockStyle.Top, BackColor = Color.FromArgb(16, 16, 24) };
            }

            // 🚀 MÜHÜR 2: ARKA PLANDA GİZLİ KALMA REZALETİNİ BİTİREN ANLIK DURUM AYNASI
            pnlSubPlatformlar.Visible = platMenuAcik;

            pnlSubPlatformlar.Controls.Clear();
            int aktifKanalSayisi = 0;
            Font idealFont = new Font("Segoe UI", 9.5F, FontStyle.Bold);

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Ad FROM Platformlar WHERE AktifMi = 1 ORDER BY Ad ASC", conn))
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            string platAd = r[0].ToString();
                            string bName = "btnDynPlatMenu_" + platAd.Replace(" ", "");

                            Button bSub = new Button
                            {
                                Name = bName,
                                Dock = DockStyle.Top,
                                Height = 40,
                                FlatStyle = FlatStyle.Flat,
                                BackColor = Color.Transparent,
                                ForeColor = Color.FromArgb(200, 200, 210),
                                Font = idealFont,
                                TextAlign = ContentAlignment.MiddleLeft,
                                Padding = new Padding(65, 0, 0, 0),
                                Cursor = Cursors.Hand
                            };
                            bSub.FlatAppearance.BorderSize = 0;

                            string guzelAd = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(platAd.ToLower());
                            bSub.Text = $"💳  {guzelAd}";

                            bSub.MouseEnter += (s, e) => { bSub.BackColor = Color.FromArgb(32, 32, 48); bSub.ForeColor = Color.White; };
                            bSub.MouseLeave += (s, e) => { bSub.BackColor = Color.Transparent; bSub.ForeColor = Color.FromArgb(200, 200, 210); };
                            bSub.Click += DynamicPlatformClick;

                            pnlSubPlatformlar.Controls.Add(bSub);
                            aktifKanalSayisi++;
                        }
                    }
                }
            }
            catch { }

            pnlSubPlatformlar.Height = aktifKanalSayisi * 40;

            if (pnlHesaplar != null)
            {
                bool mevcutGorunurluk = pnlHesaplar.Visible;
                pnlHesaplar.Controls.Clear();

                if (btnCezaAlt != null)
                {
                    btnCezaAlt.Text = "Sipariş Ceza Kesintileri";
                    btnCezaAlt.Font = idealFont; btnCezaAlt.Height = 42; btnCezaAlt.Padding = new Padding(45, 0, 0, 0); btnCezaAlt.TextAlign = ContentAlignment.MiddleLeft; btnCezaAlt.Dock = DockStyle.Top; btnCezaAlt.FlatStyle = FlatStyle.Flat; btnCezaAlt.FlatAppearance.BorderSize = 0; btnCezaAlt.ForeColor = Color.White; btnCezaAlt.Click -= OpenCezaAlt; btnCezaAlt.Click += OpenCezaAlt; pnlHesaplar.Controls.Add(btnCezaAlt);
                }

                pnlSubPlatformlar.Dock = DockStyle.Top;
                pnlHesaplar.Controls.Add(pnlSubPlatformlar);

                if (btnPlatOdeAna != null)
                {
                    btnPlatOdeAna.Font = idealFont;
                    btnPlatOdeAna.Height = 42;
                    btnPlatOdeAna.Padding = new Padding(45, 0, 0, 0);
                    btnPlatOdeAna.TextAlign = ContentAlignment.MiddleLeft;
                    btnPlatOdeAna.Dock = DockStyle.Top;
                    btnPlatOdeAna.FlatStyle = FlatStyle.Flat;
                    btnPlatOdeAna.FlatAppearance.BorderSize = 0;
                    btnPlatOdeAna.ForeColor = Color.White;
                    btnPlatOdeAna.Text = platMenuAcik ? "▼ 💸  Toplu Platform Ödemeleri" : "▶ 💸  Toplu Platform Ödemeleri";
                    btnPlatOdeAna.Click -= BtnPlatOdeAna_Click; btnPlatOdeAna.Click += BtnPlatOdeAna_Click;
                    pnlHesaplar.Controls.Add(btnPlatOdeAna);
                }

                if (btnSipOde != null) { btnSipOde.Font = idealFont; btnSipOde.Height = 42; btnSipOde.Padding = new Padding(45, 0, 0, 0); btnSipOde.TextAlign = ContentAlignment.MiddleLeft; btnSipOde.Dock = DockStyle.Top; btnSipOde.FlatStyle = FlatStyle.Flat; btnSipOde.FlatAppearance.BorderSize = 0; btnSipOde.ForeColor = Color.White; btnSipOde.Text = "Sipariş Ödemeleri"; btnSipOde.Click -= OpenSipOde; btnSipOde.Click += OpenSipOde; pnlHesaplar.Controls.Add(btnSipOde); }

                int toplamH = 84;
                if (platMenuAcik) toplamH += pnlSubPlatformlar.Height;
                if (btnCezaAlt != null) toplamH += 42;

                pnlHesaplar.Height = toplamH;
                pnlHesaplar.Dock = DockStyle.Top;
                pnlHesaplar.Visible = mevcutGorunurluk;
            }
        }

        private void BtnPlatOdeAna_Click(object sender, EventArgs e)
        {
            platMenuAcik = !platMenuAcik;
            MuhasebeMenusunuGuncelle();
        }

        private void OpenSipOde(object sender, EventArgs e) { platMenuAcik = false; MuhasebeMenusunuGuncelle(); PanelSayfaAc(new SayfaSiparisOdemeleri()); }
        private void OpenCezaAlt(object sender, EventArgs e) { platMenuAcik = false; MuhasebeMenusunuGuncelle(); PanelSayfaAc(new SayfaSiparisCezalari()); }

        private void DynamicPlatformClick(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                string text = btn.Text;
                string platName = text.Replace("CNL", "").Replace("衡量", "").Replace("💳", "").Trim().ToUpper(CultureInfo.InvariantCulture);

                // 🚀 MÜHÜR: 16 Platformun tamamını bağımsız fiziksel sınıflarına dağıtan kurumsal yönlendirme kalkanı
                if (platName.Contains("TRENDYOL"))
                {
                    PanelSayfaAc(new SayfaTrendyolOdemeleri());
                }
                else if (platName.Contains("HEPSİBURADA") || platName.Contains("HEPSI"))
                {
                    PanelSayfaAc(new SayfaHepsiburadaOdemeleri());
                }
                else if (platName.Contains("AMAZON"))
                {
                    PanelSayfaAc(new SayfaAmazonOdemeleri());
                }
                else if (platName.Contains("N11"))
                {
                    PanelSayfaAc(new SayfaN11Odemeleri());
                }
                else if (platName.Contains("ÇİÇEK") || platName.Contains("CICEK"))
                {
                    PanelSayfaAc(new SayfaCicekSepetiOdemeleri());
                }
                else if (platName.Contains("PAZARAMA"))
                {
                    PanelSayfaAc(new SayfaPazaramaOdemeleri());
                }
                else if (platName.Contains("TEKNOSA"))
                {
                    PanelSayfaAc(new SayfaTeknosaOdemeleri());
                }
                else if (platName.Contains("MEDIA") || platName.Contains("MEDİA"))
                {
                    PanelSayfaAc(new SayfaMediaMarktOdemeleri());
                }
                else if (platName.Contains("SAHİBİNDEN") || platName.Contains("SAHIBINDEN"))
                {
                    PanelSayfaAc(new SayfaSahibindenOdemeleri());
                }
                else if (platName.Contains("TEMU"))
                {
                    PanelSayfaAc(new SayfaTemuOdemeleri());
                }
                else if (platName.Contains("PASAJ") || platName.Contains("TURKCELL"))
                {
                    PanelSayfaAc(new SayfaTurkcellPasajOdemeleri());
                }
                else if (platName.Contains("PTT"))
                {
                    PanelSayfaAc(new SayfaPttAvmOdemeleri());
                }
                else if (platName.Contains("GOTURC"))
                {
                    PanelSayfaAc(new SayfaGoTurcOdemeleri());
                }
                else if (platName.Contains("GÖRMEK") || platName.Contains("GORMEK"))
                {
                    PanelSayfaAc(new SayfaGormekLazimOdemeleri());
                }
                else if (platName.ToUpper().Contains("DEF"))
                {
                    PanelSayfaAc(new SayfaIdefixOdemeleri());
                }
                else if (platName.Contains("EFETÜRK") || platName.Contains("EFETURK"))
                {
                    PanelSayfaAc(new SayfaEfeTurkOdemeleri());
                }
                else
                {
                    // Gelecekte eklenebilecek yeni pazar yerleri için sistem kilitleme emniyeti
                    SayfaPlatformOdemeleri2 sayfa = new SayfaPlatformOdemeleri2();
                    sayfa.SetPlatform(platName);
                    PanelSayfaAc(sayfa);
                }
            }
        }

        public void PanelSayfaAc(UserControl sayfa)
        {
            if (pnlOrta != null)
            {
                pnlOrta.Controls.Clear(); sayfa.Dock = DockStyle.Fill; pnlOrta.Controls.Add(sayfa);
                try { sayfa.BeginInvoke((Action)(() => { if (sayfa is SayfaStokHareketleri h) { try { h.Yenile(); } catch { } } })); } catch { }
            }
        }

        private void AnaMenuButonStiliniGiy(Button btn, string yeniMetin) { if (btn == null) return; btn.Text = yeniMetin; btn.Height = 50; btn.Dock = DockStyle.Top; btn.FlatStyle = FlatStyle.Flat; btn.FlatAppearance.BorderSize = 0; btn.BackColor = Color.FromArgb(26, 26, 39); btn.ForeColor = Color.FromArgb(154, 156, 174); btn.Font = new Font("Segoe UI", 10F, FontStyle.Bold); btn.TextAlign = ContentAlignment.MiddleLeft; btn.Padding = new Padding(15, 0, 0, 0); btn.Cursor = Cursors.Hand; btn.MouseEnter += (s, e) => { btn.BackColor = Color.FromArgb(43, 43, 64); btn.ForeColor = Color.White; }; btn.MouseLeave += (s, e) => { btn.BackColor = Color.FromArgb(26, 26, 39); btn.ForeColor = Color.FromArgb(154, 156, 174); }; }

        private void AltMenuDuzenle(Panel panel, params Button[] butonlar)
        {
            if (panel == null) return;
            bool wasVisible = panel.Visible;
            panel.Controls.Clear();
            panel.BackColor = Color.FromArgb(21, 21, 33);
            int toplamYukseklik = 0;
            Font idealFont = new Font("Segoe UI", 9.5F, FontStyle.Bold);

            for (int i = butonlar.Length - 1; i >= 0; i--)
            {
                Button btn = butonlar[i];
                if (btn != null)
                {
                    btn.UseVisualStyleBackColor = false; btn.Dock = DockStyle.Top; btn.Height = 42; btn.Margin = new Padding(0); btn.Padding = new Padding(45, 0, 0, 0); btn.FlatStyle = FlatStyle.Flat; btn.FlatAppearance.BorderSize = 0; btn.BackColor = Color.Transparent; btn.ForeColor = Color.White; btn.Font = idealFont; btn.TextAlign = ContentAlignment.MiddleLeft; btn.Cursor = Cursors.Hand; btn.AutoSize = false;
                    btn.MouseEnter -= AltButton_MouseEnter; btn.MouseEnter += AltButton_MouseEnter; btn.MouseLeave -= AltButton_MouseLeave; btn.MouseLeave += AltButton_MouseLeave;
                    panel.Controls.Add(btn); toplamYukseklik += btn.Height;
                }
            }
            panel.Height = toplamYukseklik; panel.Dock = DockStyle.Top;
            panel.Visible = wasVisible;
        }

        private void AltButton_MouseEnter(object sender, EventArgs e) { if (sender is Button btn) { btn.BackColor = Color.FromArgb(32, 32, 48); btn.ForeColor = Color.White; } }
        private void AltButton_MouseLeave(object sender, EventArgs e) { if (sender is Button btn) { btn.BackColor = Color.Transparent; btn.ForeColor = Color.White; } }
        public void AnaEkraniYukle() { PanelSayfaAc(new SayfaAnaEkran()); }

        // 🚀 MÜHÜR 3: BİR MENÜDEN BAŞKA BİR MENÜYE GEÇİLDİĞİNDE ESKİ AÇILAN HER ŞEYİ OTOMATİK KAPATAN GERÇEK AKORDEON MOTORU
        private void MenüKapatAc(Panel p)
        {
            if (p != null)
            {
                bool d = p.Visible;

                // Önce istisnasız tüm alt menü panellerini görünmez yapıyoruz
                if (pnlUrunler != null) pnlUrunler.Visible = false;
                if (pnlSiparisler != null) pnlSiparisler.Visible = false;
                if (pnlPlatformlar != null) pnlPlatformlar.Visible = false;
                if (pnlHesaplar != null) pnlHesaplar.Visible = false;
                if (pnlRaporlar != null) pnlRaporlar.Visible = false;
                if (pnlAyarlar != null) pnlAyarlar.Visible = false;

                // Eğer tıklanan yer Muhasebe (pnlHesaplar) DEĞİLSE, oradaki hakediş alt kırılımını da tamamen sıfırla/kapat
                if (p != pnlHesaplar)
                {
                    platMenuAcik = false;
                    MuhasebeMenusunuGuncelle();
                }

                p.Visible = !d;
            }
        }

        public void button1_Click(object sender, EventArgs e) { MenüKapatAc(pnlUrunler); }
        public void button2_Click(object sender, EventArgs e) { PanelSayfaAc(new SayfaUrunEkle()); }
        public void button3_Click(object sender, EventArgs e) { PanelSayfaAc(new SayfaStoklar()); }
        public void button4_Click(object sender, EventArgs e) { PanelSayfaAc(new SayfaVeriler()); }
        public void btnSiparislerAna_Click(object sender, EventArgs e) { MenüKapatAc(pnlSiparisler); }
        public void btnPlatformlarAna_Click(object sender, EventArgs e) { MenüKapatAc(pnlPlatformlar); }
        public void btnHesaplarAna_Click(object sender, EventArgs e) { MenüKapatAc(pnlHesaplar); }
        public void btnRaporlarAna_Click(object sender, EventArgs e) { MenüKapatAc(pnlRaporlar); }

        // 🚀 MÜHÜR 4: ANA SAYFAYA GEÇİLDİĞİNDE DE TÜM SİSTEMİ TERTEMİZ KAPATAN ABSOLUTE TETİKLEYİCİ
        public void AnaSayfaAc(object sender, EventArgs e)
        {
            platMenuAcik = false;
            if (pnlUrunler != null) pnlUrunler.Visible = false;
            if (pnlSiparisler != null) pnlSiparisler.Visible = false;
            if (pnlPlatformlar != null) pnlPlatformlar.Visible = false;
            if (pnlHesaplar != null) pnlHesaplar.Visible = false;
            if (pnlRaporlar != null) pnlRaporlar.Visible = false;
            if (pnlAyarlar != null) pnlAyarlar.Visible = false;
            MuhasebeMenusunuGuncelle();
            AnaEkraniYukle();
        }

        private void BtnCikis_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Mevcut oturumu kapatıp giriş ekranına dönmek istediğinize emin misiniz?", "Güvenli Oturum Kapatma", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes) { Veritabanı.OtomatikYedekle(); Application.Restart(); }
        }
        public void btnSiparisEkle_Click(object sender, EventArgs e) { PanelSayfaAc(new SayfaSiparisEkle()); }
        public void btnSiparisSorgula_Click(object sender, EventArgs e) { PanelSayfaAc(new SayfaSiparisSorgula()); }
        public void pnlOrta_Paint(object sender, PaintEventArgs e) { }
        private void SistemiFabrikaAyarlarinaDondur()
        {
            // UYARI: Bu kod çalıştırıldığında veritabanındaki tüm sipariş geçmişi kalıcı olarak silinir.
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();

                    // 1. Tüm sipariş kayıtlarını temizle
                    using (var cmd = new SqliteCommand("DELETE FROM Siparisler;", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // 2. SQLite ID sayacını sıfırla (Böylece yeni eklenen ilk sipariş 1 ID'si ile başlar)
                    using (var cmd = new SqliteCommand("DELETE FROM sqlite_sequence WHERE name='Siparisler';", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // NOT: Eğer stok hareketleri veya cari hesapları da sıfırlamak istersen, 
                    // aşağıdaki satırların başındaki yorum satırlarını (//) kaldırabilirsin:
                    // using (var cmd = new SqliteCommand("DELETE FROM StokHareketleri;", conn)) { cmd.ExecuteNonQuery(); }
                    // using (var cmd = new SqliteCommand("DELETE FROM sqlite_sequence WHERE name='StokHareketleri';", conn)) { cmd.ExecuteNonQuery(); }

                    MessageBox.Show("Veritabanındaki tüm siparişler kalıcı olarak silindi ve sayaçlar sıfırlandı! Sistem fabrika ayarlarına döndürüldü.",
                                    "Sistem Sıfırlama Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veritabanı sıfırlanırken kurumsal bir hata oluştu: " + ex.Message,
                                "Sıfırlama Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}