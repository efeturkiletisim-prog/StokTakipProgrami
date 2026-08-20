using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public partial class SayfaYedekleme : UserControl
    {
        private ListBox lstYedekler;
        private Button btnManuelYedek, btnHariciYedek, btnGeriYukle, btnYedekSil;
        private string yedekKlasorYolu;

        public SayfaYedekleme()
        {
            InitializeComponent();
            this.Load += SayfaYedekleme_Load;
        }

        private void SayfaYedekleme_Load(object sender, EventArgs e)
        {
            yedekKlasorYolu = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Yedekler");
            if (!Directory.Exists(yedekKlasorYolu)) Directory.CreateDirectory(yedekKlasorYolu);

            ArayuzYarat();
            YedekleriListele();
        }

        private void ArayuzYarat()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(245, 246, 250);

            // Üst Başlık Paneli
            Panel pnlBaslik = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(41, 41, 44) };
            Label lblBaslik = new Label { Text = "🛡️ VERİTABANI YEDEKLEME MERKEZİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 11F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlBaslik.Controls.Add(lblBaslik);
            this.Controls.Add(pnlBaslik);

            // Sol Kontrol Paneli
            Panel pnlSol = new Panel { Dock = DockStyle.Left, Width = 300, BackColor = Color.White, Padding = new Padding(15) };

            Label lblIslemler = new Label { Text = "YEDEKLEME İŞLEMLERİ", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DimGray };
            pnlSol.Controls.Add(lblIslemler);

            btnManuelYedek = new Button { Text = "⚡ Şimdi Hızlı Yedek Al", Dock = DockStyle.Top, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Margin = new Padding(0, 0, 0, 10) };
            btnManuelYedek.Click += BtnManuelYedek_Click;
            pnlSol.Controls.Add(btnManuelYedek);

            // Boşluk paneli
            pnlSol.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 15 });

            btnHariciYedek = new Button { Text = "📂 Harici Yere Yedekle (USB/Cloud)", Dock = DockStyle.Top, Height = 45, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };
            btnHariciYedek.Click += BtnHariciYedek_Click;
            pnlSol.Controls.Add(btnHariciYedek);

            pnlSol.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 40 });

            Label lblKurtarma = new Label { Text = "💥 ACİL DURUM KURTARMA", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold | FontStyle.Underline), ForeColor = Color.DarkRed };
            pnlSol.Controls.Add(lblKurtarma);

            btnGeriYukle = new Button { Text = "🔄 SEÇİLİ YEDEĞİ GERİ YÜKLE", Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            btnGeriYukle.Click += BtnGeriYukle_Click;
            pnlSol.Controls.Add(btnGeriYukle);

            pnlSol.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 15 });

            btnYedekSil = new Button { Text = "🗑️ Seçili Yedeği Kalıcı Olarak Sil", Dock = DockStyle.Top, Height = 32, BackColor = Color.FromArgb(127, 140, 141), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            btnYedekSil.Click += BtnYedekSil_Click;
            pnlSol.Controls.Add(btnYedekSil);

            this.Controls.Add(pnlSol);

            // Sağ Liste Alanı
            GroupBox grpListe = new GroupBox { Dock = DockStyle.Fill, Text = "📅 Sistemde Kayıtlı Tarihsel Yedek Listesi ( bin/Debug/Yedekler )", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Padding = new Padding(10) };
            lstYedekler = new ListBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 10F), BackColor = Color.FromArgb(248, 249, 250) };
            grpListe.Controls.Add(lstYedekler);
            this.Controls.Add(grpListe);
        }

        private void YedekleriListele()
        {
            try
            {
                lstYedekler.Items.Clear();
                string[] dosyalar = Directory.GetFiles(yedekKlasorYolu, "Yedek_*.db");

                // En güncel yedek en üstte görünsün diye diziyi ters çeviriyoruz
                Array.Reverse(dosyalar);

                foreach (string dosya in dosyalar)
                {
                    lstYedekler.Items.Add(Path.GetFileName(dosya));
                }
            }
            catch (Exception ex) { MessageBox.Show("Yedek listesi okunurken hata: " + ex.Message); }
        }

        // --- 1. İŞLEM: ŞİMDİ HIZLI YEDEK AL ---
        private void BtnManuelYedek_Click(object sender, EventArgs e)
        {
            try
            {
                string kaynakDosya = "stok.db";
                if (!File.Exists(kaynakDosya)) { MessageBox.Show("Ana veritabanı bulunamadı!"); return; }

                string yeniYedekAd = $"Yedek_{DateTime.Now:yyyy_MM_dd_HHmmss}_MANUEL.db";
                string hedefYol = Path.Combine(yedekKlasorYolu, yeniYedekAd);

                File.Copy(kaynakDosya, hedefYol, true);
                YedekleriListele();
                MessageBox.Show("Program hafızası zaman tüneline başarıyla kaydedildi!", "Yedekleme Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }

        // --- 2. İŞLEM: HARİCİ YERE YEDEKLE (USB / DRIVE) ---
        private void BtnHariciYedek_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Veritabanı Yedeği|*.db",
                FileName = $"Stok_Yedek_{DateTime.Now:yyyy_MM_dd}.db",
                Title = "Yedeği Kaydedilecek Güvenli Bölgeyi (USB/Cloud) Seçin"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    File.Copy("stok.db", sfd.FileName, true);
                    MessageBox.Show("Kritik ticari veri harici diske başarıyla klonlandı!", "Dışa Aktarım Tamam", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Harici yedekleme hatası: " + ex.Message); }
            }
        }

        // --- 3. İŞLEM: KRİTİK GERİ YÜKLEME MOTORU (RESTORE) ---
        private void BtnGeriYukle_Click(object sender, EventArgs e)
        {
            if (lstYedekler.SelectedItem == null)
            {
                MessageBox.Show("Lütfen geri yüklemek istediğiniz tarihli yedek dosyasını sağdaki listeden seçin!", "Seçim Eksik", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string seciliYedekAd = lstYedekler.SelectedItem.ToString();
            string seciliYedekTamYol = Path.Combine(yedekKlasorYolu, seciliYedekAd);

            DialogResult dr = MessageBox.Show(
                $"⚠️⚠️ DİKKAT ÇOK KRİTİK KARAR! ⚠️⚠️\n\n[{seciliYedekAd}] tarihli yedeği geri yüklemek üzeresiniz.\n" +
                "Bu işlem şu anki tüm güncel verilerinizi silecek ve dükkanı o tarihteki durumuna geri döndürecektir.\n\n" +
                "Devam etmek istediğinize EMİN MİSİNİZ?",
                "Sistem Geri Yükleme Onayı",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);

            if (dr == DialogResult.Yes)
            {
                try
                {
                    // PROFESYONEL ÖNLEM: SQLite bağlantı havuzlarını zorla boşaltıyoruz dosya kilitlenmesin 🚀
                    SqliteConnection.ClearAllPools();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // Mevcut canlı veritabanının üzerine seçilen eski yedeği çakıyoruz
                    File.Copy(seciliYedekTamYol, "stok.db", true);

                    MessageBox.Show("Sistem zaman tünelinde geriye başarıyla sarıldı! Değişikliklerin tam aktif olması için program şimdi yeniden başlatılacak.", "Geri Yükleme Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Programı yeni verilerle tertemiz açmak için baştan başlatıyoruz
                    Application.Restart();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Geri yükleme başarısız oldu! Veritabanı dosyası şu an sistem tarafından kullanılıyor olabilir.\nHata detayı: " + ex.Message, "Kilitlenme Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // --- 4. İŞLEM: YEDEK SİLME ---
        private void BtnYedekSil_Click(object sender, EventArgs e)
        {
            if (lstYedekler.SelectedItem == null) return;
            string dosyaAd = lstYedekler.SelectedItem.ToString();

            DialogResult dr = MessageBox.Show($"{dosyaAd} isimli yedek dosyasını diskten kalıcı olarak silmek istiyor musunuz?", "Yedek Dosyası Silme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    string yol = Path.Combine(yedekKlasorYolu, dosyaAd);
                    if (File.Exists(yol)) File.Delete(yol);
                    YedekleriListele();
                }
                catch (Exception ex) { MessageBox.Show("Silme hatası: " + ex.Message); }
            }
        }
    }
}