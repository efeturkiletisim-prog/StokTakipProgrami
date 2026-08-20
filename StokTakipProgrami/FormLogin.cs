using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public partial class FormLogin : Form
    {
        // Dinamik Premium UI Elemanları
        private TextBox dynTxtKullanici, dynTxtSifre;
        private Button dynBtnGiris, dynBtnCikis;

        // Pencere sürükleme değişkenleri
        private bool surukleniyor = false;
        private Point surukleBaslangicNoktasi;

        public FormLogin()
        {
            InitializeComponent();
            this.Load += FormLogin_Load;
        }

        private void FormLogin_Load(object sender, EventArgs e)
        {
            Veritabanı.Kur();
            PremiumGirisArayuzuInsaEt();
        }

        private void PremiumGirisArayuzuInsaEt()
        {
            this.Controls.Clear();

            // --- 🏢 FORM PANEL GENEL AYARLARI ---
            this.Text = "EFETURK.NET - Güvenli Giriş Paneli";
            this.Size = new Size(420, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.FromArgb(33, 33, 36);

            // Formun arka planından tutulup sürüklenmesi için olay köprüleri
            this.MouseDown += Form_MouseDown;
            this.MouseMove += Form_MouseMove;
            this.MouseUp += Form_MouseUp;

            // 1. ÜST MARKA LOGO / BAŞLIK ALANI
            Label lblBrand = new Label
            {
                Text = "EFETURK.NET",
                Location = new Point(0, 35),
                Width = 420,
                Height = 40,
                Font = new Font("Segoe UI", 20F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 215),
                TextAlign = ContentAlignment.MiddleCenter
            };
            // Logodan tutunca da sürüklensin
            lblBrand.MouseDown += Form_MouseDown;
            lblBrand.MouseMove += Form_MouseMove;
            lblBrand.MouseUp += Form_MouseUp;
            this.Controls.Add(lblBrand);

            Label lblSubBrand = new Label
            {
                Text = "BULUT ENVANTER KOMUTA MERKEZİ",
                Location = new Point(0, 75),
                Width = 420,
                Height = 20,
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
            // Alt logodan tutunca da sürüklensin
            lblSubBrand.MouseDown += Form_MouseDown;
            lblSubBrand.MouseMove += Form_MouseMove;
            lblSubBrand.MouseUp += Form_MouseUp;
            this.Controls.Add(lblSubBrand);

            // 2. MERKEZİ GİRİŞ KARTI
            Panel pnlCard = new Panel
            {
                Location = new Point(35, 120),
                Width = 350,
                Height = 340,
                BackColor = Color.FromArgb(26, 26, 28),
                Padding = new Padding(20)
            };
            this.Controls.Add(pnlCard);

            // GİRİŞ ELEMANLARI DİZİLİMİ
            pnlCard.Controls.Add(GirisElemaniYarat("👤  Kullanıcı Adı veya Personel Kodu:", out dynTxtKullanici, new Point(25, 30), 300, false));
            pnlCard.Controls.Add(GirisElemaniYarat("🔒  Şifre:", out dynTxtSifre, new Point(25, 110), 300, true));

            // 3. LÜKS GİRİŞ YAP BUTONU
            dynBtnGiris = new Button
            {
                Text = "🔒  GİRİŞ YAP",
                Location = new Point(25, 205),
                Width = 300,
                Height = 45,
                BackColor = Color.FromArgb(46, 117, 89),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            dynBtnGiris.FlatAppearance.BorderSize = 0;
            dynBtnGiris.Click += DynBtnGiris_Click;
            pnlCard.Controls.Add(dynBtnGiris);

            // 4. SADELEŞTİRİLMİŞ ÇIKIŞ YAP BUTONU
            dynBtnCikis = new Button
            {
                Text = "ÇIKIŞ YAP", // Metin tam istediğin gibi sadeleştirildi!
                Location = new Point(25, 265),
                Width = 300,
                Height = 40,
                BackColor = Color.FromArgb(53, 53, 58),
                ForeColor = Color.FromArgb(220, 76, 64),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            dynBtnCikis.FlatAppearance.BorderSize = 1;
            dynBtnCikis.FlatAppearance.BorderColor = Color.FromArgb(192, 57, 43);
            dynBtnCikis.Click += DynBtnCikis_Click;
            pnlCard.Controls.Add(dynBtnCikis);
        }

        private Panel GirisElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, bool sifreMi)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 65 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(140, 140, 145) };

            kutu = new TextBox
            {
                Location = new Point(0, 23),
                Width = genislik,
                Font = new Font("Segoe UI", 11F),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(41, 41, 44),
                ForeColor = Color.White
            };

            if (sifreMi) kutu.UseSystemPasswordChar = true;

            p.Controls.Add(lbl); p.Controls.Add(kutu);
            return p;
        }

        // --- 🖱️ HAREKET ETTİRME (SÜRÜKLEME) MOTORU ---
        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                surukleniyor = true;
                surukleBaslangicNoktasi = e.Location;
            }
        }

        private void Form_MouseMove(object sender, MouseEventArgs e)
        {
            if (surukleniyor)
            {
                this.Location = new Point(
                    (this.Location.X - surukleBaslangicNoktasi.X) + e.X,
                    (this.Location.Y - surukleBaslangicNoktasi.Y) + e.Y);
                this.Update();
            }
        }

        private void Form_MouseUp(object sender, MouseEventArgs e)
        {
            surukleniyor = false;
        }

        // --- 🔒 GİRİŞ YAPMA MOTORU ---
        private void DynBtnGiris_Click(object sender, EventArgs e)
        {
            string kadi = dynTxtKullanici.Text.Trim().ToLower();
            string sifre = dynTxtSifre.Text.Trim();

            if (string.IsNullOrWhiteSpace(kadi) || string.IsNullOrWhiteSpace(sifre)) return;

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("SELECT Rol FROM Kullanicilar WHERE KullaniciAdi=@k AND Sifre=@s", b))
                    {
                        cmd.Parameters.AddWithValue("@k", kadi);
                        cmd.Parameters.AddWithValue("@s", sifre);
                        object rol = cmd.ExecuteScalar();

                        if (rol != null)
                        {
                            Veritabanı.AktifKullanici = kadi.ToUpper();
                            Veritabanı.AktifRol = rol.ToString().ToUpper();

                            this.Hide();
                            Form1 anaEkran = new Form1();
                            anaEkran.ShowDialog();
                            this.Close();
                        }
                        else
                        {
                            MessageBox.Show("Hatalı Kullanıcı Adı veya Şifre girdiniz!", "Erişim Engellendi", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Giriş Hatası: " + ex.Message); }
        }

        // --- 🚪 ÇIKIŞ YAP MOTORU ---
        private void DynBtnCikis_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}