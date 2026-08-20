using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public partial class SayfaKullanicilar : UserControl
    {
        private TextBox dynTxtKullanici, dynTxtSifre;
        private ComboBox dynCmbRol;
        private Button dynBtnKaydet, dynBtnSil;
        private DataGridView dynDgvKullanicilar;

        public SayfaKullanicilar()
        {
            InitializeComponent();
            this.Load += SayfaKullanicilar_Load;
        }

        private void SayfaKullanicilar_Load(object sender, EventArgs e)
        {
            TabloyuGuvenceyeAl();
            ArayuzInsaEt();
            KullanicilariListele();
        }

        private void TabloyuGuvenceyeAl()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    // Kullanıcılar tablosu yoksa saniyede oluşturur
                    string sql = @"CREATE TABLE IF NOT EXISTS Kullanicilar (
                                    KullaniciAdi TEXT PRIMARY KEY,
                                    Sifre TEXT,
                                    Rol TEXT
                                  );";
                    using (var cmd = new SqliteCommand(sql, b)) { cmd.ExecuteNonQuery(); }

                    // Eğer tablo boşsa seni otomatik olarak ilk kurucu (YÖNETİCİ) yapar
                    using (var cmd = new SqliteCommand("SELECT COUNT(*) FROM Kullanicilar", b))
                    {
                        long sayi = (long)cmd.ExecuteScalar();
                        if (sayi == 0)
                        {
                            using (var cmdIns = new SqliteCommand("INSERT INTO Kullanicilar (KullaniciAdi, Sifre, Rol) VALUES ('PATRON', '1234', 'YÖNETİCİ')", b))
                            {
                                cmdIns.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(33, 33, 36); // Lüks mat antrasit tema

            // 1. ÜST BAŞLIK
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "👥  SİSTEM ERİŞİM VE PERSONEL YETKİLENDİRME MERKEZİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            // 2. SOL TARAF: KULLANICI EKLEME/GÜNCELLEME KARTI
            Panel pnlCard = new Panel { Location = new Point(25, 85), Width = 340, Height = 450, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(20) };

            pnlCard.Controls.Add(FormElemaniYarat("1. Kullanıcı / Personel Adı:", out dynTxtKullanici, new Point(20, 20), 300));
            pnlCard.Controls.Add(FormElemaniYarat("2. Sisteme Giriş Şifresi:", out dynTxtSifre, new Point(20, 100), 300));

            Label lblR = new Label { Text = "3. Sistem Yetkisi (Rol):", Location = new Point(20, 180), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynCmbRol = new ComboBox { Location = new Point(20, 205), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            dynCmbRol.Items.AddRange(new object[] { "YÖNETİCİ", "PERSONEL" });
            dynCmbRol.SelectedIndex = 1;
            pnlCard.Controls.Add(lblR); pnlCard.Controls.Add(dynCmbRol);

            dynBtnKaydet = new Button
            {
                Text = "💾 KULLANICIYI KAYDET",
                Location = new Point(20, 280),
                Width = 300,
                Height = 45,
                BackColor = Color.FromArgb(46, 117, 89),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            dynBtnKaydet.FlatAppearance.BorderSize = 0;
            dynBtnKaydet.Click += DynBtnKaydet_Click;
            pnlCard.Controls.Add(dynBtnKaydet);

            dynBtnSil = new Button
            {
                Text = "🗑️ SEÇİLİ KULLANICIYI SİL",
                Location = new Point(20, 335),
                Width = 300,
                Height = 40,
                BackColor = Color.FromArgb(192, 57, 43),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            dynBtnSil.FlatAppearance.BorderSize = 0;
            dynBtnSil.Click += DynBtnSil_Click;
            pnlCard.Controls.Add(dynBtnSil);

            this.Controls.Add(pnlCard);

            // 3. SAĞ TARAF: MEVCUT KULLANICILAR TABLOSU
            GroupBox grpTable = new GroupBox { Location = new Point(390, 85), Width = 600, Height = 450, Text = "📊 Sisteme Kayıtlı Mevcut Personeller", Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Gray };
            dynDgvKullanicilar = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(26, 26, 28),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 10F),
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, SelectionBackColor = Color.FromArgb(0, 120, 215) },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(21, 21, 33), ForeColor = Color.White, Font = new Font("Segoe UI", 9F, FontStyle.Bold) },
                EnableHeadersVisualStyles = false
            };
            dynDgvKullanicilar.CellClick += DynDgvKullanicilar_CellClick;
            grpTable.Controls.Add(dynDgvKullanicilar);
            this.Controls.Add(grpTable);
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 65 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            kutu = new TextBox { Location = new Point(0, 25), Width = genislik, Font = new Font("Segoe UI", 11F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };
            p.Controls.Add(lbl); p.Controls.Add(kutu);
            return p;
        }

        private void KullanicilariListele()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = "SELECT KullaniciAdi as 'Kullanıcı Adı', Sifre as 'Şifre', Rol as 'Sistem Yetkisi' FROM Kullanicilar ORDER BY Rol DESC";
                    using (var cmd = new SqliteCommand(sql, b))
                    {
                        DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader());
                        dynDgvKullanicilar.DataSource = dt;
                    }
                }
            }
            catch { }
        }

        private void DynDgvKullanicilar_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dynTxtKullanici.Text = dynDgvKullanicilar.Rows[e.RowIndex].Cells["Kullanıcı Adı"].Value.ToString();
                dynTxtSifre.Text = dynDgvKullanicilar.Rows[e.RowIndex].Cells["Şifre"].Value.ToString();
                dynCmbRol.Text = dynDgvKullanicilar.Rows[e.RowIndex].Cells["Sistem Yetkisi"].Value.ToString();
            }
        }

        private void DynBtnKaydet_Click(object sender, EventArgs e)
        {
            string kadi = dynTxtKullanici.Text.Trim().ToUpper();
            string sifre = dynTxtSifre.Text.Trim();
            string rol = dynCmbRol.Text;

            if (string.IsNullOrWhiteSpace(kadi) || string.IsNullOrWhiteSpace(sifre))
            {
                MessageBox.Show("Kullanıcı adı ve şifre boş bırakılamaz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = "INSERT INTO Kullanicilar (KullaniciAdi, Sifre, Rol) VALUES (@k, @s, @r) ON CONFLICT(KullaniciAdi) DO UPDATE SET Sifre=@s, Rol=@r";
                    using (var cmd = new SqliteCommand(sql, b))
                    {
                        cmd.Parameters.AddWithValue("@k", kadi);
                        cmd.Parameters.AddWithValue("@s", sifre);
                        cmd.Parameters.AddWithValue("@r", rol);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Kullanıcı sisteme başarıyla kaydedildi!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dynTxtKullanici.Clear(); dynTxtSifre.Clear();
                KullanicilariListele();
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }

        private void DynBtnSil_Click(object sender, EventArgs e)
        {
            if (dynDgvKullanicilar.SelectedRows.Count == 0) return;
            string secili = dynDgvKullanicilar.SelectedRows[0].Cells["Kullanıcı Adı"].Value.ToString();

            if (secili == Veritabanı.AktifKullanici)
            {
                MessageBox.Show("Şu an sisteme giriş yaptığınız kendi hesabınızı silemezsiniz!", "Güvenlik Engeli", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;
            }

            if (MessageBox.Show($"'{secili}' adlı kullanıcının sisteme erişimini kalıcı olarak iptal etmek istiyor musunuz?", "Kullanıcı Sil", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var cmd = new SqliteCommand("DELETE FROM Kullanicilar WHERE KullaniciAdi=@k", b))
                        {
                            cmd.Parameters.AddWithValue("@k", secili);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    dynTxtKullanici.Clear(); dynTxtSifre.Clear();
                    KullanicilariListele();
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }
        }
    }
}