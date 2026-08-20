using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public partial class SayfaPlatformEkle : UserControl
    {
        private TextBox dynTxtPlatformAdi, dynTxtHizmetBedeli;
        private Button dynBtnKaydet;

        public SayfaPlatformEkle()
        {
            InitializeComponent();
            this.Load += SayfaPlatformEkle_Load;
        }

        private void SayfaPlatformEkle_Load(object sender, EventArgs e)
        {
            ArayuzInsaEt();
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(245, 246, 250);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.White };
            Label lblT = new Label { Text = "🏢  YENİ PAZARYERİ / SATIŞ KANALI GİRİŞİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.FromArgb(41, 41, 44), TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            Panel pnlCard = new Panel { Location = new Point(25, 85), Width = 400, Height = 320, BackColor = Color.White, Padding = new Padding(30) };

            pnlCard.Controls.Add(FormElemaniYarat("Platform / Mağaza Adı (Örn: TRENDYOL)", out dynTxtPlatformAdi, new Point(30, 30), 340));
            pnlCard.Controls.Add(FormElemaniYarat("Varsayılan Sabit Komisyon Oranı (%)", out dynTxtHizmetBedeli, new Point(30, 115), 340));
            dynTxtHizmetBedeli.Text = "0";

            dynBtnKaydet = new Button { Text = "💼 PLATFORMU SİSTEME EKLE", Location = new Point(30, 210), Width = 340, Height = 48, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnKaydet.Click += dynBtnKaydet_Click;
            pnlCard.Controls.Add(dynBtnKaydet);

            this.Controls.Add(pnlCard);
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 65 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Gray };
            kutu = new TextBox { Location = new Point(0, 22), Width = genislik, Font = new Font("Segoe UI", 10.5F), BorderStyle = BorderStyle.FixedSingle };
            p.Controls.Add(lbl); p.Controls.Add(kutu);
            return p;
        }

        private void dynBtnKaydet_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(dynTxtPlatformAdi.Text)) return;
            double.TryParse(dynTxtHizmetBedeli.Text.Replace('.', ','), out double bedel);

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = "INSERT INTO Platformlar (Ad, HizmetBedeli) VALUES (@ad, @bedel) ON CONFLICT(Ad) DO UPDATE SET HizmetBedeli=@bedel";
                    using (var cmd = new SqliteCommand(sql, b))
                    {
                        cmd.Parameters.AddWithValue("@ad", dynTxtPlatformAdi.Text.Trim().ToUpper());
                        cmd.Parameters.AddWithValue("@bedel", bedel);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Satış kanalı başarıyla kaydedildi!", "Kayıt Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dynTxtPlatformAdi.Clear(); dynTxtHizmetBedeli.Text = "0";
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }
    }
}