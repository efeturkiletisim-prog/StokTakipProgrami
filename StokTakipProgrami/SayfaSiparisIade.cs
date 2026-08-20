using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public partial class SayfaSiparisIade : UserControl
    {
        private TextBox dynTxtSiparisNo, dynTxtBarkod, dynTxtAdet, dynTxtNedeni;
        private Button dynBtnIadeYap;
        private DataGridView dynDgvIadeler;

        public SayfaSiparisIade()
        {
            InitializeComponent();
            this.Load += SayfaSiparisIade_Load;
        }

        private void SayfaSiparisIade_Load(object sender, EventArgs e)
        {
            ArayuzInsaEt();
            IadeleriListele();
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(245, 246, 250);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.White };
            Label lblT = new Label { Text = "↩️  SİPARİŞ İADE GİRİŞİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.FromArgb(41, 41, 44), TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            // FORM KARTI
            Panel pnlCard = new Panel { Location = new Point(25, 85), Width = 320, Height = 440, BackColor = Color.White, Padding = new Padding(20) };
            pnlCard.Controls.Add(FormElemaniYarat("Sipariş Numarası:", out dynTxtSiparisNo, new Point(20, 20), 280));
            pnlCard.Controls.Add(FormElemaniYarat("Ürün Barkodu:", out dynTxtBarkod, new Point(20, 95), 280));
            pnlCard.Controls.Add(FormElemaniYarat("İade Edilen Adet:", out dynTxtAdet, new Point(20, 170), 280));
            dynTxtAdet.Text = "1";
            pnlCard.Controls.Add(FormElemaniYarat("İade Nedeni / Açıklama:", out dynTxtNedeni, new Point(20, 245), 280));

            dynBtnIadeYap = new Button { Text = "🔄 İADEYİ SİSTEME İŞLE", Location = new Point(20, 340), Width = 280, Height = 45, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnIadeYap.Click += dynBtnIadeYap_Click;
            pnlCard.Controls.Add(dynBtnIadeYap);
            this.Controls.Add(pnlCard);

            // SAĞ TABLO KARTI
            GroupBox grpTable = new GroupBox { Location = new Point(370, 85), Width = 500, Height = 440, Text = "📋 Yakın Zamandaki İade Hareketleri", Font = new Font("Segoe UI", 9F, FontStyle.Bold), BackColor = Color.White };
            dynDgvIadeler = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.White, RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9F) };
            grpTable.Controls.Add(dynDgvIadeler);
            this.Controls.Add(grpTable);
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 65 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Gray };
            kutu = new TextBox { Location = new Point(0, 22), Width = genislik, Font = new Font("Segoe UI", 10F), BorderStyle = BorderStyle.FixedSingle };
            p.Controls.Add(lbl); p.Controls.Add(kutu);
            return p;
        }

        private void IadeleriListele()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    string sql = "SELECT Tarih, Barkod, Miktar as 'Adet', Aciklama as 'Gerekçe' FROM StokHareketleri WHERE IslemTuru='İADE' ORDER BY Id DESC";
                    using (var cmd = new SqliteCommand(sql, b))
                    {
                        DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader());
                        dynDgvIadeler.DataSource = dt;
                    }
                }
            }
            catch { }
        }

        private void dynBtnIadeYap_Click(object sender, System.EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(dynTxtBarkod.Text) || string.IsNullOrWhiteSpace(dynTxtSiparisNo.Text)) return;

            string barkod = dynTxtBarkod.Text.Trim();
            int.TryParse(dynTxtAdet.Text, out int iadeAdet);

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var tx = b.BeginTransaction())
                    {
                        try
                        {
                            // 1. Stoğu Geri Artır
                            using (var cmd = new SqliteCommand("UPDATE Stoklar SET Adet = Adet + @a WHERE Barkod = @b", b, tx))
                            {
                                cmd.Parameters.AddWithValue("@a", iadeAdet); cmd.Parameters.AddWithValue("@b", barkod);
                                cmd.ExecuteNonQuery();
                            }
                            // 2. Sipariş Durumunu Güncelle
                            using (var cmdS = new SqliteCommand("UPDATE Siparisler SET Durum = 'İADE EDİLDİ' WHERE SiparisNo = @sn AND Barkod = @b", b, tx))
                            {
                                cmdS.Parameters.AddWithValue("@sn", dynTxtSiparisNo.Text.Trim()); cmdS.Parameters.AddWithValue("@b", barkod);
                                cmdS.ExecuteNonQuery();
                            }
                            // 📜 3. DEFTERE MÜHÜRLE
                            Veritabanı.HareketKaydet(barkod, "İADE", iadeAdet, $"İade Gerekçesi: {dynTxtNedeni.Text} (Sip No: {dynTxtSiparisNo.Text})", b, tx);

                            tx.Commit();
                            MessageBox.Show("İade başarıyla tamamlandı ve stok güncellendi!", "İade Tamam", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            dynTxtBarkod.Clear(); dynTxtSiparisNo.Clear(); dynTxtNedeni.Clear();
                            IadeleriListele();
                        }
                        catch { tx.Rollback(); throw; }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }
    }
}