using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public class SayfaGelismisStokSorgula : UserControl
    {
        private TextBox txtArama;
        private Button btnAra;
        private DataGridView dgvStoklar;

        public SayfaGelismisStokSorgula()
        {
            ArayuzElemanlariniInsaEt();
            StoklariSorgula("");
        }

        private void ArayuzElemanlariniInsaEt()
        {
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(33, 33, 36);
            this.Padding = new Padding(20);

            Label lblBaslik = new Label
            {
                Text = "🔍 STOK / ÜRÜN SORGULAMA",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblBaslik);

            Panel pnlAramaBar = new Panel { Dock = DockStyle.Top, Height = 60 };

            Label lblBilgi = new Label
            {
                Text = "Barkod Okutun veya Ürün Adı Yazın:",
                Location = new Point(0, 15),
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.LightGray
            };
            pnlAramaBar.Controls.Add(lblBilgi);

            txtArama = new TextBox
            {
                Location = new Point(250, 12),
                Width = 350,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.FromArgb(56, 189, 248),
                BorderStyle = BorderStyle.FixedSingle
            };
            txtArama.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; StoklariSorgula(txtArama.Text.Trim()); } };
            pnlAramaBar.Controls.Add(txtArama);

            btnAra = new Button
            {
                Text = "ÜRÜNÜ SORGULA",
                Location = new Point(610, 11),
                Width = 140,
                Height = 30,
                BackColor = Color.FromArgb(46, 117, 89),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnAra.FlatAppearance.BorderSize = 0;
            btnAra.Click += (s, e) => StoklariSorgula(txtArama.Text.Trim());
            pnlAramaBar.Controls.Add(btnAra);

            this.Controls.Add(pnlAramaBar);

            dgvStoklar = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(24, 24, 27),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 10F),
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,

                // ✅ ÇÖZÜM: Height özelliği stilin içinden çıkartılıp doğrudan tabloya mühürlendi
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 40,

                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, SelectionBackColor = Color.FromArgb(56, 189, 248), SelectionForeColor = Color.Black },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(20, 20, 22), ForeColor = Color.White, Font = new Font("Segoe UI", 10F, FontStyle.Bold) }
            };
            this.Controls.Add(dgvStoklar);
            dgvStoklar.BringToFront();
        }

        private void StoklariSorgula(string aramaMetni)
        {
            try
            {
                // Veritabanı dosya adını (stok.db) kendi orijinal adına göre revize edebilirsin patron
                using (var baglanti = new SqliteConnection("Data Source=stok.db"))
                {
                    baglanti.Open();
                    string sorgu = "SELECT * FROM Urunler"; // Tablo adın neyse o (Urunler / Stoklar)

                    if (!string.IsNullOrEmpty(aramaMetni))
                    {
                        sorgu += " WHERE Barkod LIKE @ara OR UrunAdi LIKE @ara";
                    }

                    using (var cmd = new SqliteCommand(sorgu, baglanti))
                    {
                        if (!string.IsNullOrEmpty(aramaMetni))
                        {
                            cmd.Parameters.AddWithValue("@ara", "%" + aramaMetni + "%");
                        }

                        // ✅ ÇÖZÜM: Bulunamayan SqliteDataAdapter yerine en hızlı DataReader motoru bağlandı
                        using (var reader = cmd.ExecuteReader())
                        {
                            DataTable dt = new DataTable();
                            dt.Load(reader);
                            dgvStoklar.DataSource = dt;
                        }
                    }
                }
            }
            catch { }
        }
    }
}