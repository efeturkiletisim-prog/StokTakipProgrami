using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public partial class SayfaKategoriMarka : UserControl
    {
        private ComboBox dynCmbAyarTuru;
        private TextBox dynTxtYeniKayit;
        private Button dynBtnEkle, dynBtnSil;
        private DataGridView dynDgvListe;
        private string aktifTablo = "Kategoriler";

        public SayfaKategoriMarka()
        {
            InitializeComponent();
            this.Load += SayfaKategoriMarka_Load;
        }

        private void SayfaKategoriMarka_Load(object sender, EventArgs e)
        {
            TablolariHazirla();
            ArayuzInsaEt();
            VerileriListele();
        }

        private void TablolariHazirla()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("CREATE TABLE IF NOT EXISTS Kategoriler (Ad TEXT PRIMARY KEY);", b)) cmd.ExecuteNonQuery();
                    using (var cmd = new SqliteCommand("CREATE TABLE IF NOT EXISTS Markalar (Ad TEXT PRIMARY KEY);", b)) cmd.ExecuteNonQuery();
                    using (var cmd = new SqliteCommand("CREATE TABLE IF NOT EXISTS Firmalar (Ad TEXT PRIMARY KEY);", b)) cmd.ExecuteNonQuery();
                }
            }
            catch { }
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "🏷️ KATEGORİ, MARKA VE TEDARİKÇİ FİRMA GİRİŞİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            Panel pnlCard = new Panel { Location = new Point(25, 85), Width = 340, Height = 360, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15) };

            Label lblTip = new Label { Text = "1. Kayıt Birimi Seçin:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynCmbAyarTuru = new ComboBox { Location = new Point(20, 45), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 11F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            dynCmbAyarTuru.Items.AddRange(new object[] { "🗂️ Kategoriler", "🏷️ Markalar", "🏢 Tedarikçi Firmalar" });
            dynCmbAyarTuru.SelectedIndex = 0;
            dynCmbAyarTuru.SelectedIndexChanged += (s, e) => {
                if (dynCmbAyarTuru.SelectedIndex == 0) aktifTablo = "Kategoriler";
                else if (dynCmbAyarTuru.SelectedIndex == 1) aktifTablo = "Markalar";
                else aktifTablo = "Firmalar";
                VerileriListele();
            };
            pnlCard.Controls.Add(lblTip); pnlCard.Controls.Add(dynCmbAyarTuru);

            Label lblAd = new Label { Text = "2. Yeni Kayıt Adı:", Location = new Point(20, 100), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynTxtYeniKayit = new TextBox { Location = new Point(20, 125), Width = 300, Font = new Font("Segoe UI", 11F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            pnlCard.Controls.Add(lblAd); pnlCard.Controls.Add(dynTxtYeniKayit);

            dynBtnEkle = new Button { Text = "➕ SİSTEME EKLE", Location = new Point(20, 185), Width = 300, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnEkle.FlatAppearance.BorderSize = 0;
            dynBtnEkle.Click += DynBtnEkle_Click;
            pnlCard.Controls.Add(dynBtnEkle);

            dynBtnSil = new Button { Text = "🗑️ SEÇİLİ OLANI SİL", Location = new Point(20, 245), Width = 300, Height = 40, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSil.FlatAppearance.BorderSize = 0;
            dynBtnSil.Click += DynBtnSil_Click;
            pnlCard.Controls.Add(dynBtnSil);

            this.Controls.Add(pnlCard);

            GroupBox grpTable = new GroupBox { Location = new Point(390, 85), Width = 560, Height = 480, Text = "📋 Sistemdeki Kayıtlar", Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Gray };
            dynDgvListe = new DataGridView
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
            grpTable.Controls.Add(dynDgvListe);
            this.Controls.Add(grpTable);
        }

        private void VerileriListele()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand($"SELECT Ad as 'Kayıtlı {aktifTablo}' FROM {aktifTablo} ORDER BY Ad ASC", b))
                    {
                        DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader());
                        dynDgvListe.DataSource = dt;
                    }
                }
            }
            catch { }
        }

        private void DynBtnEkle_Click(object sender, EventArgs e)
        {
            string veri = dynTxtYeniKayit.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(veri)) { MessageBox.Show("Lütfen bir ad girin!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand($"INSERT OR IGNORE INTO {aktifTablo} (Ad) VALUES (@v)", b))
                    {
                        cmd.Parameters.AddWithValue("@v", veri);
                        cmd.ExecuteNonQuery();
                    }
                }
                dynTxtYeniKayit.Clear(); VerileriListele();
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }

        private void DynBtnSil_Click(object sender, EventArgs e)
        {
            if (dynDgvListe.SelectedRows.Count == 0) return;
            string secili = dynDgvListe.SelectedRows[0].Cells[0].Value.ToString();

            if (MessageBox.Show($"'{secili}' kaydını kalıcı olarak silmek istiyor musunuz?", "Silme Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var cmd = new SqliteCommand($"DELETE FROM {aktifTablo} WHERE Ad=@v", b))
                        {
                            cmd.Parameters.AddWithValue("@v", secili); cmd.ExecuteNonQuery();
                        }
                    }
                    VerileriListele();
                }
                catch { }
            }
        }
    }
}