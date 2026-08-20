using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // 💸 SİPARİŞ CEZA KESİNTİLERİ VE PAZARYERİ MUTABAKAT DENETİM MERKEZİ
    public partial class SayfaSiparisCezalari : UserControl
    {
        // Sol Panel Kontrolleri
        private TextBox dynTxtSiparisNo, dynTxtCezaTutari, dynTxtUrunAdi, dynTxtPlatform, dynTxtSatisFiyati;
        private ComboBox dynCmbCezaSebebi;
        private Button dynBtnSiparisSorgula, dynBtnCezayiMuhurle;

        // Sağ Panel Kontrolleri
        private DataGridView dynDgvCezaliSiparisler;
        private Button dynBtnCezaIptal, dynBtnExcel;

        public SayfaSiparisCezalari()
        {
            InitializeComponent();
            this.Load += SayfaSiparisCezalari_Load;
        }

        private void SayfaSiparisCezalari_Load(object sender, EventArgs e)
        {
            TabloyuGuvenceyeAl();
            ArayuzInsaEt();
            CezaliSiparisleriListele();
        }

        // 🚀 MÜHÜR: Veritabanını Canlı Canlı Güncelleyen Kalkan Motoru
        private void TabloyuGuvenceyeAl()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    var sutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (var cmd = new SqliteCommand("PRAGMA table_info('Siparisler');", conn))
                    using (var r = cmd.ExecuteReader()) { while (r.Read()) sutunlar.Add(r["name"].ToString()); }

                    // Eğer sütunlar yoksa verileri bozmadan canlı olarak ekler
                    if (!sutunlar.Contains("CezaTutari"))
                    {
                        using (var cmd = new SqliteCommand("ALTER TABLE Siparisler ADD COLUMN CezaTutari REAL DEFAULT 0;", conn)) cmd.ExecuteNonQuery();
                    }
                    if (!sutunlar.Contains("CezaSebebi"))
                    {
                        using (var cmd = new SqliteCommand("ALTER TABLE Siparisler ADD COLUMN CezaSebebi TEXT;", conn)) cmd.ExecuteNonQuery();
                    }
                }
            }
            catch { }
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(33, 33, 36);

            // Üst Başlık Şeridi
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "⚠️ PAZARYERİ SİPARİŞ CEZA VE MAHKUMİYET KESİNTİLERİ GİRİŞ TERMİNALİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT);
            this.Controls.Add(pnlHeader);

            // Ana Bölücü Responsive Grid Layout
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380F)); // Giriş Formu Sabit 380px
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.Controls.Add(mainLayout);

            // ================== SOL GİRİŞ PANELİ ==================
            Panel pnlForm = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15), AutoScroll = true };

            Label lblSip = new Label { Text = "1. Sorgulanacak Sipariş Numarası:", Location = new Point(15, 15), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynTxtSiparisNo = new TextBox { Location = new Point(15, 37), Width = 210, Font = new Font("Segoe UI", 11F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            dynBtnSiparisSorgula = new Button { Text = "🔍 SORGULA", Location = new Point(235, 36), Width = 120, Height = 25, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSiparisSorgula.FlatAppearance.BorderSize = 0;
            dynBtnSiparisSorgula.Click += DynBtnSiparisSorgula_Click;
            pnlForm.Controls.Add(lblSip); pnlForm.Controls.Add(dynTxtSiparisNo); pnlForm.Controls.Add(dynBtnSiparisSorgula);

            pnlForm.Controls.Add(BilgiElemaniYarat("Eşleşen Ürün Modeli:", out dynTxtUrunAdi, new Point(15, 75), 340));
            pnlForm.Controls.Add(BilgiElemaniYarat("Satış Kanalı / Platform:", out dynTxtPlatform, new Point(15, 135), 340));
            pnlForm.Controls.Add(BilgiElemaniYarat("Brüt Satış Fiyatı (TL):", out dynTxtSatisFiyati, new Point(15, 195), 340));

            Label lblSebep = new Label { Text = "2. Kesilen Cezanın Sebebi:", Location = new Point(15, 260), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            dynCmbCezaSebebi = new ComboBox { Location = new Point(15, 282), Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            dynCmbCezaSebebi.Items.AddRange(new object[] { "TEDARİK EDEMEME", "YANLIŞ ÜRÜN", "GEÇ KARGOLAMA" });
            dynCmbCezaSebebi.SelectedIndex = 0;
            pnlForm.Controls.Add(lblSebep); pnlForm.Controls.Add(dynCmbCezaSebebi);

            pnlForm.Controls.Add(GirisElemaniYarat("3. Kesilen Ceza Tutarı (TL):", out dynTxtCezaTutari, new Point(15, 325), 340));

            dynBtnCezayiMuhurle = new Button { Text = "⚠️ CEZA KESİNTİSİNİ SİSTEME MÜHÜRLE", Location = new Point(15, 405), Width = 340, Height = 48, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnCezayiMuhurle.FlatAppearance.BorderSize = 0;
            dynBtnCezayiMuhurle.Click += DynBtnCezayiMuhurle_Click;
            pnlForm.Controls.Add(dynBtnCezayiMuhurle);

            mainLayout.Controls.Add(pnlForm, 0, 0);

            // ================== SAĞ LİSTE PANELİ ==================
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12) };
            Label lblTabloBaslik = new Label { Text = "📋 PAZARYERLERİ AKTİF CEZA VE KESİNTİ MATRİSİ LİSTESİ", Dock = DockStyle.Top, Height = 22, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            pnlRight.Controls.Add(lblTabloBaslik);

            dynDgvCezaliSiparisler = GridYarat();
            pnlRight.Controls.Add(dynDgvCezaliSiparisler);
            dynDgvCezaliSiparisler.BringToFront();

            Panel pnlRightBottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(0, 5, 0, 0) };
            dynBtnCezaIptal = new Button { Text = "🗑️ SEÇİLİ CEZA KAYDINI İPTAL ET (Kârlılığı Eski Haline Döndür)", Dock = DockStyle.Left, Width = 380, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnCezaIptal.FlatAppearance.BorderSize = 0;
            dynBtnCezaIptal.Click += DynBtnCezaIptal_Click;
            pnlRightBottom.Controls.Add(dynBtnCezaIptal);

            dynBtnExcel = new Button { Text = "📊 Excel Çıktısı", Dock = DockStyle.Right, Width = 130, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnExcel.FlatAppearance.BorderSize = 0;
            dynBtnExcel.Click += (s, e) => ExcelMotoru.DataGridViewTablosunuDisaAktar(dynDgvCezaliSiparisler, "Pazaryeri_Siparis_Cezalari");
            pnlRightBottom.Controls.Add(dynBtnExcel);

            pnlRight.Controls.Add(pnlRightBottom);
            mainLayout.Controls.Add(pnlRight, 1, 0);

            mainLayout.BringToFront();
            pnlHeader.SendToBack();
        }

        private void DynBtnSiparisSorgula_Click(object sender, EventArgs e)
        {
            string arananNo = dynTxtSiparisNo.Text.Trim();
            if (string.IsNullOrWhiteSpace(arananNo)) return;

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = "SELECT PlatformAd, SatisFiyati, IFNULL(UrunAdi, 'Bilinmeyen Ürün') as Urun, CezaTutari, CezaSebebi FROM Siparisler WHERE SiparisNo = @sn LIMIT 1";
                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@sn", arananNo);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                dynTxtUrunAdi.Text = r["Urun"].ToString();
                                dynTxtPlatform.Text = r["PlatformAd"].ToString();
                                dynTxtSatisFiyati.Text = Convert.ToDouble(r["SatisFiyati"]).ToString("N2") + " TL";
                                dynTxtCezaTutari.Text = Convert.ToDouble(r["CezaTutari"]) > 0 ? Convert.ToDouble(r["CezaTutari"]).ToString("N2") : "";
                                if (r["CezaSebebi"] != DBNull.Value && dynCmbCezaSebebi.Items.Contains(r["CezaSebebi"].ToString()))
                                    dynCmbCezaSebebi.SelectedItem = r["CezaSebebi"].ToString();
                            }
                            else
                            {
                                MessageBox.Show("Sistemde bu sipariş numarasına ait hiçbir veri bulunamadı patron!", "Kayıt Yok", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                FormuTemizle(false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }

        private void DynBtnCezayiMuhurle_Click(object sender, EventArgs e)
        {
            string sipNo = dynTxtSiparisNo.Text.Trim();
            if (string.IsNullOrWhiteSpace(sipNo) || string.IsNullOrWhiteSpace(dynTxtUrunAdi.Text))
            {
                MessageBox.Show("Lütfen önce geçerli bir sipariş numarasını sorgulayın patron!", "Hatalı İşlem", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double.TryParse(dynTxtCezaTutari.Text.Replace(".", "").Replace(",", "."), out double ceza);
            if (ceza <= 0) { MessageBox.Show("Lütfen 0'dan büyük geçerli bir ceza tutarı yazın!", "Eksik Veri"); return; }
            string sebep = dynCmbCezaSebebi.Text;

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = "UPDATE Siparisler SET CezaTutari = @ceza, CezaSebebi = @sebep WHERE SiparisNo = @sn";
                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ceza", ceza);
                        cmd.Parameters.AddWithValue("@sebep", sebep);
                        cmd.Parameters.AddWithValue("@sn", sipNo);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show($"⚠️ Ceza başarıyla mühürlendi patron! {ceza:N2} TL tutar, ilgili siparişin kârlılık matematiğinden otomatik düşülecektir.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                FormuTemizle(true);
                CezaliSiparisleriListele();
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }

        private void DynBtnCezaIptal_Click(object sender, EventArgs e)
        {
            if (dynDgvCezaliSiparisler.SelectedRows.Count == 0) return;
            string sipNo = dynDgvCezaliSiparisler.SelectedRows[0].Cells["Sipariş No"].Value.ToString();

            if (MessageBox.Show($"[{sipNo}] numaralı siparişe ait ceza kaydını kalıcı olarak silmek istiyor musunuz?", "Ceza İptal Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        conn.Open();
                        string sql = "UPDATE Siparisler SET CezaTutari = 0, CezaSebebi = NULL WHERE SiparisNo = @sn";
                        using (var cmd = new SqliteCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@sn", sipNo);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    CezaliSiparisleriListele();
                    FormuTemizle(true);
                }
                catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
            }
        }

        private void CezaliSiparisleriListele()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT Tarih, PlatformAd as 'Platform', SiparisNo as 'Sipariş No', 
                               IFNULL(UrunAdi, 'Bilinmeyen Ürün') as 'Ürün Modeli', 
                               CezaSebebi as 'Ceza Nedeni', CezaTutari as 'Ceza Tutarı (TL)' 
                        FROM Siparisler 
                        WHERE CezaTutari > 0 
                        ORDER BY rowid DESC";
                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader());
                        dynDgvCezaliSiparisler.DataSource = dt;
                        if (dynDgvCezaliSiparisler.Columns["Ceza Tutarı (TL)"] != null)
                        {
                            dynDgvCezaliSiparisler.Columns["Ceza Tutarı (TL)"].DefaultCellStyle.Format = "N2";
                            dynDgvCezaliSiparisler.Columns["Ceza Tutarı (TL)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        }
                    }
                }
            }
            catch { }
        }

        private void FormuTemizle(bool herseyiTemizle)
        {
            if (herseyiTemizle) dynTxtSiparisNo.Clear();
            dynTxtUrunAdi.Clear(); dynTxtPlatform.Clear(); dynTxtSatisFiyati.Clear(); dynTxtCezaTutari.Clear();
            dynCmbCezaSebebi.SelectedIndex = 0;
        }

        private DataGridView GridYarat()
        {
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(24, 24, 27),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9F),
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 35
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgv.DataError += (s, e) => { e.Cancel = true; };
            return dgv;
        }

        private Panel BilgiElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 55 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Regular), ForeColor = Color.Gray };
            kutu = new TextBox { Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(26, 26, 28), ForeColor = Color.Coral, ReadOnly = true };
            p.Controls.Add(lbl); p.Controls.Add(kutu);
            return p;
        }

        private Panel GirisElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 55 };
            Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray };
            kutu = new TextBox { Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10.5F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };
            p.Controls.Add(lbl); p.Controls.Add(kutu);
            return p;
        }
    }
}