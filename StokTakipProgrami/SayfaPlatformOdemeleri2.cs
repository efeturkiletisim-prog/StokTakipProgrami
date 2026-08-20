using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace StokTakipProgrami
{
    public partial class SayfaPlatformOdemeleri2 : UserControl
    {
        private string secilenPlatform = "";
        private DateTimePicker dynDtTarih;
        private TextBox dynTxtNetGelenPara, dynTxtSiparisListesi;
        private DataGridView dynDgvTahsilatGecmisi, dgvPlatformBilanço;
        private Label lblCardTahsilEdilenAdet, lblCardTahsilEdilenPara, lblCardBekleyenAdet, lblCardBekleyenPara, lblCardToplamAdet, lblCardToplamPara;

        public SayfaPlatformOdemeleri2()
        {
            InitializeComponent();
            this.Load += SayfaPlatformOdemeleri2_Load;
        }

        // 🚀 MÜHÜR: DESIGNER DOSYASININ ARADIĞI LOAD EVENT METODU KORUMA ALTINA ALINDI
        private void SayfaPlatformOdemeleri2_Load(object sender, EventArgs e)
        {
        }

        public void SetPlatform(string platformName)
        {
            secilenPlatform = platformName.ToUpper(CultureInfo.InvariantCulture);
            PlatformOdemeleriniYukle();
        }

        private void PlatformOdemeleriniYukle()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(20, 20, 22);

            Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblTitle = new Label { Text = $"💸 {secilenPlatform} MUTABAKAT VE ÖDEME TERMİNALİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI Black", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlTop.Controls.Add(lblTitle); this.Controls.Add(pnlTop);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.Controls.Add(mainLayout); mainLayout.BringToFront();

            Panel pnlCard = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15), AutoScroll = true };
            dynDtTarih = new DateTimePicker { Location = new Point(15, 45), Width = 340, Font = new Font("Segoe UI", 10.5F), Format = DateTimePickerFormat.Short };
            pnlCard.Controls.Add(new Label { Text = "1. Ödemenin Geldiği Tarih:", Location = new Point(15, 23), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray });
            pnlCard.Controls.Add(dynDtTarih);

            pnlCard.Controls.Add(FormElemaniYarat("2. Hesaba Yatan Toplam Tutar (TL):", out dynTxtNetGelenPara, new Point(15, 90), 340));
            dynTxtSiparisListesi = new TextBox { Location = new Point(15, 190), Width = 340, Height = 120, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 10F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };
            pnlCard.Controls.Add(new Label { Text = "3. Bu Ödemeye Ait Sipariş Numaraları:", Location = new Point(15, 165), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray });
            pnlCard.Controls.Add(dynTxtSiparisListesi);

            Button btnMuhurle = new Button { Text = "✅ TAHSİLAT HAKEDİŞİNİ MÜHÜRLE", Location = new Point(15, 330), Width = 340, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            btnMuhurle.Click += DynBtnOdemeyiMuhurle_Click; pnlCard.Controls.Add(btnMuhurle);
            mainLayout.Controls.Add(pnlCard, 0, 0);

            TableLayoutPanel pnlRightLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
            pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));
            pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 35F));
            pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 95F));
            mainLayout.Controls.Add(pnlRightLayout, 1, 0);

            Panel pnlUpperWrapper = new Panel { Dock = DockStyle.Fill };
            pnlUpperWrapper.Controls.Add(new Label { Text = $"📋 SİSTEME İŞLENEN {secilenPlatform} TAHSİLAT MAKBUZLARI", Dock = DockStyle.Top, Height = 22, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            dynDgvTahsilatGecmisi = GridYarat(); pnlUpperWrapper.Controls.Add(dynDgvTahsilatGecmisi); dynDgvTahsilatGecmisi.BringToFront();
            Panel pnlRightBottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(0, 5, 0, 0) };
            Button btnKaydiSil = new Button { Text = "🗑️ SEÇİLİ HAKEDİŞİ İPTAL ET", Dock = DockStyle.Left, Width = 250, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            btnKaydiSil.Click += DynBtnKaydiSil_Click; pnlRightBottom.Controls.Add(btnKaydiSil); pnlUpperWrapper.Controls.Add(pnlRightBottom);
            pnlRightLayout.Controls.Add(pnlUpperWrapper, 0, 0);

            GroupBox grpBilanço = new GroupBox { Dock = DockStyle.Fill, Text = $"📊 {secilenPlatform} CARİ BİLANÇO MATRİSİ", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange, Padding = new Padding(10, 22, 10, 10) };
            dgvPlatformBilanço = GridYarat(); grpBilanço.Controls.Add(dgvPlatformBilanço); pnlRightLayout.Controls.Add(grpBilanço, 0, 1);

            TableLayoutPanel tlpVIPCards = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Padding = new Padding(0, 5, 0, 5) };
            tlpVIPCards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F)); tlpVIPCards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F)); tlpVIPCards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            CustomVIPCardInsaEt(tlpVIPCards, "💰 TAHSİL EDİLEN", Color.FromArgb(34, 197, 94), 0, out lblCardTahsilEdilenAdet, out lblCardTahsilEdilenPara);
            CustomVIPCardInsaEt(tlpVIPCards, "⏳ BEKLEYEN NET ALACAK", Color.Orange, 1, out lblCardBekleyenAdet, out lblCardBekleyenPara);
            CustomVIPCardInsaEt(tlpVIPCards, "💎 GENEL KANAL HACMİ", Color.FromArgb(56, 189, 248), 2, out lblCardToplamAdet, out lblCardToplamPara);
            pnlRightLayout.Controls.Add(tlpVIPCards, 0, 2);

            TahsilatGecmisiniListele(); CanliBilançoListele();
        }

        private void TahsilatGecmisiniListele()
        {
            try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT Id, Tarih, Tutar as 'Yatan Net Tutar (TL)', SiparisAdedi as 'Kapatılan Sipariş' FROM PlatformTahsilatlari WHERE Platform=@p ORDER BY Id DESC", conn)) { cmd.Parameters.AddWithValue("@p", secilenPlatform); DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader()); dynDgvTahsilatGecmisi.DataSource = dt; if (dynDgvTahsilatGecmisi.Columns["Id"] != null) dynDgvTahsilatGecmisi.Columns["Id"].Visible = false; } } } catch { }
        }

        private void CanliBilançoListele()
        {
            try
            {
                DataTable dt = new DataTable(); dt.Columns.Add("Kanal"); dt.Columns.Add("Ödenen (Adet)", typeof(int)); dt.Columns.Add("Net TL", typeof(double)); dt.Columns.Add("Bekleyen (Adet)", typeof(int)); dt.Columns.Add("Kalan Net", typeof(double));
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open(); using (var cmd = new SqliteCommand("SELECT s.PlatformAd, SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN s.SatisFiyati ELSE 0 END), SUM(CASE WHEN s.Durum!='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(CASE WHEN s.Durum!='TAHSİL EDİLDİ' THEN s.SatisFiyati ELSE 0 END) FROM Siparisler s WHERE s.PlatformAd=@p GROUP BY s.PlatformAd", conn))
                    {
                        cmd.Parameters.AddWithValue("@p", secilenPlatform);
                        using (var r = cmd.ExecuteReader()) { while (r.Read()) { dt.Rows.Add(r[0].ToString(), Convert.ToInt32(r[1]), Convert.ToDouble(r[2]), Convert.ToInt32(r[3]), Convert.ToDouble(r[4])); } }
                    }
                }
                dgvPlatformBilanço.DataSource = dt;
                if (dt.Rows.Count > 0)
                {
                    lblCardTahsilEdilenAdet.Text = dt.Rows[0][1].ToString() + " Sipariş";
                    lblCardTahsilEdilenPara.Text = Convert.ToDouble(dt.Rows[0][2]).ToString("N2") + " ₺";
                    lblCardBekleyenAdet.Text = dt.Rows[0][3].ToString() + " Sipariş";
                    lblCardBekleyenPara.Text = Convert.ToDouble(dt.Rows[0][4]).ToString("N2") + " ₺";
                    double ciro = Convert.ToDouble(dt.Rows[0][2]) + Convert.ToDouble(dt.Rows[0][4]);
                    lblCardToplamAdet.Text = (Convert.ToInt32(dt.Rows[0][1]) + Convert.ToInt32(dt.Rows[0][3])) + " Sipariş";
                    lblCardToplamPara.Text = ciro.ToString("N2") + " ₺";
                }
            }
            catch { }
        }

        private void DynBtnOdemeyiMuhurle_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(dynTxtNetGelenPara.Text)) return;
            try
            {
                double tutar = ExcelHucredenSayiya(dynTxtNetGelenPara.Text); string tar = dynDtTarih.Value.ToString("dd.MM.yyyy");
                var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList();
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("INSERT INTO PlatformTahsilatlari (Platform, Tarih, Tutar, SiparisAdedi, SiparisNumaralari) VALUES (@p, @t, @v, @a, @n)", b))
                    {
                        cmd.Parameters.AddWithValue("@p", secilenPlatform); cmd.Parameters.AddWithValue("@t", tar); cmd.Parameters.AddWithValue("@v", tutar); cmd.Parameters.AddWithValue("@a", sips.Count); cmd.Parameters.AddWithValue("@n", string.Join(", ", sips)); cmd.ExecuteNonQuery();
                    }
                    foreach (var s in sips) { using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='TAHSİL EDİLDİ' WHERE SiparisNo=@sn AND PlatformAd=@p", b)) { cmd.Parameters.AddWithValue("@sn", s); cmd.Parameters.AddWithValue("@p", secilenPlatform); cmd.ExecuteNonQuery(); } }
                }
                dynTxtNetGelenPara.Clear(); dynTxtSiparisListesi.Clear(); TahsilatGecmisiniListele(); CanliBilançoListele(); MessageBox.Show("Mühürlendi patron!");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void DynBtnKaydiSil_Click(object sender, EventArgs e)
        {
            if (dynDgvTahsilatGecmisi.SelectedRows.Count == 0) return; int id = Convert.ToInt32(dynDgvTahsilatGecmisi.SelectedRows[0].Cells["Id"].Value);
            if (MessageBox.Show("Bu kaydı silmek emin misiniz?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open(); string nolar = "";
                        using (var cmd = new SqliteCommand("SELECT SiparisNumaralari FROM PlatformTahsilatlari WHERE Id=@id", b)) { cmd.Parameters.AddWithValue("@id", id); nolar = cmd.ExecuteScalar()?.ToString() ?? ""; }
                        var sips = nolar.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim());
                        foreach (var s in sips) { using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='KARGOYA VERİLDİ' WHERE SiparisNo=@sn AND PlatformAd=@p", b)) { cmd.Parameters.AddWithValue("@sn", s); cmd.Parameters.AddWithValue("@p", secilenPlatform); cmd.ExecuteNonQuery(); } }
                        using (var cmd = new SqliteCommand("DELETE FROM PlatformTahsilatlari WHERE Id=@id", b)) { cmd.Parameters.AddWithValue("@id", id); cmd.ExecuteNonQuery(); }
                    }
                    TahsilatGecmisiniListele(); CanliBilançoListele();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

        private double ExcelHucredenSayiya(object hucre) { if (hucre == null || hucre == DBNull.Value) return 0; if (hucre is double d) return d; string s = hucre.ToString().Replace("₺", "").Replace("TL", "").Replace(" ", "").Trim(); if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double resFinal)) return resFinal; return 0; }
        private DataGridView GridYarat() { DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9F) }; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgv.DefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); return dgv; }
        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan = "") { Panel p = new Panel { Location = konum, Width = genislik, Height = 52 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray }; kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
        private void CustomVIPCardInsaEt(TableLayoutPanel tlp, string b, Color c, int s, out Label lA, out Label lP) { Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Margin = new Padding(6) }; Panel pL = new Panel { Dock = DockStyle.Left, Width = 5, BackColor = c }; p.Controls.Add(pL); Label lblB = new Label { Text = b, Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.DarkGray }; lP = new Label { Text = "0,00 ₺", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI Black", 13F, FontStyle.Bold), ForeColor = Color.White }; lA = new Label { Text = "0 Sipariş", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 8.5F, FontStyle.Italic), ForeColor = Color.Gray }; p.Controls.Add(lA); p.Controls.Add(lP); p.Controls.Add(lblB); tlp.Controls.Add(p, s, 0); }
    }
}