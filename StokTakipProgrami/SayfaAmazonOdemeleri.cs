using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;
using System.IO;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;

namespace StokTakipProgrami
{
    public class AmazonDbSiparis
    {
        public string Durum { get; set; }
        public double Satis { get; set; }
        public double Komisyon { get; set; }
        public double Kargo { get; set; }
        public double Hizmet { get; set; }
        public double Ceza { get; set; }
    }

    public partial class SayfaAmazonOdemeleri : UserControl
    {
        private DateTimePicker dynDtTarih;
        private TextBox dynTxtNetGelenPara, dynTxtSiparisListesi;
        private Button dynBtnOdemeyiMuhurle, dynBtnKaydiSil, dynBtnExcel, dynBtnSiparisSorgula, dynBtnTemizle, dynBtnYukluDosyaAc, btnAmazonEkstreYukle;
        private DataGridView dynDgvTahsilatGecmisi, dgvPlatformBilanço;
        private Label lblCardTahsilEdilenAdet, lblCardTahsilEdilenPara, lblCardBekleyenAdet, lblCardBekleyenPara, lblCardToplamAdet, lblCardToplamPara, lblCardCiroAdet, lblCardCiroPara, lblCardKarAdet, lblCardKarPara;

        private DataTable sonCozumlenenTablo = null;
        private string sonSiparisKolonAdi = "Sipariş No.";
        private string sonIslemKolonAdi = "İşlem tipi";
        private string sonHakedisKolonAdi = "Toplam (TRY)";

        public SayfaAmazonOdemeleri()
        {
            InitializeComponent();
            this.Load += SayfaAmazonOdemeleri_Load;
        }

        private void SayfaAmazonOdemeleri_Load(object sender, EventArgs e)
        {
            TabloyuGuvenceyeAl();
            ArayuzInsaEt();
            TahsilatGecmisiniListele();
            CanliBilançoListele();
        }

        private void TabloyuGuvenceyeAl()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS PlatformTahsilatlari (Id INTEGER PRIMARY KEY AUTOINCREMENT, Platform TEXT, Tarih TEXT, Tutar REAL, SiparisAdedi INTEGER, SiparisNumaralari TEXT);", conn))
                    {
                        cmd.ExecuteNonQuery();
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

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(255, 153, 0) }; // Amazon Orange
            pnlHeader.Controls.Add(new Label { Text = "📦 AMAZON MUTABAKAT VE HAKEDİŞ ÖDEMELERİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.FromArgb(35, 47, 62), TextAlign = ContentAlignment.MiddleCenter }); // Amazon Dark Blue Text
            this.Controls.Add(pnlHeader);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.Controls.Add(mainLayout);

            Panel pC = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15), AutoScroll = true };

            int cY = 12;
            pC.Controls.Add(new Label { Text = "Ödeme Tarihi:", Location = new Point(15, cY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            dynDtTarih = new DateTimePicker { Location = new Point(15, cY + 20), Width = 340, Font = new Font("Segoe UI", 9.5F), Format = DateTimePickerFormat.Short };
            pC.Controls.Add(dynDtTarih);
            cY += 50;

            pC.Controls.Add(FormElemaniYarat("Hesaba Yatan Toplam Tutar (TL):", out dynTxtNetGelenPara, new Point(15, cY), 340));
            cY += 50;

            pC.Controls.Add(new Label { Text = "Sipariş Numaraları:", Location = new Point(15, cY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            dynTxtSiparisListesi = new TextBox { Location = new Point(15, cY + 20), Width = 340, Height = 85, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 9.5F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };
            pC.Controls.Add(dynTxtSiparisListesi);
            cY += 110;

            dynBtnTemizle = new Button { Text = "🧹 TEMİZLE", Location = new Point(15, cY), Width = 95, Height = 30, BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.LightGray, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnTemizle.FlatAppearance.BorderSize = 0;
            dynBtnTemizle.Click += (s, e) => {
                dynTxtSiparisListesi.Clear();
                dynTxtNetGelenPara.Clear();
                sonCozumlenenTablo = null;
                dynBtnYukluDosyaAc.Visible = false;
            };
            pC.Controls.Add(dynBtnTemizle);

            dynBtnSiparisSorgula = new Button { Text = "🔍 SİPARİŞ SORGULA", Location = new Point(116, cY), Width = 239, Height = 30, BackColor = Color.FromArgb(35, 47, 62), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSiparisSorgula.FlatAppearance.BorderSize = 0;
            dynBtnSiparisSorgula.Click += DynBtnSiparisSorgula_Click;
            pC.Controls.Add(dynBtnSiparisSorgula);
            cY += 40;

            pC.Controls.Add(new Panel { Location = new Point(15, cY), Width = 340, Height = 1, BackColor = Color.FromArgb(45, 45, 48) });
            cY += 10;

            pC.Controls.Add(new Label { Text = "Amazon Ekstre İşlemleri:", Location = new Point(15, cY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(255, 153, 0) });
            cY += 22;

            dynBtnYukluDosyaAc = new Button { Text = "📂 YÜKLÜ DOSYAYI TEKRAR AÇ", Location = new Point(15, cY), Width = 340, Height = 28, BackColor = Color.FromArgb(38, 38, 41), ForeColor = Color.Gold, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Visible = false, Cursor = Cursors.Hand };
            dynBtnYukluDosyaAc.FlatAppearance.BorderSize = 1;
            dynBtnYukluDosyaAc.FlatAppearance.BorderColor = Color.Gold;
            dynBtnYukluDosyaAc.Click += (s, e) => { if (sonCozumlenenTablo != null) MutabakatUyariEkraniniAc(sonCozumlenenTablo, "AMAZON", sonSiparisKolonAdi, sonIslemKolonAdi, sonHakedisKolonAdi); };
            pC.Controls.Add(dynBtnYukluDosyaAc);
            cY += 35;

            btnAmazonEkstreYukle = new Button { Text = "📦 AMAZON EKSTRESİ YÜKLE VE ANALİZ ET", Location = new Point(15, cY), Width = 340, Height = 48, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnAmazonEkstreYukle.FlatAppearance.BorderSize = 0;
            btnAmazonEkstreYukle.Click += BtnAmazonEkstreYukle_Click;
            pC.Controls.Add(btnAmazonEkstreYukle);
            cY += 65;

            pC.Controls.Add(new Panel { Location = new Point(15, cY), Width = 340, Height = 1, BackColor = Color.FromArgb(50, 50, 55) });
            cY += 20;

            dynBtnOdemeyiMuhurle = new Button { Text = "✅ ÖDEMEYİ KAYDET", Location = new Point(15, cY), Width = 340, Height = 42, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnOdemeyiMuhurle.FlatAppearance.BorderSize = 0;
            dynBtnOdemeyiMuhurle.Click += DynBtnOdemeyiMuhurle_Click;
            pC.Controls.Add(dynBtnOdemeyiMuhurle);

            mainLayout.Controls.Add(pC, 0, 0);

            TableLayoutPanel pRL = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
            pRL.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));
            pRL.RowStyles.Add(new RowStyle(SizeType.Percent, 35F));
            pRL.RowStyles.Add(new RowStyle(SizeType.Absolute, 95F));
            mainLayout.Controls.Add(pRL, 1, 0);

            Panel pUW = new Panel { Dock = DockStyle.Fill };
            pUW.Controls.Add(new Label { Text = "📋 SİSTEME İŞLENEN TOPLU ÖDEMELER", Dock = DockStyle.Top, Height = 22, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            dynDgvTahsilatGecmisi = GridYarat();
            dynDgvTahsilatGecmisi.CellDoubleClick += DynDgvTahsilatGecmisi_CellDoubleClick;
            dynDgvTahsilatGecmisi.DataBindingComplete += (s, e) => { dynDgvTahsilatGecmisi.ClearSelection(); };
            pUW.Controls.Add(dynDgvTahsilatGecmisi);
            dynDgvTahsilatGecmisi.BringToFront();

            Panel pRB = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(0, 5, 0, 0) };
            dynBtnKaydiSil = new Button { Text = "🗑️ SEÇİLİ ÖDEMEYİ İPTAL ET", Dock = DockStyle.Left, Width = 380, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            dynBtnKaydiSil.FlatAppearance.BorderSize = 0;
            dynBtnKaydiSil.Click += DynBtnKaydiSil_Click;
            pRB.Controls.Add(dynBtnKaydiSil);

            dynBtnExcel = new Button { Text = "📊 EXCEL RAPORU AL", Dock = DockStyle.Right, Width = 130, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            dynBtnExcel.FlatAppearance.BorderSize = 0;
            pRB.Controls.Add(dynBtnExcel);
            pUW.Controls.Add(pRB);
            pRL.Controls.Add(pUW, 0, 0);

            GroupBox grpB = new GroupBox { Dock = DockStyle.Fill, Text = "📊 AMAZON CARİ MUTABAKAT VE HESAP TABLOSU", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.FromArgb(255, 153, 0), Padding = new Padding(10, 22, 10, 10) };
            dgvPlatformBilanço = GridYarat();
            grpB.Controls.Add(dgvPlatformBilanço);
            pRL.Controls.Add(grpB, 0, 1);

            TableLayoutPanel tVIP = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, Padding = new Padding(0, 5, 0, 5) };
            for (int i = 0; i < 5; i++) tVIP.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));

            CustomVIPCardInsaEt(tVIP, "💰 TOPLAM TAHSİLAT", Color.FromArgb(34, 197, 94), 0, out lblCardTahsilEdilenAdet, out lblCardTahsilEdilenPara);
            CustomVIPCardInsaEt(tVIP, "⏳ KESİNTİLER", Color.Orange, 1, out lblCardBekleyenAdet, out lblCardBekleyenPara);
            CustomVIPCardInsaEt(tVIP, "💎 KALAN NET ALACAK", Color.FromArgb(56, 189, 248), 2, out lblCardToplamAdet, out lblCardToplamPara);
            CustomVIPCardInsaEt(tVIP, "📊 TOPLAM CİRO", Color.FromArgb(168, 85, 247), 3, out lblCardCiroAdet, out lblCardCiroPara);
            CustomVIPCardInsaEt(tVIP, "📈 NET KÂR DURUMU", Color.FromArgb(236, 72, 153), 4, out lblCardKarAdet, out lblCardKarPara);
            pRL.Controls.Add(tVIP, 0, 2);
            mainLayout.BringToFront();
            pnlHeader.SendToBack();
        }

        private Panel FormElemaniYarat(string b, out TextBox k, Point l, int w, string v = "")
        {
            Panel p = new Panel { Location = l, Width = w, Height = 52 };
            p.Controls.Add(new Label { Text = b, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            k = new TextBox { Text = v, Location = new Point(0, 18), Width = w, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            p.Controls.Add(k);
            return p;
        }

        private Panel FormElemaniYaratRenk(string b, out TextBox k, Point l, int w, string v, Color c)
        {
            Panel p = new Panel { Location = l, Width = w, Height = 48 };
            p.Controls.Add(new Label { Text = b, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = c });
            k = new TextBox { Text = v, Location = new Point(0, 16), Width = w, Font = new Font("Segoe UI", 9.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            p.Controls.Add(k);
            return p;
        }

        private DataGridView GridYarat()
        {
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9F), BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false, ColumnHeadersHeight = 32 };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgv.DataError += (s, e) => { e.Cancel = true; };
            return dgv;
        }

        private void CustomVIPCardInsaEt(TableLayoutPanel t, string b, Color c, int s, out Label lA, out Label lP)
        {
            Panel pC = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(22, 22, 26), Margin = new Padding(6), Padding = new Padding(1) };
            pC.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 3, BackColor = c });
            Panel pT = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 6, 6, 6), BackColor = Color.FromArgb(26, 26, 30) };
            Label lblB = new Label { Text = b.ToUpper(CultureInfo.CurrentCulture), Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(160, 160, 170), TextAlign = ContentAlignment.MiddleLeft };
            lP = new Label { Text = "0,00 ₺", Dock = DockStyle.Top, Height = 26, Font = new Font("Segoe UI Black", 13.5F, FontStyle.Bold), ForeColor = c, TextAlign = ContentAlignment.MiddleLeft };
            lA = new Label { Text = "0 Sipariş", Dock = DockStyle.Top, Height = 18, Font = new Font("Segoe UI", 8.0F), ForeColor = Color.FromArgb(110, 110, 120), TextAlign = ContentAlignment.MiddleLeft };
            pT.Controls.Add(lA); pT.Controls.Add(lP); pT.Controls.Add(lblB);
            lblB.BringToFront(); lP.BringToFront(); lA.BringToFront();
            pC.Controls.Add(pT); t.Controls.Add(pC, s, 0);
        }

        private bool PromptForInput(string t, string c, out string v)
        {
            Form p = new Form() { Width = 350, Height = 180, FormBorderStyle = FormBorderStyle.FixedDialog, Text = c, StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
            TextBox ib = new TextBox() { Left = 20, Top = 50, Width = 290, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, Font = new Font("Segoe UI", 11F) };
            Button cf = new Button() { Text = "Onayla", Left = 210, Width = 100, Top = 90, BackColor = Color.FromArgb(255, 153, 0), ForeColor = Color.Black, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            string r = ""; bool ok = false;
            cf.Click += (s, e) => { r = ib.Text.Trim(); ok = true; p.Close(); };
            p.Controls.Add(new Label() { Left = 20, Top = 20, Text = t, AutoSize = true, ForeColor = Color.White });
            p.Controls.Add(ib); p.Controls.Add(cf);
            p.AcceptButton = cf; p.ShowDialog();
            v = r; return ok;
        }

        private void FormMükerrerDetayGoster(string rapor, int temiz)
        {
            Form f = new Form { Text = "🛑 Çapraz Mutabakat Mükerrer Koruma", Size = new Size(680, 420), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
            Label l = new Label { Text = $"⚠️ Aşağıdaki siparişler daha önceden TAHSİL EDİLDİĞİ için tekrar kaydedilmedi.\nTemiz olan {temiz} sipariş mühürlendi!", Dock = DockStyle.Top, Height = 50, ForeColor = Color.LightCoral, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Padding = new Padding(5) };
            TextBox t = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, BackColor = Color.FromArgb(24, 24, 27), ForeColor = Color.Orange, Font = new Font("Consolas", 10F), Text = rapor };
            f.Controls.Add(t); f.Controls.Add(l); f.ShowDialog();
        }

        private double ExcelHucredenSayiya(object h)
        {
            if (h == null || h == DBNull.Value) return 0;
            if (h is double d) return d;
            if (h is decimal dc) return Convert.ToDouble(dc);
            if (h is int i) return i;
            string s = h.ToString().Replace("₺", "").Replace("TRY", "").Replace("TL", "").Replace(" ", "").Replace("\xA0", "").Trim();
            if (string.IsNullOrEmpty(s)) return 0;
            int n = s.LastIndexOf('.'); int v = s.LastIndexOf(',');
            if (n > v)
            {
                if (s.Substring(n).Length <= 3) s = s.Substring(0, n).Replace(".", "").Replace(",", "") + "." + s.Substring(n + 1);
                else s = s.Replace(".", "").Replace(",", "");
            }
            else if (v > n)
            {
                if (s.Substring(v).Length <= 3) s = s.Substring(0, v).Replace(".", "").Replace(",", "") + "." + s.Substring(v + 1);
                else s = s.Replace(".", "").Replace(",", "");
            }
            else s = s.Replace(".", "").Replace(",", "");
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double rf)) return rf;
            return 0;
        }

        private string[] SatiriParcala(string s, char a)
        {
            var r = new List<string>(); bool t = false; string h = "";
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"') t = !t;
                else if (c == a && !t) { r.Add(h.Trim('"')); h = ""; }
                else h += c;
            }
            r.Add(h.Trim('"')); return r.ToArray();
        }

        private DataTable ExcelVeCsvOku(string dosyaYolu)
        {
            DataTable dt = new DataTable();
            string ext = Path.GetExtension(dosyaYolu).ToLower();
            string icerik = "";
            try
            {
                byte[] bytes = File.ReadAllBytes(dosyaYolu);
                if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) icerik = Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
                else icerik = Encoding.GetEncoding(1254).GetString(bytes);
            }
            catch { }

            if (ext == ".xls" || ext == ".xlsx")
            {
                try
                {
                    string connStr = ext == ".xls" ? $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dosyaYolu};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\";" : $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dosyaYolu};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";
                    using (OleDbConnection conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string tS = "";
                        foreach (DataRow r in schema.Rows)
                        {
                            string sN = r["TABLE_NAME"].ToString();
                            if (!sN.ToUpper().Contains("OZET") && !sN.ToUpper().Contains("SUMMARY")) { tS = sN; break; }
                        }
                        if (string.IsNullOrEmpty(tS) && schema.Rows.Count > 0) tS = schema.Rows[0]["TABLE_NAME"].ToString();
                        if (tS.StartsWith("'") && tS.EndsWith("'")) tS = tS.Substring(1, tS.Length - 2);
                        using (OleDbDataAdapter da = new OleDbDataAdapter($"SELECT * FROM [{tS}]", conn)) { da.Fill(dt); }
                    }
                    if (dt.Rows.Count > 0) return dt;
                }
                catch { }
            }
            try
            {
                string[] lines = icerik.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                if (lines.Length > 0)
                {
                    char ayrac = lines[0].Contains(";") ? ';' : (lines[0].Contains("\t") ? '\t' : ',');
                    string[] headers = SatiriParcala(lines[0], ayrac);
                    foreach (string h in headers)
                    {
                        string cN = string.IsNullOrWhiteSpace(h) ? "Bos" : h.Trim();
                        int s = 1; string fN = cN;
                        while (dt.Columns.Contains(fN)) { fN = cN + s++; }
                        dt.Columns.Add(fN);
                    }
                    for (int i = 1; i < lines.Length; i++)
                    {
                        if (string.IsNullOrWhiteSpace(lines[i])) continue;
                        string[] cells = SatiriParcala(lines[i], ayrac);
                        DataRow row = dt.NewRow();
                        for (int j = 0; j < cells.Length && j < dt.Columns.Count; j++) row[j] = cells[j].Trim();
                        dt.Rows.Add(row);
                    }
                }
            }
            catch { }
            return dt;
        }

        private void BtnAmazonEkstreYukle_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Tüm Formatlar|*.csv;*.xlsx;*.xls;*.txt";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        DataTable dt = ExcelVeCsvOku(ofd.FileName);
                        if (dt == null || dt.Columns.Count == 0 || dt.Rows.Count == 0) { MessageBox.Show("Sistem dosyayı okuyamadı.", "Okuma Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                        int idxSip = -1, idxHak = -1, idxIslem = -1, idxKomisyon = -1, idxKargoDiger = -1;
                        int dataStartRow = 0;

                        // Amazon Sütunlarını Bul
                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            string b = dt.Columns[c].ColumnName.ToUpper().Replace(" ", "").Replace("_", "").Trim();
                            if (b.Contains("SİPARİŞNO") || b == "SIPARISNO") idxSip = c;
                            else if (b.Contains("TOPLAM(TRY)") || b == "TOPLAM" || b == "NETTUTAR") idxHak = c;
                            else if (b.Contains("İŞLEMTİPİ") || b.Contains("ISLEMTIPI")) idxIslem = c;
                            else if (b.Contains("AMAZONÜCRETLERİ") || b.Contains("AMAZONUCRETLERI") || b.Contains("KOMİSYON")) idxKomisyon = c;
                            else if (b == "DİĞER" || b == "DIGER") idxKargoDiger = c;
                        }

                        if (idxSip == -1 || idxHak == -1)
                        {
                            for (int r = 0; r < Math.Min(dt.Rows.Count, 15); r++)
                            {
                                for (int c = 0; c < dt.Columns.Count; c++)
                                {
                                    string b = dt.Rows[r][c]?.ToString().ToUpper().Replace(" ", "").Replace("_", "").Trim() ?? "";
                                    if (b.Contains("SİPARİŞNO") || b == "SIPARISNO") idxSip = c;
                                    else if (b.Contains("TOPLAM(TRY)") || b == "TOPLAM" || b == "NETTUTAR") idxHak = c;
                                    else if (b.Contains("İŞLEMTİPİ") || b.Contains("ISLEMTIPI")) idxIslem = c;
                                    else if (b.Contains("AMAZONÜCRETLERİ") || b.Contains("AMAZONUCRETLERI") || b.Contains("KOMİSYON")) idxKomisyon = c;
                                    else if (b == "DİĞER" || b == "DIGER") idxKargoDiger = c;
                                }
                                if (idxSip != -1 && idxHak != -1) { dataStartRow = r + 1; break; }
                            }
                        }

                        if (idxSip == -1 || idxHak == -1) { MessageBox.Show("Yüklenen ekstre dosyasındaki sütun başlıkları Amazon formatıyla uyuşmuyor!"); return; }

                        var sips = new HashSet<string>();
                        double totalHesabaYatan = 0;
                        var excelOzet = new Dictionary<string, double[]>(); // 0:Satis, 1:Komisyon, 2:Kargo, 3:Hizmet/Diger, 4:NetHakedis

                        for (int i = dataStartRow; i < dt.Rows.Count; i++)
                        {
                            int doluHucre = 0; for (int c = 0; c < dt.Columns.Count; c++) { if (!string.IsNullOrWhiteSpace(dt.Rows[i].ItemArray[c]?.ToString())) doluHucre++; }
                            if (doluHucre < 3) continue;

                            string sn = dt.Rows[i].ItemArray[idxSip].ToString().Trim();
                            string islemTipi = idxIslem != -1 ? dt.Rows[i].ItemArray[idxIslem].ToString().Trim() : "";

                            // Amazon "---" (çizgi) sipariş noları genellikle genel kesintidir, pass geçiyoruz.
                            if (string.IsNullOrEmpty(sn) || sn == "---" || sn.ToUpper().Contains("TOPLAM") || sn.ToUpper().Contains("TOTAL")) continue;

                            sips.Add(sn);
                            if (!excelOzet.ContainsKey(sn)) excelOzet[sn] = new double[5];

                            double rowTotal = ExcelHucredenSayiya(dt.Rows[i].ItemArray[idxHak]);
                            double rowKomisyon = idxKomisyon != -1 ? Math.Abs(ExcelHucredenSayiya(dt.Rows[i].ItemArray[idxKomisyon])) : 0;
                            double rowDiger = idxKargoDiger != -1 ? Math.Abs(ExcelHucredenSayiya(dt.Rows[i].ItemArray[idxKargoDiger])) : 0;

                            // Amazon'da İşlem Tipi Sipariş Ödemesi ise asıl satırdır. Amazon Kolay Gönderi ise Kargodur.
                            if (islemTipi.ToUpper().Contains("KOLAY GÖNDERİ") || islemTipi.ToUpper().Contains("KOLAY GONDERI") || islemTipi.ToUpper().Contains("KARGO"))
                            {
                                excelOzet[sn][2] += Math.Abs(rowTotal); // Kargo Gideri (Total eksi olduğu için abs alıyoruz)
                            }
                            else if (islemTipi.ToUpper().Contains("DİĞER") || islemTipi.ToUpper().Contains("DIGER") || islemTipi.ToUpper().Contains("VERGİ") || islemTipi.ToUpper().Contains("KDV"))
                            {
                                excelOzet[sn][3] += Math.Abs(rowTotal); // Vergi / Hizmet kesintisi
                            }
                            else
                            {
                                // Normal sipariş satırı
                                excelOzet[sn][1] += rowKomisyon;
                                excelOzet[sn][3] += rowDiger;
                            }

                            excelOzet[sn][4] += rowTotal; // O satırın totali Net Hakedişe ekleniyor
                            totalHesabaYatan += rowTotal;
                        }

                        var dbHavuz = new Dictionary<string, AmazonDbSiparis>();
                        using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                        {
                            conn.Open();
                            foreach (var s in sips)
                            {
                                using (var cmd = new SqliteCommand("SELECT Durum, SatisFiyati, KomisyonTL, KargoUcreti, (HizmetBedeli + Stopaj) as TopHizmet, IFNULL(CezaTutari,0) FROM Siparisler WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%AMAZON%'", conn))
                                {
                                    cmd.Parameters.AddWithValue("@sn", s);
                                    using (var r = cmd.ExecuteReader())
                                    {
                                        if (r.Read()) dbHavuz[s] = new AmazonDbSiparis { Durum = r[0].ToString(), Satis = Convert.ToDouble(r[1]), Komisyon = Convert.ToDouble(r[2]), Kargo = Convert.ToDouble(r[3]), Hizmet = Convert.ToDouble(r[4]), Ceza = Convert.ToDouble(r[5]) };
                                    }
                                }
                            }
                        }

                        DataTable finalTable = new DataTable();
                        finalTable.Columns.Add("Sipariş No");
                        finalTable.Columns.Add("Sistem Durumu");
                        finalTable.Columns.Add("Excel Komisyon (TL)", typeof(double));
                        finalTable.Columns.Add("Excel Kargo (TL)", typeof(double));
                        finalTable.Columns.Add("Net Hakediş (TL)", typeof(double));
                        finalTable.Columns.Add("Dedektif (Durum)");

                        int hatalar = 0;

                        foreach (var sn in sips)
                        {
                            var exData = excelOzet[sn];
                            string ded = "✔ MÜTABIK";
                            string sys = "-";

                            if (dbHavuz.TryGetValue(sn, out var db))
                            {
                                sys = db.Durum;
                                if (sys.ToUpper() == "TAHSİL EDİLDİ") { ded = "HATA: MÜKERRER TAHSİLAT"; }
                                else
                                {
                                    double beklenenNet = db.Satis - (db.Komisyon + db.Kargo + db.Hizmet + db.Ceza);
                                    if (Math.Abs(exData[4] - beklenenNet) > 0.5) ded = "HATA: HAKEDİŞ TUTAR FARKI";
                                }
                            }
                            else { sys = "SİSTEMDE YOK"; ded = "HATA: KAYITSIZ SİPARİŞ"; }

                            if (ded.StartsWith("HATA:")) hatalar++;
                            finalTable.Rows.Add(sn, sys, exData[1], exData[2], exData[4], ded);
                        }

                        dynTxtSiparisListesi.Text = string.Join(Environment.NewLine, sips);
                        dynTxtNetGelenPara.Text = totalHesabaYatan.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));
                        sonCozumlenenTablo = finalTable;
                        sonSiparisKolonAdi = "Sipariş No";
                        sonIslemKolonAdi = "";
                        sonHakedisKolonAdi = "Net Hakediş (TL)";

                        dynBtnYukluDosyaAc.Text = $"📂 YÜKLÜ DOSYAYI AÇ ({Path.GetFileName(ofd.FileName)})";
                        dynBtnYukluDosyaAc.Visible = true;

                        if (hatalar > 0) MessageBox.Show($"🚨 DİKKAT! MÜKERRER VEYA HATALI {hatalar} SİPARİŞ TESPİT EDİLDİ!", "Amazon Radarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        MutabakatUyariEkraniniAc(finalTable, "AMAZON", sonSiparisKolonAdi, sonIslemKolonAdi, sonHakedisKolonAdi);
                    }
                    catch (Exception exDosya) { MessageBox.Show("Sistem Hatası: " + exDosya.Message); }
                }
            }
        }

        private void MutabakatUyariEkraniniAc(DataTable dtSource, string dp, string sipCol, string islCol, string hakCol)
        {
            Form f = new Form { Text = $"{dp} Ekstre Önizleme", Size = new Size(1350, 650), StartPosition = FormStartPosition.CenterScreen, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };

            TableLayoutPanel tP = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(33, 33, 36) };
            tP.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            tP.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tP.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            f.Controls.Add(tP);

            tP.Controls.Add(new Label { Text = "SİPARİŞ DETAYLARI İNCELEME (Çift Tık: Ameliyat Masası | Sağ Tık: Tutar Eşitle)", Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(255, 153, 0), Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) }, 0, 0);

            Panel pG = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 10) };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), SelectionMode = DataGridViewSelectionMode.CellSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EnableHeadersVisualStyles = false, RowHeadersVisible = false };
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.ColumnHeadersHeight = 40;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.DataSource = dtSource;

            dgv.CellDoubleClick += (s, eDouble) => {
                if (eDouble.RowIndex < 0 || eDouble.ColumnIndex < 0) return;
                string col = dgv.Columns[eDouble.ColumnIndex].Name;
                string sn = dgv.Rows[eDouble.RowIndex].Cells[sipCol].Value?.ToString() ?? "";
                if (col == sipCol && !string.IsNullOrEmpty(sn) && sn != "-") { Clipboard.SetText(sn); MessageBox.Show($"Kopyalandı: {sn}"); return; }
                if (!string.IsNullOrEmpty(sn) && sn != "-")
                {
                    double tH = ExcelHucredenSayiya(dgv.Rows[eDouble.RowIndex].Cells[hakCol].Value);
                    SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);
                }
            };

            dgv.DataBindingComplete += (s, eBind) => {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    string ded = dgv.Columns.Contains("Dedektif (Durum)") ? row.Cells["Dedektif (Durum)"].Value?.ToString() ?? "" : "";
                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
                    row.DefaultCellStyle.SelectionForeColor = Color.White;

                    if (ded == "KAYDEDİLDİ" || ded == "HURDA İŞLENDİ" || ded.StartsWith("✔ MÜTABIK") || ded.StartsWith("✔ EŞİTLENDİ")) { row.DefaultCellStyle.BackColor = Color.FromArgb(46, 117, 89); row.DefaultCellStyle.ForeColor = Color.White; }
                    else if (ded.StartsWith("HATA:")) { row.DefaultCellStyle.BackColor = Color.Crimson; row.DefaultCellStyle.ForeColor = Color.White; }
                    else { row.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); row.DefaultCellStyle.ForeColor = Color.White; }
                }
                dgv.ClearSelection();
            };

            ContextMenuStrip cms = new ContextMenuStrip(); dgv.ContextMenuStrip = cms;
            dgv.MouseDown += (s, eMouse) => {
                if (eMouse.Button == MouseButtons.Right)
                {
                    var hit = dgv.HitTest(eMouse.X, eMouse.Y);
                    if (hit.Type == DataGridViewHitTestType.Cell && hit.RowIndex >= 0)
                    {
                        dgv.ClearSelection(); dgv.Rows[hit.RowIndex].Cells[hit.ColumnIndex].Selected = true;
                        var r = dgv.Rows[hit.RowIndex]; string sn = r.Cells[sipCol].Value.ToString();
                        cms.Items.Clear();

                        ToolStripMenuItem bA = new ToolStripMenuItem("✏️ İncele ve Manuel Düzelt");
                        bA.Click += (sx, exA) => {
                            if (string.IsNullOrEmpty(sn) || sn == "-") return;
                            double tH = ExcelHucredenSayiya(r.Cells[hakCol].Value);
                            SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);
                        };
                        cms.Items.Add(bA);

                        cms.Items.Add(new ToolStripSeparator());

                        ToolStripMenuItem btnHizliEsitle = new ToolStripMenuItem("⚡ Amazon Kesintilerini (Komisyon & Kargo) Eşitle");
                        btnHizliEsitle.ForeColor = Color.LimeGreen;
                        btnHizliEsitle.Font = new Font(btnHizliEsitle.Font, FontStyle.Bold);
                        btnHizliEsitle.Click += (sx, exEsitle) => {
                            if (MessageBox.Show($"[{sn}] siparişindeki Amazon kesintileri veritabanıyla eşitlenecektir. Onaylıyor musunuz?", "Toplu Eşitleme", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                            {
                                try
                                {
                                    double exKom = Convert.ToDouble(r.Cells["Excel Komisyon (TL)"].Value);
                                    double exKargo = Convert.ToDouble(r.Cells["Excel Kargo (TL)"].Value);

                                    using (var bConn = new SqliteConnection(Veritabanı.BaglantiYolu))
                                    {
                                        bConn.Open();
                                        using (var cmd = new SqliteCommand("UPDATE Siparisler SET KomisyonTL=@km, KargoUcreti=@kr WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%AMAZON%'", bConn))
                                        {
                                            cmd.Parameters.AddWithValue("@km", exKom);
                                            cmd.Parameters.AddWithValue("@kr", exKargo);
                                            cmd.Parameters.AddWithValue("@sn", sn);
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    if (dgv.Columns.Contains("Dedektif (Durum)")) r.Cells["Dedektif (Durum)"].Value = "✔ EŞİTLENDİ"; dgv.Refresh();
                                    MessageBox.Show("Sistemdeki Kargo ve Komisyon güncellendi!");
                                }
                                catch (Exception exTY) { MessageBox.Show(exTY.Message); }
                            }
                        };
                        cms.Items.Add(btnHizliEsitle);

                        cms.Show(Cursor.Position);
                    }
                }
            };

            pG.Controls.Add(dgv); tP.Controls.Add(pG, 0, 1);

            Panel pB = new Panel { Dock = DockStyle.Fill, Height = 50 }; Button btnK = new Button { Text = "Ödemeleri Onayla ve Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(255, 153, 0), ForeColor = Color.FromArgb(35, 47, 62), Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnK.FlatAppearance.BorderSize = 0; pB.SizeChanged += (sB, eB) => { btnK.Location = new Point((pB.Width - btnK.Width) / 2, (pB.Height - btnK.Height) / 2); };
            btnK.Click += (sB, eB) => { f.DialogResult = DialogResult.OK; f.Close(); }; pB.Controls.Add(btnK); tP.Controls.Add(pB, 0, 2);

            dgv.BringToFront(); f.ShowDialog();
        }

        private void DynBtnSiparisSorgula_Click(object sender, EventArgs e)
        {
            var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList();
            if (sips.Count == 0) { MessageBox.Show("Lütfen sipariş numarası giriniz."); return; }
            if (sips.Count > 1000) sips = sips.Take(1000).ToList();

            Form frmPopup = new Form { Text = $"⚖️ AMAZON MANUEL ÖDEME RADARI ({sips.Count} Sipariş)", Size = new Size(1150, 600), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
            TableLayoutPanel tP = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(33, 33, 36) };
            tP.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); tP.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); tP.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            frmPopup.Controls.Add(tP);

            tP.Controls.Add(new Label { Text = "📢 'Ödenen (TL)' sütununa tutarları yazıp Enter'a basınız. Sipariş satırına ÇİFT TIKLAYARAK ameliyat masasını açabilirsiniz.", Dock = DockStyle.Fill, ForeColor = Color.Gold, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter }, 0, 0);

            Panel pG = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 10) };
            DataGridView dgvPopup = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, Font = new Font("Segoe UI", 9.5F), EnableHeadersVisualStyles = false, BorderStyle = BorderStyle.None, ColumnHeadersHeight = 35 };
            dgvPopup.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgvPopup.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPopup.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgvPopup.DefaultCellStyle.ForeColor = Color.White;

            dgvPopup.Columns.Add("SiparisNo", "Sipariş No"); dgvPopup.Columns.Add("Durum", "Sistem Durumu"); dgvPopup.Columns.Add("Urun", "Kayıtlı Ürün");
            dgvPopup.Columns.Add("Ciro", "Brüt Satış"); dgvPopup.Columns.Add("Kesinti", "Kesintiler"); dgvPopup.Columns.Add("Kargo", "Kargo");
            dgvPopup.Columns.Add("Hakedis", "Sistem Hakedişi"); dgvPopup.Columns.Add("GelenTutar", "Ödenen (TL)"); dgvPopup.Columns.Add("Mutabakat", "Mutabakat Kontrolü");

            for (int i = 0; i < dgvPopup.Columns.Count; i++) if (i != 7) dgvPopup.Columns[i].ReadOnly = true;
            dgvPopup.Columns[3].DefaultCellStyle.Format = "N2"; dgvPopup.Columns[4].DefaultCellStyle.Format = "N2";
            dgvPopup.Columns[5].DefaultCellStyle.Format = "N2"; dgvPopup.Columns[6].DefaultCellStyle.Format = "N2"; dgvPopup.Columns[7].DefaultCellStyle.Format = "N2";
            dgvPopup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvPopup.CellDoubleClick += (sG, eG) => {
                if (eG.RowIndex < 0) return; string selectedSipNo = dgvPopup.Rows[eG.RowIndex].Cells["SiparisNo"].Value?.ToString();
                if (!string.IsNullOrEmpty(selectedSipNo)) SiparisiAmeliyatMasasinaYatir(selectedSipNo, 0, 0, dgvPopup);
            };
            pG.Controls.Add(dgvPopup); tP.Controls.Add(pG, 0, 1);

            using (var c = new SqliteConnection(Veritabanı.BaglantiYolu))
            {
                c.Open();
                foreach (string sn in sips)
                {
                    using (var m = new SqliteCommand(@"SELECT UrunAdi, IFNULL(s.SatisFiyati,0), IFNULL(s.PlatformIndirimi,0), IFNULL(s.Stopaj,0), IFNULL(s.PosBedeli,0), IFNULL(s.HizmetBedeli,0), IFNULL(s.KomisyonTL,0), IFNULL(s.KargoUcreti,0), IFNULL(s.CezaTutari,0), s.Durum FROM Siparisler s WHERE s.SiparisNo = @sn AND UPPER(s.PlatformAd) LIKE '%AMAZON%' AND UPPER(s.Durum) != 'İPTAL'", c))
                    {
                        m.Parameters.AddWithValue("@sn", sn);
                        using (var r = m.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                string sD = r["Durum"].ToString(); double sat = Convert.ToDouble(r[1]);
                                double kes = Convert.ToDouble(r[3]) + Convert.ToDouble(r[4]) + Convert.ToDouble(r[5]) + Convert.ToDouble(r[6]) + Convert.ToDouble(r[8]);
                                double kar = Convert.ToDouble(r[7]); double sH = sat + Convert.ToDouble(r[2]) - kes - kar;
                                if (sD.ToUpper() == "TAHSİL EDİLDİ")
                                {
                                    int ri = dgvPopup.Rows.Add(sn, "❌ MÜKERRER", r["UrunAdi"].ToString(), sat, kes, kar, sH, 0.00, "❌ ÖDENDİ!");
                                    dgvPopup.Rows[ri].DefaultCellStyle.BackColor = Color.FromArgb(75, 20, 20); dgvPopup.Rows[ri].DefaultCellStyle.ForeColor = Color.LightCoral;
                                }
                                else
                                {
                                    int ri = dgvPopup.Rows.Add(sn, "✔ " + sD, r["UrunAdi"].ToString(), sat, kes, kar, sH, 0.00, "✍ Bekleniyor...");
                                    dgvPopup.Rows[ri].DefaultCellStyle.BackColor = Color.FromArgb(38, 45, 38);
                                }
                            }
                            else
                            {
                                int ri = dgvPopup.Rows.Add(sn, "❌ BULUNAMADI", "Yok!", 0.00, 0.00, 0.00, 0.00, 0.00, "Bulunamadı");
                                dgvPopup.Rows[ri].DefaultCellStyle.BackColor = Color.FromArgb(55, 25, 25); dgvPopup.Rows[ri].DefaultCellStyle.ForeColor = Color.LightCoral;
                            }
                        }
                    }
                }
            }

            dgvPopup.CellValueChanged += (sC, ev) => {
                if (ev.RowIndex < 0 || ev.ColumnIndex != 7) return; DataGridViewRow rw = dgvPopup.Rows[ev.RowIndex];
                if (rw.Cells["Durum"].Value.ToString().Contains("BULUNAMADI") || rw.Cells["Durum"].Value.ToString().Contains("MÜKERRER")) { rw.Cells["GelenTutar"].Value = 0.00; return; }
                double sH = Convert.ToDouble(rw.Cells["Hakedis"].Value); double gT = 0;
                if (rw.Cells["GelenTutar"].Value != null) double.TryParse(rw.Cells["GelenTutar"].Value.ToString(), out gT);
                if (Math.Abs(sH - gT) < 0.5) { rw.Cells["Mutabakat"].Value = "✔ UYUŞUYOR"; rw.Cells["Mutabakat"].Style.ForeColor = Color.LimeGreen; rw.DefaultCellStyle.BackColor = Color.FromArgb(20, 50, 20); }
                else { rw.Cells["Mutabakat"].Value = "❌ UYUŞMAZLIK!"; rw.Cells["Mutabakat"].Style.ForeColor = Color.Crimson; rw.DefaultCellStyle.BackColor = Color.FromArgb(65, 20, 20); }
            };

            Panel pB = new Panel { Dock = DockStyle.Fill, Height = 50 }; Button bC = new Button { Text = "Kontrolü Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(255, 153, 0), ForeColor = Color.FromArgb(35, 47, 62), Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            bC.FlatAppearance.BorderSize = 0; pB.SizeChanged += (sB, eB) => { bC.Location = new Point((pB.Width - bC.Width) / 2, (pB.Height - bC.Height) / 2); };
            bC.Click += (sB, eB) => { frmPopup.DialogResult = DialogResult.OK; frmPopup.Close(); }; pB.Controls.Add(bC); tP.Controls.Add(pB, 0, 2);

            frmPopup.ShowDialog();
        }

        private void SiparisiAmeliyatMasasinaYatir(string sN, double eH, double eK, DataGridView aG)
        {
            long rI = 0; double sF = 0, kM = 0, kR = 0, hZ = 0, sT = 0, pS = 0, cZ = 0; string dR = "YOK", bR = "-", uA = "Bilinmeyen Ürün", kT = "-", mR = "-", tA = "-"; int aD = 1; double aF = 0; string iM = "-";
            try
            {
                using (var c = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    c.Open(); string q = @"SELECT s.rowid, s.SatisFiyati, s.KomisyonTL, s.KargoUcreti, s.HizmetBedeli, s.Stopaj, s.PosBedeli, IFNULL(s.CezaTutari,0), s.Durum, s.Barkod, IFNULL(u.UrunAdi, s.UrunAdi), IFNULL(u.Kategori, s.Kategori), IFNULL(u.Marka, s.Marka), IFNULL(s.Adet,1), s.Tarih, IFNULL(s.AlisFiyati,0), IFNULL(s.SatilanImei,'-') FROM Siparisler s LEFT JOIN Urunler u ON s.Barkod = u.Barkod WHERE s.SiparisNo = @sn AND UPPER(s.PlatformAd) LIKE '%AMAZON%'";
                    using (var m = new SqliteCommand(q, c))
                    {
                        m.Parameters.AddWithValue("@sn", sN);
                        using (var r = m.ExecuteReader())
                        {
                            if (r.Read()) { rI = Convert.ToInt64(r[0]); sF = Convert.ToDouble(r[1]); kM = Convert.ToDouble(r[2]); kR = Convert.ToDouble(r[3]); hZ = Convert.ToDouble(r[4]); sT = Convert.ToDouble(r[5]); pS = Convert.ToDouble(r[6]); cZ = Convert.ToDouble(r[7]); dR = r[8].ToString(); bR = r[9].ToString(); uA = r[10].ToString(); kT = r[11].ToString(); mR = r[12].ToString(); aD = Convert.ToInt32(r[13]); tA = r[14].ToString(); aF = Convert.ToDouble(r[15]); iM = r[16].ToString(); }
                            else return;
                        }
                    }
                }
            }
            catch { return; }

            Form f = new Form { Text = $"Düzeltme Masası - No: {sN}", Size = new Size(520, 660), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
            Panel pM = new Panel { Location = new Point(15, 10), Width = 475, Height = 105, BackColor = Color.FromArgb(24, 24, 27), BorderStyle = BorderStyle.FixedSingle };
            string tB = (string.IsNullOrEmpty(bR) || bR.ToUpper() == uA.ToUpper()) ? "-" : bR;
            pM.Controls.Add(new Label { Text = $"● SİSTEM DURUMU: {dR.ToUpper()}", Location = new Point(15, 12), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.Gold });
            pM.Controls.Add(new Label { Text = $"📦 Ürün: {uA} (Adet: {aD})\r\n🆔 Barkod: {tB} | Kategori: {kT} | Marka: {mR}", Location = new Point(15, 36), Size = new Size(445, 38), Font = new Font("Segoe UI", 9F), ForeColor = Color.FromArgb(56, 189, 248) });
            pM.Controls.Add(new Label { Text = $"💰 Geçmiş -> Tarih: {tA} | Alış: {aF:N2}₺ | IMEI: {iM}", Location = new Point(15, 78), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(34, 197, 94) }); f.Controls.Add(pM);

            int cY = 125;
            GroupBox gG = new GroupBox { Text = " Satış Gelirleri", Location = new Point(15, cY), Width = 475, Height = 70, ForeColor = Color.FromArgb(34, 197, 94), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            TextBox tSa; gG.Controls.Add(FormElemaniYaratRenk("Satış Fiyatı (TL):", out tSa, new Point(15, 18), 440, sF.ToString("N2"), Color.White)); f.Controls.Add(gG); cY += 75;

            GroupBox gM = new GroupBox { Text = " Ürün Maliyet Bilgisi", Location = new Point(15, cY), Width = 475, Height = 70, ForeColor = Color.FromArgb(56, 189, 248), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            TextBox tAl, tIm; gM.Controls.Add(FormElemaniYaratRenk("Alış Maliyeti (TL):", out tAl, new Point(15, 18), 210, aF.ToString("N2"), Color.White)); gM.Controls.Add(FormElemaniYaratRenk("IMEI Numarası:", out tIm, new Point(245, 18), 210, iM, Color.White)); f.Controls.Add(gM); cY += 75;

            GroupBox gD = new GroupBox { Text = " Kesintiler & Platform Vergileri", Location = new Point(15, cY), Width = 475, Height = 195, ForeColor = Color.FromArgb(239, 68, 68), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            TextBox tKo, tCe, tKr, tHz, tSt, tPo;
            gD.Controls.Add(FormElemaniYaratRenk("Komisyon (TL):", out tKo, new Point(15, 18), 210, kM.ToString("N2"), Color.White)); gD.Controls.Add(FormElemaniYaratRenk("Kargo Ücreti (TL):", out tKr, new Point(245, 18), 210, kR.ToString("N2"), Color.White));
            gD.Controls.Add(FormElemaniYaratRenk("Hizmet Bedeli (TL):", out tHz, new Point(15, 74), 210, hZ.ToString("N2"), Color.White)); gD.Controls.Add(FormElemaniYaratRenk("Ceza Tutarı (TL):", out tCe, new Point(245, 74), 210, cZ.ToString("N2"), Color.White));
            gD.Controls.Add(FormElemaniYaratRenk("Stopaj (TL):", out tSt, new Point(15, 130), 210, sT.ToString("N2"), Color.White)); gD.Controls.Add(FormElemaniYaratRenk("POS Kesintisi (TL):", out tPo, new Point(245, 130), 210, pS.ToString("N2"), Color.White)); f.Controls.Add(gD); cY += 205;

            Button bU = new Button { Text = "💾 GÜNCELLE", Location = new Point(15, cY + 25), Width = 475, Height = 45, BackColor = Color.FromArgb(255, 153, 0), ForeColor = Color.FromArgb(35, 47, 62), Font = new Font("Segoe UI", 10F, FontStyle.Bold), FlatStyle = FlatStyle.Flat }; bU.FlatAppearance.BorderSize = 0;
            bU.Click += (s, ev) => {
                try
                {
                    double.TryParse(tSa.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n1); double.TryParse(tKo.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n2); double.TryParse(tKr.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n4); double.TryParse(tHz.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n5); double.TryParse(tSt.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n6); double.TryParse(tPo.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n7); double.TryParse(tCe.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n3); double.TryParse(tAl.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nA); string nI = tIm.Text.Trim();
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET SatisFiyati=@sf, AlisFiyati=@af, SatilanImei=@im, KomisyonTL=@km, CezaTutari=@cz, KargoUcreti=@kr, HizmetBedeli=@hz, Stopaj=@st, PosBedeli=@pb WHERE rowid=@id", b))
                        {
                            cmd.Parameters.AddWithValue("@sf", n1); cmd.Parameters.AddWithValue("@af", nA); cmd.Parameters.AddWithValue("@im", nI); cmd.Parameters.AddWithValue("@km", n2); cmd.Parameters.AddWithValue("@cz", n3); cmd.Parameters.AddWithValue("@kr", n4); cmd.Parameters.AddWithValue("@hz", n5); cmd.Parameters.AddWithValue("@st", n6); cmd.Parameters.AddWithValue("@pb", n7); cmd.Parameters.AddWithValue("@id", rI); cmd.ExecuteNonQuery();
                        }
                    }
                    foreach (DataGridViewRow rX in aG.Rows) { if (rX.Cells["Sipariş No"].Value.ToString() == sN) { if (aG.Columns.Contains("Dedektif (Durum)")) { rX.Cells["Dedektif (Durum)"].Value = "✔ MÜTABIK"; } } }
                    aG.Refresh(); MessageBox.Show("Başarıyla güncellendi!"); f.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            f.Controls.Add(bU); f.ShowDialog();
        }

        private void DynBtnOdemeyiMuhurle_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(dynTxtNetGelenPara.Text) || string.IsNullOrWhiteSpace(dynTxtSiparisListesi.Text)) return;
            try
            {
                double tutar = ExcelHucredenSayiya(dynTxtNetGelenPara.Text); string plat = "AMAZON"; string tar = dynDtTarih.Value.ToString("dd.MM.yyyy");
                var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList();
                List<string> temiz = new List<string>(); StringBuilder mükerrer = new StringBuilder();
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    foreach (var s in sips)
                    {
                        using (var cmd = new SqliteCommand(@"SELECT s.Tarih, s.SatisFiyati FROM Siparisler s WHERE s.SiparisNo=@sn AND s.Durum='TAHSİL EDİLDİ' AND UPPER(s.PlatformAd) LIKE '%AMAZON%'", b))
                        {
                            cmd.Parameters.AddWithValue("@sn", s);
                            using (var r = cmd.ExecuteReader())
                            {
                                if (r.Read()) mükerrer.AppendLine($"⚠️ Sipariş No: {s} | Giriş Tarihi: {r["Tarih"]} | Sipariş Tutarı: {r["SatisFiyati"]} TL");
                                else temiz.Add(s);
                            }
                        }
                    }
                }
                if (temiz.Count == 0 && mükerrer.Length > 0) { FormMükerrerDetayGoster(mükerrer.ToString(), 0); return; }
                if (temiz.Count > 0)
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open();
                        using (var cmd = new SqliteCommand("INSERT INTO PlatformTahsilatlari (Platform, Tarih, Tutar, SiparisAdedi, SiparisNumaralari) VALUES (@p, @t, @v, @a, @n)", b))
                        {
                            cmd.Parameters.AddWithValue("@p", plat); cmd.Parameters.AddWithValue("@t", tar); cmd.Parameters.AddWithValue("@v", tutar); cmd.Parameters.AddWithValue("@a", temiz.Count); cmd.Parameters.AddWithValue("@n", string.Join(", ", temiz)); cmd.ExecuteNonQuery();
                        }
                        foreach (var s in temiz)
                        {
                            using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='TAHSİL EDİLDİ' WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%AMAZON%'", b))
                            {
                                cmd.Parameters.AddWithValue("@sn", s); cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
                dynTxtNetGelenPara.Clear(); dynTxtSiparisListesi.Clear(); sonCozumlenenTablo = null; dynBtnYukluDosyaAc.Visible = false; TahsilatGecmisiniListele(); CanliBilançoListele();
                if (mükerrer.Length > 0) FormMükerrerDetayGoster(mükerrer.ToString(), temiz.Count);
                else MessageBox.Show("Ödeme Başarıyla Kaydedildi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void DynBtnKaydiSil_Click(object sender, EventArgs e)
        {
            if (dynDgvTahsilatGecmisi.SelectedRows.Count == 0) return; int id = Convert.ToInt32(dynDgvTahsilatGecmisi.SelectedRows[0].Cells["Id"].Value);
            if (MessageBox.Show("Bu tahsilat kaydını iptal etmek istediğinize emin misiniz?", "İptal Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var tx = b.BeginTransaction()) { string nolar = ""; using (var cmd = new SqliteCommand("SELECT SiparisNumaralari FROM PlatformTahsilatlari WHERE Id = @id", b, tx)) { cmd.Parameters.AddWithValue("@id", id); object sonuc = cmd.ExecuteScalar(); if (sonuc != null) nolar = sonuc.ToString(); } List<string> siparisler = nolar.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList(); int geriAlinanSayi = 0; string updateSql = "UPDATE Siparisler SET Durum = 'KARGOYA VERİLDİ' WHERE SiparisNo = @sn AND UPPER(PlatformAd) LIKE '%AMAZON%'"; using (var cmdUpd = new SqliteCommand(updateSql, b, tx)) { foreach (var no in siparisler) { cmdUpd.Parameters.Clear(); cmdUpd.Parameters.AddWithValue("@sn", no); geriAlinanSayi += cmdUpd.ExecuteNonQuery(); } } using (var cmdDel = new SqliteCommand("DELETE FROM PlatformTahsilatlari WHERE Id = @id", b, tx)) { cmdDel.Parameters.AddWithValue("@id", id); cmdDel.ExecuteNonQuery(); } tx.Commit(); MessageBox.Show($"Hakediş kaydı iptal edildi! Alacak sınıfına geri aktarılan sipariş: {geriAlinanSayi}", "İşlem Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information); TahsilatGecmisiniListele(); CanliBilançoListele(); } } } catch (Exception exSil) { MessageBox.Show(exSil.Message); }
            }
        }

        private void DynDgvTahsilatGecmisi_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return; int id = Convert.ToInt32(dynDgvTahsilatGecmisi.Rows[e.RowIndex].Cells["Id"].Value); string nolar = "";
            try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT SiparisNumaralari FROM PlatformTahsilatlari WHERE Id=@id", conn)) { cmd.Parameters.AddWithValue("@id", id); nolar = cmd.ExecuteScalar()?.ToString() ?? ""; } } } catch { }
            Form f = new Form { Text = "Makbuz Detayı", Size = new Size(380, 450), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
            f.Controls.Add(new Label { Text = "Bu hakedişteki sipariş numaraları:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) });
            f.Controls.Add(new TextBox { Location = new Point(20, 50), Size = new Size(320, 330), Multiline = true, ScrollBars = ScrollBars.Vertical, Text = nolar.Replace(", ", Environment.NewLine), ReadOnly = true, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, Font = new Font("Consolas", 10.5F) });
            f.ShowDialog();
        }

        private void TahsilatGecmisiniListele()
        {
            try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT Id, Tarih, Platform, Tutar as 'Yatan Net Tutar (TL)', SiparisAdedi as 'Kapatılan Sipariş' FROM PlatformTahsilatlari WHERE UPPER(Platform)='AMAZON' ORDER BY Id DESC", conn)) { DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader()); dynDgvTahsilatGecmisi.DataSource = dt; if (dynDgvTahsilatGecmisi.Columns["Id"] != null) dynDgvTahsilatGecmisi.Columns["Id"].Visible = false; dynDgvTahsilatGecmisi.ClearSelection(); } } } catch { }
        }

        private void CanliBilançoListele()
        {
            try
            {
                DataTable dt = new DataTable(); dt.Columns.Add("Kapatılan Sipariş (Adet)", typeof(int)); dt.Columns.Add("Net Tahsil Edilen TL", typeof(double)); dt.Columns.Add("Toplam Kesintiler TL", typeof(double)); dt.Columns.Add("Kalan Net Alacak TL", typeof(double));
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sqlBilanço = @"SELECT SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN (IFNULL(s.SatisFiyati,0) + IFNULL(s.PlatformIndirimi,0) - (IFNULL(s.Stopaj,0) + IFNULL(s.PosBedeli,0) + IFNULL(s.HizmetBedeli,0) + IFNULL(s.KomisyonTL,0) + IFNULL(s.CezaTutari,0)) - IFNULL(s.KargoUcreti,0)) ELSE 0 END), SUM(CASE WHEN s.Durum!='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(IFNULL(s.KomisyonTL,0)), SUM(IFNULL(s.HizmetBedeli,0)), SUM(IFNULL(s.CezaTutari,0)), SUM(IFNULL(s.SatisFiyati,0)), SUM(IFNULL(s.PlatformIndirimi,0)), SUM(IFNULL(s.Stopaj,0)), SUM(IFNULL(s.PosBedeli,0)), SUM(IFNULL(s.KargoUcreti,0)) FROM Siparisler s WHERE UPPER(s.PlatformAd) LIKE '%AMAZON%' AND UPPER(s.Durum) != 'İPTAL'";
                    string sqlKar = @"SELECT IFNULL(SUM(s.Adet), 0), IFNULL(SUM((s.SatisFiyati * s.Adet) + IFNULL(s.PlatformIndirimi,0) - (IFNULL(s.AlisFiyati,0) * s.Adet) - s.KomisyonTL - s.KargoUcreti - s.HizmetBedeli - s.Stopaj - s.PosBedeli - IFNULL(s.CezaTutari,0)), 0) as NetKar FROM Siparisler s WHERE UPPER(s.PlatformAd) LIKE '%AMAZON%' AND UPPER(s.Durum) != 'İPTAL'";
                    double netKarTutar = 0; int karSiparisAdet = 0;
                    using (var cmdKar = new SqliteCommand(sqlKar, conn)) using (var rKar = cmdKar.ExecuteReader()) { if (rKar.Read()) { karSiparisAdet = Convert.ToInt32(rKar[0]); netKarTutar = Convert.ToDouble(rKar[1]); } }
                    using (var cmd = new SqliteCommand(sqlBilanço, conn)) using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read())
                        {
                            int odenenAdet = r[0] != DBNull.Value ? Convert.ToInt32(r[0]) : 0;
                            double odenenPara = r[1] != DBNull.Value ? Convert.ToDouble(r[1]) : 0;
                            int bekleyenAdet = r[2] != DBNull.Value ? Convert.ToInt32(r[2]) : 0;
                            double totKomisyon = r[3] != DBNull.Value ? Convert.ToDouble(r[3]) : 0;
                            double totHizmet = r[4] != DBNull.Value ? Convert.ToDouble(r[4]) : 0;
                            double totCeza = r[5] != DBNull.Value ? Convert.ToDouble(r[5]) : 0;
                            double totSatis = r[6] != DBNull.Value ? Convert.ToDouble(r[6]) : 0;
                            double kesintiler = totKomisyon + totHizmet + totCeza;
                            double totIndirim = r[7] != DBNull.Value ? Convert.ToDouble(r[7]) : 0;
                            double totStopaj = r[8] != DBNull.Value ? Convert.ToDouble(r[8]) : 0;
                            double totPos = r[9] != DBNull.Value ? Convert.ToDouble(r[9]) : 0;
                            double totKargo = r[10] != DBNull.Value ? Convert.ToDouble(r[10]) : 0;
                            double toplamNetHakedisMatrisi = totSatis + totIndirim - (totStopaj + totPos + totKomisyon + totHizmet + totCeza + totKargo);
                            double bakiyeKalanNetAlacak = toplamNetHakedisMatrisi - odenenPara; double toplamBrutCiro = totSatis; dt.Rows.Add(odenenAdet, odenenPara, kesintiler, bakiyeKalanNetAlacak);
                            lblCardTahsilEdilenAdet.Text = $"{odenenAdet} Gelen Ödemeler"; lblCardTahsilEdilenPara.Text = odenenPara.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardBekleyenAdet.Text = $"{odenenAdet + bekleyenAdet} İşlem Kayıtlı"; lblCardBekleyenPara.Text = kesintiler.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardToplamAdet.Text = $"{bekleyenAdet} Sipariş Beklemede"; lblCardToplamPara.Text = bakiyeKalanNetAlacak.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardCiroAdet.Text = "Mağaza Toplam Hacmi"; lblCardCiroPara.Text = toplamBrutCiro.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardKarAdet.Text = $"{karSiparisAdet} Siparişin Net Kârı"; lblCardKarPara.Text = netKarTutar.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                        }
                    }
                }
                dgvPlatformBilanço.DataSource = dt; if (dgvPlatformBilanço.Columns.Contains("Net Tahsil Edilen TL")) dgvPlatformBilanço.Columns["Net Tahsil Edilen TL"].DefaultCellStyle.Format = "N2"; if (dgvPlatformBilanço.Columns.Contains("Toplam Kesintiler TL")) dgvPlatformBilanço.Columns["Toplam Kesintiler TL"].DefaultCellStyle.Format = "N2"; if (dgvPlatformBilanço.Columns.Contains("Kalan Net Alacak TL")) dgvPlatformBilanço.Columns["Kalan Net Alacak TL"].DefaultCellStyle.Format = "N2";
            }
            catch { }
        }
    }
}