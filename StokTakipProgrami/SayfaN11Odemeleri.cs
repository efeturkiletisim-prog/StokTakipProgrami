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
    public class N11DbSiparis
    {
        public string Durum { get; set; }
        public double Satis { get; set; }
        public double Komisyon { get; set; }
        public double Kargo { get; set; }
        public double Hizmet { get; set; }
        public double Ceza { get; set; }
    }

    public partial class SayfaN11Odemeleri : UserControl
    {
        private DateTimePicker dynDtTarih;
        private TextBox dynTxtNetGelenPara, dynTxtSiparisListesi;
        private Button dynBtnOdemeyiMuhurle, dynBtnKaydiSil, dynBtnExcel, dynBtnSiparisSorgula, dynBtnTemizle, dynBtnYukluDosyaAc, dynBtnKargoGirisleri;
        private DataGridView dynDgvTahsilatGecmisi, dgvPlatformBilanço;
        private Label lblCardTahsilEdilenAdet, lblCardTahsilEdilenPara, lblCardBekleyenAdet, lblCardBekleyenPara, lblCardToplamAdet, lblCardToplamPara, lblCardCiroAdet, lblCardCiroPara, lblCardKarAdet, lblCardKarPara;
        private DataTable sonCozumlenenTablo = null;
        private string sonSiparisKolonAdi = "";
        private string sonIslemKolonAdi = "";
        private string sonHakedisKolonAdi = "";

        public SayfaN11Odemeleri() { InitializeComponent(); this.Load += SayfaN11Odemeleri_Load; }

        private void SayfaN11Odemeleri_Load(object sender, EventArgs e) { TabloyuGuvenceyeAl(); ArayuzInsaEt(); TahsilatGecmisiniListele(); CanliBilançoListele(); }

        private void TabloyuGuvenceyeAl() { try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS PlatformTahsilatlari (Id INTEGER PRIMARY KEY AUTOINCREMENT, Platform TEXT, Tarih TEXT, Tutar REAL, SiparisAdedi INTEGER, SiparisNumaralari TEXT);", conn)) { cmd.ExecuteNonQuery(); } } } catch { } }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear(); this.Dock = DockStyle.Fill; this.BackColor = Color.FromArgb(33, 33, 36);
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(194, 24, 91) };
            pnlHeader.Controls.Add(new Label { Text = "🟣 N11 MUTABAKAT VE HAKEDİŞ ÖDEMELERİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter }); this.Controls.Add(pnlHeader);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380F)); mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); this.Controls.Add(mainLayout);

            Panel pC = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15), AutoScroll = true }; int cY = 12;
            pC.Controls.Add(new Label { Text = "Ödeme Tarihi:", Location = new Point(15, cY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            dynDtTarih = new DateTimePicker { Location = new Point(15, cY + 20), Width = 340, Font = new Font("Segoe UI", 9.5F), Format = DateTimePickerFormat.Short }; pC.Controls.Add(dynDtTarih); cY += 50;

            pC.Controls.Add(FormElemaniYarat("Hesaba Yatan Toplam Tutar (TL):", out dynTxtNetGelenPara, new Point(15, cY), 340)); cY += 50;
            pC.Controls.Add(new Label { Text = "Sipariş Numaraları:", Location = new Point(15, cY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            dynTxtSiparisListesi = new TextBox { Location = new Point(15, cY + 20), Width = 340, Height = 85, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 9.5F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White }; pC.Controls.Add(dynTxtSiparisListesi); cY += 110;

            dynBtnTemizle = new Button { Text = "🧹 TEMİZLE", Location = new Point(15, cY), Width = 95, Height = 30, BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.LightGray, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand }; dynBtnTemizle.FlatAppearance.BorderSize = 0; dynBtnTemizle.Click += (s, e) => { dynTxtSiparisListesi.Clear(); dynTxtNetGelenPara.Clear(); }; pC.Controls.Add(dynBtnTemizle);
            dynBtnSiparisSorgula = new Button { Text = "🔍 SİPARİŞ SORGULA", Location = new Point(116, cY), Width = 239, Height = 30, BackColor = Color.FromArgb(194, 24, 91), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand }; dynBtnSiparisSorgula.FlatAppearance.BorderSize = 0; dynBtnSiparisSorgula.Click += DynBtnSiparisSorgula_Click; pC.Controls.Add(dynBtnSiparisSorgula); cY += 40;

            pC.Controls.Add(new Panel { Location = new Point(15, cY), Width = 340, Height = 1, BackColor = Color.FromArgb(45, 45, 48) }); cY += 10;
            pC.Controls.Add(new Label { Text = "Ekstre Yükle:", Location = new Point(15, cY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) }); cY += 22;

            dynBtnYukluDosyaAc = new Button { Text = "📂 YÜKLÜ DOSYAYI TEKRAR AÇ", Location = new Point(15, cY), Width = 340, Height = 28, BackColor = Color.FromArgb(38, 38, 41), ForeColor = Color.Gold, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Visible = false, Cursor = Cursors.Hand }; dynBtnYukluDosyaAc.FlatAppearance.BorderSize = 1; dynBtnYukluDosyaAc.FlatAppearance.BorderColor = Color.Gold; dynBtnYukluDosyaAc.Click += (s, e) => { if (sonCozumlenenTablo != null) MutabakatUyariEkraniniAc(sonCozumlenenTablo, "N11", sonSiparisKolonAdi, sonIslemKolonAdi, sonHakedisKolonAdi); }; pC.Controls.Add(dynBtnYukluDosyaAc);

            cY += 35;

            Button btnYukle = new Button { Text = "📂 EKSTRE SEÇ VE YÜKLE", Location = new Point(15, cY), Width = 340, Height = 42, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnYukle.FlatAppearance.BorderSize = 0; btnYukle.Click += DynBtnDosyaYukle_Click; pC.Controls.Add(btnYukle);

            cY += 62;

            pC.Controls.Add(new Panel { Location = new Point(15, cY), Width = 340, Height = 1, BackColor = Color.FromArgb(50, 50, 55) });

            cY += 20;

            // 1. ÖNCE "ÖDEMEYİ KAYDET" BUTONU
            dynBtnOdemeyiMuhurle = new Button { Text = "✅ ÖDEMEYİ KAYDET", Location = new Point(15, cY), Width = 340, Height = 42, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnOdemeyiMuhurle.FlatAppearance.BorderSize = 0; dynBtnOdemeyiMuhurle.Click += DynBtnOdemeyiMuhurle_Click; pC.Controls.Add(dynBtnOdemeyiMuhurle);

            cY += 65; // 🚀 MÜHÜR: İstediğin gibi az daha aşağı kaydırıldı (Boşluk 10px'den 23px'e çıkarıldı)

            // 2. HEMEN ALTINA "KARGO GİRİŞLERİ RADARI" BUTONU
            dynBtnKargoGirisleri = new Button { Text = "📦 KARGO GİRİŞLERİ RADARI", Location = new Point(15, cY), Width = 340, Height = 42, BackColor = Color.FromArgb(41, 128, 185), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnKargoGirisleri.FlatAppearance.BorderSize = 0; dynBtnKargoGirisleri.Click += DynBtnKargoGirisleri_Click; pC.Controls.Add(dynBtnKargoGirisleri);

            mainLayout.Controls.Add(pC, 0, 0);

            TableLayoutPanel pRL = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) }; pRL.RowStyles.Add(new RowStyle(SizeType.Percent, 45F)); pRL.RowStyles.Add(new RowStyle(SizeType.Percent, 35F)); pRL.RowStyles.Add(new RowStyle(SizeType.Absolute, 95F)); mainLayout.Controls.Add(pRL, 1, 0);
            Panel pUW = new Panel { Dock = DockStyle.Fill }; pUW.Controls.Add(new Label { Text = "📋 SİSTEME İŞLENEN TOPLU ÖDEMELER", Dock = DockStyle.Top, Height = 22, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray }); dynDgvTahsilatGecmisi = GridYarat(); dynDgvTahsilatGecmisi.CellDoubleClick += DynDgvTahsilatGecmisi_CellDoubleClick;
            dynDgvTahsilatGecmisi.DataBindingComplete += (s, e) => { dynDgvTahsilatGecmisi.ClearSelection(); };
            pUW.Controls.Add(dynDgvTahsilatGecmisi); dynDgvTahsilatGecmisi.BringToFront();

            Panel pRB = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(0, 5, 0, 0) }; dynBtnKaydiSil = new Button { Text = "🗑️ SEÇİLİ ÖDEMEYİ İPTAL ET", Dock = DockStyle.Left, Width = 380, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) }; dynBtnKaydiSil.FlatAppearance.BorderSize = 0; dynBtnKaydiSil.Click += DynBtnKaydiSil_Click; pRB.Controls.Add(dynBtnKaydiSil);
            dynBtnExcel = new Button { Text = "📊 EXCEL RAPORU AL", Dock = DockStyle.Right, Width = 130, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) }; dynBtnExcel.FlatAppearance.BorderSize = 0; dynBtnExcel.Click += (s, e) => { /* ExcelMotoru.DataGridViewTablosunuDisaAktar(dynDgvTahsilatGecmisi, "N11_Hakedis_Gecmisi"); */ }; pRB.Controls.Add(dynBtnExcel); pUW.Controls.Add(pRB); pRL.Controls.Add(pUW, 0, 0);

            GroupBox grpB = new GroupBox { Dock = DockStyle.Fill, Text = "📊 N11 CARİ MUTABAKAT VE HESAP TABLOSU", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange, Padding = new Padding(10, 22, 10, 10) }; dgvPlatformBilanço = GridYarat(); grpB.Controls.Add(dgvPlatformBilanço); pRL.Controls.Add(grpB, 0, 1);

            TableLayoutPanel tVIP = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, Padding = new Padding(0, 5, 0, 5) }; for (int i = 0; i < 5; i++) tVIP.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            CustomVIPCardInsaEt(tVIP, "💰 TOPLAM TAHSİLAT", Color.FromArgb(34, 197, 94), 0, out lblCardTahsilEdilenAdet, out lblCardTahsilEdilenPara); CustomVIPCardInsaEt(tVIP, "⏳ KESİNTİLER", Color.Orange, 1, out lblCardBekleyenAdet, out lblCardBekleyenPara); CustomVIPCardInsaEt(tVIP, "💎 KALAN NET ALACAK", Color.FromArgb(56, 189, 248), 2, out lblCardToplamAdet, out lblCardToplamPara); CustomVIPCardInsaEt(tVIP, "📊 TOPLAM CİRO", Color.FromArgb(168, 85, 247), 3, out lblCardCiroAdet, out lblCardCiroPara); CustomVIPCardInsaEt(tVIP, "📈 NET KÂR DURUMU", Color.FromArgb(236, 72, 153), 4, out lblCardKarAdet, out lblCardKarPara); pRL.Controls.Add(tVIP, 0, 2); mainLayout.BringToFront(); pnlHeader.SendToBack();
        }

        private Panel FormElemaniYarat(string b, out TextBox k, Point l, int w, string v = "") { Panel p = new Panel { Location = l, Width = w, Height = 52 }; p.Controls.Add(new Label { Text = b, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray }); k = new TextBox { Text = v, Location = new Point(0, 18), Width = w, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(k); return p; }
        private Panel FormElemaniYaratRenk(string b, out TextBox k, Point l, int w, string v, Color c) { Panel p = new Panel { Location = l, Width = w, Height = 48 }; p.Controls.Add(new Label { Text = b, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = c }); k = new TextBox { Text = v, Location = new Point(0, 16), Width = w, Font = new Font("Segoe UI", 9.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(k); return p; }
        private DataGridView GridYarat() { DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9F), BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false, ColumnHeadersHeight = 32 }; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold); dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgv.DefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DataError += (s, e) => { e.Cancel = true; }; return dgv; }
        private void CustomVIPCardInsaEt(TableLayoutPanel t, string b, Color c, int s, out Label lA, out Label lP) { Panel pC = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(22, 22, 26), Margin = new Padding(6), Padding = new Padding(1) }; pC.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 3, BackColor = c }); Panel pT = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 6, 6, 6), BackColor = Color.FromArgb(26, 26, 30) }; Label lblB = new Label { Text = b.ToUpper(CultureInfo.CurrentCulture), Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(160, 160, 170), TextAlign = ContentAlignment.MiddleLeft }; lP = new Label { Text = "0,00 ₺", Dock = DockStyle.Top, Height = 26, Font = new Font("Segoe UI Black", 13.5F, FontStyle.Bold), ForeColor = c, TextAlign = ContentAlignment.MiddleLeft }; lA = new Label { Text = "0 Sipariş", Dock = DockStyle.Top, Height = 18, Font = new Font("Segoe UI", 8.0F), ForeColor = Color.FromArgb(110, 110, 120), TextAlign = ContentAlignment.MiddleLeft }; pT.Controls.Add(lA); pT.Controls.Add(lP); pT.Controls.Add(lblB); lblB.BringToFront(); lP.BringToFront(); lA.BringToFront(); pC.Controls.Add(pT); t.Controls.Add(pC, s, 0); }
        private bool PromptForInput(string t, string c, out string v) { Form p = new Form() { Width = 350, Height = 180, FormBorderStyle = FormBorderStyle.FixedDialog, Text = c, StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false }; TextBox ib = new TextBox() { Left = 20, Top = 50, Width = 290, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, Font = new Font("Segoe UI", 11F) }; Button cf = new Button() { Text = "Onayla", Left = 210, Width = 100, Top = 90, BackColor = Color.FromArgb(194, 24, 91), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) }; string r = ""; bool ok = false; cf.Click += (s, e) => { r = ib.Text.Trim(); ok = true; p.Close(); }; p.Controls.Add(new Label() { Left = 20, Top = 20, Text = t, AutoSize = true, ForeColor = Color.White }); p.Controls.Add(ib); p.Controls.Add(cf); p.AcceptButton = cf; p.ShowDialog(); v = r; return ok; }
        private static void FormMükerrerDetayGoster(string rapor, int temiz) { Form f = new Form { Text = "🛑 Çapraz Mutabakat Mükerrer Koruma", Size = new Size(680, 420), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false }; Label l = new Label { Text = $"⚠️ Aşağıdaki siparişler daha önceden TAHSİL EDİLGİSİ için tekrar kaydedilmedi.\nTemiz olan {temiz} sipariş mühürlendi!", Dock = DockStyle.Top, Height = 50, ForeColor = Color.LightCoral, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Padding = new Padding(5) }; TextBox t = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, BackColor = Color.FromArgb(24, 24, 27), ForeColor = Color.Orange, Font = new Font("Consolas", 10F), Text = rapor }; f.Controls.Add(t); f.Controls.Add(l); f.ShowDialog(); }

        private double ExcelHucredenSayiya(object h) { if (h == null || h == DBNull.Value) return 0; if (h is double d) return d; if (h is decimal dc) return Convert.ToDouble(dc); if (h is int i) return i; string s = h.ToString().Replace("₺", "").Replace("TL", "").Replace(" ", "").Replace("\xA0", "").Trim(); if (string.IsNullOrEmpty(s)) return 0; int n = s.LastIndexOf('.'); int v = s.LastIndexOf(','); if (n > v) { if (s.Substring(n).Length <= 3) s = s.Substring(0, n).Replace(".", "").Replace(",", "") + "." + s.Substring(n + 1); else s = s.Replace(".", "").Replace(",", ""); } else if (v > n) { if (s.Substring(v).Length <= 3) s = s.Substring(0, v).Replace(".", "").Replace(",", "") + "." + s.Substring(v + 1); else s = s.Replace(".", "").Replace(",", ""); } else s = s.Replace(".", "").Replace(",", ""); if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double rf)) return rf; return 0; }
        private string[] SatiriParcala(string s, char a) { var r = new List<string>(); bool t = false; string h = ""; for (int i = 0; i < s.Length; i++) { char c = s[i]; if (c == '"') t = !t; else if (c == a && !t) { r.Add(h.Trim('"')); h = ""; } else h += c; } r.Add(h.Trim('"')); return r.ToArray(); }
        private double GetVal(DataRow r, params string[] cols) { foreach (string n in cols) { foreach (DataColumn c in r.Table.Columns) { string b = c.ColumnName.ToUpper().Replace(" ", "").Replace("_", "").Trim(); if (b.Contains(n) && r[c] != DBNull.Value && !string.IsNullOrWhiteSpace(r[c].ToString())) return ExcelHucredenSayiya(r[c]); } } return 0; }
        private string GetStr(DataRow r, params string[] cols) { foreach (string n in cols) { foreach (DataColumn c in r.Table.Columns) { string b = c.ColumnName.ToUpper().Replace(" ", "").Replace("_", "").Trim(); if (b.Contains(n) && r[c] != DBNull.Value && !string.IsNullOrWhiteSpace(r[c].ToString())) return r[c].ToString().Trim(); } } return ""; }

        private DataTable N11XmlSpreadsheetOku(string xml)
        {
            DataTable dt = new DataTable(); string[] sheets = xml.Split(new string[] { "<Worksheet " }, StringSplitOptions.RemoveEmptyEntries); Dictionary<string, int> colMap = new Dictionary<string, int>();
            for (int i = 1; i < sheets.Length; i++)
            {
                string sheet = sheets[i]; if (sheet.IndexOf("Name=\"Summary\"", StringComparison.OrdinalIgnoreCase) >= 0 || sheet.IndexOf("Name=\"Kılavuz\"", StringComparison.OrdinalIgnoreCase) >= 0 || sheet.IndexOf("Name=\"Ozet\"", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                int ts = sheet.IndexOf("<Table"); int te = sheet.IndexOf("</Table>"); if (ts == -1 || te == -1) continue;
                string content = sheet.Substring(ts, te - ts); string[] rows = content.Split(new string[] { "<Row", "</Row>" }, StringSplitOptions.RemoveEmptyEntries); bool first = true; int[] lMap = null;
                foreach (string row in rows)
                {
                    if (!row.Contains("<Cell")) continue; string[] cells = row.Split(new string[] { "<Cell", "</Cell>" }, StringSplitOptions.RemoveEmptyEntries); List<string> rowData = new List<string>();
                    foreach (string cell in cells) { if (!cell.Contains("<Data")) { rowData.Add(""); continue; } int ds = cell.IndexOf(">", cell.IndexOf("<Data")) + 1; int de = cell.IndexOf("</Data>"); if (ds > 0 && de > ds) { string val = cell.Substring(ds, de - ds); val = Regex.Replace(val, "<.*?>", string.Empty).Replace("&#160;", " ").Replace("&amp;", "&").Trim(); rowData.Add(val); } else rowData.Add(""); }
                    if (rowData.Count > 0)
                    {
                        if (first) { lMap = new int[rowData.Count]; for (int c = 0; c < rowData.Count; c++) { string cN = string.IsNullOrWhiteSpace(rowData[c]) ? "Bos" : rowData[c].Trim(); if (!colMap.ContainsKey(cN)) { colMap[cN] = dt.Columns.Count; dt.Columns.Add(cN); } lMap[c] = colMap[cN]; } first = false; }
                        else { if (lMap == null) continue; DataRow dr = dt.NewRow(); for (int k = 0; k < rowData.Count; k++) { if (k < lMap.Length) dr[lMap[k]] = rowData[k]; } dt.Rows.Add(dr); }
                    }
                }
            }
            return dt;
        }

        private DataTable N11HtmlExcelOku(string html)
        {
            DataTable dt = new DataTable(); string[] rows = html.Split(new string[] { "<tr", "<TR" }, StringSplitOptions.None); bool first = true;
            for (int i = 1; i < rows.Length; i++)
            {
                string rHtml = rows[i]; var cList = new List<string>(); string[] tdS = rHtml.Split(new string[] { "<td", "<TD", "<th", "<TH" }, StringSplitOptions.None);
                for (int j = 1; j < tdS.Length; j++) { int start = tdS[j].IndexOf(">"); int end = tdS[j].IndexOf("</", start > 0 ? start : 0, StringComparison.OrdinalIgnoreCase); if (start != -1 && end != -1) { string val = tdS[j].Substring(start + 1, end - start - 1); val = Regex.Replace(val, "<.*?>", string.Empty).Replace("&nbsp;", " ").Trim(); cList.Add(val); } }
                if (cList.Count > 0) { if (first) { foreach (var c in cList) { string cN = string.IsNullOrWhiteSpace(c) ? "Bos" : c.Trim(); int sfx = 1; string fN = cN; while (dt.Columns.Contains(fN)) { fN = cN + sfx++; } dt.Columns.Add(fN); } first = false; } else { while (cList.Count > dt.Columns.Count) dt.Columns.Add("Col" + dt.Columns.Count); dt.Rows.Add(cList.ToArray()); } }
            }
            return dt;
        }

        private DataTable ExcelVeCsvOku(string dosyaYolu)
        {
            DataTable dt = new DataTable(); string ext = Path.GetExtension(dosyaYolu).ToLower(); string icerik = "";
            try { byte[] bytes = File.ReadAllBytes(dosyaYolu); if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) icerik = Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3); else icerik = Encoding.GetEncoding(1254).GetString(bytes); } catch { }
            if (icerik.StartsWith("<?xml") && icerik.Contains("Workbook")) return N11XmlSpreadsheetOku(icerik);
            if (icerik.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0 || icerik.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0) return N11HtmlExcelOku(icerik);
            if (ext == ".xls" || ext == ".xlsx")
            {
                try { string connStr = ext == ".xls" ? $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dosyaYolu};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\";" : $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dosyaYolu};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";"; using (OleDbConnection conn = new OleDbConnection(connStr)) { conn.Open(); DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null); string tS = ""; foreach (DataRow r in schema.Rows) { string sN = r["TABLE_NAME"].ToString(); if (sN.ToUpper().Contains("SATIŞ") || sN.ToUpper().Contains("SATIS")) { tS = sN; break; } } if (string.IsNullOrEmpty(tS) && schema.Rows.Count > 0) tS = schema.Rows[0]["TABLE_NAME"].ToString(); if (tS.StartsWith("'") && tS.EndsWith("'")) tS = tS.Substring(1, tS.Length - 2); using (OleDbDataAdapter da = new OleDbDataAdapter($"SELECT * FROM [{tS}]", conn)) { da.Fill(dt); } } if (dt.Rows.Count > 0) return dt; } catch { }

                try { Type excelType = Type.GetTypeFromProgID("Excel.Application"); if (excelType != null) { dynamic excelApp = Activator.CreateInstance(excelType); excelApp.Visible = false; excelApp.DisplayAlerts = false; dynamic workbook = excelApp.Workbooks.Open(dosyaYolu); dynamic targetSheet = null; foreach (dynamic sheet in workbook.Sheets) { string sN = sheet.Name.ToUpper(); if (sN.Contains("SATIŞ") || sN.Contains("SATIS")) { targetSheet = sheet; break; } } if (targetSheet == null) targetSheet = workbook.Sheets[1]; dynamic range = targetSheet.UsedRange; int rows = range.Rows.Count; int cols = range.Columns.Count; object[,] values = range.Value2 as object[,]; if (values != null && rows > 0 && cols > 0) { for (int c = 1; c <= cols; c++) { string cN = values[1, c] != null ? values[1, c].ToString().Trim() : "Bos"; int s = 1; string fN = cN; while (dt.Columns.Contains(fN)) { fN = cN + s++; } dt.Columns.Add(fN); } for (int r = 2; r <= rows; r++) { DataRow dr = dt.NewRow(); bool hasData = false; for (int c = 1; c <= cols; c++) { object v = values[r, c]; if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) hasData = true; dr[c - 1] = v != null ? v.ToString().Trim() : ""; } if (hasData) dt.Rows.Add(dr); } } workbook.Close(false); excelApp.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); if (dt.Rows.Count > 0) return dt; } } catch { }
            }
            try { string[] lines = icerik.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None); if (lines.Length > 0) { char ayrac = lines[0].Contains(";") ? ';' : (lines[0].Contains("\t") ? '\t' : ','); string[] headers = SatiriParcala(lines[0], ayrac); foreach (string h in headers) { string cN = string.IsNullOrWhiteSpace(h) ? "Bos" : h.Trim(); int s = 1; string fN = cN; while (dt.Columns.Contains(fN)) { fN = cN + s++; } dt.Columns.Add(fN); } for (int i = 1; i < lines.Length; i++) { if (string.IsNullOrWhiteSpace(lines[i])) continue; string[] cells = SatiriParcala(lines[i], ayrac); DataRow row = dt.NewRow(); for (int j = 0; j < cells.Length && j < dt.Columns.Count; j++) row[j] = cells[j].Trim(); dt.Rows.Add(row); } } } catch { }
            return dt;
        }

        private void DynBtnDosyaYukle_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Tüm Formatlar|*.xlsx;*.xls;*.csv;*.txt";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        DataTable dt = ExcelVeCsvOku(ofd.FileName);
                        if (dt == null || dt.Columns.Count == 0 || dt.Rows.Count == 0) { MessageBox.Show("Sistem dosyayı okuyamadı. Excel dosyasını CSV olarak kaydedip yüklemeyi deneyin.", "Okuma Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                        int idxIslem = -1, idxSip = -1, idxHak = -1;
                        int dataStartRow = 0;

                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            string b = dt.Columns[c].ColumnName.ToUpper().Replace(" ", "").Replace("_", "").Trim();
                            if (b == "İŞLEMTİPİ" || b == "ISLEMTIPI" || b == "KAYİTTİPİ" || b == "KAYITTIPI" || b == "İŞLEMTİP" || b == "ISLEMTIP") idxIslem = c;
                            else if (b == "SİPARİŞNO" || b == "SIPARISNO" || b == "SİPARİŞNUMARASI" || b == "KAYITNO") idxSip = c;
                            else if (b == "NETSATIŞTUTARI" || b == "NETSATISTUTARI" || b == "SATICIHAKEDİŞ" || b == "NETTUTAR" || b == "HAKEDİŞ" || b == "HAKEDIS" || b == "NETTRANSFERTUTARI") idxHak = c;
                        }

                        if (idxSip == -1 || idxHak == -1)
                        {
                            for (int r = 0; r < Math.Min(dt.Rows.Count, 15); r++)
                            {
                                for (int c = 0; c < dt.Columns.Count; c++)
                                {
                                    string b = dt.Rows[r][c]?.ToString().ToUpper().Replace(" ", "").Replace("_", "").Trim() ?? "";
                                    if (b == "İŞLEMTİPİ" || b == "ISLEMTIPI" || b == "KAYİTTİPİ" || b == "KAYITTIPI" || b == "İŞLEMTİP" || b == "ISLEMTIP") idxIslem = c;
                                    else if (b == "SİPARİŞNO" || b == "SIPARISNO" || b == "SİPARİŞNUMARASI" || b == "KAYITNO") idxSip = c;
                                    else if (b == "NETSATIŞTUTARI" || b == "NETSATISTUTARI" || b == "SATICIHAKEDİŞ" || b == "NETTUTAR" || b == "HAKEDİŞ" || b == "HAKEDIS" || b == "NETTRANSFERTUTARI") idxHak = c;
                                }
                                if (idxSip != -1 && idxHak != -1) { dataStartRow = r + 1; break; }
                            }
                        }

                        if (idxSip == -1 || idxHak == -1) { MessageBox.Show("Yüklenen ekstre dosyasındaki hakediş sütun başlıkları sistem matrisiyle uyuşmuyor!"); return; }

                        var sips = new HashSet<string>(); double total = 0; var excelOzet = new Dictionary<string, double[]>();

                        for (int i = dataStartRow; i < dt.Rows.Count; i++)
                        {
                            int doluHucre = 0; for (int c = 0; c < dt.Columns.Count; c++) { if (!string.IsNullOrWhiteSpace(dt.Rows[i].ItemArray[c]?.ToString())) doluHucre++; }
                            if (doluHucre < 3) continue;

                            string sn = dt.Rows[i].ItemArray[idxSip].ToString().Trim();
                            if (string.IsNullOrEmpty(sn) || sn == "-" || sn.ToUpper().Contains("TOPLAM") || sn.ToUpper().Contains("ÖZET") || sn.ToUpper().Contains("TOTAL") || sn.ToUpper().Contains("SİPARİŞ")) continue;
                            if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                            if (!Regex.IsMatch(sn, @"^\d+$")) continue;

                            sips.Add(sn); if (!excelOzet.ContainsKey(sn)) excelOzet[sn] = new double[6];

                            string islemTxt = idxIslem != -1 ? dt.Rows[i].ItemArray[idxIslem].ToString().ToUpper(new CultureInfo("tr-TR")).Replace(" ", "") : "";
                            double orjTutar = ExcelHucredenSayiya(dt.Rows[i].ItemArray[idxHak]);
                            double mutlakTutar = Math.Abs(orjTutar);

                            double satis = GetVal(dt.Rows[i], "TUTAR", "BRÜTSATIŞ", "BRUTSATIS");
                            double kom = GetVal(dt.Rows[i], "KOMİSYON", "KOMISYON");
                            double paz = GetVal(dt.Rows[i], "PAZARLAMA");
                            double py = GetVal(dt.Rows[i], "PAZARYERİ", "PAZARYERI");
                            double hizm = Math.Abs(paz) + Math.Abs(py);

                            if (islemTxt.Contains("SATIŞ") || islemTxt.Contains("SATIS"))
                            {
                                excelOzet[sn][0] += satis;
                                excelOzet[sn][1] += Math.Abs(kom);
                                excelOzet[sn][2] += hizm;
                                excelOzet[sn][5] += orjTutar;
                            }
                            else if (islemTxt.Contains("KARGO") || islemTxt.Contains("TESLİMAT")) { excelOzet[sn][3] += mutlakTutar; excelOzet[sn][5] += orjTutar; }
                            else if (islemTxt.Contains("HURDA") || islemTxt.Contains("TAZMİN") || islemTxt.Contains("HASAR")) { excelOzet[sn][4] += mutlakTutar; excelOzet[sn][5] += orjTutar; }
                            else if (islemTxt.Contains("CEZA") || islemTxt.Contains("İHLAL") || islemTxt.Contains("GECİKME")) { excelOzet[sn][4] += mutlakTutar; excelOzet[sn][5] += orjTutar; }
                            else { excelOzet[sn][5] += orjTutar; }

                            total += orjTutar;
                        }

                        var dbHavuz = new Dictionary<string, N11DbSiparis>();
                        using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                        {
                            conn.Open();
                            foreach (var s in sips)
                            {
                                using (var cmd = new SqliteCommand("SELECT Durum, SatisFiyati, KomisyonTL, KargoUcreti, HizmetBedeli, IFNULL(CezaTutari,0) FROM Siparisler WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%N11%'", conn))
                                {
                                    cmd.Parameters.AddWithValue("@sn", s); using (var r = cmd.ExecuteReader()) { if (r.Read()) dbHavuz[s] = new N11DbSiparis { Durum = r[0].ToString(), Satis = Convert.ToDouble(r[1]), Komisyon = Convert.ToDouble(r[2]), Kargo = Convert.ToDouble(r[3]), Hizmet = Convert.ToDouble(r[4]), Ceza = Convert.ToDouble(r[5]) }; }
                                }
                            }
                        }

                        dt.Columns.Add("Dedektif (Durum)", typeof(string));
                        dt.Columns.Add("Sistem Durumu", typeof(string));

                        int hatalar = 0;
                        var finalTable = dt.Clone();

                        foreach (DataRow row in dt.Rows)
                        {
                            string sn = row[idxSip]?.ToString().Trim() ?? "";
                            if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                            if (sips.Contains(sn))
                            {
                                string ded = "✔ MÜTABIK";
                                string sys = "-";

                                if (dbHavuz.TryGetValue(sn, out var db))
                                {
                                    sys = db.Durum;
                                    if (sys.ToUpper() == "TAHSİL EDİLDİ") { ded = "HATA: MÜKERRER TAHSİLAT"; }
                                    else
                                    {
                                        if (Math.Abs(excelOzet[sn][0] - db.Satis) > 0.1 && excelOzet[sn][0] != 0) ded = "HATA: SATIŞ FİYATI FARKI";
                                        else if (Math.Abs(excelOzet[sn][1] - db.Komisyon) > 0.1 && excelOzet[sn][1] != 0) ded = "HATA: KOMİSYON FARKI";
                                        else if (Math.Abs(excelOzet[sn][3] - db.Kargo) > 0.1 && excelOzet[sn][3] != 0) ded = "HATA: KARGO FARKI";
                                        else if (Math.Abs(excelOzet[sn][2] - db.Hizmet) > 0.1 && excelOzet[sn][2] != 0) ded = "HATA: İŞLEM/HİZMET FARKI";
                                        else if (Math.Abs(excelOzet[sn][4] - db.Ceza) > 0.1 && excelOzet[sn][4] != 0) ded = "HATA: CEZA FARKI";
                                    }
                                }
                                else { sys = "SİPARİŞLER SİSTEMDE YOK"; ded = "HATA: KAYITSIZ"; }

                                if (ded.StartsWith("HATA:")) hatalar++;

                                DataRow newRow = finalTable.NewRow();
                                newRow.ItemArray = row.ItemArray.Clone() as object[];
                                newRow["Dedektif (Durum)"] = ded;
                                newRow["Sistem Durumu"] = sys;
                                finalTable.Rows.Add(newRow);
                            }
                        }

                        dynTxtSiparisListesi.Text = string.Join(Environment.NewLine, sips); dynTxtNetGelenPara.Text = total.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));
                        sonCozumlenenTablo = finalTable;
                        sonSiparisKolonAdi = finalTable.Columns[idxSip].ColumnName;
                        sonIslemKolonAdi = idxIslem != -1 ? finalTable.Columns[idxIslem].ColumnName : "";
                        sonHakedisKolonAdi = finalTable.Columns[idxHak].ColumnName;

                        dynBtnYukluDosyaAc.Text = $"📂 YÜKLÜ DOSYAYI AÇ ({Path.GetFileName(ofd.FileName)})"; dynBtnYukluDosyaAc.Visible = true;

                        if (hatalar > 0) MessageBox.Show($"🚨 DİKKAT! MÜKERRER VEYA HATALI {hatalar} SATIRDA TAHSİLAT FARKI TESPİT EDİLDİ!", "Radar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        MutabakatUyariEkraniniAc(finalTable, "N11", sonSiparisKolonAdi, sonIslemKolonAdi, sonHakedisKolonAdi);
                    }
                    catch (Exception exDosya) { MessageBox.Show("Sistem Hatası: " + exDosya.Message); }
                }
            }
        }

        private void MutabakatUyariEkraniniAc(DataTable dtSource, string dp, string sipCol, string islCol, string hakCol)
        {
            Form f = new Form { Text = $"{dp} Ekstre Önizleme", Size = new Size(1350, 650), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
            Label lbl = new Label { Text = "SİPARİŞ DETAYLARI İNCELEME.", Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Orange, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };

            TableLayoutPanel tlpPopupLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(33, 33, 36) };
            tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            f.Controls.Add(tlpPopupLayout); tlpPopupLayout.Controls.Add(lbl, 0, 0);

            Panel pnlGridContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 10) };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), SelectionMode = DataGridViewSelectionMode.CellSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EnableHeadersVisualStyles = false, RowHeadersVisible = false };
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DefaultCellStyle.SelectionForeColor = Color.White; dgv.ColumnHeadersHeight = 40; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgv.CellDoubleClick += (s, eDouble) =>
            {
                if (eDouble.RowIndex < 0 || eDouble.ColumnIndex < 0) return; string col = dgv.Columns[eDouble.ColumnIndex].Name; var r = dgv.Rows[eDouble.RowIndex]; string sn = r.Cells[sipCol].Value?.ToString() ?? "";
                if (col == sipCol && !string.IsNullOrEmpty(sn) && sn != "-") { Clipboard.SetText(sn); MessageBox.Show($"Kopyalandı: {sn}"); return; }
                if (!string.IsNullOrEmpty(sn) && sn != "-")
                {
                    double tH = 0; foreach (DataGridViewRow rX in dgv.Rows) { if (rX.Cells[sipCol].Value.ToString() == sn) { string tp = islCol != "" ? rX.Cells[islCol].Value.ToString().ToUpper() : "SATIŞ"; if (tp.Contains("SATIŞ") || tp.Contains("SİPARİŞ") || tp.Contains("SIPARIS") || tp.Contains("TUTAR") || tp.Contains("GENEL") || tp.Contains("SATIS")) { tH += ExcelHucredenSayiya(rX.Cells[hakCol].Value); } } }
                    SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);
                }
            };

            dgv.DataBindingComplete += (s, eBind) =>
            {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    string type = islCol != "" && dgv.Columns.Contains(islCol) ? row.Cells[islCol].Value?.ToString().ToUpper() ?? "" : "SATIŞ";
                    string ded = dgv.Columns.Contains("Dedektif (Durum)") ? row.Cells["Dedektif (Durum)"].Value?.ToString() ?? "" : "";
                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); row.DefaultCellStyle.SelectionForeColor = Color.White;

                    if (ded == "KAYDEDİLDİ" || ded == "HURDA İŞLENDİ" || ded.StartsWith("✔ MÜTABIK")) { row.DefaultCellStyle.BackColor = Color.FromArgb(46, 117, 89); row.DefaultCellStyle.ForeColor = Color.White; }
                    else if (ded.StartsWith("HATA:")) { row.DefaultCellStyle.BackColor = Color.Crimson; row.DefaultCellStyle.ForeColor = Color.White; }
                    else { row.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); if (type.Contains("SATIŞ") || type.Contains("SIPARIS") || type.Contains("SİPARİŞ") || type.Contains("TUTAR") || type.Contains("GENEL") || type.Contains("SATIS")) row.DefaultCellStyle.ForeColor = Color.DarkGray; else row.DefaultCellStyle.ForeColor = Color.White; }
                }
            };

            ContextMenuStrip cms = new ContextMenuStrip(); dgv.ContextMenuStrip = cms;
            dgv.MouseDown += (s, eMouse) =>
            {
                if (eMouse.Button == MouseButtons.Right)
                {
                    var hit = dgv.HitTest(eMouse.X, eMouse.Y); if (hit.Type == DataGridViewHitTestType.Cell && hit.RowIndex >= 0)
                    {
                        dgv.ClearSelection(); dgv.Rows[hit.RowIndex].Cells[hit.ColumnIndex].Selected = true; var r = dgv.Rows[hit.RowIndex]; string sn = r.Cells[sipCol].Value.ToString(); string islemTipi = islCol != "" && dgv.Columns.Contains(islCol) ? r.Cells[islCol].Value.ToString().ToUpper() : "SATIŞ"; cms.Items.Clear();

                        ToolStripMenuItem btnAmeliyatTY = new ToolStripMenuItem("✏️ İncele ve Manuel Düzelt");
                        btnAmeliyatTY.Click += (sx, exAmeliyatTY) => {
                            if (string.IsNullOrEmpty(sn) || sn == "-") return; double tH = 0; foreach (DataGridViewRow rX in dgv.Rows) { if (rX.Cells[sipCol].Value.ToString() == sn) { string tp = islCol != "" ? rX.Cells[islCol].Value.ToString().ToUpper() : "SATIŞ"; if (tp.Contains("SATIŞ") || tp.Contains("SATIS") || tp.Contains("KUPON") || tp.Contains("İNDİRİM") || tp.Contains("GENEL")) { tH += ExcelHucredenSayiya(rX.Cells[hakCol].Value); } } }
                            SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);
                        }; cms.Items.Add(btnAmeliyatTY);

                        bool isSatisVeyaKuponTY = islemTipi.Contains("SATIŞ") || islemTipi.Contains("SATIS") || islemTipi.Contains("KUPON") || islemTipi.Contains("İNDİRİM") || islemTipi.Contains("GENEL");
                        if (!isSatisVeyaKuponTY)
                        {
                            cms.Items.Add(new ToolStripSeparator()); ToolStripMenuItem btnCezaTY = new ToolStripMenuItem("⚠️ 'Ceza Tutarı' Olarak İşle");
                            btnCezaTY.Click += (sx, exCezaTY) => {
                                string targetSn = sn; if (string.IsNullOrEmpty(targetSn) || targetSn == "-") { if (PromptForInput("Sipariş No girin:", "Sipariş No", out string inpSn)) targetSn = inpSn; else return; }
                                double tut = ExcelHucredenSayiya(r.Cells[hakCol].Value);
                                if (MessageBox.Show($"[{targetSn}] siparişe {Math.Abs(tut):N2} TL ceza eklenecektir?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET CezaTutari=IFNULL(CezaTutari,0)+@c, CezaSebebi=CASE WHEN IFNULL(CezaSebebi,'')='' THEN @s ELSE CezaSebebi||' | '||@s END WHERE SiparisNo=@sn", b)) { cmd.Parameters.AddWithValue("@c", Math.Abs(tut)); cmd.Parameters.AddWithValue("@s", islemTipi); cmd.Parameters.AddWithValue("@sn", targetSn); cmd.ExecuteNonQuery(); } } foreach (DataGridViewRow gridRow in dgv.Rows) { if (gridRow.Cells[sipCol].Value.ToString() == targetSn) { if (dgv.Columns.Contains("Dedektif (Durum)")) gridRow.Cells["Dedektif (Durum)"].Value = "KAYDEDİLDİ"; } } dgv.Refresh(); }
                                    catch (Exception exTY) { MessageBox.Show(exTY.Message); }
                                }
                            }; cms.Items.Add(btnCezaTY);

                            ToolStripMenuItem btnErkenTY = new ToolStripMenuItem("⚡ 'Erken Ödeme Kesintisi' Olarak İşle");
                            btnErkenTY.Click += (sx, exErkenTY) => {
                                double tut = Math.Abs(ExcelHucredenSayiya(r.Cells[hakCol].Value));
                                if (MessageBox.Show($"{tut:N2} TL Erken Ödeme olarak düşülecek?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                {
                                    try { string yeniSip = "KESİNTİ-" + DateTime.Now.ToString("yyyyMMddHHmmss"); using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("INSERT INTO Siparisler (SiparisNo, PlatformAd, UrunAdi, Adet, SatisFiyati, CezaTutari, CezaSebebi, Tarih, Durum) VALUES (@sn, 'N11', 'Erken Ödeme Kesintisi', 0, 0, @c, 'Erken Ödeme', @t, 'BEKLEYEN')", b)) { cmd.Parameters.AddWithValue("@sn", yeniSip); cmd.Parameters.AddWithValue("@c", tut); cmd.Parameters.AddWithValue("@t", dynDtTarih.Value.ToString("dd.MM.yyyy")); cmd.ExecuteNonQuery(); } } if (dgv.Columns.Contains("Sistem Durumu")) r.Cells["Sistem Durumu"].Value = "ERKEN ÖDEME İŞLENDİ"; if (dgv.Columns.Contains("Dedektif (Durum)")) r.Cells["Dedektif (Durum)"].Value = "KAYDEDİLDİ"; dgv.Refresh(); if (!string.IsNullOrEmpty(dynTxtSiparisListesi.Text)) dynTxtSiparisListesi.Text += Environment.NewLine; dynTxtSiparisListesi.Text += yeniSip; }
                                    catch (Exception exTY2) { MessageBox.Show(exTY2.Message); }
                                }
                            }; cms.Items.Add(btnErkenTY);
                        }
                        cms.Show(Cursor.Position);
                    }
                }
            };

            Panel pBtn = new Panel { Dock = DockStyle.Fill, Height = 50 };
            Button btnC = new Button { Text = "Ödemeleri Onayla ve Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(194, 24, 91), ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnC.FlatAppearance.BorderSize = 0; pBtn.SizeChanged += (sB, eB) => { btnC.Location = new Point((pBtn.Width - btnC.Width) / 2, (pBtn.Height - btnC.Height) / 2); };
            btnC.Click += (sB, eB) => {
                bool hataliVar = false; foreach (DataGridViewRow rx in dgv.Rows) { if (dgv.Columns.Contains("Dedektif (Durum)") && rx.Cells["Dedektif (Durum)"].Value.ToString().StartsWith("HATA:")) { hataliVar = true; break; } }
                if (hataliVar) { MessageBox.Show("Hatalı (Kırmızı) kayıtlar var! Lütfen sağ tıklayıp Ceza/Hurda olarak işleyin veya çift tıklayıp manuel düzeltin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                f.DialogResult = DialogResult.OK; f.Close();
            }; pBtn.Controls.Add(btnC);
            pnlGridContainer.Controls.Add(dgv); tlpPopupLayout.Controls.Add(pnlGridContainer, 0, 1); tlpPopupLayout.Controls.Add(pBtn, 0, 2);
            dgv.BringToFront(); dgv.DataSource = dtSource; f.ShowDialog();
        }

        private void DynBtnSiparisSorgula_Click(object sender, EventArgs e)
        {
            var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList();
            if (sips.Count == 0) { MessageBox.Show("Lütfen sipariş numarası giriniz.", "Sipariş Numarası Eksik", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (sips.Count > 1000) sips = sips.Take(1000).ToList();

            Form frmPopup = new Form { Text = "⚖️ N11 MANUEL MUTABAKAT VE ÖDEME RADARI (Sorgulanan: " + sips.Count + " Sipariş)", Size = new Size(1150, 600), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };

            TableLayoutPanel tlpPopupLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(33, 33, 36) };
            tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); frmPopup.Controls.Add(tlpPopupLayout);

            Panel pnlTopInfo = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28) };
            pnlTopInfo.Controls.Add(new Label { Text = "📢 'N11 Ödenen (TL)' sütununa gelen net hakediş tutarlarını yazıp Enter'a basınız. Sipariş satırına ÇİFT TIKLAYARAK ameliyat masasını açabilirsiniz.", Dock = DockStyle.Fill, ForeColor = Color.Gold, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter });
            tlpPopupLayout.Controls.Add(pnlTopInfo, 0, 0);

            Panel pnlGridContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 10) };
            DataGridView dgvPopup = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, Font = new Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular), EnableHeadersVisualStyles = false, BorderStyle = BorderStyle.None, ColumnHeadersHeight = 35 };
            dgvPopup.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgvPopup.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgvPopup.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgvPopup.DefaultCellStyle.ForeColor = Color.White;

            dgvPopup.Columns.Add("SiparisNo", "Sipariş No"); dgvPopup.Columns.Add("Durum", "Sistem Durumu"); dgvPopup.Columns.Add("Urun", "Kayıtlı Ürün"); dgvPopup.Columns.Add("Ciro", "Brüt Satış (TL)"); dgvPopup.Columns.Add("Kesinti", "Kesintiler (TL)"); dgvPopup.Columns.Add("Kargo", "Kargo (TL)"); dgvPopup.Columns.Add("Hakedis", "Sistem Hakedişi (TL)"); dgvPopup.Columns.Add("GelenTutar", "N11 Ödenen (TL)"); dgvPopup.Columns.Add("Mutabakat", "Mutabakat Kontrolü");

            for (int i = 0; i < dgvPopup.Columns.Count; i++) if (i != 7) dgvPopup.Columns[i].ReadOnly = true;
            dgvPopup.Columns[3].DefaultCellStyle.Format = "N2"; dgvPopup.Columns[4].DefaultCellStyle.Format = "N2"; dgvPopup.Columns[5].DefaultCellStyle.Format = "N2"; dgvPopup.Columns[6].DefaultCellStyle.Format = "N2"; dgvPopup.Columns[7].DefaultCellStyle.Format = "N2"; dgvPopup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvPopup.CellDoubleClick += (sGrid, eGrid) => {
                if (eGrid.RowIndex < 0) return; string selectedSipNo = dgvPopup.Rows[eGrid.RowIndex].Cells["SiparisNo"].Value?.ToString();
                if (!string.IsNullOrEmpty(selectedSipNo) && selectedSipNo != "-") SiparisiAmeliyatMasasinaYatir(selectedSipNo, 0, 0, dgvPopup);
            };

            pnlGridContainer.Controls.Add(dgvPopup); tlpPopupLayout.Controls.Add(pnlGridContainer, 0, 1);

            Panel pnlPopupBtn = new Panel { Dock = DockStyle.Fill };
            Button btnPopupClose = new Button { Text = "Mutabakat Kontrolünü Onayla ve Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(194, 24, 91), ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnPopupClose.FlatAppearance.BorderSize = 0; btnPopupClose.Click += (sB, eB) => { frmPopup.DialogResult = DialogResult.OK; frmPopup.Close(); };
            btnPopupClose.Location = new Point((pnlPopupBtn.Width - btnPopupClose.Width) / 2, (pnlPopupBtn.Height - btnPopupClose.Height) / 2); pnlPopupBtn.SizeChanged += (sB, eB) => { btnPopupClose.Location = new Point((pnlPopupBtn.Width - btnPopupClose.Width) / 2, (pnlPopupBtn.Height - btnPopupClose.Height) / 2); };
            pnlPopupBtn.Controls.Add(btnPopupClose); tlpPopupLayout.Controls.Add(pnlPopupBtn, 0, 2);

            using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
            {
                conn.Open();
                foreach (string sNo in sips)
                {
                    string query = @"SELECT UrunAdi, IFNULL(s.SatisFiyati,0) as Satis, IFNULL(s.PlatformIndirimi,0) as Indirim, IFNULL(s.Stopaj,0) as Stopaj, IFNULL(s.PosBedeli,0) as Pos, IFNULL(s.HizmetBedeli,0) as Hizmet, IFNULL(s.KomisyonTL,0) as Kom, IFNULL(s.KargoUcreti,0) as Kargo, IFNULL(s.CezaTutari,0) as Ceza, s.Durum FROM Siparisler s WHERE s.SiparisNo = @sn AND UPPER(s.PlatformAd) LIKE '%N11%' AND UPPER(s.Durum) != 'İPTAL'";
                    using (var cmd = new SqliteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@sn", sNo);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                string sysDurum = r["Durum"].ToString(); double satis = Convert.ToDouble(r["Satis"]); double kesintiler = Convert.ToDouble(r["Stopaj"]) + Convert.ToDouble(r["Pos"]) + Convert.ToDouble(r["Hizmet"]) + Convert.ToDouble(r["Kom"]) + Convert.ToDouble(r["Ceza"]); double kargo = Convert.ToDouble(r["Kargo"]); double sHakedis = satis + Convert.ToDouble(r["Indirim"]) - kesintiler - kargo;

                                if (sysDurum.ToUpper() == "TAHSİL EDİLDİ") { int rowIdx = dgvPopup.Rows.Add(sNo, "❌ MÜKERRER TAHSİLAT", r["UrunAdi"].ToString(), satis, kesintiler, kargo, sHakedis, 0.00, "❌ DAHA ÖNCE ÖDENDİ!"); dgvPopup.Rows[rowIdx].DefaultCellStyle.BackColor = Color.FromArgb(75, 20, 20); dgvPopup.Rows[rowIdx].DefaultCellStyle.ForeColor = Color.LightCoral; }
                                else { int rowIdx = dgvPopup.Rows.Add(sNo, "✔ " + sysDurum, r["UrunAdi"].ToString(), satis, kesintiler, kargo, sHakedis, 0.00, "✍ Tutar Bekleniyor..."); dgvPopup.Rows[rowIdx].DefaultCellStyle.BackColor = Color.FromArgb(38, 45, 38); }
                            }
                            else { int rowIdx = dgvPopup.Rows.Add(sNo, "❌ BULUNAMADI", "Sistemde bu sipariş numarası yok!", 0.00, 0.00, 0.00, 0.00, 0.00, "Bulunamadı"); dgvPopup.Rows[rowIdx].DefaultCellStyle.BackColor = Color.FromArgb(55, 25, 25); dgvPopup.Rows[rowIdx].DefaultCellStyle.ForeColor = Color.LightCoral; }
                        }
                    }
                }
            }

            dgvPopup.CellValueChanged += (s, ev) =>
            {
                if (ev.RowIndex < 0 || ev.ColumnIndex != 7) return; DataGridViewRow row = dgvPopup.Rows[ev.RowIndex];
                if (row.Cells["Durum"].Value.ToString().Contains("BULUNAMADI") || row.Cells["Durum"].Value.ToString().Contains("MÜKERRER")) { row.Cells["GelenTutar"].Value = 0.00; return; }

                double sHakedis = Convert.ToDouble(row.Cells["Hakedis"].Value); double gelenTutar = 0;
                if (row.Cells["GelenTutar"].Value != null) double.TryParse(row.Cells["GelenTutar"].Value.ToString(), out gelenTutar);

                if (Math.Abs(sHakedis - gelenTutar) < 0.001) { row.Cells["Mutabakat"].Value = "✔ TAM UYUŞUYOR"; row.Cells["Mutabakat"].Style.ForeColor = Color.LimeGreen; row.Cells["Mutabakat"].Style.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold); row.DefaultCellStyle.BackColor = Color.FromArgb(20, 50, 20); }
                else { row.Cells["Mutabakat"].Value = "❌ UYUŞMAZLIK HATASI!"; row.Cells["Mutabakat"].Style.ForeColor = Color.Crimson; row.Cells["Mutabakat"].Style.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold); row.DefaultCellStyle.BackColor = Color.FromArgb(65, 20, 20); }
            };

            frmPopup.ShowDialog();
        }

        private void SiparisiAmeliyatMasasinaYatir(string sipNo, double exH, double exK, DataGridView anaGrid)
        {
            long rId = 0; double sF = 0, km = 0, kr = 0, hz = 0, st = 0, ps = 0, cz = 0; string durum = "YOK", brkd = "-", uAdi = "Bilinmeyen Ürün", kat = "-", mrk = "-", tar = "-"; int adet = 1; double alisFiyati = 0; string imei = "-";
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open(); string q = @"SELECT s.rowid, s.SatisFiyati, s.KomisyonTL, s.KargoUcreti, s.HizmetBedeli, s.Stopaj, s.PosBedeli, IFNULL(s.CezaTutari,0), s.Durum, s.Barkod, IFNULL(u.UrunAdi, s.UrunAdi), IFNULL(u.Kategori, s.Kategori), IFNULL(u.Marka, s.Marka), IFNULL(s.Adet,1), s.Tarih, IFNULL(s.AlisFiyati,0), IFNULL(s.SatilanImei,'-') FROM Siparisler s LEFT JOIN Urunler u ON s.Barkod = u.Barkod WHERE s.SiparisNo = @sn AND UPPER(s.PlatformAd) LIKE '%N11%'";
                    using (var cmd = new SqliteCommand(q, conn)) { cmd.Parameters.AddWithValue("@sn", sipNo); using (var r = cmd.ExecuteReader()) { if (r.Read()) { rId = Convert.ToInt64(r[0]); sF = Convert.ToDouble(r[1]); km = Convert.ToDouble(r[2]); kr = Convert.ToDouble(r[3]); hz = Convert.ToDouble(r[4]); st = Convert.ToDouble(r[5]); ps = Convert.ToDouble(r[6]); cz = Convert.ToDouble(r[7]); durum = r[8].ToString(); brkd = r[9].ToString(); uAdi = r[10].ToString(); kat = r[11].ToString(); mrk = r[12].ToString(); adet = Convert.ToInt32(r[13]); tar = r[14].ToString(); alisFiyati = Convert.ToDouble(r[15]); imei = r[16].ToString(); } else return; } }
                }
            }
            catch { return; }
            Form frm = new Form { Text = $"Düzeltme Masası - No: {sipNo}", Size = new Size(520, 660), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };

            Panel pnlMeta = new Panel { Location = new Point(15, 10), Width = 475, Height = 105, BackColor = Color.FromArgb(24, 24, 27), BorderStyle = BorderStyle.FixedSingle };
            string temizBarkod = (string.IsNullOrEmpty(brkd) || brkd.ToUpper() == uAdi.ToUpper()) ? "-" : brkd;
            Label lblDurumX = new Label { Text = "● SİSTEM DURUMU: " + durum.ToUpper(), Location = new Point(15, 12), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.Gold };
            Label lblKimlikX = new Label { Text = $"📦 Ürün: {uAdi} (Adet: {adet})\r\n🆔 Barkod: {temizBarkod} | Kategori: {kat} | Marka: {mrk}", Location = new Point(15, 36), Size = new Size(445, 38), Font = new Font("Segoe UI", 9F, FontStyle.Regular), ForeColor = Color.FromArgb(56, 189, 248) };
            Label lblFinansX = new Label { Text = $"💰 Kayıtlı Finansal Geçmiş -> Tarih: {tar} | Alış: {alisFiyati.ToString("N2") + "₺ | IMEI: " + imei}", Location = new Point(15, 78), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(34, 197, 94) };
            pnlMeta.Controls.Add(lblDurumX); pnlMeta.Controls.Add(lblKimlikX); pnlMeta.Controls.Add(lblFinansX); frm.Controls.Add(pnlMeta);

            int currentY = 125;
            GroupBox grpGelir = new GroupBox { Text = " Satış Gelirleri", Location = new Point(15, currentY), Width = 475, Height = 70, ForeColor = Color.FromArgb(34, 197, 94), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            TextBox tSat; grpGelir.Controls.Add(FormElemaniYaratRenk("Satış Fiyatı (TL):", out tSat, new Point(15, 18), 440, sF.ToString("N2"), Color.White)); frm.Controls.Add(grpGelir); currentY += 75;
            GroupBox grpMaliyet = new GroupBox { Text = " Ürün Maliyet & Envanter Bilgisi", Location = Point.Empty, Width = 475, Height = 70, ForeColor = Color.FromArgb(56, 189, 248), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            grpMaliyet.Location = new Point(15, currentY);
            TextBox tAlis, tImei; grpMaliyet.Controls.Add(FormElemaniYaratRenk("Alış Maliyeti (TL):", out tAlis, new Point(15, 18), 210, alisFiyati.ToString("N2"), Color.White)); grpMaliyet.Controls.Add(FormElemaniYaratRenk("IMEI Numarası:", out tImei, new Point(245, 18), 210, imei, Color.White)); frm.Controls.Add(grpMaliyet); currentY += 75;
            GroupBox grpGider = new GroupBox { Text = " Kesintiler & Platform Vergileri", Location = new Point(15, currentY), Width = 475, Height = 195, ForeColor = Color.FromArgb(239, 68, 68), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
            TextBox tKoms, tCez, tKrg, tHzm, tStpj, tPos;
            grpGider.Controls.Add(FormElemaniYaratRenk("Komisyon (TL):", out tKoms, new Point(15, 18), 210, km.ToString("N2"), Color.White)); grpGider.Controls.Add(FormElemaniYaratRenk("Kargo Ücreti (TL):", out tKrg, new Point(245, 18), 210, kr.ToString("N2"), Color.White));
            grpGider.Controls.Add(FormElemaniYaratRenk("Hizmet Bedeli (TL):", out tHzm, new Point(15, 74), 210, hz.ToString("N2"), Color.White)); grpGider.Controls.Add(FormElemaniYaratRenk("Ceza Tutarı (TL):", out tCez, new Point(245, 74), 210, cz.ToString("N2"), Color.White));
            grpGider.Controls.Add(FormElemaniYaratRenk("Stopaj (TL):", out tStpj, new Point(15, 130), 210, st.ToString("N2"), Color.White)); grpGider.Controls.Add(FormElemaniYaratRenk("POS Kesintisi (TL):", out tPos, new Point(245, 130), 210, ps.ToString("N2"), Color.White));
            frm.Controls.Add(grpGider); currentY += 205;

            frm.Controls.Add(new Label { Text = $"💡 Brüt Hakediş: {exH:N2} ₺", Location = new Point(15, currentY), AutoSize = true, ForeColor = Color.LightSkyBlue, Font = new Font("Segoe UI", 9F, FontStyle.Bold) });
            Button btn = new Button { Text = "💾 GÜNCELLE", Location = new Point(15, currentY + 25), Width = 475, Height = 45, BackColor = Color.FromArgb(194, 24, 91), ForeColor = Color.White, Font = new Font("Segoe UI", 10F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => {
                try
                {
                    double.TryParse(tSat.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n1); double.TryParse(tKoms.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n2); double.TryParse(tKrg.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n4); double.TryParse(tHzm.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n5); double.TryParse(tStpj.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n6); double.TryParse(tPos.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n7); double.TryParse(tCez.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n3); double.TryParse(tAlis.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nAlis); string nImei = tImei.Text.Trim();
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET SatisFiyati=@sf, AlisFiyati=@af, SatilanImei=@im, KomisyonTL=@km, KargoUcreti=@kr, HizmetBedeli=@hb, Stopaj=@st, PosBedeli=@pb, CezaTutari=@cz WHERE rowid=@id", b)) { cmd.Parameters.AddWithValue("@sf", n1); cmd.Parameters.AddWithValue("@af", nAlis); cmd.Parameters.AddWithValue("@im", nImei); cmd.Parameters.AddWithValue("@km", n2); cmd.Parameters.AddWithValue("@kr", n4); cmd.Parameters.AddWithValue("@hb", n5); cmd.Parameters.AddWithValue("@st", n6); cmd.Parameters.AddWithValue("@pb", n7); cmd.Parameters.AddWithValue("@cz", n3); cmd.Parameters.AddWithValue("@id", rId); cmd.ExecuteNonQuery(); } }

                    string gridSipCol = "";
                    foreach (DataGridViewColumn c in anaGrid.Columns) { if (c.Name.ToUpper().Contains("SİPARİŞ") || c.Name.ToUpper().Contains("SIPARIS") || c.Name.ToUpper().Contains("KOD")) { gridSipCol = c.Name; break; } }
                    if (gridSipCol != "") { foreach (DataGridViewRow rowX in anaGrid.Rows) { if (rowX.Cells[gridSipCol].Value != null && rowX.Cells[gridSipCol].Value.ToString() == sipNo) { if (anaGrid.Columns.Contains("Dedektif (Durum)")) rowX.Cells["Dedektif (Durum)"].Value = "✔ MÜTABIK"; if (anaGrid.Columns.Contains("Dedektif (Kargo Durumu)")) rowX.Cells["Dedektif (Kargo Durumu)"].Value = "✔ MÜTABIK"; } } }

                    anaGrid.Refresh(); MessageBox.Show("Başarıyla güncellendi!", "İşlem Tamam", MessageBoxButtons.OK, MessageBoxIcon.Information); frm.Close();
                }
                catch (Exception exAmeliyat) { MessageBox.Show(exAmeliyat.Message); }
            };
            frm.Controls.Add(btn); frm.ShowDialog();
        }

        private void TahsilatGecmisiniListele()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Id, Tarih, Platform, Tutar as 'Yatan Net Tutar (TL)', SiparisAdedi as 'Kapatılan Sipariş' FROM PlatformTahsilatlari WHERE UPPER(Platform)='N11' ORDER BY Id DESC", conn))
                    {
                        DataTable dt = new DataTable();
                        dt.Load(cmd.ExecuteReader());
                        dynDgvTahsilatGecmisi.DataSource = dt;

                        if (dynDgvTahsilatGecmisi.Columns["Id"] != null)
                            dynDgvTahsilatGecmisi.Columns["Id"].Visible = false;

                        dynDgvTahsilatGecmisi.ClearSelection();
                    }
                }
            }
            catch { }
        }

        private void CanliBilançoListele()
        {
            try
            {
                DataTable dt = new DataTable(); dt.Columns.Add("Kapatılan Sipariş (Adet)", typeof(int)); dt.Columns.Add("Net Tahsil Edilen TL", typeof(double)); dt.Columns.Add("Toplam Kesintiler TL", typeof(double)); dt.Columns.Add("Kalan Net Alacak TL", typeof(double));
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sqlBilanço = @"SELECT SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN (IFNULL(s.SatisFiyati,0) + IFNULL(s.PlatformIndirimi,0) - (IFNULL(s.Stopaj,0) + IFNULL(s.PosBedeli,0) + IFNULL(s.HizmetBedeli,0) + IFNULL(s.KomisyonTL,0) + IFNULL(s.CezaTutari,0)) - IFNULL(s.KargoUcreti,0)) ELSE 0 END), SUM(CASE WHEN s.Durum!='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(IFNULL(s.KomisyonTL,0)), SUM(IFNULL(s.HizmetBedeli,0)), SUM(IFNULL(s.CezaTutari,0)), SUM(IFNULL(s.SatisFiyati,0)), SUM(IFNULL(s.PlatformIndirimi,0)), SUM(IFNULL(s.Stopaj,0)), SUM(IFNULL(s.PosBedeli,0)), SUM(IFNULL(s.KargoUcreti,0)) FROM Siparisler s WHERE UPPER(s.PlatformAd) LIKE '%N11%' AND UPPER(s.Durum) != 'İPTAL'";
                    string sqlKar = @"SELECT IFNULL(SUM(s.Adet), 0), IFNULL(SUM((s.SatisFiyati * s.Adet) + IFNULL(s.PlatformIndirimi,0) - (IFNULL(s.AlisFiyati,0) * s.Adet) - s.KomisyonTL - s.KargoUcreti - s.HizmetBedeli - s.Stopaj - s.PosBedeli - IFNULL(s.CezaTutari,0)), 0) as NetKar FROM Siparisler s WHERE UPPER(s.PlatformAd) LIKE '%N11%' AND UPPER(s.Durum) != 'İPTAL'";
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

        private void DynBtnKaydiSil_Click(object sender, EventArgs e)
        {
            if (dynDgvTahsilatGecmisi.SelectedRows.Count == 0) return; int id = Convert.ToInt32(dynDgvTahsilatGecmisi.SelectedRows[0].Cells["Id"].Value);
            if (MessageBox.Show("Bu tahsilat kaydını iptal etmek istediğinize emin misiniz?", "Hakediş İptal Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                    {
                        b.Open(); using (var tx = b.BeginTransaction())
                        {
                            string nolar = "";
                            using (var cmd = new SqliteCommand("SELECT SiparisNumaralari FROM PlatformTahsilatlari WHERE Id = @id", b, tx)) { cmd.Parameters.AddWithValue("@id", id); object sonuc = cmd.ExecuteScalar(); if (sonuc != null) nolar = sonuc.ToString(); }
                            List<string> siparisler = nolar.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
                            int geriAlinanSayi = 0;
                            string updateSql = "UPDATE Siparisler SET Durum = 'KARGOYA VERİLDİ' WHERE SiparisNo = @sn AND UPPER(PlatformAd) LIKE '%N11%'";
                            using (var cmdUpd = new SqliteCommand(updateSql, b, tx))
                            {
                                foreach (var no in siparisler) { cmdUpd.Parameters.Clear(); cmdUpd.Parameters.AddWithValue("@sn", no); geriAlinanSayi += cmdUpd.ExecuteNonQuery(); }
                            }
                            using (var cmdDel = new SqliteCommand("DELETE FROM PlatformTahsilatlari WHERE Id = @id", b, tx)) { cmdDel.Parameters.AddWithValue("@id", id); cmdDel.ExecuteNonQuery(); }
                            tx.Commit();
                            MessageBox.Show($"Hakediş kaydı iptal edildi! Alacak sınıfına geri aktarılan sipariş: {geriAlinanSayi}", "İşlem Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            TahsilatGecmisiniListele(); CanliBilançoListele();
                        }
                    }
                }
                catch (Exception exSil) { MessageBox.Show(exSil.Message); }
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

        private void DynBtnOdemeyiMuhurle_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(dynTxtNetGelenPara.Text) || string.IsNullOrWhiteSpace(dynTxtSiparisListesi.Text)) return;
            try
            {
                double tutar = ExcelHucredenSayiya(dynTxtNetGelenPara.Text); string plat = "N11"; string tar = dynDtTarih.Value.ToString("dd.MM.yyyy");
                var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList();

                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open(); using (var tx = b.BeginTransaction())
                    {
                        string insertSql = "INSERT INTO PlatformTahsilatlari (Platform, Tarih, Tutar, SiparisAdedi, SiparisNumaralari) VALUES (@p, @tar, @tut, @adet, @nolar)";
                        using (var cmdIns = new SqliteCommand(insertSql, b, tx)) { cmdIns.Parameters.AddWithValue("@p", plat); cmdIns.Parameters.AddWithValue("@tar", tar); cmdIns.Parameters.AddWithValue("@tut", tutar); cmdIns.Parameters.AddWithValue("@adet", sips.Count); cmdIns.Parameters.AddWithValue("@nolar", string.Join(", ", sips)); cmdIns.ExecuteNonQuery(); }
                        int kapananSiparisSayisi = 0;
                        string updateSql = "UPDATE Siparisler SET Durum = 'TAHSİL EDİLDİ' WHERE SiparisNo = @sn AND UPPER(PlatformAd) LIKE '%N11%'";
                        using (var cmdUpd = new SqliteCommand(updateSql, b, tx)) { foreach (var no in sips) { cmdUpd.Parameters.Clear(); cmdUpd.Parameters.AddWithValue("@sn", no); kapananSiparisSayisi += cmdUpd.ExecuteNonQuery(); } }
                        tx.Commit();
                        MessageBox.Show($"Başarıyla Mühürlendi patron!\n\nPlatform: {plat}\nHesaba Geçen: {tutar:N2} TL\nKapatılan Sipariş: {kapananSiparisSayisi} Adet", "Tahsilat Tamam", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dynTxtNetGelenPara.Clear(); dynTxtSiparisListesi.Clear(); sonCozumlenenTablo = null; dynBtnYukluDosyaAc.Visible = false;
                        TahsilatGecmisiniListele(); CanliBilançoListele();
                    }
                }
            }
            catch (Exception exMuhur) { MessageBox.Show(exMuhur.Message); }
        }

        private void DynBtnKargoGirisleri_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Tüm Formatlar|*.xlsx;*.xls;*.csv;*.txt";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        DataTable dt = ExcelVeCsvOku(ofd.FileName);
                        if (dt == null || dt.Columns.Count == 0 || dt.Rows.Count == 0) { MessageBox.Show("Sistem kargo dosyasını okuyamadı.", "Okuma Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                        int idxSip = -1, idxKargo = -1;
                        int dataStartRow = 0;

                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            string b = dt.Columns[c].ColumnName.ToUpper().Replace(" ", "").Replace("_", "").Trim();
                            if (b.Contains("KOD") || b.Contains("SİPARİŞ") || b.Contains("SIPARIS") || b.Contains("KAYITNO")) idxSip = c;
                            if ((b.Contains("KARGO") || b.Contains("TUTAR") || b.Contains("BEDEL") || b.Contains("ÜCRET") || b.Contains("UCRET") || b.Contains("DESİ") || b.Contains("DESI") || b.Contains("TL") || b.Contains("KG"))
                                && !b.Contains("ŞİRKET") && !b.Contains("SIRKET") && !b.Contains("ADI") && !b.Contains("AD"))
                                idxKargo = c;
                        }

                        if (idxSip == -1 || idxKargo == -1)
                        {
                            for (int r = 0; r < Math.Min(dt.Rows.Count, 15); r++)
                            {
                                for (int c = 0; c < dt.Columns.Count; c++)
                                {
                                    string b = dt.Rows[r][c]?.ToString().ToUpper().Replace(" ", "").Replace("_", "").Trim() ?? "";
                                    if (b.Contains("KOD") || b.Contains("SİPARİŞ") || b.Contains("SIPARIS")) idxSip = c;
                                    if ((b.Contains("TUTAR") || b.Contains("KARGO") || b.Contains("DESİ") || b.Contains("TL")) && !b.Contains("ŞİRKET")) idxKargo = c;
                                }
                                if (idxSip != -1 && idxKargo != -1) { dataStartRow = r + 1; break; }
                            }
                        }

                        if (idxSip == -1 || idxKargo == -1) { MessageBox.Show("Kargo ekstre sütun başlıkları sistemle uyuşmuyor!"); return; }

                        var sips = new HashSet<string>();
                        var excelKargoOzet = new Dictionary<string, double>();

                        for (int i = dataStartRow; i < dt.Rows.Count; i++)
                        {
                            int doluHucre = 0; for (int c = 0; c < dt.Columns.Count; c++) { if (!string.IsNullOrWhiteSpace(dt.Rows[i].ItemArray[c]?.ToString())) doluHucre++; }
                            if (doluHucre < 3) continue;

                            string sn = dt.Rows[i].ItemArray[idxSip].ToString().Trim();
                            if (string.IsNullOrEmpty(sn) || sn == "-" || sn.ToUpper().Contains("TOPLAM") || sn.ToUpper().Contains("OZET")) continue;
                            if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);
                            if (!Regex.IsMatch(sn, @"^\d+$")) continue;

                            sips.Add(sn);

                            string rawKargo = dt.Rows[i].ItemArray[idxKargo].ToString().ToUpper();
                            if (rawKargo.Contains("/")) rawKargo = rawKargo.Split('/')[0];
                            if (rawKargo.Contains("TL")) rawKargo = rawKargo.Replace("TL", "");

                            double kargoBedeli = Math.Abs(ExcelHucredenSayiya(rawKargo.Trim()));
                            if (!excelKargoOzet.ContainsKey(sn)) excelKargoOzet[sn] = 0;
                            excelKargoOzet[sn] = kargoBedeli;
                        }

                        var dbHavuz = new Dictionary<string, double>();
                        using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                        {
                            conn.Open();
                            foreach (var s in sips)
                            {
                                using (var cmd = new SqliteCommand("SELECT IFNULL(KargoUcreti,0) FROM Siparisler WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%N11%'", conn))
                                {
                                    cmd.Parameters.AddWithValue("@sn", s);
                                    object val = cmd.ExecuteScalar();
                                    if (val != null) dbHavuz[s] = Convert.ToDouble(val);
                                }
                            }
                        }

                        dt.Columns.Add("Sistem Kargo Ücreti (TL)", typeof(double));
                        dt.Columns.Add("Dedektif (Kargo Durumu)", typeof(string));

                        int hatalar = 0;
                        var finalTable = dt.Clone();

                        foreach (DataRow row in dt.Rows)
                        {
                            string sn = row[idxSip]?.ToString().Trim() ?? "";
                            if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                            if (sips.Contains(sn))
                            {
                                string ded = "✔ MÜTABIK";
                                double dbKargo = 0;

                                if (dbHavuz.TryGetValue(sn, out dbKargo))
                                {
                                    if (Math.Abs(excelKargoOzet[sn] - dbKargo) > 0.1)
                                    {
                                        ded = "HATA: KARGO FARKI";
                                        hatalar++;
                                    }
                                }
                                else
                                {
                                    ded = "HATA: KAYITSIZ";
                                    hatalar++;
                                }

                                DataRow newRow = finalTable.NewRow();
                                newRow.ItemArray = row.ItemArray.Clone() as object[];
                                newRow["Sistem Kargo Ücreti (TL)"] = dbKargo;
                                newRow["Dedektif (Kargo Durumu)"] = ded;
                                finalTable.Rows.Add(newRow);
                            }
                        }

                        if (hatalar > 0) MessageBox.Show($"🚨 DİKKAT! Sistemle uyuşmayan {hatalar} adet kargo uyuşmazlığı tespit edildi!", "Kargo Kontrol Radarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        KargoUyariEkraniniAc(finalTable, "N11", finalTable.Columns[idxSip].ColumnName, finalTable.Columns[idxKargo].ColumnName);
                    }
                    catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
                }
            }
        }

        private void KargoUyariEkraniniAc(DataTable dtSource, string dp, string sipCol, string hakCol)
        {
            Form f = new Form { Text = $"{dp} Kargo Fatura Önizleme ve İnceleme Radarı", Size = new Size(1350, 650), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
            Label lbl = new Label { Text = "📦 GİRİLEN KARGO VERİLERİ VE SİSTEM EŞLEŞTİRME EKRANI (Çift Tık: Sipariş İnceleme Masası | Sağ Tık: Hızlı Eşitleme)", Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Orange, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };

            TableLayoutPanel tlpPopupLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(33, 33, 36) };
            tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            f.Controls.Add(tlpPopupLayout); tlpPopupLayout.Controls.Add(lbl, 0, 0);

            Panel pnlGridContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 10) };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), SelectionMode = DataGridViewSelectionMode.CellSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EnableHeadersVisualStyles = false, RowHeadersVisible = false };
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DefaultCellStyle.SelectionForeColor = Color.White; dgv.ColumnHeadersHeight = 40; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgv.CellDoubleClick += (s, eDouble) =>
            {
                if (eDouble.RowIndex < 0 || eDouble.ColumnIndex < 0) return; string col = dgv.Columns[eDouble.ColumnIndex].Name; var r = dgv.Rows[eDouble.RowIndex]; string sn = r.Cells[sipCol].Value?.ToString() ?? "";
                if (col == sipCol && !string.IsNullOrEmpty(sn) && sn != "-") { Clipboard.SetText(sn); MessageBox.Show($"Kopyalandı: {sn}"); return; }
                if (!string.IsNullOrEmpty(sn) && sn != "-")
                {
                    SiparisiAmeliyatMasasinaYatir(sn, 0, 0, dgv);
                }
            };

            dgv.DataBindingComplete += (s, eBind) =>
            {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    string ded = dgv.Columns.Contains("Dedektif (Kargo Durumu)") ? row.Cells["Dedektif (Kargo Durumu)"].Value?.ToString() ?? "" : "";
                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); row.DefaultCellStyle.SelectionForeColor = Color.White;

                    if (ded.StartsWith("✔ MÜTABIK")) { row.DefaultCellStyle.BackColor = Color.FromArgb(46, 117, 89); row.DefaultCellStyle.ForeColor = Color.White; }
                    else if (ded.StartsWith("HATA:")) { row.DefaultCellStyle.BackColor = Color.Crimson; row.DefaultCellStyle.ForeColor = Color.White; }
                    else { row.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); row.DefaultCellStyle.ForeColor = Color.White; }
                }
            };

            ContextMenuStrip cms = new ContextMenuStrip(); dgv.ContextMenuStrip = cms;
            dgv.MouseDown += (s, eMouse) =>
            {
                if (eMouse.Button == MouseButtons.Right)
                {
                    var hit = dgv.HitTest(eMouse.X, eMouse.Y); if (hit.Type == DataGridViewHitTestType.Cell && hit.RowIndex >= 0)
                    {
                        dgv.ClearSelection(); dgv.Rows[hit.RowIndex].Cells[hit.ColumnIndex].Selected = true; var r = dgv.Rows[hit.RowIndex]; string sn = r.Cells[sipCol].Value.ToString(); cms.Items.Clear();

                        if (!string.IsNullOrEmpty(sn) && sn != "-")
                        {
                            ToolStripMenuItem btnAmeliyatTY = new ToolStripMenuItem("✏️ İncele ve Manuel Düzelt (Ameliyat Masası)");
                            btnAmeliyatTY.Click += (sx, exAmeliyatTY) => { SiparisiAmeliyatMasasinaYatir(sn, 0, 0, dgv); };
                            cms.Items.Add(btnAmeliyatTY);

                            cms.Items.Add(new ToolStripSeparator());

                            string rawVal = r.Cells[hakCol].Value?.ToString().ToUpper() ?? "0";
                            if (rawVal.Contains("/")) rawVal = rawVal.Split('/')[0];
                            if (rawVal.Contains("TL")) rawVal = rawVal.Replace("TL", "");
                            double excelKargo = Math.Abs(ExcelHucredenSayiya(rawVal.Trim()));

                            ToolStripMenuItem btnKargoEsitle = new ToolStripMenuItem($"⚡ Kargo Tutarını Sistemle Eşitle ({excelKargo:N2} TL)");
                            btnKargoEsitle.Click += (sx, exCezaTY) => {
                                if (MessageBox.Show($"[{sn}] numaralı siparişin sistem kargo ücretini, yüklenen dosyadaki tutar olan {excelKargo:N2} TL ile güncellemek ve eşitlemek istiyor musunuz?", "Kargo Fiyatı Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    try
                                    {
                                        using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                                        {
                                            b.Open();
                                            using (var cmd = new SqliteCommand("UPDATE Siparisler SET KargoUcreti=@kargo WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%N11%'", b))
                                            {
                                                cmd.Parameters.AddWithValue("@kargo", excelKargo);
                                                cmd.Parameters.AddWithValue("@sn", sn);
                                                cmd.ExecuteNonQuery();
                                            }
                                        }

                                        DataRowView drv = r.DataBoundItem as DataRowView;
                                        if (drv != null)
                                        {
                                            if (drv.Row.Table.Columns.Contains("Sistem Kargo Ücreti (TL)")) drv.Row["Sistem Kargo Ücreti (TL)"] = excelKargo;
                                            if (drv.Row.Table.Columns.Contains("Dedektif (Kargo Durumu)")) drv.Row["Dedektif (Kargo Durumu)"] = "✔ MÜTABIK";
                                        }
                                        else
                                        {
                                            if (dgv.Columns.Contains("Sistem Kargo Ücreti (TL)")) r.Cells["Sistem Kargo Ücreti (TL)"].Value = excelKargo;
                                            if (dgv.Columns.Contains("Dedektif (Kargo Durumu)")) r.Cells["Dedektif (Kargo Durumu)"].Value = "✔ MÜTABIK";
                                        }

                                        dgv.Refresh();
                                        MessageBox.Show("Sistem kargo fiyatı başarıyla eşitlendi!", "Tamam", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    }
                                    catch (Exception exTY) { MessageBox.Show(exTY.Message); }
                                }
                            }; cms.Items.Add(btnKargoEsitle);
                        }
                        cms.Show(Cursor.Position);
                    }
                }
            };

            Panel pBtn = new Panel { Dock = DockStyle.Fill, Height = 50 };
            Button btnC = new Button { Text = "Kargo İncelemesini Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(194, 24, 91), ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnC.FlatAppearance.BorderSize = 0; pBtn.SizeChanged += (sB, eB) => { btnC.Location = new Point((pBtn.Width - btnC.Width) / 2, (pBtn.Height - btnC.Height) / 2); };
            btnC.Click += (sB, eB) => { f.Close(); };
            pBtn.Controls.Add(btnC);

            pnlGridContainer.Controls.Add(dgv); tlpPopupLayout.Controls.Add(pnlGridContainer, 0, 1); tlpPopupLayout.Controls.Add(pBtn, 0, 2);
            dgv.BringToFront(); dgv.DataSource = dtSource; f.ShowDialog();
        }
    }
}