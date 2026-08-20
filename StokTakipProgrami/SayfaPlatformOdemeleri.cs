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



namespace StokTakipProgrami

{

    // 🚀 MÜHÜR: MODEL SINIFLARI DERLEYİCİNİN GÖREBİLECEĞİ EN GÜVENLİ ALANDADIR

    public class DbSiparis { public string Durum; public double Satis, Komisyon, Kargo, Hizmet, Pos, Stopaj, Ceza; }

    public class HbSiparisOzet { public double ExcelSatis, ExcelKomisyon, ExcelKargo, ExcelHizmet; }



    public partial class SayfaPlatformOdemeleri : UserControl

    {

        private ComboBox dynCmbPlatform, dynCmbDosyaPlatformu;

        private DateTimePicker dynDtTarih;

        private TextBox dynTxtNetGelenPara, dynTxtSiparisListesi;

        private Button dynBtnOdemeyiMuhurle, dynBtnKaydiSil, dynBtnExcel, dynBtnSiparisSorgula, dynBtnTemizle;

        private DataGridView dynDgvTahsilatGecmisi, dgvPlatformBilanço;

        private Label lblCardTahsilEdilenAdet, lblCardTahsilEdilenPara, lblCardBekleyenAdet, lblCardBekleyenPara, lblCardToplamAdet, lblCardToplamPara;



        public SayfaPlatformOdemeleri() { InitializeComponent(); this.Load += SayfaPlatformOdemeleri_Load; }



        private void SayfaPlatformOdemeleri_Load(object sender, EventArgs e) { TabloyuGuvenceyeAl(); ArayuzInsaEt(); PlatformlariYukle(); TahsilatGecmisiniListele(); CanliBilançoListele(); }



        private void TabloyuGuvenceyeAl()

        {

            try

            {

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))

                {

                    conn.Open();

                    string sql = @"CREATE TABLE IF NOT EXISTS PlatformTahsilatlari (Id INTEGER PRIMARY KEY AUTOINCREMENT, Platform TEXT, Tarih TEXT, Tutar REAL, SiparisAdedi INTEGER, SiparisNumaralari TEXT);";

                    using (var cmd = new SqliteCommand(sql, conn)) { cmd.ExecuteNonQuery(); }

                }

            }

            catch { }

        }



        private void ArayuzInsaEt()

        {

            this.Controls.Clear(); this.Dock = DockStyle.Fill; this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(26, 26, 39) };

            Label lblT = new Label { Text = "💸 PLATFORM TOPLU HAKEDİŞ VE MUTABAKAT DENETİM TERMİNALİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };

            pnlHeader.Controls.Add(lblT); this.Controls.Add(pnlHeader);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };

            mainLayout.ColumnStyles.Add(new Panel { Width = 380 }.Dock == DockStyle.Left ? new ColumnStyle(SizeType.Absolute, 380F) : new ColumnStyle(SizeType.Absolute, 380F)); mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); this.Controls.Add(mainLayout);

            Panel pnlCard = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15), AutoScroll = true };

            dynCmbPlatform = new ComboBox { Location = new Point(15, 37), Width = 340, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

            pnlCard.Controls.Add(new Label { Text = "1. Ödemeyi Yapan Platform:", Location = new Point(15, 15), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray }); pnlCard.Controls.Add(dynCmbPlatform);

            dynDtTarih = new DateTimePicker { Location = new Point(15, 102), Width = 340, Font = new Font("Segoe UI", 10.5F), Format = DateTimePickerFormat.Short };

            pnlCard.Controls.Add(new Label { Text = "2. Ödemenin Geldiği Tarih:", Location = new Point(15, 80), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray }); pnlCard.Controls.Add(dynDtTarih);

            pnlCard.Controls.Add(FormElemaniYarat("3. Hesaba Yatan Toplam Tutar (TL):", out dynTxtNetGelenPara, new Point(15, 145), 340));

            dynTxtSiparisListesi = new TextBox { Location = new Point(15, 250), Width = 340, Height = 95, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 10F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };

            pnlCard.Controls.Add(new Label { Text = "4. Bu Ödemeye Ait Sipariş Numaraları:", Location = new Point(15, 220), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.DarkGray }); pnlCard.Controls.Add(dynTxtSiparisListesi);

            dynBtnTemizle = new Button { Text = "🧹 KUTULARI TEMİZLE", Location = new Point(15, 355), Width = 340, Height = 25, BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.LightGray, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) };

            dynBtnTemizle.FlatAppearance.BorderSize = 0; dynBtnTemizle.Click += (s, e) => { dynTxtSiparisListesi.Clear(); dynTxtNetGelenPara.Clear(); }; pnlCard.Controls.Add(dynBtnTemizle);

            dynBtnSiparisSorgula = new Button { Text = "🔍 MANUEL SORGULA", Location = new Point(15, 390), Width = 340, Height = 30, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) };

            dynBtnSiparisSorgula.FlatAppearance.BorderSize = 0; dynBtnSiparisSorgula.Click += DynBtnSiparisSorgula_Click; pnlCard.Controls.Add(dynBtnSiparisSorgula);

            dynBtnOdemeyiMuhurle = new Button { Text = "✅ TAHSİLAT HAKEDİŞİNİ MÜHÜRLE", Location = new Point(15, 430), Width = 340, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };

            dynBtnOdemeyiMuhurle.FlatAppearance.BorderSize = 0; dynBtnOdemeyiMuhurle.Click += DynBtnOdemeyiMuhurle_Click; pnlCard.Controls.Add(dynBtnOdemeyiMuhurle);

            Panel pnlAyrac = new Panel { Location = new Point(15, 490), Width = 340, Height = 1, BackColor = Color.FromArgb(50, 50, 55) }; pnlCard.Controls.Add(pnlAyrac);

            pnlCard.Controls.Add(new Label { Text = "5. DOSYADAN AKILLI MUTABAKAT (BETA):", Location = new Point(15, 505), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) });

            dynCmbDosyaPlatformu = new ComboBox { Location = new Point(15, 530), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat }; pnlCard.Controls.Add(dynCmbDosyaPlatformu);

            Button dynBtnMutabakatYukle = new Button { Text = "📂 DOSYA SEÇ VE EŞLEŞTİR", Location = new Point(175, 529), Width = 180, Height = 30, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) };

            dynBtnMutabakatYukle.FlatAppearance.BorderSize = 0; dynBtnMutabakatYukle.Click += DynBtnDosyaYukle_Click; pnlCard.Controls.Add(dynBtnMutabakatYukle);

            mainLayout.Controls.Add(pnlCard, 0, 0);

            TableLayoutPanel pnlRightLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };

            pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 45F)); pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 35F)); pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 95F)); mainLayout.Controls.Add(pnlRightLayout, 1, 0);

            Panel pnlUpperWrapper = new Panel { Dock = DockStyle.Fill }; pnlUpperWrapper.Controls.Add(new Label { Text = "📋 SİSTEME İŞLENEN TOPLU TAHSİLAT MAKBUZLARI", Dock = DockStyle.Top, Height = 22, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray });

            dynDgvTahsilatGecmisi = GridYarat(); dynDgvTahsilatGecmisi.CellDoubleClick += DynDgvTahsilatGecmisi_CellDoubleClick; pnlUpperWrapper.Controls.Add(dynDgvTahsilatGecmisi); dynDgvTahsilatGecmisi.BringToFront();

            Panel pnlRightBottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(0, 5, 0, 0) };

            dynBtnKaydiSil = new Button { Text = "🗑️ SEÇİLİ HAKEDİŞİ İPTAL ET", Dock = DockStyle.Left, Width = 380, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

            dynBtnKaydiSil.FlatAppearance.BorderSize = 0; dynBtnKaydiSil.Click += DynBtnKaydiSil_Click; pnlRightBottom.Controls.Add(dynBtnKaydiSil);

            dynBtnExcel = new Button { Text = "📊 Excel Rapor", Dock = DockStyle.Right, Width = 130, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

            dynBtnExcel.FlatAppearance.BorderSize = 0; dynBtnExcel.Click += (s, e) => ExcelMotoru.DataGridViewTablosunuDisaAktar(dynDgvTahsilatGecmisi, "Hakedis_Gecmisi");

            pnlRightBottom.Controls.Add(dynBtnExcel); pnlUpperWrapper.Controls.Add(pnlRightBottom); pnlRightLayout.Controls.Add(pnlUpperWrapper, 0, 0);

            GroupBox grpBilanço = new GroupBox { Dock = DockStyle.Fill, Text = "📊 PAZARYERLERİ CARİ BİLANÇO MATRİSİ", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange, Padding = new Padding(10, 22, 10, 10) };

            dgvPlatformBilanço = GridYarat(); grpBilanço.Controls.Add(dgvPlatformBilanço); pnlRightLayout.Controls.Add(grpBilanço, 0, 1);

            TableLayoutPanel tlpVIPCards = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Padding = new Padding(0, 5, 0, 5) };

            tlpVIPCards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F)); tlpVIPCards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F)); tlpVIPCards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));

            CustomVIPCardInsaEt(tlpVIPCards, "💰 TOPLAM TAHSİL EDİLEN", Color.FromArgb(34, 197, 94), 0, out lblCardTahsilEdilenAdet, out lblCardTahsilEdilenPara);

            CustomVIPCardInsaEt(tlpVIPCards, "⏳ TOPLAM KALAN NET ALACAK", Color.Orange, 1, out lblCardBekleyenAdet, out lblCardBekleyenPara);

            CustomVIPCardInsaEt(tlpVIPCards, "💎 GENEL MAĞAZA HACMİ (CİRO)", Color.FromArgb(56, 189, 248), 2, out lblCardToplamAdet, out lblCardToplamPara);

            pnlRightLayout.Controls.Add(tlpVIPCards, 0, 2); mainLayout.BringToFront(); pnlHeader.SendToBack();

        }



        private bool PromptForInput(string t, string c, out string v)

        {

            Form p = new Form() { Width = 350, Height = 180, FormBorderStyle = FormBorderStyle.FixedDialog, Text = c, StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };

            TextBox ib = new TextBox() { Left = 20, Top = 50, Width = 290, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, Font = new Font("Segoe UI", 11F) };

            Button cf = new Button() { Text = "Onayla", Left = 210, Width = 100, Top = 90, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

            string r = ""; bool ok = false; cf.Click += (s, e) => { r = ib.Text.Trim(); ok = true; p.Close(); };

            p.Controls.Add(new Label() { Left = 20, Top = 20, Text = t, AutoSize = true, ForeColor = Color.White }); p.Controls.Add(ib); p.Controls.Add(cf); p.AcceptButton = cf; p.ShowDialog(); v = r; return ok;

        }



        private bool PromptForHurda(string dSn, out string oSn, out double oPo, out double oMz)

        {

            string tSn = dSn; double tP = 0, tZ = 0;

            Form p = new Form() { Width = 380, Height = 280, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "🛠️ Hurda Kapat", StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };

            TextBox txt1 = new TextBox() { Left = 20, Top = 35, Width = 320, Text = dSn, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };

            TextBox txt2 = new TextBox() { Left = 20, Top = 95, Width = 320, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };

            TextBox txt3 = new TextBox() { Left = 20, Top = 155, Width = 320, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };

            Button cf = new Button() { Text = "✔ Onayla", Left = 20, Width = 320, Top = 200, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White };

            bool isOk = false; cf.Click += (s, e) => { tSn = txt1.Text.Trim(); double.TryParse(txt2.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out tP); double.TryParse(txt3.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out tZ); if (!string.IsNullOrEmpty(tSn)) { isOk = true; p.Close(); } };

            p.Controls.Add(txt1); p.Controls.Add(txt2); p.Controls.Add(txt3); p.Controls.Add(cf); p.ShowDialog(); oSn = tSn; oPo = tP; oMz = tZ; return isOk;

        }



        private void DynDgvTahsilatGecmisi_CellDoubleClick(object sender, DataGridViewCellEventArgs e)

        {

            if (e.RowIndex < 0) return; int id = Convert.ToInt32(dynDgvTahsilatGecmisi.Rows[e.RowIndex].Cells["Id"].Value); string nolar = "";

            try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT SiparisNumaralari FROM PlatformTahsilatlari WHERE Id=@id", conn)) { cmd.Parameters.AddWithValue("@id", id); nolar = cmd.ExecuteScalar()?.ToString() ?? ""; } } } catch { }

            Form f = new Form { Text = $"Detay", Size = new Size(380, 450), StartPosition = FormStartPosition.CenterParent };

            TextBox t = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, Text = nolar.Replace(", ", Environment.NewLine), ReadOnly = true }; f.Controls.Add(t); f.ShowDialog();

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

                        foreach (var s in sips) { using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='KARGOYA VERİLDİ' WHERE SiparisNo=@sn", b)) { cmd.Parameters.AddWithValue("@sn", s); cmd.ExecuteNonQuery(); } }

                        using (var cmd = new SqliteCommand("DELETE FROM PlatformTahsilatlari WHERE Id=@id", b)) { cmd.Parameters.AddWithValue("@id", id); cmd.ExecuteNonQuery(); }

                    }

                    TahsilatGecmisiniListele(); CanliBilançoListele();

                }

                catch (Exception exSil) { MessageBox.Show(exSil.Message); }

            }

        }



        private void DynBtnOdemeyiMuhurle_Click(object sender, EventArgs e)

        {

            if (dynCmbPlatform.SelectedIndex <= 0 || string.IsNullOrEmpty(dynTxtNetGelenPara.Text)) return;

            try

            {

                double tutar = ExcelHucredenSayiya(dynTxtNetGelenPara.Text); string plat = dynCmbPlatform.Text; string tar = dynDtTarih.Value.ToString("dd.MM.yyyy");

                var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList();

                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))

                {

                    b.Open();

                    using (var cmd = new SqliteCommand("INSERT INTO PlatformTahsilatlari (Platform, Tarih, Tutar, SiparisAdedi, SiparisNumaralari) VALUES (@p, @t, @v, @a, @n)", b))

                    {

                        cmd.Parameters.AddWithValue("@p", plat); cmd.Parameters.AddWithValue("@t", tar); cmd.Parameters.AddWithValue("@v", tutar); cmd.Parameters.AddWithValue("@a", sips.Count); cmd.Parameters.AddWithValue("@n", string.Join(", ", sips)); cmd.ExecuteNonQuery();

                    }

                    foreach (var s in sips) { using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='TAHSİL EDİLDİ' WHERE SiparisNo=@sn", b)) { cmd.Parameters.AddWithValue("@sn", s); cmd.ExecuteNonQuery(); } }

                }

                dynTxtNetGelenPara.Clear(); dynTxtSiparisListesi.Clear(); TahsilatGecmisiniListele(); CanliBilançoListele(); MessageBox.Show("Mühürlendi!");

            }

            catch (Exception exMuhur) { MessageBox.Show(exMuhur.Message); }

        }



        private double ExcelHucredenSayiya(object hucre)

        {

            if (hucre == null || hucre == DBNull.Value) return 0; if (hucre is double d) return d;

            if (hucre is decimal dc) return Convert.ToDouble(dc); if (hucre is int i) return i;

            if (hucre is long lg) return lg; if (hucre is float fl) return fl;

            string s = hucre.ToString().Replace("₺", "").Replace("TL", "").Replace(" ", "").Replace("\xA0", "").Trim(); if (string.IsNullOrEmpty(s)) return 0;

            int sonNokta = s.LastIndexOf('.'); int sonVirgul = s.LastIndexOf(',');

            if (sonNokta > sonVirgul) { if (s.Length - 1 - sonNokta <= 2) { s = s.Substring(0, sonNokta).Replace(".", "").Replace(",", "") + "." + s.Substring(sonNokta + 1); } else { s = s.Replace(".", "").Replace(",", ""); } }

            else if (sonVirgul > sonNokta) { if (s.Length - 1 - sonVirgul <= 2) { s = s.Substring(0, sonVirgul).Replace(".", "").Replace(",", "") + "." + s.Substring(sonVirgul + 1); } else { s = s.Replace(".", "").Replace(",", ""); } }

            else { s = s.Replace(".", "").Replace(",", ""); }

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double resFinal)) return resFinal; return 0;

        }



        private DataTable ExcelVeCsvOku(string dosyaYolu)

        {

            DataTable dt = new DataTable(); string ext = Path.GetExtension(dosyaYolu).ToLower();

            if (ext == ".xls" || ext == ".xlsx")

            {

                string connStr = ext == ".xls" ? $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dosyaYolu};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\";" : $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dosyaYolu};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";

                using (OleDbConnection conn = new OleDbConnection(connStr))

                {

                    conn.Open(); DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null); string sn = schema.Rows[0]["TABLE_NAME"].ToString();

                    if (sn.StartsWith("'") && sn.EndsWith("'")) sn = sn.Substring(1, sn.Length - 2);

                    using (OleDbDataAdapter da = new OleDbDataAdapter($"SELECT * FROM [{sn}]", conn)) { da.Fill(dt); }

                }

            }

            else

            {

                string[] lines = File.ReadAllLines(dosyaYolu, System.Text.Encoding.Default);

                if (lines.Length > 0)

                {

                    char ayrac = lines[0].Contains(";") ? ';' : ','; string[] headers = SatiriParcala(lines[0], ayrac); foreach (string h in headers) dt.Columns.Add(h.Trim());

                    for (int i = 1; i < lines.Length; i++) { if (string.IsNullOrWhiteSpace(lines[i])) continue; string[] cells = SatiriParcala(lines[i], ayrac); DataRow row = dt.NewRow(); for (int j = 0; j < cells.Length && j < dt.Columns.Count; j++) row[j] = cells[j].Trim(); dt.Rows.Add(row); }

                }

            }

            return dt;

        }



        // 🚀 MÜHÜR 1: CS0103 HATASINI KÖKTEN ÇÖZEN VE YERİ TAMAMEN SABİTLENEN HÜCRE PARÇALAYICI

        private string[] SatiriParcala(string satir, char ayrac)

        {

            var res = new List<string>(); bool ti = false; string hc = "";

            for (int i = 0; i < satir.Length; i++) { char c = satir[i]; if (c == '"') ti = !ti; else if (c == ayrac && !ti) { res.Add(hc.Trim('"')); hc = ""; } else hc += c; }

            res.Add(hc.Trim('"')); return res.ToArray();

        }



        private void DynBtnDosyaYukle_Click(object sender, EventArgs e)

        {

            if (dynCmbDosyaPlatformu.SelectedIndex <= 0) return;

            string dp = dynCmbDosyaPlatformu.Text.ToUpper(); bool isTY = dp == "TRENDYOL", isHB = dp == "HEPSİBURADA";

            using (OpenFileDialog ofd = new OpenFileDialog())

            {

                ofd.Filter = "Excel Dosyaları|*.xlsx;*.xls;*.csv";

                if (ofd.ShowDialog() == DialogResult.OK)

                {

                    try

                    {

                        DataTable dt = ExcelVeCsvOku(ofd.FileName); int idxIslem = -1, idxSip = -1, idxHak = -1, idxKom = -1, idxAcik = -1;

                        for (int i = 0; i < dt.Columns.Count; i++)

                        {

                            string b = dt.Columns[i].ColumnName.ToUpper().Trim();

                            if (b == "İŞLEM TİPİ" || b == "ISLEM TIPI" || b == "KAYIT TİPİ") idxIslem = i;

                            else if (b == "SİPARİŞ NO" || b == "SIPARIS NO" || b == "SİPARİŞ NUMARASI") idxSip = i;

                            else if (b == "SATICI HAKEDİŞ" || b == "NET TUTAR" || b == "TUTAR") idxHak = i;

                            else if (b == "TY HAKEDİŞ" || b == "KOMİSYON TUTARI") idxKom = i;

                            else if (b.Contains("AÇIKLAMA") || b.Contains("ÜRÜN ADI")) idxAcik = i;

                        }

                        if (idxSip == -1 || idxHak == -1) { MessageBox.Show("Sütunlar uyuşmuyor!"); return; }

                        var sips = new HashSet<string>(); double total = 0;

                        for (int i = 0; i < dt.Rows.Count; i++)

                        {

                            int doluHucre = 0; for (int c = 0; c < dt.Columns.Count; c++) { if (!string.IsNullOrWhiteSpace(dt.Rows[i][c].ToString())) doluHucre++; }

                            string sn = dt.Rows[i][idxSip].ToString().Trim(); if (isHB && (string.IsNullOrEmpty(sn) || sn == "-" || doluHucre < 5)) continue;

                            if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                            if (sn != "-") sips.Add(sn); total += ExcelHucredenSayiya(dt.Rows[i][idxHak]);

                        }

                        var dbHavuz = new Dictionary<string, DbSiparis>();

                        using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))

                        {

                            conn.Open();

                            foreach (var s in sips)

                            {

                                using (var cmd = new SqliteCommand("SELECT Durum, SatisFiyati, KomisyonTL, KargoUcreti, HizmetBedeli, IFNULL(CezaTutari,0) FROM Siparisler WHERE SiparisNo=@sn", conn))

                                {

                                    cmd.Parameters.AddWithValue("@sn", s); using (var r = cmd.ExecuteReader()) { if (r.Read()) dbHavuz[s] = new DbSiparis { Durum = r[0].ToString(), Satis = Convert.ToDouble(r[1]), Komisyon = Convert.ToDouble(r[2]), Kargo = Convert.ToDouble(r[3]), Hizmet = Convert.ToDouble(r[4]), Ceza = Convert.ToDouble(r[5]) }; }

                                }

                            }

                        }

                        DataTable dtHam = new DataTable(); dtHam.Columns.Add("Sipariş No"); dtHam.Columns.Add("İşlem Tipi"); dtHam.Columns.Add("Açıklama / Ürün"); dtHam.Columns.Add("Tutar (TL)", typeof(double));

                        if (isTY) dtHam.Columns.Add("Komisyon (TL)", typeof(double));

                        dtHam.Columns.Add("Dedektif (Durum)"); dtHam.Columns.Add("Sistem Durumu");

                        int hatalar = 0;



                        #region 🟠 TRENDYOL DOSYA MOTORU

                        if (isTY)

                        {

                            foreach (DataRow row in dt.Rows)

                            {

                                string sn = row[idxSip].ToString().Trim(); if (sn.ToUpper().Contains("TOPLAM") || sn.ToUpper().Contains("TOTAL")) continue;

                                if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                                if (string.IsNullOrEmpty(sn)) sn = "-";

                                string isl = idxIslem != -1 ? row[idxIslem].ToString() : "SATIŞ"; double tut = ExcelHucredenSayiya(row[idxHak]); double kom = idxKom != -1 ? ExcelHucredenSayiya(row[idxKom]) : 0;

                                string sys = "-", ded = "✔ KUSURSUZ"; if (dbHavuz.TryGetValue(sn, out var db)) { sys = db.Durum; if (!isl.ToUpper().Contains("KUPON") && !isl.ToUpper().Contains("İNDİRİM") && !isl.ToUpper().Contains("SATIŞ")) { if (Math.Abs(Math.Abs(tut) - db.Ceza) > 0.1) ded = "HATA: CEZA FARKI"; } } else if (sn != "-") { sys = "SİSTEMDE YOK"; ded = "HATA: KAYITSIZ"; } else { sys = "-"; ded = "GENEL İŞLEM"; }

                                if (ded.StartsWith("HATA:")) hatalar++; dtHam.Rows.Add(sn, isl, idxAcik != -1 ? row[idxAcik].ToString() : "", tut, kom, ded, sys);

                            }

                        }

                        #endregion

                        #region 📋 HEPSİBURADA DOSYA MOTORU

                        else if (isHB)

                        {

                            var hbO = new Dictionary<string, HbSiparisOzet>();

                            foreach (DataRow row in dt.Rows)

                            {

                                int dh = 0; for (int c = 0; c < dt.Columns.Count; c++) { if (!string.IsNullOrWhiteSpace(row[c].ToString())) dh++; }

                                string sn = row[idxSip].ToString().Trim(); if (string.IsNullOrEmpty(sn) || sn == "-" || dh < 5) continue;

                                if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                                if (!hbO.ContainsKey(sn)) hbO[sn] = new HbSiparisOzet(); string isl = idxIslem != -1 ? row[idxIslem].ToString().ToUpper() : ""; double tut = ExcelHucredenSayiya(row[idxHak]);

                                if (isl.Contains("SİPARİŞ")) hbO[sn].ExcelSatis += tut; else if (isl.Contains("KOMİSYON")) hbO[sn].ExcelKomisyon += Math.Abs(tut); else if (isl.Contains("KARGO")) hbO[sn].ExcelKargo += Math.Abs(tut); else hbO[sn].ExcelHizmet += Math.Abs(tut);

                            }

                            foreach (DataRow row in dt.Rows)

                            {

                                int dh = 0; for (int c = 0; c < dt.Columns.Count; c++) { if (!string.IsNullOrWhiteSpace(row[c].ToString())) dh++; }

                                string sn = row[idxSip].ToString().Trim(); if (string.IsNullOrEmpty(sn) || sn == "-" || dh < 5) continue;

                                if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                                string isl = idxIslem != -1 ? row[idxIslem].ToString() : "SATIŞ"; string ded = "✔ KUSURSUZ";

                                if (dbHavuz.TryGetValue(sn, out var db))

                                {

                                    if (hbO.TryGetValue(sn, out var oz))

                                    {

                                        if (Math.Abs(oz.ExcelSatis - db.Satis) > 0.1 || Math.Abs(oz.ExcelKargo - db.Kargo) > 0.1 || Math.Abs(oz.ExcelHizmet - db.Hizmet) > 0.1) ded = "HATA: FARK VAR";

                                    }

                                }

                                else ded = "HATA: KAYITSIZ";

                                if (ded.StartsWith("HATA:")) hatalar++; dtHam.Rows.Add(sn, isl, idxAcik != -1 ? row[idxAcik].ToString() : "", ExcelHucredenSayiya(row[idxHak]), ded, db != null ? db.Durum : "SİSTEMDE YOK");

                            }

                        }

                        #endregion

                        if (dynCmbPlatform.Items.Contains(dp)) dynCmbPlatform.SelectedItem = dp;

                        dynTxtSiparisListesi.Text = string.Join(Environment.NewLine, sips); dynTxtNetGelenPara.Text = total.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));

                        if (hatalar > 0) MessageBox.Show($"🚨 DİKKAT! {hatalar} SATIRDA HATA VAR PATRON!", "Radar", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        MutabakatUyariEkraniniAc(dtHam, dp);

                    }

                    catch (Exception exDosya) { MessageBox.Show(exDosya.Message); }

                }

            }

        }



        private void MutabakatUyariEkraniniAc(DataTable dtSource, string dp)

        {

            Form f = new Form { Text = $"{dp} Ekstre Önizleme", Size = new Size(1350, 650), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };

            Label lbl = new Label { Text = "Sipariş No hücresine çift tıklayarak panoya kopyalayabilirsiniz. Diğer hücrelere Çift Tıklayarak profesyonel döküm matrisini açabilirsiniz.", Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Orange, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };

            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), SelectionMode = DataGridViewSelectionMode.CellSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EnableHeadersVisualStyles = false, RowHeadersVisible = false };

            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DefaultCellStyle.SelectionForeColor = Color.White; dgv.ColumnHeadersHeight = 40; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;



            dgv.CellDoubleClick += (s, eDouble) => {

                if (eDouble.RowIndex < 0 || eDouble.ColumnIndex < 0) return;

                string col = dgv.Columns[eDouble.ColumnIndex].Name; var r = dgv.Rows[eDouble.RowIndex]; string sn = r.Cells["Sipariş No"].Value?.ToString() ?? "";

                if (col == "Sipariş No" && !string.IsNullOrEmpty(sn) && sn != "-") { Clipboard.SetText(sn); MessageBox.Show($"Kopyalandı: {sn}"); return; }

                if (!string.IsNullOrEmpty(sn) && sn != "-")

                {

                    double tH = 0; foreach (DataGridViewRow rX in dgv.Rows) { if (rX.Cells["Sipariş No"].Value.ToString() == sn) { string tp = rX.Cells["İşlem Tipi"].Value.ToString().ToUpper(); if (tp.Contains("SATIŞ") || tp.Contains("SİPARİŞ") || tp.Contains("SIPARIS") || tp.Contains("KUPON")) { tH += Convert.ToDouble(rX.Cells["Tutar (TL)"].Value); } } }

                    SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);

                }

            };



            dgv.DataBindingComplete += (s, eBind) => {

                foreach (DataGridViewRow row in dgv.Rows)

                {

                    string type = row.Cells["İşlem Tipi"].Value?.ToString().ToUpper() ?? ""; string ded = row.Cells["Dedektif (Durum)"].Value?.ToString() ?? "";

                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); row.DefaultCellStyle.SelectionForeColor = Color.White;

                    if (ded == "MÜDAHALE EDİLDİ" || ded == "HURDA İŞLENDİ" || ded == "✔ KUSURSUZ") { row.DefaultCellStyle.BackColor = Color.FromArgb(46, 117, 89); row.DefaultCellStyle.ForeColor = Color.White; }

                    else if (ded.StartsWith("HATA:")) { row.DefaultCellStyle.BackColor = Color.Crimson; row.DefaultCellStyle.ForeColor = Color.White; }

                    else { row.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); if (type.Contains("SATIŞ") || type.Contains("SIPARIS") || type.Contains("SİPARİŞ") || type.Contains("KUPON")) row.DefaultCellStyle.ForeColor = Color.DarkGray; else row.DefaultCellStyle.ForeColor = Color.White; }

                }

            };



            ContextMenuStrip cms = new ContextMenuStrip(); dgv.ContextMenuStrip = cms;

            dgv.MouseDown += (s, eMouse) => {

                if (eMouse.Button == MouseButtons.Right)

                {

                    var hit = dgv.HitTest(eMouse.X, eMouse.Y); if (hit.Type == DataGridViewHitTestType.Cell && hit.RowIndex >= 0)

                    {

                        dgv.ClearSelection(); dgv.Rows[hit.RowIndex].Cells[hit.ColumnIndex].Selected = true; var r = dgv.Rows[hit.RowIndex]; string sn = r.Cells["Sipariş No"].Value.ToString(); string islemTipi = r.Cells["İşlem Tipi"].Value.ToString().ToUpper(); cms.Items.Clear();



                        // 🚀 MÜHÜR 2: CS1028 ENGELLEYİCİ - ÖNİŞLEMCİ YÖNERGELERİ KESKİNCE KOŞUL BLOKLARININ İÇİNE ALINDI

                        if (dp == "TRENDYOL")

                        {

                            #region [TRENDYOL SAĞ TIK İZOLASYONU]

                            ToolStripMenuItem btnAmeliyatTY = new ToolStripMenuItem("✏️ İncele ve Manuel Düzelt (Ameliyat Masası)");

                            btnAmeliyatTY.Click += (sx, exAmeliyatTY) => {

                                if (string.IsNullOrEmpty(sn) || sn == "-") return; double tH = 0; foreach (DataGridViewRow rX in dgv.Rows) { if (rX.Cells["Sipariş No"].Value.ToString() == sn) { string tp = rX.Cells["İşlem Tipi"].Value.ToString().ToUpper(); if (tp.Contains("SATIŞ") || tp.Contains("SATIS") || tp.Contains("KUPON") || tp.Contains("İNDİRİM")) { tH += Convert.ToDouble(rX.Cells["Tutar (TL)"].Value); } } }

                                SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);

                            }; cms.Items.Add(btnAmeliyatTY);

                            bool isSatisVeyaKuponTY = islemTipi.Contains("SATIŞ") || islemTipi.Contains("SATIS") || islemTipi.Contains("KUPON") || islemTipi.Contains("İNDİRİM");

                            if (!isSatisVeyaKuponTY)

                            {

                                cms.Items.Add(new ToolStripSeparator()); ToolStripMenuItem btnCezaTY = new ToolStripMenuItem("⚠️ Kesintiyi 'Ceza Tutarı' Olarak İşle");

                                btnCezaTY.Click += (sx, exCezaTY) => {

                                    string targetSn = sn; if (string.IsNullOrEmpty(targetSn) || targetSn == "-") { if (PromptForInput("Sipariş No girin:", "Eksik", out string inpSn)) targetSn = inpSn; else return; }

                                    double tut = Convert.ToDouble(r.Cells["Tutar (TL)"].Value);

                                    if (MessageBox.Show($"[{targetSn}] siparişe {Math.Abs(tut):N2} TL ceza eklenecektir?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)

                                    {

                                        try

                                        {

                                            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET CezaTutari=IFNULL(CezaTutari,0)+@c, CezaSebebi=CASE WHEN IFNULL(CezaSebebi,'')='' THEN @s ELSE CezaSebebi||' | '||@s END WHERE SiparisNo=@sn", b)) { cmd.Parameters.AddWithValue("@c", Math.Abs(tut)); cmd.Parameters.AddWithValue("@s", islemTipi); cmd.Parameters.AddWithValue("@sn", targetSn); cmd.ExecuteNonQuery(); } }

                                            foreach (DataGridViewRow gridRow in dgv.Rows) { if (gridRow.Cells["Sipariş No"].Value.ToString() == targetSn) { gridRow.Cells["Dedektif (Durum)"].Value = "MÜDAHALE EDİLDİ"; } }

                                            dgv.Refresh();

                                        }

                                        catch (Exception exTY) { MessageBox.Show(exTY.Message); }

                                    }

                                }; cms.Items.Add(btnCezaTY);

                                ToolStripMenuItem btnErkenTY = new ToolStripMenuItem("⚡ Kesintiyi 'Erken Ödeme Kesintisi' Olarak İşle");

                                btnErkenTY.Click += (sx, exErkenTY) => {

                                    double tut = Math.Abs(Convert.ToDouble(r.Cells["Tutar (TL)"].Value));

                                    if (MessageBox.Show($"{tut:N2} TL Erken Ödeme olarak düşülecek?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)

                                    {

                                        try

                                        {

                                            string yeniSip = "KESİNTİ-" + DateTime.Now.ToString("yyyyMMddHHmmss");

                                            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("INSERT INTO Siparisler (SiparisNo, PlatformAd, UrunAdi, Adet, SatisFiyati, CezaTutari, CezaSebebi, Tarih, Durum) VALUES (@sn, 'TRENDYOL', 'Erken Ödeme Kesintisi', 0, 0, @c, 'Erken Ödeme', @t, 'BEKLEYEN')", b)) { cmd.Parameters.AddWithValue("@sn", yeniSip); cmd.Parameters.AddWithValue("@c", tut); cmd.Parameters.AddWithValue("@t", dynDtTarih.Value.ToString("dd.MM.yyyy")); cmd.ExecuteNonQuery(); } }

                                            r.Cells["Sistem Durumu"].Value = "ERKEN ÖDEME İŞLENDİ"; r.Cells["Dedektif (Durum)"].Value = "MÜDAHALE EDİLDİ"; dgv.Refresh();

                                            if (!string.IsNullOrEmpty(dynTxtSiparisListesi.Text)) dynTxtSiparisListesi.Text += Environment.NewLine; dynTxtSiparisListesi.Text += yeniSip;

                                        }

                                        catch (Exception exTY2) { MessageBox.Show(exTY2.Message); }

                                    }

                                }; cms.Items.Add(btnErkenTY);

                            }

                            #endregion

                        }

                        else if (dp == "HEPSİBURADA")

                        {

                            #region [HEPSİBURADA SAĞ TIK İZOLASYONU]

                            ToolStripMenuItem btnAmeliyatHB = new ToolStripMenuItem("✏️ İncele ve Manuel Düzelt (Ameliyat Masası)");

                            btnAmeliyatHB.Click += (sx, exAmeliyatHB) => {

                                if (string.IsNullOrEmpty(sn) || sn == "-") return; double tH = 0; foreach (DataGridViewRow rX in dgv.Rows) { if (rX.Cells["Sipariş No"].Value.ToString() == sn) { string tp = rX.Cells["İşlem Tipi"].Value.ToString().ToUpper(); if (tp.Contains("SİPARİŞ") || tp.Contains("SIPARIS")) tH += Convert.ToDouble(rX.Cells["Tutar (TL)"].Value); } }

                                SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);

                            }; cms.Items.Add(btnAmeliyatHB);

                            if (islemTipi.Contains("HURDA") || r.Cells["Açıklama / Ürün"].Value.ToString().ToUpper().Contains("HURDA"))

                            {

                                cms.Items.Add(new ToolStripSeparator()); ToolStripMenuItem btnH = new ToolStripMenuItem("🛠️ İhtilaflıyı Bitir (Hurda Geliri İşle)");

                                btnH.Click += (sx, exHurdaHB) => {

                                    if (PromptForHurda(sn == "-" ? "" : sn, out string nS, out double pO, out double mZ))

                                    {

                                        try

                                        {

                                            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))

                                            {

                                                b.Open(); string bar = ""; int ad = 1; using (var cmd = new SqliteCommand("SELECT Barkod, Adet FROM Siparisler WHERE SiparisNo=@sn", b)) { cmd.Parameters.AddWithValue("@sn", nS); using (var read = cmd.ExecuteReader()) { if (read.Read()) { bar = read[0].ToString(); ad = Convert.ToInt32(read[1]); } } }

                                                using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='İADE', AlisFiyati=MAX(0, AlisFiyati-@po), CezaTutari=IFNULL(CezaTutari,0)+@mz, CezaSebebi='Hurda' WHERE SiparisNo=@sn", b)) { cmd.Parameters.AddWithValue("@po", pO); cmd.Parameters.AddWithValue("@mz", mZ); cmd.Parameters.AddWithValue("@sn", nS); cmd.ExecuteNonQuery(); }

                                                using (var cmd = new SqliteCommand("INSERT INTO Stoklar (Barkod, UrunAdi, Kategori, Marka, Adet, StokMiktari, AlisFiyati) SELECT Barkod||'-2EL', UrunAdi||' (2. EL)', Kategori, Marka, @a, @a, MAX(0, AlisFiyati-@po) FROM Stoklar WHERE Barkod=@ob", b)) { cmd.Parameters.AddWithValue("@a", ad); cmd.Parameters.AddWithValue("@po", pO); cmd.Parameters.AddWithValue("@ob", bar); cmd.ExecuteNonQuery(); }

                                            }

                                            foreach (DataGridViewRow rH in dgv.Rows) { if (rH.Cells["Sipariş No"].Value.ToString() == nS) { rH.Cells["Sistem Durumu"].Value = "İHTİLAFLI BİTTİ"; rH.Cells["Dedektif (Durum)"].Value = "HURDA İŞLENDİ"; } }

                                            dgv.Refresh();

                                        }

                                        catch (Exception exHB) { MessageBox.Show(exHB.Message); }

                                    }

                                }; cms.Items.Add(btnH);

                            }

                            bool isStandartSatisHB = islemTipi.Contains("SİPARİŞ") || islemTipi.Contains("SIPARIS") || islemTipi.Contains("KOMİSYON") || islemTipi.Contains("KARGO") || islemTipi.Contains("HİZMET") || islemTipi.Contains("İŞLEM");

                            if (!isStandartSatisHB)

                            {

                                cms.Items.Add(new ToolStripSeparator()); ToolStripMenuItem btnCezaHB = new ToolStripMenuItem("⚠️ Kesintiyi 'Ceza Tutarı' Olarak İşle");

                                btnCezaHB.Click += (sx, exCezaHB) => {

                                    string targetSn = sn; if (string.IsNullOrEmpty(targetSn) || targetSn == "-") { if (PromptForInput("Sipariş No girin:", "Eksik", out string inpSn)) targetSn = inpSn; else return; }

                                    double tut = Convert.ToDouble(r.Cells["Tutar (TL)"].Value);

                                    if (MessageBox.Show($"[{targetSn}] siparişe {Math.Abs(tut):N2} TL ceza eklenecektir?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)

                                    {

                                        try

                                        {

                                            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET CezaTutari=IFNULL(CezaTutari,0)+@c, CezaSebebi=CASE WHEN IFNULL(CezaSebebi,'')='' THEN @s ELSE CezaSebebi||' | '||@s END WHERE SiparisNo=@sn", b)) { cmd.Parameters.AddWithValue("@c", Math.Abs(tut)); cmd.Parameters.AddWithValue("@s", islemTipi); cmd.Parameters.AddWithValue("@sn", targetSn); cmd.ExecuteNonQuery(); } }

                                            foreach (DataGridViewRow gridRow in dgv.Rows) { if (gridRow.Cells["Sipariş No"].Value.ToString() == targetSn) { gridRow.Cells["Dedektif (Durum)"].Value = "MÜDAHALE EDİLDİ"; } }

                                            dgv.Refresh();

                                        }

                                        catch (Exception exHB2) { MessageBox.Show(exHB2.Message); }

                                    }

                                }; cms.Items.Add(btnCezaHB);

                                ToolStripMenuItem btnErkenHB = new ToolStripMenuItem("⚡ Kesintiyi 'Erken Ödeme Kesintisi' Olarak İşle");

                                btnErkenHB.Click += (sx, exErkenHB) => {

                                    double tut = Math.Abs(Convert.ToDouble(r.Cells["Tutar (TL)"].Value));

                                    if (MessageBox.Show($"{tut:N2} TL Erken Ödeme olarak düşülecek?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)

                                    {

                                        try

                                        {

                                            string yeniSip = "KESİNTİ-" + DateTime.Now.ToString("yyyyMMddHHmmss");

                                            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("INSERT INTO Siparisler (SiparisNo, PlatformAd, UrunAdi, Adet, SatisFiyati, CezaTutari, CezaSebebi, Tarih, Durum) VALUES (@sn, 'HEPSİBURADA', 'Erken Ödeme Kesintisi', 0, 0, @c, 'Erken Ödeme', @t, 'BEKLEYEN')", b)) { cmd.Parameters.AddWithValue("@sn", yeniSip); cmd.Parameters.AddWithValue("@c", tut); cmd.Parameters.AddWithValue("@t", dynDtTarih.Value.ToString("dd.MM.yyyy")); cmd.ExecuteNonQuery(); } }

                                            r.Cells["Sistem Durumu"].Value = "ERKEN ÖDEME İŞLENDİ"; r.Cells["Dedektif (Durum)"].Value = "MÜDAHALE EDİLDİ"; dgv.Refresh();

                                            if (!string.IsNullOrEmpty(dynTxtSiparisListesi.Text)) dynTxtSiparisListesi.Text += Environment.NewLine; dynTxtSiparisListesi.Text += yeniSip;

                                        }

                                        catch (Exception exHB3) { MessageBox.Show(exHB3.Message); }

                                    }

                                }; cms.Items.Add(btnErkenHB);

                            }

                            #endregion

                        }

                        cms.Show(Cursor.Position);

                    }

                }

            };

            Panel pBtn = new Panel { Dock = DockStyle.Bottom, Height = 50 };

            Button btnC = new Button { Text = "Mutabakatı Onayla ve Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };

            btnC.FlatAppearance.BorderSize = 0; pBtn.SizeChanged += (s, e) => { btnC.Location = new Point((pBtn.Width - btnC.Width) / 2, (pBtn.Height - btnC.Height) / 2); };

            btnC.Click += (s, e) => { f.DialogResult = DialogResult.OK; f.Close(); }; pBtn.Controls.Add(btnC);

            f.Controls.Add(dgv); f.Controls.Add(lbl); f.Controls.Add(pBtn); dgv.BringToFront(); dgv.DataSource = dtSource; f.ShowDialog();

        }



        private void SiparisiAmeliyatMasasinaYatir(string sipNo, double exH, double exK, DataGridView anaGrid)

        {

            long rId = 0; double sF = 0, km = 0, kr = 0, hz = 0, st = 0, ps = 0, cz = 0; string durum = "YOK", brkd = "-", uAdi = "Bilinmeyen Ürün", kat = "-", mrk = "-", tar = "-"; int adet = 1; double alisFiyati = 0; string imei = "-";

            try

            {

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))

                {

                    conn.Open(); string q = @"SELECT s.rowid, s.SatisFiyati, s.KomisyonTL, s.KargoUcreti, s.HizmetBedeli, s.Stopaj, s.PosBedeli, IFNULL(s.CezaTutari,0), s.Durum, s.Barkod, IFNULL(u.UrunAdi, s.UrunAdi), IFNULL(u.Kategori, s.Kategori), IFNULL(u.Marka, s.Marka), IFNULL(s.Adet,1), s.Tarih, IFNULL(s.AlisFiyati,0), IFNULL(s.SatilanImei,'-') FROM Siparisler s LEFT JOIN Urunler u ON s.Barkod = u.Barkod WHERE s.SiparisNo = @sn";

                    using (var cmd = new SqliteCommand(q, conn)) { cmd.Parameters.AddWithValue("@sn", sipNo); using (var r = cmd.ExecuteReader()) { if (r.Read()) { rId = Convert.ToInt64(r[0]); sF = Convert.ToDouble(r[1]); km = Convert.ToDouble(r[2]); kr = Convert.ToDouble(r[3]); hz = Convert.ToDouble(r[4]); st = Convert.ToDouble(r[5]); ps = Convert.ToDouble(r[6]); cz = Convert.ToDouble(r[7]); durum = r[8].ToString(); brkd = r[9].ToString(); uAdi = r[10].ToString(); kat = r[11].ToString(); mrk = r[12].ToString(); adet = Convert.ToInt32(r[13]); tar = r[14].ToString(); alisFiyati = Convert.ToDouble(r[15]); imei = r[16].ToString(); } else return; } }

                }

            }

            catch { return; }

            Form frm = new Form { Text = $"Düzeltme Masası - No: {sipNo}", Size = new Size(520, 660), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };



            Panel pnlMeta = new Panel { Location = new Point(15, 10), Width = 475, Height = 105, BackColor = Color.FromArgb(24, 24, 27), BorderStyle = BorderStyle.FixedSingle };

            string temizBarkod = (string.IsNullOrEmpty(brkd) || brkd.ToUpper() == uAdi.ToUpper()) ? "-" : brkd;

            Label lblDurumX = new Label { Text = $"● SİSTEM DURUMU: {durum.ToUpper()}", Location = new Point(15, 12), AutoSize = true, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.Gold };

            Label lblKimlikX = new Label { Text = $"📦 Ürün: {uAdi} (Adet: {adet})\r\n🆔 Barkod: {temizBarkod} | Kategori: {kat} | Marka: {mrk}", Location = new Point(15, 36), Size = new Size(445, 38), Font = new Font("Segoe UI", 9F, FontStyle.Regular), ForeColor = Color.FromArgb(56, 189, 248) };

            Label lblFinansX = new Label { Text = $"💰 Kayıtlı Finansal Geçmiş -> Tarih: {tar} | Alış: {alisFiyati:N2}₺ | IMEI: {imei}", Location = new Point(15, 78), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.FromArgb(34, 197, 94) };

            pnlMeta.Controls.Add(lblDurumX); pnlMeta.Controls.Add(lblKimlikX); pnlMeta.Controls.Add(lblFinansX); frm.Controls.Add(pnlMeta);



            int currentY = 125;

            GroupBox grpGelir = new GroupBox { Text = " Satış Gelirleri", Location = new Point(15, currentY), Width = 475, Height = 70, ForeColor = Color.FromArgb(34, 197, 94), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

            TextBox tSat; grpGelir.Controls.Add(FormElemaniYaratRenk("Satış Fiyatı (TL):", out tSat, new Point(15, 18), 440, sF.ToString("N2"), Color.White)); frm.Controls.Add(grpGelir); currentY += 75;

            GroupBox grpMaliyet = new GroupBox { Text = " Ürün Maliyet & Envanter Bilgisi", Location = new Point(15, currentY), Width = 475, Height = 70, ForeColor = Color.FromArgb(56, 189, 248), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

            TextBox tAlis, tImei; grpMaliyet.Controls.Add(FormElemaniYaratRenk("Alış Maliyeti (TL):", out tAlis, new Point(15, 18), 210, alisFiyati.ToString("N2"), Color.White)); grpMaliyet.Controls.Add(FormElemaniYaratRenk("IMEI Numarası:", out tImei, new Point(245, 18), 210, imei, Color.White)); frm.Controls.Add(grpMaliyet); currentY += 75;

            GroupBox grpGider = new GroupBox { Text = " Kesintiler & Platform Vergileri", Location = new Point(15, currentY), Width = 475, Height = 195, ForeColor = Color.FromArgb(239, 68, 68), Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

            TextBox tKoms, tCez, tKrg, tHzm, tStpj, tPos;

            grpGider.Controls.Add(FormElemaniYaratRenk("Komisyon (TL):", out tKoms, new Point(15, 18), 210, km.ToString("N2"), Color.White)); grpGider.Controls.Add(FormElemaniYaratRenk("Kargo Ücreti (TL):", out tKrg, new Point(245, 18), 210, kr.ToString("N2"), Color.White));

            grpGider.Controls.Add(FormElemaniYaratRenk("Hizmet Bedeli (TL):", out tHzm, new Point(15, 74), 210, hz.ToString("N2"), Color.White)); grpGider.Controls.Add(FormElemaniYaratRenk("Ceza Tutarı (TL):", out tCez, new Point(245, 74), 210, cz.ToString("N2"), Color.White));

            grpGider.Controls.Add(FormElemaniYaratRenk("Stopaj (TL):", out tStpj, new Point(15, 130), 210, st.ToString("N2"), Color.White)); grpGider.Controls.Add(FormElemaniYaratRenk("POS Kesintisi (TL):", out tPos, new Point(245, 130), 210, ps.ToString("N2"), Color.White));

            frm.Controls.Add(grpGider); currentY += 205;



            frm.Controls.Add(new Label { Text = $"💡 Dosyadan Tahmin Edilen Brüt Hakediş: {exH:N2} ₺", Location = new Point(15, currentY), AutoSize = true, ForeColor = Color.LightSkyBlue, Font = new Font("Segoe UI", 9F, FontStyle.Bold) });

            Button btn = new Button { Text = "💾 HESAPLA VE SİSTEMİ GÜNCELLE", Location = new Point(15, currentY + 25), Width = 475, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, Font = new Font("Segoe UI", 10F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };

            btn.FlatAppearance.BorderSize = 0;

            btn.Click += (s, e) => {

                try

                {

                    double.TryParse(tSat.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n1); double.TryParse(tKoms.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n2); double.TryParse(tKrg.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n4); double.TryParse(tHzm.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n5); double.TryParse(tStpj.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n6); double.TryParse(tPos.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n7); double.TryParse(tCez.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n3); double.TryParse(tAlis.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nAlis); string nImei = tImei.Text.Trim();

                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET SatisFiyati=@sf, AlisFiyati=@af, SatilanImei=@im, KomisyonTL=@km, CezaTutari=@cz, KargoUcreti=@kr, HizmetBedeli=@hz, Stopaj=@st, PosBedeli=@pb WHERE rowid=@id", b)) { cmd.Parameters.AddWithValue("@sf", n1); cmd.Parameters.AddWithValue("@af", nAlis); cmd.Parameters.AddWithValue("@im", nImei); cmd.Parameters.AddWithValue("@km", n2); cmd.Parameters.AddWithValue("@cz", n3); cmd.Parameters.AddWithValue("@kr", n4); cmd.Parameters.AddWithValue("@hz", n5); cmd.Parameters.AddWithValue("@st", n6); cmd.Parameters.AddWithValue("@pb", n7); cmd.Parameters.AddWithValue("@id", rId); cmd.ExecuteNonQuery(); } }

                    foreach (DataGridViewRow rowX in anaGrid.Rows) { if (rowX.Cells["Sipariş No"].Value.ToString() == sipNo) { rowX.Cells["Dedektif (Durum)"].Value = "✔ KUSURSUZ"; } }

                    anaGrid.Refresh(); MessageBox.Show("Siparişin tüm kârlılık matrisi başarıyla güncellendi ve kırmızılık söndürüldü!"); frm.Close();

                }

                catch (Exception exAmeliyat) { MessageBox.Show(exAmeliyat.Message); }

            };

            frm.Controls.Add(btn); frm.ShowDialog();

        }



        private Panel FormElemaniYaratRenk(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan, Color baslikRenk)

        {

            Panel p = new Panel { Location = konum, Width = genislik, Height = 48 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = baslikRenk };

            kutu = new TextBox { Text = varsayilan, Location = new Point(0, 16), Width = genislik, Font = new Font("Segoe UI", 9.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p;

        }



        private void DynBtnSiparisSorgula_Click(object sender, EventArgs e)

        {

            var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList(); if (sips.Count == 0) return;

            try

            {

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))

                {

                    conn.Open(); var pNames = new List<string>(); var cmd = new SqliteCommand { Connection = conn };

                    for (int i = 0; i < sips.Count; i++) { string p = "@sn" + i; pNames.Add(p); cmd.Parameters.AddWithValue(p, sips[i]); }

                    cmd.CommandText = $@"SELECT Tarih, PlatformAd, SiparisNo, UrunAdi, Adet, SatisFiyati, Durum FROM Siparisler WHERE SiparisNo IN ({string.Join(", ", pNames)})";

                    DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader());

                    Form f = new Form { Text = "Önizleme", Size = new Size(900, 400) }; DataGridView d = new DataGridView { Dock = DockStyle.Fill, DataSource = dt }; f.Controls.Add(d); f.ShowDialog();

                }

            }

            catch (Exception exSorgu) { MessageBox.Show(exSorgu.Message); }

        }



        private void PlatformlariYukle()

        {

            try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT Ad FROM Platformlar", conn)) using (var r = cmd.ExecuteReader()) { dynCmbPlatform.Items.Clear(); dynCmbDosyaPlatformu.Items.Clear(); dynCmbPlatform.Items.Add(""); dynCmbDosyaPlatformu.Items.Add(""); while (r.Read()) { string s = r[0].ToString().ToUpper(); dynCmbPlatform.Items.Add(s); dynCmbDosyaPlatformu.Items.Add(s); } } } } catch { }

            if (dynCmbPlatform.Items.Count <= 1) { dynCmbPlatform.Items.AddRange(new object[] { "TRENDYOL", "HEPSİBURADA" }); dynCmbDosyaPlatformu.Items.AddRange(new object[] { "TRENDYOL", "HEPSİBURADA" }); }

            dynCmbPlatform.SelectedIndex = 0; dynCmbDosyaPlatformu.SelectedIndex = 0;

        }



        private void TahsilatGecmisiniListele()

        {

            try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT Id, Tarih, Platform, Tutar as 'Yatan Net Tutar (TL)', SiparisAdedi as 'Kapatılan Sipariş' FROM PlatformTahsilatlari ORDER BY Id DESC", conn)) { DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader()); dynDgvTahsilatGecmisi.DataSource = dt; if (dynDgvTahsilatGecmisi.Columns["Id"] != null) dynDgvTahsilatGecmisi.Columns["Id"].Visible = false; } } } catch { }

        }



        private void CanliBilançoListele()

        {

            try

            {

                DataTable dt = new DataTable(); dt.Columns.Add("Kanal"); dt.Columns.Add("Ödenen (Adet)", typeof(int)); dt.Columns.Add("Net TL", typeof(double)); dt.Columns.Add("Bekleyen (Adet)", typeof(int)); dt.Columns.Add("Kalan Net", typeof(double));

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))

                {

                    conn.Open(); using (var cmd = new SqliteCommand("SELECT s.PlatformAd, SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN s.SatisFiyati ELSE 0 END), SUM(CASE WHEN s.Durum!='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END), SUM(CASE WHEN s.Durum!='TAHSİL EDİLDİ' THEN s.SatisFiyati ELSE 0 END) FROM Siparisler s GROUP BY s.PlatformAd", conn))

                    using (var r = cmd.ExecuteReader()) { while (r.Read()) { dt.Rows.Add(r[0].ToString(), Convert.ToInt32(r[1]), Convert.ToDouble(r[2]), Convert.ToInt32(r[3]), Convert.ToDouble(r[4])); } }

                }

                dgvPlatformBilanço.DataSource = dt;

            }

            catch { }

        }



        private DataGridView GridYarat()

        {

            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9F) };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgv.DefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); return dgv;

        }



        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan = "")

        {

            Panel p = new Panel { Location = konum, Width = genislik, Height = 52 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };

            kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p;

        }



        private void CustomVIPCardInsaEt(TableLayoutPanel tlp, string b, Color c, int s, out Label lA, out Label lP)

        {

            Panel p = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Margin = new Padding(6) }; Panel pL = new Panel { Dock = DockStyle.Left, Width = 5, BackColor = c }; p.Controls.Add(pL);

            Label lblB = new Label { Text = b, Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.DarkGray };

            lP = new Label { Text = "0,00 ₺", Dock = DockStyle.Top, Height = 25, Font = new Font("Segoe UI Black", 13F, FontStyle.Bold), ForeColor = Color.White };

            lA = new Label { Text = "0 Sipariş", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 8.5F, FontStyle.Italic), ForeColor = Color.Gray };

            p.Controls.Add(lA); p.Controls.Add(lP); p.Controls.Add(lblB); tlp.Controls.Add(p, s, 0);

        }

    }

}

