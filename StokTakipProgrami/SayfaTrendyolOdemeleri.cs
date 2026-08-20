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

namespace StokTakipProgrami
{
    public class TyDbSiparis
    {
        public string Durum { get; set; }
        public double Satis { get; set; }
        public double Komisyon { get; set; }
        public double Kargo { get; set; }
        public double Hizmet { get; set; }
        public double Ceza { get; set; }
    }

    public partial class SayfaTrendyolOdemeleri : UserControl
    {
        private DateTimePicker dynDtTarih;
        private TextBox dynTxtNetGelenPara, dynTxtSiparisListesi;
        private Button dynBtnOdemeyiMuhurle, dynBtnKaydiSil, dynBtnExcel, dynBtnSiparisSorgula, dynBtnTemizle;
        private DataGridView dynDgvTahsilatGecmisi, dgvPlatformBilanço;
        private Label lblCardTahsilEdilenAdet, lblCardTahsilEdilenPara, lblCardBekleyenAdet, lblCardBekleyenPara, lblCardToplamAdet, lblCardToplamPara;
        private Label lblCardCiroAdet, lblCardCiroPara;
        private Label lblCardKarAdet, lblCardKarPara;
        private Button dynBtnYukluDosyaAc;
        private Button dynBtnKargoHizmetMutabakati; // 🚀 YENİ EKLENEN BUTON
        private DataTable sonCozumlenenTablo = null;

        public SayfaTrendyolOdemeleri() { InitializeComponent(); this.Load += SayfaTrendyolOdemeleri_Load; }

        private void SayfaTrendyolOdemeleri_Load(object sender, EventArgs e) { TabloyuGuvenceyeAl(); ArayuzInsaEt(); TahsilatGecmisiniListele(); CanliBilançoListele(); }

        private void TabloyuGuvenceyeAl() { try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); string sql = @"CREATE TABLE IF NOT EXISTS PlatformTahsilatlari (Id INTEGER PRIMARY KEY AUTOINCREMENT, Platform TEXT, Tarih TEXT, Tutar REAL, SiparisAdedi INTEGER, SiparisNumaralari TEXT);"; using (var cmd = new SqliteCommand(sql, conn)) { cmd.ExecuteNonQuery(); } } } catch { } }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear(); this.Dock = DockStyle.Fill; this.BackColor = Color.FromArgb(33, 33, 36);
            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(230, 126, 34) };
            Label lblT = new Label { Text = "🧡 TRENDYOL MUTABAKAT, HAKEDİŞ ÖDEMELERİ", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT); this.Controls.Add(pnlHeader);
            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380F)); mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); this.Controls.Add(mainLayout);
            Panel pnlCard = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(15), AutoScroll = true };

            int curY = 12;
            Label lbl1 = new Label { Text = "Ödeme Tarihi:", Location = new Point(15, curY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            pnlCard.Controls.Add(lbl1);
            dynDtTarih = new DateTimePicker { Location = new Point(15, curY + 20), Width = 340, Font = new Font("Segoe UI", 9.5F), Format = DateTimePickerFormat.Short };
            pnlCard.Controls.Add(dynDtTarih); curY += 50;
            pnlCard.Controls.Add(FormElemaniYarat("Hesaba Yatan Toplam Tutar (TL):", out dynTxtNetGelenPara, new Point(15, curY), 340)); curY += 50;

            Label lbl3 = new Label { Text = "Sipariş Numaraları:", Location = new Point(15, curY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            pnlCard.Controls.Add(lbl3);
            dynTxtSiparisListesi = new TextBox { Location = new Point(15, curY + 20), Width = 340, Height = 85, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 9.5F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };
            pnlCard.Controls.Add(dynTxtSiparisListesi); curY += 110;

            dynBtnTemizle = new Button { Text = "🧹 TEMİZLE", Location = new Point(15, curY), Width = 95, Height = 30, BackColor = Color.FromArgb(50, 50, 55), ForeColor = Color.LightGray, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnTemizle.FlatAppearance.BorderSize = 0; dynBtnTemizle.Click += (s, e) => { dynTxtSiparisListesi.Clear(); dynTxtNetGelenPara.Clear(); }; pnlCard.Controls.Add(dynBtnTemizle);

            dynBtnSiparisSorgula = new Button { Text = "🔍 SİPARİŞ NUMARALARINI SORGULA", Location = new Point(116, curY), Width = 239, Height = 30, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSiparisSorgula.FlatAppearance.BorderSize = 0; dynBtnSiparisSorgula.Click += DynBtnSiparisSorgula_Click; pnlCard.Controls.Add(dynBtnSiparisSorgula); curY += 40;

            Panel pnlDiv = new Panel { Location = new Point(15, curY), Width = 340, Height = 1, BackColor = Color.FromArgb(45, 45, 48) }; pnlCard.Controls.Add(pnlDiv); curY += 10;

            Label lbl4 = new Label { Text = "Ekstre Yükle:", Location = new Point(15, curY), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(56, 189, 248) };
            pnlCard.Controls.Add(lbl4); curY += 22;

            dynBtnYukluDosyaAc = new Button { Text = "📂 YÜKLÜ DOSYAYI TEKRAR AÇ", Location = new Point(15, curY), Width = 340, Height = 28, BackColor = Color.FromArgb(38, 38, 41), ForeColor = Color.Gold, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Visible = false, Cursor = Cursors.Hand };
            dynBtnYukluDosyaAc.FlatAppearance.BorderSize = 1; dynBtnYukluDosyaAc.FlatAppearance.BorderColor = Color.Gold; dynBtnYukluDosyaAc.Click += (s, e) => { if (sonCozumlenenTablo != null) MutabakatUyariEkraniniAc(sonCozumlenenTablo, "TRENDYOL"); };
            pnlCard.Controls.Add(dynBtnYukluDosyaAc);

            Button dynBtnMutabakatYukle = new Button { Text = "📂 EKSTRE SEÇ VE YÜKLE", Location = new Point(15, curY + 33), Width = 340, Height = 32, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnMutabakatYukle.FlatAppearance.BorderSize = 0; dynBtnMutabakatYukle.Click += DynBtnDosyaYukle_Click; pnlCard.Controls.Add(dynBtnMutabakatYukle); curY += 80;

            Panel pnlAyrac = new Panel { Location = new Point(15, curY), Width = 340, Height = 1, BackColor = Color.FromArgb(50, 50, 55) }; pnlCard.Controls.Add(pnlAyrac); curY += 12;

            dynBtnOdemeyiMuhurle = new Button { Text = "✅ ÖDEMEYİ KAYDET", Location = new Point(15, curY), Width = 340, Height = 42, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnOdemeyiMuhurle.FlatAppearance.BorderSize = 0; dynBtnOdemeyiMuhurle.Click += DynBtnOdemeyiMuhurle_Click; pnlCard.Controls.Add(dynBtnOdemeyiMuhurle);

            curY += 65; // 🚀 MÜHÜR: Kargo butonu ile aradaki boşluk artırıldı (Biraz daha aşağı alındı)

            // 🚀 KARGO VE HİZMET BEDELİ MUTABAKATI BUTONU
            dynBtnKargoHizmetMutabakati = new Button { Text = "📦 KARGO VE HİZMET BEDELİ MUTABAKATI", Location = new Point(15, curY), Width = 340, Height = 42, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnKargoHizmetMutabakati.FlatAppearance.BorderSize = 0; dynBtnKargoHizmetMutabakati.Click += DynBtnKargoHizmetMutabakati_Click; pnlCard.Controls.Add(dynBtnKargoHizmetMutabakati);

            mainLayout.Controls.Add(pnlCard, 0, 0);

            TableLayoutPanel pnlRightLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
            pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 45F)); pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 35F)); pnlRightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 95F)); mainLayout.Controls.Add(pnlRightLayout, 1, 0);

            Panel pnlUpperWrapper = new Panel { Dock = DockStyle.Fill }; pnlUpperWrapper.Controls.Add(new Label { Text = "📋 SİSTEME İŞLENEN TOPLU ÖDEMELER", Dock = DockStyle.Top, Height = 22, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.DarkGray });
            dynDgvTahsilatGecmisi = GridYarat(); dynDgvTahsilatGecmisi.CellDoubleClick += DynDgvTahsilatGecmisi_CellDoubleClick; pnlUpperWrapper.Controls.Add(dynDgvTahsilatGecmisi); dynDgvTahsilatGecmisi.BringToFront();

            Panel pnlRightBottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(0, 5, 0, 0) };
            dynBtnKaydiSil = new Button { Text = "🗑️ SEÇİLİ ÖDEMEYİ İPTAL ET", Dock = DockStyle.Left, Width = 380, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnKaydiSil.FlatAppearance.BorderSize = 0; dynBtnKaydiSil.Click += DynBtnKaydiSil_Click; pnlRightBottom.Controls.Add(dynBtnKaydiSil);

            dynBtnExcel = new Button { Text = "📊 EXCEL RAPORU AL", Dock = DockStyle.Right, Width = 130, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnExcel.FlatAppearance.BorderSize = 0; dynBtnExcel.Click += (s, e) => { }; pnlRightBottom.Controls.Add(dynBtnExcel); pnlUpperWrapper.Controls.Add(pnlRightBottom); pnlRightLayout.Controls.Add(pnlUpperWrapper, 0, 0);

            GroupBox grpBilanço = new GroupBox { Dock = DockStyle.Fill, Text = "📊 TRENDYOL CARİ MUTABAKAT VE HESAP TABLOSU", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange, Padding = new Padding(10, 22, 10, 10) };
            dgvPlatformBilanço = GridYarat(); grpBilanço.Controls.Add(dgvPlatformBilanço); pnlRightLayout.Controls.Add(grpBilanço, 0, 1);

            TableLayoutPanel tlpVIPCards = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 1, Padding = new Padding(0, 5, 0, 5) };
            for (int i = 0; i < 5; i++) tlpVIPCards.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));

            CustomVIPCardInsaEt(tlpVIPCards, "💰 TOPLAM TAHSİLAT", Color.FromArgb(34, 197, 94), 0, out lblCardTahsilEdilenAdet, out lblCardTahsilEdilenPara);
            CustomVIPCardInsaEt(tlpVIPCards, "⏳ KESİNTİLER", Color.Orange, 1, out lblCardBekleyenAdet, out lblCardBekleyenPara);
            CustomVIPCardInsaEt(tlpVIPCards, "💎 KALAN NET ALACAK", Color.FromArgb(56, 189, 248), 2, out lblCardToplamAdet, out lblCardToplamPara);
            CustomVIPCardInsaEt(tlpVIPCards, "📊 TOPLAM CİRO", Color.FromArgb(168, 85, 247), 3, out lblCardCiroAdet, out lblCardCiroPara);
            CustomVIPCardInsaEt(tlpVIPCards, "📈 NET KÂR DURUMU", Color.FromArgb(236, 72, 153), 4, out lblCardKarAdet, out lblCardKarPara);

            pnlRightLayout.Controls.Add(tlpVIPCards, 0, 2); mainLayout.BringToFront(); pnlHeader.SendToBack();
        }

        private bool PromptForInput(string t, string c, out string v) { Form p = new Form() { Width = 350, Height = 180, FormBorderStyle = FormBorderStyle.FixedDialog, Text = c, StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false }; TextBox ib = new TextBox() { Left = 20, Top = 50, Width = 290, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, Font = new Font("Segoe UI", 11F) }; Button cf = new Button() { Text = "Onayla", Left = 210, Width = 100, Top = 90, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold) }; string r = ""; bool ok = false; cf.Click += (s, e) => { r = ib.Text.Trim(); ok = true; p.Close(); }; p.Controls.Add(new Label() { Left = 20, Top = 20, Text = t, AutoSize = true, ForeColor = Color.White }); p.Controls.Add(ib); p.Controls.Add(cf); p.AcceptButton = cf; p.ShowDialog(); v = r; return ok; }
        private bool PromptForHurda(string dSn, out string oSn, out double oPo, out double oMz) { string tSn = dSn; double tP = 0, tZ = 0; Form p = new Form() { Width = 380, Height = 280, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "🛠️ Hurda Kapat", StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false }; TextBox txt1 = new TextBox() { Left = 20, Top = 35, Width = 320, Text = dSn, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White }; TextBox txt2 = new TextBox() { Left = 20, Top = 95, Width = 320, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White }; TextBox txt3 = new TextBox() { Left = 20, Top = 155, Width = 320, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White }; Button cf = new Button() { Text = "✔ Onayla", Left = 20, Width = 320, Top = 200, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat }; cf.FlatAppearance.BorderSize = 0; bool isOk = false; cf.Click += (s, e) => { tSn = txt1.Text.Trim(); double.TryParse(txt2.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out tP); double.TryParse(txt3.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out tZ); if (!string.IsNullOrEmpty(tSn)) { isOk = true; p.Close(); } }; p.Controls.Add(txt1); p.Controls.Add(txt2); p.Controls.Add(txt3); p.Controls.Add(cf); p.ShowDialog(); oSn = tSn; oPo = tP; oMz = tZ; return isOk; }

        private void DynDgvTahsilatGecmisi_CellDoubleClick(object sender, DataGridViewCellEventArgs e) { if (e.RowIndex < 0) return; int id = Convert.ToInt32(dynDgvTahsilatGecmisi.Rows[e.RowIndex].Cells["Id"].Value); string nolar = ""; try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT SiparisNumaralari FROM PlatformTahsilatlari WHERE Id=@id", conn)) { cmd.Parameters.AddWithValue("@id", id); nolar = cmd.ExecuteScalar()?.ToString() ?? ""; } } } catch { } Form f = new Form { Text = $"Makbuz Detayı", Size = new Size(380, 450), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false }; Label lblInfo = new Label { Text = "Bu hakedişteki sipariş numaraları:", Location = new Point(20, 20), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) }; TextBox t = new TextBox { Location = new Point(20, 50), Size = new Size(320, 330), Multiline = true, ScrollBars = ScrollBars.Vertical, Text = nolar.Replace(", ", Environment.NewLine), ReadOnly = true, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, Font = new Font("Consolas", 10.5F) }; f.Controls.Add(lblInfo); f.Controls.Add(t); f.ShowDialog(); }

        private void DynBtnKaydiSil_Click(object sender, EventArgs e) { if (dynDgvTahsilatGecmisi.SelectedRows.Count == 0) return; int id = Convert.ToInt32(dynDgvTahsilatGecmisi.SelectedRows[0].Cells["Id"].Value); if (MessageBox.Show("Bu tahsilat kaydını iptal etmek istediğinize emin misiniz?", "Hakediş İptal Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) { try { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var tx = b.BeginTransaction()) { string nolar = ""; using (var cmd = new SqliteCommand("SELECT SiparisNumaralari FROM PlatformTahsilatlari WHERE Id=@id", b, tx)) { cmd.Parameters.AddWithValue("@id", id); object sonuc = cmd.ExecuteScalar(); if (sonuc != null) nolar = sonuc.ToString(); } List<string> siparisler = nolar.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList(); int geriAlinanSayi = 0; string updateSql = "UPDATE Siparisler SET Durum = 'KARGOYA VERİLDİ' WHERE SiparisNo = @sn"; using (var cmdUpd = new SqliteCommand(updateSql, b, tx)) { foreach (var no in siparisler) { cmdUpd.Parameters.Clear(); cmdUpd.Parameters.AddWithValue("@sn", no); geriAlinanSayi += cmdUpd.ExecuteNonQuery(); } } using (var cmdDel = new SqliteCommand("DELETE FROM PlatformTahsilatlari WHERE Id = @id", b, tx)) { cmdDel.Parameters.AddWithValue("@id", id); cmdDel.ExecuteNonQuery(); } tx.Commit(); MessageBox.Show($"Hakediş kaydı iptal edildi! Alacak sınıfına geri aktarılan sipariş: {geriAlinanSayi}", "İşlem Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information); TahsilatGecmisiniListele(); CanliBilançoListele(); } } } catch (Exception exSil) { MessageBox.Show(exSil.Message); } } }

        private void DynBtnOdemeyiMuhurle_Click(object sender, EventArgs e) { if (string.IsNullOrEmpty(dynTxtNetGelenPara.Text)) return; try { double tutar = ExcelHucredenSayiya(dynTxtNetGelenPara.Text); string plat = "TRENDYOL"; string tar = dynDtTarih.Value.ToString("dd.MM.yyyy"); var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList(); List<string> temizSiparisler = new List<string>(); StringBuilder mükerrerRaporu = new StringBuilder(); using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); foreach (var s in sips) { using (var cmdCheck = new SqliteCommand(@"SELECT s.Tarih, s.SatisFiyati FROM Siparisler s WHERE s.SiparisNo=@sn AND s.Durum='TAHSİL EDİLDİ' AND UPPER(s.PlatformAd) LIKE '%TRENDYOL%'", b)) { cmdCheck.Parameters.AddWithValue("@sn", s); using (var reader = cmdCheck.ExecuteReader()) { if (reader.Read()) { string eskiTarih = reader["Tarih"].ToString(); string eskiTutar = reader["SatisFiyati"].ToString(); mükerrerRaporu.AppendLine($"⚠️ Sipariş No: {s} | Giriş Tarihi: {eskiTarih} | Sipariş Tutarı: {eskiTutar} TL"); } else { temizSiparisler.Add(s); } } } } } if (temizSiparisler.Count == 0 && mükerrerRaporu.Length > 0) { FormMükerrerDetayGoster(mükerrerRaporu.ToString(), 0); return; } if (temizSiparisler.Count > 0) { using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("INSERT INTO PlatformTahsilatlari (Platform, Tarih, Tutar, SiparisAdedi, SiparisNumaralari) VALUES (@p, @t, @v, @a, @n)", b)) { cmd.Parameters.AddWithValue("@p", plat); cmd.Parameters.AddWithValue("@t", tar); cmd.Parameters.AddWithValue("@v", tutar); cmd.Parameters.AddWithValue("@a", temizSiparisler.Count); cmd.Parameters.AddWithValue("@n", string.Join(", ", temizSiparisler)); cmd.ExecuteNonQuery(); } foreach (var s in temizSiparisler) { using (var cmd = new SqliteCommand("UPDATE Siparisler SET Durum='TAHSİL EDİLDİ' WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", b)) { cmd.Parameters.AddWithValue("@sn", s); cmd.ExecuteNonQuery(); } } } } dynTxtNetGelenPara.Clear(); dynTxtSiparisListesi.Clear(); sonCozumlenenTablo = null; dynBtnYukluDosyaAc.Visible = false; TahsilatGecmisiniListele(); CanliBilançoListele(); if (mükerrerRaporu.Length > 0) { FormMükerrerDetayGoster(mükerrerRaporu.ToString(), temizSiparisler.Count); } else { MessageBox.Show("Ödeme Başarıyla Kaydedildi!", "Kayıt Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information); } } catch (Exception exMuhur) { MessageBox.Show(exMuhur.Message); } }

        private static void FormMükerrerDetayGoster(string raporIcerigi, int basariliAdet) { Form frmRapor = new Form { Text = "🛑 Çapraz Mutabakat Mükerrer Koruma Kalkanı", Size = new Size(680, 420), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false }; Label lblInfo = new Label { Text = $"⚠️ Aşağıdaki siparişler daha önceden TAHSİL EDİLDİĞİ için tekrar kaydedilmedi.\nTemiz olan geriye kalan {basariliAdet} yeni sipariş ise başarıyla sisteme mühürlendi patron!", Dock = DockStyle.Top, Height = 50, ForeColor = Color.LightCoral, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Padding = new Padding(5) }; TextBox txtRapor = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, BackColor = Color.FromArgb(24, 24, 27), ForeColor = Color.Orange, Font = new Font("Consolas", 10F), Text = raporIcerigi }; frmRapor.Controls.Add(txtRapor); frmRapor.Controls.Add(lblInfo); frmRapor.ShowDialog(); }

        private double ExcelHucredenSayiya(object hucre) { if (hucre == null || hucre == DBNull.Value) return 0; if (hucre is double d) return d; if (hucre is decimal dc) return Convert.ToDouble(dc); if (hucre is int i) return i; if (hucre is long lg) return lg; if (hucre is float fl) return fl; string s = hucre.ToString().Replace("₺", "").Replace("TL", "").Replace(" ", "").Replace("\xA0", "").Trim(); if (string.IsNullOrEmpty(s)) return 0; int sonNokta = s.LastIndexOf('.'); int sonVirgul = s.LastIndexOf(','); if (sonNokta > sonVirgul) { if (s.Substring(sonNokta).Length <= 3) { s = s.Substring(0, sonNokta).Replace(".", "").Replace(",", "") + "." + s.Substring(sonNokta + 1); } else { s = s.Replace(".", "").Replace(",", ""); } } else if (sonVirgul > sonNokta) { if (s.Substring(sonVirgul).Length <= 3) { s = s.Substring(0, sonVirgul).Replace(".", "").Replace(",", "") + "." + s.Substring(sonVirgul + 1); } else { s = s.Replace(".", "").Replace(",", ""); } } else { s = s.Replace(".", "").Replace(",", ""); } if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double resFinal)) return resFinal; return 0; }

        private DataTable ExcelVeCsvOku(string dosyaYolu) { DataTable dt = new DataTable(); string ext = Path.GetExtension(dosyaYolu).ToLower(); if (ext == ".xls" || ext == ".xlsx") { string connStr = ext == ".xls" ? $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dosyaYolu};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\";" : $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dosyaYolu};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";"; using (OleDbConnection conn = new OleDbConnection(connStr)) { conn.Open(); DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null); string sn = schema.Rows[0]["TABLE_NAME"].ToString(); if (sn.StartsWith("'") && sn.EndsWith("'")) sn = sn.Substring(1, sn.Length - 2); using (OleDbDataAdapter da = new OleDbDataAdapter($"SELECT * FROM [{sn}]", conn)) { da.Fill(dt); } } } else { string[] lines = File.ReadAllLines(dosyaYolu, System.Text.Encoding.Default); if (lines.Length > 0) { char ayrac = lines[0].Contains(";") ? ';' : ','; string[] headers = SatiriParcala(lines[0], ayrac); foreach (string h in headers) dt.Columns.Add(h.Trim()); for (int i = 1; i < lines.Length; i++) { if (string.IsNullOrWhiteSpace(lines[i])) continue; string[] cells = SatiriParcala(lines[i], ayrac); DataRow row = dt.NewRow(); for (int j = 0; j < cells.Length && j < dt.Columns.Count; j++) row[j] = cells[j].Trim(); dt.Rows.Add(row); } } } return dt; }

        private string[] SatiriParcala(string satir, char ayrac) { var res = new List<string>(); bool ti = false; string hc = ""; for (int i = 0; i < satir.Length; i++) { char c = satir[i]; if (c == '"') ti = !ti; else if (c == ayrac && !ti) { res.Add(hc.Trim('"')); hc = ""; } else hc += c; } res.Add(hc.Trim('"')); return res.ToArray(); }

        // 🚀 1. ÖZEL KARGO HİZMET MUTABAKATI BUTONU METODU
        private void DynBtnKargoHizmetMutabakati_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Dosyaları|*.xlsx;*.xls;*.csv";
                ofd.Title = "Trendyol Kargo ve Hizmet Faturalarını Seçin";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        DataTable dt = ExcelVeCsvOku(ofd.FileName);
                        int idxSip = -1, idxGonderiKrg = -1, idxIadeKrg = -1, idxPltHizmet = -1;

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            string b = dt.Columns[i].ColumnName.ToUpper().Trim();
                            if (b == "SİPARİŞ NO" || b == "SIPARIS NO" || b == "SİPARİŞ NUMARASI") idxSip = i;
                            else if (b == "GÖNDERİ KARGO BEDELİ" || b == "GONDERİ KARGO" || b.Contains("GÖNDERİ")) idxGonderiKrg = i;
                            else if (b == "İADE KARGO BEDELİ" || b == "IADE KARGO" || b.Contains("İADE")) idxIadeKrg = i;
                            else if (b == "PLATFORM HİZMET BEDELİ" || b == "PLATFORM HIZMET" || b.Contains("PLATFORM")) idxPltHizmet = i;
                        }

                        if (idxSip == -1) { MessageBox.Show("Sipariş numarası sütunu bulunamadı!"); return; }

                        DataTable dtMutabakat = new DataTable();
                        dtMutabakat.Columns.Add("Sipariş No");
                        dtMutabakat.Columns.Add("Sistem Durumu");
                        dtMutabakat.Columns.Add("Dosya Kargo (TL)", typeof(double));
                        dtMutabakat.Columns.Add("Sistem Kargo (TL)", typeof(double));
                        dtMutabakat.Columns.Add("Dosya Hizmet (TL)", typeof(double));
                        dtMutabakat.Columns.Add("Sistem Hizmet (TL)", typeof(double));
                        dtMutabakat.Columns.Add("Dedektif (Durum)");

                        int hatalar = 0;

                        using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                        {
                            conn.Open();
                            foreach (DataRow row in dt.Rows)
                            {
                                string sn = row.ItemArray[idxSip]?.ToString().Trim() ?? "";
                                if (string.IsNullOrEmpty(sn) || sn == "-" || sn.ToUpper().Contains("TOPLAM") || sn.ToUpper().Contains("TOTAL")) continue;
                                if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);

                                double exGonderi = idxGonderiKrg != -1 ? Math.Abs(ExcelHucredenSayiya(row.ItemArray[idxGonderiKrg])) : 0;
                                double exIade = idxIadeKrg != -1 ? Math.Abs(ExcelHucredenSayiya(row.ItemArray[idxIadeKrg])) : 0;
                                double exTotalKargo = exGonderi + exIade;
                                double exHizmet = idxPltHizmet != -1 ? Math.Abs(ExcelHucredenSayiya(row.ItemArray[idxPltHizmet])) : 0;

                                using (var cmd = new SqliteCommand("SELECT Durum, KargoUcreti, HizmetBedeli FROM Siparisler WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", conn))
                                {
                                    cmd.Parameters.AddWithValue("@sn", sn);
                                    using (var r = cmd.ExecuteReader())
                                    {
                                        if (r.Read())
                                        {
                                            string sysDurum = r[0].ToString();
                                            double dbKargo = Convert.ToDouble(r[1]);
                                            double dbHizmet = Convert.ToDouble(r[2]);

                                            string ded = "✔ MÜTABIK";
                                            bool kargoHatali = Math.Abs(exTotalKargo - dbKargo) > 0.001;
                                            bool hizmetHatali = Math.Abs(exHizmet - dbHizmet) > 0.001;

                                            if (kargoHatali && hizmetHatali) ded = "HATA: KARGO VE HİZMET FARKI";
                                            else if (kargoHatali) ded = "HATA: KARGO FARKI";
                                            else if (hizmetHatali) ded = "HATA: HİZMET FARKI";

                                            if (ded.StartsWith("HATA:")) hatalar++;

                                            dtMutabakat.Rows.Add(sn, sysDurum, exTotalKargo, dbKargo, exHizmet, dbHizmet, ded);
                                        }
                                        else
                                        {
                                            dtMutabakat.Rows.Add(sn, "KAYIT YOK", exTotalKargo, 0, exHizmet, 0, "HATA: SİPARİŞ BULUNAMADI");
                                            hatalar++;
                                        }
                                    }
                                }
                            }
                        }

                        if (hatalar > 0) MessageBox.Show($"DİKKAT! {hatalar} siparişte Kargo veya Hizmet faturası farkı tespit edildi.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        KargoHizmetMutabakatEkraniniAc(dtMutabakat);
                    }
                    catch (Exception exDosya) { MessageBox.Show("Sistem Hatası: " + exDosya.Message); }
                }
            }
        }

        // 🚀 2. KARGO/HİZMET ÖZEL EKRAN METODU
        private void KargoHizmetMutabakatEkraniniAc(DataTable dtSource)
        {
            Form f = new Form { Text = "📦 Özel Kargo ve Hizmet Mutabakat Ekranı", Size = new Size(1100, 600), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
            Label lbl = new Label { Text = "KARGO VE HİZMET BEDELİ KIYASLAMASI (Çift Tık: İnceleme Masası | Sağ Tık: Eşitleme ve Hata Detayı)", Dock = DockStyle.Top, Height = 40, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(56, 189, 248), Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };

            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EnableHeadersVisualStyles = false, RowHeadersVisible = false };
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DefaultCellStyle.SelectionForeColor = Color.White; dgv.ColumnHeadersHeight = 40; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgv.DataSource = dtSource;
            if (dgv.Columns.Contains("Dosya Kargo (TL)")) dgv.Columns["Dosya Kargo (TL)"].DefaultCellStyle.Format = "N2";
            if (dgv.Columns.Contains("Sistem Kargo (TL)")) dgv.Columns["Sistem Kargo (TL)"].DefaultCellStyle.Format = "N2";
            if (dgv.Columns.Contains("Dosya Hizmet (TL)")) dgv.Columns["Dosya Hizmet (TL)"].DefaultCellStyle.Format = "N2";
            if (dgv.Columns.Contains("Sistem Hizmet (TL)")) dgv.Columns["Sistem Hizmet (TL)"].DefaultCellStyle.Format = "N2";

            // 🚀 EKLENEN ÇİFT TIK ÖZELLİĞİ
            dgv.CellDoubleClick += (s, eDouble) => {
                if (eDouble.RowIndex < 0) return;
                var r = dgv.Rows[eDouble.RowIndex];
                string sn = dgv.Columns.Contains("Sipariş No") ? r.Cells["Sipariş No"].Value?.ToString() ?? "" : "";
                if (!string.IsNullOrEmpty(sn) && sn != "-")
                {
                    SiparisiAmeliyatMasasinaYatir(sn, 0, 0, dgv);
                }
            };

            dgv.DataBindingComplete += (s, eBind) => {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    string ded = row.Cells["Dedektif (Durum)"].Value?.ToString() ?? "";
                    if (ded == "✔ MÜTABIK" || ded == "✔ EŞİTLENDİ") { row.DefaultCellStyle.BackColor = Color.FromArgb(46, 117, 89); row.DefaultCellStyle.ForeColor = Color.White; }
                    else if (ded.StartsWith("HATA:")) { row.DefaultCellStyle.BackColor = Color.Crimson; row.DefaultCellStyle.ForeColor = Color.White; }
                    else { row.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); row.DefaultCellStyle.ForeColor = Color.White; }
                }
                dgv.ClearSelection();
            };

            ContextMenuStrip cms = new ContextMenuStrip();

            dgv.MouseDown += (s, eMouse) => {
                if (eMouse.Button == MouseButtons.Right)
                {
                    var hit = dgv.HitTest(eMouse.X, eMouse.Y);
                    if (hit.Type == DataGridViewHitTestType.Cell && hit.RowIndex >= 0)
                    {
                        dgv.ClearSelection(); dgv.Rows[hit.RowIndex].Selected = true; var r = dgv.Rows[hit.RowIndex];
                        string sn = r.Cells["Sipariş No"].Value?.ToString() ?? "";
                        string ded = r.Cells["Dedektif (Durum)"].Value?.ToString() ?? "";

                        cms.Items.Clear();

                        // HER DURUMDA AMELİYAT MASASI ÇIKSIN
                        ToolStripMenuItem btnAmeliyat = new ToolStripMenuItem("✏️ İncele ve Manuel Düzelt (Ameliyat Masası)");
                        btnAmeliyat.Click += (sx, exAmeliyat) => {
                            if (string.IsNullOrEmpty(sn) || sn == "-") return;
                            SiparisiAmeliyatMasasinaYatir(sn, 0, 0, dgv);
                        };
                        cms.Items.Add(btnAmeliyat);
                        cms.Items.Add(new ToolStripSeparator());

                        // 🚀 HATA VARSA DETAY VE EŞİTLEME ÇIKSIN
                        if (ded.StartsWith("HATA:") && !ded.Contains("BULUNAMADI"))
                        {
                            double dKrg = Convert.ToDouble(r.Cells["Dosya Kargo (TL)"].Value);
                            double sKrg = Convert.ToDouble(r.Cells["Sistem Kargo (TL)"].Value);
                            double dHzm = Convert.ToDouble(r.Cells["Dosya Hizmet (TL)"].Value);
                            double sHzm = Convert.ToDouble(r.Cells["Sistem Hizmet (TL)"].Value);

                            ToolStripMenuItem btnHataDetay = new ToolStripMenuItem("❌ HATAYI GÖSTER");
                            btnHataDetay.ForeColor = Color.Crimson;
                            btnHataDetay.Font = new Font(btnHataDetay.Font, FontStyle.Bold);
                            btnHataDetay.Click += (sx, ex) => {
                                string hataNedeni = "";
                                if (Math.Abs(dKrg - sKrg) > 0.001) hataNedeni += $"Kargo Farkı: \nDosya ({dKrg:N2} TL) | Sistem ({sKrg:N2} TL)\n\n";
                                if (Math.Abs(dHzm - sHzm) > 0.001) hataNedeni += $"Hizmet Farkı: \nDosya ({dHzm:N2} TL) | Sistem ({sHzm:N2} TL)";

                                MessageBox.Show(string.IsNullOrEmpty(hataNedeni) ? "Bilinmeyen bir hata oluştu." : hataNedeni, "Hata Detayı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            };
                            cms.Items.Add(btnHataDetay);
                            cms.Items.Add(new ToolStripSeparator());

                            ToolStripMenuItem btnEsitle = new ToolStripMenuItem("⚡ Dosyadaki Rakamları Sisteme İşle (Otomatik Eşitle)");
                            btnEsitle.Font = new Font(btnEsitle.Font, FontStyle.Bold);
                            btnEsitle.Click += (sx, ex) => {
                                if (MessageBox.Show($"[{sn}] numaralı siparişin;\n\nKargo Tutarını: {dKrg:N2} TL\nHizmet Tutarını: {dHzm:N2} TL\n\nOlarak sistemde eşitlemek ve kaydetmek istiyor musunuz?", "Otomatik Eşitleme Onayı", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    try
                                    {
                                        using (var bc = new SqliteConnection(Veritabanı.BaglantiYolu))
                                        {
                                            bc.Open();
                                            using (var cmdu = new SqliteCommand("UPDATE Siparisler SET KargoUcreti=@k, HizmetBedeli=@h WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", bc))
                                            {
                                                cmdu.Parameters.AddWithValue("@k", dKrg);
                                                cmdu.Parameters.AddWithValue("@h", dHzm);
                                                cmdu.Parameters.AddWithValue("@sn", sn);
                                                cmdu.ExecuteNonQuery();
                                            }
                                        }

                                        // ARAYÜZÜ DOSYA DEĞERLERİYLE GÜNCELLE VE MÜTABIK YAP
                                        r.Cells["Sistem Kargo (TL)"].Value = dKrg;
                                        r.Cells["Sistem Hizmet (TL)"].Value = dHzm;
                                        r.Cells["Dedektif (Durum)"].Value = "✔ EŞİTLENDİ";
                                        dgv.Refresh();
                                    }
                                    catch (Exception exTY) { MessageBox.Show(exTY.Message); }
                                }
                            };
                            cms.Items.Add(btnEsitle);
                        }

                        // 🚀 MENÜYÜ FARENİN TIKLANDIĞI NOKTADA AÇ
                        cms.Show(dgv, eMouse.Location);
                    }
                }
            };

            f.Controls.Add(dgv); f.Controls.Add(lbl);
            dgv.BringToFront(); f.ShowDialog();
        }

        // 🚀 SENİN ORİJİNAL SİPARİŞ ÖDEMELERİ YÜKLEME KODUN (HİÇ DOKUNULMADI!)
        private void DynBtnDosyaYukle_Click(object sender, EventArgs e)
        {
            string dp = "TRENDYOL";
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
                        if (idxSip == -1 || idxHak == -1) { MessageBox.Show("Yüklenen ekstre dosyasındaki hakediş sütun başlıkları sistem matrisiyle uyuşmuyor!"); return; }
                        var sips = new HashSet<string>(); double total = 0;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            int doluHucre = 0; for (int c = 0; c < dt.Columns.Count; c++) { if (!string.IsNullOrWhiteSpace(dt.Rows[i].ItemArray[c].ToString())) doluHucre++; }
                            string sn = dt.Rows[i].ItemArray[idxSip].ToString().Trim(); if (string.IsNullOrEmpty(sn) || sn == "-") continue;
                            if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);
                            if (sn != "-") sips.Add(sn); total += ExcelHucredenSayiya(dt.Rows[i].ItemArray[idxHak]);
                        }
                        var dbHavuz = new Dictionary<string, HbDbSiparis>();
                        using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                        {
                            conn.Open();
                            foreach (var s in sips)
                            {
                                using (var cmd = new SqliteCommand("SELECT Durum, SatisFiyati, KomisyonTL, KargoUcreti, HizmetBedeli, IFNULL(CezaTutari,0) FROM Siparisler WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", conn))
                                {
                                    cmd.Parameters.AddWithValue("@sn", s); using (var r = cmd.ExecuteReader()) { if (r.Read()) dbHavuz[s] = new HbDbSiparis { Durum = r[0].ToString(), Satis = Convert.ToDouble(r[1]), Komisyon = Convert.ToDouble(r[2]), Kargo = Convert.ToDouble(r[3]), Hizmet = Convert.ToDouble(r[4]), Ceza = Convert.ToDouble(r[5]) }; }
                                }
                            }
                        }
                        DataTable dtHam = new DataTable(); dtHam.Columns.Add("Sipariş No"); dtHam.Columns.Add("İşlem Tipi"); dtHam.Columns.Add("Açıklama / Ürün"); dtHam.Columns.Add("Tutar (TL)", typeof(double));
                        dtHam.Columns.Add("Komisyon (TL)", typeof(double));
                        dtHam.Columns.Add("Dedektif (Durum)"); dtHam.Columns.Add("Sistem Durumu");
                        int hatalar = 0;

                        foreach (DataRow row in dt.Rows)
                        {
                            string sn = row.ItemArray[idxSip].ToString().Trim(); if (sn.ToUpper().Contains("TOPLAM") || sn.ToUpper().Contains("TOTAL")) continue;
                            if (sn.EndsWith(",0") || sn.EndsWith(".0")) sn = sn.Substring(0, sn.Length - 2);
                            if (string.IsNullOrEmpty(sn)) sn = "-";
                            string isl = idxIslem != -1 ? row.ItemArray[idxIslem].ToString() : "SATIŞ"; double tut = ExcelHucredenSayiya(row.ItemArray[idxHak]); double kom = idxKom != -1 ? ExcelHucredenSayiya(row.ItemArray[idxKom]) : 0;
                            string sys = "-", ded = "✔ MÜTABIK";

                            if (dbHavuz.TryGetValue(sn, out var db))
                            {
                                sys = db.Durum;
                                if (db.Durum.ToUpper() == "TAHSİL EDİLDİ")
                                {
                                    ded = "HATA: MÜKERRER TAHSİLAT";
                                }
                                else if (!isl.ToUpper().Contains("KUPON") && !isl.ToUpper().Contains("İNDİRİM") && !isl.ToUpper().Contains("SATIŞ"))
                                {
                                    if (Math.Abs(Math.Abs(tut) - db.Ceza) > 0.1) ded = "HATA: CEZA FARKI";
                                }
                            }
                            else if (sn != "-") { sys = "SİPARİŞLER SİSTEMDE YOK"; ded = "HATA: KAYITSIZ"; }
                            else { sys = "-"; ded = "GENEL İŞLEM"; }

                            if (ded.StartsWith("HATA:")) hatalar++; dtHam.Rows.Add(sn, isl, idxAcik != -1 ? row.ItemArray[idxAcik].ToString() : "", tut, kom, ded, sys);
                        }

                        dynTxtSiparisListesi.Text = string.Join(Environment.NewLine, sips); dynTxtNetGelenPara.Text = total.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));

                        sonCozumlenenTablo = dtHam;
                        dynBtnYukluDosyaAc.Text = $"📂 YÜKLÜ DOSYAYI AÇ ({Path.GetFileName(ofd.FileName)})";
                        dynBtnYukluDosyaAc.Visible = true;

                        if (hatalar > 0) MessageBox.Show($"🚨 DİKKAT! MÜKERRER VEYA HATALI {hatalar} SATIRDA TAHSİLAT FARKI TESPİT EDİLDİ!", "Radar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        MutabakatUyariEkraniniAc(dtHam, dp);
                    }
                    catch (Exception exDosya) { MessageBox.Show(exDosya.Message); }
                }
            }
        }

        // 🚀 SENİN ORİJİNAL SAĞ TIK EKRANIN (HİÇ DOKUNULMADI!)
        private void MutabakatUyariEkraniniAc(DataTable dtSource, string dp)
        {
            Form f = new Form { Text = $"{dp} Ekstre Önizleme", Size = new Size(1350, 650), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
            Label lbl = new Label { Text = "SİPARİŞ DETAYLARI İNCELEME (Çift Tık: Ameliyat Masası | Sağ Tık: Hızlı Tutar Eşitleme ve Detaylar)", Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Orange, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };

            TableLayoutPanel tlpPopupLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(33, 33, 36) };
            tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            f.Controls.Add(tlpPopupLayout); tlpPopupLayout.Controls.Add(lbl, 0, 0);

            Panel pnlGridContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 10) };
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, EnableHeadersVisualStyles = false, RowHeadersVisible = false };
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DefaultCellStyle.SelectionForeColor = Color.White; dgv.ColumnHeadersHeight = 40; dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgv.DataSource = dtSource;
            if (dgv.Columns.Contains("HataDetay")) dgv.Columns["HataDetay"].Visible = false;

            if (dgv.Columns.Contains("Satış (TL)")) dgv.Columns["Satış (TL)"].DefaultCellStyle.Format = "N2";
            if (dgv.Columns.Contains("Komisyon (TL)")) dgv.Columns["Komisyon (TL)"].DefaultCellStyle.Format = "N2";
            if (dgv.Columns.Contains("Kargo (TL)")) dgv.Columns["Kargo (TL)"].DefaultCellStyle.Format = "N2";
            if (dgv.Columns.Contains("Hizmet (TL)")) dgv.Columns["Hizmet (TL)"].DefaultCellStyle.Format = "N2";

            dgv.CellDoubleClick += (s, eDouble) => {
                if (eDouble.RowIndex < 0) return;
                var r = dgv.Rows[eDouble.RowIndex]; string sn = "";
                if (dgv.Columns.Contains("Sipariş No")) sn = r.Cells["Sipariş No"].Value?.ToString() ?? "";
                else if (dgv.Columns.Contains("Sipariş Numarası")) sn = r.Cells["Sipariş Numarası"].Value?.ToString() ?? "";

                if (!string.IsNullOrEmpty(sn) && sn != "-")
                {
                    double tH = 0;
                    if (dgv.Columns.Contains("Satış (TL)")) tH = Convert.ToDouble(r.Cells["Satış (TL)"].Value);
                    SiparisiAmeliyatMasasinaYatir(sn, Math.Abs(tH), 0, dgv);
                }
            };

            dgv.DataBindingComplete += (s, eBind) => {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    string ded = dgv.Columns.Contains("Dedektif (Durum)") ? (row.Cells["Dedektif (Durum)"].Value?.ToString() ?? "") : "";
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
                    var hit = dgv.HitTest(eMouse.X, eMouse.Y); if (hit.Type == DataGridViewHitTestType.Cell && hit.RowIndex >= 0)
                    {
                        dgv.ClearSelection(); dgv.Rows[hit.RowIndex].Selected = true; var r = dgv.Rows[hit.RowIndex];
                        string sn = "";
                        if (dgv.Columns.Contains("Sipariş No")) sn = r.Cells["Sipariş No"].Value?.ToString() ?? "";
                        else if (dgv.Columns.Contains("Sipariş Numarası")) sn = r.Cells["Sipariş Numarası"].Value?.ToString() ?? "";

                        string dedektifDurum = dgv.Columns.Contains("Dedektif (Durum)") ? (r.Cells["Dedektif (Durum)"].Value?.ToString() ?? "") : "";
                        string hataAciklamasi = dgv.Columns.Contains("HataDetay") ? (r.Cells["HataDetay"].Value?.ToString() ?? "") : "";

                        cms.Items.Clear();

                        if (dedektifDurum.StartsWith("HATA:"))
                        {
                            ToolStripMenuItem btnHataDetay = new ToolStripMenuItem("❌ HATAYI GÖSTER");
                            btnHataDetay.ForeColor = Color.Crimson;
                            btnHataDetay.Font = new Font(btnHataDetay.Font, FontStyle.Bold);
                            btnHataDetay.Click += (sx, ex) => { MessageBox.Show(string.IsNullOrEmpty(hataAciklamasi) ? "Hata detayı bulunamadı." : hataAciklamasi, "Hata Detayı", MessageBoxButtons.OK, MessageBoxIcon.Warning); };
                            cms.Items.Add(btnHataDetay);
                            cms.Items.Add(new ToolStripSeparator());
                        }

                        ToolStripMenuItem btnAmeliyatTY = new ToolStripMenuItem("✏️ İncele ve Manuel Düzelt (Ameliyat Masası)");
                        btnAmeliyatTY.Click += (sx, exAmeliyatTY) => {
                            if (string.IsNullOrEmpty(sn) || sn == "-") return;
                            double tH = dgv.Columns.Contains("Satış (TL)") ? Math.Abs(Convert.ToDouble(r.Cells["Satış (TL)"].Value)) : 0;
                            SiparisiAmeliyatMasasinaYatir(sn, tH, 0, dgv);
                        }; cms.Items.Add(btnAmeliyatTY);

                        cms.Items.Add(new ToolStripSeparator());

                        try
                        {
                            double exKargo = 0, exHizmet = 0, exKom = 0, exSatis = 0;

                            if (dgv.Columns.Contains("Kargo (TL)")) exKargo += Math.Abs(ExcelHucredenSayiya(r.Cells["Kargo (TL)"].Value));
                            if (dgv.Columns.Contains("Gönderi Kargo Bedeli")) exKargo += Math.Abs(ExcelHucredenSayiya(r.Cells["Gönderi Kargo Bedeli"].Value));
                            if (dgv.Columns.Contains("İade Kargo Bedeli")) exKargo += Math.Abs(ExcelHucredenSayiya(r.Cells["İade Kargo Bedeli"].Value));

                            if (dgv.Columns.Contains("Hizmet (TL)")) exHizmet += Math.Abs(ExcelHucredenSayiya(r.Cells["Hizmet (TL)"].Value));
                            if (dgv.Columns.Contains("Platform Hizmet Bedeli")) exHizmet += Math.Abs(ExcelHucredenSayiya(r.Cells["Platform Hizmet Bedeli"].Value));

                            if (dgv.Columns.Contains("Komisyon (TL)")) exKom += Math.Abs(ExcelHucredenSayiya(r.Cells["Komisyon (TL)"].Value));
                            if (dgv.Columns.Contains("Satış (TL)")) exSatis += Math.Abs(ExcelHucredenSayiya(r.Cells["Satış (TL)"].Value));

                            using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                            {
                                b.Open();
                                using (var cmd = new SqliteCommand("SELECT SatisFiyati, KomisyonTL, KargoUcreti, HizmetBedeli FROM Siparisler WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", b))
                                {
                                    cmd.Parameters.AddWithValue("@sn", sn);
                                    using (var dbR = cmd.ExecuteReader())
                                    {
                                        if (dbR.Read())
                                        {
                                            double dbSatis = Convert.ToDouble(dbR[0]);
                                            double dbKom = Convert.ToDouble(dbR[1]);
                                            double dbKargo = Convert.ToDouble(dbR[2]);
                                            double dbHizmet = Convert.ToDouble(dbR[3]);

                                            bool kargoFarkli = Math.Abs(exKargo - dbKargo) > 0.001;
                                            bool hizmetFarkli = Math.Abs(exHizmet - dbHizmet) > 0.001;

                                            if (Math.Abs(exKom - dbKom) > 0.001)
                                            {
                                                ToolStripMenuItem mItem = new ToolStripMenuItem($"⚡ Komisyon Farkını Eşitle ({exKom:N2} TL)");
                                                mItem.Click += (sx, ex) => {
                                                    if (MessageBox.Show($"[{sn}] komisyon tutarını {exKom:N2} TL olarak güncellemek istiyor musunuz?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                                    {
                                                        using (var bc = new SqliteConnection(Veritabanı.BaglantiYolu))
                                                        {
                                                            bc.Open();
                                                            using (var cmdu = new SqliteCommand("UPDATE Siparisler SET KomisyonTL=@v WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", bc))
                                                            {
                                                                cmdu.Parameters.AddWithValue("@v", exKom);
                                                                cmdu.Parameters.AddWithValue("@sn", sn);
                                                                cmdu.ExecuteNonQuery();
                                                            }
                                                        }
                                                        if (dgv.Columns.Contains("Dedektif (Durum)")) r.Cells["Dedektif (Durum)"].Value = "✔ EŞİTLENDİ";
                                                        dgv.Refresh();
                                                    }
                                                };
                                                cms.Items.Add(mItem);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch { }

                        cms.Items.Add(new ToolStripSeparator());
                        ToolStripMenuItem btnCezaTY = new ToolStripMenuItem("⚠️ 'Ceza Tutarı' Olarak İşle");
                        btnCezaTY.Click += (sx, exCezaTY) => {
                            string targetSn = sn; if (string.IsNullOrEmpty(targetSn) || targetSn == "-") { if (PromptForInput("Sipariş No girin:", "Sipariş No", out string inpSn)) targetSn = inpSn; else return; }
                            double tut = dgv.Columns.Contains("Satış (TL)") ? Math.Abs(Convert.ToDouble(r.Cells["Satış (TL)"].Value)) : 0;
                            if (MessageBox.Show($"[{targetSn}] siparişe {tut:N2} TL ceza eklenecektir?", "Onay", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                try
                                {
                                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET CezaTutari=IFNULL(CezaTutari,0)+@c, CezaSebebi=CASE WHEN IFNULL(CezaSebebi,'')='' THEN @s ELSE CezaSebebi||' | '||@s END WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", b)) { cmd.Parameters.AddWithValue("@c", tut); cmd.Parameters.AddWithValue("@s", "MANUEL CEZA"); cmd.Parameters.AddWithValue("@sn", targetSn); cmd.ExecuteNonQuery(); } }
                                    if (dgv.Columns.Contains("Dedektif (Durum)")) r.Cells["Dedektif (Durum)"].Value = "KAYDEDİLDİ";
                                    dgv.Refresh();
                                }
                                catch (Exception exTY) { MessageBox.Show(exTY.Message); }
                            }
                        }; cms.Items.Add(btnCezaTY);

                        ToolStripMenuItem btnHurda = new ToolStripMenuItem("🛠️ İhtilaflı Siparişi Tamamla (Hurda/İade İşlemi)");
                        btnHurda.Click += (sx, exHurda) => {
                            string targetSn = sn; if (string.IsNullOrEmpty(targetSn) || targetSn == "-") return;
                            double tut = dgv.Columns.Contains("Satış (TL)") ? Math.Abs(Convert.ToDouble(r.Cells["Satış (TL)"].Value)) : 0;
                            if (MessageBox.Show($"{targetSn} numaralı sipariş İADE olarak kapatılıp, ürün (2. El) olarak {tut:N2} TL hurda bedeli düşülerek stoğa eklenecek. Onaylıyor musunuz?", "İhtilaf Tamamlama", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                try
                                {
                                    using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                                    {
                                        conn.Open(); string sBarkod = "", sUrunAdi = "", sDurum = ""; double sKoms = 0, sKrg = 0, sHzm = 0, sStpj = 0, sPos = 0;
                                        using (var c = new SqliteCommand("SELECT Barkod, UrunAdi, Durum, KomisyonTL, KargoUcreti, HizmetBedeli, Stopaj, PosBedeli FROM Siparisler WHERE SiparisNo=@sn AND UPPER(PlatformAd) LIKE '%TRENDYOL%'", conn))
                                        {
                                            c.Parameters.AddWithValue("@sn", targetSn);
                                            using (var rd = c.ExecuteReader()) { if (rd.Read()) { sBarkod = rd["Barkod"].ToString(); sUrunAdi = rd["UrunAdi"].ToString(); sDurum = rd["Durum"].ToString().ToUpper(); sKoms = Convert.ToDouble(rd["KomisyonTL"]); sKrg = Convert.ToDouble(rd["KargoUcreti"]); sHzm = Convert.ToDouble(rd["HizmetBedeli"]); sStpj = Convert.ToDouble(rd["Stopaj"]); sPos = Convert.ToDouble(rd["PosBedeli"]); } else { MessageBox.Show("Sipariş veritabanında bulunamadı!"); return; } }
                                        }
                                        if (sDurum != "İHTİLAFLI") { if (MessageBox.Show($"Bu siparişin durumu '{sDurum}', 'İHTİLAFLI' değil! Yine de zorla işleme devam edilsin mi?", "Durum", MessageBoxButtons.YesNo) == DialogResult.No) return; }
                                        double orjAlis = 0; using (var c = new SqliteCommand("SELECT AlisFiyati FROM Urunler WHERE Barkod=@b OR UrunAdi=@u LIMIT 1", conn)) { c.Parameters.AddWithValue("@b", sBarkod); c.Parameters.AddWithValue("@u", sUrunAdi); var obj = c.ExecuteScalar(); if (obj != null && obj != DBNull.Value) orjAlis = Convert.ToDouble(obj); }
                                        double kesintilerToplami = sKoms + sKrg + sHzm + sStpj + sPos; double yeniAlis = orjAlis - tut + kesintilerToplami; if (yeniAlis < 0) yeniAlis = 0;
                                        using (var c = new SqliteCommand("UPDATE Siparisler SET Durum='İADE', GuncellenmeTarihi=@dt WHERE SiparisNo=@sn", conn)) { c.Parameters.AddWithValue("@dt", DateTime.Now.ToString("dd.MM.yyyy")); c.Parameters.AddWithValue("@sn", targetSn); c.ExecuteNonQuery(); }
                                        string yBarkod = sBarkod + "-2EL"; string yUrunAdi = sUrunAdi + " (2. El)"; bool varMi = false; using (var c = new SqliteCommand("SELECT COUNT(*) FROM Urunler WHERE Barkod=@b", conn)) { c.Parameters.AddWithValue("@b", yBarkod); varMi = Convert.ToInt32(c.ExecuteScalar()) > 0; }
                                        if (varMi) { using (var c = new SqliteCommand("UPDATE Urunler SET StokMiktari = StokMiktari + 1, AlisFiyati=@af WHERE Barkod=@b", conn)) { c.Parameters.AddWithValue("@af", yeniAlis); c.Parameters.AddWithValue("@b", yBarkod); c.ExecuteNonQuery(); } }
                                        else { using (var c = new SqliteCommand("INSERT INTO Urunler (Barkod, UrunAdi, AlisFiyati, SatisFiyati, StokMiktari, Kategori, Marka, Aciklama) VALUES (@b, @u, @af, @sf, 1, '2. EL', 'Tasnif Dışı', 'Hurda/İade')", conn)) { c.Parameters.AddWithValue("@b", yBarkod); c.Parameters.AddWithValue("@u", yUrunAdi); c.Parameters.AddWithValue("@af", yeniAlis); c.Parameters.AddWithValue("@sf", yeniAlis * 1.5); c.ExecuteNonQuery(); } }
                                        if (dgv.Columns.Contains("Dedektif (Durum)")) r.Cells["Dedektif (Durum)"].Value = "HURDA İŞLENDİ";
                                        dgv.Refresh(); MessageBox.Show("İhtilaflı sipariş İADE olarak kapatıldı, ürün 2. El stoğuna eklendi!", "Başarılı");
                                    }
                                }
                                catch (Exception ex) { MessageBox.Show(ex.Message); }
                            }
                        }; cms.Items.Add(btnHurda);

                        cms.Show(Cursor.Position);
                    }
                }
            };
            Panel pBtn = new Panel { Dock = DockStyle.Fill, Height = 50 };
            Button btnC = new Button { Text = "Ödemeleri Onayla ve Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnC.FlatAppearance.BorderSize = 0; pBtn.SizeChanged += (sB, eB) => { btnC.Location = new Point((pBtn.Width - btnC.Width) / 2, (pBtn.Height - btnC.Height) / 2); };
            btnC.Click += (sB, eB) => {
                bool hataliVar = false; if (dgv.Columns.Contains("Dedektif (Durum)")) { foreach (DataGridViewRow rx in dgv.Rows) { if (rx.Cells["Dedektif (Durum)"].Value.ToString().StartsWith("HATA:")) { hataliVar = true; break; } } }
                if (hataliVar) { MessageBox.Show("Hatalı (Kırmızı) kayıtlar var! Lütfen sağ tıklayıp hataları giderin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                f.DialogResult = DialogResult.OK; f.Close();
            }; pBtn.Controls.Add(btnC);
            pnlGridContainer.Controls.Add(dgv); tlpPopupLayout.Controls.Add(pnlGridContainer, 0, 1); tlpPopupLayout.Controls.Add(pBtn, 0, 2);
            dgv.BringToFront(); f.ShowDialog();
        }

        private void SiparisiAmeliyatMasasinaYatir(string sipNo, double exH, double exK, DataGridView anaGrid)
        {
            long rId = 0; double sF = 0, km = 0, kr = 0, hz = 0, st = 0, ps = 0, cz = 0; string durum = "YOK", brkd = "-", uAdi = "Bilinmeyen Ürün", kat = "-", mrk = "-", tar = "-"; int adet = 1; double alisFiyati = 0; string imei = "-";
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open(); string q = @"SELECT s.rowid, s.SatisFiyati, s.KomisyonTL, s.KargoUcreti, s.HizmetBedeli, s.Stopaj, s.PosBedeli, IFNULL(s.CezaTutari,0), s.Durum, s.Barkod, IFNULL(u.UrunAdi, s.UrunAdi), IFNULL(u.Kategori, s.Kategori), IFNULL(u.Marka, s.Marka), IFNULL(s.Adet,1), s.Tarih, IFNULL(s.AlisFiyati,0), IFNULL(s.SatilanImei,'-') FROM Siparisler s LEFT JOIN Urunler u ON s.Barkod = u.Barkod WHERE s.SiparisNo = @sn AND UPPER(s.PlatformAd) LIKE '%TRENDYOL%'";
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
            Button btn = new Button { Text = "💾 GÜNCELLE", Location = new Point(15, currentY + 25), Width = 475, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, Font = new Font("Segoe UI", 10F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => {
                try
                {
                    double.TryParse(tSat.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n1); double.TryParse(tKoms.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n2); double.TryParse(tKrg.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n4); double.TryParse(tHzm.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n5); double.TryParse(tStpj.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n6); double.TryParse(tPos.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n7); double.TryParse(tCez.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double n3); double.TryParse(tAlis.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double nAlis); string nImei = tImei.Text.Trim();
                    using (var b = new SqliteConnection(Veritabanı.BaglantiYolu)) { b.Open(); using (var cmd = new SqliteCommand("UPDATE Siparisler SET SatisFiyati=@sf, AlisFiyati=@af, SatilanImei=@im, KomisyonTL=@km, CezaTutari=@cz, KargoUcreti=@kr, HizmetBedeli=@hz, Stopaj=@st, PosBedeli=@pb WHERE rowid=@id", b)) { cmd.Parameters.AddWithValue("@sf", n1); cmd.Parameters.AddWithValue("@af", nAlis); cmd.Parameters.AddWithValue("@im", nImei); cmd.Parameters.AddWithValue("@km", n2); cmd.Parameters.AddWithValue("@cz", n3); cmd.Parameters.AddWithValue("@kr", n4); cmd.Parameters.AddWithValue("@hz", n5); cmd.Parameters.AddWithValue("@st", n6); cmd.Parameters.AddWithValue("@pb", n7); cmd.Parameters.AddWithValue("@id", rId); cmd.ExecuteNonQuery(); } }
                    foreach (DataGridViewRow rowX in anaGrid.Rows) { if (rowX.Cells["Sipariş No"].Value.ToString() == sipNo) { if (anaGrid.Columns.Contains("Dedektif (Durum)")) rowX.Cells["Dedektif (Durum)"].Value = "✔ MÜTABIK"; } }
                    anaGrid.Refresh(); System.Windows.Forms.MessageBox.Show("Başarıyla güncellendi!"); frm.Close();
                }
                catch (Exception exAmeliyat) { System.Windows.Forms.MessageBox.Show(exAmeliyat.Message); }
            };
            frm.Controls.Add(btn); frm.ShowDialog();
        }

        private Panel FormElemaniYaratRenk(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan, Color baslikRenk)
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 48 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = baslikRenk };
            kutu = new TextBox { Text = varsayilan, Location = new Point(0, 16), Width = genislik, Font = new Font("Segoe UI", 9.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p;
        }

        private void TahsilatGecmisiniListele()
        {
            try { using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu)) { conn.Open(); using (var cmd = new SqliteCommand("SELECT Id, Tarih, Platform, Tutar as 'Yatan Net Tutar (TL)', SiparisAdedi as 'Kapatılan Sipariş' FROM PlatformTahsilatlari WHERE UPPER(Platform)='TRENDYOL' ORDER BY Id DESC", conn)) { DataTable dt = new DataTable(); dt.Load(cmd.ExecuteReader()); dynDgvTahsilatGecmisi.DataSource = dt; if (dynDgvTahsilatGecmisi.Columns["Id"] != null) dynDgvTahsilatGecmisi.Columns["Id"].Visible = false; dynDgvTahsilatGecmisi.ClearSelection(); } } } catch { }
        }

        private void DynBtnSiparisSorgula_Click(object sender, EventArgs e)
        {
            var sips = dynTxtSiparisListesi.Text.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Distinct().ToList();
            if (sips.Count == 0)
            {
                MessageBox.Show("Lütfen sipariş numarası giriniz.", "Sipariş Numarası Eksik", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (sips.Count > 1000) sips = sips.Take(1000).ToList();

            Form frmPopup = new Form
            {
                Text = $"⚖️ TRENDYOL MANUEL MUTABAKAT VE ÖDEME RADARI (Sorgulanan: {sips.Count} Sipariş)",
                Size = new Size(1150, 600),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.FromArgb(33, 33, 36),
                ForeColor = Color.White,
                ShowIcon = false
            };

            TableLayoutPanel tlpPopupLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.FromArgb(33, 33, 36) };
            tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
            tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlpPopupLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            frmPopup.Controls.Add(tlpPopupLayout);

            Panel pnlTopInfo = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28) };
            pnlTopInfo.Controls.Add(new Label { Text = "📢 'Trendyol Ödenen (TL)' sütununa gelen net hakediş tutarlarını yazıp Enter'a basınız. Sipariş satırına ÇİFT TIKLAYARAK ameliyat masasını açabilirsiniz.", Dock = DockStyle.Fill, ForeColor = Color.Gold, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter });
            tlpPopupLayout.Controls.Add(pnlTopInfo, 0, 0);

            Panel pnlGridContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 10) };

            DataGridView dgvPopup = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, Font = new Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular), EnableHeadersVisualStyles = false, BorderStyle = BorderStyle.None, ColumnHeadersHeight = 35 };
            dgvPopup.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            dgvPopup.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPopup.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
            dgvPopup.DefaultCellStyle.ForeColor = Color.White;

            dgvPopup.Columns.Add("SiparisNo", "Sipariş No");
            dgvPopup.Columns.Add("Durum", "Sistem Durumu");
            dgvPopup.Columns.Add("Urun", "Kayıtlı Ürün");
            dgvPopup.Columns.Add("Ciro", "Brüt Satış (TL)");
            dgvPopup.Columns.Add("Kesinti", "Kesintiler (TL)");
            dgvPopup.Columns.Add("Kargo", "Kargo (TL)");
            dgvPopup.Columns.Add("Hakedis", "Sistem Hakedişi (TL)");
            dgvPopup.Columns.Add("GelenTutar", "Trendyol Ödenen (TL)");
            dgvPopup.Columns.Add("Mutabakat", "Mutabakat Kontrolü");

            for (int i = 0; i < dgvPopup.Columns.Count; i++) if (i != 7) dgvPopup.Columns[i].ReadOnly = true;
            dgvPopup.Columns[3].DefaultCellStyle.Format = "N2";
            dgvPopup.Columns[4].DefaultCellStyle.Format = "N2";
            dgvPopup.Columns[5].DefaultCellStyle.Format = "N2";
            dgvPopup.Columns[6].DefaultCellStyle.Format = "N2";
            dgvPopup.Columns[7].DefaultCellStyle.Format = "N2";
            dgvPopup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvPopup.CellDoubleClick += (sGrid, eGrid) =>
            {
                if (eGrid.RowIndex < 0) return;
                string selectedSipNo = dgvPopup.Rows[eGrid.RowIndex].Cells["SiparisNo"].Value?.ToString();
                if (!string.IsNullOrEmpty(selectedSipNo) && selectedSipNo != "-")
                {
                    SiparisiAmeliyatMasasinaYatir(selectedSipNo, 0, 0, dgvPopup);
                }
            };

            pnlGridContainer.Controls.Add(dgvPopup);
            tlpPopupLayout.Controls.Add(pnlGridContainer, 0, 1);

            Panel pnlPopupBtn = new Panel { Dock = DockStyle.Fill };
            Button btnPopupClose = new Button { Text = "Mutabakat Kontrolünü Onayla ve Kapat", Width = 280, Height = 35, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnPopupClose.FlatAppearance.BorderSize = 0;
            pnlPopupBtn.SizeChanged += (sB, eB) => { btnPopupClose.Location = new Point((pnlPopupBtn.Width - btnPopupClose.Width) / 2, (pnlPopupBtn.Height - btnPopupClose.Height) / 2); };
            btnPopupClose.Click += (sB, eB) => { frmPopup.DialogResult = DialogResult.OK; frmPopup.Close(); };
            pnlPopupBtn.Controls.Add(btnPopupClose);
            tlpPopupLayout.Controls.Add(pnlPopupBtn, 0, 2);

            using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
            {
                conn.Open();
                foreach (string sNo in sips)
                {
                    string query = @"SELECT UrunAdi, IFNULL(s.SatisFiyati,0) as Satis, IFNULL(s.PlatformIndirimi,0) as Indirim,
                                     IFNULL(s.Stopaj,0) as Stopaj, IFNULL(s.PosBedeli,0) as Pos, IFNULL(s.HizmetBedeli,0) as Hizmet,
                                     IFNULL(s.KomisyonTL,0) as Kom, IFNULL(s.KargoUcreti,0) as Kargo, IFNULL(s.CezaTutari,0) as Ceza, s.Durum
                                     FROM Siparisler s 
                                     WHERE s.SiparisNo = @sn AND UPPER(s.PlatformAd) LIKE '%TRENDYOL%' AND UPPER(s.Durum) != 'İPTAL'";

                    using (var cmd = new SqliteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@sn", sNo);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                string urun = r["UrunAdi"].ToString();
                                double satis = Convert.ToDouble(r["Satis"]);
                                double indirim = Convert.ToDouble(r["Indirim"]);
                                double stopaj = Convert.ToDouble(r["Stopaj"]);
                                double pos = Convert.ToDouble(r["Pos"]);
                                double hizmet = Convert.ToDouble(r["Hizmet"]);
                                double kom = Convert.ToDouble(r["Kom"]);
                                double kargo = Convert.ToDouble(r["Kargo"]);
                                double ceza = Convert.ToDouble(r["Ceza"]);
                                string sysDurum = r["Durum"].ToString();

                                double kesintiler = stopaj + pos + hizmet + kom + ceza;
                                double sHakedis = satis + indirim - kesintiler - kargo;

                                if (sysDurum.ToUpper() == "TAHSİL EDİLDİ")
                                {
                                    int rowIdx = dgvPopup.Rows.Add(sNo, "❌ MÜKERRER TAHSİLAT", urun, satis, kesintiler, kargo, sHakedis, 0.00, "❌ DAHA ÖNCE ÖDENDİ!");
                                    dgvPopup.Rows[rowIdx].DefaultCellStyle.BackColor = Color.FromArgb(75, 20, 20);
                                    dgvPopup.Rows[rowIdx].DefaultCellStyle.ForeColor = Color.LightCoral;
                                }
                                else
                                {
                                    int rowIdx = dgvPopup.Rows.Add(sNo, $"✔ {sysDurum}", urun, satis, kesintiler, kargo, sHakedis, 0.00, "✍ Tutar Bekleniyor...");
                                    dgvPopup.Rows[rowIdx].DefaultCellStyle.BackColor = Color.FromArgb(38, 45, 38);
                                }
                            }
                            else
                            {
                                int rowIdx = dgvPopup.Rows.Add(sNo, "❌ BULUNAMADI", "Sistemde bu sipariş numarası yok!", 0.00, 0.00, 0.00, 0.00, 0.00, "Bulunamadı");
                                dgvPopup.Rows[rowIdx].DefaultCellStyle.BackColor = Color.FromArgb(55, 25, 25);
                                dgvPopup.Rows[rowIdx].DefaultCellStyle.ForeColor = Color.LightCoral;
                            }
                        }
                    }
                }
            }

            dgvPopup.CellValueChanged += (s, ev) =>
            {
                if (ev.RowIndex < 0 || ev.ColumnIndex != 7) return;
                DataGridViewRow row = dgvPopup.Rows[ev.RowIndex];
                if (row.Cells["Durum"].Value.ToString().Contains("BULUNAMADI") || row.Cells["Durum"].Value.ToString().Contains("MÜKERRER")) { row.Cells["GelenTutar"].Value = 0.00; return; }

                double sHakedis = Convert.ToDouble(row.Cells["Hakedis"].Value);
                double gelenTutar = 0;
                if (row.Cells["GelenTutar"].Value != null) double.TryParse(row.Cells["GelenTutar"].Value.ToString(), out gelenTutar);

                if (Math.Abs(sHakedis - gelenTutar) < 0.001)
                {
                    row.Cells["Mutabakat"].Value = "✔ TAM UYUŞUYOR";
                    row.Cells["Mutabakat"].Style.ForeColor = Color.LimeGreen;
                    row.Cells["Mutabakat"].Style.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
                    row.DefaultCellStyle.BackColor = Color.FromArgb(20, 50, 20);
                }
                else
                {
                    row.Cells["Mutabakat"].Value = "❌ UYUŞMAZLIK HATASI!";
                    row.Cells["Mutabakat"].Style.ForeColor = Color.Crimson;
                    row.Cells["Mutabakat"].Style.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
                    row.DefaultCellStyle.BackColor = Color.FromArgb(65, 20, 20);
                }
            };

            frmPopup.ShowDialog();
        }

        private void CanliBilançoListele()
        {
            try
            {
                DataTable dt = new DataTable(); dt.Columns.Add("Kapatılan Sipariş (Adet)", typeof(int)); dt.Columns.Add("Net Tahsil Edilen TL", typeof(double)); dt.Columns.Add("Toplam Kesintiler TL", typeof(double)); dt.Columns.Add("Kalan Net Alacak TL", typeof(double));

                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sqlBilanço = @"
                        SELECT 
                            SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END),
                            SUM(CASE WHEN s.Durum='TAHSİL EDİLDİ' THEN (IFNULL(s.SatisFiyati,0) + IFNULL(s.PlatformIndirimi,0) - (IFNULL(s.Stopaj,0) + IFNULL(s.PosBedeli,0) + IFNULL(s.HizmetBedeli,0) + IFNULL(s.KomisyonTL,0) + IFNULL(s.CezaTutari,0)) - IFNULL(s.KargoUcreti,0)) ELSE 0 END),
                            SUM(CASE WHEN s.Durum!='TAHSİL EDİLDİ' THEN s.Adet ELSE 0 END),
                            SUM(IFNULL(s.KomisyonTL,0)),
                            SUM(IFNULL(s.HizmetBedeli,0)),
                            SUM(IFNULL(s.CezaTutari,0)),
                            SUM(IFNULL(s.SatisFiyati,0)),
                            SUM(IFNULL(s.PlatformIndirimi,0)),
                            SUM(IFNULL(s.Stopaj,0)),
                            SUM(IFNULL(s.PosBedeli,0)),
                            SUM(IFNULL(s.KargoUcreti,0))
                        FROM Siparisler s 
                        WHERE UPPER(s.PlatformAd) LIKE '%TRENDYOL%' AND UPPER(s.Durum) != 'İPTAL'";

                    string sqlKar = @"
                        SELECT 
                            IFNULL(SUM(s.Adet), 0),
                            IFNULL(SUM((s.SatisFiyati * s.Adet) + IFNULL(s.PlatformIndirimi,0) - (IFNULL(s.AlisFiyati,0) * s.Adet) - s.KomisyonTL - s.KargoUcreti - s.HizmetBedeli - s.Stopaj - s.PosBedeli - IFNULL(s.CezaTutari,0)), 0) as NetKar
                        FROM Siparisler s 
                        WHERE UPPER(s.PlatformAd) LIKE '%TRENDYOL%' AND UPPER(s.Durum) != 'İPTAL'";

                    double netKarTutar = 0; int karSiparisAdet = 0;
                    using (var cmdKar = new SqliteCommand(sqlKar, conn)) using (var rKar = cmdKar.ExecuteReader()) { if (rKar.Read()) { karSiparisAdet = Convert.ToInt32(rKar[0]); netKarTutar = Convert.ToDouble(rKar[1]); } }

                    using (var cmd = new SqliteCommand(sqlBilanço, conn))
                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read())
                        {
                            int odenenAdet = r[0] != DBNull.Value ? Convert.ToInt32(r[0]) : 0; double odenenPara = r[1] != DBNull.Value ? Convert.ToDouble(r[1]) : 0; int bekleyenAdet = r[2] != DBNull.Value ? Convert.ToInt32(r[2]) : 0;
                            double totKomisyon = r[3] != DBNull.Value ? Convert.ToDouble(r[3]) : 0; double totHizmet = r[4] != DBNull.Value ? Convert.ToDouble(r[4]) : 0; double totCeza = r[5] != DBNull.Value ? Convert.ToDouble(r[5]) : 0;
                            double totSatis = r[6] != DBNull.Value ? Convert.ToDouble(r[6]) : 0; double kesintiler = totKomisyon + totHizmet + totCeza;
                            double totIndirim = r[7] != DBNull.Value ? Convert.ToDouble(r[7]) : 0; double totStopaj = r[8] != DBNull.Value ? Convert.ToDouble(r[8]) : 0; double totPos = r[9] != DBNull.Value ? Convert.ToDouble(r[9]) : 0; double totKargo = r[10] != DBNull.Value ? Convert.ToDouble(r[10]) : 0;

                            double toplamNetHakedisMatrisi = totSatis + totIndirim - (totStopaj + totPos + totKomisyon + totHizmet + totCeza + totKargo);
                            double bakiyeKalanNetAlacak = toplamNetHakedisMatrisi - odenenPara; double toplamBrutCiro = totSatis;

                            dt.Rows.Add(odenenAdet, odenenPara, kesintiler, bakiyeKalanNetAlacak);

                            lblCardTahsilEdilenAdet.Text = $"{odenenAdet} Gelen Ödemeler"; lblCardTahsilEdilenPara.Text = odenenPara.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardBekleyenAdet.Text = $"{odenenAdet + bekleyenAdet} İşlem Kayıtlı"; lblCardBekleyenPara.Text = kesintiler.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardToplamAdet.Text = $"{bekleyenAdet} Sipariş Beklemede"; lblCardToplamPara.Text = bakiyeKalanNetAlacak.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardCiroAdet.Text = "Mağaza Toplam Hacmi"; lblCardCiroPara.Text = toplamBrutCiro.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                            lblCardKarAdet.Text = $"{karSiparisAdet} Siparişin Net Kârı"; lblCardKarPara.Text = netKarTutar.ToString("N2", CultureInfo.GetCultureInfo("tr-TR")) + " ₺";
                        }
                    }
                }
                dgvPlatformBilanço.DataSource = dt;
                if (dgvPlatformBilanço.Columns.Contains("Net Tahsil Edilen TL")) dgvPlatformBilanço.Columns["Net Tahsil Edilen TL"].DefaultCellStyle.Format = "N2";
                if (dgvPlatformBilanço.Columns.Contains("Toplam Kesintiler TL")) dgvPlatformBilanço.Columns["Toplam Kesintiler TL"].DefaultCellStyle.Format = "N2";
                if (dgvPlatformBilanço.Columns.Contains("Kalan Net Alacak TL")) dgvPlatformBilanço.Columns["Kalan Net Alacak TL"].DefaultCellStyle.Format = "N2";
            }
            catch { }
        }

        private DataGridView GridYarat()
        {
            DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9F), BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false, ColumnHeadersHeight = 32 };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold); dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36); dgv.DefaultCellStyle.ForeColor = Color.White; dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215); dgv.DataError += (s, e) => { e.Cancel = true; };
            dgv.DataBindingComplete += (s, e) => { dgv.ClearSelection(); };
            return dgv;
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan = "")
        {
            Panel p = new Panel { Location = konum, Width = genislik, Height = 52 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p;
        }

        private void CustomVIPCardInsaEt(TableLayoutPanel tlp, string b, Color c, int s, out Label lA, out Label lP)
        {
            Panel pCard = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(22, 22, 26), Margin = new Padding(6), Padding = new Padding(1) };
            Panel pAccent = new Panel { Dock = DockStyle.Top, Height = 3, BackColor = c }; pCard.Controls.Add(pAccent);
            Panel pTextContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 6, 6, 6), BackColor = Color.FromArgb(26, 26, 30) };
            Label lblB = new Label { Text = b.ToUpper(CultureInfo.CurrentCulture), Dock = DockStyle.Top, Height = 20, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.FromArgb(160, 160, 170), TextAlign = ContentAlignment.MiddleLeft };
            lP = new Label { Text = "0,00 ₺", Dock = DockStyle.Top, Height = 26, Font = new Font("Segoe UI Black", 13.5F, FontStyle.Bold), ForeColor = c, TextAlign = ContentAlignment.MiddleLeft };
            lA = new Label { Text = "0 Sipariş Kayıtlı", Dock = DockStyle.Top, Height = 18, Font = new Font("Segoe UI", 8.0F, FontStyle.Regular), ForeColor = Color.FromArgb(110, 110, 120), TextAlign = ContentAlignment.MiddleLeft };
            pTextContainer.Controls.Add(lA); pTextContainer.Controls.Add(lP); pTextContainer.Controls.Add(lblB); lblB.BringToFront(); lP.BringToFront(); lA.BringToFront(); pCard.Controls.Add(pTextContainer); tlp.Controls.Add(pCard, s, 0);
        }
    }
}