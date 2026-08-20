using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Xml.Linq;

namespace StokTakipProgrami
{
    // ⚡ MANUEL VE EXCEL ŞABLONLU TOPLU VERİ GİRİŞ MERKEZİ
    // DÜZELTMELER:
    // 1. OleDb kaldırıldı → ClosedXML ile .xlsx okuma eklendi (kurulum gerektirmez, NuGet paketi)
    // 2. Mükerrer sipariş kontrolü: aynı sipariş no varsa atla, hata verme
    // 3. Platformlar tablosu sütun adı "Ad" olarak sabitlendi (dinamik arama kaldırıldı)
    // 4. Tarih ayrıştırma Excel OADate (sayısal tarih) desteği eklendi
    // 5. SafeDouble virgül/nokta karışıklığı tam düzeltildi
    // 6. StokHareketleri tablosu altyapı güvencesi eklendi
    // 7. Stok adedi Excel'den yüklenince azaltılmıyor (geçmiş aktarım, stok etkilenmez)

    public partial class SayfaVeriler : UserControl
    {
        private DataGridView dynDgvVeriGiris;
        private ComboBox dynCmbHedefTablo;
        private Button dynBtnSatirEkle, dynBtnVerileriMuhurle, dynBtnExcelYukle, dynBtnSablonIndir, dynBtnSiparisSorgula;
        private DataTable veriTablosu;

        public SayfaVeriler()
        {
            InitializeComponent();
            this.Load += SayfaVeriler_Load;
        }

        private void SayfaVeriler_Load(object sender, EventArgs e)
        {
            ArayuzInsaEt();
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear();
            this.BackColor = Color.FromArgb(33, 33, 36);

            Label lblSecim = new Label { Text = "Giriş Alanı:", Location = new Point(15, 15), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.White };
            dynCmbHedefTablo = new ComboBox { Location = new Point(90, 12), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 9.5F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White };
            dynCmbHedefTablo.Items.AddRange(new object[] { "📦 STOKLAR (Yeni Verileri Toplu Tanıt)", "🛒 SİPARİŞLER (Geçmiş Satışları Aktar)" });
            dynCmbHedefTablo.SelectedIndexChanged += DynCmbHedefTablo_SelectedIndexChanged;
            this.Controls.Add(lblSecim);
            this.Controls.Add(dynCmbHedefTablo);

            dynBtnSablonIndir = new Button { Text = "📥 Şablon İndir", Location = new Point(280, 11), Width = 100, Height = 26, BackColor = Color.FromArgb(142, 68, 173), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSablonIndir.FlatAppearance.BorderSize = 0;
            dynBtnSablonIndir.Click += DynBtnSablonIndir_Click;
            this.Controls.Add(dynBtnSablonIndir);

            dynBtnSatirEkle = new Button { Text = "➕ Satır Ekle", Location = new Point(390, 11), Width = 90, Height = 26, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSatirEkle.FlatAppearance.BorderSize = 0;
            dynBtnSatirEkle.Click += (s, e) => { if (veriTablosu != null) veriTablosu.Rows.Add(veriTablosu.NewRow()); };
            this.Controls.Add(dynBtnSatirEkle);

            dynBtnExcelYukle = new Button { Text = "📊 Excel / CSV Yükle", Location = new Point(490, 11), Width = 120, Height = 26, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnExcelYukle.FlatAppearance.BorderSize = 0;
            dynBtnExcelYukle.Click += DynBtnExcelYukle_Click;
            this.Controls.Add(dynBtnExcelYukle);

            dynBtnVerileriMuhurle = new Button { Text = "💾 VERİLERİ SİSTEME KAYDET", Location = new Point(620, 11), Width = 355, Height = 26, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnVerileriMuhurle.FlatAppearance.BorderSize = 0;
            dynBtnVerileriMuhurle.Click += DynBtnVerileriMuhurle_Click;
            this.Controls.Add(dynBtnVerileriMuhurle);

            dynBtnSiparisSorgula = new Button { Text = "🔍 TABLODAKİ SİPARİŞLERİ DETAYLI ÖNİZLE / SORGULA", Location = new Point(15, 590), Width = 350, Height = 28, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            dynBtnSiparisSorgula.FlatAppearance.BorderSize = 0;
            dynBtnSiparisSorgula.Click += DynBtnSiparisSorgula_Click;
            this.Controls.Add(dynBtnSiparisSorgula);

            dynDgvVeriGiris = new DataGridView
            {
                Location = new Point(15, 48),
                Width = 960,
                Height = 532,
                AllowUserToAddRows = true,
                BackgroundColor = Color.FromArgb(26, 26, 28),
                RowHeadersVisible = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9.5F),
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, SelectionBackColor = Color.FromArgb(0, 120, 215) },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(21, 21, 33), ForeColor = Color.White, Font = new Font("Segoe UI", 9F, FontStyle.Bold) },
                EnableHeadersVisualStyles = false,
                EditMode = DataGridViewEditMode.EditOnEnter
            };
            this.Controls.Add(dynDgvVeriGiris);
        }

        // --- SİPARİŞ NO ÇÖZÜMLEME ---
        private string SiparisNoCozumle(object val)
        {
            if (val == null || string.IsNullOrWhiteSpace(val.ToString())) return "";

            if (val is double d) return ((long)d).ToString();
            if (val is float f) return ((long)f).ToString();
            if (val is decimal dec) return ((long)dec).ToString();

            string s = val.ToString().Trim();
            // "219385251.0" gibi ondalıklı gelenleri temizle
            if (s.Contains(".") && s.EndsWith(".0")) s = s.Substring(0, s.Length - 2);
            if (s.Contains(",") && s.EndsWith(",0")) s = s.Substring(0, s.Length - 2);
            // Bilimsel gösterim: "2.19385E+08" gibi
            if (s.Contains("E") || s.Contains("e"))
            {
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
                    return ((long)parsed).ToString();
            }
            return s;
        }

        // --- TARİH ÇÖZÜMLEME (Excel OADate dahil) ---
        private string FormatliTarihGetir(object val)
        {
            if (val == null || string.IsNullOrWhiteSpace(val.ToString()))
                return DateTime.Now.ToString("dd.MM.yyyy");

            string s = val.ToString().Trim();

            // Excel OADate: sayısal tarih (örn: 45678)
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate) && oaDate > 1000 && oaDate < 100000)
            {
                try
                {
                    DateTime dt = DateTime.FromOADate(oaDate);
                    return dt.ToString("dd.MM.yyyy");
                }
                catch { }
            }

            // Standart format denemeleri
            string[] formatlar = {
                "dd.MM.yyyy", "d.M.yyyy", "dd/MM/yyyy", "d/M/yyyy",
                "yyyy-MM-dd", "yyyy.MM.dd", "MM/dd/yyyy",
                "dd-MM-yyyy", "d-M-yyyy"
            };

            s = s.Replace("/", ".").Replace("-", ".");
            string[] parts = s.Split('.');
            if (parts.Length == 3)
            {
                try
                {
                    if (parts[0].Length == 4) // yyyy.MM.dd
                    {
                        int yil = int.Parse(parts[0]);
                        int ay = int.Parse(parts[1]);
                        int gun = int.Parse(parts[2]);
                        return $"{gun:D2}.{ay:D2}.{yil}";
                    }
                    else
                    {
                        int gun = int.Parse(parts[0]);
                        int ay = int.Parse(parts[1]);
                        int yil = int.Parse(parts[2]);
                        if (yil < 100) yil += 2000;
                        return $"{gun:D2}.{ay:D2}.{yil}";
                    }
                }
                catch { }
            }

            return DateTime.Now.ToString("dd.MM.yyyy");
        }

        // --- GÜVENLİ DOUBLE DÖNÜŞÜM ---
        private double SafeDouble(object val)
        {
            if (val == null || string.IsNullOrWhiteSpace(val.ToString())) return 0;
            try
            {
                string s = val.ToString().Trim()
                    .Replace("TL", "").Replace("₺", "").Replace(" ", "").Replace("\u00A0", "");

                // Hem nokta hem virgül varsa: Türkçe format (1.500,50) → virgül ondalık
                if (s.Contains(",") && s.Contains("."))
                    s = s.Replace(".", "").Replace(",", ".");
                else
                    s = s.Replace(",", ".");

                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double res))
                    return res;
            }
            catch { }
            return 0;
        }

        private int SafeInt(object val)
        {
            if (val == null || string.IsNullOrWhiteSpace(val.ToString())) return 1;
            string s = val.ToString().Trim();
            // "1.0" gibi double formatında gelebilir
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                return Math.Max(1, (int)d);
            if (int.TryParse(s, out int res)) return Math.Max(1, res);
            return 1;
        }

        private void DynCmbHedefTablo_SelectedIndexChanged(object sender, EventArgs e)
        {
            veriTablosu = new DataTable();

            if (dynCmbHedefTablo.SelectedIndex == 0)
            {
                veriTablosu.Columns.Add("Barkod");
                veriTablosu.Columns.Add("Kategori");
                veriTablosu.Columns.Add("Marka");
                veriTablosu.Columns.Add("Ürün Adı");
            }
            else
            {
                veriTablosu.Columns.Add("Tarih");
                veriTablosu.Columns.Add("Kategori");
                veriTablosu.Columns.Add("Barkod");
                veriTablosu.Columns.Add("Marka");
                veriTablosu.Columns.Add("Satılan IMEI");
                veriTablosu.Columns.Add("Durum");
                veriTablosu.Columns.Add("Ürün Adı");
                veriTablosu.Columns.Add("Adet");
                veriTablosu.Columns.Add("Sipariş No");
                veriTablosu.Columns.Add("Alış Fiyatı");
                veriTablosu.Columns.Add("Platform İndirimi");
                veriTablosu.Columns.Add("Satış Fiyatı");
                veriTablosu.Columns.Add("KDV Hariç");
                veriTablosu.Columns.Add("Platform");
                veriTablosu.Columns.Add("Kargo Firması");
                veriTablosu.Columns.Add("Desi");
                veriTablosu.Columns.Add("Kargo Ücreti");
                veriTablosu.Columns.Add("Komisyon");
                veriTablosu.Columns.Add("Hizmet Bedeli");
                veriTablosu.Columns.Add("Stopaj");
                veriTablosu.Columns.Add("POS Bedeli");
            }

            dynDgvVeriGiris.DataSource = veriTablosu;
            for (int i = 0; i < 5; i++) veriTablosu.Rows.Add(veriTablosu.NewRow());
        }

        private void DynBtnSablonIndir_Click(object sender, EventArgs e)
        {
            if (dynCmbHedefTablo.SelectedIndex < 0)
            {
                MessageBox.Show("Lütfen giriş alanı seçiniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Uyumlu CSV Dosyası (*.csv)|*.csv";
                string baslikSatiri;

                if (dynCmbHedefTablo.SelectedIndex == 0)
                {
                    sfd.FileName = "Toplu_Stok_Veri_Ekleme_Sablonu.csv";
                    baslikSatiri = "Barkod;Kategori;Marka;Ürün Adı";
                }
                else
                {
                    sfd.FileName = "Toplu_Siparis_Finans_Matris_Sablonu.csv";
                    baslikSatiri = "Tarih;Kategori;Barkod;Marka;IMEI;Durum (Örn: ÖDENDİ);Ürün Adı;Adet;Sipariş No;Alış Fiyatı;Platform İndirimi;Satış Fiyatı;KDV Hariç;Platform;Kargo Firması;Desi;Kargo Bedeli;Komisyon;Hizmet Bedeli;Stopaj;POS Bedeli";
                }

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        byte[] bom = new byte[] { 0xEF, 0xBB, 0xBF };
                        using (var fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                        {
                            fs.Write(bom, 0, bom.Length);
                            using (var sw = new StreamWriter(fs, System.Text.Encoding.UTF8))
                            {
                                sw.WriteLine(baslikSatiri);
                                // Örnek satır ekle
                                if (dynCmbHedefTablo.SelectedIndex == 1)
                                    sw.WriteLine("25.06.2025;CEP TELEFONU;APPLE;-;-;ÖDENDİ;IPHONE 15 PRO 256GB;1;TRD123456789;18000,00;0,00;25000,00;21008,40;TRENDYOL;YURTİÇİ;1;0,00;1250,00;0,00;208,40;0,00");
                            }
                        }
                        MessageBox.Show("Şablon başarıyla indirildi!\n\nDosyayı Excel ile açıp doldurun, ardından buradan yükleyin.", "Şablon İndirildi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Şablon indirilirken hata oluştu:\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // --- CSV OKUMA MOTORU ---
        private DataTable CsvDosyasiniOku(string dosyaYolu)
        {
            DataTable dt = new DataTable();

            // Encoding tespiti (BOM varsa UTF-8, yoksa default)
            System.Text.Encoding enc = System.Text.Encoding.Default;
            byte[] bom = new byte[3];
            using (var fs = new FileStream(dosyaYolu, FileMode.Open, FileAccess.Read))
            {
                fs.Read(bom, 0, 3);
                if (bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) enc = System.Text.Encoding.UTF8;
            }

            string icerik = File.ReadAllText(dosyaYolu, enc);

            // Ayırıcı otomatik tespit: noktalı virgül mü, virgül mü, tab mı?
            char ayirici = ';';
            string[] ilkSatirlar = icerik.Split('\n');
            if (ilkSatirlar.Length > 0)
            {
                string ilk = ilkSatirlar[0];
                int noktaliVirgulSayisi = ilk.Count(c => c == ';');
                int virgulSayisi = ilk.Count(c => c == ',');
                int tabSayisi = ilk.Count(c => c == '\t');
                if (tabSayisi > noktaliVirgulSayisi && tabSayisi > virgulSayisi) ayirici = '\t';
                else if (virgulSayisi > noktaliVirgulSayisi * 2) ayirici = ',';
            }

            bool tirnakIci = false;
            var satir = new List<string>();
            var hucre = new System.Text.StringBuilder();
            bool ilkSatir = true;

            for (int i = 0; i < icerik.Length; i++)
            {
                char c = icerik[i];

                if (c == '"')
                {
                    if (!tirnakIci) tirnakIci = true;
                    else if (i + 1 < icerik.Length && icerik[i + 1] == '"') { hucre.Append('"'); i++; }
                    else tirnakIci = false;
                }
                else if (c == ayirici && !tirnakIci)
                {
                    satir.Add(hucre.ToString().Trim());
                    hucre.Clear();
                }
                else if ((c == '\r' || c == '\n') && !tirnakIci)
                {
                    if (c == '\r' && i + 1 < icerik.Length && icerik[i + 1] == '\n') i++;
                    satir.Add(hucre.ToString().Trim());
                    hucre.Clear();

                    if (ilkSatir && satir.Any(x => !string.IsNullOrWhiteSpace(x)))
                    {
                        // İlk satır başlık: sütun ekle
                        for (int col = 0; col < satir.Count; col++)
                            dt.Columns.Add("Col" + col);
                        ilkSatir = false;
                        satir.Clear();
                        continue;
                    }

                    if (!ilkSatir && satir.Any(x => !string.IsNullOrWhiteSpace(x)))
                    {
                        DataRow dr = dt.NewRow();
                        for (int col = 0; col < Math.Min(satir.Count, dt.Columns.Count); col++)
                            dr[col] = satir[col];
                        dt.Rows.Add(dr);
                    }
                    satir.Clear();
                }
                else
                {
                    hucre.Append(c);
                }
            }

            // Son satır
            if (hucre.Length > 0 || satir.Count > 0)
            {
                satir.Add(hucre.ToString().Trim());
                if (!ilkSatir && satir.Any(x => !string.IsNullOrWhiteSpace(x)) && dt.Columns.Count > 0)
                {
                    DataRow dr = dt.NewRow();
                    for (int col = 0; col < Math.Min(satir.Count, dt.Columns.Count); col++)
                        dr[col] = satir[col];
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        // --- EXCEL (.xlsx) OKUMA MOTORU (Saf C# - hiçbir ek paket gerektirmez) ---
        // .xlsx dosyası aslında bir ZIP arşividir. İçindeki XML dosyalarını okuyarak veriyi çekeriz.
        private DataTable ExcelDosyasiniOku(string dosyaYolu)
        {
            DataTable dt = new DataTable();

            using (var zip = System.IO.Compression.ZipFile.OpenRead(dosyaYolu))
            {
                // 1. Paylaşılan string tablosunu oku (sharedStrings.xml)
                var stringler = new List<string>();
                var sharedEntry = zip.GetEntry("xl/sharedStrings.xml");
                if (sharedEntry != null)
                {
                    using (var stream = sharedEntry.Open())
                    {
                        var xdoc = System.Xml.Linq.XDocument.Load(stream);
                        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        foreach (var si in xdoc.Descendants(ns + "si"))
                        {
                            string metin = string.Concat(si.Descendants(ns + "t").Select(t => t.Value));
                            stringler.Add(metin);
                        }
                    }
                }

                // 2. Stil bilgisini oku (tarih formatlarını tespit için)
                var tarihFormatlari = new HashSet<int>();
                var stylesEntry = zip.GetEntry("xl/styles.xml");
                if (stylesEntry != null)
                {
                    using (var stream = stylesEntry.Open())
                    {
                        var xdoc = System.Xml.Linq.XDocument.Load(stream);
                        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        var numFmts = xdoc.Descendants(ns + "numFmt")
                            .Where(n => {
                                string fmt = n.Attribute("formatCode")?.Value?.ToLower() ?? "";
                                return fmt.Contains("y") || fmt.Contains("d") || fmt.Contains("m");
                            })
                            .Select(n => int.Parse(n.Attribute("numFmtId")?.Value ?? "0"))
                            .ToHashSet();
                        // Yerleşik Excel tarih format ID'leri: 14-17, 22
                        foreach (int id in new[] { 14, 15, 16, 17, 22 }) numFmts.Add(id);

                        int xfIdx = 0;
                        foreach (var xf in xdoc.Descendants(ns + "cellXf"))
                        {
                            int numFmtId = int.Parse(xf.Attribute("numFmtId")?.Value ?? "0");
                            if (numFmts.Contains(numFmtId)) tarihFormatlari.Add(xfIdx);
                            xfIdx++;
                        }
                    }
                }

                // 3. İlk sayfayı bul
                string sayfaYolu = "xl/worksheets/sheet1.xml";
                var workbookEntry = zip.GetEntry("xl/workbook.xml");
                if (workbookEntry != null)
                {
                    using (var stream = workbookEntry.Open())
                    {
                        var xdoc = System.Xml.Linq.XDocument.Load(stream);
                        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        var ilkSayfa = xdoc.Descendants(ns + "sheet").FirstOrDefault();
                        if (ilkSayfa != null)
                        {
                            // Sayfa ID'sini relationships'tan bul
                            string rId = ilkSayfa.Attribute("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")?.Value;
                            var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                            if (relsEntry != null && rId != null)
                            {
                                using (var rStream = relsEntry.Open())
                                {
                                    var rDoc = System.Xml.Linq.XDocument.Load(rStream);
                                    var rel = rDoc.Descendants().FirstOrDefault(x => x.Attribute("Id")?.Value == rId);
                                    string target = rel?.Attribute("Target")?.Value;
                                    if (!string.IsNullOrEmpty(target))
                                        sayfaYolu = "xl/" + target.TrimStart('/');
                                }
                            }
                        }
                    }
                }

                // 4. Sayfa verisini oku
                var sheetEntry = zip.GetEntry(sayfaYolu);
                if (sheetEntry == null) throw new Exception("Sayfa bulunamadı: " + sayfaYolu);

                using (var stream = sheetEntry.Open())
                {
                    var xdoc = System.Xml.Linq.XDocument.Load(stream);
                    System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                    int maxSutun = 0;
                    var satirlar = new List<List<string>>();

                    foreach (var row in xdoc.Descendants(ns + "row"))
                    {
                        var satirHucreler = new List<string>();
                        int oncekiSutun = 0;

                        foreach (var c in row.Elements(ns + "c"))
                        {
                            // Hücre adresi (örn: "C5") → sütun indeksi
                            string adres = c.Attribute("r")?.Value ?? "";
                            int sutunIdx = AdresSutunIndexiHesapla(adres) - 1;

                            // Boş sütunları doldur
                            while (oncekiSutun < sutunIdx)
                            {
                                satirHucreler.Add("");
                                oncekiSutun++;
                            }

                            string tip = c.Attribute("t")?.Value ?? "";
                            string v = c.Element(ns + "v")?.Value ?? "";
                            string deger = "";

                            if (tip == "s") // Paylaşılan string
                            {
                                if (int.TryParse(v, out int idx) && idx < stringler.Count)
                                    deger = stringler[idx];
                            }
                            else if (tip == "inlineStr") // Inline string
                            {
                                deger = c.Descendants(ns + "t").FirstOrDefault()?.Value ?? "";
                            }
                            else if (!string.IsNullOrEmpty(v))
                            {
                                // Tarih kontrolü
                                string sIdx = c.Attribute("s")?.Value ?? "";
                                if (int.TryParse(sIdx, out int styleIdx) && tarihFormatlari.Contains(styleIdx))
                                {
                                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate))
                                    {
                                        try { deger = DateTime.FromOADate(oaDate).ToString("dd.MM.yyyy"); }
                                        catch { deger = v; }
                                    }
                                    else deger = v;
                                }
                                else deger = v;
                            }

                            satirHucreler.Add(deger);
                            oncekiSutun = sutunIdx + 1;
                            if (oncekiSutun > maxSutun) maxSutun = oncekiSutun;
                        }

                        satirlar.Add(satirHucreler);
                    }

                    // Sütun ekle
                    for (int i = 0; i < maxSutun; i++) dt.Columns.Add("Col" + i);

                    // Satırları DataTable'a ekle
                    foreach (var satirHucreler in satirlar)
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < Math.Min(satirHucreler.Count, maxSutun); i++)
                            dr[i] = satirHucreler[i];
                        dt.Rows.Add(dr);
                    }
                }
            }

            return dt;
        }

        // "C5" → 3, "AA1" → 27
        private int AdresSutunIndexiHesapla(string adres)
        {
            int sonuc = 0;
            foreach (char ch in adres)
            {
                if (!char.IsLetter(ch)) break;
                sonuc = sonuc * 26 + (char.ToUpper(ch) - 'A' + 1);
            }
            return sonuc;
        }

        private void DynBtnExcelYukle_Click(object sender, EventArgs e)
        {
            if (dynCmbHedefTablo.SelectedIndex < 0)
            {
                MessageBox.Show("Lütfen önce 'Giriş Alanı' seçiniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int seciliMod = dynCmbHedefTablo.SelectedIndex;

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Desteklenen Dosyalar (*.xlsx;*.csv)|*.xlsx;*.csv|Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv";
                ofd.Title = "Yüklenecek Finansal Listeyi Seçin";

                if (ofd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    DataTable excelDt;
                    string uzanti = Path.GetExtension(ofd.FileName).ToLower();

                    if (uzanti == ".csv")
                        excelDt = CsvDosyasiniOku(ofd.FileName);
                    else
                        excelDt = ExcelDosyasiniOku(ofd.FileName);

                    if (excelDt.Rows.Count == 0)
                    {
                        MessageBox.Show("Dosyada okunabilir satır bulunamadı! Dosyanın dolu olduğundan ve doğru formatta olduğundan emin olun.", "Boş Dosya", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Sütun indekslerini tespit et (başlık satırı otomatik algıla)
                    int idxTarih = 0, idxKategori = 1, idxBarkod = 2, idxMarka = 3, idxImei = -1;
                    int idxDurum = 5, idxUrunAdi = 6, idxAdet = 7, idxSipNo = 8;
                    int idxAlis = 9, idxIndirim = 10, idxSatis = 11, idxKdv = 12;
                    int idxPlatform = 13, idxKargoFirma = 14, idxDesi = 15, idxKargoUcret = 16;
                    int idxKomisyon = 17, idxHizmet = 18, idxStopaj = 19, idxPos = 20;

                    // İlk satırda başlık mı var kontrol et
                    DataRow ilkSatir = excelDt.Rows[0];
                    bool baslikSatiriVarMi = false;
                    for (int col = 0; col < excelDt.Columns.Count; col++)
                    {
                        string txt = ilkSatir[col]?.ToString()?.ToLower().Trim() ?? "";
                        if (txt.Contains("tarih") || txt.Contains("barkod") || txt.Contains("sipari") || txt.Contains("platform"))
                        {
                            baslikSatiriVarMi = true;
                            break;
                        }
                    }

                    if (baslikSatiriVarMi)
                    {
                        for (int col = 0; col < excelDt.Columns.Count; col++)
                        {
                            string txt = ilkSatir[col]?.ToString()?.ToLower().Trim() ?? "";
                            if (txt.Contains("tarih")) idxTarih = col;
                            else if (txt.Contains("kategori")) idxKategori = col;
                            else if (txt.Contains("barkod")) idxBarkod = col;
                            else if (txt.Contains("marka")) idxMarka = col;
                            else if (txt.Contains("imei")) idxImei = col;
                            else if (txt.Contains("durum")) idxDurum = col;
                            else if (txt.Contains("ürün") || txt.Contains("urun") || txt.Contains("ü")) idxUrunAdi = col;
                            else if (txt.Contains("adet")) idxAdet = col;
                            else if (txt.Contains("sipariş") || txt.Contains("siparis") || txt.Contains("no")) idxSipNo = col;
                            else if (txt.Contains("alış") || txt.Contains("alis") || txt.Contains("maliyet")) idxAlis = col;
                            else if (txt.Contains("indirim")) idxIndirim = col;
                            else if (txt.Contains("satış") || txt.Contains("satis")) idxSatis = col;
                            else if (txt.Contains("kdv")) idxKdv = col;
                            else if (txt.Contains("platform")) idxPlatform = col;
                            else if (txt.Contains("kargo firm")) idxKargoFirma = col;
                            else if (txt.Contains("desi")) idxDesi = col;
                            else if (txt.Contains("kargo") && (txt.Contains("bedel") || txt.Contains("ücret") || txt.Contains("ucret"))) idxKargoUcret = col;
                            else if (txt.Contains("komisyon")) idxKomisyon = col;
                            else if (txt.Contains("hizmet")) idxHizmet = col;
                            else if (txt.Contains("stopaj")) idxStopaj = col;
                            else if (txt.Contains("pos")) idxPos = col;
                        }
                    }

                    dynDgvVeriGiris.SuspendLayout();
                    veriTablosu.BeginLoadData();
                    veriTablosu.Clear();

                    int aktarilanSatirSayaci = 0;
                    string sonSiparisNo = "";
                    string sonTarih = "";
                    string sonPlatform = "";

                    int baslangicSatiri = (seciliMod == 1 && baslikSatiriVarMi) ? 1 : 0;

                    for (int rowIdx = baslangicSatiri; rowIdx < excelDt.Rows.Count; rowIdx++)
                    {
                        DataRow row = excelDt.Rows[rowIdx];

                        string Hucre(int idx)
                        {
                            if (idx < 0 || idx >= row.ItemArray.Length) return "";
                            return row[idx]?.ToString()?.Trim() ?? "";
                        }

                        if (seciliMod == 0) // STOK MODU
                        {
                            string barkod = Hucre(idxBarkod);
                            string kategori = Hucre(idxKategori);
                            string marka = Hucre(idxMarka);
                            string urunAdi = Hucre(idxUrunAdi);

                            if (string.IsNullOrWhiteSpace(barkod) && string.IsNullOrWhiteSpace(urunAdi)) continue;
                            if (barkod.ToLower() == "barkod") continue; // başlık satırı tekrarı

                            veriTablosu.Rows.Add(
                                barkod,
                                string.IsNullOrWhiteSpace(kategori) ? "GENEL" : kategori.ToUpper(),
                                string.IsNullOrWhiteSpace(marka) ? "-" : marka.ToUpper(),
                                urunAdi.ToUpper()
                            );
                            aktarilanSatirSayaci++;
                        }
                        else // SİPARİŞ MODU
                        {
                            string sipNo = SiparisNoCozumle(row.ItemArray.Length > idxSipNo ? row[idxSipNo] : null);
                            string tarih = Hucre(idxTarih);
                            string urunAdi = Hucre(idxUrunAdi);
                            string barkod = Hucre(idxBarkod);

                            // Başlık satırı tekrarını atla
                            if (sipNo.ToLower().Contains("sipari") || tarih.ToLower().Contains("tarih") || barkod.ToLower().Contains("barkod")) continue;
                            // Tamamen boş satırı atla
                            if (string.IsNullOrWhiteSpace(sipNo) && string.IsNullOrWhiteSpace(urunAdi) && string.IsNullOrWhiteSpace(tarih)) continue;

                            // Boş hücreleri önceki satırdan al (merged cell / birleştirilmiş hücre desteği)
                            if (!string.IsNullOrWhiteSpace(sipNo)) sonSiparisNo = sipNo; else sipNo = sonSiparisNo;
                            if (!string.IsNullOrWhiteSpace(tarih)) sonTarih = tarih; else tarih = sonTarih;
                            string platform = Hucre(idxPlatform);
                            if (!string.IsNullOrWhiteSpace(platform)) sonPlatform = platform.ToUpper(); else platform = sonPlatform;

                            if (string.IsNullOrWhiteSpace(sipNo)) continue;

                            if (string.IsNullOrWhiteSpace(barkod) || barkod == "-")
                                barkod = string.IsNullOrWhiteSpace(urunAdi) ? "BILINMEYEN" : urunAdi.ToUpper();

                            string imei = idxImei >= 0 ? SiparisNoCozumle(row.ItemArray.Length > idxImei ? row[idxImei] : null) : "-";
                            if (string.IsNullOrWhiteSpace(imei)) imei = "-";

                            string durum = Hucre(idxDurum).ToUpper();

                            DataRow yeniSatir = veriTablosu.NewRow();
                            yeniSatir["Tarih"] = FormatliTarihGetir(tarih);
                            yeniSatir["Kategori"] = string.IsNullOrWhiteSpace(Hucre(idxKategori)) ? "GENEL" : Hucre(idxKategori).ToUpper();
                            yeniSatir["Barkod"] = barkod;
                            yeniSatir["Marka"] = string.IsNullOrWhiteSpace(Hucre(idxMarka)) ? "-" : Hucre(idxMarka).ToUpper();
                            yeniSatir["Satılan IMEI"] = imei;
                            yeniSatir["Durum"] = durum;
                            yeniSatir["Ürün Adı"] = urunAdi.ToUpper();
                            yeniSatir["Adet"] = Hucre(idxAdet);
                            yeniSatir["Sipariş No"] = sipNo;
                            yeniSatir["Alış Fiyatı"] = Hucre(idxAlis);
                            yeniSatir["Platform İndirimi"] = Hucre(idxIndirim);
                            yeniSatir["Satış Fiyatı"] = Hucre(idxSatis);
                            yeniSatir["KDV Hariç"] = Hucre(idxKdv);
                            yeniSatir["Platform"] = string.IsNullOrWhiteSpace(platform) ? "GENEL" : platform;
                            yeniSatir["Kargo Firması"] = Hucre(idxKargoFirma).ToUpper();
                            yeniSatir["Desi"] = Hucre(idxDesi);
                            yeniSatir["Kargo Ücreti"] = Hucre(idxKargoUcret);
                            yeniSatir["Komisyon"] = Hucre(idxKomisyon);
                            yeniSatir["Hizmet Bedeli"] = Hucre(idxHizmet);
                            yeniSatir["Stopaj"] = Hucre(idxStopaj);
                            yeniSatir["POS Bedeli"] = Hucre(idxPos);

                            veriTablosu.Rows.Add(yeniSatir);
                            aktarilanSatirSayaci++;
                        }
                    }

                    veriTablosu.EndLoadData();
                    dynDgvVeriGiris.ResumeLayout();

                    MessageBox.Show(
                        $"✅ Başarılı! {aktarilanSatirSayaci} kayıt önizleme tablosuna yüklendi.\n\n" +
                        "Verileri kontrol edin, ardından 'VERİLERİ SİSTEME KAYDET' butonuna basın.",
                        "Yükleme Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Dosya okunurken hata oluştu!\n\n" +
                        "HATA: " + ex.Message + "\n\n" +
                        "CSV dosyası için noktalı virgül (;) ayırıcı kullandığınızdan emin olun.\n" +
                        "Excel dosyası için ClosedXML NuGet paketi kurulu olmalıdır.",
                        "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // --- VERİTABANI ALTYAPI GÜVENCESİ ---
        private void SiparisTablosunuGuvenceyeAl(SqliteConnection conn, SqliteTransaction tx)
        {
            try
            {
                var sutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var cmd = new SqliteCommand("PRAGMA table_info('Siparisler');", conn, tx))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) sutunlar.Add(r["name"].ToString());

                var eklenecekSutunlar = new Dictionary<string, string>
                {
                    { "UrunAdi", "TEXT" },
                    { "Kategori", "TEXT" },
                    { "Marka", "TEXT" },
                    { "AlisFiyati", "REAL DEFAULT 0" },
                    { "CezaTutari", "REAL DEFAULT 0" },
                    { "CezaSebebi", "TEXT" }
                };

                foreach (var kv in eklenecekSutunlar)
                {
                    if (!sutunlar.Contains(kv.Key))
                        using (var cmd = new SqliteCommand($"ALTER TABLE Siparisler ADD COLUMN {kv.Key} {kv.Value};", conn, tx))
                            cmd.ExecuteNonQuery();
                }

                // StokHareketleri tablosu
                using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS StokHareketleri 
                    (Id INTEGER PRIMARY KEY AUTOINCREMENT, Tarih TEXT, Barkod TEXT, IslemTuru TEXT, 
                    Miktar INTEGER, EskiStok INTEGER DEFAULT 0, YeniStok INTEGER DEFAULT 0, 
                    Aciklama TEXT DEFAULT '-', IslemYapan TEXT);", conn, tx))
                    cmd.ExecuteNonQuery();

                // StokHareketleri sütun güvencesi
                var logSutunlar = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var cmd = new SqliteCommand("PRAGMA table_info('StokHareketleri');", conn, tx))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) logSutunlar.Add(r["name"].ToString());

                if (!logSutunlar.Contains("IslemYapan"))
                    using (var cmd = new SqliteCommand("ALTER TABLE StokHareketleri ADD COLUMN IslemYapan TEXT;", conn, tx))
                        cmd.ExecuteNonQuery();
            }
            catch { }
        }

        // --- VERİLERİ KAYDET ---
        private void DynBtnVerileriMuhurle_Click(object sender, EventArgs e)
        {
            if (veriTablosu == null || veriTablosu.Rows.Count == 0 || dynCmbHedefTablo.SelectedItem == null)
            {
                MessageBox.Show("Mühürlenecek hiçbir veri bulunamadı!", "Boş İşlem", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            dynDgvVeriGiris.EndEdit();

            int basariliKayit = 0;
            int atlananMukerrer = 0;
            int atlananBos = 0;

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var tx = b.BeginTransaction())
                    {
                        try
                        {
                            if (dynCmbHedefTablo.SelectedIndex == 0) // STOK/ÜRÜN MODU
                            {
                                string sql = "INSERT INTO Urunler (Barkod, Kategori, Marka, UrunAdi) VALUES (@b, @k, @m, @u) ON CONFLICT(Barkod) DO UPDATE SET UrunAdi=@u, Kategori=@k, Marka=@m";
                                using (var cmd = new SqliteCommand(sql, b, tx))
                                {
                                    foreach (DataRow r in veriTablosu.Rows)
                                    {
                                        string barkod = r["Barkod"]?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(barkod)) { atlananBos++; continue; }

                                        string kat = r["Kategori"]?.ToString()?.ToUpper() ?? "GENEL";
                                        if (string.IsNullOrWhiteSpace(kat)) kat = "GENEL";

                                        try
                                        {
                                            using (var cmdKat = new SqliteCommand("INSERT OR IGNORE INTO Kategoriler (Ad) VALUES (@ad)", b, tx))
                                            { cmdKat.Parameters.AddWithValue("@ad", kat); cmdKat.ExecuteNonQuery(); }
                                        }
                                        catch { }

                                        cmd.Parameters.Clear();
                                        cmd.Parameters.AddWithValue("@b", barkod);
                                        cmd.Parameters.AddWithValue("@k", kat);
                                        cmd.Parameters.AddWithValue("@m", r["Marka"]?.ToString()?.ToUpper() ?? "-");
                                        cmd.Parameters.AddWithValue("@u", r["Ürün Adı"]?.ToString()?.ToUpper() ?? "BİLİNMEYEN ÜRÜN");
                                        cmd.ExecuteNonQuery();
                                        basariliKayit++;
                                    }
                                }
                            }
                            else // SİPARİŞ MODU
                            {
                                SiparisTablosunuGuvenceyeAl(b, tx);

                                // Mevcut sipariş numaralarını önceden çek (hızlı kontrol için)
                                var mevcutSiparisler = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                using (var cmdCheck = new SqliteCommand("SELECT SiparisNo FROM Siparisler", b, tx))
                                using (var rCheck = cmdCheck.ExecuteReader())
                                    while (rCheck.Read()) mevcutSiparisler.Add(rCheck[0].ToString());

                                string sql = @"INSERT INTO Siparisler 
                                    (SiparisNo, PlatformAd, Barkod, Adet, SatisFiyati, PlatformIndirimi, 
                                     Stopaj, PosBedeli, HizmetBedeli, KomisyonTL, KargoUcreti, Desi, 
                                     KargoFirmasi, SatilanImei, Durum, Tarih, UrunAdi, Kategori, Marka, AlisFiyati) 
                                    VALUES 
                                    (@sn, @p, @b, @adet, @sf, @pi, @st, @pb, @hb, @km, @kr, @d, @kf, @im, @dur, @t, @u, @k, @m, @af)";

                                using (var cmd = new SqliteCommand(sql, b, tx))
                                {
                                    foreach (DataRow r in veriTablosu.Rows)
                                    {
                                        string sipNo = r["Sipariş No"]?.ToString()?.Trim();
                                        string barkod = r["Barkod"]?.ToString()?.Trim();
                                        string urunAdi = r["Ürün Adı"]?.ToString()?.Trim() ?? "BİLİNMEYEN ÜRÜN";
                                        string kategori = r["Kategori"]?.ToString()?.Trim() ?? "GENEL";
                                        string marka = r["Marka"]?.ToString()?.Trim() ?? "-";

                                        if (string.IsNullOrWhiteSpace(sipNo)) { atlananBos++; continue; }

                                        // MÜKERRER KONTROL
                                        if (mevcutSiparisler.Contains(sipNo)) { atlananMukerrer++; continue; }

                                        if (string.IsNullOrWhiteSpace(barkod) || barkod == "-")
                                            barkod = string.IsNullOrWhiteSpace(urunAdi) ? "BILINMEYEN" : urunAdi.ToUpper();

                                        if (string.IsNullOrWhiteSpace(kategori)) kategori = "GENEL";

                                        try
                                        {
                                            using (var cmdKat = new SqliteCommand("INSERT OR IGNORE INTO Kategoriler (Ad) VALUES (@ad)", b, tx))
                                            { cmdKat.Parameters.AddWithValue("@ad", kategori.ToUpper()); cmdKat.ExecuteNonQuery(); }
                                        }
                                        catch { }

                                        // Platform ekle (Platformlar tablosuna)
                                        string platform = r["Platform"]?.ToString()?.Trim()?.ToUpper() ?? "GENEL";
                                        if (string.IsNullOrWhiteSpace(platform)) platform = "GENEL";
                                        if (platform != "GENEL")
                                        {
                                            try
                                            {
                                                using (var cmdPlat = new SqliteCommand(
                                                    "INSERT INTO Platformlar (Ad) SELECT @p WHERE NOT EXISTS (SELECT 1 FROM Platformlar WHERE Ad = @p)", b, tx))
                                                {
                                                    cmdPlat.Parameters.AddWithValue("@p", platform);
                                                    cmdPlat.ExecuteNonQuery();
                                                }
                                            }
                                            catch { }
                                        }

                                        int adet = SafeInt(r["Adet"]);
                                        double sf = SafeDouble(r["Satış Fiyatı"]);
                                        double pi = SafeDouble(r["Platform İndirimi"]);
                                        double s_t = SafeDouble(r["Stopaj"]);
                                        double pb = SafeDouble(r["POS Bedeli"]);
                                        double hb = SafeDouble(r["Hizmet Bedeli"]);
                                        double km = SafeDouble(r["Komisyon"]);
                                        double kr = SafeDouble(r["Kargo Ücreti"]);
                                        double af = SafeDouble(r["Alış Fiyatı"]);
                                        string tarihFormatli = FormatliTarihGetir(r["Tarih"]);
                                        string durumSon = r["Durum"]?.ToString()?.Trim()?.ToUpper() ?? "";

                                        cmd.Parameters.Clear();
                                        cmd.Parameters.AddWithValue("@sn", sipNo);
                                        cmd.Parameters.AddWithValue("@p", platform);
                                        cmd.Parameters.AddWithValue("@b", barkod);
                                        cmd.Parameters.AddWithValue("@adet", adet);
                                        cmd.Parameters.AddWithValue("@sf", sf);
                                        cmd.Parameters.AddWithValue("@pi", pi);
                                        cmd.Parameters.AddWithValue("@st", s_t);
                                        cmd.Parameters.AddWithValue("@pb", pb);
                                        cmd.Parameters.AddWithValue("@hb", hb);
                                        cmd.Parameters.AddWithValue("@km", km);
                                        cmd.Parameters.AddWithValue("@kr", kr);
                                        cmd.Parameters.AddWithValue("@d", r["Desi"]?.ToString() ?? "1");
                                        cmd.Parameters.AddWithValue("@kf", r["Kargo Firması"]?.ToString()?.ToUpper() ?? "-");
                                        cmd.Parameters.AddWithValue("@im", r["Satılan IMEI"]?.ToString() ?? "-");
                                        cmd.Parameters.AddWithValue("@dur", durumSon);
                                        cmd.Parameters.AddWithValue("@t", tarihFormatli);
                                        cmd.Parameters.AddWithValue("@u", urunAdi);
                                        cmd.Parameters.AddWithValue("@k", kategori);
                                        cmd.Parameters.AddWithValue("@m", marka);
                                        cmd.Parameters.AddWithValue("@af", af);
                                        cmd.ExecuteNonQuery();

                                        // Stok hareketi logu (geçmiş aktarım - stok adedini ETKİLEMİYOR)
                                        try
                                        {
                                            using (var cmdLog = new SqliteCommand(@"INSERT INTO StokHareketleri 
                                                (Tarih, Barkod, IslemTuru, Miktar, EskiStok, YeniStok, Aciklama, IslemYapan) 
                                                VALUES (@tarih, @b, 'GEÇMİŞ SATIŞ', @a, 0, 0, @ack, @yapan)", b, tx))
                                            {
                                                cmdLog.Parameters.AddWithValue("@tarih", tarihFormatli + " 00:00:00");
                                                cmdLog.Parameters.AddWithValue("@b", barkod);
                                                cmdLog.Parameters.AddWithValue("@a", adet);
                                                cmdLog.Parameters.AddWithValue("@ack", $"Excel Aktarım | Sipariş: {sipNo} | Ürün: {urunAdi}");
                                                cmdLog.Parameters.AddWithValue("@yapan", Veritabanı.AktifKullanici);
                                                cmdLog.ExecuteNonQuery();
                                            }
                                        }
                                        catch { }

                                        mevcutSiparisler.Add(sipNo); // Aynı oturumda tekrar eklemeyi engelle
                                        basariliKayit++;
                                    }
                                }
                            }

                            tx.Commit();

                            try { SayfaStokHareketleri.OnStokHareketEklendi?.Invoke(); } catch { }

                            string mesaj = $"✅ {basariliKayit} kayıt başarıyla sisteme eklendi.";
                            if (atlananMukerrer > 0) mesaj += $"\n⚠️ {atlananMukerrer} mükerrer sipariş atlandı (zaten kayıtlı).";
                            if (atlananBos > 0) mesaj += $"\nℹ️ {atlananBos} boş satır atlandı.";

                            MessageBox.Show(mesaj, "İşlem Tamamlandı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            veriTablosu.Clear();
                            for (int i = 0; i < 5; i++) veriTablosu.Rows.Add(veriTablosu.NewRow());
                        }
                        catch (Exception exInner) { tx.Rollback(); throw exInner; }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kayıt sırasında hata oluştu:\n\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- SİPARİŞ SORGULA ---
        private void DynBtnSiparisSorgula_Click(object sender, EventArgs e)
        {
            if (veriTablosu == null || veriTablosu.Rows.Count == 0 || !veriTablosu.Columns.Contains("Sipariş No"))
            {
                MessageBox.Show("Tabloda sorgulanacak sipariş verisi bulunamadı!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var nolar = new List<string>();
            foreach (DataRow r in veriTablosu.Rows)
            {
                string no = r["Sipariş No"]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(no)) nolar.Add(no);
            }
            if (nolar.Count == 0) { MessageBox.Show("Geçerli sipariş numarası bulunamadı!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            SiparisSorgulamaMotoru(nolar.Distinct().ToList());
        }

        private void SiparisSorgulamaMotoru(List<string> siparisler)
        {
            if (siparisler == null || siparisler.Count == 0) return;
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    var paramNames = new List<string>();
                    var cmd = new SqliteCommand { Connection = conn };
                    for (int i = 0; i < siparisler.Count; i++)
                    {
                        string pName = "@sn" + i;
                        paramNames.Add(pName);
                        cmd.Parameters.AddWithValue(pName, siparisler[i]);
                    }
                    cmd.CommandText = $@"SELECT Tarih, PlatformAd as 'Platform', SiparisNo as 'Sipariş No', 
                        IFNULL(UrunAdi, Barkod) as 'Ürün', Adet, SatisFiyati as 'Satış Fiyatı', 
                        PlatformIndirimi as 'İndirim',
                        (IFNULL(Stopaj,0)+IFNULL(PosBedeli,0)+IFNULL(HizmetBedeli,0)+IFNULL(KomisyonTL,0)+IFNULL(KargoUcreti,0)) as 'Toplam Kesinti', 
                        Durum 
                        FROM Siparisler WHERE SiparisNo IN ({string.Join(", ", paramNames)}) ORDER BY Tarih DESC";

                    DataTable dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Bu sipariş numaraları veritabanında bulunamadı.", "Kayıt Yok", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    Form frmSorgu = new Form { Text = $"🔎 Sipariş Önizleme - {dt.Rows.Count} Kayıt", Size = new Size(1100, 520), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ForeColor = Color.White, ShowIcon = false };
                    DataGridView dgv = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, BackgroundColor = Color.FromArgb(24, 24, 27), RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Font = new Font("Segoe UI", 9.5F), BorderStyle = BorderStyle.None, EnableHeadersVisualStyles = false };
                    dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
                    dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Orange;
                    dgv.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
                    dgv.DefaultCellStyle.ForeColor = Color.White;
                    dgv.DataSource = dt;
                    frmSorgu.Controls.Add(dgv);
                    frmSorgu.ShowDialog();
                }
            }
            catch (Exception ex) { MessageBox.Show("Sorgulama hatası: " + ex.Message); }
        }

        public void button1_Click(object sender, EventArgs e) { }
        public void button2_Click(object sender, EventArgs e) { }
        public void button3_Click(object sender, EventArgs e) { }
        public void button4_Click(object sender, EventArgs e) { }
    }
}