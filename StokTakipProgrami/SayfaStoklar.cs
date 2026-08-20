using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.Sqlite;
using System.Globalization;

namespace StokTakipProgrami
{
    // 📋 ÜRÜN EKLE / ASIL STOK ENVANTER GİRİŞ EKRANI (2. EL HURDA/İADE MOTORU VE SIFIR HATALI FORM)
    public partial class SayfaStoklar : UserControl
    {
        private TextBox stkTxtBarkod, stkTxtAdet, stkTxtAlisFiyati, stkTxtImei;
        private ComboBox stkCmbFirma;
        private Label stkLblUrunAdi, stkLblKategori;
        private DateTimePicker stkDtpTarih;
        private Button stkBtnKaydet, stkBtnUrunAra;
        private DataGridView stkDgvStokListesi;

        // 🚀 MÜHÜR: Yeni / 2. El Durum Seçimi İçin Radyo Butonlar
        private RadioButton rbSifir, rbIkinciEl;

        private string otomatikKategori = "GENEL";
        private List<string> aktifImeiListesi = new List<string>();

        public SayfaStoklar()
        {
            InitializeComponent();
            TabloAltyapisiniGuvenceyeAl();
            ArayuzInsaEt();
            DinamikFirmalariYukle();
        }

        public void SayfaStoklar_Load(object sender, EventArgs e)
        {
            MevcutStoklariListele();
            if (stkTxtBarkod != null) stkTxtBarkod.Focus();
        }

        private void TabloAltyapisiniGuvenceyeAl()
        {
            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var cmd = new SqliteCommand("CREATE TABLE IF NOT EXISTS Firmalar (Ad TEXT UNIQUE);", b)) { cmd.ExecuteNonQuery(); }
                }
            }
            catch { }
        }

        private void ArayuzInsaEt()
        {
            this.Controls.Clear(); this.Dock = DockStyle.Fill; this.BackColor = Color.FromArgb(33, 33, 36);

            Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(26, 26, 39) };
            Label lblT = new Label { Text = "📋 ASIL STOK / ENVANTER GİRİŞ MERKEZİ (ÜRÜN EKLE)", Dock = DockStyle.Fill, Font = new Font("Segoe UI", 12F, FontStyle.Bold), ForeColor = Color.White, TextAlign = ContentAlignment.MiddleCenter };
            pnlHeader.Controls.Add(lblT); this.Controls.Add(pnlHeader);

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = Color.Transparent };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380F)); // SOL: FORM
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));  // SAĞ: LİSTE
            this.Controls.Add(mainLayout);

            // --- 🟢 SOL PANEL (GİRİŞ FORMU) ---
            Panel pnlLeftForm = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(26, 26, 28), Padding = new Padding(10) };

            Label lblVeriEkleBaslik = new Label { Text = "🟢 STOK GİRİŞİ", Location = new Point(15, 15), AutoSize = true, Font = new Font("Segoe UI Black", 14F, FontStyle.Bold), ForeColor = Color.White };
            pnlLeftForm.Controls.Add(lblVeriEkleBaslik);

            pnlLeftForm.Controls.Add(FormElemaniYarat("Barkod (Okut/Yaz):", out stkTxtBarkod, new Point(15, 55), 200));
            stkTxtBarkod.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; UrunKartiniSorgula(); } };

            stkBtnUrunAra = new Button { Text = "🔍 ARA", Location = new Point(225, 73), Width = 110, Height = 30, BackColor = Color.FromArgb(230, 126, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            stkBtnUrunAra.FlatAppearance.BorderSize = 0; stkBtnUrunAra.Click += (s, e) => UrunKartiniSorgula();
            pnlLeftForm.Controls.Add(stkBtnUrunAra);

            Panel pnlCardInfo = new Panel { Location = new Point(15, 120), Width = 320, Height = 60, BackColor = Color.FromArgb(33, 33, 36), BorderStyle = BorderStyle.FixedSingle };
            stkLblUrunAdi = new Label { Text = "Ürün Adı: Kart bekleniyor...", Location = new Point(10, 8), AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ForeColor = Color.Gray };
            stkLblKategori = new Label { Text = "Kategori: -", Location = new Point(10, 32), AutoSize = true, Font = new Font("Segoe UI", 8.5F), ForeColor = Color.Gold };
            pnlCardInfo.Controls.Add(stkLblUrunAdi); pnlCardInfo.Controls.Add(stkLblKategori);
            pnlLeftForm.Controls.Add(pnlCardInfo);

            pnlLeftForm.Controls.Add(FormElemaniYarat("Adet:", out stkTxtAdet, new Point(15, 190), 80, "1"));
            pnlLeftForm.Controls.Add(FormElemaniYarat("Alış (Maliyet TL):", out stkTxtAlisFiyati, new Point(115, 190), 220, "0,00"));

            stkTxtAlisFiyati.Leave += StkTxtAlisFiyati_Leave;

            // 🚀 MÜHÜR: Ürün Durumu (Sıfır / 2. El) Seçimi
            Panel pnlDurum = new Panel { Location = new Point(15, 255), Width = 320, Height = 55 };
            Label lblDurum = new Label { Text = "Ürün Kondisyonu:", Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            rbSifir = new RadioButton { Text = "Sıfır (Orijinal)", Location = new Point(5, 20), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.White, Checked = true, Cursor = Cursors.Hand };
            rbIkinciEl = new RadioButton { Text = "2. El (Hurda / İade)", Location = new Point(140, 20), AutoSize = true, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.Orange, Cursor = Cursors.Hand };

            // Eğer radyo buton değişirse anında barkodu tarayarak isme (2. EL) ekinin yansımasını sağlar
            rbSifir.CheckedChanged += (s, e) => { if (rbSifir.Checked) UrunKartiniSorgula(); };
            rbIkinciEl.CheckedChanged += (s, e) => { if (rbIkinciEl.Checked) UrunKartiniSorgula(); };

            pnlDurum.Controls.Add(lblDurum); pnlDurum.Controls.Add(rbSifir); pnlDurum.Controls.Add(rbIkinciEl);
            pnlLeftForm.Controls.Add(pnlDurum);

            pnlLeftForm.Controls.Add(FormComboYarat("Tedarikçi Firma:", out stkCmbFirma, new Point(15, 320), 320));

            pnlLeftForm.Controls.Add(FormElemaniYarat("IMEI Numarası:", out stkTxtImei, new Point(15, 385), 320));
            stkTxtImei.Enabled = false;

            stkTxtAdet.TextChanged += (s, e) => {
                int.TryParse(stkTxtAdet.Text, out int adet);
                if (otomatikKategori.Contains("TELEFON") || otomatikKategori.Contains("CEP"))
                {
                    if (adet > 1)
                    {
                        stkTxtImei.ReadOnly = true;
                        stkTxtImei.Text = "Çoklu giriş için buraya tıklayın...";
                    }
                    else
                    {
                        stkTxtImei.ReadOnly = false;
                        if (stkTxtImei.Text.Contains("tıklayın") || stkTxtImei.Text.Contains("GİRİLDİ")) stkTxtImei.Clear();
                        aktifImeiListesi.Clear();
                    }
                }
            };

            stkTxtImei.Click += (s, e) => {
                int.TryParse(stkTxtAdet.Text, out int adet);
                if (adet > 1 && (otomatikKategori.Contains("TELEFON") || otomatikKategori.Contains("CEP")))
                {
                    CokluImeiPenceresiAc(adet);
                }
            };

            Label lblTarih = new Label { Text = "Stok Giriş Tarihi:", Location = new Point(15, 450), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray };
            stkDtpTarih = new DateTimePicker { Location = new Point(15, 470), Width = 320, Font = new Font("Segoe UI", 10F), Format = DateTimePickerFormat.Short };
            pnlLeftForm.Controls.Add(lblTarih); pnlLeftForm.Controls.Add(stkDtpTarih);

            stkBtnKaydet = new Button { Text = "💾 KAYDET", Location = new Point(15, 515), Width = 320, Height = 45, BackColor = Color.FromArgb(46, 117, 89), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
            stkBtnKaydet.FlatAppearance.BorderSize = 0; stkBtnKaydet.Click += StkBtnKaydet_Click;
            pnlLeftForm.Controls.Add(stkBtnKaydet);

            // --- 📋 SAĞ PANEL (LİSTE ALANI) ---
            Panel pnlRightList = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
            Label lblListTitle = new Label { Text = "📦 MEVCUT STOK ENVANTERİ (Son Eklenen En Üstte)", Dock = DockStyle.Top, Height = 30, Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.DarkGray };
            pnlRightList.Controls.Add(lblListTitle);

            stkDgvStokListesi = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.FromArgb(24, 24, 27),
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 9.5F),
                BorderStyle = BorderStyle.None,
            };

            stkDgvStokListesi.DataError += (s, e) => { e.Cancel = true; };
            pnlRightList.Controls.Add(stkDgvStokListesi);
            stkDgvStokListesi.BringToFront();

            mainLayout.Controls.Add(pnlLeftForm, 0, 0);
            mainLayout.Controls.Add(pnlRightList, 1, 0);
        }

        private void StkTxtAlisFiyati_Leave(object sender, EventArgs e)
        {
            string hamVeri = stkTxtAlisFiyati.Text.Trim().Replace("₺", "").Trim();
            if (string.IsNullOrEmpty(hamVeri)) return;
            if (double.TryParse(hamVeri.Replace(".", "").Replace(",", "."), out double dblVal))
            {
                stkTxtAlisFiyati.Text = dblVal.ToString("N2", CultureInfo.GetCultureInfo("tr-TR"));
            }
        }

        private void DinamikFirmalariYukle()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT Ad FROM Firmalar ORDER BY Ad ASC", conn))
                    using (var r = cmd.ExecuteReader())
                    {
                        stkCmbFirma.Items.Clear();
                        while (r.Read()) stkCmbFirma.Items.Add(r["Ad"].ToString());
                    }
                }
            }
            catch { }
            if (stkCmbFirma.Items.Count == 0) stkCmbFirma.Items.AddRange(new object[] { "TANIMSIZ FİRMA" });
            stkCmbFirma.SelectedIndex = 0;
        }

        private void MevcutStoklariListele()
        {
            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    string sql = @"
                        SELECT s.Barkod, u.UrunAdi as 'Ürün Adı', u.Marka as 'Marka', 
                               s.Adet, s.StokMiktari as 'Mevcut Stok', s.AlisFiyati as 'Maliyet (TL)', s.Firma, s.Tarih 
                        FROM Stoklar s 
                        LEFT JOIN Urunler u ON s.Barkod = u.Barkod 
                        ORDER BY s.rowid DESC";
                    using (var cmd = new SqliteCommand(sql, conn))
                    {
                        DataTable dt = new DataTable();
                        using (var reader = cmd.ExecuteReader()) { dt.Load(reader); }
                        stkDgvStokListesi.DataSource = dt;
                    }
                }

                stkDgvStokListesi.EnableHeadersVisualStyles = false;
                stkDgvStokListesi.ColumnHeadersVisible = true;
                stkDgvStokListesi.ColumnHeadersHeight = 35;
                stkDgvStokListesi.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
                stkDgvStokListesi.ColumnHeadersDefaultCellStyle.ForeColor = Color.Orange;
                stkDgvStokListesi.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);

                stkDgvStokListesi.DefaultCellStyle.BackColor = Color.FromArgb(33, 33, 36);
                stkDgvStokListesi.DefaultCellStyle.ForeColor = Color.White;
                stkDgvStokListesi.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);

                if (stkDgvStokListesi.Columns["Maliyet (TL)"] != null)
                {
                    stkDgvStokListesi.Columns["Maliyet (TL)"].DefaultCellStyle.Format = "N2";
                }

                stkDgvStokListesi.ClearSelection();
            }
            catch { }
        }

        private void UrunKartiniSorgula()
        {
            string b = stkTxtBarkod.Text.Trim(); if (string.IsNullOrEmpty(b)) return;

            // 🚀 MÜHÜR: Barkod okutulduğunda eğer zaten -2EL eki varsa kök barkodu arat
            string arananBarkod = b;
            if (arananBarkod.EndsWith("-2EL")) arananBarkod = arananBarkod.Replace("-2EL", "");

            try
            {
                using (var conn = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT UrunAdi, Kategori, Marka FROM Urunler WHERE Barkod=@b", conn))
                    {
                        cmd.Parameters.AddWithValue("@b", arananBarkod);
                        using (var r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                // 🚀 MÜHÜR: 2. El seçili ise Ürün Adına (2. EL) damgası vurulur!
                                string gercekAd = r["UrunAdi"].ToString();
                                if (rbIkinciEl.Checked && !gercekAd.Contains("(2. EL)")) gercekAd += " (2. EL)";

                                stkLblUrunAdi.Text = "Ürün Adı: " + gercekAd;
                                stkLblUrunAdi.ForeColor = rbIkinciEl.Checked ? Color.Orange : Color.White;

                                otomatikKategori = r["Kategori"].ToString().ToUpper();
                                string mrk = r["Marka"].ToString();
                                if (string.IsNullOrEmpty(mrk)) mrk = "-";

                                stkLblKategori.Text = $"Kategori: {otomatikKategori} | Marka: {mrk}";
                                stkLblKategori.ForeColor = Color.Gold;

                                if (otomatikKategori.Contains("TELEFON") || otomatikKategori.Contains("CEP"))
                                {
                                    stkTxtImei.Enabled = true; stkTxtImei.BackColor = Color.FromArgb(41, 24, 24);
                                    stkTxtAdet.Enabled = true; stkTxtImei.Focus();
                                }
                                else
                                {
                                    stkTxtImei.Clear(); stkTxtImei.Enabled = false; stkTxtImei.BackColor = Color.FromArgb(41, 41, 44);
                                    stkTxtAdet.Enabled = true; stkTxtAdet.Focus();
                                }
                            }
                            else
                            {
                                stkLblUrunAdi.Text = "⚠️ KART BULUNAMADI! Önce 'Veri Ekle'den kart açın."; stkLblUrunAdi.ForeColor = Color.Crimson;
                                stkLblKategori.Text = "Kategori: -"; otomatikKategori = "GENEL";
                                stkTxtImei.Clear(); stkTxtImei.Enabled = false; stkTxtAdet.Enabled = true;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void CokluImeiPenceresiAc(int adet)
        {
            Form frmImei = new Form { Text = "📱 Çoklu Cihaz IMEI Kayıt", Size = new Size(380, 450), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(33, 33, 36), ShowIcon = false, MinimizeBox = false, MaximizeBox = false };

            Panel pnlScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(15) };
            Button btnOnay = new Button { Text = "✔ IMEI NUMARALARINI ONAYLA", Dock = DockStyle.Bottom, Height = 50, BackColor = Color.ForestGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnOnay.FlatAppearance.BorderSize = 0;

            Label lblBilgi = new Label { Text = $"Toplam {adet} cihaz için IMEI giriniz:", Dock = DockStyle.Top, Height = 30, ForeColor = Color.Gold, Font = new Font("Segoe UI", 10F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
            pnlScroll.Controls.Add(lblBilgi);

            List<TextBox> dinamikKutular = new List<TextBox>();
            int yPos = 40;
            for (int i = 0; i < adet; i++)
            {
                Label l = new Label { Text = $"{i + 1}. Cihaz IMEI:", Location = new Point(20, yPos), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };
                TextBox t = new TextBox { Location = new Point(20, yPos + 20), Width = 300, Font = new Font("Segoe UI", 11F), BackColor = Color.FromArgb(45, 45, 48), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

                if (i < aktifImeiListesi.Count) t.Text = aktifImeiListesi[i];

                pnlScroll.Controls.Add(l); pnlScroll.Controls.Add(t);
                dinamikKutular.Add(t);
                yPos += 65;
            }

            btnOnay.Click += (sX, eX) => {
                aktifImeiListesi.Clear();
                foreach (var kutu in dinamikKutular)
                {
                    if (!string.IsNullOrWhiteSpace(kutu.Text)) aktifImeiListesi.Add(kutu.Text.Trim());
                }
                if (aktifImeiListesi.Count < adet) { MessageBox.Show("Lütfen tüm kutuları doldurun!", "Eksik Veri", MessageBoxButtons.OK, MessageBoxIcon.Warning); }

                stkTxtImei.Text = $"[{aktifImeiListesi.Count} ADET IMEI GİRİLDİ]";
                frmImei.Close();
            };

            frmImei.Controls.Add(pnlScroll);
            frmImei.Controls.Add(btnOnay);
            frmImei.ShowDialog();
        }

        private void StkBtnKaydet_Click(object sender, EventArgs e)
        {
            string bkd = stkTxtBarkod.Text.Trim(); string firma = stkCmbFirma.Text; string tarih = stkDtpTarih.Value.ToString("dd.MM.yyyy");
            if (string.IsNullOrEmpty(bkd) || stkLblUrunAdi.Text.Contains("BULUNAMADI")) { MessageBox.Show("Lütfen geçerli bir ürün kartı barkodu okutun!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            // 🚀 MÜHÜR: Eğer radyo buton "2. El" ise stok barkodu otomatik olarak -2EL olur.
            if (rbIkinciEl.Checked && !bkd.EndsWith("-2EL"))
            {
                bkd = bkd + "-2EL";
            }

            int.TryParse(stkTxtAdet.Text, out int girilenAdet); if (girilenAdet <= 0) girilenAdet = 1;

            if (stkTxtImei.Enabled)
            {
                if (girilenAdet == 1)
                {
                    if (string.IsNullOrWhiteSpace(stkTxtImei.Text) || stkTxtImei.Text.Contains("tıklayın"))
                    {
                        MessageBox.Show("Lütfen cihazın IMEI numarasını kutuya giriniz!", "Eksik IMEI", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
                    }
                    aktifImeiListesi.Clear();
                    aktifImeiListesi.Add(stkTxtImei.Text.Trim());
                }
                else
                {
                    if (aktifImeiListesi.Count != girilenAdet)
                    {
                        MessageBox.Show($"Eksik IMEI girişi! Lütfen kutuya tıklayıp {girilenAdet} adet cihazın IMEI'sini eksiksiz girin.", "IMEI Adet Uyuşmazlığı", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
                    }
                }
            }

            double.TryParse(stkTxtAlisFiyati.Text.Replace(".", "").Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double girilenAlis);
            string urunName = stkLblUrunAdi.Text.Replace("Ürün Adı: ", "").Trim();

            try
            {
                using (var b = new SqliteConnection(Veritabanı.BaglantiYolu))
                {
                    b.Open();
                    using (var tx = b.BeginTransaction())
                    {
                        bool urunVarMi = false; int mevcutAdet = 0; double mevcutAlis = 0;
                        using (var cmd = new SqliteCommand("SELECT Adet, AlisFiyati FROM Stoklar WHERE Barkod=@b", b, tx))
                        {
                            cmd.Parameters.AddWithValue("@b", bkd);
                            using (var r = cmd.ExecuteReader()) { if (r.Read()) { urunVarMi = true; mevcutAdet = Convert.ToInt32(r["Adet"]); mevcutAlis = Convert.ToDouble(r["AlisFiyati"]); } }
                        }

                        if (urunVarMi)
                        {
                            double yeniAlisMaliyeti = ((mevcutAdet * mevcutAlis) + (girilenAdet * girilenAlis)) / (mevcutAdet + girilenAdet);
                            int yeniToplamStok = mevcutAdet + girilenAdet;

                            using (var cmdUp = new SqliteCommand("UPDATE Stoklar SET Adet=@t, StokMiktari=@t, AlisFiyati=@a, Firma=@f, Tarih=@dt WHERE Barkod=@b", b, tx))
                            {
                                cmdUp.Parameters.AddWithValue("@t", yeniToplamStok); cmdUp.Parameters.AddWithValue("@a", yeniAlisMaliyeti);
                                cmdUp.Parameters.AddWithValue("@f", string.IsNullOrEmpty(firma) ? "-" : firma);
                                cmdUp.Parameters.AddWithValue("@dt", tarih); cmdUp.Parameters.AddWithValue("@b", bkd);
                                cmdUp.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            using (var cmdIn = new SqliteCommand("INSERT INTO Stoklar (Barkod, UrunAdi, Kategori, Adet, StokMiktari, AlisFiyati, SatisFiyati, Firma, Tarih) VALUES (@b, @u, @k, @adet, @adet, @alis, 0, @f, @t)", b, tx))
                            {
                                cmdIn.Parameters.AddWithValue("@b", bkd); cmdIn.Parameters.AddWithValue("@u", urunName); cmdIn.Parameters.AddWithValue("@k", otomatikKategori);
                                cmdIn.Parameters.AddWithValue("@adet", girilenAdet); cmdIn.Parameters.AddWithValue("@alis", girilenAlis);
                                cmdIn.Parameters.AddWithValue("@f", string.IsNullOrEmpty(firma) ? "-" : firma); cmdIn.Parameters.AddWithValue("@t", tarih);
                                cmdIn.ExecuteNonQuery();
                            }
                        }

                        if (stkTxtImei.Enabled && aktifImeiListesi.Count > 0)
                        {
                            foreach (string imei in aktifImeiListesi)
                            {
                                using (var cmdI = new SqliteCommand("INSERT INTO Imeiler (ImeiNo, Barkod, Durum) VALUES (@i, @b, 'STOKTA')", b, tx))
                                {
                                    cmdI.Parameters.AddWithValue("@i", imei); cmdI.Parameters.AddWithValue("@b", bkd); cmdI.ExecuteNonQuery();
                                }
                            }
                        }

                        tx.Commit();
                        MessageBox.Show("🎉 Stok başarıyla envantere dâhil edildi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        EkranıTemizle();
                        MevcutStoklariListele();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message); }
        }

        private void EkranıTemizle()
        {
            stkTxtBarkod.Clear();
            stkTxtImei.Clear(); stkTxtImei.Enabled = false; stkTxtImei.BackColor = Color.FromArgb(41, 41, 44);
            aktifImeiListesi.Clear();
            stkTxtAdet.Text = "1"; stkTxtAlisFiyati.Text = "0,00";
            if (stkCmbFirma.Items.Count > 0) stkCmbFirma.SelectedIndex = 0;
            stkLblUrunAdi.Text = "Ürün Adı: Kart bekleniyor..."; stkLblUrunAdi.ForeColor = Color.Gray;
            stkLblKategori.Text = "Kategori: -";
            rbSifir.Checked = true; // Temizlendiğinde varsayılan sıfıra döner
            stkDtpTarih.Value = DateTime.Now; stkTxtBarkod.Focus();
        }

        private Panel FormElemaniYarat(string baslik, out TextBox kutu, Point konum, int genislik, string varsayilan = "") { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DarkGray }; kutu = new TextBox { Text = varsayilan, Location = new Point(0, 18), Width = genislik, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
        private Panel FormComboYarat(string baslik, out ComboBox kutu, Point konum, int genislik) { Panel p = new Panel { Location = konum, Width = genislik, Height = 55 }; Label lbl = new Label { Text = baslik, Location = new Point(0, 0), AutoSize = true, ForeColor = Color.DarkGray, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) }; kutu = new ComboBox { Location = new Point(0, 18), Width = genislik, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(41, 41, 44), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 10F) }; p.Controls.Add(lbl); p.Controls.Add(kutu); return p; }
    }
}