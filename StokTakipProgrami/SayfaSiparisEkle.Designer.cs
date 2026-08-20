namespace StokTakipProgrami
{
    partial class SayfaSiparisEkle
    {
        /// <summary> 
        ///Gerekli tasarımcı değişkeni.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        ///Kullanılan tüm kaynakları temizleyin.
        /// </summary>
        ///<param name="disposing">yönetilen kaynaklar dispose edilmeliyse doğru; aksi halde yanlış.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Bileşen Tasarımcısı üretimi kod

        /// <summary> 
        /// Tasarımcı desteği için gerekli metot - bu metodun 
        ///içeriğini kod düzenleyici ile değiştirmeyin.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtSiparisNo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtBarkod = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnUrunCagir = new System.Windows.Forms.Button();
            this.lblUrunAdi = new System.Windows.Forms.Label();
            this.lblKategori = new System.Windows.Forms.Label();
            this.numAdet = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSatisFiyati = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.dtpSiparisTarihi = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbDurum = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btnSiparisKaydet = new System.Windows.Forms.Button();
            this.lblAlisFiyati = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbPlatform = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.cmbKargoFirmasi = new System.Windows.Forms.ComboBox();
            this.numDesi = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.txtKargoBedeli = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtStopaj = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtPlatformIndirimi = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.txtNetKar = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.dgvGunlukSiparisler = new System.Windows.Forms.DataGridView();
            this.txtImei = new System.Windows.Forms.TextBox();
            this.lblImei = new System.Windows.Forms.Label();
            this.dgvSepet = new System.Windows.Forms.DataGridView();
            this.btnSepeteEkle = new System.Windows.Forms.Button();
            this.txtKomisyon = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.txtHizmetBedeli = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.txtPosBedeli = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.txtGunlukSiparis = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.txtGunlukKar = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numAdet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDesi)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGunlukSiparisler)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSepet)).BeginInit();
            this.SuspendLayout();
            // 
            // txtSiparisNo
            // 
            this.txtSiparisNo.Location = new System.Drawing.Point(137, 417);
            this.txtSiparisNo.Name = "txtSiparisNo";
            this.txtSiparisNo.Size = new System.Drawing.Size(225, 29);
            this.txtSiparisNo.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(221, 386);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sipariş No";
            // 
            // txtBarkod
            // 
            this.txtBarkod.Location = new System.Drawing.Point(617, 135);
            this.txtBarkod.Name = "txtBarkod";
            this.txtBarkod.Size = new System.Drawing.Size(225, 29);
            this.txtBarkod.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(706, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(74, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "Barkod";
            // 
            // btnUrunCagir
            // 
            this.btnUrunCagir.Location = new System.Drawing.Point(682, 187);
            this.btnUrunCagir.Name = "btnUrunCagir";
            this.btnUrunCagir.Size = new System.Drawing.Size(121, 38);
            this.btnUrunCagir.TabIndex = 4;
            this.btnUrunCagir.Text = "Ara";
            this.btnUrunCagir.UseVisualStyleBackColor = true;
            // 
            // lblUrunAdi
            // 
            this.lblUrunAdi.AutoSize = true;
            this.lblUrunAdi.Location = new System.Drawing.Point(672, 317);
            this.lblUrunAdi.Name = "lblUrunAdi";
            this.lblUrunAdi.Size = new System.Drawing.Size(64, 25);
            this.lblUrunAdi.TabIndex = 5;
            this.lblUrunAdi.Text = "label3";
            // 
            // lblKategori
            // 
            this.lblKategori.AutoSize = true;
            this.lblKategori.Location = new System.Drawing.Point(468, 317);
            this.lblKategori.Name = "lblKategori";
            this.lblKategori.Size = new System.Drawing.Size(64, 25);
            this.lblKategori.TabIndex = 6;
            this.lblKategori.Text = "label3";
            // 
            // numAdet
            // 
            this.numAdet.Location = new System.Drawing.Point(623, 414);
            this.numAdet.Name = "numAdet";
            this.numAdet.Size = new System.Drawing.Size(120, 29);
            this.numAdet.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(653, 386);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 25);
            this.label4.TabIndex = 10;
            this.label4.Text = "Adet";
            // 
            // txtSatisFiyati
            // 
            this.txtSatisFiyati.Location = new System.Drawing.Point(787, 417);
            this.txtSatisFiyati.Name = "txtSatisFiyati";
            this.txtSatisFiyati.Size = new System.Drawing.Size(120, 29);
            this.txtSatisFiyati.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(822, 386);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 25);
            this.label5.TabIndex = 12;
            this.label5.Text = "Fiyat";
            // 
            // dtpSiparisTarihi
            // 
            this.dtpSiparisTarihi.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpSiparisTarihi.Location = new System.Drawing.Point(932, 423);
            this.dtpSiparisTarihi.Name = "dtpSiparisTarihi";
            this.dtpSiparisTarihi.Size = new System.Drawing.Size(160, 29);
            this.dtpSiparisTarihi.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(998, 386);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(57, 25);
            this.label6.TabIndex = 14;
            this.label6.Text = "Tarih";
            // 
            // cmbDurum
            // 
            this.cmbDurum.FormattingEnabled = true;
            this.cmbDurum.Items.AddRange(new object[] {
            "Bekliyor",
            "Tamamlandı",
            "İptal Edildi"});
            this.cmbDurum.Location = new System.Drawing.Point(1168, 136);
            this.cmbDurum.Name = "cmbDurum";
            this.cmbDurum.Size = new System.Drawing.Size(193, 32);
            this.cmbDurum.TabIndex = 15;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(1229, 89);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 25);
            this.label7.TabIndex = 16;
            this.label7.Text = "Durum";
            // 
            // btnSiparisKaydet
            // 
            this.btnSiparisKaydet.Location = new System.Drawing.Point(1026, 665);
            this.btnSiparisKaydet.Name = "btnSiparisKaydet";
            this.btnSiparisKaydet.Size = new System.Drawing.Size(266, 49);
            this.btnSiparisKaydet.TabIndex = 17;
            this.btnSiparisKaydet.Text = "Siparişi Kaydet";
            this.btnSiparisKaydet.UseVisualStyleBackColor = true;
            // 
            // lblAlisFiyati
            // 
            this.lblAlisFiyati.AutoSize = true;
            this.lblAlisFiyati.Location = new System.Drawing.Point(934, 317);
            this.lblAlisFiyati.Name = "lblAlisFiyati";
            this.lblAlisFiyati.Size = new System.Drawing.Size(64, 25);
            this.lblAlisFiyati.TabIndex = 18;
            this.lblAlisFiyati.Text = "label3";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(925, 269);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(95, 25);
            this.label3.TabIndex = 19;
            this.label3.Text = "Alış Fiyatı";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(677, 269);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(88, 25);
            this.label8.TabIndex = 20;
            this.label8.Text = "Ürün Adı";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(473, 269);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(85, 25);
            this.label9.TabIndex = 21;
            this.label9.Text = "Kategori";
            // 
            // cmbPlatform
            // 
            this.cmbPlatform.FormattingEnabled = true;
            this.cmbPlatform.Location = new System.Drawing.Point(395, 420);
            this.cmbPlatform.Name = "cmbPlatform";
            this.cmbPlatform.Size = new System.Drawing.Size(178, 32);
            this.cmbPlatform.TabIndex = 22;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(449, 386);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(83, 25);
            this.label10.TabIndex = 23;
            this.label10.Text = "Platform";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(221, 544);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(65, 25);
            this.label11.TabIndex = 24;
            this.label11.Text = "Kargo";
            // 
            // cmbKargoFirmasi
            // 
            this.cmbKargoFirmasi.FormattingEnabled = true;
            this.cmbKargoFirmasi.Location = new System.Drawing.Point(135, 584);
            this.cmbKargoFirmasi.Name = "cmbKargoFirmasi";
            this.cmbKargoFirmasi.Size = new System.Drawing.Size(218, 32);
            this.cmbKargoFirmasi.TabIndex = 25;
            // 
            // numDesi
            // 
            this.numDesi.Location = new System.Drawing.Point(438, 587);
            this.numDesi.Name = "numDesi";
            this.numDesi.Size = new System.Drawing.Size(120, 29);
            this.numDesi.TabIndex = 26;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(466, 547);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(51, 25);
            this.label12.TabIndex = 27;
            this.label12.Text = "Desi";
            // 
            // txtKargoBedeli
            // 
            this.txtKargoBedeli.BackColor = System.Drawing.SystemColors.ControlLight;
            this.txtKargoBedeli.Location = new System.Drawing.Point(641, 597);
            this.txtKargoBedeli.Name = "txtKargoBedeli";
            this.txtKargoBedeli.ReadOnly = true;
            this.txtKargoBedeli.Size = new System.Drawing.Size(119, 29);
            this.txtKargoBedeli.TabIndex = 28;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(641, 560);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(124, 25);
            this.label13.TabIndex = 29;
            this.label13.Text = "Kargo Bedeli";
            // 
            // txtStopaj
            // 
            this.txtStopaj.BackColor = System.Drawing.SystemColors.ControlDark;
            this.txtStopaj.Location = new System.Drawing.Point(1045, 598);
            this.txtStopaj.Name = "txtStopaj";
            this.txtStopaj.ReadOnly = true;
            this.txtStopaj.Size = new System.Drawing.Size(100, 29);
            this.txtStopaj.TabIndex = 30;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(1074, 560);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(68, 25);
            this.label14.TabIndex = 31;
            this.label14.Text = "Stopaj";
            // 
            // txtPlatformIndirimi
            // 
            this.txtPlatformIndirimi.Location = new System.Drawing.Point(843, 597);
            this.txtPlatformIndirimi.Name = "txtPlatformIndirimi";
            this.txtPlatformIndirimi.Size = new System.Drawing.Size(155, 29);
            this.txtPlatformIndirimi.TabIndex = 32;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(849, 560);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(149, 25);
            this.label15.TabIndex = 33;
            this.label15.Text = "Platform İndirimi";
            // 
            // txtNetKar
            // 
            this.txtNetKar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.857143F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtNetKar.ForeColor = System.Drawing.Color.Green;
            this.txtNetKar.Location = new System.Drawing.Point(1197, 423);
            this.txtNetKar.Name = "txtNetKar";
            this.txtNetKar.ReadOnly = true;
            this.txtNetKar.Size = new System.Drawing.Size(146, 34);
            this.txtNetKar.TabIndex = 34;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(1249, 377);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(43, 25);
            this.label16.TabIndex = 35;
            this.label16.Text = "Kar";
            // 
            // dgvGunlukSiparisler
            // 
            this.dgvGunlukSiparisler.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvGunlukSiparisler.Location = new System.Drawing.Point(39, 977);
            this.dgvGunlukSiparisler.Name = "dgvGunlukSiparisler";
            this.dgvGunlukSiparisler.RowHeadersWidth = 72;
            this.dgvGunlukSiparisler.RowTemplate.Height = 31;
            this.dgvGunlukSiparisler.Size = new System.Drawing.Size(1313, 234);
            this.dgvGunlukSiparisler.TabIndex = 36;
            // 
            // txtImei
            // 
            this.txtImei.Location = new System.Drawing.Point(1123, 303);
            this.txtImei.Name = "txtImei";
            this.txtImei.Size = new System.Drawing.Size(186, 29);
            this.txtImei.TabIndex = 37;
            this.txtImei.Visible = false;
            // 
            // lblImei
            // 
            this.lblImei.AutoSize = true;
            this.lblImei.Location = new System.Drawing.Point(1173, 253);
            this.lblImei.Name = "lblImei";
            this.lblImei.Size = new System.Drawing.Size(88, 25);
            this.lblImei.TabIndex = 38;
            this.lblImei.Text = "IMEI No:";
            this.lblImei.Visible = false;
            // 
            // dgvSepet
            // 
            this.dgvSepet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSepet.Location = new System.Drawing.Point(154, 775);
            this.dgvSepet.Name = "dgvSepet";
            this.dgvSepet.RowHeadersWidth = 72;
            this.dgvSepet.RowTemplate.Height = 31;
            this.dgvSepet.Size = new System.Drawing.Size(1125, 150);
            this.dgvSepet.TabIndex = 39;
            // 
            // btnSepeteEkle
            // 
            this.btnSepeteEkle.Location = new System.Drawing.Point(356, 655);
            this.btnSepeteEkle.Name = "btnSepeteEkle";
            this.btnSepeteEkle.Size = new System.Drawing.Size(176, 41);
            this.btnSepeteEkle.TabIndex = 40;
            this.btnSepeteEkle.Text = "Sepete Ekle";
            this.btnSepeteEkle.UseVisualStyleBackColor = true;
            // 
            // txtKomisyon
            // 
            this.txtKomisyon.Location = new System.Drawing.Point(137, 512);
            this.txtKomisyon.Name = "txtKomisyon";
            this.txtKomisyon.Size = new System.Drawing.Size(216, 29);
            this.txtKomisyon.TabIndex = 41;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(211, 474);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(99, 25);
            this.label17.TabIndex = 42;
            this.label17.Text = "Komisyon";
            // 
            // txtHizmetBedeli
            // 
            this.txtHizmetBedeli.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.txtHizmetBedeli.Location = new System.Drawing.Point(413, 512);
            this.txtHizmetBedeli.Name = "txtHizmetBedeli";
            this.txtHizmetBedeli.ReadOnly = true;
            this.txtHizmetBedeli.Size = new System.Drawing.Size(145, 29);
            this.txtHizmetBedeli.TabIndex = 43;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(418, 474);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(131, 25);
            this.label18.TabIndex = 44;
            this.label18.Text = "Hizmet Bedeli";
            // 
            // txtPosBedeli
            // 
            this.txtPosBedeli.Location = new System.Drawing.Point(623, 512);
            this.txtPosBedeli.Name = "txtPosBedeli";
            this.txtPosBedeli.Size = new System.Drawing.Size(181, 29);
            this.txtPosBedeli.TabIndex = 45;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(653, 474);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(105, 25);
            this.label19.TabIndex = 46;
            this.label19.Text = "Pos Bedeli";
            // 
            // txtGunlukSiparis
            // 
            this.txtGunlukSiparis.Location = new System.Drawing.Point(22, 108);
            this.txtGunlukSiparis.Name = "txtGunlukSiparis";
            this.txtGunlukSiparis.ReadOnly = true;
            this.txtGunlukSiparis.Size = new System.Drawing.Size(117, 29);
            this.txtGunlukSiparis.TabIndex = 47;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(17, 63);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(122, 25);
            this.label20.TabIndex = 48;
            this.label20.Text = "Sipariş Adeti";
            // 
            // txtGunlukKar
            // 
            this.txtGunlukKar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtGunlukKar.ForeColor = System.Drawing.Color.Green;
            this.txtGunlukKar.Location = new System.Drawing.Point(22, 226);
            this.txtGunlukKar.Name = "txtGunlukKar";
            this.txtGunlukKar.ReadOnly = true;
            this.txtGunlukKar.Size = new System.Drawing.Size(117, 31);
            this.txtGunlukKar.TabIndex = 49;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(22, 174);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(110, 25);
            this.label21.TabIndex = 50;
            this.label21.Text = "Günlük Kar";
            // 
            // SayfaSiparisEkle
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label21);
            this.Controls.Add(this.txtGunlukKar);
            this.Controls.Add(this.label20);
            this.Controls.Add(this.txtGunlukSiparis);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.txtPosBedeli);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.txtHizmetBedeli);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.txtKomisyon);
            this.Controls.Add(this.btnSepeteEkle);
            this.Controls.Add(this.dgvSepet);
            this.Controls.Add(this.lblImei);
            this.Controls.Add(this.txtImei);
            this.Controls.Add(this.dgvGunlukSiparisler);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.txtNetKar);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.txtPlatformIndirimi);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.txtStopaj);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.txtKargoBedeli);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.numDesi);
            this.Controls.Add(this.cmbKargoFirmasi);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.cmbPlatform);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblAlisFiyati);
            this.Controls.Add(this.btnSiparisKaydet);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cmbDurum);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.dtpSiparisTarihi);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtSatisFiyati);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.numAdet);
            this.Controls.Add(this.lblKategori);
            this.Controls.Add(this.lblUrunAdi);
            this.Controls.Add(this.btnUrunCagir);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtBarkod);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtSiparisNo);
            this.Name = "SayfaSiparisEkle";
            this.Size = new System.Drawing.Size(1392, 1247);
            this.Load += new System.EventHandler(this.SayfaSiparisEkle_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numAdet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDesi)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvGunlukSiparisler)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSepet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSiparisNo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtBarkod;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnUrunCagir;
        private System.Windows.Forms.Label lblUrunAdi;
        private System.Windows.Forms.Label lblKategori;
        private System.Windows.Forms.NumericUpDown numAdet;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtSatisFiyati;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker dtpSiparisTarihi;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbDurum;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnSiparisKaydet;
        private System.Windows.Forms.Label lblAlisFiyati;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbPlatform;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cmbKargoFirmasi;
        private System.Windows.Forms.NumericUpDown numDesi;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtKargoBedeli;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtStopaj;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txtPlatformIndirimi;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtNetKar;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.DataGridView dgvGunlukSiparisler;
        private System.Windows.Forms.TextBox txtImei;
        private System.Windows.Forms.Label lblImei;
        private System.Windows.Forms.DataGridView dgvSepet;
        private System.Windows.Forms.Button btnSepeteEkle;
        private System.Windows.Forms.TextBox txtKomisyon;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtHizmetBedeli;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox txtPosBedeli;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox txtGunlukSiparis;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox txtGunlukKar;
        private System.Windows.Forms.Label label21;
    }
}
