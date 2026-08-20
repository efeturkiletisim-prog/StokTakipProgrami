namespace StokTakipProgrami
{
    partial class SayfaUrunEkle
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
            this.txtBarkod = new System.Windows.Forms.TextBox();
            this.btnBarkodBul = new System.Windows.Forms.Button();
            this.lblKategori = new System.Windows.Forms.Label();
            this.lblMarka = new System.Windows.Forms.Label();
            this.lblUrunAdi = new System.Windows.Forms.Label();
            this.cmbFirma = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpAlisTarihi = new System.Windows.Forms.DateTimePicker();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.txtAlisFiyati = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtImei = new System.Windows.Forms.TextBox();
            this.lblImei = new System.Windows.Forms.Label();
            this.btnStokKaydet = new System.Windows.Forms.Button();
            this.dgvStoklar = new System.Windows.Forms.DataGridView();
            this.label4 = new System.Windows.Forms.Label();
            this.numAdet = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvStoklar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numAdet)).BeginInit();
            this.SuspendLayout();
            // 
            // txtBarkod
            // 
            this.txtBarkod.Location = new System.Drawing.Point(756, 106);
            this.txtBarkod.Name = "txtBarkod";
            this.txtBarkod.Size = new System.Drawing.Size(174, 29);
            this.txtBarkod.TabIndex = 0;
            // 
            // btnBarkodBul
            // 
            this.btnBarkodBul.Location = new System.Drawing.Point(756, 161);
            this.btnBarkodBul.Name = "btnBarkodBul";
            this.btnBarkodBul.Size = new System.Drawing.Size(174, 36);
            this.btnBarkodBul.TabIndex = 1;
            this.btnBarkodBul.Text = "Ürün Çağır";
            this.btnBarkodBul.UseVisualStyleBackColor = true;
            // 
            // lblKategori
            // 
            this.lblKategori.AutoSize = true;
            this.lblKategori.Location = new System.Drawing.Point(609, 225);
            this.lblKategori.Name = "lblKategori";
            this.lblKategori.Size = new System.Drawing.Size(64, 25);
            this.lblKategori.TabIndex = 2;
            this.lblKategori.Text = "label1";
            // 
            // lblMarka
            // 
            this.lblMarka.AutoSize = true;
            this.lblMarka.Location = new System.Drawing.Point(780, 225);
            this.lblMarka.Name = "lblMarka";
            this.lblMarka.Size = new System.Drawing.Size(64, 25);
            this.lblMarka.TabIndex = 3;
            this.lblMarka.Text = "label2";
            // 
            // lblUrunAdi
            // 
            this.lblUrunAdi.AutoSize = true;
            this.lblUrunAdi.Location = new System.Drawing.Point(963, 225);
            this.lblUrunAdi.Name = "lblUrunAdi";
            this.lblUrunAdi.Size = new System.Drawing.Size(64, 25);
            this.lblUrunAdi.TabIndex = 4;
            this.lblUrunAdi.Text = "label3";
            // 
            // cmbFirma
            // 
            this.cmbFirma.FormattingEnabled = true;
            this.cmbFirma.Location = new System.Drawing.Point(734, 286);
            this.cmbFirma.Name = "cmbFirma";
            this.cmbFirma.Size = new System.Drawing.Size(186, 32);
            this.cmbFirma.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(615, 289);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 25);
            this.label1.TabIndex = 6;
            this.label1.Text = "Firma";
            // 
            // dtpAlisTarihi
            // 
            this.dtpAlisTarihi.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpAlisTarihi.Location = new System.Drawing.Point(734, 370);
            this.dtpAlisTarihi.Name = "dtpAlisTarihi";
            this.dtpAlisTarihi.Size = new System.Drawing.Size(138, 29);
            this.dtpAlisTarihi.TabIndex = 7;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(798, 383);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 72;
            this.dataGridView1.RowTemplate.Height = 31;
            this.dataGridView1.Size = new System.Drawing.Size(8, 8);
            this.dataGridView1.TabIndex = 8;
            // 
            // txtAlisFiyati
            // 
            this.txtAlisFiyati.Location = new System.Drawing.Point(734, 415);
            this.txtAlisFiyati.Name = "txtAlisFiyati";
            this.txtAlisFiyati.Size = new System.Drawing.Size(138, 29);
            this.txtAlisFiyati.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(615, 370);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 25);
            this.label2.TabIndex = 10;
            this.label2.Text = "Tarih";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(615, 419);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 25);
            this.label3.TabIndex = 11;
            this.label3.Text = "Fiyat";
            // 
            // txtImei
            // 
            this.txtImei.Location = new System.Drawing.Point(1113, 286);
            this.txtImei.Multiline = true;
            this.txtImei.Name = "txtImei";
            this.txtImei.Size = new System.Drawing.Size(186, 322);
            this.txtImei.TabIndex = 12;
            // 
            // lblImei
            // 
            this.lblImei.AutoSize = true;
            this.lblImei.Location = new System.Drawing.Point(994, 286);
            this.lblImei.Name = "lblImei";
            this.lblImei.Size = new System.Drawing.Size(88, 25);
            this.lblImei.TabIndex = 13;
            this.lblImei.Text = "IMEI No:";
            // 
            // btnStokKaydet
            // 
            this.btnStokKaydet.Location = new System.Drawing.Point(757, 533);
            this.btnStokKaydet.Name = "btnStokKaydet";
            this.btnStokKaydet.Size = new System.Drawing.Size(138, 36);
            this.btnStokKaydet.TabIndex = 14;
            this.btnStokKaydet.Text = "Kaydet";
            this.btnStokKaydet.UseVisualStyleBackColor = true;
            // 
            // dgvStoklar
            // 
            this.dgvStoklar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvStoklar.Location = new System.Drawing.Point(163, 646);
            this.dgvStoklar.Name = "dgvStoklar";
            this.dgvStoklar.RowHeadersWidth = 72;
            this.dgvStoklar.RowTemplate.Height = 31;
            this.dgvStoklar.Size = new System.Drawing.Size(1385, 150);
            this.dgvStoklar.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(620, 328);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 25);
            this.label4.TabIndex = 16;
            this.label4.Text = "Adet:";
            // 
            // numAdet
            // 
            this.numAdet.Location = new System.Drawing.Point(744, 328);
            this.numAdet.Name = "numAdet";
            this.numAdet.Size = new System.Drawing.Size(120, 29);
            this.numAdet.TabIndex = 17;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(808, 63);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(74, 25);
            this.label5.TabIndex = 18;
            this.label5.Text = "Barkod";
            // 
            // SayfaUrunEkle
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label5);
            this.Controls.Add(this.numAdet);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.dgvStoklar);
            this.Controls.Add(this.btnStokKaydet);
            this.Controls.Add(this.lblImei);
            this.Controls.Add(this.txtImei);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtAlisFiyati);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.dtpAlisTarihi);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbFirma);
            this.Controls.Add(this.lblUrunAdi);
            this.Controls.Add(this.lblMarka);
            this.Controls.Add(this.lblKategori);
            this.Controls.Add(this.btnBarkodBul);
            this.Controls.Add(this.txtBarkod);
            this.Name = "SayfaUrunEkle";
            this.Size = new System.Drawing.Size(1721, 1176);
            this.Load += new System.EventHandler(this.SayfaUrunEkle_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvStoklar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numAdet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtBarkod;
        private System.Windows.Forms.Button btnBarkodBul;
        private System.Windows.Forms.Label lblKategori;
        private System.Windows.Forms.Label lblMarka;
        private System.Windows.Forms.Label lblUrunAdi;
        private System.Windows.Forms.ComboBox cmbFirma;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpAlisTarihi;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox txtAlisFiyati;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtImei;
        private System.Windows.Forms.Label lblImei;
        private System.Windows.Forms.Button btnStokKaydet;
        private System.Windows.Forms.DataGridView dgvStoklar;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numAdet;
        private System.Windows.Forms.Label label5;
    }
}
