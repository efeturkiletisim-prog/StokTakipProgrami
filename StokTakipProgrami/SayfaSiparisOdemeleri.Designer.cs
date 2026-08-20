namespace StokTakipProgrami
{
    partial class SayfaSiparisOdemeleri
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
            this.txtSiparisNumaralari = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtGelenTutar = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbPlatformSec = new System.Windows.Forms.TextBox();
            this.dtpOdemeTarihi = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.btnOdemeKaydet = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.lblOkunanAdet = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtSiparisNumaralari
            // 
            this.txtSiparisNumaralari.Location = new System.Drawing.Point(139, 78);
            this.txtSiparisNumaralari.Multiline = true;
            this.txtSiparisNumaralari.Name = "txtSiparisNumaralari";
            this.txtSiparisNumaralari.Size = new System.Drawing.Size(276, 1172);
            this.txtSiparisNumaralari.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(203, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(171, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sipariş Numaraları";
            // 
            // txtGelenTutar
            // 
            this.txtGelenTutar.Location = new System.Drawing.Point(538, 496);
            this.txtGelenTutar.Name = "txtGelenTutar";
            this.txtGelenTutar.Size = new System.Drawing.Size(194, 29);
            this.txtGelenTutar.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(616, 454);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "Tutar";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(901, 453);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 25);
            this.label3.TabIndex = 4;
            this.label3.Text = "Platform";
            // 
            // cmbPlatformSec
            // 
            this.cmbPlatformSec.Location = new System.Drawing.Point(855, 496);
            this.cmbPlatformSec.Name = "cmbPlatformSec";
            this.cmbPlatformSec.Size = new System.Drawing.Size(219, 29);
            this.cmbPlatformSec.TabIndex = 5;
            // 
            // dtpOdemeTarihi
            // 
            this.dtpOdemeTarihi.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpOdemeTarihi.Location = new System.Drawing.Point(1187, 495);
            this.dtpOdemeTarihi.Name = "dtpOdemeTarihi";
            this.dtpOdemeTarihi.Size = new System.Drawing.Size(200, 29);
            this.dtpOdemeTarihi.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(1256, 435);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 25);
            this.label4.TabIndex = 7;
            this.label4.Text = "Tarih";
            // 
            // btnOdemeKaydet
            // 
            this.btnOdemeKaydet.Location = new System.Drawing.Point(735, 615);
            this.btnOdemeKaydet.Name = "btnOdemeKaydet";
            this.btnOdemeKaydet.Size = new System.Drawing.Size(132, 42);
            this.btnOdemeKaydet.TabIndex = 8;
            this.btnOdemeKaydet.Text = "Kaydet";
            this.btnOdemeKaydet.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(1538, 434);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 25);
            this.label5.TabIndex = 9;
            this.label5.Text = "Adet";
            // 
            // lblOkunanAdet
            // 
            this.lblOkunanAdet.AutoSize = true;
            this.lblOkunanAdet.Location = new System.Drawing.Point(1538, 495);
            this.lblOkunanAdet.Name = "lblOkunanAdet";
            this.lblOkunanAdet.Size = new System.Drawing.Size(64, 25);
            this.lblOkunanAdet.TabIndex = 10;
            this.lblOkunanAdet.Text = "label6";
            // 
            // SayfaSiparisOdemeleri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblOkunanAdet);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnOdemeKaydet);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.dtpOdemeTarihi);
            this.Controls.Add(this.cmbPlatformSec);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtGelenTutar);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtSiparisNumaralari);
            this.Name = "SayfaSiparisOdemeleri";
            this.Size = new System.Drawing.Size(1790, 1305);
            this.Load += new System.EventHandler(this.SayfaSiparisOdemeleri_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSiparisNumaralari;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtGelenTutar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox cmbPlatformSec;
        private System.Windows.Forms.DateTimePicker dtpOdemeTarihi;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnOdemeKaydet;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblOkunanAdet;
    }
}
