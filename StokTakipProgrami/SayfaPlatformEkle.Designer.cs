namespace StokTakipProgrami
{
    partial class SayfaPlatformEkle
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
            this.txtPlatformAdi = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtHizmetBedeli = new System.Windows.Forms.Label();
            this.txtPosBedeli = new System.Windows.Forms.Label();
            this.txtEkBedel = new System.Windows.Forms.Label();
            this.btnPlatformKaydet = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtPlatformAdi
            // 
            this.txtPlatformAdi.Location = new System.Drawing.Point(564, 465);
            this.txtPlatformAdi.Name = "txtPlatformAdi";
            this.txtPlatformAdi.Size = new System.Drawing.Size(240, 29);
            this.txtPlatformAdi.TabIndex = 0;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(564, 541);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(240, 29);
            this.textBox2.TabIndex = 1;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(564, 699);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(240, 29);
            this.textBox3.TabIndex = 3;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(564, 623);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(240, 29);
            this.textBox4.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(639, 428);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 25);
            this.label1.TabIndex = 4;
            this.label1.Text = "Platform Adı";
            // 
            // txtHizmetBedeli
            // 
            this.txtHizmetBedeli.AutoSize = true;
            this.txtHizmetBedeli.Location = new System.Drawing.Point(625, 513);
            this.txtHizmetBedeli.Name = "txtHizmetBedeli";
            this.txtHizmetBedeli.Size = new System.Drawing.Size(131, 25);
            this.txtHizmetBedeli.TabIndex = 5;
            this.txtHizmetBedeli.Text = "Hizmet Bedeli";
            // 
            // txtPosBedeli
            // 
            this.txtPosBedeli.AutoSize = true;
            this.txtPosBedeli.Location = new System.Drawing.Point(592, 595);
            this.txtPosBedeli.Name = "txtPosBedeli";
            this.txtPosBedeli.Size = new System.Drawing.Size(190, 25);
            this.txtPosBedeli.TabIndex = 6;
            this.txtPosBedeli.Text = "Pos/Güvence Bedeli";
            // 
            // txtEkBedel
            // 
            this.txtEkBedel.AutoSize = true;
            this.txtEkBedel.Location = new System.Drawing.Point(651, 671);
            this.txtEkBedel.Name = "txtEkBedel";
            this.txtEkBedel.Size = new System.Drawing.Size(90, 25);
            this.txtEkBedel.TabIndex = 7;
            this.txtEkBedel.Text = "Ek Bedel";
            // 
            // btnPlatformKaydet
            // 
            this.btnPlatformKaydet.Location = new System.Drawing.Point(564, 806);
            this.btnPlatformKaydet.Name = "btnPlatformKaydet";
            this.btnPlatformKaydet.Size = new System.Drawing.Size(240, 68);
            this.btnPlatformKaydet.TabIndex = 8;
            this.btnPlatformKaydet.Text = "Kaydet";
            this.btnPlatformKaydet.UseVisualStyleBackColor = true;
            // 
            // SayfaPlatformEkle
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnPlatformKaydet);
            this.Controls.Add(this.txtEkBedel);
            this.Controls.Add(this.txtPosBedeli);
            this.Controls.Add(this.txtHizmetBedeli);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.txtPlatformAdi);
            this.Name = "SayfaPlatformEkle";
            this.Size = new System.Drawing.Size(1444, 1306);
            this.Load += new System.EventHandler(this.SayfaPlatformEkle_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtPlatformAdi;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label txtHizmetBedeli;
        private System.Windows.Forms.Label txtPosBedeli;
        private System.Windows.Forms.Label txtEkBedel;
        private System.Windows.Forms.Button btnPlatformKaydet;
    }
}
