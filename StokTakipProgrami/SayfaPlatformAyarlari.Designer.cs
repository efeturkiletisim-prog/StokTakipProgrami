namespace StokTakipProgrami
{
    partial class SayfaPlatformAyarlari
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
            this.dgvPlatformlar = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPlatformAdi = new System.Windows.Forms.TextBox();
            this.txtHizmetBedeli = new System.Windows.Forms.TextBox();
            this.txtPosBedeli = new System.Windows.Forms.TextBox();
            this.txtEkBedel = new System.Windows.Forms.TextBox();
            this.btnPlatformGuncelle = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPlatformlar)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvPlatformlar
            // 
            this.dgvPlatformlar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPlatformlar.Location = new System.Drawing.Point(547, 841);
            this.dgvPlatformlar.Name = "dgvPlatformlar";
            this.dgvPlatformlar.RowHeadersWidth = 72;
            this.dgvPlatformlar.RowTemplate.Height = 31;
            this.dgvPlatformlar.Size = new System.Drawing.Size(716, 150);
            this.dgvPlatformlar.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(607, 396);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Platform";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(602, 476);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 25);
            this.label2.TabIndex = 2;
            this.label2.Text = "Hizmet Bedeli";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(602, 555);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(105, 25);
            this.label3.TabIndex = 3;
            this.label3.Text = "Pos Bedeli";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(607, 639);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(90, 25);
            this.label4.TabIndex = 4;
            this.label4.Text = "Ek Bedel";
            // 
            // txtPlatformAdi
            // 
            this.txtPlatformAdi.Location = new System.Drawing.Point(818, 392);
            this.txtPlatformAdi.Name = "txtPlatformAdi";
            this.txtPlatformAdi.Size = new System.Drawing.Size(184, 29);
            this.txtPlatformAdi.TabIndex = 5;
            // 
            // txtHizmetBedeli
            // 
            this.txtHizmetBedeli.Location = new System.Drawing.Point(818, 476);
            this.txtHizmetBedeli.Name = "txtHizmetBedeli";
            this.txtHizmetBedeli.Size = new System.Drawing.Size(184, 29);
            this.txtHizmetBedeli.TabIndex = 6;
            // 
            // txtPosBedeli
            // 
            this.txtPosBedeli.Location = new System.Drawing.Point(818, 555);
            this.txtPosBedeli.Name = "txtPosBedeli";
            this.txtPosBedeli.Size = new System.Drawing.Size(184, 29);
            this.txtPosBedeli.TabIndex = 7;
            // 
            // txtEkBedel
            // 
            this.txtEkBedel.Location = new System.Drawing.Point(818, 634);
            this.txtEkBedel.Name = "txtEkBedel";
            this.txtEkBedel.Size = new System.Drawing.Size(184, 29);
            this.txtEkBedel.TabIndex = 8;
            // 
            // btnPlatformGuncelle
            // 
            this.btnPlatformGuncelle.Location = new System.Drawing.Point(818, 723);
            this.btnPlatformGuncelle.Name = "btnPlatformGuncelle";
            this.btnPlatformGuncelle.Size = new System.Drawing.Size(183, 60);
            this.btnPlatformGuncelle.TabIndex = 9;
            this.btnPlatformGuncelle.Text = "Güncelle";
            this.btnPlatformGuncelle.UseVisualStyleBackColor = true;
            // 
            // SayfaPlatformAyarlari
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnPlatformGuncelle);
            this.Controls.Add(this.txtEkBedel);
            this.Controls.Add(this.txtPosBedeli);
            this.Controls.Add(this.txtHizmetBedeli);
            this.Controls.Add(this.txtPlatformAdi);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgvPlatformlar);
            this.Name = "SayfaPlatformAyarlari";
            this.Size = new System.Drawing.Size(1803, 1278);
            this.Load += new System.EventHandler(this.SayfaPlatformAyarlari_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPlatformlar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvPlatformlar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPlatformAdi;
        private System.Windows.Forms.TextBox txtHizmetBedeli;
        private System.Windows.Forms.TextBox txtPosBedeli;
        private System.Windows.Forms.TextBox txtEkBedel;
        private System.Windows.Forms.Button btnPlatformGuncelle;
    }
}
