namespace StokTakipProgrami
{
    partial class SayfaAnaEkran
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
            this.btnKisayolSiparisEkle = new System.Windows.Forms.Button();
            this.btnKisayolSiparisSorgula = new System.Windows.Forms.Button();
            this.btnKisayolUrunEkle = new System.Windows.Forms.Button();
            this.btnKisayolStokSorgula = new System.Windows.Forms.Button();
            this.pnlOrta = new System.Windows.Forms.Panel();
            this.pnlOrta.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnKisayolSiparisEkle
            // 
            this.btnKisayolSiparisEkle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(204)))), ((int)(((byte)(113)))));
            this.btnKisayolSiparisEkle.FlatAppearance.BorderSize = 0;
            this.btnKisayolSiparisEkle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnKisayolSiparisEkle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnKisayolSiparisEkle.ForeColor = System.Drawing.Color.White;
            this.btnKisayolSiparisEkle.Location = new System.Drawing.Point(209, 174);
            this.btnKisayolSiparisEkle.Name = "btnKisayolSiparisEkle";
            this.btnKisayolSiparisEkle.Size = new System.Drawing.Size(200, 150);
            this.btnKisayolSiparisEkle.TabIndex = 0;
            this.btnKisayolSiparisEkle.Text = "Sipariş Ekle";
            this.btnKisayolSiparisEkle.UseVisualStyleBackColor = false;
            // 
            // btnKisayolSiparisSorgula
            // 
            this.btnKisayolSiparisSorgula.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(152)))), ((int)(((byte)(219)))));
            this.btnKisayolSiparisSorgula.FlatAppearance.BorderSize = 0;
            this.btnKisayolSiparisSorgula.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnKisayolSiparisSorgula.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnKisayolSiparisSorgula.ForeColor = System.Drawing.Color.White;
            this.btnKisayolSiparisSorgula.Location = new System.Drawing.Point(592, 174);
            this.btnKisayolSiparisSorgula.Name = "btnKisayolSiparisSorgula";
            this.btnKisayolSiparisSorgula.Size = new System.Drawing.Size(200, 150);
            this.btnKisayolSiparisSorgula.TabIndex = 1;
            this.btnKisayolSiparisSorgula.Text = "Sipariş Sorgula";
            this.btnKisayolSiparisSorgula.UseVisualStyleBackColor = false;
            // 
            // btnKisayolUrunEkle
            // 
            this.btnKisayolUrunEkle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(155)))), ((int)(((byte)(89)))), ((int)(((byte)(182)))));
            this.btnKisayolUrunEkle.FlatAppearance.BorderSize = 0;
            this.btnKisayolUrunEkle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnKisayolUrunEkle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnKisayolUrunEkle.ForeColor = System.Drawing.Color.White;
            this.btnKisayolUrunEkle.Location = new System.Drawing.Point(209, 385);
            this.btnKisayolUrunEkle.Name = "btnKisayolUrunEkle";
            this.btnKisayolUrunEkle.Size = new System.Drawing.Size(200, 150);
            this.btnKisayolUrunEkle.TabIndex = 2;
            this.btnKisayolUrunEkle.Text = "Ürün Ekle";
            this.btnKisayolUrunEkle.UseVisualStyleBackColor = false;
            // 
            // btnKisayolStokSorgula
            // 
            this.btnKisayolStokSorgula.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(126)))), ((int)(((byte)(34)))));
            this.btnKisayolStokSorgula.FlatAppearance.BorderSize = 0;
            this.btnKisayolStokSorgula.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnKisayolStokSorgula.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnKisayolStokSorgula.ForeColor = System.Drawing.Color.White;
            this.btnKisayolStokSorgula.Location = new System.Drawing.Point(592, 385);
            this.btnKisayolStokSorgula.Name = "btnKisayolStokSorgula";
            this.btnKisayolStokSorgula.Size = new System.Drawing.Size(200, 150);
            this.btnKisayolStokSorgula.TabIndex = 3;
            this.btnKisayolStokSorgula.Text = "Stok Sorgula";
            this.btnKisayolStokSorgula.UseVisualStyleBackColor = false;
            // 
            // pnlOrta
            // 
            this.pnlOrta.Controls.Add(this.btnKisayolSiparisSorgula);
            this.pnlOrta.Controls.Add(this.btnKisayolUrunEkle);
            this.pnlOrta.Controls.Add(this.btnKisayolSiparisEkle);
            this.pnlOrta.Controls.Add(this.btnKisayolStokSorgula);
            this.pnlOrta.Location = new System.Drawing.Point(297, 179);
            this.pnlOrta.Name = "pnlOrta";
            this.pnlOrta.Size = new System.Drawing.Size(969, 721);
            this.pnlOrta.TabIndex = 4;
            this.pnlOrta.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // SayfaAnaEkran
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(244)))), ((int)(((byte)(248)))));
            this.Controls.Add(this.pnlOrta);
            this.Name = "SayfaAnaEkran";
            this.Size = new System.Drawing.Size(1349, 1023);
            this.pnlOrta.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.Button btnKisayolSiparisEkle;
        public System.Windows.Forms.Button btnKisayolSiparisSorgula;
        public System.Windows.Forms.Button btnKisayolUrunEkle;
        public System.Windows.Forms.Button btnKisayolStokSorgula;
        private System.Windows.Forms.Panel pnlOrta;
    }
}
