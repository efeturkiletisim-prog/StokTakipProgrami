namespace StokTakipProgrami
{
    partial class SayfaSiparisSorgula
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
            this.txtSiparisAra = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSiparisAra = new System.Windows.Forms.Button();
            this.btnYenile = new System.Windows.Forms.Button();
            this.dgvSiparisler = new System.Windows.Forms.DataGridView();
            this.cmbDurumGuncelle = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnDurumGuncelle = new System.Windows.Forms.Button();
            this.btnSiparisSil = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSiparisler)).BeginInit();
            this.SuspendLayout();
            // 
            // txtSiparisAra
            // 
            this.txtSiparisAra.Location = new System.Drawing.Point(602, 362);
            this.txtSiparisAra.Name = "txtSiparisAra";
            this.txtSiparisAra.Size = new System.Drawing.Size(237, 29);
            this.txtSiparisAra.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(675, 321);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Sipariş ara";
            // 
            // btnSiparisAra
            // 
            this.btnSiparisAra.Location = new System.Drawing.Point(620, 416);
            this.btnSiparisAra.Name = "btnSiparisAra";
            this.btnSiparisAra.Size = new System.Drawing.Size(75, 49);
            this.btnSiparisAra.TabIndex = 2;
            this.btnSiparisAra.Text = "Ara";
            this.btnSiparisAra.UseVisualStyleBackColor = true;
            // 
            // btnYenile
            // 
            this.btnYenile.Location = new System.Drawing.Point(736, 416);
            this.btnYenile.Name = "btnYenile";
            this.btnYenile.Size = new System.Drawing.Size(80, 49);
            this.btnYenile.TabIndex = 3;
            this.btnYenile.Text = "Yenile";
            this.btnYenile.UseVisualStyleBackColor = true;
            // 
            // dgvSiparisler
            // 
            this.dgvSiparisler.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSiparisler.Location = new System.Drawing.Point(26, 748);
            this.dgvSiparisler.Name = "dgvSiparisler";
            this.dgvSiparisler.RowHeadersWidth = 72;
            this.dgvSiparisler.RowTemplate.Height = 31;
            this.dgvSiparisler.Size = new System.Drawing.Size(1533, 449);
            this.dgvSiparisler.TabIndex = 4;
            // 
            // cmbDurumGuncelle
            // 
            this.cmbDurumGuncelle.FormattingEnabled = true;
            this.cmbDurumGuncelle.Items.AddRange(new object[] {
            "Bekliyor",
            "Tamamlandı",
            "İptal Edildi"});
            this.cmbDurumGuncelle.Location = new System.Drawing.Point(661, 1203);
            this.cmbDurumGuncelle.Name = "cmbDurumGuncelle";
            this.cmbDurumGuncelle.Size = new System.Drawing.Size(251, 32);
            this.cmbDurumGuncelle.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(542, 1210);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 25);
            this.label2.TabIndex = 6;
            this.label2.Text = "Durum";
            // 
            // btnDurumGuncelle
            // 
            this.btnDurumGuncelle.Location = new System.Drawing.Point(588, 1241);
            this.btnDurumGuncelle.Name = "btnDurumGuncelle";
            this.btnDurumGuncelle.Size = new System.Drawing.Size(431, 48);
            this.btnDurumGuncelle.TabIndex = 7;
            this.btnDurumGuncelle.Text = "Seçili Siparişin Durumunu Güncelle";
            this.btnDurumGuncelle.UseVisualStyleBackColor = true;
            // 
            // btnSiparisSil
            // 
            this.btnSiparisSil.Location = new System.Drawing.Point(1282, 1228);
            this.btnSiparisSil.Name = "btnSiparisSil";
            this.btnSiparisSil.Size = new System.Drawing.Size(164, 47);
            this.btnSiparisSil.TabIndex = 8;
            this.btnSiparisSil.Text = "Siparişi Sil";
            this.btnSiparisSil.UseVisualStyleBackColor = true;
            // 
            // SayfaSiparisSorgula
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnSiparisSil);
            this.Controls.Add(this.btnDurumGuncelle);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbDurumGuncelle);
            this.Controls.Add(this.dgvSiparisler);
            this.Controls.Add(this.btnYenile);
            this.Controls.Add(this.btnSiparisAra);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtSiparisAra);
            this.Name = "SayfaSiparisSorgula";
            this.Size = new System.Drawing.Size(1595, 1302);
            this.Load += new System.EventHandler(this.SayfaSiparisSorgula_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSiparisler)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSiparisAra;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSiparisAra;
        private System.Windows.Forms.Button btnYenile;
        private System.Windows.Forms.DataGridView dgvSiparisler;
        private System.Windows.Forms.ComboBox cmbDurumGuncelle;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnDurumGuncelle;
        private System.Windows.Forms.Button btnSiparisSil;
    }
}
