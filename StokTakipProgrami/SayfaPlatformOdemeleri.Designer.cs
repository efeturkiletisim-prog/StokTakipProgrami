namespace StokTakipProgrami
{
    partial class SayfaPlatformOdemeleri
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
            this.cmbPlatformFiltre = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dgvOdemelerGecmisi = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOdemelerGecmisi)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbPlatformFiltre
            // 
            this.cmbPlatformFiltre.FormattingEnabled = true;
            this.cmbPlatformFiltre.Location = new System.Drawing.Point(775, 602);
            this.cmbPlatformFiltre.Name = "cmbPlatformFiltre";
            this.cmbPlatformFiltre.Size = new System.Drawing.Size(277, 32);
            this.cmbPlatformFiltre.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(881, 563);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Platform";
            // 
            // dgvOdemelerGecmisi
            // 
            this.dgvOdemelerGecmisi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvOdemelerGecmisi.Location = new System.Drawing.Point(268, 811);
            this.dgvOdemelerGecmisi.Name = "dgvOdemelerGecmisi";
            this.dgvOdemelerGecmisi.RowHeadersWidth = 72;
            this.dgvOdemelerGecmisi.RowTemplate.Height = 31;
            this.dgvOdemelerGecmisi.Size = new System.Drawing.Size(1256, 150);
            this.dgvOdemelerGecmisi.TabIndex = 2;
            // 
            // SayfaPlatformOdemeleri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dgvOdemelerGecmisi);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbPlatformFiltre);
            this.Name = "SayfaPlatformOdemeleri";
            this.Size = new System.Drawing.Size(1784, 1276);
            this.Load += new System.EventHandler(this.SayfaPlatformOdemeleri_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvOdemelerGecmisi)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbPlatformFiltre;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgvOdemelerGecmisi;
    }
}
