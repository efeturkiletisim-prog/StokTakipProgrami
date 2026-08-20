namespace StokTakipProgrami
{
    partial class SayfaStoklar
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
            this.txtStokAra = new System.Windows.Forms.TextBox();
            this.btnStokAra = new System.Windows.Forms.Button();
            this.btnYenile = new System.Windows.Forms.Button();
            this.dgvStoklar = new System.Windows.Forms.DataGridView();
            this.lblToplamAdet = new System.Windows.Forms.Label();
            this.lblToplamTutar = new System.Windows.Forms.Label();
            this.txtDuzenleAlisFiyati = new System.Windows.Forms.TextBox();
            this.numDuzenleAdet = new System.Windows.Forms.NumericUpDown();
            this.btnStokGuncelle = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnStokSil = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvStoklar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDuzenleAdet)).BeginInit();
            this.SuspendLayout();
            // 
            // txtStokAra
            // 
            this.txtStokAra.Location = new System.Drawing.Point(532, 411);
            this.txtStokAra.Name = "txtStokAra";
            this.txtStokAra.Size = new System.Drawing.Size(270, 29);
            this.txtStokAra.TabIndex = 0;
            // 
            // btnStokAra
            // 
            this.btnStokAra.Location = new System.Drawing.Point(680, 479);
            this.btnStokAra.Name = "btnStokAra";
            this.btnStokAra.Size = new System.Drawing.Size(122, 40);
            this.btnStokAra.TabIndex = 1;
            this.btnStokAra.Text = "Ara";
            this.btnStokAra.UseVisualStyleBackColor = true;
            // 
            // btnYenile
            // 
            this.btnYenile.Location = new System.Drawing.Point(498, 479);
            this.btnYenile.Name = "btnYenile";
            this.btnYenile.Size = new System.Drawing.Size(151, 40);
            this.btnYenile.TabIndex = 2;
            this.btnYenile.Text = "Listeyi Yenile";
            this.btnYenile.UseVisualStyleBackColor = true;
            // 
            // dgvStoklar
            // 
            this.dgvStoklar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvStoklar.Location = new System.Drawing.Point(155, 776);
            this.dgvStoklar.Name = "dgvStoklar";
            this.dgvStoklar.RowHeadersWidth = 72;
            this.dgvStoklar.RowTemplate.Height = 31;
            this.dgvStoklar.Size = new System.Drawing.Size(1221, 150);
            this.dgvStoklar.TabIndex = 3;
            // 
            // lblToplamAdet
            // 
            this.lblToplamAdet.AutoSize = true;
            this.lblToplamAdet.Location = new System.Drawing.Point(1182, 247);
            this.lblToplamAdet.Name = "lblToplamAdet";
            this.lblToplamAdet.Size = new System.Drawing.Size(173, 25);
            this.lblToplamAdet.TabIndex = 4;
            this.lblToplamAdet.Text = "Toplam Stok Adeti";
            // 
            // lblToplamTutar
            // 
            this.lblToplamTutar.AutoSize = true;
            this.lblToplamTutar.Location = new System.Drawing.Point(1182, 309);
            this.lblToplamTutar.Name = "lblToplamTutar";
            this.lblToplamTutar.Size = new System.Drawing.Size(194, 25);
            this.lblToplamTutar.TabIndex = 5;
            this.lblToplamTutar.Text = "Toplam Stok Maliyeti";
            // 
            // txtDuzenleAlisFiyati
            // 
            this.txtDuzenleAlisFiyati.Location = new System.Drawing.Point(303, 656);
            this.txtDuzenleAlisFiyati.Name = "txtDuzenleAlisFiyati";
            this.txtDuzenleAlisFiyati.Size = new System.Drawing.Size(169, 29);
            this.txtDuzenleAlisFiyati.TabIndex = 6;
            // 
            // numDuzenleAdet
            // 
            this.numDuzenleAdet.Location = new System.Drawing.Point(581, 656);
            this.numDuzenleAdet.Name = "numDuzenleAdet";
            this.numDuzenleAdet.Size = new System.Drawing.Size(120, 29);
            this.numDuzenleAdet.TabIndex = 7;
            // 
            // btnStokGuncelle
            // 
            this.btnStokGuncelle.Location = new System.Drawing.Point(342, 716);
            this.btnStokGuncelle.Name = "btnStokGuncelle";
            this.btnStokGuncelle.Size = new System.Drawing.Size(103, 34);
            this.btnStokGuncelle.TabIndex = 8;
            this.btnStokGuncelle.Text = "Güncelle";
            this.btnStokGuncelle.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(370, 625);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 25);
            this.label1.TabIndex = 9;
            this.label1.Text = "Fiyat";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(605, 612);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 25);
            this.label2.TabIndex = 10;
            this.label2.Text = "Adet";
            // 
            // btnStokSil
            // 
            this.btnStokSil.Location = new System.Drawing.Point(598, 716);
            this.btnStokSil.Name = "btnStokSil";
            this.btnStokSil.Size = new System.Drawing.Size(94, 34);
            this.btnStokSil.TabIndex = 11;
            this.btnStokSil.Text = "Sil";
            this.btnStokSil.UseVisualStyleBackColor = true;
            // 
            // SayfaStoklar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnStokSil);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnStokGuncelle);
            this.Controls.Add(this.numDuzenleAdet);
            this.Controls.Add(this.txtDuzenleAlisFiyati);
            this.Controls.Add(this.lblToplamTutar);
            this.Controls.Add(this.lblToplamAdet);
            this.Controls.Add(this.dgvStoklar);
            this.Controls.Add(this.btnYenile);
            this.Controls.Add(this.btnStokAra);
            this.Controls.Add(this.txtStokAra);
            this.Name = "SayfaStoklar";
            this.Size = new System.Drawing.Size(1451, 1324);
            this.Load += new System.EventHandler(this.SayfaStoklar_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvStoklar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDuzenleAdet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtStokAra;
        private System.Windows.Forms.Button btnStokAra;
        private System.Windows.Forms.Button btnYenile;
        private System.Windows.Forms.DataGridView dgvStoklar;
        private System.Windows.Forms.Label lblToplamAdet;
        private System.Windows.Forms.Label lblToplamTutar;
        private System.Windows.Forms.TextBox txtDuzenleAlisFiyati;
        private System.Windows.Forms.NumericUpDown numDuzenleAdet;
        private System.Windows.Forms.Button btnStokGuncelle;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnStokSil;
    }
}
