namespace StokTakipProgrami
{
    partial class SayfaKarZarar
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
            this.dtpBaslangic = new System.Windows.Forms.DateTimePicker();
            this.dtpBitis = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnHesapla = new System.Windows.Forms.Button();
            this.dgvDetayliRapor = new System.Windows.Forms.DataGridView();
            this.lblToplamCiro = new System.Windows.Forms.Label();
            this.lblToplamMaliyet = new System.Windows.Forms.Label();
            this.lblIadeZarar = new System.Windows.Forms.Label();
            this.lblPlatformGider = new System.Windows.Forms.Label();
            this.lblNetKar = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDetayliRapor)).BeginInit();
            this.SuspendLayout();
            // 
            // dtpBaslangic
            // 
            this.dtpBaslangic.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpBaslangic.Location = new System.Drawing.Point(590, 375);
            this.dtpBaslangic.Name = "dtpBaslangic";
            this.dtpBaslangic.Size = new System.Drawing.Size(200, 29);
            this.dtpBaslangic.TabIndex = 0;
            // 
            // dtpBitis
            // 
            this.dtpBitis.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpBitis.Location = new System.Drawing.Point(913, 375);
            this.dtpBitis.Name = "dtpBitis";
            this.dtpBitis.Size = new System.Drawing.Size(200, 29);
            this.dtpBitis.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(614, 322);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(151, 25);
            this.label1.TabIndex = 2;
            this.label1.Text = "Başlangıç Tarihi";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(965, 322);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(102, 25);
            this.label2.TabIndex = 3;
            this.label2.Text = "Bitiş Tarihi";
            // 
            // btnHesapla
            // 
            this.btnHesapla.Location = new System.Drawing.Point(789, 443);
            this.btnHesapla.Name = "btnHesapla";
            this.btnHesapla.Size = new System.Drawing.Size(126, 41);
            this.btnHesapla.TabIndex = 4;
            this.btnHesapla.Text = "Ara";
            this.btnHesapla.UseVisualStyleBackColor = true;
            // 
            // dgvDetayliRapor
            // 
            this.dgvDetayliRapor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDetayliRapor.Location = new System.Drawing.Point(277, 701);
            this.dgvDetayliRapor.Name = "dgvDetayliRapor";
            this.dgvDetayliRapor.RowHeadersWidth = 72;
            this.dgvDetayliRapor.RowTemplate.Height = 31;
            this.dgvDetayliRapor.Size = new System.Drawing.Size(1314, 306);
            this.dgvDetayliRapor.TabIndex = 5;
            // 
            // lblToplamCiro
            // 
            this.lblToplamCiro.AutoSize = true;
            this.lblToplamCiro.Location = new System.Drawing.Point(289, 593);
            this.lblToplamCiro.Name = "lblToplamCiro";
            this.lblToplamCiro.Size = new System.Drawing.Size(64, 25);
            this.lblToplamCiro.TabIndex = 6;
            this.lblToplamCiro.Text = "label3";
            // 
            // lblToplamMaliyet
            // 
            this.lblToplamMaliyet.AutoSize = true;
            this.lblToplamMaliyet.Location = new System.Drawing.Point(525, 593);
            this.lblToplamMaliyet.Name = "lblToplamMaliyet";
            this.lblToplamMaliyet.Size = new System.Drawing.Size(64, 25);
            this.lblToplamMaliyet.TabIndex = 7;
            this.lblToplamMaliyet.Text = "label4";
            // 
            // lblIadeZarar
            // 
            this.lblIadeZarar.AutoSize = true;
            this.lblIadeZarar.Location = new System.Drawing.Point(937, 593);
            this.lblIadeZarar.Name = "lblIadeZarar";
            this.lblIadeZarar.Size = new System.Drawing.Size(64, 25);
            this.lblIadeZarar.TabIndex = 9;
            this.lblIadeZarar.Text = "label5";
            // 
            // lblPlatformGider
            // 
            this.lblPlatformGider.AutoSize = true;
            this.lblPlatformGider.Location = new System.Drawing.Point(701, 593);
            this.lblPlatformGider.Name = "lblPlatformGider";
            this.lblPlatformGider.Size = new System.Drawing.Size(64, 25);
            this.lblPlatformGider.TabIndex = 8;
            this.lblPlatformGider.Text = "label6";
            // 
            // lblNetKar
            // 
            this.lblNetKar.AutoSize = true;
            this.lblNetKar.Location = new System.Drawing.Point(1136, 593);
            this.lblNetKar.Name = "lblNetKar";
            this.lblNetKar.Size = new System.Drawing.Size(64, 25);
            this.lblNetKar.TabIndex = 10;
            this.lblNetKar.Text = "label7";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(1136, 523);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(78, 25);
            this.label8.TabIndex = 15;
            this.label8.Text = "Net Kar";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(923, 523);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(126, 25);
            this.label9.TabIndex = 14;
            this.label9.Text = "İade Zararları";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(701, 523);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(160, 25);
            this.label10.TabIndex = 13;
            this.label10.Text = "Platform Giderleri";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(493, 523);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(145, 25);
            this.label11.TabIndex = 12;
            this.label11.Text = "Toplam Maliyet";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(289, 523);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(119, 25);
            this.label12.TabIndex = 11;
            this.label12.Text = "Toplam Ciro";
            // 
            // SayfaKarZarar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.lblNetKar);
            this.Controls.Add(this.lblIadeZarar);
            this.Controls.Add(this.lblPlatformGider);
            this.Controls.Add(this.lblToplamMaliyet);
            this.Controls.Add(this.lblToplamCiro);
            this.Controls.Add(this.dgvDetayliRapor);
            this.Controls.Add(this.btnHesapla);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dtpBitis);
            this.Controls.Add(this.dtpBaslangic);
            this.Name = "SayfaKarZarar";
            this.Size = new System.Drawing.Size(1733, 1358);
            this.Load += new System.EventHandler(this.SayfaKarZarar_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDetayliRapor)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtpBaslangic;
        private System.Windows.Forms.DateTimePicker dtpBitis;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnHesapla;
        private System.Windows.Forms.DataGridView dgvDetayliRapor;
        private System.Windows.Forms.Label lblToplamCiro;
        private System.Windows.Forms.Label lblToplamMaliyet;
        private System.Windows.Forms.Label lblIadeZarar;
        private System.Windows.Forms.Label lblPlatformGider;
        private System.Windows.Forms.Label lblNetKar;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
    }
}
