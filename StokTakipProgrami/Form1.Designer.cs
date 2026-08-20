namespace StokTakipProgrami
{
    partial class Form1
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

        #region Windows Form Designer üretilen kod

        /// <summary>
        /// Tasarımcı desteği için gerekli metot - bu metodun 
        ///içeriğini kod düzenleyici ile değiştirmeyin.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.pnlRaporlar = new System.Windows.Forms.Panel();
            this.btnRaporlarAna = new System.Windows.Forms.Button();
            this.pnlHesaplar = new System.Windows.Forms.Panel();
            this.btnPlatformOdemeleriAlt = new System.Windows.Forms.Button();
            this.btnSiparisOdemeleriAlt = new System.Windows.Forms.Button();
            this.btnHesaplarAna = new System.Windows.Forms.Button();
            this.pnlPlatformlar = new System.Windows.Forms.Panel();
            this.btnPlatformAyarlariAlt = new System.Windows.Forms.Button();
            this.btnPlatformEkleAlt = new System.Windows.Forms.Button();
            this.btnPlatformlarAna = new System.Windows.Forms.Button();
            this.pnlUrunler = new System.Windows.Forms.Panel();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.pnlSiparisler = new System.Windows.Forms.Panel();
            this.btnSiparislerAna = new System.Windows.Forms.Button();
            this.btnSiparisEkle = new System.Windows.Forms.Button();
            this.btnSiparisSorgula = new System.Windows.Forms.Button();
            this.btnAnaSayfa = new System.Windows.Forms.Button();
            this.pnlOrta = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.pnlHesaplar.SuspendLayout();
            this.pnlPlatformlar.SuspendLayout();
            this.pnlUrunler.SuspendLayout();
            this.pnlSiparisler.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(35)))), ((int)(((byte)(40)))), ((int)(((byte)(45)))));
            this.panel1.Controls.Add(this.pnlRaporlar);
            this.panel1.Controls.Add(this.btnRaporlarAna);
            this.panel1.Controls.Add(this.pnlHesaplar);
            this.panel1.Controls.Add(this.btnHesaplarAna);
            this.panel1.Controls.Add(this.pnlPlatformlar);
            this.panel1.Controls.Add(this.btnPlatformlarAna);
            this.panel1.Controls.Add(this.pnlSiparisler);
            this.panel1.Controls.Add(this.btnAnaSayfa);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(357, 1293);
            this.panel1.TabIndex = 0;
            // 
            // pnlRaporlar
            // 
            this.pnlRaporlar.Location = new System.Drawing.Point(15, 1004);
            this.pnlRaporlar.Name = "pnlRaporlar";
            this.pnlRaporlar.Size = new System.Drawing.Size(274, 223);
            this.pnlRaporlar.TabIndex = 10;
            // 
            // btnRaporlarAna
            // 
            this.btnRaporlarAna.FlatAppearance.BorderSize = 0;
            this.btnRaporlarAna.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRaporlarAna.Font = new System.Drawing.Font("Segoe UI", 9.857143F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnRaporlarAna.ForeColor = System.Drawing.Color.White;
            this.btnRaporlarAna.Location = new System.Drawing.Point(15, 940);
            this.btnRaporlarAna.Name = "btnRaporlarAna";
            this.btnRaporlarAna.Size = new System.Drawing.Size(271, 48);
            this.btnRaporlarAna.TabIndex = 9;
            this.btnRaporlarAna.Text = "Raporlar";
            this.btnRaporlarAna.UseVisualStyleBackColor = true;
            // 
            // pnlHesaplar
            // 
            this.pnlHesaplar.Controls.Add(this.btnPlatformOdemeleriAlt);
            this.pnlHesaplar.Controls.Add(this.btnSiparisOdemeleriAlt);
            this.pnlHesaplar.Location = new System.Drawing.Point(12, 738);
            this.pnlHesaplar.Name = "pnlHesaplar";
            this.pnlHesaplar.Size = new System.Drawing.Size(280, 181);
            this.pnlHesaplar.TabIndex = 8;
            // 
            // btnPlatformOdemeleriAlt
            // 
            this.btnPlatformOdemeleriAlt.Location = new System.Drawing.Point(3, 69);
            this.btnPlatformOdemeleriAlt.Name = "btnPlatformOdemeleriAlt";
            this.btnPlatformOdemeleriAlt.Size = new System.Drawing.Size(274, 43);
            this.btnPlatformOdemeleriAlt.TabIndex = 1;
            this.btnPlatformOdemeleriAlt.Text = "Platform Ödemeleri";
            this.btnPlatformOdemeleriAlt.UseVisualStyleBackColor = true;
            // 
            // btnSiparisOdemeleriAlt
            // 
            this.btnSiparisOdemeleriAlt.Location = new System.Drawing.Point(0, 14);
            this.btnSiparisOdemeleriAlt.Name = "btnSiparisOdemeleriAlt";
            this.btnSiparisOdemeleriAlt.Size = new System.Drawing.Size(274, 43);
            this.btnSiparisOdemeleriAlt.TabIndex = 0;
            this.btnSiparisOdemeleriAlt.Text = "Sipariş Ödemeleri";
            this.btnSiparisOdemeleriAlt.UseVisualStyleBackColor = true;
            // 
            // btnHesaplarAna
            // 
            this.btnHesaplarAna.FlatAppearance.BorderSize = 0;
            this.btnHesaplarAna.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnHesaplarAna.Font = new System.Drawing.Font("Segoe UI", 9.857143F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnHesaplarAna.ForeColor = System.Drawing.Color.White;
            this.btnHesaplarAna.Location = new System.Drawing.Point(7, 690);
            this.btnHesaplarAna.Name = "btnHesaplarAna";
            this.btnHesaplarAna.Size = new System.Drawing.Size(285, 42);
            this.btnHesaplarAna.TabIndex = 7;
            this.btnHesaplarAna.Text = "Hesaplar";
            this.btnHesaplarAna.UseVisualStyleBackColor = true;
            // 
            // pnlPlatformlar
            // 
            this.pnlPlatformlar.Controls.Add(this.btnPlatformAyarlariAlt);
            this.pnlPlatformlar.Controls.Add(this.btnPlatformEkleAlt);
            this.pnlPlatformlar.Location = new System.Drawing.Point(4, 486);
            this.pnlPlatformlar.Name = "pnlPlatformlar";
            this.pnlPlatformlar.Size = new System.Drawing.Size(297, 186);
            this.pnlPlatformlar.TabIndex = 6;
            // 
            // btnPlatformAyarlariAlt
            // 
            this.btnPlatformAyarlariAlt.Location = new System.Drawing.Point(3, 67);
            this.btnPlatformAyarlariAlt.Name = "btnPlatformAyarlariAlt";
            this.btnPlatformAyarlariAlt.Size = new System.Drawing.Size(298, 38);
            this.btnPlatformAyarlariAlt.TabIndex = 1;
            this.btnPlatformAyarlariAlt.Text = "Platform Ayarları";
            this.btnPlatformAyarlariAlt.UseVisualStyleBackColor = true;
            // 
            // btnPlatformEkleAlt
            // 
            this.btnPlatformEkleAlt.Location = new System.Drawing.Point(0, 23);
            this.btnPlatformEkleAlt.Name = "btnPlatformEkleAlt";
            this.btnPlatformEkleAlt.Size = new System.Drawing.Size(297, 38);
            this.btnPlatformEkleAlt.TabIndex = 0;
            this.btnPlatformEkleAlt.Text = "Platform Ekle";
            this.btnPlatformEkleAlt.UseVisualStyleBackColor = true;
            // 
            // btnPlatformlarAna
            // 
            this.btnPlatformlarAna.FlatAppearance.BorderSize = 0;
            this.btnPlatformlarAna.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPlatformlarAna.Font = new System.Drawing.Font("Segoe UI", 9.857143F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnPlatformlarAna.ForeColor = System.Drawing.Color.White;
            this.btnPlatformlarAna.Location = new System.Drawing.Point(3, 441);
            this.btnPlatformlarAna.Name = "btnPlatformlarAna";
            this.btnPlatformlarAna.Size = new System.Drawing.Size(295, 39);
            this.btnPlatformlarAna.TabIndex = 5;
            this.btnPlatformlarAna.Text = "Platformlar";
            this.btnPlatformlarAna.UseVisualStyleBackColor = true;
            // 
            // pnlUrunler
            // 
            this.pnlUrunler.Controls.Add(this.button4);
            this.pnlUrunler.Controls.Add(this.button3);
            this.pnlUrunler.Controls.Add(this.button2);
            this.pnlUrunler.Controls.Add(this.button1);
            this.pnlUrunler.Location = new System.Drawing.Point(15, 118);
            this.pnlUrunler.Name = "pnlUrunler";
            this.pnlUrunler.Size = new System.Drawing.Size(301, 210);
            this.pnlUrunler.TabIndex = 1;
            this.pnlUrunler.Visible = false;
            // 
            // button4
            // 
            this.button4.Dock = System.Windows.Forms.DockStyle.Top;
            this.button4.Location = new System.Drawing.Point(0, 116);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(301, 37);
            this.button4.TabIndex = 2;
            this.button4.Text = "Veri Ekle/Sil";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Dock = System.Windows.Forms.DockStyle.Top;
            this.button3.Location = new System.Drawing.Point(0, 79);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(301, 37);
            this.button3.TabIndex = 1;
            this.button3.Text = "Stoklar";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Dock = System.Windows.Forms.DockStyle.Top;
            this.button2.Location = new System.Drawing.Point(0, 42);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(301, 37);
            this.button2.TabIndex = 0;
            this.button2.Text = "Ürün Ekle";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Top;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Segoe UI", 9.857143F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(0, 0);
            this.button1.Margin = new System.Windows.Forms.Padding(6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(301, 42);
            this.button1.TabIndex = 0;
            this.button1.Text = "Ürünler";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pnlSiparisler
            // 
            this.pnlSiparisler.Controls.Add(this.btnSiparislerAna);
            this.pnlSiparisler.Controls.Add(this.btnSiparisEkle);
            this.pnlSiparisler.Controls.Add(this.btnSiparisSorgula);
            this.pnlSiparisler.Controls.Add(this.pnlUrunler);
            this.pnlSiparisler.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlSiparisler.Location = new System.Drawing.Point(0, 50);
            this.pnlSiparisler.Name = "pnlSiparisler";
            this.pnlSiparisler.Size = new System.Drawing.Size(357, 275);
            this.pnlSiparisler.TabIndex = 3;
            this.pnlSiparisler.Visible = false;
            // 
            // btnSiparislerAna
            // 
            this.btnSiparislerAna.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnSiparislerAna.FlatAppearance.BorderSize = 0;
            this.btnSiparislerAna.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSiparislerAna.Font = new System.Drawing.Font("Segoe UI", 9.857143F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnSiparislerAna.ForeColor = System.Drawing.Color.White;
            this.btnSiparislerAna.Location = new System.Drawing.Point(0, 70);
            this.btnSiparislerAna.Name = "btnSiparislerAna";
            this.btnSiparislerAna.Size = new System.Drawing.Size(357, 42);
            this.btnSiparislerAna.TabIndex = 2;
            this.btnSiparislerAna.Text = "Siparişler";
            this.btnSiparislerAna.UseVisualStyleBackColor = true;
            this.btnSiparislerAna.Click += new System.EventHandler(this.btnSiparislerAna_Click);
            // 
            // btnSiparisEkle
            // 
            this.btnSiparisEkle.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnSiparisEkle.Location = new System.Drawing.Point(0, 35);
            this.btnSiparisEkle.Name = "btnSiparisEkle";
            this.btnSiparisEkle.Size = new System.Drawing.Size(357, 35);
            this.btnSiparisEkle.TabIndex = 0;
            this.btnSiparisEkle.Text = "Sipariş Ekle";
            this.btnSiparisEkle.UseVisualStyleBackColor = true;
            this.btnSiparisEkle.Click += new System.EventHandler(this.btnSiparisEkle_Click);
            // 
            // btnSiparisSorgula
            // 
            this.btnSiparisSorgula.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnSiparisSorgula.Location = new System.Drawing.Point(0, 0);
            this.btnSiparisSorgula.Name = "btnSiparisSorgula";
            this.btnSiparisSorgula.Size = new System.Drawing.Size(357, 35);
            this.btnSiparisSorgula.TabIndex = 1;
            this.btnSiparisSorgula.Text = "Sipariş Sorgula";
            this.btnSiparisSorgula.UseVisualStyleBackColor = true;
            this.btnSiparisSorgula.Click += new System.EventHandler(this.btnSiparisSorgula_Click);
            // 
            // btnAnaSayfa
            // 
            this.btnAnaSayfa.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnAnaSayfa.Location = new System.Drawing.Point(0, 0);
            this.btnAnaSayfa.Name = "btnAnaSayfa";
            this.btnAnaSayfa.Size = new System.Drawing.Size(357, 50);
            this.btnAnaSayfa.TabIndex = 4;
            this.btnAnaSayfa.Text = "ANA SAYFA";
            this.btnAnaSayfa.UseVisualStyleBackColor = true;
            // 
            // pnlOrta
            // 
            this.pnlOrta.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlOrta.Location = new System.Drawing.Point(357, 0);
            this.pnlOrta.Margin = new System.Windows.Forms.Padding(6);
            this.pnlOrta.Name = "pnlOrta";
            this.pnlOrta.Size = new System.Drawing.Size(1757, 1293);
            this.pnlOrta.TabIndex = 1;
            this.pnlOrta.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlOrta_Paint);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(2114, 1293);
            this.Controls.Add(this.pnlOrta);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "Form1";
            this.Text = "Form1";
            this.panel1.ResumeLayout(false);
            this.pnlHesaplar.ResumeLayout(false);
            this.pnlPlatformlar.ResumeLayout(false);
            this.pnlUrunler.ResumeLayout(false);
            this.pnlSiparisler.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel pnlOrta;
        private System.Windows.Forms.Panel pnlUrunler;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel pnlSiparisler;
        private System.Windows.Forms.Button btnSiparisSorgula;
        private System.Windows.Forms.Button btnSiparisEkle;
        private System.Windows.Forms.Button btnSiparislerAna;
        private System.Windows.Forms.Button btnAnaSayfa;
        private System.Windows.Forms.Button btnPlatformlarAna;
        private System.Windows.Forms.Panel pnlPlatformlar;
        private System.Windows.Forms.Button btnPlatformAyarlariAlt;
        private System.Windows.Forms.Button btnPlatformEkleAlt;
        private System.Windows.Forms.Panel pnlHesaplar;
        private System.Windows.Forms.Button btnPlatformOdemeleriAlt;
        private System.Windows.Forms.Button btnSiparisOdemeleriAlt;
        private System.Windows.Forms.Button btnHesaplarAna;
        private System.Windows.Forms.Panel pnlRaporlar;
        private System.Windows.Forms.Button btnRaporlarAna;
    }
}

