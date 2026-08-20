using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace StokTakipProgrami
{
    public static class ExcelMotoru
    {
        // 🚀 HERHANGİ BİR TABLOYU SİYAH KUTUDAN EXCEL'E FIRLATAN JİLET GİBİ MOTOR
        public static void DataGridViewTablosunuDisaAktar(DataGridView dgv, string dosyaAdi)
        {
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("Dışa aktarılacak herhangi bir veri bulunamadı!", "Tablo Boş", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel CSV Dosyası (*.csv)|*.csv";
            sfd.Title = "Tabloyu Excel'e Kaydet";
            // Dosya adına otomatik olarak günün tarihini ve saatini mühürler
            sfd.FileName = dosyaAdi + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // UTF8 BOM kullanarak Türkçe karakterlerin (Ş, Ğ, Ç vb.) Excel'de bozuk çıkmasını %100 engelliyoruz
                    using (StreamWriter sw = new StreamWriter(sfd.FileName, false, new UTF8Encoding(true)))
                    {
                        // 1. SÜTUN BAŞLIKLARINI DİZ
                        for (int i = 0; i < dgv.Columns.Count; i++)
                        {
                            sw.Write(dgv.Columns[i].HeaderText);
                            // Türkçe Excel için sütun ayırıcı noktalı virgüldür (;)
                            if (i < dgv.Columns.Count - 1) sw.Write(";");
                        }
                        sw.WriteLine(); // Alt satıra in

                        // 2. İÇERİDEKİ VERİLERİ DİZ
                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            for (int j = 0; j < dgv.Columns.Count; j++)
                            {
                                string hucreVerisi = dgv.Rows[i].Cells[j].Value?.ToString() ?? "";

                                // Verinin içinde yanlışlıkla noktalı virgül veya enter varsa Excel'i bozmaması için temizliyoruz
                                hucreVerisi = hucreVerisi.Replace(";", ",").Replace("\n", " ").Replace("\r", "");

                                sw.Write(hucreVerisi);

                                if (j < dgv.Columns.Count - 1) sw.Write(";");
                            }
                            sw.WriteLine(); // Sonraki satıra geç
                        }
                    }
                    MessageBox.Show("Veriler başarıyla Excel formatında dışa aktarıldı!\n\nDosya Konumu:\n" + sfd.FileName, "Başarılı İhracat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Dışa aktarma sırasında kritik bir hata oluştu:\n" + ex.Message, "Sistem Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}