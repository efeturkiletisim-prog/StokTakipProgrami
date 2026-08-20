using System;
using System.Windows.Forms;

namespace StokTakipProgrami
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Programı Giriş Paneli ile başlatıyoruz
            Application.Run(new FormLogin());

            // --- PROGRAM KAPANIRKEN SESSİZCE OTOMATİK YEDEK ALINIR ---
            Veritabanı.OtomatikYedekle();
        }
    }
}