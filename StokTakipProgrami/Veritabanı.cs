using System;
using System.IO;
using Microsoft.Data.Sqlite;

namespace StokTakipProgrami
{
    public static class Veritabanı
    {
        public static string BaglantiYolu = "Data Source=stok.db";
        public static string AktifKullanici = "Admin";
        public static string AktifRol = "ADMIN"; // ADMIN veya PERSONEL

        public static void Kur()
        {
            using (var b = new SqliteConnection(BaglantiYolu))
            {
                b.Open();

                // 1. ÜRÜNLER TABLOSU
                string qUrunler = @"CREATE TABLE IF NOT EXISTS Urunler (
                                    Barkod TEXT PRIMARY KEY,
                                    UrunAdi TEXT,
                                    Kategori TEXT);";
                using (var cmd = new SqliteCommand(qUrunler, b)) cmd.ExecuteNonQuery();

                // 2. STOKLAR TABLOSU
                string qStoklar = @"CREATE TABLE IF NOT EXISTS Stoklar (
                                    Barkod TEXT PRIMARY KEY,
                                    Adet INTEGER,
                                    AlisFiyati REAL,
                                    Imei TEXT,
                                    KritikLimit INTEGER DEFAULT 0);";
                using (var cmd = new SqliteCommand(qStoklar, b)) cmd.ExecuteNonQuery();

                // 3. PLATFORMLAR TABLOSU
                string qPlatformlar = @"CREATE TABLE IF NOT EXISTS Platformlar (
                                        Ad TEXT PRIMARY KEY,
                                        HizmetBedeli REAL DEFAULT 0);";
                using (var cmd = new SqliteCommand(qPlatformlar, b)) cmd.ExecuteNonQuery();

                // 4. SİPARİŞLER TABLOSU
                string qSiparisler = @"CREATE TABLE IF NOT EXISTS Siparisler (
                                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                        Barkod TEXT,
                                        SiparisNo TEXT,
                                        Adet INTEGER,
                                        SatisFiyati REAL,
                                        SiparisTarihi TEXT,
                                        Durum TEXT DEFAULT 'BEKLEYEN',
                                        PlatformAd TEXT,
                                        SatilanImei TEXT,
                                        OAnkiAlisFiyati REAL DEFAULT 0,
                                        OAnkiKargoMaliyeti REAL DEFAULT 0);";
                using (var cmd = new SqliteCommand(qSiparisler, b)) cmd.ExecuteNonQuery();

                // Ensure backward-compatible column 'Tarih' exists because other modules insert into that column
                SutunEkle(b, "Siparisler", "Tarih", "TEXT");

                // 5. PERSONEL / KULLANICI TABLOSU
                string qKullanicilar = @"CREATE TABLE IF NOT EXISTS Kullanicilar (
                                            KullaniciAdi TEXT PRIMARY KEY,
                                            Sifre TEXT,
                                            Rol TEXT DEFAULT 'PERSONEL');";
                using (var cmd = new SqliteCommand(qKullanicilar, b)) cmd.ExecuteNonQuery();

                // 6. ÇOK KATMANLI KARGO ANLAŞMA MATRİSİ
                string qKargoMatrisi = @"CREATE TABLE IF NOT EXISTS KargoMatrisi (
                                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                            PlatformAd TEXT,
                                            KargoFirma TEXT,
                                            Desi INTEGER,
                                            Fiyat REAL,
                                            UNIQUE(PlatformAd, KargoFirma, Desi));";
                using (var cmd = new SqliteCommand(qKargoMatrisi, b)) cmd.ExecuteNonQuery();

                // 7. 📜 STOK HAREKETLERİ KARA KAPLI DEFTER TABLOSU (YENİ!)
                string qStokHareketleri = @"CREATE TABLE IF NOT EXISTS StokHareketleri (
                                                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                Barkod TEXT,
                                                IslemTuru TEXT,  -- 'GİRİŞ', 'ÇIKIŞ', 'İADE', 'DÜZELTME'
                                                Miktar INTEGER,
                                                IslemYapan TEXT,
                                                Tarih TEXT,
                                                Aciklama TEXT);";
                using (var cmd = new SqliteCommand(qStokHareketleri, b)) cmd.ExecuteNonQuery();

                // Eski tablolara yeni sütun güvenlik korumaları
                SutunEkle(b, "Stoklar", "KritikLimit", "INTEGER DEFAULT 0");
                SutunEkle(b, "Siparisler", "OAnkiKargoMaliyeti", "REAL DEFAULT 0");

                // Varsayılan kullanıcılar
                string qAdminEkle = "INSERT INTO Kullanicilar (KullaniciAdi, Sifre, Rol) VALUES ('admin', '1234', 'ADMIN') ON CONFLICT(KullaniciAdi) DO NOTHING";
                using (var cmd = new SqliteCommand(qAdminEkle, b)) cmd.ExecuteNonQuery();

                string qPersonelEkle = "INSERT INTO Kullanicilar (KullaniciAdi, Sifre, Rol) VALUES ('personel', '1234', 'PERSONEL') ON CONFLICT(KullaniciAdi) DO NOTHING";
                using (var cmd = new SqliteCommand(qPersonelEkle, b)) cmd.ExecuteNonQuery();
            }
        }

        private static void SutunEkle(SqliteConnection baglanti, string tabloAd, string sutunAd, string sutunTip)
        {
            try
            {
                string sorgu = $"ALTER TABLE {tabloAd} ADD COLUMN {sutunAd} {sutunTip};";
                using (var cmd = new SqliteCommand(sorgu, baglanti)) cmd.ExecuteNonQuery();
            }
            catch { }
        }

        // --- 📜 HER YERDEN AYAK İZİ BIRAKAN AKILLI DEFTER MOTORU ---
        // Bu metot projenin neresinden çağrılırsa çağrılsın o anki oturum sahibini otomatik loglar
        public static void HareketKaydet(string barkod, string islemTuru, int miktar, string aciklama, SqliteConnection hariciBaglanti = null, SqliteTransaction hariciTx = null)
        {
            string sorgu = @"INSERT INTO StokHareketleri (Barkod, IslemTuru, Miktar, IslemYapan, Tarih, Aciklama) 
                             VALUES (@b, @tur, @mik, @yapan, @tarih, @aciklama)";

            string tamZaman = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");

            try
            {
                // Eğer halihazırda açık bir transaction içinden çağrılıyorsa onu kullan (Kilitlenmeyi önler)
                if (hariciBaglanti != null && hariciTx != null)
                {
                    using (var cmd = new SqliteCommand(sorgu, hariciBaglanti, hariciTx))
                    {
                        cmd.Parameters.AddWithValue("@b", barkod); cmd.Parameters.AddWithValue("@tur", islemTuru.ToUpper());
                        cmd.Parameters.AddWithValue("@mik", miktar); cmd.Parameters.AddWithValue("@yapan", AktifKullanici);
                        cmd.Parameters.AddWithValue("@tarih", tamZaman); cmd.Parameters.AddWithValue("@aciklama", aciklama);
                        cmd.ExecuteNonQuery();
                    }

                    // Notify UI pages that a new hareket was added
                    try { SayfaStokHareketleri.OnStokHareketEklendi?.Invoke(); } catch { }
                }
                else // Bağımsız bir işlemse kendi bağlantısını açar kapatır
                {
                    using (var b = new SqliteConnection(BaglantiYolu))
                    {
                        b.Open();
                        using (var cmd = new SqliteCommand(sorgu, b))
                        {
                            cmd.Parameters.AddWithValue("@b", barkod); cmd.Parameters.AddWithValue("@tur", islemTuru.ToUpper());
                            cmd.Parameters.AddWithValue("@mik", miktar); cmd.Parameters.AddWithValue("@yapan", AktifKullanici);
                            cmd.Parameters.AddWithValue("@tarih", tamZaman); cmd.Parameters.AddWithValue("@aciklama", aciklama);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    // Notify UI pages that a new hareket was added
                    try { SayfaStokHareketleri.OnStokHareketEklendi?.Invoke(); } catch { }
                }
            }
            catch { /* Arka planda sessizce mühürler, ana akışı bozmaz */ }
        }

        // --- OTOMATİK TARİHLİ YEDEKLLEME MOTORU ---
        public static void OtomatikYedekle()
        {
            try
            {
                string kaynakDosya = "stok.db";
                if (!File.Exists(kaynakDosya)) return;

                string yedekKlasor = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Yedekler");
                if (!Directory.Exists(yedekKlasor)) Directory.CreateDirectory(yedekKlasor);

                string yedekDosyaAdi = $"Yedek_{DateTime.Now:yyyy_MM_dd_HHmmss}.db";
                string hedefYol = Path.Combine(yedekKlasor, yedekDosyaAdi);

                File.Copy(kaynakDosya, hedefYol, true);
            }
            catch { }
        }
    }
}