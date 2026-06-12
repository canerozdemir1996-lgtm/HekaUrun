# HekaUrun

Bu uygulama Python ile yazildi ve iki ana sekmeden olusur:

1. `Urun Sorgulama`
   Excel veya SQL veri kaynagindan urun ailesi, kirilim, stok kodu ve ozellikleri ceker. Aranabilir dropdown alanlariyla filtreleme yapar, genel arama yapabilir, tek kirilim veya tum aileyi vitrin gorunumunde listeleyebilir. Ayrica stok koduna gore gorsel bulur, kart icinde birden fazla urun gorseli arasinda gecis yapabilir ve ilgili kanal klasorlerini Explorer'da acabilir.

2. `Toplu Isim Degistirme`
   Belirttigin klasor altindaki `01_ Urun Gorselleri` klasorlerinde stok kodu ile baslayan urun klasorlerini bulur, gorselleri kurallarina gore yeniden adlandirir veya farkli bir cikis klasorune orijinallere dokunmadan cikartir. Filtrelenebilir, gorsel onizlemeli ve geri alinabilir bir akis kullanir. Sadece secili satirlar uygulanir.

3. `Yurtdisi Etiket Basligi`
   Ana ayarlarda secilen veri kaynagindan barkod/etiket bilgisini ceker. Stok kodu veya aciklama ile arama yapar, urun ve koli barkod bilgilerini gosterir, ornek yurtdisi etiket tasarimina uygun HEKA onizlemesi olusturur ve PDF olarak kaydeder. Secili urun icin etiket metinleri, barkod/HKA bilgisi ve temel tasarim konumlari elle duzenlenebilir.

## Calistirma

```powershell
python app.py
```

EXE paketi olusturmak icin:

```powershell
.\build_exe.ps1 -Clean
```

Programi `dist\UrunYonetimMasasi_v3\UrunYonetimMasasi_v3.exe` yolundan calistir. Baska bilgisayara tasirken sadece exe'yi degil, `dist\UrunYonetimMasasi_v3` klasorunun tamamini kopyala. `build` klasoru PyInstaller'in gecici calisma klasorudur; oradaki exe calistirilmaz.

Tasima icin varsayilan build yerel `settings.json` ve `product_index.sqlite` dosyalarini pakete koymaz. Yeni bilgisayarda ilk acilista `Ayarlar > Ayarlar` bolumunden Excel ve klasor yollarini o bilgisayara gore sec. Sadece ayni disk/ag harflerine sahip bilgisayara paket hazirliyorsan `.\build_exe.ps1 -Clean -WithLocalSettings` kullanabilirsin; indeks dosyasi mutlak yollar tuttugu icin baska bilgisayara tasirken `-WithIndex` kullanma.

macOS uygulama paketi olusturmak icin bu klasoru Mac'e kopyala ve Terminal'de:

```bash
chmod +x build_mac.sh
./build_mac.sh --clean
```

Mac build cikisi `dist/UrunYonetimMasasi_v3.app` olarak olusur. PyInstaller macOS `.app` paketini sadece macOS uzerinde uretebildigi icin Windows makineden dogrudan Mac build alinmaz. Ayni ag/disk yollarini kullanan bir Mac icin yerel ayarlari da pakete koymak istersen `./build_mac.sh --clean --with-local-settings --with-index` kullan. SQL Server ODBC ile baglanilacaksa Mac'e Microsoft ODBC Driver for SQL Server kurulmalidir.

GitHub Actions Mac build iki artifact uretir:

- `UrunYonetimMasasi_v3-macOS-arm64`: Apple Silicon / M1-M2-M3-M4 Mac'ler icin.
- `UrunYonetimMasasi_v3-macOS-intel-x64`: Intel islemcili Mac'ler icin.

Mac uygulamayi engellerse once sag tik > Open dene. Gerekirse Terminal'de `xattr -dr com.apple.quarantine UrunYonetimMasasi_v3.app` calistir. Cift tiklayinca hic tepki yoksa `UrunYonetimMasasi_v3.app/Contents/MacOS/UrunYonetimMasasi_v3 --smoke-test` komutu startup hatasini terminale/log dosyasina dusurur.

## Gerekli Kutuphaneler

```powershell
python -m pip install -r requirements.txt
```

## Ayarlar

Uygulamayi acinca menuden `Ayarlar > Ayarlar` yoluyla su alanlari doldurabilirsin:

- Veri kaynagi
- Excel dosyasi
- SQL baglanti
- SQL tablo veya gorunum
- SQL sorgu (opsiyonel)
- Sayfa adi
- Baslik satiri
- Urun ailesi sutunu
- Kirilim sutunu
- Stok kodu sutunu
- Ozellik sutunlari
- Ozellik baslik eslestirme
- Explorer arama ana klasoru
- Onizleme gorsel klasoru
- Gorsel uzantilari
- Stok kodu regex
- Acik / kapali / teknik cizim anahtar kelimeleri
- Son secimler ve pencere boyutu otomatik olarak `settings.json` dosyasina kaydedilir
- Eger uygulama klasorune yazma izni yoksa ayar ve indeks dosyalari Windows kullanici veri klasorunde saklanir.

## Notlar

- Excel tarafinda `xlsx` veya `xlsm` formatlari desteklenir.
- Varsayilan veri kaynagi `LOGODATA` SQL baglantisidir: `192.168.10.12 / LOGODATA / HRMS_122_MALZEMEBILGILERI`. Eski ayar dosyasi varsa ilk acilista bu SQL profiline otomatik tasinir.
- Guvenlik icin SQL parolasi kaynak koda yazilmaz. Parolayi uygulama ayarlarindan girebilir veya build/calistirma ortaminda `HEKA_LOGODATA_SQL_PASSWORD` degiskeniyle verebilirsin.
- SQL tarafinda SQLite dosya/URI veya ODBC baglanti metni desteklenir. ODBC/SQL Server icin paket `pyodbc` icerir; `SQL Sorgu` bos birakilirsa `SQL Tablo veya Gorunum` degeri kullanilir. Ekrandaki `sa` kullanicisi parolasiz baglanamazsa uygulama otomatik olarak Windows kimligiyle bir kez daha dener.
- `Ozellik Sutunlari` bos birakilirsa aile, kirilim ve stok disindaki tum sutunlar ozellik olarak gosterilir.
- Excel basliklari `11`, `12`, `13` gibi geliyorsa `Ozellik Baslik Eslestirme` alanina `11=Material`, `12=Fiyat` gibi satirlar yazabilirsin.
- Urun kartlarindaki `B2B`, `WEB`, `INSTAGRAM`, `BEYMEN` butonlari secili stok kodunu ilgili kanal klasoru altinda arayip Explorer'da acar.
- Urun kartlarinda `<` ve `>` tuslari ile ayni urune ait gorseller arasinda gezebilirsin.
- Urun vitrin kartlarinda `Stok Kopyala`, `Karti Kopyala` ve `Yollari Kopyala` butonlariyla stok kodu, kart ozeti veya bulunan dosya yollarini panoya alabilirsin.
- `Yurtdisi Etiket Basligi` sekmesi ayri bir YDK Excel dosyasi istemez; etiket bilgilerini ana Excel ayarlarindan okur. Gorsel klasoru ve PDF cikis klasoru ayarlanabilir. PDF kaydi varsayilan olarak `YURTDISI ETIKET_gun.ay.yil` klasorune yapilir.
- Toplu isim degistirme onizlemesinde aile/grup satirlari acilip kapanabilir. `Sec` kolonundaki kutularla tek dosya veya grup bazli secim yapabilirsin.
- `Tum Satirlar`, `Sadece Secili`, `Sadece Hazir`, `Sadece Cakismalar`, `Sadece Manuel`, `Sadece Teknik` filtreleriyle rename listesini daraltabilirsin.
- `Yeni` kolonuna cift tiklayarak gorsel bazli hedef dosya adina manuel mudahale edebilirsin. Uzanti korunur.
- Sagdaki onizleme paneli secili gorseli, mevcut/yeni adini ve kok klasorden sonraki yolunu gosterir.
- `Son Islemi Geri Al` butonu en son basarili toplu rename batch'ini eski dosya adlarina dondurur.
- `Farkli klasore cikart` secenegi acilirsa dosyalar kaynakta yeniden adlandirilmaz; secilen cikis klasorune kaynak klasor agaci korunarak kopyalanir. Ornek: `Cikis/GAMBIA/01_URUN GORSELLERI/02_B2B/StokKlasoru/yeni_ad.ext`.
- Uygulama ikonu `assets/app_icon.ico`, arayuz logosu `assets/logo_white.png` dosyasindan yuklenir. Build script'i bu assets klasorunu exe paketine ekler.
- `Kokten Sonra` kolonu secilen kok klasorden sonraki yolu `.../alt/klasor/dosya.jpg` biciminde gosterir.
- Secimler ve manuel hedef adlar `settings.json` icinde saklanir.
- Ayni hedef ad daha once planlandiysa veya klasorde zaten varsa sistem dosyanin ustune yazmamak icin `_2`, `_3` gibi ekler ve kural kolonunda `isim cakismasi` olarak gosterir.
- `Indeksi Yenile` komutu stok koduna gore fotograf, ana klasor ve kanal klasorlerini `product_index.sqlite` dosyasina kaydeder. Sonraki aramalar once bu indeksi kullanir, indeks yetersizse eski tarama mantigina duser.
- `Kanal Raporu` B2B, WEB, INSTAGRAM ve BEYMEN klasorlerinin secili kapsamda var/yok durumunu listeler.
- Urun sorgulamada gorsel ve Explorer aramalari secili stoklar icin toplu yapilir; ayni klasor her stok icin tekrar tekrar taranmaz.
- Toplu isim degistirme bolumunde dosyalar once gecici isimlerle tasinip sonra hedef ada alinir; bu sayede cakisma riski azaltilir.
