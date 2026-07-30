İdari Hakediş Şablonu — Kullanım

Bu klasördeki CSV dosyaları uygulamanın beklediği sayfalar için örnek şablondur.
Dosyalar:
- idari_hakedis_IsProgrami.csv -> Sayfa adı: IsProgrami (İş Programı ve İmalatlar)
- idari_hakedis_Endeks.csv -> Sayfa adı: Endeks (Endeks değerleri)
- idari_hakedis_AltEndeks.csv -> Sayfa adı: AltEndeks (Ağırlık, Katsayı, Temel Endeks)
- idari_hakedis_B.csv -> Sayfa adı: B (B katsayıları)

Excel (.xlsx) oluşturma:
1) CSV dosyalarını Excel ile açın.
2) Her CSV için yeni bir çalışma sayfası oluşturun ve sayfa adını yukarıdaki gibi ayarlayın (IsProgrami, Endeks, AltEndeks, B).
3) Tek bir .xlsx dosyası olarak kaydedin (Farklı kaydet -> Excel Çalışma Kitabı (*.xlsx)).
4) Uygulamaya yüklemek için oluşturduğunuz .xlsx dosyasını kullanın.

Notlar:
- Tarih formatı örneği: Oca 22, Mayıs 26 vb. parse_turkish_date bu tür girişleri işler.
- Ondalık ayırıcı olarak virgül ya da nokta olabilir; uygulama temizleme yapar.
- Eğer isterseniz, doğrudan .xlsx olarak repo'ya eklememi talep edebilirsiniz.
