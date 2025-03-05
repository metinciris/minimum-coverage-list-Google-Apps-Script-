# Kapsama Dışı Gen Listesi Oluşturucu (Google Apps Script)

Bu araç, hedefli dizileme verilerindeki düşük kapsama sahip gen bölgelerinin listesini Google Sheets'ten otomatik olarak oluşturmak için tasarlanmıştır. Belirli bir eşik değerinin altında kapsama sahip genleri ve bölgeleri belirlemek için Google Sheets'teki verileri analiz eder.

Şu bilgileri daha rahat elde etmenizi sağlar:
* okuma derinliğinin altındaki okumalar varyant olarak yansımamaktadır.
* okuma derinliğinin altındaki bölgeler kapsamamaktadır.
* okuma derinliğinin altındaki varyantlar saptanmamaktadır.

## Özellikler

*   Google Sheets'teki verileri analiz edebilir.
*   Kullanıcı tarafından belirlenebilir kapsama eşik değerleri (50 ve 100 için ayrı tuşlar).
*   Sonuçları sıralı bir şekilde listeler.
*   Sonuçları doğrudan Google Sheets'e yazdırır.
*   Kullanımı kolay menü arayüzü.

## Gereksinimler

*   Google Hesabı
*   Google Sheets

## Kurulum

1. (paired)_Per-region_QC_for_Targeted_Sequencing.xls tablonuzu Google tablolar içinde içeri aktarın. Dosya--> içe aktar --> yükle --> yeni sayfa oluştur.
2.   Tablonun aşağıdaki sütunları içermesi gerekmektedir:
    *   Chromosome
    *   Region
    *   Name
    *   Annotation type
    *   Target region length
    *   Target region length with coverage above 50
    *   Percentage with coverage above 50
    *   Read count
    *   Broken read count
    *   Non-specific read count
    *   Base count
    *   GC %
    *   Min coverage
    *   Max coverage
    *   Mean coverage
    *   Median coverage
    *   Zero coverage bases
    *   Mean coverage (excluding zero coverage)
    *   Median coverage (excluding zero coverage)
3.  Google Sheets'te, "Araçlar" (Tools) menüsüne gidin ve "Komut Dosyası Düzenleyicisi" (Script editor) seçeneğini tıklayın.
4.  Açılan Google Apps Script düzenleyicisine aşağıdaki kodu yapıştırın:

    ```
    function onOpen() {
      // Menü öğelerini oluşturun
      var ui = SpreadsheetApp.getUi();
      ui.createMenu('Kapsama Analizi')
          .addItem('Kapsama-50 Analizi', 'analyzeCoverage50')
          .addItem('Kapsama-100 Analizi', 'analyzeCoverage100')
          .addToUi();
    }

    function analyzeCoverage(coverageThreshold) {
      // Aktif sayfayı alın
      var sheet = SpreadsheetApp.getActiveSheet();

      // Veri aralığını tanımlayın (A1'den başlayarak)
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();

      // Sonuçları saklamak için bir nesne oluşturun
      var results = {};

      // Başlık satırını atlayın
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var geneName = row[2]; // "Name" sütunu (C sütunu, 3. sütun)
        var region = row[1];   // "Region" sütunu (B sütunu, 2. sütun)
        var minCoverage = row; // "Min coverage" sütunu (M sütunu, 13. sütun)

        // Kapsama eşiğin altında ise ve minCoverage geçerli bir sayı ise
        var minCoverageNumber = Number(minCoverage); // Sayıya dönüştürmeyi dene

        if (!isNaN(minCoverageNumber) && minCoverageNumber < coverageThreshold) {
          if (!results[geneName]) {
            results[geneName] = [];
          }
          results[geneName].push(region);
        }
      }

      // Sonuçları biçimlendirin
      var formattedResults = "- Aşağıdaki genlerin ilgili bölgeleri bu dizileme çalışmasında kapsanmamıştır (okuma derinliği <" + coverageThreshold + "):\n";
      for (var gene in results) {
        formattedResults += "  - " + gene + ": " + results[gene].join(", ") + "\n";
      }

      // "kapsam" sayfasını alın veya oluşturun
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var coverageSheet = ss.getSheetByName("kapsam");
      if (!coverageSheet) {
        coverageSheet = ss.insertSheet("kapsam");
      }

      // Başlık satırını A1'e yazdırın
      coverageSheet.getRange(1, 1).setValue("- Aşağıdaki genlerin ilgili bölgeleri bu dizileme çalışmasında kapsanmamıştır (okuma derinliği <" + coverageThreshold + "):");

      // "kapsam" sayfasını temizleyin (başlık hariç)
      coverageSheet.getRange("A2:Z").clearContent();

      // Sonuçları "kapsam" sayfasına A2'den itibaren yazdırın
      coverageSheet.getRange(2, 1).setValue(formattedResults);
      SpreadsheetApp.flush(); // Verilerin hemen görüntülenmesini sağlayın
    }

    function analyzeCoverage50() {
      analyzeCoverage(50);
    }

    function analyzeCoverage100() {
      analyzeCoverage(100);
    }
    ```

5.  Kodu kaydedin ve düzenleyiciyi kapatın.

## Kullanım

1.  Google Sheets'te, verilerinizi içeren tabloyu açın. İçeri aktarılan tablo aktif olmalıdır.
2.  Menü çubuğunda, "Kapsama Analizi" adlı yeni bir menü göreceksiniz.
3.  "Kapsama Analizi" menüsünden, istediğiniz eşik değerine göre "Kapsama-50 Analizi" veya "Kapsama-100 Analizi" seçeneğini tıklayın.
4.  Kodun Google E-Tablolarınıza erişmesine izin vermeniz istenebilir. İzinleri verin.
5.  Sonuçlar, "kapsam" adlı yeni bir sayfada görüntülenecektir.

## Örnek Sonuçlar

Aşağıdaki örnek, eşik değeri 100 olarak ayarlandığında elde edilen tipik bir sonucu göstermektedir.

```
- Aşağıdaki genlerin ilgili bölgeleri bu dizileme çalışmasında kapsanmamıştır (okuma derinliği <100):
  - BRCA1: 17:43044295..43044450, 17:43047612..43047890
  - EGFR: 7:55019017..55019230
  - TP53: 17:7571720..7571900, 17:7573282..7573455, 17:7574003..7574188, 17:7577122..7577300
```

Bu sonuç, BRCA1 geninin exon2 ve exon5 bölgelerinin, EGFR geninin exon7 bölgesinin ve TP53 geninin intron4 bölgesinin, belirtilen dizileme çalışmasında 100'ün altında okuma derinliğine sahip olduğunu gösterir.

## Katkıda Bulunma

Katkılarınızı bekliyoruz! Lütfen bir "pull request" oluşturarak veya sorunları bildirerek katkıda bulunun.

## Lisans

Bu proje MIT lisansı altında lisanslanmıştır.
