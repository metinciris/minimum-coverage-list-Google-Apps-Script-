/**************************************************************
 *  TMB Hesabı (Qiagen/CLC CSV odaklı) – Menü: "TMB Hesabı"
 *  Çakışmayı önlemek için onOpen yazmıyoruz.
 *  İlk kullanımda Script Editor'dan addTMBMenu() -> Run edin.
 *  qiAgen panel çıktısında tüm alanlar seçilmiş csv eksport edin
 *  --- ELLE DEĞİŞTİRİLEBİLİR ALANLAR (BAŞLANGIÇ) ---
 **************************************************************/

// PANEL boyutları (Mb) – KENDİ DOĞRULANMIŞ DEĞERLERİNİZİ YAZIN
const PANEL_MB_LUNG  = 1.200; // TODO: kesin değeri yazın
const PANEL_MB_SOLID = 1.200; // TODO: kesin değeri yazın

// Varsayılan (korumacı) filtre eşikleri
const QUAL_MIN   = 200.0; // kalite
const DP_MIN     = 300;   // toplam derinlik
const ALT_MIN    = 15;    // alternatif okuma sayısı
const VAF_MIN    = 0.10;
const VAF_MAX    = 0.80;
const REQUIRE_PASS = true;
const STR_FILTER   = true; // homopolimer/STR tipi adayları ele

// Rapora yazılacak ilk 20 varyant liste sınırı
const TOP_N_TO_LIST = 20;

// CSV başlık aliasları (Qiagen/CLC değişimlerine dayanıklı)
const HEADER_ALIASES = {
  chrom:        ["Chromosome","chr","#CHROM","CHROM"],
  pos:          ["Start Position","POS","Position","Start position","Start"],
  ref:          ["Reference Allele","REF","Reference"],
  alt:          ["Sample Allele","ALT","Alternate","Alt"],
  qual:         ["QUAL","Quality","Score","Variant Score"],
  filter:       ["FILTER","Filter","Filters"],
  dp:           ["Read Depth","DP","Coverage","Total Read Count","Total Read depth"],
  ad:           ["AD","Allele Depth","Sample allele depth","Allelic Depth"],
  vaf:          ["VAF","Variant Allele Frequency","Allele Frequency","AF"],
  // Qiagen CSV’de ins/del ayrı kolonlarda gelebilir:
  ins:          ["Inserted Bases","Inserted base(s)","Insertion"],
  del:          ["Deleted Bases","Deleted base(s)","Deletion"],
  // Bazen tek hücrede "A(123)/G(45)" benzeri sayım da olabilir:
  alleleCounts: ["Allele read counts","Allele Read Counts"]
};

/**************************************************************
 *  --- ELLE DEĞİŞTİRİLEBİLİR ALANLAR (BİTİŞ) ---
 **************************************************************/

/** Menüyü ekler (coverage.gs ile ÇAKIŞMAZ, onOpen yazmıyoruz) */
function addTMBMenu(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TMB Hesabı')
    .addItem('Akciğer panel TMB yükü', 'tmbLung')
    .addItem('Solid panel TMB yükü',   'tmbSolid')
    .addToUi();
}

/** Menü fonksiyonları */
function tmbLung(){  computeTMBForActiveSheet_("Akciğer (Qiagen/CLC CSV)", PANEL_MB_LUNG); }
function tmbSolid(){ computeTMBForActiveSheet_("Solid (Qiagen/CLC CSV)",   PANEL_MB_SOLID); }

/** Aktif sayfadaki CSV’den TMB hesapla ve "tbm" sayfasına rapor yaz */
function computeTMBForActiveSheet_(panelName, panelMb){
  try{
    const sh = SpreadsheetApp.getActiveSheet();
    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2)
      throw new Error("Aktif sayfada veri yok.");

    // Başlık eşlemesi
    const header = (values[0] || []).map(x => (x||"").toString().trim());
    const idx = buildHeaderIndexMap_(header, HEADER_ALIASES);

    // CSV doğrulaması (en azından CHROM & POS olmalı)
    if (idx.chrom < 0 || idx.pos < 0)
      throw new Error('Beklenen CSV sütunları eksik: en az "Chromosome" ve "Start Position" gerekir.');

    // Satırları dolaş ve aday varyantları çıkar
    const bestByLocus = new Map(); // "chr:pos" -> kayıt (en iyi DP>QUAL)
    for (let r=1; r<values.length; r++){
      const row = values[r];
      const parsed = parseVariantRow_(row, idx);
      if (!parsed) continue; // satırdan varyant üretilemedi

      // Filtreler
      if (REQUIRE_PASS && parsed.filter && parsed.filter.toString().toUpperCase() !== "PASS") continue;
      if (isNaN(parsed.qual) || parsed.qual < QUAL_MIN) continue;
      if (isNaN(parsed.dp)   || parsed.dp   < DP_MIN)   continue;
      if (isNaN(parsed.alt)  || parsed.alt  < ALT_MIN)  continue;
      if (isNaN(parsed.vaf)  || parsed.vaf  < VAF_MIN || parsed.vaf > VAF_MAX) continue;
      if (STR_FILTER && isLikelySTR_(parsed.ref, parsed.altAllele)) continue;

      // Lokus tekilleştirme (aynı CHROM:POS -> en iyi DP, eşitse QUAL)
      const key = parsed.chrom + ":" + parsed.pos;
      if (!bestByLocus.has(key)){
        bestByLocus.set(key, parsed);
      } else {
        const prev = bestByLocus.get(key);
        const better = (parsed.dp > prev.dp) || (parsed.dp === prev.dp && parsed.qual > prev.qual);
        if (better) bestByLocus.set(key, parsed);
      }
    }

    const qualified = Array.from(bestByLocus.values());
    const tmb = (panelMb && panelMb > 0) ? (qualified.length / panelMb) : null;

    // Raporu yaz
    writeTmbReportSheet_("tbm", {
      today: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Istanbul", "yyyy-MM-dd"),
      sampleName: sh.getName(),
      sourceSheet: sh.getName(),
      panelName: panelName,
      panelMb: panelMb,
      thresholds: `QUAL≥${QUAL_MIN}, DP≥${DP_MIN}, ALT≥${ALT_MIN}, VAF≥${VAF_MIN.toFixed(2)}–${VAF_MAX.toFixed(2)}`,
      requirePass: REQUIRE_PASS,
      strFilter: STR_FILTER,
      qualifiedCount: qualified.length,
      tmb: tmb,
      topList: qualified
        .sort((a,b)=> (b.dp - a.dp) || (b.qual - a.qual))
        .slice(0, TOP_N_TO_LIST)
    });

    SpreadsheetApp.getUi().alert("TMB hesabı tamamlandı: " + (tmb!=null ? tmb.toFixed(2) : "-") + " (variants/Mb). Rapora 'tbm' sayfasından bakabilirsiniz.");

  } catch(err){
    SpreadsheetApp.getUi().alert("Hata: " + err.message);
  }
}

/** Başlık alias eşlemesi -> index haritası */
function buildHeaderIndexMap_(headerRow, aliases){
  const toIndex = (name) => headerRow.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const map = {};
  for (const key in aliases){
    const arr = aliases[key];
    let idx = -1;
    for (let i=0;i<arr.length;i++){
      const j = toIndex(arr[i]);
      if (j >= 0){ idx = j; break; }
    }
    map[key] = idx;
  }
  return map;
}

/** Qiagen CSV satırı -> normalize edilmiş varyant kaydı */
function parseVariantRow_(row, idx){
  // Zorunlu alanlar
  const chrom = safeCell_(row, idx.chrom);
  const pos   = toNumber_(row[idx.pos]);

  if (!chrom || isNaN(pos)) return null;

  // REF / ALT temel değerleri
  let ref = safeCell_(row, idx.ref);
  let altAllele = safeCell_(row, idx.alt);

  // ALT boşsa ins/del kolonlarından üret
  if (!altAllele){
    const insVal = safeCell_(row, idx.ins);
    const delVal = safeCell_(row, idx.del);
    if (delVal){ ref = delVal; altAllele = "-"; }
    else if (insVal){ ref = "-"; altAllele = insVal; }
  }

  // Bazı CSV’lerde allele sayıları tek hücrede olabilir (örn. "A(120);G(30)")
  // AD / ALT hesaplaması
  let dp = toNumber_(row[idx.dp]);
  let altCount = null;

  // 1) AD kolonundan (örn. "120,30" ya da "ref,alt")
  if (idx.ad >= 0){
    const adRaw = (row[idx.ad]||"").toString();
    const nums = adRaw.split(/[;,\s\/\|]+/).map(x => Number(x)).filter(x => !isNaN(x));
    if (nums.length >= 2){
      // varsayım: [ref, alt, ...]
      const refAD = nums[0], altAD = nums[1];
      if (!isNaN(refAD) && !isNaN(altAD)){
        altCount = Number(altAD);
        if (isNaN(dp) || dp <= 0){
          const sum = nums.reduce((a,b)=>a+b,0);
          if (!isNaN(sum) && sum > 0) dp = sum;
        }
      }
    }
  }

  // 2) alleleCounts kolonundan (örn. "A(120) G(30)")
  if (altCount == null && idx.alleleCounts >= 0){
    const countsRaw = (row[idx.alleleCounts]||"").toString();
    const matches = countsRaw.match(/\((\d+)\)/g); // parantez içindeki sayılar
    if (matches && matches.length >= 2){
      const nums = matches.map(s => Number(s.replace(/[()]/g,'') )).filter(x=>!isNaN(x));
      if (nums.length >= 2){
        altCount = nums[1];
        if (isNaN(dp) || dp <= 0){
          const sum = nums.reduce((a,b)=>a+b,0);
          if (!isNaN(sum) && sum > 0) dp = sum;
        }
      }
    }
  }

  // 3) VAF varsa ALT ≈ round(VAF*DP)
  let vaf = (idx.vaf>=0) ? toNumber_(row[idx.vaf]) : NaN;
  if (!isNaN(vaf) && (vaf > 1.0)) vaf = vaf / 100.0; // % ise 0-1’e çevir
  if ((altCount == null || isNaN(altCount)) && !isNaN(vaf) && !isNaN(dp) && dp > 0){
    altCount = Math.round(vaf * dp);
  }

  // 4) Hiçbiri yoksa, ALT_COUNT belirsiz kalabilir
  if (altCount == null || isNaN(altCount)) altCount = NaN;

  // QUAL / FILTER
  const qual = (idx.qual>=0) ? toNumber_(row[idx.qual]) : NaN;
  const filter = (idx.filter>=0) ? (row[idx.filter]||"").toString().trim() : "";

  // VAF’ı yeniden tutarlı hesapla (eğer mümkünse)
  let vafFinal = (!isNaN(vaf)) ? vaf : NaN;
  if (isNaN(vafFinal) && !isNaN(altCount) && !isNaN(dp) && dp>0){
    vafFinal = altCount / dp;
  }

  // Normalize alanlar
  return {
    chrom: chrom.toString().replace(/^chr/i, ''),
    pos: Number(pos),
    ref: (ref||"").toString(),
    altAllele: (altAllele||"").toString(),
    dp: isNaN(dp) ? NaN : Number(dp),
    alt: isNaN(altCount) ? NaN : Number(altCount),
    qual: isNaN(qual) ? NaN : Number(qual),
    vaf: isNaN(vafFinal) ? NaN : Number(vafFinal),
    filter: filter
  };
}

/** Basit STR/homopolimer sezgisi */
function isLikelySTR_(ref, alt){
  const s1 = (ref||"").toString().replace(/-/g,'');
  const s2 = (alt||"").toString().replace(/-/g,'');
  const s = (s1.length >= s2.length) ? s1 : s2;
  if (!s) return false;
  // tek bazın ≥3 tekrarı (AAA, TTTT, vb.)
  if (/^(A+|T+|C+|G+)$/.test(s) && s.length >= 3) return true;
  // ref==alt veya çok kısa/sıfır içerik
  if (s1 === s2) return true;
  return false;
}

/** Hücre okuma yardımcıları */
function safeCell_(row, idx){ return (idx>=0 ? (row[idx]||"").toString().trim() : ""); }
function toNumber_(v){
  if (v === null || v === undefined) return NaN;
  if (typeof v === "number") return v;
  const s = v.toString().replace(',','.');
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}

/** "tbm" sayfasına rapor ve liste yaz */
function writeTmbReportSheet_(sheetName, info){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clear();

  // Başlık blok (A1: tek hücre)
  const lines = [
    "*** TMB RAPORU ***",
    "Tarih: " + info.today,
    "Örnek: " + info.sampleName,
    "Kaynak sayfa: " + info.sourceSheet,
    "Panel: " + info.panelName,
    "Panel boyutu (Mb): " + (info.panelMb != null ? info.panelMb.toFixed(3) : "-"),
    "Filtre Eşikleri: " + info.thresholds,
    "Filter PASS şartı: " + (info.requirePass ? "Evet" : "Hayır"),
    "STR artefakt elemesi: " + (info.strFilter ? "Evet" : "Hayır"),
    "",
    "Nitelikli varyant sayısı: " + info.qualifiedCount,
    "TMB (varyant/Mb): " + (info.tmb != null ? info.tmb.toFixed(2) : "-"),
    "",
    "Notlar:",
    "- Bu değer panel tabanlı, filtrelenmiş teknik TMB’dir (variants/Mb).",
    "- Klinik raporlama için doğrulanmış panel Mb ve eşiklerinizi esas alın.",
    "- CSV’de PASS/AD/VAF bulunmuyorsa, muhafazakâr varsayımlar (ALT≈VAF×DP vb.) kullanılmış olabilir.",
    "- Germline/benign/silent dışlama yapılmadığından, klinik karar desteği için tek başına kullanılmamalıdır."
  ];
  sh.getRange(1,1).setValue(lines.join("\n"));
  sh.getRange(1,1).setWrap(true);

  // Başlık satırı (liste) – B sütunundan itibaren
  const header = ["CHROM","POS","REF","ALT","QUAL","DP","ALT_AD","VAF"];
  sh.getRange(1,3,1,header.length).setValues([header]);

  // İlk N nitelikli varyantı yaz
  const rows = info.topList.map(r => [
    r.chrom, r.pos, r.ref, r.altAllele,
    isNaN(r.qual) ? "" : r.qual,
    isNaN(r.dp)   ? "" : r.dp,
    isNaN(r.alt)  ? "" : r.alt,
    (isNaN(r.vaf) ? "" : r.vaf.toFixed(3))
  ]);
  if (rows.length){
    sh.getRange(2,3,rows.length, header.length).setValues(rows);
  }

  // Otomatik genişlik
  sh.autoResizeColumns(1, 10);
}
