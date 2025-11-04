/**************************************************************
 *  T M B   (Qiagen/CLC CSV) – Tek Dosya, Çakışmasız Kurulum  *
 *  Menü: TMB  →  (1) Lung MSI • 3833   (2) Solid • 3204
 *  Çıktı: "tbm" adlı sayfaya özet rapor
 *
 *  © Aile için hazırlanmıştır. (Kasım 2025)
 **************************************************************/

/** ===========================================================
 *  >>>>>  PANEL ALANLARI (Mb) – GÜNCEL KONSERVATİF DEĞERLER <<<<<
 *  Sadece bu ikisini kendi doğrulanmış değerlerinize göre güncelleyin.
 *  CDHS-53206Z-3833 (Lung MSI) : 0.279–0.280 Mb → konservatif 0.279
 *  CDHS-53205Z-3204 (Solid)    : 0.231 Mb
 *  ===========================================================
 */
const CONST_LUNG_MB  = 0.279;  // ✏️ Gerekirse güncelleyin
const CONST_SOLID_MB = 0.231;  // ✏️ Gerekirse güncelleyin

/** ===========================================================
 *  >>>>>  TEKNİK EŞİKLER (Sıkı ve Gevşek) – İsteğe göre değiştirilebilir
 *  QUAL, DP, ALT (≈ VAF×DP), VAF aralığı ve PASS şartı
 *  ===========================================================
 */
const STRICT_THRESH = { QUAL: 200, DP: 300, ALT: 15, VAF_MIN: 0.10, VAF_MAX: 0.80, REQUIRE_PASS: true,  STR_FILTER: true  };
const LOOSE_THRESH  = { QUAL:  50, DP: 100, ALT:  5, VAF_MIN: 0.05, VAF_MAX: 0.95, REQUIRE_PASS: false, STR_FILTER: false };

/** ===========================================================
 *  MENÜ – Coverage kodunuzdaki menüden ayrı bir menü oluşturur
 *  Not: Aynı projede birden çok onOpen() olabilir; ikisi de çalışır.
 *  ===========================================================
 */
function onOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('TMB')
      .addItem('TMB (Lung MSI • CDHS-53206Z-3833)', 'tmbLung')
      .addItem('TMB (Solid • CDHS-53205Z-3204)',    'tmbSolid')
      .addToUi();
  }catch(e){
    // Sessiz geç
  }
}

/** İki buton **/
function tmbLung(){ runTmb_(CONST_LUNG_MB,  'Akciğer (CDHS-53206Z-3833)'); }
function tmbSolid(){ runTmb_(CONST_SOLID_MB, 'Solid (CDHS-53205Z-3204)'); }

/** ===========================================================
 *  BAŞLIK EŞLEME – Qiagen/CLC CSV için esnek karşılıklar
 *  (sütun adları küçük farklarla değişebilir)
 *  ===========================================================
 */
function headerMap_(headers){
  const norm = headers.map(h => (h||'').toString().trim().toLowerCase());

  function idxOf(aliases){
    for (let a of aliases){
      const j = norm.indexOf(a.toLowerCase());
      if (j >= 0) return j;
    }
    return -1;
  }

  return {
    chrom: idxOf(['chromosome','chrom']),
    pos: idxOf(['position','start position','pos']),
    end: idxOf(['end position','end']),
    ref: idxOf(['reference allele','ref']),
    alt: idxOf(['sample allele','alt','alternate allele']),
    varType: idxOf(['variation type','variant type','type']),
    qual: idxOf(['call quality','qual','sample genotype quality']),
    dp: idxOf(['read depth','depth','dp']),
    vaf: idxOf(['allele fraction','vaf','variant allele frequency']),
    pass: idxOf(['sample upstream filtering','filter','filters'])
  };
}

/** Sayısal parse (virgül/nokta dayanıklı) */
function toNumber_(v){
  if (v === null || v === undefined) return NaN;
  if (typeof v === 'number') return v;
  const s = v.toString().replace('%','').replace(/\s/g,'').replace(',','.');
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}

/** VAF normalizasyonu: yüzde geldiyse 0–1 aralığına çevir */
function normalizeVaf_(raw){
  let v = toNumber_(raw);
  if (isNaN(v)) return NaN;
  if (v > 1.0) v = v/100.0;  // 47.3 → 0.473
  if (v < 0) v = 0;
  if (v > 1) v = 1;
  return v;
}

/** Basit STR/homopolimer sezgisi (sıkı modda kapatır) */
function looksLikeSTR_(ref, alt){
  // Tek bazın 3+ tekrarından oluşan küçük insersyon/delesyonları kaba ele
  const R = (ref||'').toString().toUpperCase();
  const A = (alt||'').toString().toUpperCase();
  const isIndel = (R.length !== A.length);
  if (!isIndel) return false;

  function homopolymer(s){
    if (!s || s.length < 3) return false;
    return /^([ACGT])\1{2,}$/.test(s); // AAA, CCCC, TTTTT
  }
  return homopolymer(R) || homopolymer(A);
}

/** Varyantın eşiğe uyup uymadığını kontrol et */
function passVariant_(row, idx, thresh){
  const qual = (idx.qual>=0) ? toNumber_(row[idx.qual]) : NaN;
  const dp   = (idx.dp>=0)   ? toNumber_(row[idx.dp])   : NaN;
  const vaf  = (idx.vaf>=0)  ? normalizeVaf_(row[idx.vaf]) : NaN;

  // ALT okun sayısı türetme: ALT ≈ VAF × DP
  const altReads = (!isNaN(vaf) && !isNaN(dp)) ? Math.round(vaf * dp) : NaN;

  // PASS/FILTER
  let isPass = true;
  if (idx.pass >= 0){
    const f = (row[idx.pass]||'').toString().toLowerCase();
    // Qiagen export’ta genelde "Pass" veya boş
    isPass = f.indexOf('pass') >= 0 || f === '' || f === 'ok';
  }

  // STR/homopolimer elemesi (yalnızca sıkı mod)
  const ref = (idx.ref>=0) ? row[idx.ref] : '';
  const alt = (idx.alt>=0) ? row[idx.alt] : '';
  if (thresh.STR_FILTER && looksLikeSTR_(ref, alt)) return {ok:false};

  // Eşikler
  if (!isNaN(thresh.QUAL) && !isNaN(qual) && qual < thresh.QUAL) return {ok:false};
  if (!isNaN(thresh.DP)   && !isNaN(dp)   && dp   < thresh.DP)   return {ok:false};
  if (!isNaN(thresh.ALT)  && !isNaN(altReads) && altReads < thresh.ALT) return {ok:false};
  if (!isNaN(thresh.VAF_MIN) && !isNaN(vaf) && vaf < thresh.VAF_MIN) return {ok:false};
  if (!isNaN(thresh.VAF_MAX) && !isNaN(vaf) && vaf > thresh.VAF_MAX) return {ok:false};
  if (thresh.REQUIRE_PASS && !isPass) return {ok:false};

  return {
    ok: true,
    chrom: (idx.chrom>=0)? row[idx.chrom] : '',
    pos:   (idx.pos>=0)?   row[idx.pos]   : '',
    ref:   ref,
    alt:   alt,
    qual:  qual,
    dp:    dp,
    vaf:   vaf,
    altReads: altReads
  };
}

/** TMB ana çalıştırıcı */
function runTmb_(panelMb, panelLabel){
  const sh = SpreadsheetApp.getActiveSheet();
  const name = sh.getName();
  const values = sh.getDataRange().getDisplayValues();

  if (!values || values.length < 2)
    return uiAlert_('Aktif sayfada veri yok.');

  const headers = values[0];
  const idx = headerMap_(headers);

  // Minimum gereken sütunlar: Chromosome, Position, Reference/Sample Allele
  if (idx.chrom < 0 || idx.pos < 0 || (idx.ref < 0 && idx.alt < 0)){
    return uiAlert_('Beklenen VCF/CSV sütunları bulunamadı (Chromosome, Position, Reference/Sample Allele).');
  }

  // Her iki mod için say
  const resStrict = [];
  const resLoose  = [];

  for (let r=1; r<values.length; r++){
    const row = values[r];

    // Sıkı
    const s = passVariant_(row, idx, STRICT_THRESH);
    if (s.ok) resStrict.push(s);

    // Gevşek
    const g = passVariant_(row, idx, LOOSE_THRESH);
    if (g.ok) resLoose.push(g);
  }

  // TMB hesapları
  const tmbStrict = (panelMb > 0) ? (resStrict.length / panelMb) : 0;
  const tmbLoose  = (panelMb > 0) ? (resLoose.length  / panelMb) : 0;

  // Çıktı sayfası
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = ss.getSheetByName('tbm') || ss.insertSheet('tbm');
  out.clear();

  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'Europe/Istanbul', 'yyyy-MM-dd');

  const lines = [];
  lines.push('*** TMB RAPORU ***');
  lines.push('Tarih: ' + today);
  lines.push('Örnek: ' + name);
  lines.push('Kaynak sayfa: ' + name);
  lines.push('Panel: ' + panelLabel + ' (Qiagen/CLC CSV)');
  lines.push('Panel boyutu (Mb): ' + panelMb.toFixed(3));
  lines.push('');

  lines.push('[Sıkı] Filtre Eşikleri: QUAL≥' + STRICT_THRESH.QUAL +
             ', DP≥' + STRICT_THRESH.DP +
             ', ALT≥' + STRICT_THRESH.ALT +
             ', VAF≥' + STRICT_THRESH.VAF_MIN.toFixed(2) + '–' + STRICT_THRESH.VAF_MAX.toFixed(2));
  lines.push('[Sıkı] Filter PASS şartı: ' + (STRICT_THRESH.REQUIRE_PASS ? 'Evet' : 'Hayır'));
  lines.push('[Sıkı] STR artefakt elemesi: ' + (STRICT_THRESH.STR_FILTER ? 'Evet' : 'Hayır'));
  lines.push('[Sıkı] Nitelikli varyant sayısı: ' + resStrict.length);
  lines.push('[Sıkı] TMB (varyant/Mb): ' + tmbStrict.toFixed(2));
  lines.push('');

  lines.push('[Gevşek (teknik)] Filtre Eşikleri: QUAL≥' + LOOSE_THRESH.QUAL +
             ', DP≥' + LOOSE_THRESH.DP +
             ', ALT≥' + LOOSE_THRESH.ALT +
             ', VAF≥' + LOOSE_THRESH.VAF_MIN.toFixed(2) + '–' + LOOSE_THRESH.VAF_MAX.toFixed(2));
  lines.push('[Gevşek (teknik)] Filter PASS şartı: ' + (LOOSE_THRESH.REQUIRE_PASS ? 'Evet' : 'Hayır'));
  lines.push('[Gevşek (teknik)] STR artefakt elemesi: ' + (LOOSE_THRESH.STR_FILTER ? 'Evet' : 'Hayır'));
  lines.push('[Gevşek (teknik)] Nitelikli varyant sayısı: ' + resLoose.length);
  lines.push('[Gevşek (teknik)] TMB (varyant/Mb): ' + tmbLoose.toFixed(2));
  lines.push('');

  lines.push('Notlar:');
  lines.push('- TMB değerleri panel tabanlıdır; klinik raporlama için doğrulanmış panel Mb ve eşikleriniz esas alınmalıdır.');
  lines.push('- CSV’de PASS/AD/VAF olmayabilir. Bu durumda ALT≈VAF×DP gibi türetmeler kullanılır.');
  lines.push('- Sıkı koşul 0 dönerse, keşif amaçlı gevşek koşullu teknik TMB de referans için verilmiştir.');
  lines.push('- Germline/benign/silent dışlama yapılmadığından, klinik karar desteği için tek başına kullanılmamalıdır.');

  // Yaz
  out.getRange(1,1,lines.length,1).setValues(lines.map(s=>[s]));

  // İlk 30 nitelikli varyant (Sıkı) – özet tablo
  const tableHead = ['CHROM','POS','REF','ALT','QUAL','DP','VAF','ALT_EST'];
  const preview = resStrict.slice(0,30).map(v => [
    v.chrom, v.pos, v.ref, v.alt,
    isNaN(v.qual)? '' : v.qual,
    isNaN(v.dp)?   '' : v.dp,
    isNaN(v.vaf)?  '' : v.vaf.toFixed(3),
    isNaN(v.altReads)? '' : v.altReads
  ]);

  if (preview.length){
    const startRow = lines.length + 2;
    out.getRange(startRow,1,1,tableHead.length).setValues([tableHead]);
    out.getRange(startRow+1,1,preview.length,tableHead.length).setValues(preview);
  }

  out.setColumnWidths(1, 1, 520);
  SpreadsheetApp.flush();
  uiToast_('TMB tamam: Sıkı=' + resStrict.length + ' ('
           + tmbStrict.toFixed(2) + '/Mb), Gevşek=' + resLoose.length + ' ('
           + tmbLoose.toFixed(2) + '/Mb).');
}

/** Küçük yardımcılar */
function uiAlert_(msg){ SpreadsheetApp.getUi().alert(msg); }
function uiToast_(msg){ SpreadsheetApp.getActive().toast(msg, 'TMB', 6); }
