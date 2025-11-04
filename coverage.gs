function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Kapsama Analizi')
    .addItem('Kapsama<50 Analizi', 'analyzeCoverage50')
    .addItem('Kapsama<100 Analizi', 'analyzeCoverage100')
    .addItem('İki Çalışma Karşılaştırma (50)', 'compareTwoSheets50')
    .addItem('Sadece A3 Özeti Yaz', 'writeSummaryOnlyA3')
    .addToUi();
    addTMBMenu_(); // TMB menüsünü de ekle


}

/** ===================== YARDIMCILAR ===================== **/

function getHeaderIndexMap(values){
  const h = values[0] || [];
  return {
    chrom: h.indexOf('Chromosome'),
    region: h.indexOf('Region'),
    gene: h.indexOf('Name'),
    minCov: h.indexOf('Min coverage'),
    meanCov: h.indexOf('Mean coverage'),
    medianCov: h.indexOf('Median coverage'),
    readCount: h.indexOf('Read count'),
    baseCount: h.indexOf('Base count'),
    pct50: h.indexOf('Percentage with coverage above 50'),
    lenTarget: h.indexOf('Target region length')
  };
}

// "chr1:120572524..120572615" ya da "120572524..120572615" -> uzunluk (bp)
function parseRegionLength_(regionCell){
  if (typeof regionCell !== 'string') return null;
  const m = regionCell.replace(/,/g,'').match(/(\d+)\.\.(\d+)/);
  if (!m) return null;
  const start = Number(m[1]), end = Number(m[2]);
  if (isNaN(start) || isNaN(end) || end < start) return null;
  return (end - start + 1);
}

// Per-region sayfasından toplamlar (Read count ve Base count sütun toplamı)
function getPerRegionTotals_(sheet){
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return {reads:null, bases:null};
  const idx = getHeaderIndexMap(values);
  let sumReads = 0, sumBases = 0;
  const haveReads = (idx.readCount >= 0), haveBases = (idx.baseCount >= 0);

  for (let i=1;i<values.length;i++){
    if (haveReads){
      const v = values[i][idx.readCount];
      if (typeof v === 'number') sumReads += v;
      else if (v !== '' && !isNaN(Number(v))) sumReads += Number(v);
    }
    if (haveBases){
      const b = values[i][idx.baseCount];
      if (typeof b === 'number') sumBases += b;
      else if (b !== '' && !isNaN(Number(b))) sumBases += Number(b);
    }
  }
  return {
    reads: haveReads ? Math.round(sumReads) : null,
    bases: haveBases ? Math.round(sumBases) : null
  };
}

// Metinden ondalık/tam sayı çek (virgül/nokta ayıklar)
function extractFloatFromText_(txt){
  if (!txt) return null;
  const s = txt.toString();
  const mAll = s.match(/(\d+(?:[.,]\d+)?)/g);
  if (!mAll) return null;
  for (let i=0;i<mAll.length;i++){
    const n = Number(mAll[i].replace(',', '.'));
    if (!isNaN(n)) return n;
  }
  return null;
}

// Q metriklerini ara (bütün çalışma kitabı genelinde, sol-üst 120x12 bloklar)
function findQualityMetricsAllSheets_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let umiQ = null, inputQ = null, q30 = null;

  // Named ranges öncelikli (opsiyonel)
  try {
    const qUmiNR = ss.getRangeByName('Q_UMI');
    if (qUmiNR) umiQ = extractFloatFromText_(qUmiNR.getDisplayValue());
    const qInNR = ss.getRangeByName('Q_INPUT');
    if (qInNR) inputQ = extractFloatFromText_(qInNR.getDisplayValue());
    const q30NR = ss.getRangeByName('Q30_PERCENT');
    if (q30NR) q30 = extractFloatFromText_(q30NR.getDisplayValue());
  } catch(e){}

  // Etiket tarama (TR/EN; UMI/consensus ve Q30 için)
  const RE_ANY_Q = /(umi.*consensus.*q|consensus.*q|umi.*q|average\s*q|mean\s*q|avg\s*q|q-?\s*score|base\s*quality|q30|q\s*30|ortalama\s*q|q\s*skoru|taban\s*kalitesi)/i;
  const RE_UMI = /(umi|consensus)/i;
  const RE_Q30 = /(q30|q\s*30)/i;

  for (let s=0;s<sheets.length;s++){
    const sh = sheets[s];
    const rng = sh.getRange(1,1,Math.min(120, sh.getMaxRows()), Math.min(12, sh.getMaxColumns()));
    const vals = rng.getDisplayValues();

    for (let r=0; r<vals.length; r++){
      for (let c=0; c<vals[r].length; c++){
        const cell = (vals[r][c] || '').toString();
        if (RE_ANY_Q.test(cell)){
          // 1) Aynı hücre
          let val = extractFloatFromText_(cell);
          // 2) Sağ
          if (val == null && c+1 < vals[r].length) val = extractFloatFromText_(vals[r][c+1]);
          // 3) Alt
          if (val == null && r+1 < vals.length) val = extractFloatFromText_(vals[r+1][c]);

          if (val != null){
            if (RE_Q30.test(cell)){
              if (q30 == null) q30 = val;
            } else if (RE_UMI.test(cell)){
              if (umiQ == null) umiQ = val;
            } else {
              if (inputQ == null) inputQ = val;
            }
          }
        }
      }
    }
  }

  return { umiQ, inputQ, q30 };
}

// Tek hücrelik <eşik metni> (A2)
// *** DEĞİŞİKLİK: Gen anahtarları alfabetik sıralanıyor ***
function buildBelowThresholdText_(results, coverageThreshold){
  const genesSorted = Object.keys(results || {})
    .filter(function(g){ return g != null && g !== ''; })
    .sort(function(a,b){
      return a.toString().localeCompare(b.toString(), 'en', {sensitivity:'base'});
    });

  let txt = '- Aşağıdaki genlerin ilgili bölgeleri bu dizileme çalışmasında kapsanmamıştır (okuma derinliği <' + coverageThreshold + '):\n';
  genesSorted.forEach(function(gene){
    // Bölge (region) sırasını aynen bırakıyoruz.
    txt += '  - ' + gene + ': ' + (results[gene] || []).join(', ') + '\n';
  });
  return txt;
}

/** ===================== A3 METNİ ===================== **/

// A3 – 1. satır (Seçenek 2: UMI sonrası Q + Q30 vurgulu)
function buildQcSummaryLine1_(perRegionSheet, threshold){
  const values = perRegionSheet.getDataRange().getValues();
  if (!values || values.length < 2) return "Özet: veri bulunamadı.";
  const idx = getHeaderIndexMap(values);
  if (idx.region < 0 || idx.gene < 0 || idx.minCov < 0) return "Özet: başlıklar eksik (Name/Region/Min coverage).";

  let totalLen = 0, lenBelow = 0, lenAbove50 = 0, lenAbove100 = 0;
  let meanWeightedSum = 0; let medianList = [];
  let countAll = 0, countBelow = 0, countAbove50 = 0, countAbove100 = 0;

  for (let i=1; i<values.length; i++){
    const r = values[i];
    const minCov = Number(r[idx.minCov]);
    const meanCov = (idx.meanCov >= 0) ? Number(r[idx.meanCov]) : NaN;
    const medianCov = (idx.medianCov >= 0) ? Number(r[idx.medianCov]) : NaN;
    const len = parseRegionLength_(r[idx.region]);
    const L = (len && len > 0) ? len : 1;

    totalLen += L; countAll += 1;
    if (!isNaN(meanCov)) meanWeightedSum += meanCov * L;
    if (!isNaN(medianCov)) medianList.push(medianCov);

    if (!isNaN(minCov)){
      if (minCov < 50){ lenBelow += L; countBelow += 1; }
      if (minCov >= 50){ lenAbove50 += L; countAbove50 += 1; }
      if (minCov >= 100){ lenAbove100 += L; countAbove100 += 1; }
    }
  }

  const panelMean = (totalLen > 0) ? (meanWeightedSum / totalLen) : null;
  medianList.sort(function(a,b){return a-b;});
  const medIdx = Math.floor(medianList.length/2);
  const panelMedian = medianList.length ? (medianList.length%2 ? medianList[medIdx] : (medianList[medIdx-1]+medianList[medIdx])/2) : null;

  const pctFull50 = totalLen>0 ? (100 * lenAbove50 / totalLen) : (100 * countAbove50 / Math.max(1,countAll));
  const pctFull100 = totalLen>0 ? (100 * lenAbove100 / totalLen) : (100 * countAbove100 / Math.max(1,countAll));

  // Per-region toplam okuma
  const totalsPR = getPerRegionTotals_(perRegionSheet);

  // Tüm çalışma kitabından Q metrikleri
  const q = findQualityMetricsAllSheets_();

  const parts = [];
  parts.push('Özet (eşik ' + threshold + 'x):');
  if (totalsPR.reads && totalsPR.reads > 0) parts.push('Toplam okuma ~ ' + totalsPR.reads);

  // Seçenek 2: UMI sonrası Q + Q30
  if (q.umiQ != null)  parts.push('UMI sonrası (hata düzeltilmiş) ort. Q ~ ' + q.umiQ.toFixed(2));
  if (q.q30  != null)  parts.push('Q30 ~ ' + q.q30.toFixed(2) + '%');

  parts.push('Panel ortalama kapsama ~ ' + (panelMean ? panelMean.toFixed(1) : '-') + 'x');
  parts.push('panel medyan kapsama ~ ' + (panelMedian ? panelMedian.toFixed(0) : '-') + 'x');
  parts.push('tüm hedeflerin >=50x oranı ~ ' + pctFull50.toFixed(2) + '%');
  parts.push('>=100x oranı ~ ' + pctFull100.toFixed(2) + '%');
  parts.push('<' + threshold + 'x hedef sayısı ' + countBelow + (totalLen>0 ? (', toplam ~ ' + lenBelow + ' bp') : '') + '.');

  return parts.join(' ; ');
}

// A3 – 2. satır (toplamlar)
function buildQcSummaryLine2_(perRegionSheet){
  const totals = getPerRegionTotals_(perRegionSheet);
  const readsStr = (totals.reads != null) ? totals.reads : '-';
  const basesStr = (totals.bases != null) ? totals.bases : '-';
  return 'Toplam okuma ~ ' + readsStr + ' ; Toplam base ~ ' + basesStr;
}

/** ===================== ANA İŞLEVLER ===================== **/

function analyzeCoverage(coverageThreshold) {
  const sheet = SpreadsheetApp.getActiveSheet(); // Per-region sayfası aktif olmalı
  const values = sheet.getDataRange().getValues();
  const idx = getHeaderIndexMap(values);

  const results = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const geneName = row[idx.gene];
    const region = row[idx.region];
    const minCoverageNumber = Number(row[idx.minCov]);
    if (!isNaN(minCoverageNumber) && minCoverageNumber < coverageThreshold) {
      if (!results[geneName]) results[geneName] = [];
      results[geneName].push(region);
    }
  }

  const formattedResults = buildBelowThresholdText_(results, coverageThreshold);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverageSheet = ss.getSheetByName("kapsam") || ss.insertSheet("kapsam");

  coverageSheet.getRange(1,1).setValue('Okuma derinliği <' + coverageThreshold + ' olan hedefler');
  coverageSheet.getRange("A2:Z").clearContent();
  coverageSheet.getRange(2,1).setValue(formattedResults);

  // A3: iki satır tek hücre (Seçenek 2 kalıbı)
  const line1 = buildQcSummaryLine1_(sheet, coverageThreshold);
  const line2 = buildQcSummaryLine2_(sheet);
  coverageSheet.getRange(3,1).setValue(line1 + '\n' + line2);

  SpreadsheetApp.flush();
}

function analyzeCoverage50(){ analyzeCoverage(50); }
function analyzeCoverage100(){ analyzeCoverage(100); }

function compareTwoSheets50() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(function(s){ return s.getName().indexOf('Per-region') === 0; });
  if (sheets.length < 2) {
    SpreadsheetApp.getUi().alert('En az iki "Per-region..." sayfası bulunamadı!');
    return;
  }

  const resultsList = sheets.map(function(s){ return getLowCoverageRegions_(s, 50); });
  const genes = Array.from(resultsList[0].keys()).filter(function(g){ return resultsList.every(function(m){ return m.has(g); }); });

  const commonResults = {};
  genes.forEach(function(g){
    const base = Array.from(resultsList[0].get(g));
    const inter = base.filter(function(r){ return resultsList.every(function(m){ return m.get(g).has(r); }); });
    if (inter.length) commonResults[g] = inter;
  });

  // *** DEĞİŞİKLİK: Gen isimlerini alfabetik sırala, bölgelerin sırası korunur. ***
  let formatted = '- Aşağıdaki genlerin ilgili bölgeleri her iki/çoklu dizileme çalışmasında da kapsanmamıştır (okuma derinliği <50):\n';
  Object.keys(commonResults)
    .sort(function(a,b){ return a.toString().localeCompare(b.toString(), 'en', {sensitivity:'base'}); })
    .forEach(function(g){
      formatted += '  - ' + g + ': ' + commonResults[g].join(', ') + '\n';
    });

  const out = ss.getSheetByName("karsilastirma") || ss.insertSheet("karsilastirma");
  out.clear();
  out.getRange(1,1).setValue(formatted);

  // A3 (kapsam sayfası) güncelle
  const ksheet = ss.getSheetByName("kapsam") || ss.insertSheet("kapsam");
  const line1 = buildQcSummaryLine1_(sheets[0], 50);
  const line2 = buildQcSummaryLine2_(sheets[0]);
  ksheet.getRange(3,1).setValue(line1 + '\n' + line2);

  SpreadsheetApp.flush();
}

// Eşik altındaki bölgeler: Map{gene -> Set(regions)}
function getLowCoverageRegions_(sheet, threshold){
  const v = sheet.getDataRange().getValues();
  const idx = getHeaderIndexMap(v);
  const map = new Map();
  for (let i=1; i<v.length; i++){
    const min = Number(v[i][idx.minCov]);
    if (!isNaN(min) && min < threshold){
      const g = v[i][idx.gene], r = v[i][idx.region];
      if (!map.has(g)) map.set(g, new Set());
      map.get(g).add(r);
    }
  }
  return map;
}

// Sadece A3 özetini yaz (aktif Per-region sayfasından)
function writeSummaryOnlyA3(){
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ksheet = ss.getSheetByName("kapsam") || ss.insertSheet("kapsam");
  const line1 = buildQcSummaryLine1_(sheet, 50);
  const line2 = buildQcSummaryLine2_(sheet);
  ksheet.getRange(3,1).setValue(line1 + '\n' + line2);
}
