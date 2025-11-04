/**** =======================  TMB RAPORLAMA (Qiagen/CLC CSV)  ======================= ****
 * Menü: TMB ▸ Akciğer (teknik) / Akciğer (klinik) / Solid (teknik) / Solid (klinik)
 * - Kaynak: Aktif sayfa (Qiagen/CLC CSV export – sütun adları Türkçe/İngilizce karışık olabilir)
 * - Çıktı: "tbm" sayfasına rapor metni + ilk N varyant listesi yazılır (varsa üzerine ekler)
 * - Panel Mb: Menüyü seçtiğin panele göre sabit (konservatif)
 *
 * NOT:
 *  - “Klinik” mod: yalnızca raporlanabilir (kodlayan-nonsynonymous SNV + kodlayan indeller + splice) sayılır,
 *    benign/likely benign ve gnomAD≥0.001 olanlar dışlanır; intronik/UTR/promoter hariç tutulur (splice hariç).
 *  - “Teknik” mod: panel tabanlı ham sayım (eşiklerle), PASS/AD/VAF yoksa türev kullanır (ALT≈VAF×DP).
 *
 * Güvenli varsayılanlar:
 *  Strict thresholds: QUAL≥200, DP≥300, ALT≥15, 0.10≤VAF≤0.80, PASS, STR eleme açık
 *  Loose  thresholds: QUAL≥ 50, DP≥100, ALT≥ 5, 0.05≤VAF≤0.95, PASS gerekmez, STR eleme kapalı
 ****/

/////////////////////////// KULLANICI DÜZENLEYEBİLİR SABİTLER ///////////////////////////

// Panel boyutları (Mb) – konservatif
const PANEL_MB_LUNG  = 0.279;   // CDHS-53206Z-3833
const PANEL_MB_SOLID = 0.231;   // CDHS-53205Z-3204

// Rapor formatı
const MAX_LIST_VARIANTS = 120;  // rapora liste düşülecek maksimum varyant sayısı

// Klinik mod kriterleri
const KEEP_IMPACTS = new Set(['missense','nonsense','stop gain','stop_gain','frameshift','splice site','splice_site']);
const EXCLUDE_CLASSES = new Set(['benign','likely benign','likely_benign']);
const EXCLUDE_REGIONS = [/intronic/i, /utr/i, /promoter/i, /ncrna/i];  // splice varsa yine tutulur
const GNOMAD_MAX_AF = 0.001; // >= ise klinikten çıkar

/////////////////////////////// MENÜ /////////////////////////////////////////////////

function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('TMB')
    .addItem('Akciğer paneli (teknik TMB)', 'menu_lung_tech')
    .addItem('Akciğer paneli (klinik TMB)', 'menu_lung_clin')
    .addSeparator()
    .addItem('Solid panel (teknik TMB)', 'menu_solid_tech')
    .addItem('Solid panel (klinik TMB)', 'menu_solid_clin')
    .addSeparator()
    .addItem('Son TBM çıktısını temizle', 'menu_clear_tbm')
    .addToUi();
}

function menu_lung_tech(){ runTMBForActiveSheet_(PANEL_MB_LUNG, 'Akciğer (CDHS-53206Z-3833)', /*clinical*/false); }
function menu_lung_clin(){ runTMBForActiveSheet_(PANEL_MB_LUNG, 'Akciğer (CDHS-53206Z-3833)', /*clinical*/true ); }
function menu_solid_tech(){ runTMBForActiveSheet_(PANEL_MB_SOLID, 'Solid (CDHS-53205Z-3204)', /*clinical*/false); }
function menu_solid_clin(){ runTMBForActiveSheet_(PANEL_MB_SOLID, 'Solid (CDHS-53205Z-3204)', /*clinical*/true ); }

function menu_clear_tbm(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('tbm') || ss.insertSheet('tbm');
  sh.clear();
  sh.getRange(1,1).setValue('TBM çıktısı temizlendi.');
}

//////////////////////////////// KALBİ ////////////////////////////////////////////////

function runTMBForActiveSheet_(panelMb, panelLabel, clinicalMode){
  const src = SpreadsheetApp.getActiveSheet();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const out = ss.getSheetByName('tbm') || ss.insertSheet('tbm');

  // CSV başlık çözümle
  const values = src.getDataRange().getValues();
  if (!values || values.length < 2){
    uiWarn_('Aktif sayfada veri yok.');
    return;
  }
  const idx = headerIndexMap_(values[0]);
  const need = ['chrom','pos','ref','alt'];
  const miss = need.filter(k => idx[k] < 0);
  if (miss.length){
    uiWarn_('Beklenen VCF/CSV sütunları bulunamadı (Chromosome, Position, Reference Allele, Sample Allele). Eksik: ' + miss.join(', '));
    return;
  }

  // İki profil: Strict & Loose (her ikisini hesaplayıp rapora koyuyoruz)
  const strict = {
    name: 'Sıkı',
    qualMin: 200, dpMin: 300, altMin: 15,
    vafMin: 0.10, vafMax: 0.80,
    requirePass: true,
    dropSTR: true
  };
  const loose = {
    name: 'Gevşek (teknik)',
    qualMin: 50, dpMin: 100, altMin: 5,
    vafMin: 0.05, vafMax: 0.95,
    requirePass: false,
    dropSTR: false
  };

  const sampleName = src.getName();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const resultStrict = countQualified_(values, idx, strict, clinicalMode);
  const resultLoose  = countQualified_(values, idx, loose,  clinicalMode);

  // Rapor metni
  const report = buildReport_({
    date: today,
    sample: sampleName,
    sheet: sampleName,
    panelLabel,
    panelMb,
    results: [resultStrict, resultLoose],
    clinicalMode,
    topList: pickTopList_(values, idx, resultStrict.passedRows, MAX_LIST_VARIANTS)
  });

  // Yaz
  out.clear();
  writeMultiline_(out, 1, 1, report);
  out.setColumnWidths(1, 1, 760);
  SpreadsheetApp.flush();
  uiInfo_('TMB raporu "tbm" sayfasına yazıldı.');
}

/////////////////////////////// SAYIM & FİLTRELER /////////////////////////////////////

function countQualified_(values, idx, prof, clinicalMode){
  let count = 0;
  const passedRows = [];

  for (let r=1; r<values.length; r++){
    const row = values[r];
    if (!row || row.length === 0) continue;

    // Zorunlu alanlar
    const chrom = str_(row[idx.chrom]);
    const pos   = toInt_(row[idx.pos]);
    const ref   = str_(row[idx.ref]);
    const alt   = str_(row[idx.alt]);

    if (!chrom || isNaN(pos) || !ref){ continue; } // ALT boş (ör. del/ins gösterimi) olabilir; yine de sayımda yer vereceğiz.

    // Metri̇kler
    const qual = (idx.qual >= 0) ? toNumber_(row[idx.qual]) : NaN;
    const dp   = (idx.dp   >= 0) ? toNumber_(row[idx.dp])   : NaN;
    let   vaf  = (idx.vaf  >= 0) ? toNumber_(row[idx.vaf])  : NaN;

    // VAF yüzde mi geldi? (örn. 99.83) → 0.9983'e çevir
    if (!isNaN(vaf) && vaf > 1) vaf = vaf / 100.0;

    // ALT (AD) yoksa türet: ALT ≈ round(VAF * DP)
    let altReads = (idx.altAd >= 0) ? toInt_(row[idx.altAd]) : NaN;
    if (isNaN(altReads) && !isNaN(vaf) && !isNaN(dp)) altReads = Math.round(vaf * dp);

    // PASS?
    const passTxt = (idx.passFlag >= 0) ? str_(row[idx.passFlag]) : '';
    const isPass  = /pass/i.test(passTxt);

    // Eşikler
    if (!isNaN(prof.qualMin) && !(qual >= prof.qualMin)) continue;
    if (!isNaN(prof.dpMin)   && !(dp   >= prof.dpMin))   continue;
    if (!isNaN(prof.altMin)  && !(altReads >= prof.altMin)) continue;
    if (!isNaN(prof.vafMin)  && !(vaf  >= prof.vafMin))  continue;
    if (!isNaN(prof.vafMax)  && !(vaf <= prof.vafMax))   continue;
    if (prof.requirePass && !isPass) continue;

    // Basit STR/tekrar artefakt elemesi (çok basitleştirilmiş sezgi)
    if (prof.dropSTR){
      const isIndel = /del|ins/i.test(str_(row[idx.varType]));
      if (isIndel && looksLikeSTR_(ref, alt)) continue;
    }

    // Klinik mod filtreleri
    if (clinicalMode){
      const impact = (idx.transImpact>=0) ? row[idx.transImpact] : '';
      const cls    = (idx.classif>=0)     ? row[idx.classif]     : '';
      const region = (idx.region >=0)     ? row[idx.region]      : '';
      const gmaf   = (idx.gnomad>=0)      ? toNumber_(row[idx.gnomad]) : NaN;

      // Bölge kodlayan mı? (splice her zaman dahil)
      if (!regionIsCoding_(region)) continue;

      // Etki tipi uygun mu?
      if (!isAllowedImpact_(impact)) continue;

      // Benign sınıflar hariç
      if (isExcludedClass_(cls)) continue;

      // Popülasyon sık varyant hariç
      if (!isNaN(gmaf) && gmaf >= GNOMAD_MAX_AF) continue;
    }

    count++;
    passedRows.push(r);
  }

  return { profile: prof, count, passedRows };
}

function looksLikeSTR_(ref, alt){
  // 1) Tek bazın 4+ tekrarı (AAAA...), 2) homopolimerik ins/del tahmini
  const s1 = (ref||'').toString();
  const s2 = (alt||'').toString();
  const seq = (s1.length >= s2.length ? s1 : s2);
  if (!seq) return false;
  const m = seq.match(/^([ACGT])\1{3,}$/i); // >=4 aynı nükleotid
  return !!m;
}

function regionIsCoding_(region){
  const txt = str_(region);
  if (/splice/i.test(txt)) return true;
  for (let i=0;i<EXCLUDE_REGIONS.length;i++){
    if (EXCLUDE_REGIONS[i].test(txt)) return false;
  }
  return true;
}

function isAllowedImpact_(impact){
  const s = str_(impact).toLowerCase();
  // "missense; synonymous" gibi birleşik hücreler olabilir
  for (const k of KEEP_IMPACTS){
    if (s.indexOf(k) >= 0) return true;
  }
  return false;
}

function isExcludedClass_(cls){
  const s = str_(cls).toLowerCase();
  for (const k of EXCLUDE_CLASSES){
    if (s.indexOf(k) >= 0) return true;
  }
  return false;
}

//////////////////////////////// RAPORLA ///////////////////////////////////////////////

function buildReport_({date, sample, sheet, panelLabel, panelMb, results, clinicalMode, topList}){
  const lines = [];
  lines.push('*** TMB RAPORU ***');
  lines.push('Tarih: ' + date);
  lines.push('Örnek: ' + sample);
  lines.push('Kaynak sayfa: ' + sheet);
  lines.push('Panel: ' + panelLabel + ' (Qiagen/CLC CSV)');
  lines.push('Panel boyutu (Mb): ' + panelMb.toFixed(3));
  lines.push('');

  results.forEach(r => {
    const tag = '[' + r.profile.name + (clinicalMode ? ' – KLINIK' : '') + ']';
    lines.push(tag + ' Filtre Eşikleri: QUAL≥' + r.profile.qualMin + ', DP≥' + r.profile.dpMin + ', ALT≥' + r.profile.altMin +
               ', VAF≥' + r.profile.vafMin.toFixed(2) + '–' + r.profile.vafMax.toFixed(2));
    lines.push(tag + ' Filter PASS şartı: ' + (r.profile.requirePass ? 'Evet' : 'Hayır'));
    lines.push(tag + ' STR artefakt elemesi: ' + (r.profile.dropSTR ? 'Evet' : 'Hayır'));
    lines.push(tag + ' Nitelikli varyant sayısı: ' + r.count);
    const tmb = (panelMb > 0) ? (r.count / panelMb) : 0;
    lines.push(tag + ' TMB (varyant/Mb): ' + tmb.toFixed(2));
    lines.push('');
  });

  lines.push('Notlar:');
  if (clinicalMode){
    lines.push('- Bu mod, klinik raporlanabilir TMB’yi hedefler: kodlayan nonsynonymous SNV/indel ve splice; synonymous/intronic/UTR/promoter hariç; benign/likely benign ve gnomAD≥'+GNOMAD_MAX_AF+' dışlanır.');
  } else {
    lines.push('- Bu değer panel tabanlı teknik/ham TMB’dir; klinik raporlama için doğrulanmış panel Mb ve eşikler esas alınmalıdır.');
  }
  lines.push('- CSV’de PASS/AD/VAF sütunları yoksa, VAF ve DP’den ALT≈VAF×DP türetimi yapılır.');
  lines.push('- Eşleşik normal yoksa yüksek VAF (~%50/~%100) varyantlar germline olabilir; klinik yorumda dikkate alınmalıdır.');
  lines.push('');

  if (topList && topList.length){
    lines.push('CHROM\tPOS\tREF\tALT');
    topList.forEach(v => {
      lines.push([v.chrom, v.pos, v.ref, v.alt].join('\t'));
    });
  }

  return lines.join('\n');
}

function pickTopList_(values, idx, passedRows, N){
  const out = [];
  const take = Math.min(N, passedRows.length);
  for (let i=0;i<take;i++){
    const r = passedRows[i];
    const row = values[r];
    out.push({
      chrom: str_(row[idx.chrom]),
      pos: toInt_(row[idx.pos]),
      ref: str_(row[idx.ref]),
      alt: str_(row[idx.alt])
    });
  }
  return out;
}

///////////////////////////////// YARDIMCILAR //////////////////////////////////////////

function uiWarn_(msg){ SpreadsheetApp.getUi().alert('TMB: ' + msg); }
function uiInfo_(msg){ SpreadsheetApp.getUi().alert('TMB: ' + msg); }

function writeMultiline_(sheet, r, c, text){
  const lines = text.split('\n');
  const rng = sheet.getRange(r, c, lines.length, 1);
  const vals = lines.map(s => [s]);
  rng.setValues(vals);
  rng.setWrap(true);
}

function str_(v){ return (v==null) ? '' : String(v); }
function toInt_(v){
  if (v==null || v==='') return NaN;
  const n = Number(String(v).replace(',', '.'));
  return Math.round(n);
}
function toNumber_(v){
  if (v==null || v==='') return NaN;
  const s = String(v).replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}

function headerIndexMap_(headerRow){
  const H = headerRow.map(x => (x||'').toString());
  const L = H.map(x => x.toLowerCase().trim());

  function idxOf(keys){
    for (let k of keys){
      const kk = k.toLowerCase();
      // tam eşleşme
      let i = L.indexOf(kk);
      if (i>=0) return i;
      // kısmi içerme (örn. "Sample Genotype Quality" ~ "Genotype Quality")
      i = L.findIndex(t => t.indexOf(kk) >= 0);
      if (i>=0) return i;
    }
    return -1;
  }

  return {
    chrom:     idxOf(['chromosome','chrom','chr']),
    pos:       idxOf(['position','start position','pos']),
    ref:       idxOf(['reference allele','ref']),
    alt:       idxOf(['sample allele','alt','alternate allele']),
    qual:      idxOf(['call quality','qual','sample genotype quality']),
    dp:        idxOf(['read depth','dp','coverage']),
    vaf:       idxOf(['allele fraction','variant allele fraction','vaf']),
    altAd:     idxOf(['alt ad','alt count','alt reads']), // çoğu qiagen csv’de yok
    passFlag:  idxOf(['sample upstream filtering','filter']),
    varType:   idxOf(['variation type','var type']),
    transImpact: idxOf(['translation impact','consequence','effect']),
    classif:     idxOf(['classification','clinvar','interpretation']),
    gnomad:      idxOf(['gnomad frequency','gnomad','af']),
    region:      idxOf(['gene region','region','location'])
  };
}
