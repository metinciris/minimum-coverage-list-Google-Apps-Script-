/**** =======================  TMB RAPORLAMA (Qiagen/CLC CSV)  ======================= ****
 * Bu dosya coverage.gs ile ÇAKIŞMAZ: onOpen() TANIMLANMAZ.
 * Menü eklemek için ya:
 *   1) coverage.gs içindeki mevcut onOpen() sonuna addTMBMenu_(); satırını ekle, veya
 *   2) Apps Script’te Tetikleyiciler → addTMBMenu_ için “Açıldığında” kurulabilir tetikleyici oluştur.
 *
 * Menü: TMB ▸ Akciğer (teknik) / Akciğer (klinik) / Solid (teknik) / Solid (klinik) / Temizle
 *
 * Notlar:
 * - “Klinik” mod: kodlayan non-synonymous SNV + kodlayan indeller + splice; synonymous/intronic/UTR/promoter hariç;
 *   benign/likely benign ve gnomAD≥0.001 dışlanır.
 * - “Teknik” mod: panel tabanlı ham sayım; PASS/AD/VAF yoksa ALT≈VAF×DP türetilir.
 ****/

/////////////////////////// KULLANICI DÜZENLEYEBİLİR SABİTLER ///////////////////////////

// Panel boyutları (Mb) – konservatif
const TMB_PANEL_MB_LUNG  = 0.279;   // CDHS-53206Z-3833
const TMB_PANEL_MB_SOLID = 0.231;   // CDHS-53205Z-3204

// Rapor formatı
const TMB_MAX_LIST_VARIANTS = 120;  // rapora düşülecek maksimum varyant sayısı

// Klinik mod kriterleri
const TMB_KEEP_IMPACTS = new Set(['missense','nonsense','stop gain','stop_gain','frameshift','splice site','splice_site']);
const TMB_EXCLUDE_CLASSES = new Set(['benign','likely benign','likely_benign']);
const TMB_EXCLUDE_REGIONS = [/intronic/i, /utr/i, /promoter/i, /ncrna/i];  // splice yine dahildir
const TMB_GNOMAD_MAX_AF = 0.001; // >= ise klinikten çıkar

//////////////////////////////// MENÜ EKLEYİCİ //////////////////////////////////////////
/** coverage.gs onOpen()’inden veya kurulabilir tetikleyiciden çağırın */
function addTMBMenu_(){
  SpreadsheetApp.getUi()
    .createMenu('TMB')
    .addItem('Akciğer paneli (teknik TMB)', 'tmb_menu_lung_tech')
    .addItem('Akciğer paneli (klinik TMB)', 'tmb_menu_lung_clin')
    .addSeparator()
    .addItem('Solid panel (teknik TMB)', 'tmb_menu_solid_tech')
    .addItem('Solid panel (klinik TMB)', 'tmb_menu_solid_clin')
    .addSeparator()
    .addItem('Son TBM çıktısını temizle', 'tmb_menu_clear')
    .addToUi();
}

// İstersen bir kere çalıştırıp menünün geldiğini test etmek için:
function tmb_runAddMenuOnce(){ addTMBMenu_(); }

/////////////////////////////// MENÜ İŞLEVCİLERİ ///////////////////////////////////////

function tmb_menu_lung_tech(){ tmb_runForActiveSheet_(TMB_PANEL_MB_LUNG,  'Akciğer (CDHS-53206Z-3833)', /*clinical*/false); }
function tmb_menu_lung_clin(){ tmb_runForActiveSheet_(TMB_PANEL_MB_LUNG,  'Akciğer (CDHS-53206Z-3833)', /*clinical*/true ); }
function tmb_menu_solid_tech(){tmb_runForActiveSheet_(TMB_PANEL_MB_SOLID, 'Solid (CDHS-53205Z-3204)' , /*clinical*/false); }
function tmb_menu_solid_clin(){tmb_runForActiveSheet_(TMB_PANEL_MB_SOLID, 'Solid (CDHS-53205Z-3204)' , /*clinical*/true ); }

function tmb_menu_clear(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('tbm') || ss.insertSheet('tbm');
  sh.clear();
  sh.getRange(1,1).setValue('TBM çıktısı temizlendi.');
}

//////////////////////////////// ANA AKIŞ //////////////////////////////////////////////

function tmb_runForActiveSheet_(panelMb, panelLabel, clinicalMode){
  const src = SpreadsheetApp.getActiveSheet();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const out = ss.getSheetByName('tbm') || ss.insertSheet('tbm');

  // CSV başlık/indeksleri
  const values = src.getDataRange().getValues();
  if (!values || values.length < 2){ return tmb_uiWarn_('Aktif sayfada veri yok.'); }
  const idx = tmb_headerIndexMap_(values[0]);
  const need = ['chrom','pos','ref','alt'];
  const miss = need.filter(k => idx[k] < 0);
  if (miss.length){
    return tmb_uiWarn_(
      'Beklenen VCF/CSV sütunları bulunamadı (Chromosome, Position, Reference Allele, Sample Allele). Eksik: ' +
      miss.join(', ')
    );
  }

  // İki profil: Sıkı & Gevşek
  const strict = { name:'Sıkı', qualMin:200, dpMin:300, altMin:15, vafMin:0.10, vafMax:0.80, requirePass:true,  dropSTR:true  };
  const loose  = { name:'Gevşek (teknik)', qualMin:50,  dpMin:100, altMin:5,  vafMin:0.05, vafMax:0.95, requirePass:false, dropSTR:false };

  const sampleName = src.getName();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const resStrict = tmb_countQualified_(values, idx, strict, clinicalMode);
  const resLoose  = tmb_countQualified_(values, idx, loose,  clinicalMode);

  // Rapor üret
  const report = tmb_buildReport_({
    date: today,
    sample: sampleName,
    sheet: sampleName,
    panelLabel,
    panelMb,
    results: [resStrict, resLoose],
    clinicalMode,
    topList: tmb_pickTopList_(values, idx, resStrict.passedRows, TMB_MAX_LIST_VARIANTS)
  });

  // Yaz
  out.clear();
  tmb_writeMultiline_(out, 1, 1, report);
  out.setColumnWidths(1, 1, 760);
  SpreadsheetApp.flush();
  tmb_uiInfo_('TMB raporu "tbm" sayfasına yazıldı.');
}

/////////////////////////////// SAYIM & FİLTRELER /////////////////////////////////////

function tmb_countQualified_(values, idx, prof, clinicalMode){
  let count = 0;
  const passedRows = [];

  for (let r=1; r<values.length; r++){
    const row = values[r];
    if (!row || row.length === 0) continue;

    // Zorunlu alanlar
    const chrom = tmb_str_(row[idx.chrom]);
    const pos   = tmb_toInt_(row[idx.pos]);
    const ref   = tmb_str_(row[idx.ref]);
    const alt   = tmb_str_(row[idx.alt]); // boş (del) olabilir fakat tutarız

    if (!chrom || isNaN(pos) || !ref) continue;

    // Metri̇kler
    const qual = (idx.qual >= 0) ? tmb_toNumber_(row[idx.qual]) : NaN;
    const dp   = (idx.dp   >= 0) ? tmb_toNumber_(row[idx.dp])   : NaN;
    let   vaf  = (idx.vaf  >= 0) ? tmb_toNumber_(row[idx.vaf])  : NaN;
    if (!isNaN(vaf) && vaf > 1) vaf = vaf/100.0; // yüzde gelmiş olabilir

    // ALT (AD) yoksa türet: ALT ≈ round(VAF * DP)
    let altReads = (idx.altAd >= 0) ? tmb_toInt_(row[idx.altAd]) : NaN;
    if (isNaN(altReads) && !isNaN(vaf) && !isNaN(dp)) altReads = Math.round(vaf * dp);

    // PASS?
    const passTxt = (idx.passFlag >= 0) ? tmb_str_(row[idx.passFlag]) : '';
    const isPass  = /pass/i.test(passTxt);

    // Eşikler
    if (!isNaN(prof.qualMin) && !(qual >= prof.qualMin)) continue;
    if (!isNaN(prof.dpMin)   && !(dp   >= prof.dpMin))   continue;
    if (!isNaN(prof.altMin)  && !(altReads >= prof.altMin)) continue;
    if (!isNaN(prof.vafMin)  && !(vaf  >= prof.vafMin))  continue;
    if (!isNaN(prof.vafMax)  && !(vaf <= prof.vafMax))   continue;
    if (prof.requirePass && !isPass) continue;

    // STR/tekrar sezgisel eleme
    if (prof.dropSTR){
      const isIndel = /del|ins/i.test(tmb_str_(row[idx.varType]));
      if (isIndel && tmb_looksLikeSTR_(ref, alt)) continue;
    }

    // Klinik mod filtreleri
    if (clinicalMode){
      const impact = (idx.transImpact>=0) ? row[idx.transImpact] : '';
      const cls    = (idx.classif>=0)     ? row[idx.classif]     : '';
      const region = (idx.region >=0)     ? row[idx.region]      : '';
      const gmaf   = (idx.gnomad>=0)      ? tmb_toNumber_(row[idx.gnomad]) : NaN;

      if (!tmb_regionIsCoding_(region)) continue;       // splice her zaman dahil
      if (!tmb_isAllowedImpact_(impact)) continue;      // allowed etkiler
      if (tmb_isExcludedClass_(cls)) continue;          // benign/likely benign
      if (!isNaN(gmaf) && gmaf >= TMB_GNOMAD_MAX_AF) continue; // popüler varyant dışla
    }

    count++;
    passedRows.push(r);
  }

  return { profile: prof, count, passedRows };
}

function tmb_looksLikeSTR_(ref, alt){
  const s1 = (ref||'').toString();
  const s2 = (alt||'').toString();
  const seq = (s1.length >= s2.length ? s1 : s2);
  if (!seq) return false;
  const m = seq.match(/^([ACGT])\1{3,}$/i); // ≥4 aynı nükleotid
  return !!m;
}

function tmb_regionIsCoding_(region){
  const txt = tmb_str_(region);
  if (/splice/i.test(txt)) return true;
  for (let i=0;i<TMB_EXCLUDE_REGIONS.length;i++){
    if (TMB_EXCLUDE_REGIONS[i].test(txt)) return false;
  }
  return true;
}

function tmb_isAllowedImpact_(impact){
  const s = tmb_str_(impact).toLowerCase();
  for (const k of TMB_KEEP_IMPACTS){
    if (s.indexOf(k) >= 0) return true;
  }
  return false;
}

function tmb_isExcludedClass_(cls){
  const s = tmb_str_(cls).toLowerCase();
  for (const k of TMB_EXCLUDE_CLASSES){
    if (s.indexOf(k) >= 0) return true;
  }
  return false;
}

//////////////////////////////// RAPORLA ///////////////////////////////////////////////

function tmb_buildReport_({date, sample, sheet, panelLabel, panelMb, results, clinicalMode, topList}){
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
    lines.push('- Klinik TMB: kodlayan nonsynonymous SNV/indel + splice; synonymous/intronic/UTR/promoter hariç; benign/likely benign ve gnomAD≥'+TMB_GNOMAD_MAX_AF+' dışlanır.');
  } else {
    lines.push('- Teknik/ham TMB: panel tabanlı sayımdır; raporlama için doğrulanmış panel Mb ve eşikler esas alınmalıdır.');
  }
  lines.push('- CSV’de PASS/AD/VAF sütunları yoksa, VAF ve DP’den ALT≈VAF×DP türetimi yapılır.');
  lines.push('- Eşleşik normal yoksa yüksek VAF (~%50/~%100) varyantlar germline olabilir; klinik yorumda dikkate alınmalıdır.');
  lines.push('');

  if (topList && topList.length){
    lines.push('CHROM\tPOS\tREF\tALT');
    topList.forEach(v => lines.push([v.chrom, v.pos, v.ref, v.alt].join('\t')));
  }

  return lines.join('\n');
}

function tmb_pickTopList_(values, idx, passedRows, N){
  const out = [];
  const take = Math.min(N, passedRows.length);
  for (let i=0;i<take;i++){
    const r = passedRows[i];
    const row = values[r];
    out.push({
      chrom: tmb_str_(row[idx.chrom]),
      pos: tmb_toInt_(row[idx.pos]),
      ref: tmb_str_(row[idx.ref]),
      alt: tmb_str_(row[idx.alt])
    });
  }
  return out;
}

///////////////////////////////// YARDIMCILAR //////////////////////////////////////////

function tmb_uiWarn_(msg){ SpreadsheetApp.getUi().alert('TMB: ' + msg); }
function tmb_uiInfo_(msg){ SpreadsheetApp.getUi().alert('TMB: ' + msg); }

function tmb_writeMultiline_(sheet, r, c, text){
  const lines = text.split('\n');
  const rng = sheet.getRange(r, c, lines.length, 1);
  const vals = lines.map(s => [s]);
  rng.setValues(vals);
  rng.setWrap(true);
}

function tmb_str_(v){ return (v==null) ? '' : String(v); }
function tmb_toInt_(v){
  if (v==null || v==='') return NaN;
  const n = Number(String(v).replace(',', '.'));
  return Math.round(n);
}
function tmb_toNumber_(v){
  if (v==null || v==='') return NaN;
  const s = String(v).replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}

function tmb_headerIndexMap_(headerRow){
  const H = headerRow.map(x => (x||'').toString());
  const L = H.map(x => x.toLowerCase().trim());

  function idxOf(keys){
    for (let k of keys){
      const kk = k.toLowerCase();
      let i = L.indexOf(kk);
      if (i>=0) return i;
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
    altAd:     idxOf(['alt ad','alt count','alt reads']),
    passFlag:  idxOf(['sample upstream filtering','filter']),
    varType:   idxOf(['variation type','var type']),
    transImpact: idxOf(['translation impact','consequence','effect']),
    classif:     idxOf(['classification','clinvar','interpretation']),
    gnomad:      idxOf(['gnomad frequency','gnomad','af']),
    region:      idxOf(['gene region','region','location'])
  };
}
/** ================== KALICI MENÜ: KURULABİLİR onOpen TETİKLEYİCİSİ ================== **/

// Bir kez çalıştır: tetikleyiciyi kurar ve menüyü hemen ekler
function tmb_bootstrap(){
  tmb_installOpenTrigger_();
  addTMBMenu_();
  SpreadsheetApp.getUi().alert('TMB: Kurulum tamam. Sayfayı her açtığınızda TMB menüsü otomatik eklenecek.');
}

// Manuel test için (artık tetikleyiciyi de garanti altına alır)
function tmb_runAddMenuOnce(){
  tmb_installOpenTrigger_();
  addTMBMenu_();
}

// Zaten “çalışma sayfası açıldığında” addTMBMenu_ tetikleyicisi var mı?
function tmb_hasOpenTrigger_(){
  const me = ScriptApp.getProjectTriggers();
  return me.some(t =>
    t.getHandlerFunction && t.getHandlerFunction() === 'addTMBMenu_' &&
    t.getEventType && t.getEventType() === ScriptApp.EventType.ON_OPEN
  );
}

// Yoksa oluştur
function tmb_installOpenTrigger_(){
  if (tmb_hasOpenTrigger_()) return;
  ScriptApp.newTrigger('addTMBMenu_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}
