/**************************************************************
 * tbm.gs — Panel tabanlı TMB (variants/Mb) hesaplama
 * ÇAKIŞMASIN diye onOpen **YOK**. Menü eklemek isterseniz,
 * coverage.gs içindeki onOpen()'a aşağıdaki tek satırı ekleyin:
 *   if (typeof addTmbMenu_ === 'function') addTmbMenu_();
 **************************************************************/

/** ====================== KULLANICI AYARLARI ====================== **/

// TODO: Panel boyutları (Mb) — ihtiyacınıza göre güncelleyin
const PANEL_MB_LUNG  = 1.200;  // Akciğer (ör. 3833 panel)
const PANEL_MB_SOLID = 1.200;  // Solid (ör. 3204 panel)

// TODO: Filtre eşikleri (istenirse değiştirin)
const QUAL_MIN = 50.0;
const DP_MIN   = 100;
const ALT_MIN  = 5;       // ALT sayısı yoksa ~ VAF*DP ile tahmin edilir
const VAF_MIN  = 0.05;    // 5%

// TODO: PASS şartı — CSV’de Filter kolonu varsa PASS arar; yoksa atlar
const REQUIRE_PASS = true;

// TODO: Basit STR/homopolimer elemesi — default: KAPALI
const STR_FILTER = false;

/** ====================== MENÜ (opsiyonel) ====================== **/
// BUNU coverage.gs -> onOpen() İÇİNDEN ÇAĞIRIN:
//   if (typeof addTmbMenu_ === 'function') addTmbMenu_();
function addTmbMenu_(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TMB')
    .addItem('Akciğer panel TMB yükü', 'runTMB_Lung')
    .addItem('Solid panel TMB yükü',    'runTMB_Solid')
    .addToUi();
}

/** ====================== BAŞLIK ALIASES ====================== **/

// Qiagen/CLC CSV ihracatlarına göre esnek başlık eşleme
const HEADER_ALIASES = {
  chrom: ["Chromosome","#CHROM","CHROM","Chrom","Chr","Contig","Chromosome No"],
  pos:   ["Start Position","Start","POS","Position","Variant Position","Genomic Position","Start position (bp)"],
  ref:   ["Reference Allele","REF","Ref Allele","Ref"],
  alt:   ["Sample Allele","ALT","Alt Allele","Alt","Alt allele"],
  qual:  ["Sample Call Quality","QUAL","Call Quality","Quality","Qual"],
  dp:    ["Sample Read Depth","DP","Depth","Read Depth","Coverage","Total Depth","Total depth"],
  vaf:   ["Sample Allele Fraction","VAF","AF","Allele Frequency","Allele fraction","Fraction"],
  gt:    ["Sample Genotype","GT","Genotype"],
  ad:    ["AD","Allelic Depth","Sample Allele Depth","Alt allele depth"],
  filter:["Filter","FILTER","Variant Filter","Call Filter"]
};

function headerIndexMap_(headers){
  const norm = headers.map(h => (h||"").toString().trim().toLowerCase());
  function idxOfAny(names){
    for (const n of names){
      const i = norm.indexOf((n||"").toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  }
  const m = {};
  Object.keys(HEADER_ALIASES).forEach(k => m[k] = idxOfAny(HEADER_ALIASES[k]));
  return m;
}

/** ====================== YARDIMCILAR ====================== **/

function asNumber_(v){
  if (v == null) return NaN;
  const s = v.toString().replace(',', '.').trim();
  if (s.endsWith('%')) {
    const x = parseFloat(s.slice(0, -1));
    return isNaN(x) ? NaN : (x/100.0);
  }
  const x = parseFloat(s);
  return isNaN(x) ? NaN : x;
}

function maxAltFromAD_(adCell){
  // AD biçimleri: "123,45", veya "45" (sadece ALT), veya "123|45", veya çoklu ALT "123,20,25"
  if (adCell == null) return NaN;
  const s = adCell.toString().trim();
  const parts = s.split(/[,\|;]/).map(t => t.trim()).filter(Boolean);
  if (!parts.length) return NaN;
  // Eğer ilk değer REF, diğerleri ALT ise: max(ALT’lar)
  if (parts.length >= 2) {
    const nums = parts.map(p => asNumber_(p)).filter(x => !isNaN(x));
    if (!nums.length) return NaN;
    // Geleneksel VCF AD: [REF, ALT1, ALT2...]
    const alts = nums.slice(1);
    return alts.length ? Math.max.apply(null, alts) : (nums.length>1 ? nums[1] : nums[0]);
  }
  // Tek değer ise onu ALT sayalım
  const n = asNumber_(parts[0]);
  return isNaN(n) ? NaN : n;
}

function isLikelySTR_(ref, alt){
  if (!STR_FILTER) return false;
  const R = (ref||"").toString().toUpperCase();
  const A = (alt||"").toString().toUpperCase();
  // Homopolimer uzun ekleme/silme: ≥4 aynı baz ardışık
  const rep = /^(A{4,}|C{4,}|G{4,}|T{4,})$/;
  return rep.test(R) || rep.test(A);
}

function passFilter_(filterCell){
  if (!REQUIRE_PASS) return true;
  if (filterCell == null) return true; // kolon yoksa veya boşsa: geç
  const s = filterCell.toString().toUpperCase().trim();
  return (s === "PASS" || s === "." || s === "");
}

/** ====================== ANA AKIŞ ====================== **/

function runTMB_Lung(){    runTMBFromActiveSheet_(PANEL_MB_LUNG,  "Akciğer"); }
function runTMB_Solid(){   runTMBFromActiveSheet_(PANEL_MB_SOLID, "Solid");   }

function runTMBFromActiveSheet_(panelMb, panelLabel){
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) {
    SpreadsheetApp.getUi().alert("Aktif sayfada veri bulunamadı.");
    return;
  }

  const headers = values[0];
  const idx = headerIndexMap_(headers);

  if (idx.chrom < 0 || idx.pos < 0){
    SpreadsheetApp.getUi().alert(
      "CSV/VCF sayfasında konum sütunları bulunamadı (Chromosome, Start Position vb.).\n" +
      "Qiagen export’ta bu iki kolonu mutlaka ekleyin."
    );
    return;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  let qualified = []; // {chrom,pos,ref,alt,qual,dp,altMax,vaf,filter}

  for (let r=1; r<values.length; r++){
    const row = values[r];

    const chrom  = (idx.chrom  >=0) ? row[idx.chrom]  : "";
    const pos    = (idx.pos    >=0) ? row[idx.pos]    : "";
    const ref    = (idx.ref    >=0) ? row[idx.ref]    : "";
    const alt    = (idx.alt    >=0) ? row[idx.alt]    : "";
    const qual   = (idx.qual   >=0) ? asNumber_(row[idx.qual])  : NaN;
    const dp     = (idx.dp     >=0) ? asNumber_(row[idx.dp])    : NaN;
    let vaf      = (idx.vaf    >=0) ? asNumber_(row[idx.vaf])   : NaN;
    const adCell = (idx.ad     >=0) ? row[idx.ad]               : null;
    const filter = (idx.filter >=0) ? row[idx.filter]           : "";

    if (!passFilter_(filter)) continue;
    if (!isNaN(qual) && qual < QUAL_MIN) continue;
    if (!isNaN(dp)   && dp   < DP_MIN)   continue;

    // AD’den ALT yakalamayı dene
    let altMax = maxAltFromAD_(adCell);

    // VAF yoksa AD/DP’den tahmin, AD yoksa VAF*DP ile ALT tahmin
    if (isNaN(vaf) || vaf <= 0){
      if (!isNaN(altMax) && !isNaN(dp) && dp > 0){
        vaf = altMax / dp;
      }
    }
    if ((isNaN(altMax) || altMax <= 0) && !isNaN(vaf) && !isNaN(dp)){
      altMax = Math.round(vaf * dp);
    }

    if (!isNaN(vaf) && vaf < VAF_MIN) continue;
    if (!isNaN(altMax) && altMax < ALT_MIN) continue;
    if (isLikelySTR_(ref, alt)) continue;

    qualified.push({
      chrom: chrom,
      pos: pos,
      ref: ref,
      alt: alt,
      qual: isNaN(qual) ? "" : qual,
      dp:   isNaN(dp)   ? "" : dp,
      altMax: isNaN(altMax) ? "" : altMax,
      vaf:  isNaN(vaf)  ? "" : vaf,
      filter: (filter||"")
    });
  }

  const tmb = (panelMb > 0) ? (qualified.length / panelMb) : NaN;

  // Çıktı sayfası
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = ss.getSheetByName("tbm") || ss.insertSheet("tbm");
  out.clear();

  // Başlık blok
  const sampleName = sheet.getName();
  const lines = [
    "*** TMB RAPORU ***",
    "Tarih: " + today,
    "Örnek: " + sampleName,
    "Kaynak sayfa: " + sampleName,
    "Panel: " + panelLabel + " (Qiagen/CLC CSV)",
    "Panel boyutu (Mb): " + panelMb.toFixed(3),
    "Filtre Eşikleri: QUAL≥" + QUAL_MIN + ", DP≥" + DP_MIN + ", ALT≥" + ALT_MIN + ", VAF≥" + VAF_MIN,
    "Filter PASS şartı: " + (REQUIRE_PASS ? "Evet" : "Hayır"),
    "STR artefakt elemesi: " + (STR_FILTER ? "Evet" : "Hayır"),
    "",
    "Nitelikli varyant sayısı: " + qualified.length,
    "TMB (varyant/Mb): " + (isNaN(tmb) ? "-" : tmb.toFixed(2)),
    "",
    "Notlar:",
    "- Bu değer panel tabanlı, filtrelenmiş ham TMB'dir.",
    "- Klinik raporlama için doğrulanmış panel Mb ve eşiklerinizi esas alın.",
    "- CSV’de PASS/AD/VAF yoksa, uygun yaklaşımlar (PASS serbest, ALT≈VAF×DP) uygulanır.",
    ""
  ];
  out.getRange(1,1,lines.length,1).setValues(lines.map(s => [s]));

  // Detay tablo başlık
  const headerRow = ["CHROM","POS","REF","ALT","QUAL","DP","ALT_AD_MAX","VAF","FILTER"];
  out.getRange(lines.length+1,1,1,headerRow.length).setValues([headerRow]);

  // Detay tablo satırlar
  if (qualified.length){
    const rows = qualified.slice(0, Math.max(qualified.length,1)).map(v => ([
      v.chrom, v.pos, v.ref, v.alt,
      v.qual === "" ? "" : Number(v.qual),
      v.dp   === "" ? "" : Number(v.dp),
      v.altMax === "" ? "" : Number(v.altMax),
      v.vaf  === "" ? "" : Number(v.vaf),
      v.filter
    ]));
    out.getRange(lines.length+2,1,rows.length, headerRow.length).setValues(rows);
  }

  out.autoResizeColumns(1, headerRow.length);
  out.activate();
}
