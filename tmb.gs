/**************************************************************
 *  TMB Hesabı (Qiagen/CLC CSV) panelden export tümü seçili csv çıktısını al– Menü: "TMB Hesabı"
 *  onOpen ÇAKIŞMASI YOK; menüyü addTMBMenu() ile ekliyoruz.
 **************************************************************/

// ================== ELLE DÜZENLE (Panel Mb) ==================
const PANEL_MB_LUNG  = 1.200; // TODO: kesin değeri yazın
const PANEL_MB_SOLID = 1.200; // TODO: kesin değeri yazın
// =============================================================

// ======= SIKI AYARLAR (öncelikli deneme / daha seçici) =======
const STRICT = {
  QUAL_MIN:   200.0,
  DP_MIN:     300,
  ALT_MIN:    15,
  VAF_MIN:    0.10,
  VAF_MAX:    0.80,
  REQUIRE_PASS: true,
  STR_FILTER:   true,
  LABEL: "Sıkı"
};

// ======= GEVŞEK AYARLAR (fallback / keşif amaçlı) ============
const RELAXED = {
  QUAL_MIN:   50.0,
  DP_MIN:     100,
  ALT_MIN:    5,
  VAF_MIN:    0.05,
  VAF_MAX:    0.95,
  REQUIRE_PASS: false,
  STR_FILTER:   false,
  LABEL: "Gevşek (teknik)"
};

const TOP_N_TO_LIST = 20;

// CSV başlık aliasları
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
  ins:          ["Inserted Bases","Inserted base(s)","Insertion"],
  del:          ["Deleted Bases","Deleted base(s)","Deletion"],
  alleleCounts: ["Allele read counts","Allele Read Counts"]
};

/** Menüyü ekler (coverage.gs ile çakışmaz) */
function addTMBMenu(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TMB Hesabı')
    .addItem('Akciğer panel TMB yükü', 'tmbLung')
    .addItem('Solid panel TMB yükü',   'tmbSolid')
    .addToUi();
}

function tmbLung(){  runTMB_("Akciğer (Qiagen/CLC CSV)", PANEL_MB_LUNG); }
function tmbSolid(){ runTMB_("Solid (Qiagen/CLC CSV)",   PANEL_MB_SOLID); }

/** Ana akış: önce SIKI, sonuç 0 ise GEVŞEK hesapla; raporu yaz */
function runTMB_(panelName, panelMb){
  try{
    const sh = SpreadsheetApp.getActiveSheet();
    const values = sh.getDataRange().getValues();
    if (!values || values.length < 2) throw new Error("Aktif sayfada veri yok.");

    const header = (values[0] || []).map(x => (x||"").toString().trim());
    const idx = buildHeaderIndexMap_(header, HEADER_ALIASES);
    if (idx.chrom < 0 || idx.pos < 0){
      throw new Error('Beklenen CSV sütunları eksik: en az "Chromosome" ve "Start Position" gerekir.');
    }

    // 1) SIKI koş
    const strictRes = collectQualified_(values, idx, STRICT);
    // 2) 0 ise GEVŞEK koş
    let relaxedRes = null;
    if (strictRes.qualified.length === 0){
      relaxedRes = collectQualified_(values, idx, RELAXED);
    }

    // Raporu yaz
    writeTmbReportSheet_("tbm", {
      today: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Istanbul", "yyyy-MM-dd"),
      sampleName: sh.getName(),
      sourceSheet: sh.getName(),
      panelName: panelName,
      panelMb: panelMb,
      strict: buildReportBlock_(strictRes, panelMb, STRICT),
      relaxed: relaxedRes ? buildReportBlock_(relaxedRes, panelMb, RELAXED) : null
    });

    const msg = (strictRes.qualified.length > 0)
      ? `TMB (Sıkı): ${(strictRes.tmb(panelMb)).toFixed(2)}`
      : (relaxedRes && relaxedRes.qualified.length > 0)
          ? `Sıkı=0 çıktı; Gevşek (teknik) TMB: ${(relaxedRes.tmb(panelMb)).toFixed(2)}`
          : `Sıkı=0, Gevşek=0 (veri eşiklere uymuyor).`;
    SpreadsheetApp.getUi().alert("Tamam: " + msg);

  } catch(err){
    SpreadsheetApp.getUi().alert("Hata: " + err.message);
  }
}

/** Tek koşum: satırları parse et, filtrele, teşhis sayımlarını topla */
function collectQualified_(values, idx, conf){
  const bestByLocus = new Map(); // chr:pos -> record (en iyi dp→qual)
  const diag = {
    totalRows: values.length-1,
    parsedOk: 0,
    dropNoALT: 0,
    dropPASS: 0,
    dropQUAL: 0,
    dropDP:   0,
    dropALT:  0,
    dropVAF:  0,
    dropSTR:  0
  };

  for (let r=1; r<values.length; r++){
    const parsed = parseVariantRow_(values[r], idx);
    if (!parsed){ continue; }

    diag.parsedOk++;

    // ALT/VAF/DP eksiklikleri (hesaplanabilen her şeyi hesaplıyoruz)
    // 1) PASS
    if (conf.REQUIRE_PASS){
      const f = (parsed.filter || "").toString().trim().toUpperCase();
      if (f !== "PASS"){ diag.dropPASS++; continue; }
    }

    // 2) QUAL
    if (isNaN(parsed.qual) || parsed.qual < conf.QUAL_MIN){ diag.dropQUAL++; continue; }

    // 3) DP
    if (isNaN(parsed.dp)   || parsed.dp   < conf.DP_MIN){ diag.dropDP++; continue; }

    // 4) ALT
    if (isNaN(parsed.alt)  || parsed.alt  < conf.ALT_MIN){ diag.dropALT++; continue; }

    // 5) VAF
    if (isNaN(parsed.vaf) || parsed.vaf < conf.VAF_MIN || parsed.vaf > conf.VAF_MAX){ diag.dropVAF++; continue; }

    // 6) STR
    if (conf.STR_FILTER && isLikelySTR_(parsed.ref, parsed.altAllele)){ diag.dropSTR++; continue; }

    // 7) ALT gerçekten hiç üretilemediyse (nadiren)
    if (isNaN(parsed.alt)){ diag.dropNoALT++; continue; }

    const key = parsed.chrom + ":" + parsed.pos;
    const prev = bestByLocus.get(key);
    if (!prev){
      bestByLocus.set(key, parsed);
    } else {
      const better = (parsed.dp > prev.dp) || (parsed.dp === prev.dp && parsed.qual > prev.qual);
      if (better) bestByLocus.set(key, parsed);
    }
  }

  const qualified = Array.from(bestByLocus.values());
  const tmb = (panelMb) => (panelMb && panelMb>0) ? (qualified.length / panelMb) : NaN;

  // İlk N için sıralı
  const topList = qualified
    .slice()
    .sort((a,b)=> (b.dp - a.dp) || (b.qual - a.qual))
    .slice(0, TOP_N_TO_LIST);

  return { qualified, topList, diag, conf, tmb };
}

/** Rapor bloğu (metin + liste başlığı + top N satırlar) */
function buildReportBlock_(res, panelMb, conf){
  return {
    label: conf.LABEL,
    thresholds: `QUAL≥${conf.QUAL_MIN}, DP≥${conf.DP_MIN}, ALT≥${conf.ALT_MIN}, VAF≥${conf.VAF_MIN.toFixed(2)}–${conf.VAF_MAX.toFixed(2)}`,
    requirePass: conf.REQUIRE_PASS,
    strFilter: conf.STR_FILTER,
    qualifiedCount: res.qualified.length,
    tmbValue: (panelMb && panelMb>0) ? (res.qualified.length / panelMb) : NaN,
    topList: res.topList,
    diag: res.diag
  };
}

/** Başlık eşlemesi -> index haritası */
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

/** CSV satırı -> normalize varyant */
function parseVariantRow_(row, idx){
  const chrom = safeCell_(row, idx.chrom);
  const pos   = toNumber_(row[idx.pos]);
  if (!chrom || isNaN(pos)) return null;

  let ref = safeCell_(row, idx.ref);
  let altAllele = safeCell_(row, idx.alt);

  // ALT boşsa ins/del’den üret
  if (!altAllele){
    const insVal = safeCell_(row, idx.ins);
    const delVal = safeCell_(row, idx.del);
    if (delVal){ ref = delVal; altAllele = "-"; }
    else if (insVal){ ref = "-"; altAllele = insVal; }
  }

  // DP / ALT / VAF türetme
  let dp = toNumber_(row[idx.dp]);
  let altCount = null;

  // AD: "ref,alt"
  if (idx.ad >= 0){
    const adRaw = (row[idx.ad]||"").toString();
    const nums = adRaw.split(/[;,\s\/\|]+/).map(x => Number(x)).filter(x => !isNaN(x));
    if (nums.length >= 2){
      const refAD = nums[0], altAD = nums[1];
      altCount = Number(altAD);
      if (isNaN(dp) || dp <= 0){
        const sum = nums.reduce((a,b)=>a+b,0);
        if (!isNaN(sum) && sum > 0) dp = sum;
      }
    }
  }

  // Allele read counts: "A(120) G(30)" vb.
  if (altCount == null && idx.alleleCounts >= 0){
    const countsRaw = (row[idx.alleleCounts]||"").toString();
    const matches = countsRaw.match(/\((\d+)\)/g);
    if (matches && matches.length >= 2){
      const nums = matches.map(s => Number(s.replace(/[()]/g,''))).filter(x=>!isNaN(x));
      if (nums.length >= 2){
        altCount = nums[1];
        if (isNaN(dp) || dp <= 0){
          const sum = nums.reduce((a,b)=>a+b,0);
          if (!isNaN(sum) && sum > 0) dp = sum;
        }
      }
    }
  }

  // VAF varsa ALT ≈ round(VAF*DP)
  let vaf = (idx.vaf>=0) ? toNumber_(row[idx.vaf]) : NaN;
  if (!isNaN(vaf) && vaf > 1.0) vaf = vaf / 100.0; // % ise 0-1

  if ((altCount == null || isNaN(altCount)) && !isNaN(vaf) && !isNaN(dp) && dp > 0){
    altCount = Math.round(vaf * dp);
  }

  if (altCount == null) altCount = NaN;

  const qual = (idx.qual>=0) ? toNumber_(row[idx.qual]) : NaN;
  const filter = (idx.filter>=0) ? (row[idx.filter]||"").toString().trim() : "";

  let vafFinal = (!isNaN(vaf)) ? vaf : NaN;
  if (isNaN(vafFinal) && !isNaN(altCount) && !isNaN(dp) && dp>0){
    vafFinal = altCount / dp;
  }

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
  if (/^(A+|T+|C+|G+)$/.test(s) && s.length >= 3) return true;
  if (s1 === s2) return true;
  return false;
}

/** Yardımcılar */
function safeCell_(row, idx){ return (idx>=0 ? (row[idx]||"").toString().trim() : ""); }
function toNumber_(v){
  if (v === null || v === undefined) return NaN;
  if (typeof v === "number") return v;
  const s = v.toString().replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}

/** "tbm" sayfasına RAPOR + TEŞHİS + LİSTE yaz */
function writeTmbReportSheet_(sheetName, all){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clear();

  const lines = [
    "*** TMB RAPORU ***",
    "Tarih: " + all.today,
    "Örnek: " + all.sampleName,
    "Kaynak sayfa: " + all.sourceSheet,
    "Panel: " + all.panelName,
    "Panel boyutu (Mb): " + (all.panelMb != null ? all.panelMb.toFixed(3) : "-"),
    ""
  ];

  // SIKI blok
  lines.push(`[${all.strict.label}] Filtre Eşikleri: ${all.strict.thresholds}`);
  lines.push(`[${all.strict.label}] Filter PASS şartı: ${all.strict.requirePass ? "Evet" : "Hayır"}`);
  lines.push(`[${all.strict.label}] STR artefakt elemesi: ${all.strict.strFilter ? "Evet" : "Hayır"}`);
  lines.push(`[${all.strict.label}] Nitelikli varyant sayısı: ${all.strict.qualifiedCount}`);
  lines.push(`[${all.strict.label}] TMB (varyant/Mb): ` + (isNaN(all.strict.tmbValue) ? "-" : all.strict.tmbValue.toFixed(2)));
  lines.push("");

  // GEVŞEK blok (varsa)
  if (all.relaxed){
    lines.push(`[${all.relaxed.label}] Filtre Eşikleri: ${all.relaxed.thresholds}`);
    lines.push(`[${all.relaxed.label}] Filter PASS şartı: ${all.relaxed.requirePass ? "Evet" : "Hayır"}`);
    lines.push(`[${all.relaxed.label}] STR artefakt elemesi: ${all.relaxed.strFilter ? "Evet" : "Hayır"}`);
    lines.push(`[${all.relaxed.label}] Nitelikli varyant sayısı: ${all.relaxed.qualifiedCount}`);
    lines.push(`[${all.relaxed.label}] TMB (varyant/Mb): ` + (isNaN(all.relaxed.tmbValue) ? "-" : all.relaxed.tmbValue.toFixed(2)));
    lines.push("");
  }

  // Notlar
  lines.push("Notlar:");
  lines.push("- TMB değerleri panel tabanlıdır; klinik raporlama için doğrulanmış panel Mb ve eşikleriniz esas alınmalıdır.");
  lines.push("- CSV’de PASS/AD/VAF yoksa, ALT≈VAF×DP vb. türetmeler yapılmış olabilir.");
  lines.push("- Sıkı koşul 0 dönerse, keşif amaçlı gevşek koşullu teknik TMB de hesaplanır.");
  lines.push("- Germline/benign/silent dışlama yapılmadığından, klinik karar desteği için tek başına kullanılmamalıdır.");
  lines.push("");

  sh.getRange(1,1).setValue(lines.join("\n"));
  sh.getRange(1,1).setWrap(true);

  // Teşhis (A sütununda devam)
  let rowStart = 2 + (lines.length); // yaklaşık; basit yerleşim
  rowStart = Math.max(10, lines.length + 2);

  const diagHeader = [
    ["", "Toplam Satır", "Parse OK", "PASS", "QUAL", "DP", "ALT", "VAF", "STR/Homopol", "ALT yok"]
  ];
  const dS = all.strict.diag, dR = all.relaxed ? all.relaxed.diag : null;
  const diagStrict = [
    ["[Sıkı] Elenen sayıları", dS.totalRows, dS.parsedOk, dS.dropPASS, dS.dropQUAL, dS.dropDP, dS.dropALT, dS.dropVAF, dS.dropSTR, dS.dropNoALT]
  ];
  const diagRelaxed = dR ? [
    ["[Gevşek] Elenen sayıları", dR.totalRows, dR.parsedOk, dR.dropPASS, dR.dropQUAL, dR.dropDP, dR.dropALT, dR.dropVAF, dR.dropSTR, dR.dropNoALT]
  ] : [];

  sh.getRange(rowStart,1,1,diagHeader[0].length).setValues(diagHeader);
  sh.getRange(rowStart+1,1,diagStrict.length,diagHeader[0].length).setValues(diagStrict);
  if (diagRelaxed.length){
    sh.getRange(rowStart+1+diagStrict.length,1,diagRelaxed.length,diagHeader[0].length).setValues(diagRelaxed);
  }

  // Liste başlığı (C sütunundan)
  const listHeader = ["LABEL","CHROM","POS","REF","ALT","QUAL","DP","ALT_AD","VAF"];
  const topRows = [];
  const pushRows = (label, list) => {
    list.forEach(r => {
      topRows.push([
        label, r.chrom, r.pos, r.ref, r.altAllele,
        isNaN(r.qual) ? "" : r.qual,
        isNaN(r.dp)   ? "" : r.dp,
        isNaN(r.alt)  ? "" : r.alt,
        (isNaN(r.vaf) ? "" : r.vaf.toFixed(3))
      ]);
    });
  };

  pushRows("[Sıkı]", all.strict.topList);
  if (all.relaxed) pushRows("[Gevşek]", all.relaxed.topList);

  sh.getRange(1,3,1,listHeader.length).setValues([listHeader]);
  if (topRows.length){
    sh.getRange(2,3,topRows.length,listHeader.length).setValues(topRows);
  }

  sh.autoResizeColumns(1, 12);
}
