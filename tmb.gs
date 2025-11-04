/*  TMB Raporlama – Qiagen/CLC CSV (Google Sheets)
 *  @author: you
 *  Not: TR yerelinde ondalık virgül gelirse otomatik nokta çevrilir.
 */

const PANELS = {
  "SOLID_3204": 0.231,   // CDHS-53205Z-3204 konservatif ~0.231 Mb
  "LUNG_3833": 0.280,    // CDHS-53206Z-3833 konservatif ~0.280 Mb
};

const HEADER_ALIASES = {
  chrom: ["Chromosome", "#CHROM", "CHROM"],
  pos:   ["Start Position", "POS", "Start"],
  ref:   ["Reference Allele", "REF"],
  alt:   ["Sample Allele", "ALT"],
  qual:  ["Sample Call Quality", "QUAL"],
  dp:    ["Sample Read Depth", "DP"],
  vaf:   ["Sample Allele Fraction", "VAF", "AF", "Allele Frequency"],
  gt:    ["Sample Genotype", "GT"]
};

const DEFAULT_THRESH = { QUAL:50, DP:100, ALT:5, VAF:0.05 };

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("TMB")
    .addItem("Ayar sayfasını hazırla", "ensureSettings")
    .addItem("Aktif sayfadan TMB hesapla", "runTMBFromActiveSheet")
    .addToUi();
}

function ensureSettings() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Ayarlar");
  if (!sh) sh = ss.insertSheet("Ayarlar");
  sh.clear();

  sh.getRange("A1").setValue("Parametre");
  sh.getRange("B1").setValue("Değer");

  const rows = [
    ["Panel", "SOLID_3204"],             // B2 seçmeli
    ["Panel Mb (CUSTOM için)", ""],      // B3
    ["QUAL eşik", DEFAULT_THRESH.QUAL],  // B4
    ["DP eşik", DEFAULT_THRESH.DP],      // B5
    ["ALT (yaklaşık) eşik", DEFAULT_THRESH.ALT], // B6
    ["VAF eşik", DEFAULT_THRESH.VAF],    // B7
    ["Filter PASS gerekli mi?", "Evet"]  // B8 (CSV’de 'Filter' yoksa yok sayılır)
  ];
  sh.getRange(2,1,rows.length,2).setValues(rows);

  // Panel seçimi için doğrulama
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["SOLID_3204","LUNG_3833","CUSTOM"], true)
    .setAllowInvalid(false).build();
  sh.getRange("B2").setDataValidation(rule);

  SpreadsheetApp.getUi().alert("Ayarlar sayfası hazır.");
}

function runTMBFromActiveSheet() {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ss.getActiveSheet();
  if (dataSheet.getName() === "Ayarlar" || dataSheet.getName() === "TMB") {
    SpreadsheetApp.getUi().alert("Lütfen CSV’nin bulunduğu veri sayfasını aktif yapın.");
    return;
  }
  const cfg = readConfig_();
  const result = computeTMB_(dataSheet, cfg);
  writeTmbSheet_(ss, dataSheet.getName(), cfg, result);
  SpreadsheetApp.getUi().alert("TMB hesaplandı ve ‘TMB’ sayfasına yazıldı.");
}

function readConfig_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Ayarlar");
  if (!sh) { ensureSettings(); sh = ss.getSheetByName("Ayarlar"); }

  const kv = Object.fromEntries(
    sh.getRange(2,1,10,2).getValues().filter(r => r[0])
  );
  const panel = (kv["Panel"] || "SOLID_3204").toString().trim();
  let panelMb = PANELS[panel] || null;
  if (panel === "CUSTOM") {
    const v = (kv["Panel Mb (CUSTOM için)"] || "").toString().trim().replace(",", ".");
    panelMb = v ? parseFloat(v) : null;
  }
  const thresh = {
    QUAL: Number(kv["QUAL eşik"] || DEFAULT_THRESH.QUAL),
    DP:   Number(kv["DP eşik"]   || DEFAULT_THRESH.DP),
    ALT:  Number(kv["ALT (yaklaşık) eşik"] || DEFAULT_THRESH.ALT),
    VAF:  Number((kv["VAF eşik"] || DEFAULT_THRESH.VAF).toString().replace(",", ".")),
    PASS: ((kv["Filter PASS gerekli mi?"] || "Evet")+"").toLowerCase().startsWith("e")
  };
  return { panel, panelMb, thresh };
}

function computeTMB_(sheet, cfg) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("Veri bulunamadı.");

  const headers = values[0].map(h => (h||"").toString().trim());
  const map = headerMap_(headers);

  const idx = (name) => map[name] ?? -1;
  const iChrom = idx("chrom"), iPos = idx("pos"), iRef = idx("ref"), iAlt = idx("alt");
  const iQual  = idx("qual"),  iDp  = idx("dp"),  iVaf = idx("vaf"), iGt = idx("gt");

  // En azından CHROM, POS, REF/ALT, QUAL/DP/VAF’den bazıları bulunmalı
  if (iChrom<0 || iPos<0 || (iRef<0 && iAlt<0)) {
    throw new Error("Beklenen VCF/CSV sütunları bulunamadı (Chromosome, Start Position, Reference/Sample Allele).");
  }

  // “Filter” kolonu varsa PASS kontrol edelim
  const iFilter = headers.findIndex(h => /filter/i.test(h));

  const passRows = [];
  const seen = new Set(); // CHROM:POS uniqleştirme

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row[iChrom] && !row[iPos]) continue;

    const key = `${row[iChrom]}:${row[iPos]}`;
    if (seen.has(key)) continue;
    seen.add(key);

    const qual = parseNum_(row[iQual]);
    const dp   = parseNum_(row[iDp]);
    let vaf    = parseNum_(row[iVaf]);
    const gt   = (iGt>=0 ? (row[iGt]||"").toString() : "");

    // ALT var mı?
    const altVal = (iAlt>=0 ? (row[iAlt]||"").toString() : "");
    const hasAlt = (gt && /1/.test(gt)) || (altVal && altVal !== ".");

    // VAF yoksa ALT tahmini yapılamaz → satırı atla
    if (!hasAlt) continue;

    // Eşikler
    if (qual!=null && qual < cfg.thresh.QUAL) continue;
    if (dp!=null   && dp   < cfg.thresh.DP)   continue;

    // VAF normalize (yüzde gelmişse)
    if (vaf != null && vaf > 1.0) vaf = vaf/100.0;

    // ALT ~ VAF*DP
    let altApprox = null;
    if (vaf!=null && dp!=null) altApprox = Math.round(vaf * dp);
    if (vaf!=null && vaf < cfg.thresh.VAF) continue;
    if (altApprox!=null && altApprox < cfg.thresh.ALT) continue;

    // PASS filtresi varsa
    if (cfg.thresh.PASS && iFilter>=0) {
      const fl = (row[iFilter]||"").toString().toUpperCase();
      if (fl && fl!=="PASS" && fl!=="." ) continue;
    }

    passRows.push({
      chrom: row[iChrom], pos: row[iPos],
      ref: iRef>=0 ? row[iRef] : "",
      alt: altVal,
      qual, dp, vaf, altApprox
    });
  }

  const qualified = passRows.length;
  const panelMb = cfg.panelMb;
  const tmb = panelMb ? (qualified / panelMb) : null;

  // İlk 20
  const top = passRows.slice(0,20).map(v => {
    const vafStr = (v.vaf!=null) ? v.vaf.toFixed(3) : "";
    return `${v.chrom}:${v.pos} ${v.ref}>${v.alt} | ${numStr_(v.qual)} | ${numStr_(v.dp)} | ${numStr_(v.altApprox)} | ${vafStr}`;
  });

  return { qualified, tmb, top, countTotal: seen.size };
}

function writeTmbSheet_(ss, sourceSheetName, cfg, res) {
  let sh = ss.getSheetByName("TMB");
  if (!sh) sh = ss.insertSheet("TMB");
  sh.clear();

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const panelMbTxt = cfg.panelMb ? cfg.panelMb.toFixed(3) : "—";

  const info = [
    ["*** TMB RAPORU ***",""],
    ["Tarih", today],
    ["Kaynak sayfa", sourceSheetName],
    ["Panel", cfg.panel],
    ["Panel boyutu (Mb)", panelMbTxt],
    ["Filtre Eşikleri", `QUAL≥${cfg.thresh.QUAL}, DP≥${cfg.thresh.DP}, ALT≥${cfg.thresh.ALT}, VAF≥${cfg.thresh.VAF}`],
    ["Filter PASS şartı", cfg.thresh.PASS ? "Evet" : "Hayır"],
    ["Nitelikli varyant sayısı", res.qualified],
    ["TMB (varyant/Mb)", (res.tmb!=null) ? res.tmb.toFixed(2) : "Panel Mb girilmeli"]
  ];
  sh.getRange(1,1,info.length,2).setValues(info);

  sh.getRange(12,1).setValue("İlk 20 nitelikli varyant (CHROM:POS REF>ALT | QUAL | DP | ALT≈VAF×DP | VAF):");
  if (res.top.length) {
    const rows = res.top.map(s => [s]);
    sh.getRange(13,1,rows.length,1).setValues(rows);
  }

  const aciklama =
`Notlar:
- Bu değer panel tabanlı, eşiklerle filtrelenmiş ham TMB’dir (anotasyon yapılmadı).
- Klinik rapor için kurumunuzun doğrulanmış panel Mb ve eşikleri kullanılmalıdır.
- PASS/DP/QUAL/VAF/ALT koşulları basit kalite kontrolleridir; STR/mikrosatellit artefaktları anotasyon yapılmadan tam ayıklanamaz.
- Aynı pozisyondaki çoklu alleller konuma göre benzersiz kabul edilmiştir.`;

  sh.getRange(9,1,1,2).setNote(aciklama);
  sh.autoResizeColumns(1,2);
}

function headerMap_(headers) {
  const map = {};
  for (const key in HEADER_ALIASES) {
    const idx = headers.findIndex(h => HEADER_ALIASES[key].some(a => a.toLowerCase() === h.toLowerCase()));
    if (idx >= 0) map[key] = idx;
  }
  return map;
}

function parseNum_(v) {
  if (v === null || v === "" || v === undefined) return null;
  const s = (""+v).replace(",", ".").replace(/\s+/g,"").trim();
  const n = Number(s);
  return isNaN(n) ? null : n;
}
function numStr_(v){ return (v==null ? "" : (typeof v==="number" ? v.toString() : (""+v))); }
