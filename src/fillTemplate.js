// Vyplnění firemní šablony přímou úpravou XML uvnitř .xlsx souboru.
// Na rozdíl od ExcelJS nezasahuje do ničeho jiného (podmíněné formátování,
// ověření dat, vzorce i styly zůstávají bajt po bajtu zachované).

let _jszip = null;
async function getJSZip() {
  if (!_jszip) {
    const mod = await import("jszip");
    _jszip = mod.default || mod;
  }
  return _jszip;
}

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const colToNum = (letters) => letters.split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);

// Nastaví hodnotu buňky v XML listu. value: null = vyprázdnit, number = číslo, string = text (inline).
// Zachovává atribut stylu (s="...") existující buňky.
export function setCellInXml(xml, ref, value) {
  const re = new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>[\\s\\S]*?</c>)');
  const m = xml.match(re);

  const build = (attrs) => {
    const a = attrs.replace(/\s+t="[^"]*"/g, ""); // původní typ pryč, nastavíme vlastní
    if (value === null || value === undefined || value === "") {
      return '<c r="' + ref + '"' + a + "/>";
    }
    if (typeof value === "number") {
      return '<c r="' + ref + '"' + a + "><v>" + value + "</v></c>";
    }
    return (
      '<c r="' + ref + '"' + a + ' t="inlineStr"><is><t xml:space="preserve">' +
      xmlEscape(value) +
      "</t></is></c>"
    );
  };

  if (m) return xml.replace(re, build(m[1]));

  // Buňka v XML není (prázdné buňky bez stylu se nezapisují) — vložit do řádku
  if (value === null || value === undefined || value === "") return xml;
  const rowNum = ref.match(/\d+/)[0];
  const myCol = colToNum(ref.match(/[A-Z]+/)[0]);
  const rowRe = new RegExp('(<row r="' + rowNum + '"[^>]*>)([\\s\\S]*?)(</row>)');
  const rm = xml.match(rowRe);
  if (!rm) throw new Error("Řádek " + rowNum + " v listu nenalezen.");
  const cellRe = /<c r="([A-Z]+)(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
  let insertAt = rm[2].length;
  let cm;
  while ((cm = cellRe.exec(rm[2]))) {
    if (colToNum(cm[1]) > myCol) {
      insertAt = cm.index;
      break;
    }
  }
  const inner = rm[2].slice(0, insertAt) + build("") + rm[2].slice(insertAt);
  return xml.replace(rowRe, rm[1] + inner + rm[3]);
}

// Najde cestu k XML souboru listu podle jeho názvu
async function sheetPath(zip, sheetName) {
  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  let rId = null;
  for (const tag of wbXml.match(/<sheet\b[^>]*\/?>/g) || []) {
    const nm = tag.match(/name="([^"]*)"/);
    const id = tag.match(/r:id="(rId\d+)"/);
    if (nm && id && nm[1] === sheetName) {
      rId = id[1];
      break;
    }
  }
  if (!rId) throw new Error(`List ${sheetName} v šabloně nenalezen.`);
  const rel = (relsXml.match(new RegExp('<Relationship [^>]*Id="' + rId + '"[^>]*/>')) || [])[0];
  const target = rel && rel.match(/Target="([^"]*)"/);
  if (!target) throw new Error(`Cesta k listu ${sheetName} nenalezena.`);
  return "xl/" + target[1].replace(/^\//, "").replace(/^xl\//, "");
}

/**
 * Vyplní měsíční list šablony a vrátí ArrayBuffer hotového souboru.
 * @param {ArrayBuffer} templateBuf  - obsah šablony
 * @param {string} sheetName         - název listu (např. "Červenec")
 * @param {string} employeeName      - jméno do O2
 * @param {Object} dayCells          - { [den: number]: {E,F,G,H,I,J} } (null hodnoty se přeskočí)
 */
export async function fillOfficialSheet(templateBuf, sheetName, employeeName, dayCells) {
  const JSZip = await getJSZip();
  const zip = await JSZip.loadAsync(templateBuf);
  const path = await sheetPath(zip, sheetName);
  let xml = await zip.file(path).async("string");

  // Jméno zaměstnance
  xml = setCellInXml(xml, "O2", employeeName);

  // Vyčistit ručně vyplňované sloupce E–J (řádky 7–37 = dny 1–31)
  for (let r = 7; r <= 37; r++) {
    for (const col of ["E", "F", "G", "H", "I", "J"]) {
      xml = setCellInXml(xml, col + r, null);
    }
  }

  // Vyplnit zadané dny
  for (const [day, cells] of Object.entries(dayCells)) {
    const r = 6 + parseInt(day);
    if (r < 7 || r > 37) continue;
    for (const col of ["E", "F", "G", "H", "I", "J"]) {
      const v = cells[col];
      if (v !== null && v !== undefined) xml = setCellInXml(xml, col + r, v);
    }
  }

  zip.file(path, stripCachedValues(xml));
  await finalizeWorkbook(zip);
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

// ---------- Zahraniční pracovní cesta ----------

// Odstraní z buněk se vzorcem uložené výsledky, aby se nikde neukázala
// stará hodnota ze šablony — čtečka musí vzorec spočítat znovu.
export function stripCachedValues(xml) {
  return xml.replace(/<c\b[^>]*>[\s\S]*?<\/c>/g, (cell) => {
    if (cell.indexOf("<f") === -1) return cell;
    return cell.replace(/<v>[\s\S]*?<\/v>/g, "");
  });
}

// Odstraní calcChain a vynutí přepočet všech vzorců při otevření souboru.
// Bez toho by Excel mohl zobrazit staré uložené hodnoty ze šablony.
async function finalizeWorkbook(zip) {
  zip.remove("xl/calcChain.xml");
  const ct = zip.file("[Content_Types].xml");
  if (ct) {
    let ctXml = await ct.async("string");
    ctXml = ctXml.replace(/<Override[^>]*calcChain\.xml"[^>]*\/>/g, "");
    zip.file("[Content_Types].xml", ctXml);
  }
  const wbFile = zip.file("xl/workbook.xml");
  if (wbFile) {
    let wbXml = await wbFile.async("string");
    if (/<calcPr[^>]*\/>/.test(wbXml)) {
      wbXml = wbXml.replace(/<calcPr([^>]*)\/>/, (m, attrs) => {
        const a = attrs.replace(/\s+fullCalcOnLoad="[^"]*"/g, "");
        return `<calcPr${a} fullCalcOnLoad="1"/>`;
      });
    } else {
      wbXml = wbXml.replace(/<\/workbook>/, '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    zip.file("xl/workbook.xml", wbXml);
  }
}

// ISO datum -> Excel sériové číslo (epocha 1899-12-30)
function dateToSerial(iso) {
  const d = new Date(iso + "T00:00:00Z");
  return Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
}

// "7:30" -> zlomek dne
function timeToFraction(t) {
  if (!t) return null;
  const m = String(t).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return (parseInt(m[1]) * 60 + parseInt(m[2])) / 1440;
}

/**
 * Vyplní formulář zahraniční pracovní cesty a vrátí ArrayBuffer hotového souboru.
 * List "Vzor": C7=jméno, C8=účel, C9=cíl, C10=země, J7/J8=začátek/konec cesty,
 * J9=doprava, C13=kurz ČNB, J13=kapesné, J15=záloha,
 * řádky 29–42: B=datum, C=místo pobytu, D=od, E=do, F=bezplatná jídla.
 */
export async function fillForeignTrip(templateBuf, trip) {
  const JSZip = await getJSZip();
  const zip = await JSZip.loadAsync(templateBuf);
  const path = await sheetPath(zip, "Vzor");
  let xml = await zip.file(path).async("string");

  const set = (ref, val) => { xml = setCellInXml(xml, ref, val); };

  set("C7", trip.employeeName || null);
  set("C8", trip.ucel || null);
  set("C9", trip.cil || null);
  set("C10", trip.zeme || null);
  set("J7", trip.zacatek ? dateToSerial(trip.zacatek) : null);
  set("J8", trip.konec ? dateToSerial(trip.konec) : null);
  set("J9", trip.doprava || null);
  set("C13", trip.kurz ? Number(trip.kurz) : null);
  set("J13", trip.kapesne ? Number(trip.kapesne) : 0);
  set("J15", trip.zaloha ? Number(trip.zaloha) : 0);

  // Denní řádky 29–42 (max 14 dní)
  const days = (trip.days || []).slice(0, 14);
  for (let i = 0; i < 14; i++) {
    const r = 29 + i;
    const d = days[i];
    if (!d) {
      ["B", "C", "D", "E", "F"].forEach((c) => set(c + r, null));
      continue;
    }
    set("B" + r, d.date ? dateToSerial(d.date) : null);
    set("C" + r, d.misto || null);
    set("D" + r, timeToFraction(d.od));
    set("E" + r, timeToFraction(d.do));
    set("F" + r, d.jidla ? Number(d.jidla) : null);
  }

  zip.file(path, stripCachedValues(xml));
  await finalizeWorkbook(zip);
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

/**
 * Načte tabulku zemí ze šablony (list Sazby): B=země, C=měnový kód, G=sazba 2025.
 */
export async function readCountryRates(templateBuf) {
  const JSZip = await getJSZip();
  const zip = await JSZip.loadAsync(templateBuf);
  const path = await sheetPath(zip, "Sazby");
  const xml = await zip.file(path).async("string");

  let shared = [];
  const ssFile = zip.file("xl/sharedStrings.xml");
  if (ssFile) {
    const ss = await ssFile.async("string");
    shared = (ss.match(/<si>[\s\S]*?<\/si>/g) || []).map((si) =>
      (si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [])
        .map((t) => t.replace(/<[^>]+>/g, ""))
        .join("")
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
    );
  }

  const cellVal = (rowXml, col) => {
    const m = rowXml.match(new RegExp('<c r="' + col + '\\d+"([^>]*)>([\\s\\S]*?)</c>'));
    if (!m) return null;
    const isShared = /t="s"/.test(m[1]);
    const v = (m[2].match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    if (v == null) {
      const inline = (m[2].match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1];
      return inline != null ? inline : null;
    }
    return isShared ? shared[parseInt(v)] : v;
  };

  const out = [];
  for (const rowXml of xml.match(/<row[^>]*>[\s\S]*?<\/row>/g) || []) {
    const rNum = parseInt((rowXml.match(/<row r="(\d+)"/) || [])[1] || "0");
    if (rNum < 9 || rNum > 183) continue;
    const zeme = cellVal(rowXml, "B");
    const mena = cellVal(rowXml, "C");
    const sazba = cellVal(rowXml, "G");
    if (zeme && mena && sazba) out.push({ zeme, mena, sazba: parseFloat(sazba) });
  }
  return out;
}
