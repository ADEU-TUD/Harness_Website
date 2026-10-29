// Reading attachments (PDF, Word, PowerPoint, Excel, OpenDocument, text, images) and writing Word/Excel files.
/* global JSZip, pdfjsLib */

const MAX_TEXT = 120000;
const MAX_FILE = 25 * 1024 * 1024;

if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";

const xml = (text) => new DOMParser().parseFromString(text, "application/xml");
const ext = (name) => (name.includes(".") ? name.split(".").pop().toLowerCase() : "");

/** Reads [file]: {kind: "image", name, url} or {kind: "file", name, text, truncated, note}. */
export async function readAttachment(file) {
  if (file.size > MAX_FILE) throw new Error(`„${file.name}“ ist größer als 25 MB.`);
  if (file.type.startsWith("image/")) return { kind: "image", name: file.name, url: await imageUrl(file) };
  const e = ext(file.name);
  let text;
  let note = null;
  if (e === "pdf") {
    text = await pdfText(file);
    if (text.replace(/--- Seite \d+ ---/g, "").trim().length < 20) note = "Die PDF-Datei enthält kaum Text (Scan?) – Bilder von Seiten gehen in der Browserversion noch nicht.";
  } else if (e === "docx") text = await docxText(file);
  else if (e === "pptx") text = await pptxText(file);
  else if (e === "xlsx") text = await xlsxText(file);
  else if (["odt", "odp", "ods"].includes(e)) text = await odfText(file);
  else if (["doc", "ppt", "xls"].includes(e)) throw new Error(`„${file.name}“ ist im alten Office-Format – bitte als .${e}x speichern.`);
  else text = await file.text();
  const truncated = text.length > MAX_TEXT;
  return { kind: "file", name: file.name, text: truncated ? text.slice(0, MAX_TEXT) : text, truncated, note };
}

async function imageUrl(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = reject;
      i.src = url;
    });
    const scale = Math.min(1, 1600 / Math.max(img.width, img.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(img.width * scale);
    canvas.height = Math.round(img.height * scale);
    canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL("image/jpeg", 0.85);
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function pdfText(file) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pages = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const content = await (await pdf.getPage(n)).getTextContent();
    let line = "";
    const lines = [];
    for (const item of content.items) {
      line += item.str;
      if (item.hasEOL) { lines.push(line); line = ""; }
    }
    if (line) lines.push(line);
    pages.push(`--- Seite ${n} ---\n` + lines.join("\n"));
    if (pages.join("\n").length > MAX_TEXT) break;
  }
  return pages.join("\n\n");
}

function texts(node, tag) {
  return Array.from(node.getElementsByTagName(tag)).map((t) => t.textContent).join("");
}

async function docxText(file) {
  const zip = await JSZip.loadAsync(file);
  const doc = xml(await zip.file("word/document.xml").async("string"));
  const body = doc.getElementsByTagName("w:body")[0];
  const out = [];
  for (const el of Array.from(body ? body.children : [])) {
    if (el.tagName === "w:p") {
      const text = texts(el, "w:t");
      if (!text.trim()) continue;
      const style = el.getElementsByTagName("w:pStyle")[0]?.getAttribute("w:val") || "";
      const level = /^(Heading|berschrift)(\d)/i.exec(style.replace("Ü", ""))?.[2] || (/^Title$/i.test(style) ? "1" : null);
      const list = el.getElementsByTagName("w:numPr").length > 0;
      out.push(level ? "#".repeat(Number(level)) + " " + text : list ? "- " + text : text);
    } else if (el.tagName === "w:tbl") {
      const rows = Array.from(el.getElementsByTagName("w:tr")).map((tr) =>
        Array.from(tr.getElementsByTagName("w:tc")).map((tc) => texts(tc, "w:t").replace(/\|/g, "/")));
      out.push(markdownTable(rows));
    }
  }
  return out.join("\n\n");
}

async function pptxText(file) {
  const zip = await JSZip.loadAsync(file);
  const num = (p) => Number(/(\d+)\.xml$/.exec(p)?.[1] || 0);
  const slides = Object.keys(zip.files).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p)).sort((a, b) => num(a) - num(b));
  const out = [];
  for (const path of slides) {
    const doc = xml(await zip.file(path).async("string"));
    const paras = Array.from(doc.getElementsByTagName("a:p")).map((p) => texts(p, "a:t")).filter((t) => t.trim());
    let block = `## Folie ${num(path)}${paras.length ? ": " + paras[0] : ""}\n` + paras.slice(1).map((t) => "- " + t).join("\n");
    const notes = zip.file(`ppt/notesSlides/notesSlide${num(path)}.xml`);
    if (notes) {
      const n = Array.from(xml(await notes.async("string")).getElementsByTagName("a:p")).map((p) => texts(p, "a:t")).filter((t) => t.trim() && !/^\d+$/.test(t.trim()));
      if (n.length) block += "\nNotizen: " + n.join(" ");
    }
    out.push(block);
  }
  return out.join("\n\n");
}

async function xlsxText(file) {
  const zip = await JSZip.loadAsync(file);
  const shared = zip.file("xl/sharedStrings.xml")
    ? Array.from(xml(await zip.file("xl/sharedStrings.xml").async("string")).getElementsByTagName("si")).map((si) => texts(si, "t"))
    : [];
  const book = xml(await zip.file("xl/workbook.xml").async("string"));
  const names = Array.from(book.getElementsByTagName("sheet")).map((s) => s.getAttribute("name"));
  const sheets = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))
    .sort((a, b) => Number(/(\d+)\.xml$/.exec(a)[1]) - Number(/(\d+)\.xml$/.exec(b)[1]));
  const out = [];
  for (const [i, path] of sheets.entries()) {
    const doc = xml(await zip.file(path).async("string"));
    const rows = Array.from(doc.getElementsByTagName("row")).slice(0, 500).map((row) => {
      const cells = [];
      for (const c of Array.from(row.getElementsByTagName("c"))) {
        const col = colIndex(c.getAttribute("r") || "");
        const v = c.getElementsByTagName("v")[0]?.textContent ?? texts(c, "t");
        cells[col >= 0 ? col : cells.length] = c.getAttribute("t") === "s" ? shared[Number(v)] ?? "" : v;
      }
      return Array.from(cells, (x) => (x ?? "").replace(/\|/g, "/"));
    });
    out.push(`## ${names[i] || "Tabelle " + (i + 1)}\n\n` + markdownTable(rows));
  }
  return out.join("\n\n");
}

function colIndex(ref) {
  const letters = /^[A-Z]+/.exec(ref)?.[0];
  if (!letters) return -1;
  return [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

async function odfText(file) {
  const zip = await JSZip.loadAsync(file);
  const doc = xml(await zip.file("content.xml").async("string"));
  return Array.from(doc.getElementsByTagName("text:p")).map((p) => p.textContent).filter((t) => t.trim()).join("\n");
}

function markdownTable(rows) {
  const width = Math.max(1, ...rows.map((r) => r.length));
  const line = (r) => "| " + Array.from({ length: width }, (_, i) => (r[i] ?? "").replace(/\n/g, " ")).join(" | ") + " |";
  if (!rows.length) return "";
  return [line(rows[0]), "| " + Array(width).fill("---").join(" | ") + " |", ...rows.slice(1).map(line)].join("\n");
}

// ------------------------------------------------------------------ writing

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Markdown tables of [md] as rows of cells. */
export function tablesOf(md) {
  const tables = [];
  let cur = null;
  let title = null;
  for (const raw of md.split("\n")) {
    const line = raw.trim();
    if (line.startsWith("#")) title = line.replace(/^#+\s*/, "");
    if (line.startsWith("|")) {
      if (/^\|[\s:|-]+\|$/.test(line)) continue;
      if (!cur) { cur = { title: title || `Tabelle ${tables.length + 1}`, rows: [] }; tables.push(cur); }
      cur.rows.push(line.replace(/^\||\|$/g, "").split("|").map((c) => c.trim()));
    } else if (line) cur = null;
  }
  return tables;
}

/** Inline Markdown (**fett**, *kursiv*, `code`) as Word runs. */
function runs(text) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g;
  let last = 0;
  let m;
  const run = (t, props = "") => out.push(`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`);
  while ((m = re.exec(text))) {
    if (m.index > last) run(text.slice(last, m.index));
    const t = m[0];
    if (t.startsWith("**")) run(t.slice(2, -2), "<w:b/>");
    else if (t.startsWith("`")) run(t.slice(1, -1), '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>');
    else run(t.slice(1, -1), "<w:i/>");
    last = m.index + t.length;
  }
  if (last < text.length) run(text.slice(last));
  return out.join("");
}

const plain = (s) => s.replace(/\[([^\]]+)\]\([^)]+\)/g, "$1");

/** A Word document from Markdown: headings, paragraphs, lists, tables, quotes. */
export async function markdownToDocx(md) {
  const body = [];
  const lines = md.replace(/\r/g, "").split("\n");
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const t = line.trim();
    if (!t) continue;
    if (t.startsWith("|")) {
      const block = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) block.push(lines[i++].trim());
      i--;
      const rows = block.filter((l) => !/^\|[\s:|-]+\|$/.test(l)).map((l) => l.replace(/^\||\|$/g, "").split("|").map((c) => c.trim()));
      const cell = (c, head) => `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p>${runs(plain(head ? `**${c}**` : c))}</w:p></w:tc>`;
      body.push(`<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${
        ["top", "left", "bottom", "right", "insideH", "insideV"].map((b) => `<w:${b} w:val="single" w:sz="4" w:color="BFBFBF"/>`).join("")
      }</w:tblBorders></w:tblPr>${rows.map((r, n) => `<w:tr>${r.map((c) => cell(c, n === 0)).join("")}</w:tr>`).join("")}</w:tbl><w:p/>`);
      continue;
    }
    const h = /^(#{1,3})\s+(.*)$/.exec(t);
    if (h) { body.push(`<w:p><w:pPr><w:pStyle w:val="Heading${h[1].length}"/></w:pPr>${runs(plain(h[2]))}</w:p>`); continue; }
    const li = /^([-*+]|\d+[.)])\s+(?:\[[ xX]\]\s+)?(.*)$/.exec(t);
    if (li) {
      const indent = Math.min(3, Math.floor((line.length - line.trimStart().length) / 2));
      const mark = /\d/.test(li[1]) ? li[1] : "•";
      body.push(`<w:p><w:pPr><w:ind w:left="${360 + indent * 360}" w:hanging="360"/></w:pPr>${runs(mark + "\t" + plain(li[2]))}</w:p>`);
      continue;
    }
    if (t.startsWith(">")) { body.push(`<w:p><w:pPr><w:ind w:left="567"/></w:pPr>${runs(plain(t.replace(/^>\s?/, "")))}</w:p>`); continue; }
    if (/^(-{3,}|\*{3,})$/.test(t)) { body.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="BFBFBF"/></w:pBdr></w:pPr></w:p>'); continue; }
    body.push(`<w:p>${runs(plain(t))}</w:p>`);
  }
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  const heading = (n, size) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:color w:val="00305D"/><w:sz w:val="${size}"/></w:rPr></w:style>`;
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${heading(1, 32)}${heading(2, 26)}${heading(3, 23)}<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style></w:styles>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

/** An Excel file: every Markdown table of [md] becomes a sheet; numbers stay numbers. */
export async function markdownToXlsx(md) {
  const tables = tablesOf(md);
  if (!tables.length) tables.push({ title: "Tabelle 1", rows: md.split("\n").filter((l) => l.trim()).map((l) => l.split(/\t|;/)) });
  const col = (i) => { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const used = new Set();
  const sheetName = (t) => {
    let n = t.replace(/[\\/?*[\]:]/g, " ").trim().slice(0, 28) || "Tabelle";
    let k = 2;
    const base = n;
    while (used.has(n.toLowerCase())) n = `${base} ${k++}`;
    used.add(n.toLowerCase());
    return n;
  };
  const zip = new JSZip();
  const names = tables.map((t) => sheetName(t.title));
  tables.forEach((t, s) => {
    const rows = t.rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => {
      const ref = `${col(ci)}${ri + 1}`;
      // German numbers: 1.234,5 and 3,5 → 1234.5 and 3.5; 1.000 is a thousand.
      const num = v.replace(/\s/g, "")
        .replace(/^(-?\d{1,3}(\.\d{3})+|-?\d+),(\d+)$/, (_, a, __, b) => a.replace(/\./g, "") + "." + b)
        .replace(/^(-?\d{1,3})((\.\d{3})+)$/, (_, a, b) => a + b.replace(/\./g, ""));
      if (ri > 0 && /^-?\d+(\.\d+)?$/.test(num)) return `<c r="${ref}"><v>${num}</v></c>`;
      return `<c r="${ref}" t="inlineStr"${ri === 0 ? ' s="1"' : ""}><is><t xml:space="preserve">${esc(plain(v.replace(/\*\*/g, "")))}</t></is></c>`;
    }).join("")}</row>`).join("");
    zip.file(`xl/worksheets/sheet${s + 1}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>${rows}</sheetData></worksheet>`);
  });
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${names.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${names.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`);
  zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
