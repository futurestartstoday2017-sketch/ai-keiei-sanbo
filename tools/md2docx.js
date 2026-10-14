// Markdown（本プロジェクトのレポート形式）を Word (.docx) に変換する
// 使い方: node md2docx.js input.md output.docx
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, HeadingLevel,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  AlignmentType, LevelFormat, Footer, PageNumber, TableLayoutType,
} = require("docx");

const [, , inPath, outPath] = process.argv;
const lines = fs.readFileSync(inPath, "utf8").replace(/\r\n/g, "\n").split("\n");

const FONT = { ascii: "Yu Gothic", hAnsi: "Yu Gothic", eastAsia: "Yu Gothic", cs: "Yu Gothic" };
const ACCENT = "1F4E79";
const PAGE_W = 11906, MARGIN = 1020, CONTENT_W = PAGE_W - MARGIN * 2;
const BODY_SIZE = 20, TABLE_SIZE = 17; // half-points

// ---------- inline ----------
function inlineRuns(text, base = {}) {
  const runs = [];
  const re = /\*\*(.+?)\*\*|\[([^\]]+)\]\(([^)]+)\)|`([^`]+)`/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    if (m[1] !== undefined) {
      runs.push(...inlineRuns(m[1], { ...base, bold: true }));
    } else if (m[2] !== undefined) {
      const url = m[3];
      if (/^https?:\/\//.test(url)) {
        runs.push(new ExternalHyperlink({
          link: url,
          children: [new TextRun({ text: m[2], ...base, style: "Hyperlink" })],
        }));
      } else {
        runs.push(new TextRun({ text: m[2], ...base }));
      }
    } else if (m[4] !== undefined) {
      runs.push(new TextRun({ text: m[4], ...base, font: "Consolas", color: "6A1B9A" }));
    }
    last = re.lastIndex;
  }
  if (last < text.length) runs.push(new TextRun({ text: text.slice(last), ...base }));
  return runs;
}

// ---------- tables ----------
const splitRow = (l) => l.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());
const visLen = (s) => s.replace(/\*\*|`/g, "").replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
  .split(/<br>/).reduce((a, p) => a + [...p].reduce((n, ch) => n + (ch.charCodeAt(0) > 255 ? 2 : 1), 0), 0);

function buildTable(rows) {
  const header = splitRow(rows[0]);
  const body = rows.slice(2).map(splitRow);
  const nCol = header.length;
  // 列幅：内容量の平方根に比例（極端な偏りを抑える）＋最小幅
  const weights = [];
  for (let c = 0; c < nCol; c++) {
    const maxLen = Math.max(visLen(header[c] || ""), ...body.map((r) => visLen(r[c] || "")));
    weights.push(Math.max(Math.sqrt(maxLen), 5.6));
  }
  const total = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((w / total) * CONTENT_W));
  widths[nCol - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);

  const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const mkCell = (text, c, isHeader) => new TableCell({
    width: { size: widths[c], type: WidthType.DXA },
    borders,
    shading: isHeader ? { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: (text || "").split(/<br>/).map((part) => new Paragraph({
      spacing: { after: 20, line: 276 },
      children: inlineRuns(part, isHeader
        ? { bold: true, color: "FFFFFF", size: TABLE_SIZE }
        : { size: TABLE_SIZE }),
    })),
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, c) => mkCell(h, c, true)) }),
      ...body.map((r) => new TableRow({
        cantSplit: false,
        children: Array.from({ length: nCol }, (_, c) => mkCell(r[c], c, false)),
      })),
    ],
  });
}

// ---------- blocks ----------
const children = [];
let numberedInstance = 0;
let inNumbered = false;
const spacer = () => new Paragraph({ spacing: { after: 60 }, children: [] });

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  const trimmed = line.trim();
  if (!/^\d+\.\s/.test(trimmed)) inNumbered = trimmed === "" ? inNumbered : false;

  if (trimmed === "") continue;

  // table
  if (trimmed.startsWith("|") && i + 1 < lines.length && /^\|?\s*:?-{3,}/.test(lines[i + 1].trim())) {
    const rows = [];
    while (i < lines.length && lines[i].trim().startsWith("|")) rows.push(lines[i++]);
    i--;
    children.push(buildTable(rows));
    children.push(spacer());
    continue;
  }

  // heading
  const h = /^(#{1,4})\s+(.*)$/.exec(trimmed);
  if (h) {
    const level = h[1].length;
    const map = { 1: HeadingLevel.TITLE, 2: HeadingLevel.HEADING_1, 3: HeadingLevel.HEADING_2, 4: HeadingLevel.HEADING_3 };
    children.push(new Paragraph({ heading: map[level], children: inlineRuns(h[2]) }));
    continue;
  }

  // horizontal rule
  if (/^-{3,}$/.test(trimmed)) {
    children.push(new Paragraph({
      spacing: { before: 120, after: 120 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF", space: 1 } },
      children: [],
    }));
    continue;
  }

  // blockquote
  if (trimmed.startsWith(">")) {
    children.push(new Paragraph({
      spacing: { before: 60, after: 120, line: 300 },
      indent: { left: 200 },
      shading: { fill: "EEF3F8", type: ShadingType.CLEAR, color: "auto" },
      border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 8 } },
      children: inlineRuns(trimmed.replace(/^>\s?/, ""), { size: 19 }),
    }));
    continue;
  }

  // bullet list
  const b = /^(\s*)[-*]\s+(.*)$/.exec(line);
  if (b) {
    const level = Math.min(Math.floor(b[1].length / 2), 1);
    children.push(new Paragraph({
      numbering: { reference: "bullets", level },
      spacing: { after: 60, line: 300 },
      children: inlineRuns(b[2]),
    }));
    continue;
  }

  // numbered list
  const n = /^(\s*)\d+\.\s+(.*)$/.exec(line);
  if (n) {
    if (!inNumbered) { numberedInstance++; inNumbered = true; }
    children.push(new Paragraph({
      numbering: { reference: "numbers", level: 0, instance: numberedInstance },
      spacing: { after: 60, line: 300 },
      children: inlineRuns(n[2]),
    }));
    continue;
  }

  // paragraph
  children.push(new Paragraph({ spacing: { after: 120, line: 300 }, children: inlineRuns(trimmed) }));
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: BODY_SIZE } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 34, bold: true, color: ACCENT },
        paragraph: { spacing: { after: 240 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 6 } } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 28, bold: true, color: ACCENT },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 24, bold: true, color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true, color: "404040" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "●", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 420, hanging: 280 } }, run: { size: 14, color: ACCENT } } },
        { level: 1, format: LevelFormat.BULLET, text: "○", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 840, hanging: 280 } }, run: { size: 14, color: ACCENT } } },
      ] },
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 420, hanging: 320 } } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1134, bottom: 1134, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES], size: 16, color: "808080" })],
    })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(outPath, buf); console.log("wrote", outPath); });
