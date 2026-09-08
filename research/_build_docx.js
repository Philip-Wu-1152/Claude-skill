const fs = require('fs');
const d = require('docx');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  TableOfContents, PageBreak, LevelFormat, convertInchesToTwip
} = d;

const CN = "宋体", CNH = "黑体", EN = "Times New Roman";
const F  = { ascii: EN, hAnsi: EN, eastAsia: CN, cs: EN };
const FH = { ascii: EN, hAnsi: EN, eastAsia: CNH, cs: EN };
const CONTENT_W = 9026; // A4 minus 1in margins, DXA

// ---------- inline parsing: **bold**, *italic*, `code` ----------
function inline(text, base = {}) {
  const runs = [];
  const re = /(\*\*[^*]+\*\*|\*[^*\n]+\*|`[^`]+`)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index), font: F, ...base }));
    const t = m[0];
    if (t.startsWith('**')) runs.push(new TextRun({ text: t.slice(2, -2), bold: true, font: F, ...base }));
    else if (t.startsWith('`')) runs.push(new TextRun({ text: t.slice(1, -1), font: { ascii: "Consolas", hAnsi: "Consolas", eastAsia: CN }, ...base }));
    else runs.push(new TextRun({ text: t.slice(1, -1), italics: true, font: F, ...base }));
    last = m.index + t.length;
  }
  if (last < text.length) runs.push(new TextRun({ text: text.slice(last), font: F, ...base }));
  return runs.length ? runs : [new TextRun({ text: "", font: F, ...base })];
}

function splitRow(line) {
  return line.replace(/^\s*\|/, '').replace(/\|\s*$/, '').split('|').map(s => s.trim());
}

function buildTable(rows) {
  const header = rows[0];
  const body = rows.slice(1);
  const n = header.length;
  const base = Math.floor(CONTENT_W / n);
  const widths = Array(n).fill(base);
  widths[n - 1] = CONTENT_W - base * (n - 1);

  const mkCell = (txt, isHead, w) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: isHead ? { type: ShadingType.CLEAR, fill: "E8E8E8" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      spacing: { before: 20, after: 20 },
      children: inline(txt, { size: 19, bold: isHead || undefined })
    })]
  });

  return new Table({
    columnWidths: widths,
    width: { size: CONTENT_W, type: WidthType.DXA },
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => mkCell(h, true, widths[i])) }),
      ...body.map(r => new TableRow({
        children: Array.from({ length: n }, (_, i) => mkCell(r[i] ?? "", false, widths[i]))
      }))
    ]
  });
}

function convert(md) {
  const lines = md.split('\n');
  const out = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    const t = line.trim();

    if (t === '') { i++; continue; }

    // horizontal rule
    if (/^-{3,}$/.test(t) || /^\*{3,}$/.test(t)) {
      out.push(new Paragraph({
        spacing: { before: 120, after: 120 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AAAAAA", space: 1 } },
        children: [new TextRun({ text: "", font: F })]
      }));
      i++; continue;
    }

    // table
    if (t.startsWith('|') && i + 1 < lines.length && /^\s*\|[\s:|-]+\|\s*$/.test(lines[i + 1])) {
      const rows = [splitRow(t)];
      i += 2;
      while (i < lines.length && lines[i].trim().startsWith('|')) { rows.push(splitRow(lines[i].trim())); i++; }
      out.push(buildTable(rows));
      out.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "", font: F, size: 12 })] }));
      continue;
    }

    // headings
    const h = t.match(/^(#{1,4})\s+(.*)$/);
    if (h) {
      const lvl = h[1].length;
      const sizes = { 1: 34, 2: 28, 3: 24, 4: 22 };
      const hl = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3, 4: HeadingLevel.HEADING_4 }[lvl];
      out.push(new Paragraph({
        heading: hl,
        spacing: { before: lvl <= 2 ? 320 : 240, after: 140 },
        keepNext: true,
        children: inline(h[2], { size: sizes[lvl], bold: true, font: FH, color: "1A1A1A" })
          .map(r => r)
      }));
      i++; continue;
    }

    // blockquote (consume consecutive)
    if (t.startsWith('>')) {
      const buf = [];
      while (i < lines.length && lines[i].trim().startsWith('>')) {
        buf.push(lines[i].trim().replace(/^>\s?/, ''));
        i++;
      }
      buf.filter(x => x !== '').forEach(x => {
        out.push(new Paragraph({
          spacing: { before: 60, after: 60, line: 300 },
          indent: { left: 340 },
          border: { left: { style: BorderStyle.SINGLE, size: 12, color: "999999", space: 8 } },
          shading: { type: ShadingType.CLEAR, fill: "F5F5F5" },
          children: inline(x, { size: 19, color: "333333" })
        }));
      });
      continue;
    }

    // bullet list
    const b = t.match(/^[-*]\s+(.*)$/);
    if (b) {
      out.push(new Paragraph({
        bullet: { level: 0 },
        spacing: { before: 40, after: 40, line: 320 },
        children: inline(b[1], { size: 21 })
      }));
      i++; continue;
    }

    // numbered list
    const nl = t.match(/^(\d+)\.\s+(.*)$/);
    if (nl) {
      out.push(new Paragraph({
        numbering: { reference: "num-list", level: 0 },
        spacing: { before: 40, after: 40, line: 320 },
        children: inline(nl[2], { size: 21 })
      }));
      i++; continue;
    }

    // plain paragraph
    out.push(new Paragraph({
      spacing: { before: 60, after: 60, line: 340 },
      indent: { firstLine: 420 },
      alignment: AlignmentType.JUSTIFIED,
      children: inline(t, { size: 21 })
    }));
    i++;
  }
  return out;
}

// ---------- assemble ----------
const base = '/home/user/Claude-skill/research/';
let main = fs.readFileSync(base + '研究背景与文献综述_v5.md', 'utf8');
let appx = fs.readFileSync(base + '附录A_海外疑似白马窑产品的文献线索.md', 'utf8');

// strip the top-level H1 from each (we build our own title page)
main = main.replace(/^#\s+.*\n/, '').replace(/^##\s+研究背景与研究综述（第五稿 v5）\n/m, '');
appx = appx.replace(/^#\s+.*\n/, '');

const titlePage = [
  new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "广东惠州白马窑仿龙泉青瓷工艺探析", bold: true, size: 40, font: FH })] }),
  new Paragraph({ spacing: { after: 700 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "及微（痕）量元素数据库构建", bold: true, size: 40, font: FH })] }),
  new Paragraph({ spacing: { after: 1400 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "研究背景与研究综述", size: 30, font: FH, color: "444444" })] }),
  ...[
    ["项　目", "广东省哲学社会科学规划 2025 年度常规项目（青年项目）"],
    ["学科分类", "历史学（考古学）"],
    ["负 责 人", "吴　博（中山大学）"],
    ["稿　次", "第五稿 v5（含附录 A）"],
  ].map(([k, v]) => new Paragraph({
    spacing: { after: 130 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: k + "：" + v, size: 22, font: F })]
  })),
  new Paragraph({ spacing: { before: 1800 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: (() => { const n = new Date(); return `${n.getFullYear()} 年 ${n.getMonth() + 1} 月 ${n.getDate()} 日`; })(), size: 21, font: F, color: "666666" })] }),
  new Paragraph({ children: [new PageBreak()] }),
];

const toc = [
  new Paragraph({ spacing: { after: 240 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "目　　录", bold: true, size: 30, font: FH })] }),
  new TableOfContents("目录", { hyperlink: true, headingStyleRange: "1-3" }),
  new Paragraph({ children: [new PageBreak()] }),
];

const appxTitle = [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: "附录 A　海内外「疑似白马窑」产品的文献线索", bold: true, size: 34, font: FH })] }),
];

const doc = new Document({
  creator: "吴博",
  title: "广东惠州白马窑仿龙泉青瓷工艺探析及微（痕）量元素数据库构建——研究背景与研究综述",
  description: "省社科规划课题研究背景与研究综述",
  styles: {
    default: {
      document: { run: { font: F, size: 21 } },
      heading1: { run: { font: FH, size: 34, bold: true, color: "1A1A1A" }, paragraph: { spacing: { before: 320, after: 160 } } },
      heading2: { run: { font: FH, size: 28, bold: true, color: "1A1A1A" }, paragraph: { spacing: { before: 300, after: 140 } } },
      heading3: { run: { font: FH, size: 24, bold: true, color: "1A1A1A" }, paragraph: { spacing: { before: 240, after: 120 } } },
      heading4: { run: { font: FH, size: 22, bold: true, color: "333333" }, paragraph: { spacing: { before: 200, after: 100 } } },
    },
  },
  numbering: {
    config: [{
      reference: "num-list",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 620, hanging: 340 } } } }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children: [...titlePage, ...toc, ...convert(main), ...appxTitle, ...convert(appx)],
  }],
});

Packer.toBuffer(doc).then(buf => {
  const out = base + '白马窑研究背景与文献综述.docx';
  fs.writeFileSync(out, buf);
  console.log('written:', out, (buf.length / 1024).toFixed(0) + ' KB');
});
