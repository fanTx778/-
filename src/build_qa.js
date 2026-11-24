// 生成 out/答辩问答摘要.docx —— 由 docs/06_答辩问答摘要.md 转换
// 版式：A4；页边距上下 2.54 cm、左右 3.17 cm；正文小四宋体 24 磅行距首行缩进 2 字符
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
  WidthType, VerticalAlign, LineRuleType, BorderStyle, Header, Footer, PageNumber,
} = require('docx');
const { SZ, EN, CN_SONG, CN_HEI, charTw } = require('./lib');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'docs', '06_答辩问答摘要.md');
const OUT = path.join(ROOT, 'out', '答辩问答摘要.docx');

/* ---------- 行内格式：**粗体**、`等宽` ---------- */
function runs(text, { size = SZ.xiaosi, bold = false, cn = CN_SONG, color } = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*|`[^`]+`)/g).filter((p) => p !== '');
  const out = [];
  for (const p of parts) {
    if (/^\*\*[^*]+\*\*$/.test(p)) {
      out.push(new TextRun({
        text: p.slice(2, -2), size, bold: true,
        font: { ascii: EN, hAnsi: EN, eastAsia: cn, cs: EN }, color,
      }));
    } else if (/^`[^`]+`$/.test(p)) {
      out.push(new TextRun({
        text: p.slice(1, -1), size: size - 2, bold,
        font: { ascii: 'Consolas', hAnsi: 'Consolas', eastAsia: CN_SONG, cs: 'Consolas' },
        color: '9C3A1E',
      }));
    } else {
      out.push(new TextRun({ text: p, size, bold, font: { ascii: EN, hAnsi: EN, eastAsia: cn, cs: EN }, color }));
    }
  }
  return out;
}

function para(text, o = {}) {
  const {
    size = SZ.xiaosi, bold = false, cn = CN_SONG, align = 'both', line = 480,
    before = 0, after = 0, indentChars = 2, color, leftChars = 0,
  } = o;
  const alignMap = {
    both: AlignmentType.JUSTIFIED, left: AlignmentType.LEFT,
    center: AlignmentType.CENTER, right: AlignmentType.RIGHT,
  };
  let indent;
  if (indentChars === 0 && !leftChars) indent = undefined;
  else indent = { firstLine: charTw(size) * indentChars, left: leftChars ? charTw(size) * leftChars : undefined };
  return new Paragraph({
    alignment: alignMap[align],
    spacing: { line, lineRule: LineRuleType.EXACT, before, after },
    indent,
    children: runs(text, { size, bold, cn, color }),
  });
}

// 列表项：悬挂缩进，行首符号
function bullet(text, { size = SZ.xiaosi, marker = '·' } = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 480, lineRule: LineRuleType.EXACT },
    indent: { left: charTw(size) * 2, hanging: charTw(size) },
    children: runs(`${marker} ${text}`, { size }),
  });
}

/* ---------- 三线表 ---------- */
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
function cell(text, { width, bold = false, align = 'left', size = SZ.xiaowu } = {}) {
  const isHead = bold;
  return new TableCell({
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    borders: {
      top: NONE, left: NONE, right: NONE,
      bottom: isHead
        ? { style: BorderStyle.SINGLE, size: 6, color: '000000' }
        : NONE,
    },
    children: [new Paragraph({
      alignment: align === 'center' ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { line: 300, lineRule: LineRuleType.EXACT },
      children: runs(text, { size, bold }),
    })],
  });
}

function table(rows, cols) {
  const widths = cols === 2 ? [32, 68] : Array.from({ length: cols }, () => Math.round(100 / cols));
  const trs = rows.map((r, ri) => new TableRow({
    tableHeader: ri === 0,
    cantSplit: true,
    children: r.map((c, ci) => cell(c, { width: widths[ci], bold: ri === 0, align: ri === 0 ? 'center' : (ci === 0 ? 'center' : 'left') })),
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 12, color: '000000' },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: '000000' },
      left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE,
    },
    rows: trs,
  });
}

/* ---------- 解析 Markdown ---------- */
function parse(md) {
  const lines = md.split(/\r?\n/);
  const blocks = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const t = line.trim();
    if (!t) { i++; continue; }
    if (t.startsWith('| ')) {
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        const row = lines[i].trim();
        if (!/^\|[\s:|-]+\|$/.test(row)) {
          const cells = row.replace(/^\|/, '').replace(/\|$/, '').split('|').map((x) => x.trim());
          rows.push(cells);
        }
        i++;
      }
      if (rows.length) blocks.push({ t: 'table', rows });
      continue;
    }
    if (t.startsWith('### ')) { blocks.push({ t: 'h3', s: t.slice(4) }); i++; continue; }
    if (t.startsWith('## ')) { blocks.push({ t: 'h2', s: t.slice(3) }); i++; continue; }
    if (t.startsWith('# ')) { blocks.push({ t: 'h1', s: t.slice(2) }); i++; continue; }
    if (t.startsWith('> ')) { blocks.push({ t: 'quote', s: t.slice(2) }); i++; continue; }
    if (t.startsWith('- ')) { blocks.push({ t: 'li', s: t.slice(2) }); i++; continue; }
    blocks.push({ t: 'p', s: t });
    i++;
  }
  return blocks;
}

function build(blocks) {
  const out = [];
  for (const b of blocks) {
    if (b.t === 'h1') {
      out.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 480, lineRule: LineRuleType.EXACT, before: 0, after: 240 },
        children: runs(b.s, { size: 32, bold: false, cn: CN_HEI }),
      }));
    } else if (b.t === 'h2') {
      out.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 480, lineRule: LineRuleType.EXACT, before: 360, after: 120 },
        keepNext: true, keepLines: true,
        children: runs(b.s, { size: SZ.xiaosi, bold: false, cn: CN_HEI }),
      }));
    } else if (b.t === 'h3') {
      out.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 480, lineRule: LineRuleType.EXACT, before: 240, after: 60 },
        keepNext: true, keepLines: true,
        children: runs(b.s, { size: SZ.xiaosi, bold: true }),
      }));
    } else if (b.t === 'quote') {
      out.push(para(b.s, { size: SZ.xiaowu, color: '595959', indentChars: 0, leftChars: 2, after: 120 }));
    } else if (b.t === 'li') {
      out.push(bullet(b.s));
    } else if (b.t === 'table') {
      out.push(table(b.rows, b.rows[0].length));
      out.push(new Paragraph({ spacing: { line: 120, lineRule: LineRuleType.EXACT }, children: [new TextRun({ text: '', size: 12 })] }));
    } else {
      out.push(para(b.s));
    }
  }
  return out;
}

const md = fs.readFileSync(SRC, 'utf8');
const children = build(parse(md));

const doc = new Document({
  creator: '范锋育',
  title: '开题答辩问答摘要',
  styles: { default: { document: { run: { font: { ascii: EN, hAnsi: EN, eastAsia: CN_SONG, cs: EN }, size: SZ.xiaosi } } } },
  sections: [{
    properties: {
      page: {
        margin: { top: 1440, bottom: 1440, left: 1797, right: 1797 },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, font: { ascii: EN, hAnsi: EN, eastAsia: CN_SONG } })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log('已生成：' + path.relative(ROOT, OUT));
});
