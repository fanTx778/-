// 统一样式：字体、行距、缩进、页边距。
// 单位换算：1 pt = 20 twips；字号 size 为 half-point（小四 = 24）。
const {
  Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType,
  BorderStyle, PageBreak, VerticalAlign, LineRuleType,
} = require('docx');

const EN = 'Times New Roman';
const CN_SONG = '宋体';
const CN_HEI = '黑体';

const SZ = {
  xiaoer: 36, // 小二 18pt
  san: 32,    // 三号 16pt
  xiaosi: 24, // 小四 12pt
  wu: 21,     // 五号 10.5pt
  xiaowu: 18, // 小五 9pt
};

const TW = (pt) => Math.round(pt * 20);           // pt -> twips
const charTw = (size) => Math.round((size / 2) * 20); // 一个字符宽（twips）

// 把 [1] / [2,3] / [4-6] 识别为参考文献角标
const REF_RE = /(\[\d+(?:\s*[,\-–—]\s*\d+)*\])/;

function makeRuns(text, { size = SZ.xiaosi, bold = false, cn = CN_SONG, ref = true, italic = false } = {}) {
  const pieces = String(text).split(ref ? REF_RE : /$^/);
  const out = [];
  for (const p of pieces) {
    if (p === '') continue;
    const isRef = ref && REF_RE.test(p) && /^\[\d/.test(p);
    out.push(new TextRun({
      text: p,
      size,
      bold,
      italics: italic,
      superScript: isRef || undefined,
      font: { ascii: EN, hAnsi: EN, eastAsia: cn, cs: EN },
    }));
  }
  return out;
}

// 段落：默认正文样式（宋体小四、24磅行距、首行缩进2字符、两端对齐）
function P(text, o = {}) {
  const {
    size = SZ.xiaosi, bold = false, cn = CN_SONG, align = 'both',
    line = 480, before = 0, after = 0, indentChars = 2, ref = true, italic = false,
    leftIndent = 0, hangingChars = 0,
  } = o;
  const alignMap = {
    both: AlignmentType.JUSTIFIED,
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT,
  };
  let indent;
  if (hangingChars) {
    indent = { left: charTw(size) * (leftIndent + hangingChars), hanging: charTw(size) * hangingChars };
  } else if (indentChars) {
    indent = { firstLine: charTw(size) * indentChars, left: leftIndent || undefined };
  } else {
    indent = leftIndent ? { left: leftIndent } : undefined;
  }
  return new Paragraph({
    alignment: alignMap[align],
    spacing: { line, lineRule: LineRuleType.EXACT, before, after },
    indent,
    children: makeRuns(text, { size, bold, cn, ref, italic }),
  });
}

// 空段（行距同正文）
function Blank(line = 480) {
  return new Paragraph({ spacing: { line, lineRule: LineRuleType.EXACT }, children: [new TextRun({ text: '', size: SZ.xiaosi })] });
}

function Break() {
  return new Paragraph({ children: [new PageBreak()] });
}

// 表格单元格：默认宋体五号居中、单行
function Cell(text, o = {}) {
  const { size = SZ.wu, bold = false, align = 'center', width = 0, cn = CN_SONG, valign = 'center' } = o;
  const alignMap = { center: AlignmentType.CENTER, left: AlignmentType.LEFT };
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    verticalAlign: valign === 'center' ? VerticalAlign.CENTER : VerticalAlign.TOP,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: alignMap[align],
      spacing: { line: 360, lineRule: LineRuleType.EXACT },
      children: makeRuns(text, { size, bold, cn }),
    })],
  });
}

// 段前段后 0.5 行 = 240 twips（行距 24 磅时）
const HALF_LINE = 240;

module.exports = { SZ, TW, charTw, P, Blank, Break, Cell, Table, TableRow, WidthType, BorderStyle, EN, CN_SONG, CN_HEI, HALF_LINE, makeRuns };
