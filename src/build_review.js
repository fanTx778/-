// 生成 out/文献综述.docx
// 依据模板：页边距上下 2.54cm、左右 3.17cm；正文小四宋体 24 磅行距首行缩进 2 字符；
// 一级标题黑体小四不加粗、段前段后 0.5 行；摘要「摘要：」加粗、100-200 字；参考文献五号。
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LineRuleType, ImageRun, BorderStyle, WidthType, HeightRule, VerticalAlign, TableCell } = require('docx');
const { SZ, P, Blank, Break, Cell, Table, TableRow, EN, CN_SONG, CN_HEI, HALF_LINE, makeRuns } = require('./lib');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'docs', '01_文献综述.md');
const OUT = path.join(ROOT, 'out', '文献综述.docx');
const COVER_LOGO = path.join(ROOT, 'assets', '校徽.jpeg');

// ---- 待填项：改这里后重新生成即可 ----
const INFO = {
  college: '应用技术学院',
  major: '机械设计制造及其自动化',
  student: '范锋育',
  teacher: '高大伟',
  date: '2025年12月25日',
};

function parse(md) {
  const lines = md.split(/\r?\n/);
  let title = '';
  const abstract = [];
  const keywords = [];
  const body = [];      // {type:'h1'|'h2'|'p', text}
  const refs = [];
  let mode = null;

  for (const raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    if (line.startsWith('### ')) { mode = 'body'; body.push({ type: 'h2', text: line.slice(4).trim() }); continue; }
    if (line.startsWith('## ')) {
      const h = line.slice(3).trim();
      if (h === '摘要') { mode = 'abstract'; continue; }
      if (h === '关键词') { mode = 'keywords'; continue; }
      if (h === '参考文献') { mode = 'refs'; continue; }
      mode = 'body';
      body.push({ type: 'h1', text: h });
      continue;
    }
    if (line.startsWith('# ')) { title = line.slice(2).trim(); continue; }
    if (mode === 'abstract') abstract.push(line);
    else if (mode === 'keywords') keywords.push(line);
    else if (mode === 'refs') refs.push(line);
    else if (mode === 'body') body.push({ type: 'p', text: line });
  }
  return { title, abstract: abstract.join(''), keywords: keywords.join(''), body, refs };
}

// —— 封面：版式逐项对齐模板（../文献综述模板25.10.20.docx）——
// 模板封面：4 个空行 + 校徽 + 2 个空行 + 「毕业设计」小二黑体 + 空行 + 「文献综述」26 磅黑体
//          + 空行 + 题目 18 磅黑体 + 2 个空行 + 信息表（仅下框线、左缩进 1025）+ 空行 + 日期 15 磅
function coverBlank({ align, size = 28 } = {}) {
  return new Paragraph({
    alignment: align,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
    // 空段落的高度由“段落标记”字号决定（1.5 倍行距 × 字号），故同时设置段落标记 run
    run: { size, font: { ascii: EN, hAnsi: EN, eastAsia: CN_HEI, cs: EN } },
    children: [new TextRun({ text: '', size })],
  });
}

function coverLine(text, { size, align = AlignmentType.CENTER, line = 360, rule = LineRuleType.AUTO } = {}) {
  return new Paragraph({
    alignment: align,
    spacing: { line, lineRule: rule, before: 0, after: 0 },
    children: [new TextRun({ text, size, font: { ascii: EN, hAnsi: EN, eastAsia: CN_HEI, cs: EN } })],
  });
}

// 单元格：模板只在“填写值”一列画填空下划线，标签列不画线
const LINE = { style: BorderStyle.SINGLE, size: 4, color: 'auto' };
const NO_LINE = { style: BorderStyle.NONE, size: 0, color: 'auto' };

function coverCell(text, { width, align, valign = VerticalAlign.TOP, borders }) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: valign,
    margins: { top: 0, bottom: 0, left: 108, right: 108 },
    ...(borders ? { borders } : {}),
    children: [new Paragraph({
      alignment: align,
      spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
      children: [new TextRun({ text, size: 30, font: { ascii: EN, hAnsi: EN, eastAsia: CN_HEI, cs: EN } })],
    })],
  });
}

function cover(title) {
  const kids = [];
  // 空段字号按模板设置：行距为 1.5 倍，字号决定空行高度
  kids.push(coverBlank({ align: AlignmentType.RIGHT, size: 21 }));
  kids.push(coverBlank({ align: AlignmentType.RIGHT, size: 21 }));
  kids.push(coverBlank({ align: AlignmentType.RIGHT, size: 21 }));
  kids.push(coverBlank({ align: AlignmentType.CENTER, size: 21 }));

  // 校徽：模板中宽 4497705 EMU × 高 1214120 EMU（354.1 × 95.6 pt）
  kids.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
    run: { size: 21, font: { ascii: EN, hAnsi: EN, eastAsia: CN_HEI, cs: EN } },
    children: [new ImageRun({
      type: 'jpg',
      data: fs.readFileSync(COVER_LOGO),
      transformation: { width: 472.2, height: 127.5 },
    })],
  }));
  kids.push(coverBlank({ align: AlignmentType.CENTER, size: 21 }));
  kids.push(coverBlank({ align: AlignmentType.CENTER, size: 28 }));

  kids.push(coverLine('毕业设计', { size: 72 }));
  kids.push(coverBlank({ size: 28, align: AlignmentType.JUSTIFIED }));
  kids.push(coverLine('文献综述', { size: 52 }));
  kids.push(coverBlank({ size: 28, align: AlignmentType.JUSTIFIED }));

  kids.push(coverLine(title, { size: 36, line: 480, rule: LineRuleType.EXACT }));
  kids.push(coverBlank({ size: 28, align: AlignmentType.JUSTIFIED }));
  kids.push(coverBlank({ size: 21 }));

  const rows = [
    ['院（系）名称', INFO.college],
    ['专业名称', INFO.major],
    ['学生姓名', INFO.student],
    ['指导教师', INFO.teacher],
  ].map(([k, v], i) => {
    // 模板：第 1 行值列仅下线；第 2~4 行值列上下线；第 4 行标签列屏蔽外框底线
    const valueBorders = i === 0 ? { bottom: LINE } : { top: LINE, bottom: LINE };
    const labelBorders = i === 3 ? { bottom: { style: BorderStyle.NIL, size: 0, color: 'auto' } } : undefined;
    return new TableRow({
      height: { value: [616, 616, 601, 631][i], rule: HeightRule.ATLEAST },
      children: [
        coverCell(k, { width: 2448, align: AlignmentType.DISTRIBUTE, borders: labelBorders }),
        coverCell(v, { width: 4371, align: AlignmentType.CENTER, valign: VerticalAlign.CENTER, borders: valueBorders }),
      ],
    });
  });

  kids.push(new Table({
    width: { size: 6819, type: WidthType.DXA },
    columnWidths: [2448, 4371],
    indent: { size: 1025, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'auto' },
      right: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'auto' },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    },
    rows,
  }));

  kids.push(coverBlank({ size: 28 }));
  kids.push(coverLine(INFO.date, { size: 30 }));
  kids.push(coverBlank({ size: 32 }));
  kids.push(Break());
  return kids;
}

function bodyBlock(parsed) {
  const out = [];
  // 正文题目
  out.push(P(parsed.title, {
    size: SZ.san, cn: CN_HEI, bold: false, align: 'center', line: 480,
    before: HALF_LINE, after: 0, indentChars: 0,
  }));

  // 摘要：整段首行缩进 2 字符，「摘要：」加粗
  const abstractKids = [
    new TextRun({ text: '摘要：', bold: true, size: SZ.xiaosi, font: { ascii: EN, hAnsi: EN, eastAsia: CN_SONG } }),
    ...makeRuns(parsed.abstract, { size: SZ.xiaosi, cn: CN_SONG, ref: false }),
  ];
  out.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 480, lineRule: LineRuleType.EXACT, before: 0, after: 0 },
    indent: { firstLine: 480 },
    children: abstractKids,
  }));

  // 关键词
  out.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 480, lineRule: LineRuleType.EXACT, before: 0, after: 0 },
    indent: { firstLine: 480 },
    children: [
      new TextRun({ text: '关键词：', bold: true, size: SZ.xiaosi, font: { ascii: EN, hAnsi: EN, eastAsia: CN_SONG } }),
      ...makeRuns(parsed.keywords, { size: SZ.xiaosi, cn: CN_SONG, ref: false }),
    ],
  }));

  // 正文
  for (const b of parsed.body) {
    if (b.type === 'h1') {
      out.push(P(b.text, {
        size: SZ.xiaosi, cn: CN_HEI, bold: false, align: 'left',
        line: 480, before: HALF_LINE, after: HALF_LINE, indentChars: 0,
      }));
    } else if (b.type === 'h2') {
      out.push(P(b.text, {
        size: SZ.xiaosi, cn: CN_SONG, bold: true, align: 'left',
        line: 480, before: HALF_LINE, after: 0, indentChars: 0,
      }));
    } else {
      out.push(P(b.text, { size: SZ.xiaosi, cn: CN_SONG, align: 'both', line: 480 }));
    }
  }

  // 参考文献
  out.push(P('参考文献', {
    size: SZ.xiaosi, cn: CN_HEI, bold: false, align: 'left',
    line: 480, before: HALF_LINE, after: HALF_LINE, indentChars: 0,
  }));
  for (const r of parsed.refs) {
    out.push(P(r, {
      size: SZ.wu, cn: CN_SONG, align: 'both', line: 400,
      indentChars: 0, hangingChars: 0, leftIndent: 0,
    }));
  }
  return out;
}

async function main() {
  const parsed = parse(fs.readFileSync(SRC, 'utf8'));
  const children = [...cover(parsed.title), ...bodyBlock(parsed)];

  const doc = new Document({
    creator: 'bishe',
    title: '文献综述',
    styles: {
      default: {
        document: {
          run: { font: { ascii: EN, hAnsi: EN, eastAsia: CN_SONG }, size: SZ.xiaosi },
          paragraph: { spacing: { line: 480, lineRule: LineRuleType.EXACT } },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          // 上下 2.54cm = 1440 twips；左右 3.17cm = 1798 twips
          margin: { top: 1440, bottom: 1440, left: 1798, right: 1798 },
          size: { width: 11906, height: 16838 },
        },
      },
      children,
    }],
  });

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT, buf);
  console.log('生成:', OUT, (buf.length / 1024).toFixed(1) + ' KB');
  console.log('正文字数（Word 口径估算）:', wordCount(parsed));
}

function wordCount(p) {
  const t = p.body.map(b => b.text).join('');
  const cjk = (t.match(/[\u4e00-\u9fa5]/g) || []).length;
  const en = (t.match(/[A-Za-z]+/g) || []).length;
  const num = (t.match(/[0-9]+(?:\.[0-9]+)?/g) || []).length;
  return cjk + en + num;
}

main().catch(e => { console.error(e); process.exit(1); });
