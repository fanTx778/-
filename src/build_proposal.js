// 生成 out/开题报告.docx —— 表格式，宋体五号 20 磅行距
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
  WidthType, VerticalAlign, LineRuleType, BorderStyle,
} = require('docx');
const { SZ, EN, CN_SONG, CN_HEI, makeRuns } = require('./lib');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'docs', '02_开题报告.md');
const OUT = path.join(ROOT, 'out', '开题报告.docx');

// ---- 待填项 ----
const INFO = {
  title: '室内粉墙升降平台结构设计及运动分析',
  source: '教师拟订',
  type: 'AX',
  teacher: '高大伟',
  student: '范锋育',
  major: '机械设计制造及其自动化',
  sno: '2328050011',
  signDate: '2026 年 1 月 5 日',
};

// 正文段落（五号宋体，20 磅行距）
function p(text, o = {}) {
  const { bold = false, align = 'both', line = 400, before = 0, after = 0, indentChars = 2, size = SZ.wu, cn = CN_SONG } = o;
  const alignMap = { both: AlignmentType.JUSTIFIED, left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT };
  return new Paragraph({
    alignment: alignMap[align],
    spacing: { line, lineRule: LineRuleType.EXACT, before, after },
    indent: indentChars ? { firstLine: Math.round((size / 2) * 20 * indentChars) } : undefined,
    children: makeRuns(text, { size, bold, cn }),
  });
}

function parse(md) {
  const lines = md.split(/\r?\n/);
  const blocks = [];
  for (const raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    if (line.startsWith('### ')) { blocks.push({ t: 'h2', s: line.slice(4).trim() }); continue; }
    if (line.startsWith('## ')) { blocks.push({ t: 'h1', s: line.slice(3).trim() }); continue; }
    if (line.startsWith('# ')) continue;
    blocks.push({ t: 'p', s: line });
  }
  return blocks;
}

function buildBody(blocks) {
  const out = [];
  for (const b of blocks) {
    if (b.t === 'h1') {
      // 宋体五号加粗，段前段后 0.5 行（20 磅行距下 = 200 twips）
      out.push(p(b.s, { bold: true, align: 'left', before: 200, after: 200, indentChars: 0 }));
    } else if (b.t === 'h2') {
      out.push(p(b.s, { bold: false, align: 'left', before: 100, after: 0, indentChars: 0 }));
    } else {
      out.push(p(b.s, { indentChars: 2 }));
    }
  }
  return out;
}

function headCell(text, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: { fill: 'FFFFFF' },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 320, lineRule: LineRuleType.EXACT },
      children: makeRuns(text, { size: SZ.wu, bold: false, cn: CN_SONG }),
    })],
  });
}

function valueCell(text, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: opts.align === 'left' ? AlignmentType.LEFT : AlignmentType.CENTER,
      spacing: { line: 320, lineRule: LineRuleType.EXACT },
      children: makeRuns(text, { size: SZ.wu, bold: false, cn: CN_SONG }),
    })],
  });
}

function wideCell(children, span = 6) {
  return new TableCell({
    columnSpan: span,
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children,
  });
}

async function main() {
  const blocks = parse(fs.readFileSync(SRC, 'utf8'));

  const W = [1400, 1500, 1400, 1400, 1400, 1526]; // 合计约 8626 twips
  const rows = [];

  // 课题名称
  rows.push(new TableRow({
    children: [headCell('课题名称', W[0]), wideCell([
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 320, lineRule: LineRuleType.EXACT },
        children: makeRuns(INFO.title, { size: SZ.wu, cn: CN_SONG }),
      }),
    ], 5)],
  }));
  // 课题来源 / 课题类型 / 指导教师
  rows.push(new TableRow({
    children: [
      headCell('课题来源', W[0]), valueCell(INFO.source, W[1]),
      headCell('课题类型', W[2]), valueCell(INFO.type, W[3]),
      headCell('指导教师', W[4]), valueCell(INFO.teacher, W[5]),
    ],
  }));
  // 学生 / 专业 / 学号
  rows.push(new TableRow({
    children: [
      headCell('学生姓名', W[0]), valueCell(INFO.student, W[1]),
      headCell('专    业', W[2]), valueCell(INFO.major, W[3]),
      headCell('学    号', W[4]), valueCell(INFO.sno, W[5]),
    ],
  }));
  // 正文
  rows.push(new TableRow({ children: [wideCell(buildBody(blocks))] }));
  // 签名
  rows.push(new TableRow({
    children: [wideCell([
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { line: 400, lineRule: LineRuleType.EXACT, before: 200, after: 200 },
        children: makeRuns('指导教师签名：                      ' + INFO.signDate + '      ', { size: SZ.wu, cn: CN_SONG }),
      }),
    ])],
  }));
  // 尾注
  const notes = [
    '课题来源：（1）教师拟订；（2）学生建议；（3）企业和社会征集；（4）科研单位提供',
    '课题类型：（1）A—工程设计（艺术设计）；B—技术开发；C—软件工程；D—理论研究；E—调研报告',
    '　　　　　（2）X—真实课题；Y—模拟课题；Z—虚拟课题',
    '要求（1）、（2）均要填，如 AY、BX 等。',
  ].map(t => p(t, { size: SZ.xiaowu, align: 'left', line: 320, indentChars: 0 }));
  rows.push(new TableRow({ children: [wideCell(notes)] }));

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: W,
    rows,
  });

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 400, lineRule: LineRuleType.EXACT, before: 0, after: 200 },
      children: [new TextRun({ text: '黄河科技学院毕业设计开题报告表', bold: true, size: SZ.xiaoer, font: { ascii: EN, hAnsi: EN, eastAsia: CN_HEI } })],
    }),
    table,
  ];

  const doc = new Document({
    creator: 'bishe',
    title: '开题报告',
    styles: {
      default: {
        document: {
          run: { font: { ascii: EN, hAnsi: EN, eastAsia: CN_SONG }, size: SZ.wu },
          paragraph: { spacing: { line: 400, lineRule: LineRuleType.EXACT } },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
          size: { width: 11906, height: 16838 },
        },
      },
      children,
    }],
  });

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT, buf);

  const bodyText = blocks.filter(b => b.t === 'p').map(b => b.s).join('');
  const cjk = (bodyText.match(/[\u4e00-\u9fa5]/g) || []).length;
  const en = (bodyText.match(/[A-Za-z]+/g) || []).length;
  console.log('生成:', OUT, (buf.length / 1024).toFixed(1) + ' KB');
  console.log('正文字数（Word 口径估算）:', cjk + en, '（汉字', cjk, '+ 英文词', en, '）');
}

main().catch(e => { console.error(e); process.exit(1); });
