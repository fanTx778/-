// 生成 out/课题申报表.docx —— 表格式，宋体五号 20 磅行距
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
  WidthType, VerticalAlign, LineRuleType,
} = require('docx');
const { SZ, EN, CN_SONG, CN_HEI, makeRuns } = require('./lib');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'docs', '05_课题申报表.md');
const OUT = path.join(ROOT, 'out', '课题申报表.docx');

const INFO = {
  title: '室内粉墙升降平台结构设计及运动分析',
  source: '教师拟订',
  type: 'AX',
  teacher: '高大伟',
  post: '【讲师】',
  unit: '黄河科技学院',
  practice: '否',
  place: '—',
};

function p(text, o = {}) {
  const { bold = false, align = 'both', line = 400, before = 0, after = 0, indentChars = 2, size = SZ.wu, cn = CN_SONG } = o;
  const am = { both: AlignmentType.JUSTIFIED, left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT };
  return new Paragraph({
    alignment: am[align],
    spacing: { line, lineRule: LineRuleType.EXACT, before, after },
    indent: indentChars ? { firstLine: Math.round((size / 2) * 20 * indentChars) } : undefined,
    children: makeRuns(text, { size, bold, cn }),
  });
}

function parse(md) {
  const secs = [];
  let cur = null;
  for (const raw of md.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('# ')) continue;
    if (line.startsWith('## ')) { cur = { title: line.slice(3).trim(), paras: [] }; secs.push(cur); continue; }
    if (cur) cur.paras.push(line);
  }
  return secs;
}

function headCell(text, width, span = 1) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    columnSpan: span > 1 ? span : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 320, lineRule: LineRuleType.EXACT },
      children: makeRuns(text, { size: SZ.wu, cn: CN_SONG }),
    })],
  });
}

function valCell(text, width, span = 1, align = 'center') {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    columnSpan: span > 1 ? span : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: align === 'left' ? AlignmentType.LEFT : AlignmentType.CENTER,
      spacing: { line: 320, lineRule: LineRuleType.EXACT },
      children: makeRuns(text, { size: SZ.wu, cn: CN_SONG }),
    })],
  });
}

function blockCell(children, span = 6) {
  return new TableCell({
    columnSpan: span,
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children,
  });
}

async function main() {
  const secs = parse(fs.readFileSync(SRC, 'utf8'));
  const W = [1500, 1400, 1400, 1400, 1400, 1926];

  const body = [];
  for (const s of secs) {
    body.push(p(s.title, { bold: true, align: 'left', before: 180, after: 100, indentChars: 0 }));
    for (const t of s.paras) body.push(p(t));
  }

  const rows = [
    // 课题名称
    new TableRow({ children: [headCell('课题名称', W[0]), valCell(INFO.title, 0, 5)] }),
    // 课题来源 / 课题类型
    new TableRow({ children: [headCell('课题来源', W[0]), valCell(INFO.source, W[1]), headCell('课题类型', W[2]), valCell(INFO.type, 0, 3)] }),
    // 指导教师 / 技术职务 / 工作单位
    new TableRow({ children: [headCell('指导教师', W[0]), valCell(INFO.teacher, W[1]), headCell('技术职务', W[2]), valCell(INFO.post, W[3]), headCell('工作单位', W[4]), valCell(INFO.unit, W[5])] }),
    // 是否需要实习 / 建议实习地点
    new TableRow({ children: [headCell('该课题是否需要实习', 1900), valCell(INFO.practice, W[1]), headCell('建议实习地点', W[2]), valCell(INFO.place, 0, 3)] }),
    // 正文
    new TableRow({ children: [blockCell(body)] }),
    // 审批意见
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 3,
          verticalAlign: VerticalAlign.TOP,
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [
            p('专业（教研室）意见：', { indentChars: 0 }),
            p('', { indentChars: 0 }),
            p('', { indentChars: 0 }),
            p('负责人签名：　　　　　　　2025 年 11 月 5 日', { align: 'right', indentChars: 0 }),
          ],
        }),
        new TableCell({
          columnSpan: 3,
          verticalAlign: VerticalAlign.TOP,
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [
            p('院（系）意见：', { indentChars: 0 }),
            p('', { indentChars: 0 }),
            p('', { indentChars: 0 }),
            p('负责人签名：　　　　　　2025 年 11 月 15 日', { align: 'right', indentChars: 0 }),
          ],
        }),
      ],
    }),
    // 尾注
    new TableRow({
      children: [blockCell([
        p('课题来源：（1）教师拟订；（2）学生建议；（3）企业和社会征集；（4）科研单位提供', { size: SZ.xiaowu, align: 'left', line: 320, indentChars: 0 }),
        p('课题类型：（1）A—工程设计（艺术设计）；B—技术开发；C—软件工程；D—理论研究；E—调研报告', { size: SZ.xiaowu, align: 'left', line: 320, indentChars: 0 }),
        p('　　　　　（2）X—真实课题；Y—模拟课题；Z—虚拟课题', { size: SZ.xiaowu, align: 'left', line: 320, indentChars: 0 }),
        p('要求（1）、（2）均要填，如 AX、BX 等。', { size: SZ.xiaowu, align: 'left', line: 320, indentChars: 0 }),
      ])],
    }),
  ];

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 400, lineRule: LineRuleType.EXACT, before: 0, after: 240 },
      children: [new TextRun({ text: '黄河科技学院毕业设计课题申报表', bold: true, size: SZ.san, font: { ascii: EN, hAnsi: EN, eastAsia: CN_HEI } })],
    }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: W, rows }),
  ];

  const doc = new Document({
    creator: 'bishe',
    title: '课题申报表',
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
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }, size: { width: 11906, height: 16838 } },
      },
      children,
    }],
  });

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT, buf);
  console.log('生成:', OUT, (buf.length / 1024).toFixed(1) + ' KB');

  const t = secs.flatMap(s => s.paras).join('');
  const cjk = (t.match(/[\u4e00-\u9fa5]/g) || []).length;
  const en = (t.match(/[A-Za-z]+/g) || []).length;
  console.log('正文字数（Word 口径估算）:', cjk + en);
}

main().catch(e => { console.error(e); process.exit(1); });
