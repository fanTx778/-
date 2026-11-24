// 生成 out/任务书.docx
// 页面：左 3cm，上/下/右 2.5cm；正文宋体小四 24 磅行距首行缩进 2 字符
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LineRuleType } = require('docx');
const { SZ, P, Blank, EN, CN_SONG, CN_HEI, HALF_LINE, makeRuns } = require('./lib');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'docs', '04_任务书.md');
const OUT = path.join(ROOT, 'out', '任务书.docx');

const INFO = {
  collegeLabel: '应用技术',
  collegeSuffix: '学院',
  major: '机械设计制造及其自动化',
  gradeYear: '2023',
  gradeClass: '1',
  sno: '2328050011',
  student: '范锋育',
  teacher: '高大伟',
  title: '室内粉墙升降平台结构设计及运动分析',
  startDate: '2025 年 11 月 15 日',
  endDate: '2026 年 5 月 15 日',
  defence: '2026 年 5 月 16 日',
};

// 个人信息行：标签不加下划线，填写值加下划线（与任务书模板一致）
function infoLine(parts) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 480, lineRule: LineRuleType.EXACT },
    children: parts.map((p) => new TextRun({
      text: p.t,
      size: SZ.xiaosi,
      underline: p.u ? { type: 'single' } : undefined,
      font: { ascii: EN, hAnsi: EN, eastAsia: CN_SONG, cs: EN },
    })),
  });
}

function parse(md) {
  const sections = [];
  let cur = null;
  for (const raw of md.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('# ')) continue;
    if (line.startsWith('## ')) {
      cur = { title: line.slice(3).trim(), paras: [] };
      sections.push(cur);
      continue;
    }
    if (cur) cur.paras.push(line);
  }
  return sections;
}

async function main() {
  const sections = parse(fs.readFileSync(SRC, 'utf8'));
  const kids = [];

  // 标题
  kids.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 480, lineRule: LineRuleType.EXACT, before: 0, after: HALF_LINE },
    children: [new TextRun({ text: '黄河科技学院本科毕业设计任务书', bold: true, size: SZ.san, font: { ascii: EN, hAnsi: EN, eastAsia: CN_HEI } })],
  }));

  // 学院 / 专业 / 年级班级
  kids.push(infoLine([
    { t: '　　', u: true },
    { t: INFO.collegeLabel, u: true },
    { t: '　', u: true },
    { t: INFO.collegeSuffix },
    { t: '　　', u: true },
    { t: INFO.major, u: true },
    { t: '　', u: true },
    { t: '专业' },
    { t: `　${INFO.gradeYear}　`, u: true },
    { t: '级' },
    { t: `　${INFO.gradeClass}　`, u: true },
    { t: '班' },
  ]));
  kids.push(infoLine([
    { t: '学号' },
    { t: `　　　${INFO.sno}　　　`, u: true },
    { t: '学生' },
    { t: `　　　${INFO.student}　　　`, u: true },
    { t: '指导教师' },
    { t: `　　　${INFO.teacher}　　　`, u: true },
  ]));
  kids.push(Blank(480));

  // 题目
  kids.push(P('毕业设计题目', { size: SZ.xiaosi, cn: CN_HEI, align: 'left', before: HALF_LINE, after: 120, indentChars: 0 }));
  kids.push(P(INFO.title, { size: SZ.xiaosi, align: 'center', indentChars: 0 }));

  // 工作内容与基本要求
  kids.push(P('毕业设计工作内容与基本要求', { size: SZ.xiaosi, cn: CN_HEI, align: 'left', before: HALF_LINE, after: 0, indentChars: 0 }));

  for (const sec of sections) {
    if (sec.title !== '引言') {
      // 一级标题：宋体小四，左对齐，段前 0.5 行，段后 0
      kids.push(P(sec.title, { size: SZ.xiaosi, align: 'left', before: HALF_LINE, after: 0, indentChars: 0 }));
    }
    const isRef = sec.title.startsWith('三、');
    for (const t of sec.paras) {
      if (isRef) {
        kids.push(P(t, { size: SZ.wu, align: 'both', line: 400, indentChars: 0, hangingChars: 0, leftIndent: 0 }));
      } else {
        kids.push(P(t, { size: SZ.xiaosi, align: 'both', line: 480 }));
      }
    }
  }

  // 时间
  kids.push(Blank(480));
  kids.push(P(`毕业设计时间：　${INFO.startDate} 至　${INFO.endDate}`, { size: SZ.xiaosi, align: 'left', indentChars: 0 }));
  kids.push(P(`计划答辩时间：　${INFO.defence}`, { size: SZ.xiaosi, align: 'left', indentChars: 0 }));
  kids.push(Blank(480));
  kids.push(P('专业（教研室）审批意见：', { size: SZ.xiaosi, align: 'left', indentChars: 0 }));
  kids.push(Blank(480), Blank(480));
  kids.push(P('审批人（签字）：', { size: SZ.xiaosi, align: 'left', indentChars: 0 }));

  const doc = new Document({
    creator: 'bishe',
    title: '任务书',
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
          // 左 3cm = 1701 twips；上/下/右 2.5cm = 1417 twips
          margin: { top: 1417, bottom: 1417, left: 1701, right: 1417 },
          size: { width: 11906, height: 16838 },
        },
      },
      children: kids,
    }],
  });

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(OUT, buf);
  console.log('生成:', OUT, (buf.length / 1024).toFixed(1) + ' KB');
}

main().catch(e => { console.error(e); process.exit(1); });
