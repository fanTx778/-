// 生成 out/开题答辩.pptx —— 开题报告 + 文献综述汇报（16:9，16 页）
// 定位：画面只放给老师看的关键词与图，详细讲法写进每页备注（speaker notes）。
// 内容口径与 docs/01_文献综述.md、docs/02_开题报告.md 一致，改内容请同步三处。
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'out', '开题答辩.pptx');
const ROUTE_IMG = path.join(ROOT, 'assets', '技术路线图.png');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.333 × 7.5 in
pptx.author = '范锋育';
pptx.company = '黄河科技学院';
pptx.title = '室内粉墙升降平台结构设计及运动分析 — 开题答辩';
pptx.subject = '毕业设计开题报告与文献综述汇报';

const FONT = '微软雅黑';
const NAVY = '15315B';
const BLUE = '2F5F9E';
const ACCENT = 'C9662B';
const INK = '333C4A';
const MUTED = '7A8794';
const LINE = 'DCE4EF';
const CARD = 'F6F9FC';

const W = 13.333;
const MX = 0.62;          // 左右页边距
const CW = W - MX * 2;    // 内容宽 12.093

/* ---------- 基础元件 ---------- */
function txt(slide, s, o) {
  slide.addText(s, Object.assign({ fontFace: FONT, color: INK, fontSize: 12, margin: 0 }, o));
}

function rect(slide, x, y, w, h, color, extra = {}) {
  slide.addShape(pptx.ShapeType.rect, Object.assign({ x, y, w, h, fill: { color }, line: { color, width: 0 } }, extra));
}

function round(slide, x, y, w, h, color, extra = {}) {
  slide.addShape(pptx.ShapeType.roundRect, Object.assign({
    x, y, w, h, rectRadius: 0.05, fill: { color }, line: { color: LINE, width: 1 },
  }, extra));
}

function pageHead(slide, kicker, title) {
  rect(slide, MX, 0.62, 0.075, 0.78, ACCENT);
  txt(slide, kicker, { x: MX + 0.22, y: 0.6, w: 9, h: 0.28, fontSize: 10.5, color: MUTED, charSpacing: 1.2 });
  txt(slide, title, { x: MX + 0.22, y: 0.86, w: CW - 0.2, h: 0.55, fontSize: 25, bold: true, color: NAVY, fit: 'shrink' });
}

function pageFoot(slide, no) {
  slide.addShape(pptx.ShapeType.line, { x: MX, y: 7.06, w: CW, h: 0, line: { color: 'E5E9F0', width: 0.75 } });
  txt(slide, '室内粉墙升降平台结构设计及运动分析 · 开题答辩', { x: MX, y: 7.12, w: 8, h: 0.26, fontSize: 9, color: 'A0AAB6' });
  txt(slide, String(no).padStart(2, '0'), { x: W - MX - 0.7, y: 7.12, w: 0.7, h: 0.26, fontSize: 9, color: 'A0AAB6', align: 'right' });
}

// 卡片：标题 + 正文（正文只用短句，\n 手动分行）
function card(slide, o) {
  const { x, y, w, h, head, body, tag, color = BLUE, headSize = 16, bodySize = 13.5, fill = CARD } = o;
  round(slide, x, y, w, h, fill);
  if (tag) {
    rect(slide, x + 0.26, y + 0.26, 0.075, 0.2, color);
    txt(slide, tag, { x: x + 0.44, y: y + 0.22, w: w - 0.72, h: 0.2, fontSize: 9, color: color, bold: true, charSpacing: 0.8 });
    txt(slide, head, { x: x + 0.26, y: y + 0.5, w: w - 0.52, h: 0.32, fontSize: headSize, bold: true, color: NAVY });
    txt(slide, body, { x: x + 0.26, y: y + 0.9, w: w - 0.52, h: h - 1.06, fontSize: bodySize, color: INK, lineSpacingMultiple: 1.3, valign: 'middle' });
  } else {
    rect(slide, x + 0.26, y + 0.24, 0.075, 0.3, color);
    txt(slide, head, { x: x + 0.46, y: y + 0.2, w: w - 0.72, h: 0.34, fontSize: headSize, bold: true, color: NAVY });
    txt(slide, body, { x: x + 0.26, y: y + 0.66, w: w - 0.52, h: h - 0.86, fontSize: bodySize, color: INK, lineSpacingMultiple: 1.3, valign: 'middle' });
  }
}

// 表格：首行表头（深蓝底白字），其余白/浅底交替
function table(slide, rows, o) {
  const { x, y, w, colW, rowH = 0.46, fontSize = 12, headSize = 12 } = o;
  const data = rows.map((r, ri) => r.map((c, ci) => {
    const isHead = ri === 0;
    return {
      text: c,
      options: {
        fill: { color: isHead ? NAVY : (ri % 2 === 0 ? 'F6F9FC' : 'FFFFFF') },
        color: isHead ? 'FFFFFF' : (ci === 0 ? NAVY : INK),
        bold: isHead || ci === 0,
        fontSize: isHead ? headSize : fontSize,
        align: ci === 0 ? 'center' : 'left',
        valign: 'middle',
        margin: [0.04, 0.08, 0.04, 0.08],
      },
    };
  }));
  slide.addTable(data, {
    x, y, w, colW, rowH,
    fontFace: FONT,
    border: { type: 'solid', color: LINE, pt: 1 },
    autoPage: false,
  });
}

/* ================= 1 封面 ================= */
{
  const s = pptx.addSlide();
  s.background = { color: 'FFFFFF' };
  s.addShape(pptx.ShapeType.ellipse, { x: 9.4, y: -1.5, w: 5.6, h: 5.6, fill: { color: 'EEF3F9' }, line: { color: 'EEF3F9', width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 11.4, y: 4.6, w: 3.2, h: 3.2, fill: { color: 'F4F7FB' }, line: { color: 'F4F7FB', width: 0 } });
  rect(s, 0, 0, 0.22, 7.5, NAVY);

  txt(s, '黄河科技学院 · 2026 届本科毕业设计（论文）', { x: 0.95, y: 1.28, w: 10, h: 0.3, fontSize: 12, color: MUTED, charSpacing: 1 });
  txt(s, '室内粉墙升降平台结构设计及运动分析', { x: 0.95, y: 1.82, w: 11.6, h: 0.8, fontSize: 34, bold: true, color: NAVY });
  rect(s, 0.98, 2.76, 0.9, 0.06, ACCENT);
  txt(s, '开题报告与文献综述汇报', { x: 0.98, y: 2.98, w: 8, h: 0.4, fontSize: 19, color: BLUE });

  // 信息两列
  const info = [
    ['汇 报 人', '范锋育'],
    ['学    号', '2328050011'],
    ['专    业', '机械设计制造及其自动化'],
    ['指导教师', '高大伟'],
  ];
  info.forEach((it, i) => {
    const x = 0.98 + (i % 2) * 4.3;
    const y = 4.02 + Math.floor(i / 2) * 0.82;
    txt(s, it[0], { x, y, w: 1.3, h: 0.3, fontSize: 10.5, color: MUTED, charSpacing: 1 });
    txt(s, it[1], { x: x + 1.25, y: y - 0.04, w: 2.9, h: 0.36, fontSize: 14, bold: true, color: INK });
  });

  rect(s, 0.98, 5.86, 3.2, 0.02, LINE);
  txt(s, '汇报时间：2026 年 1 月', { x: 0.98, y: 6.0, w: 5, h: 0.3, fontSize: 11, color: MUTED });
  txt(s, '课题类型：AX（工程设计 + 真实课题）', { x: 8.2, y: 6.0, w: 4.3, h: 0.3, fontSize: 10.5, color: MUTED, align: 'right' });

  s.addNotes([
    '开场（约 20 秒）：各位老师好，我是机械设计制造及其自动化专业的范锋育，指导教师高大伟老师。',
    '我的题目是《室内粉墙升降平台结构设计及运动分析》，课题类型 AX，工程设计类真实课题。',
    '下面从选题背景、文献综述、设计方案、预期成果四部分汇报。',
  ]);
}

/* ================= 2 提纲 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, 'CONTENTS', '汇报提纲');
  const items = [
    ['01', '选题背景与意义', '抹灰作业的现状 · 抹灰机械化三个层次'],
    ['02', '文献综述与研究空白', '16 篇核心文献 · 四条线索 · 三处空白'],
    ['03', '设计方案与技术路线', '结构比选 · 主要参数 · 难点对策 · 运动分析'],
    ['04', '预期成果、进度与条件', '成果清单 · 14 周安排 · 可行性'],
  ];
  items.forEach((it, i) => {
    const y = 1.72 + i * 1.3;
    txt(s, it[0], { x: MX + 0.1, y, w: 1.1, h: 0.7, fontSize: 30, bold: true, color: ACCENT, fontFace: 'Arial' });
    txt(s, it[1], { x: MX + 1.25, y: y - 0.02, w: 6.0, h: 0.42, fontSize: 20, bold: true, color: NAVY });
    txt(s, it[2], { x: MX + 1.25, y: y + 0.44, w: 9.8, h: 0.34, fontSize: 12, color: MUTED });
    if (i < 3) s.addShape(pptx.ShapeType.line, { x: MX + 1.25, y: y + 1.0, w: CW - 1.25, h: 0, line: { color: 'EDF1F7', width: 1 } });
  });
  pageFoot(s, 2);
  s.addNotes(['四部分，按 8–10 分钟准备；第 9、10 页是提问高发页，讲慢一点。']);
}

/* ================= 3 选题背景 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '01 选题背景与意义', '室内抹灰：效率、质量、安全三重压力');
  const cw = (CW - 0.3 * 2) / 3;
  const cards = [
    ['效率低', '两人一整天\n只抹一面墙'],
    ['质量离散大', '同墙不同区段\n厚度差常超 5 mm'],
    ['安全与健康', '坠落、腰颈椎劳损\n砂浆碱性腐蚀'],
  ];
  cards.forEach((c, i) => card(s, {
    x: MX + i * (cw + 0.3), y: 1.68, w: cw, h: 2.0, head: c[0], body: c[1], bodySize: 15,
  }));

  round(s, MX, 4.0, CW, 0.95, 'F2F6FB');
  txt(s, '验收指标', { x: MX + 0.28, y: 4.24, w: 1.2, h: 0.3, fontSize: 11, bold: true, color: ACCENT });
  txt(s, '抹灰层厚 15–20 mm；平整度允许偏差 普通 4 mm / 高级 3 mm',
    { x: MX + 1.35, y: 4.0, w: CW - 1.6, h: 0.95, fontSize: 14, color: INK, valign: 'middle' });

  rect(s, MX, 5.3, CW, 1.15, NAVY);
  txt(s, '抹灰层的厚度均匀性，最终由承载平台的运动精度决定。',
    { x: MX + 0.4, y: 5.3, w: CW - 0.8, h: 1.15, fontSize: 19, bold: true, color: 'FFFFFF', valign: 'middle' });
  pageFoot(s, 3);

  s.addNotes([
    '现状三句话：效率低——一面普通住宅内墙，两人配合一整天；质量离散——平整度、厚度全凭手感，同墙不同区段厚度差常超过 5 mm；安全与健康——坠落风险、腰颈椎劳损、砂浆碱性腐蚀。',
    '验收指标是设计公差的来源：GB 50210 规定抹灰层 15–20 mm，普通抹灰平整度允许偏差 4 mm、高级 3 mm，配合 JGJ/T 220。平台的位姿偏差必须控制在这一量级。',
    '引出结论：抹灰质量最终取决于承载平台的运动精度，这是本课题的切入点。',
  ]);
}

/* ================= 4 机械化三层次 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '01 选题背景与意义', '抹灰机械化：三层演进，瓶颈在平台');
  const cw = (CW - 0.62 * 2) / 3;
  const blocks = [
    ['第一层 · 单机替代', '喷涂机、砂浆泵\n替代上料与甩浆', '找平、压实仍靠人'],
    ['第二层 · 工序集成（主流）', '喷浆 + 压实 + 抹平\n一台机器一次成型', '陆龙福、阮学云、王博等'],
    ['第三层 · 整机自动化（少数）', '定位—贴墙—升降—抹灰\n全过程自主', '瓶颈：承载平台'],
  ];
  blocks.forEach((b, i) => {
    const x = MX + i * (cw + 0.62);
    round(s, x, 1.62, cw, 2.0, i === 2 ? 'FFF6F0' : CARD, { line: { color: i === 2 ? 'F0D8C8' : LINE, width: 1 } });
    txt(s, b[0], { x: x + 0.26, y: 1.8, w: cw - 0.52, h: 0.32, fontSize: 14, bold: true, color: i === 2 ? ACCENT : NAVY });
    txt(s, b[1], { x: x + 0.26, y: 2.24, w: cw - 0.52, h: 0.8, fontSize: 13.5, color: INK, lineSpacingMultiple: 1.3, valign: 'top' });
    txt(s, b[2], { x: x + 0.26, y: 3.1, w: cw - 0.52, h: 0.34, fontSize: 11, color: MUTED });
    if (i < 2) txt(s, '›', { x: x + cw + 0.1, y: 2.3, w: 0.45, h: 0.6, fontSize: 32, color: 'B9C6D6', align: 'center' });
  });

  round(s, MX, 3.9, CW, 0.9, 'F2F6FB');
  txt(s, '判断', { x: MX + 0.28, y: 4.12, w: 1.1, h: 0.3, fontSize: 11, bold: true, color: ACCENT });
  txt(s, '抹灰头已成熟，制约整机落地的是平台的刚度、平稳性与贴墙精度。',
    { x: MX + 1.2, y: 3.9, w: CW - 1.45, h: 0.9, fontSize: 14, color: INK, valign: 'middle' });

  txt(s, '室内粉墙的三条刚性约束', { x: MX, y: 5.02, w: 6, h: 0.3, fontSize: 12.5, bold: true, color: NAVY });
  const cons = [
    ['受限空间', '层高 3.0–4.5 m\n过 0.8 m 门洞'],
    ['贴墙作业', '全行程平行度\n≤3 mm'],
    ['单侧偏载', '人 + 砂浆桶靠墙\n另有顶紧反力'],
  ];
  const cw2 = (CW - 0.3 * 2) / 3;
  cons.forEach((c, i) => card(s, {
    x: MX + i * (cw2 + 0.3), y: 5.38, w: cw2, h: 1.3, head: c[0], body: c[1], headSize: 13, bodySize: 11.5,
  }));
  pageFoot(s, 4);

  s.addNotes([
    '三层演进：单机替代（喷涂机、砂浆泵替代上料甩浆，找平压实仍靠人）；工序集成（喷浆、压实、抹平一体，机器沿墙行走一次成型，是目前主流，陆龙福、阮学云、孔令盼、王博都属这一层）；整机自动化（定位、贴墙、升降、抹灰、移位全过程自主，成果尚少）。',
    '关键判断：抹灰头的喷浆、压实、抹平均有成熟方案，制约整机落地的是室内受限空间下平台的刚度、运动平稳性与贴墙精度。',
    '三条刚性约束说明通用高空作业平台不能直接套用，引出后面的方案附加机构。',
  ]);
}

/* ================= 5 文献概况 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '02 文献综述与研究空白', '文献调研：16 篇核心文献');
  const nums = [
    ['16', '篇', '核心文献'],
    ['14', '篇', '期刊论文（87.5%）'],
    ['3', '篇', '英文文献'],
    ['69%', '', '2021 年以后'],
  ];
  const nw = 2.72;
  nums.forEach((n, i) => {
    const x = MX + (i % 2) * (nw + 0.24);
    const y = 1.62 + Math.floor(i / 2) * 1.92;
    round(s, x, y, nw, 1.74, CARD);
    txt(s, [
      { text: n[0], options: { fontSize: 36, bold: true, color: ACCENT, fontFace: 'Arial' } },
      { text: ' ' + n[1], options: { fontSize: 11, color: MUTED } },
    ], { x: x + 0.24, y: y + 0.2, w: nw - 0.48, h: 0.6 });
    txt(s, n[2], { x: x + 0.24, y: y + 0.92, w: nw - 0.48, h: 0.3, fontSize: 13, bold: true, color: NAVY });
  });

  const rx = MX + 2 * nw + 0.24 + 0.35;
  const rw = W - MX - rx;
  card(s, {
    tag: '检 索', x: rx, y: 1.62, w: rw, h: 2.4,
    head: '数据库与检索',
    body: '知网、万方、维普、国家科技期刊平台 + Crossref\n检索词：粉墙机 / 自动抹灰机 / 升降平台 / 剪叉机构\n先场景 → 后装备 → 再方法，逐层收敛',
    bodySize: 13, headSize: 15,
  });
  card(s, {
    tag: '依 据', x: rx, y: 4.22, w: rw, h: 2.1,
    head: '标准与工具书',
    body: 'GB 50210、JGJ/T 220、GB/T 25849—2010\nGB/T 700、GB/T 6892；《机械设计手册》等\nGB/T 25849 是稳定性校核的主要依据',
    bodySize: 13, headSize: 15,
  });
  pageFoot(s, 5);

  s.addNotes([
    '16 篇核心文献：期刊 14 篇（占 87.5%，含英文 3 篇）、学位论文 1 篇、专著 1 部（濮良贵《机械设计》）；2021 年以后占 69%，以近五年为主。',
    '检索顺序是“先场景、后装备、再方法”：先用 GB 50210 和 JGJ/T 220 确定 15–20 mm 层厚、4 mm 平整度这些工艺参数，再检索抹灰机与升降平台，最后补运动学、动力学、稳定性的方法文献。',
    '模板要求不少于 12 篇、期刊 80% 以上、含 2 篇英文，这几项都满足。',
  ]);
}

/* ================= 6 四条线索 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '02 文献综述与研究空白', '四条线索：方法齐全，对象多为通用平台');
  const cw = (CW - 0.3) / 2;
  const items = [
    ['2.1 抹灰机械整机', '抹灰头研究细致：喷浆、压实、抹平、喷嘴流场\n承载平台常以“立柱 + 升降机构”带过'],
    ['2.2 升降平台结构形式', '剪叉式研究最充分；桅柱式紧凑；导轨附着式贴墙最准\n国外 Gurskyi、Qiao、Ogbemhe'],
    ['2.3 运动学与动力学', '解析法（虚位移、几何闭环）+ 多体仿真（ADAMS）\n误差用微分法传递'],
    ['2.4 稳定性与调平', '稳定系数法、刚柔耦合（仿真与试验差 8.78%）、四点调平'],
  ];
  items.forEach((it, i) => {
    const x = MX + (i % 2) * (cw + 0.3);
    const y = 1.62 + Math.floor(i / 2) * 2.14;
    card(s, { x, y, w: cw, h: 1.98, head: it[0], body: it[1], headSize: 15.5, bodySize: 12.5 });
  });
  round(s, MX, 5.98, CW, 0.86, 'F2F6FB');
  txt(s, '评述', { x: MX + 0.28, y: 6.18, w: 1.0, h: 0.28, fontSize: 11, bold: true, color: ACCENT });
  txt(s, '方法体系完整，但研究对象普遍是通用高空作业平台。',
    { x: MX + 1.1, y: 5.98, w: CW - 1.35, h: 0.86, fontSize: 14, color: INK, valign: 'middle' });
  pageFoot(s, 6);

  s.addNotes([
    '2.1 整机：陆龙福（喷浆-压实-抹平一体）、阮学云（气动智能抹灰机）、孔令盼（喷嘴流场优化）、王博（整机结构优化）。共同倾向是抹灰头研究很细，承载平台一句话带过。',
    '2.2 结构形式：剪叉式研究最充分——毕涛用 RecurDyn、刘焕友用 ADAMS 加 TRIZ、付恒用虚位移原理、龙文龙做精度分析、陈晋市做皮带式；国外 Gurskyi 给解析式，Qiao 做设计-建模-控制，Ogbemhe 做多物理场。桅柱式紧凑、垂直度好；导轨附着式贴墙最准但要固定导轨。',
    '2.3 运动分析：解析法物理意义清晰，仿真法能输出铰点力时间历程；龙文龙把误差归为杆长、铰点位置、驱动误差三类，用微分法传递。',
    '2.4 稳定性：邬伟骏用稳定系数法，高明星做刚柔耦合、仿真与试验偏差 8.78%，赵腾做稳定性试验，张文浩设计调平机构。',
    '评述：方法体系完整，但对象普遍是通用平台，这正是空白所在。',
  ]);
}

/* ================= 7 研究空白 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '02 文献综述与研究空白', '三处空白，构成课题的研究空间');
  const gaps = [
    ['01', '复合工况下的稳定性', '缺“贴墙顶紧 + 单侧偏载”的倾覆分析'],
    ['02', '面向工艺的运动描述', '缺贴墙点轨迹与位姿偏差'],
    ['03', '受限空间的专用化平台', '通用平台不能直接套用'],
  ];
  gaps.forEach((g, i) => {
    const y = 1.66 + i * 1.28;
    round(s, MX, y, CW, 1.12, CARD);
    rect(s, MX + 0.22, y + 0.2, 0.72, 0.72, i === 0 ? NAVY : i === 1 ? BLUE : ACCENT);
    txt(s, g[0], { x: MX + 0.22, y: y + 0.2, w: 0.72, h: 0.72, fontSize: 15, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', fontFace: 'Arial' });
    txt(s, g[1], { x: MX + 1.15, y: y + 0.18, w: 4.2, h: 0.36, fontSize: 15.5, bold: true, color: NAVY });
    txt(s, g[2], { x: MX + 1.15, y: y + 0.6, w: CW - 1.45, h: 0.4, fontSize: 13.5, color: INK });
  });

  rect(s, MX, 5.66, CW, 1.05, NAVY);
  txt(s, '本课题定位', { x: MX + 0.32, y: 5.78, w: 1.6, h: 0.26, fontSize: 11, bold: true, color: 'F0B48A', charSpacing: 1 });
  txt(s, '剪叉为主体 + 贴墙导向 + 四点调平；解析与仿真互校，输出贴墙点位姿偏差。',
    { x: MX + 0.32, y: 6.06, w: CW - 0.64, h: 0.52, fontSize: 15, bold: true, color: 'FFFFFF' });
  pageFoot(s, 7);

  s.addNotes([
    '空白一：既有稳定性分析的前提都是“四周无约束、载荷对称”，缺少贴墙顶紧加单侧偏载的复合工况分析。',
    '空白二：运动分析多止步于举升角与驱动力、平台高度的内部关系，缺少平台边缘贴墙点的轨迹和它与理想竖直线的偏差——而这正是施工质量最关心的指标。',
    '空白三：通用平台按室外或工业工况设计，室内通过性、轻量化、贴墙定距缺少系统方案。',
    '所以本课题定位是：剪叉为主体，附加贴墙导向和四点调平，用解析加 ADAMS 互校，把贴墙点位姿偏差作为核心输出。',
  ]);
}

/* ================= 8 结构形式比选 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '03 设计方案与技术路线', '结构形式比选：为什么选剪叉式');
  table(s, [
    ['结构形式', '承载能力', '收缩比 / 整机高度', '贴墙精度', '室内通过性', '成熟度 / 成本'],
    ['剪叉式', '强', '大，收缩后低', '低，需附加导向', '好', '最成熟，成本低'],
    ['桅柱式', '中', '小，多级立柱', '较好', '好', '同步控制复杂'],
    ['导轨附着式', '中', '—（沿墙爬升）', '最高，导轨定距', '差，需固定导轨', '破坏免安装前提'],
  ], { x: MX, y: 1.7, w: CW, colW: [1.95, 1.65, 2.35, 2.6, 1.95, 1.6], rowH: 0.68, fontSize: 12 });

  txt(s, '比选维度：承载能力 / 收缩高度 / 贴墙精度 / 室内通过性 / 成本',
    { x: MX, y: 4.5, w: CW, h: 0.3, fontSize: 11.5, color: MUTED });

  round(s, MX, 4.95, CW, 1.8, CARD);
  txt(s, '选定：剪叉式 + 附加机构补短板', { x: MX + 0.28, y: 5.1, w: CW - 0.56, h: 0.34, fontSize: 16, bold: true, color: NAVY });
  const add = [
    ['保留', '承载强、收缩比大、成熟、成本低'],
    ['补足 1', '前缘贴墙导向轮 → 悬臂变两点支承'],
    ['补足 2', '丝杠调距 → 间隙 20–50 mm 主动可调'],
    ['补足 3', '四角螺旋支腿 → 消除初始倾斜'],
  ];
  add.forEach((a, i) => {
    const y = 5.55 + i * 0.3;
    txt(s, a[0], { x: MX + 0.3, y, w: 0.9, h: 0.28, fontSize: 10.5, bold: true, color: i === 0 ? BLUE : ACCENT });
    txt(s, a[1], { x: MX + 1.2, y, w: CW - 1.5, h: 0.28, fontSize: 12.5, color: INK });
  });
  pageFoot(s, 8);

  s.addNotes([
    '按五个维度比较：承载能力、收缩比、贴墙精度、室内通过性、成本。',
    '剪叉式承载强、收缩比大、结构最成熟、成本最低，短板是高位横向刚度低、不能主动贴墙；桅柱式紧凑、垂直度好，但承载有限、多级同步控制复杂；导轨附着式贴墙精度最高，但要在墙面固定导轨，破坏免安装、可移动的前提。',
    '室内约束是过门洞、可移动、能贴墙。剪叉收缩后整机高不超过 1.6 m，通过性最好；它的短板可以用机构补，另外两种的短板是原理性的。',
    '所以保留剪叉主体的优点，补三样：导向轮、丝杠调距、四角调平。',
  ]);
}

/* ================= 9 设计目标与参数 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '03 设计方案与技术路线', '设计目标与主要参数');
  table(s, [
    ['类别', '指标', '要求'],
    ['适用工况', '层高 / 门洞 / 地面', '3.0–4.5 m；净宽 ≥0.8 m；平整度 ≤10 mm/2 m'],
    ['承载', '额定载荷', '200 kg（1 人 + 砂浆桶 + 工具）'],
    ['几何', '台面 / 整机宽 / 举升 / 收缩', '1200×600 mm；≤780 mm；2.8 m；≤1.6 m'],
    ['运动', '升降速度 / 贴墙间隙', '0.08–0.15 m/s；20–50 mm 可调'],
    ['精度', '平行度 / 调平', '全行程 ≤3 mm；调平后 ≤0.5°'],
    ['材料', '主体 / 整机质量', '6061-T6 铝合金型材；≤200 kg'],
    ['安全', '稳定系数 / 装置', 'K ≥1.5；机械防坠 + 限位；平衡阀 + 液压锁'],
  ], { x: MX, y: 1.62, w: CW, colW: [1.5, 3.2, 7.4], rowH: 0.56, fontSize: 11.5 });

  txt(s, '2.8 m 由层高反推：台面 2.8 m + 人站立作业高度可覆盖 4.5 m 顶部，且不顶楼板；780 mm 保证过 800 mm 门洞。',
    { x: MX, y: 6.24, w: CW, h: 0.5, fontSize: 11, color: MUTED, lineSpacingMultiple: 1.2 });
  pageFoot(s, 9);

  s.addNotes([
    '参数来源逐个说：层高 3.0–4.5 m 决定台面 2.8 m——台面再高，人站上去会顶到楼板；最初按 4.0 m 设计，后来发现冲突才降到 2.8 m，剪叉也从 3 层减到 2 层。',
    '载荷 200 kg = 人 80 + 砂浆桶 60 + 工具 20，留约 25% 余量。',
    '整机宽 ≤780 mm 是为了过 800 mm 门洞，两侧各留 10 mm；收缩高 ≤1.6 m 便于搬运。',
    '精度：贴墙间隙 20–50 mm 可调，全行程平行度 ≤3 mm，调平后水平度 ≤0.5°，稳定系数 K ≥1.5。材料 6061-T6，整机不超过 200 kg。',
  ]);
}

/* ================= 10 难点与对策 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '03 设计方案与技术路线', '关键难点与对策');
  table(s, [
    ['难点', '成因', '对策'],
    ['高位横向刚度低', '铰接间隙逐级放大', '前缘导向轮顶墙，两点支承'],
    ['无法主动贴墙', '剪叉轨迹严格竖直', '丝杠调节间隙 20–50 mm'],
    ['地面初始倾斜', '误差转为楔形偏差', '四角螺旋支腿 + 水平仪'],
    ['偏载倾覆风险', '人员、料具靠墙侧', '最不利组合校核，K ≥1.5'],
    ['室内通过性受限', '整机进深偏大', '宽 ≤780 mm，收缩 ≤1.6 m'],
  ], { x: MX, y: 1.62, w: CW, colW: [2.5, 4.2, 5.4], rowH: 0.71, fontSize: 12.5 });

  round(s, MX, 6.0, CW, 0.72, 'FFF6F0', { line: { color: 'F0D8C8', width: 1 } });
  txt(s, '四项对策都不改变剪叉主体，只做“附加约束 + 附加自由度”。',
    { x: MX + 0.32, y: 6.0, w: CW - 0.6, h: 0.72, fontSize: 13, bold: true, color: ACCENT, valign: 'middle' });
  pageFoot(s, 10);

  s.addNotes([
    '难点一：高位横向刚度低，原因是剪叉臂铰接间隙经多层串联逐级放大；对策是前缘两侧贴墙导向轮顶住墙面，把悬臂变成两点支承。',
    '难点二：剪叉升降轨迹严格竖直，不能主动贴墙；对策是把导向轮做成丝杠调节式，间隙 20–50 mm 主动可调。',
    '难点三：地面有 1–3 度初始倾斜，调平误差会变成楔形偏差；对策是四角螺旋支腿加水平仪或倾角传感器。',
    '难点四：偏载倾覆，人员和砂浆桶靠墙侧，还有顶紧反力；对策是按满载加单侧偏载加贴墙顶紧的最不利组合校核，稳定系数不低于 1.5。',
    '难点五：室内通过性，整机进深控制在 780 mm 以内，收缩高度 1.6 m。',
  ]);
}

/* ================= 11 技术路线 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '03 设计方案与技术路线', '技术路线：五步推进');
  // 图 1800×1660，高 4.7 in → 宽 5.10 in
  s.addImage({ path: ROUTE_IMG, x: 0.8, y: 1.72, w: 5.1, h: 4.7 });
  txt(s, '图 1  技术路线图', { x: 0.8, y: 6.5, w: 5.1, h: 0.26, fontSize: 9.5, color: MUTED, align: 'center' });

  const steps = [
    ['①', '方案比选', '剪叉 / 桅柱 / 导轨附着，五维比较'],
    ['②', '机构补充', '贴墙导向、丝杠调距、四角调平'],
    ['③', '结构设计与校核', '2 层剪叉、铰点、液压缸；三关键件校核'],
    ['④', '运动分析', '解析建 θ–s–h + 间隙误差预算；ADAMS 铰点力'],
    ['⑤', '图纸与说明书', 'A0 装配图、A3 零件图、分析报告'],
  ];
  steps.forEach((st, i) => {
    const y = 1.66 + i * 1.0;
    txt(s, st[0], { x: 6.35, y: y + 0.02, w: 0.42, h: 0.4, fontSize: 16, bold: true, color: ACCENT });
    txt(s, st[1], { x: 6.82, y, w: 5.9, h: 0.32, fontSize: 15, bold: true, color: NAVY });
    txt(s, st[2], { x: 6.82, y: y + 0.34, w: 5.9, h: 0.36, fontSize: 12, color: MUTED });
  });
  pageFoot(s, 11);

  s.addNotes([
    '五步：方案比选定结构；机构补充补短板；结构设计与强度校核定尺寸；运动分析做解析、误差预算和铰点力提取；最后出图纸和说明书。',
    '第三步的关键是剪叉层数、铰点布置和液压缸安装位置——缸位按驱动力峰值和驱动行程两个指标比较确定。',
    '第四步的重点是误差预算：先算理想运动关系，再把杆长、铰点、间隙、变形折算成贴墙间隙的变化；ADAMS 主要用来取三种工况的铰点力。下一页展开。',
  ]);
}

/* ================= 12 运动分析 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '03 设计方案与技术路线', '运动分析：误差预算 + 铰点力');
  const cw = (CW - 0.3) / 2;
  card(s, {
    tag: '解 析 法', x: MX, y: 1.62, w: cw, h: 2.45,
    head: '几何关系 + 误差传递',
    body: '举升角 θ → 驱动行程 s、平台高度 h\n误差源 → 贴墙间隙变化范围',
    bodySize: 14, headSize: 15.5,
  });
  card(s, {
    tag: '仿 真 法', x: MX + cw + 0.3, y: 1.62, w: cw, h: 2.45,
    head: 'SolidWorks + ADAMS',
    body: '运动学校验 + 准静态铰点力\n铰接间隙用衬套等效',
    bodySize: 14, headSize: 15.5,
  });

  txt(s, '三种计算工况', { x: MX, y: 4.34, w: 4, h: 0.3, fontSize: 12.5, bold: true, color: NAVY });
  const cases = [
    ['空载', '自重与摩擦'],
    ['额定载荷 200 kg', '人、料、工具居中'],
    ['单侧偏载', '靠墙集中 + 顶紧反力'],
  ];
  const cw3 = (CW - 0.24 * 2) / 3;
  cases.forEach((c, i) => {
    const x = MX + i * (cw3 + 0.24);
    round(s, x, 4.72, cw3, 0.9, i === 2 ? 'FFF6F0' : CARD, { line: { color: i === 2 ? 'F0D8C8' : LINE, width: 1 } });
    txt(s, c[0], { x: x + 0.24, y: 4.82, w: cw3 - 0.48, h: 0.3, fontSize: 13.5, bold: true, color: i === 2 ? ACCENT : NAVY });
    txt(s, c[1], { x: x + 0.24, y: 5.14, w: cw3 - 0.48, h: 0.34, fontSize: 11, color: MUTED });
  });

  txt(s, '用途：铰点力 → 强度校核与驱动选型；误差预算 → 检验平行度 ≤3 mm 是否可达。\n接触模型仅作探索，不作结论依据。',
    { x: MX, y: 5.86, w: CW, h: 0.66, fontSize: 12, color: INK, lineSpacingMultiple: 1.3 });
  pageFoot(s, 12);

  s.addNotes([
    '解析法：以举升角为参量，用几何关系和闭环矢量方程推出驱动行程与平台高度。注意理想剪叉是竖直升降，贴墙间隙在理想情况下恒定，所以偏差要靠误差模型算——把杆长、铰点位置、铰接间隙、构件变形折算成间隙变化量，按最不利情况叠加，看能否落在 3 mm 以内。',
    '仿真法：SolidWorks 建模后导入 ADAMS，做运动学校验，再用准静态分析提三种工况的铰点力；铰接间隙用衬套等效，比接触模型稳定。',
    '两者是分工关系：解析法保证规律和误差预算，仿真给铰点力和载荷偏置的趋势。',
    '三种工况：空载、额定 200 kg 居中、单侧偏载加顶紧，用来给强度校核、驱动选型和间隙判断做依据。',
    '边界讲清楚：平行度用误差预算、水平度用支腿行程反算、稳定系数用计算，不做样机实测；导向轮接触只作探索。',
  ]);
}

/* ================= 13 预期成果 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '04 预期成果、进度与条件', '预期成果');
  const outs = [
    ['01', '文献综述', '≥3000 字（已完成）'],
    ['02', '开题报告', '≥2000 字'],
    ['03', '装配图', 'A0 一张，三视图 + 技术特性表'],
    ['04', '零件图', 'A3 三至四张'],
    ['05', '三维模型', '整机装配 + 仿真动画'],
    ['06', '分析报告', '解析 + ADAMS + 三工况'],
    ['07', '说明书', '≥8000 字'],
  ];
  outs.forEach((o, i) => {
    const col = i < 4 ? 0 : 1;
    const row = i < 4 ? i : i - 4;
    const x = MX + col * (CW / 2 + 0.15);
    const y = 1.72 + row * 1.14;
    const w = CW / 2 - 0.15;
    txt(s, o[0], { x, y: y + 0.06, w: 0.55, h: 0.34, fontSize: 15, bold: true, color: ACCENT, fontFace: 'Arial' });
    txt(s, o[1], { x: x + 0.62, y, w: 1.6, h: 0.34, fontSize: 15, bold: true, color: NAVY });
    txt(s, o[2], { x: x + 2.3, y: y + 0.04, w: w - 2.3, h: 0.4, fontSize: 12, color: INK });
    s.addShape(pptx.ShapeType.line, { x, y: y + 0.9, w, h: 0, line: { color: 'EDF1F7', width: 1 } });
  });
  round(s, MX, 6.36, CW, 0.5, 'F2F6FB');
  txt(s, '图纸 5 张 · 模型 1 套 · 分析报告 1 份 · 说明书 1 份',
    { x: MX + 0.26, y: 6.36, w: CW - 0.5, h: 0.5, fontSize: 11.5, color: BLUE, valign: 'middle' });
  pageFoot(s, 13);
  s.addNotes(['成果七项，对应任务书要求：文献综述、开题报告已完成；后面是 A0 装配图 1 张、A3 零件图 3–4 张、三维模型、分析报告、8000 字以上的说明书。']);
}

/* ================= 14 进度 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '04 预期成果、进度与条件', '进度安排：14 周（2025-11-15 — 2026-05-15）');
  const phases = [
    ['1–2 周', '搜集资料', '文献综述已完成'],
    ['3–4 周', '开题答辩', '方案与参数确定'],
    ['5–6 周', '结构设计', '剪叉、导向、调平'],
    ['7–8 周', '选型与校核', '强度校核、图纸'],
    ['9–10 周', '三维建模', '运动学推导'],
    ['11–12 周', 'ADAMS 仿真', '三工况对比优化'],
    ['13–14 周', '整理成果', '说明书、答辩'],
  ];
  const bw = (CW - 0.13 * 6) / 7;
  s.addShape(pptx.ShapeType.line, { x: MX + 0.2, y: 2.06, w: CW - 0.4, h: 0, line: { color: 'C9D5E4', width: 2 } });
  phases.forEach((p, i) => {
    const x = MX + i * (bw + 0.13);
    const now = i === 1;
    s.addShape(pptx.ShapeType.ellipse, {
      x: x + bw / 2 - 0.09, y: 1.97, w: 0.18, h: 0.18,
      fill: { color: now ? ACCENT : 'FFFFFF' }, line: { color: now ? ACCENT : '9FB3CB', width: 2 },
    });
    txt(s, p[0], { x, y: 1.52, w: bw, h: 0.3, fontSize: 11.5, bold: true, color: now ? ACCENT : NAVY, align: 'center' });
    round(s, x, 2.34, bw, 1.9, now ? 'FFF6F0' : CARD, { line: { color: now ? 'F0D8C8' : LINE, width: 1 } });
    txt(s, p[1], { x: x + 0.12, y: 2.48, w: bw - 0.24, h: 0.62, fontSize: 12, bold: true, color: NAVY, lineSpacingMultiple: 1.2 });
    txt(s, p[2], { x: x + 0.12, y: 3.14, w: bw - 0.24, h: 0.7, fontSize: 10, color: MUTED, lineSpacingMultiple: 1.2 });
  });

  txt(s, '当前：开题阶段（第 3–4 周）',
    { x: MX, y: 4.56, w: CW, h: 0.34, fontSize: 13, bold: true, color: ACCENT });

  const goals = [
    ['2025.12', '开题答辩通过，方案定稿'],
    ['2026.01—02', '结构设计、强度校核、图纸初稿'],
    ['2026.03—04', '三维建模、运动分析、ADAMS 仿真'],
    ['2026.05', '说明书定稿，5 月 16 日答辩'],
  ];
  round(s, MX, 5.0, CW, 1.7, CARD);
  goals.forEach((g, i) => {
    const y = 5.14 + i * 0.38;
    txt(s, g[0], { x: MX + 0.3, y, w: 1.35, h: 0.3, fontSize: 10, bold: true, color: BLUE });
    txt(s, g[1], { x: MX + 1.75, y, w: CW - 2.0, h: 0.3, fontSize: 10.5, color: INK });
  });
  pageFoot(s, 14);
  s.addNotes(['14 周节点：前 4 周已完成文献综述与开题；5–8 周结构设计与校核；9–12 周建模与仿真；13–14 周整理成果。计划 2026 年 5 月 16 日答辩。']);
}

/* ================= 15 条件 ================= */
{
  const s = pptx.addSlide();
  pageHead(s, '04 预期成果、进度与条件', '完成条件与可行性');
  const cw = (CW - 0.3) / 2;
  const items = [
    ['课程基础', '机械原理、机械设计、理论力学、材料力学等已修完'],
    ['工具条件', 'AutoCAD、SolidWorks、ADAMS、MATLAB'],
    ['资料条件', '知网、万方、维普 + 国标与《机械设计手册》'],
    ['指导与检测', '导师指导；卡尺、百分表、水平仪可检测加工件'],
  ];
  items.forEach((it, i) => {
    const x = MX + (i % 2) * (cw + 0.3);
    const y = 1.62 + Math.floor(i / 2) * 1.92;
    card(s, { x, y, w: cw, h: 1.7, head: it[0], body: it[1], headSize: 16, bodySize: 13 });
  });
  rect(s, MX, 5.52, CW, 0.9, NAVY);
  txt(s, '知识、工具、资料与指导条件均已具备，可按计划完成。',
    { x: MX + 0.32, y: 5.52, w: CW - 0.6, h: 0.9, fontSize: 16, bold: true, color: 'FFFFFF', valign: 'middle' });
  pageFoot(s, 15);
  s.addNotes(['课程上主干课已修完；工具上 AutoCAD、SolidWorks、ADAMS、MATLAB 都能用；资料上图书馆数据库齐全；指导上导师做机械结构与机构分析，实验室有卡尺、百分表、水平仪，可对加工件检测。']);
}

/* ================= 16 结束 ================= */
{
  const s = pptx.addSlide();
  s.background = { color: NAVY };
  s.addShape(pptx.ShapeType.ellipse, { x: 9.2, y: -1.8, w: 5.8, h: 5.8, fill: { color: '1D3E6E' }, line: { color: '1D3E6E', width: 0 } });
  rect(s, 0, 0, 0.22, 7.5, ACCENT);
  txt(s, 'THANKS FOR YOUR ATTENTION', { x: 1.0, y: 1.9, w: 9, h: 0.3, fontSize: 11, color: '9FB3CB', charSpacing: 2 });
  txt(s, '请各位老师批评指正', { x: 1.0, y: 2.4, w: 10.5, h: 0.9, fontSize: 40, bold: true, color: 'FFFFFF' });
  rect(s, 1.03, 3.46, 1.1, 0.06, ACCENT);
  txt(s, '室内粉墙升降平台结构设计及运动分析', { x: 1.0, y: 3.72, w: 10, h: 0.4, fontSize: 16, color: 'D6E2F0' });
  txt(s, '汇报人：范锋育（2328050011）　指导教师：高大伟　机械设计制造及其自动化',
    { x: 1.0, y: 4.34, w: 11, h: 0.34, fontSize: 11.5, color: '9FB3CB' });
  pageFoot(s, 16);
  s.addNotes(['结束语：以上是我的汇报，请各位老师批评指正。']);
}

pptx.writeFile({ fileName: OUT }).then(() => {
  console.log('已生成：' + path.relative(ROOT, OUT));
}).catch((e) => {
  console.error(e);
  process.exit(1);
});
