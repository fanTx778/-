// 校验生成的 docx：页边距、字体、行距、首行缩进、上标角标、字数、参考文献统计
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

function readZip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('not a zip');
  const n = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const files = {};
  for (let k = 0; k < n; k++) {
    const method = buf.readUInt16LE(off + 10);
    const csize = buf.readUInt32LE(off + 20);
    const nlen = buf.readUInt16LE(off + 28);
    const elen = buf.readUInt16LE(off + 30);
    const clen = buf.readUInt16LE(off + 32);
    const lho = buf.readUInt32LE(off + 42);
    const name = buf.toString('utf8', off + 46, off + 46 + nlen);
    files[name] = { method, csize, lho };
    off += 46 + nlen + elen + clen;
  }
  return (nm) => {
    const f = files[nm];
    if (!f) return null;
    const l = f.lho;
    const nl = buf.readUInt16LE(l + 26);
    const el = buf.readUInt16LE(l + 28);
    const st = l + 30 + nl + el;
    const raw = buf.subarray(st, st + f.csize);
    return f.method === 0 ? raw : zlib.inflateRawSync(raw);
  };
}

function textOf(xml) {
  return xml.replace(/<w:p[ >]/g, '\n<w:p ').replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

function count(xml, re) {
  return (xml.match(re) || []).length;
}

function checkDocx(file, opt) {
  const get = readZip(fs.readFileSync(file));
  const xml = get('word/document.xml').toString('utf8');
  const txt = textOf(xml);
  const cjk = (txt.match(/[\u4e00-\u9fa5]/g) || []).length;
  const en = (txt.match(/[A-Za-z]+/g) || []).length;

  const pgMar = (xml.match(/<w:pgMar[^>]*\/>/) || [''])[0];
  const margin = {};
  for (const k of ['top', 'bottom', 'left', 'right']) {
    const m = pgMar.match(new RegExp('w:' + k + '="(\\d+)"'));
    margin[k] = m ? +m[1] : null;
  }
  const got = {
    '宋体字符数': count(xml, /w:eastAsia="宋体"/g),
    '黑体字符数': count(xml, /w:eastAsia="黑体"/g),
    'Times New Roman': count(xml, /w:ascii="Times New Roman"/g),
    [`行距${opt.line}`]: count(xml, new RegExp('w:line="' + opt.line + '" w:lineRule="exact"', 'g')),
    '首行缩进2字符段': count(xml, new RegExp('w:firstLine="' + (opt.size * 20) + '"', 'g')),
    ...(opt.needRef === false ? {} : { '上标角标': count(xml, /w:vertAlign w:val="superscript"/g) }),
    ...(opt.needTable === false ? {} : { '表格数': count(xml, /<w:tbl>/g) }),
  };

  const pgOk = Math.abs(margin.top - opt.margin.top) <= 2
    && Math.abs(margin.bottom - opt.margin.bottom) <= 2
    && Math.abs(margin.left - opt.margin.left) <= 2
    && Math.abs(margin.right - opt.margin.right) <= 2;

  console.log('\n=== ' + path.basename(file) + ' ===');
  console.log('文件大小', (fs.statSync(file).size / 1024).toFixed(1), 'KB');
  console.log('页边距 (twips)', JSON.stringify(margin), pgOk ? '✓ 符合' : '✗ 不符，期望 ' + JSON.stringify(opt.margin));
  console.log('汉字', cjk, '英文词', en, '合计约', cjk + en, '字  | 模板要求 ≥' + opt.minWords, cjk + en >= opt.minWords ? '✓' : '✗');
  for (const [k, v] of Object.entries(got)) console.log(' ', k, v, v > 0 ? '✓' : '✗');
  return { xml, txt, cjk, en, pgOk, got };
}

const ROOT = path.join(__dirname, '..');
let fail = 0;

// 文献综述
const reviewPath = path.join(ROOT, 'out', '文献综述.docx');
if (fs.existsSync(reviewPath)) {
  const r = checkDocx(reviewPath, {
    line: 480, size: 24, minWords: 3000,
    margin: { top: 1440, bottom: 1440, left: 1798, right: 1798 },
  });
  // 参考文献统计
  const refs = (r.txt.split('参考文献')[1] || '').split('\n').map(s => s.trim()).filter(s => /^\[\d+\]/.test(s));
  const en = refs.filter(s => /\[J\]\.\s*[A-Z]/.test(s)).length;
  const journal = refs.filter(s => /\[J\]/.test(s)).length;
  const degree = refs.filter(s => /\[D\]/.test(s)).length;
  const book = refs.filter(s => /\[M\]/.test(s)).length;
  console.log(`参考文献 ${refs.length} 篇：期刊 ${journal}（${(journal / refs.length * 100).toFixed(0)}%）/ 学位 ${degree} / 专著 ${book} / 英文 ${en}`);
  if (refs.length < 12) { console.log('✗ 参考文献不足 12 篇'); fail++; }
  if (journal / refs.length < 0.8) { console.log('✗ 期刊占比不足 80%'); fail++; }
  if (en < 2) { console.log('✗ 英文文献不足 2 篇'); fail++; }
  if (book > 5) { console.log('✗ 专著超过 5 篇'); fail++; }
} else {
  console.log('缺少 out/文献综述.docx');
  fail++;
}

// 开题报告
const propPath = path.join(ROOT, 'out', '开题报告.docx');
if (fs.existsSync(propPath)) {
  checkDocx(propPath, {
    line: 400, size: 21, minWords: 2000,
    needRef: false,
    margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
  });
}

// 任务书（左 3cm，上/下/右 2.5cm）
const taskPath = path.join(ROOT, 'out', '任务书.docx');
if (fs.existsSync(taskPath)) {
  checkDocx(taskPath, {
    line: 480, size: 24, minWords: 800, needRef: false, needTable: false,
    margin: { top: 1417, bottom: 1417, left: 1701, right: 1417 },
  });
} else {
  console.log('缺少 out/任务书.docx');
}

// 课题申报表
const decPath = path.join(ROOT, 'out', '课题申报表.docx');
if (fs.existsSync(decPath)) {
  checkDocx(decPath, {
    line: 400, size: 21, minWords: 1000, needRef: false,
    margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
  });
} else {
  console.log('缺少 out/课题申报表.docx');
}

console.log(fail ? '\n有 ' + fail + ' 项未通过' : '\n全部检查通过');
process.exit(fail ? 1 : 0);