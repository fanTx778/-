// 依次生成全部 docx
const { execFileSync } = require('child_process');
const path = require('path');

const scripts = ['build_review.js', 'build_proposal.js', 'build_task.js', 'build_declare.js', 'build_qa.js', 'build_ppt.js'];
for (const s of scripts) {
  const p = path.join(__dirname, s);
  console.log('\n>>> ' + s);
  execFileSync(process.execPath, [p], { stdio: 'inherit' });
}
console.log('\n全部生成完成');
