/**
 * Smart Billing — Git Auto Sync
 * pull --rebase → add -A → commit (ถ้ามีการเปลี่ยนแปลง) → push
 * ใช้: npm run sync   หรือ   node scripts/git-sync.js "ข้อความ commit"
 */
const { execSync } = require('child_process');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const msg = process.argv.slice(2).join(' ') ||
  ('Auto-sync ' + new Date().toLocaleString('th-TH'));

function run(cmd) {
  return execSync(cmd, { cwd: ROOT, stdio: ['ignore', 'pipe', 'pipe'] }).toString().trim();
}

try {
  run('git pull --rebase --autostash origin main');
  console.log('⬇️  Pulled latest from GitHub');

  run('git add -A');
  let hasChanges = true;
  try {
    run('git diff --cached --quiet');
    hasChanges = false;
  } catch (e) { hasChanges = true; }

  if (!hasChanges) {
    console.log('✅ ไม่มีการเปลี่ยนแปลง — ทุกอย่างอัปเดตแล้ว');
  } else {
    run('git commit -m "' + msg.replace(/"/g, '\\"') + '"');
    console.log('📝 Committed: ' + msg);
  }

  run('git push origin main');
  console.log('⬆️  Pushed to github.com/itman-ncc/invoice.git');
} catch (e) {
  console.error('❌ Git sync failed:', e.message);
  process.exit(1);
}
