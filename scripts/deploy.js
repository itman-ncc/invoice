/**
 * Smart Billing — Auto Deploy Script
 * 1) clasp push (อัปโหลดโค้ดล่าสุด)
 * 2) ถ้ามี .deployment-id → clasp redeploy (URL /exec คงเดิม)
 *    ถ้ายังไม่มี      → clasp deploy (สร้าง deployment แรก + บันทึก id)
 */
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ID_FILE = path.join(__dirname, '..', '.deployment-id');

function run(cmd, capture) {
  console.log('> ' + cmd);
  const out = execSync(cmd, {
    cwd: path.join(__dirname, '..'),
    stdio: capture ? ['ignore', 'pipe', 'pipe'] : 'inherit'
  });
  const s = out ? out.toString().trim() : '';
  if (capture && s) console.log(s);
  return s;
}

try {
  // Step 1: push (-f จำเป็นบน clasp 3.x: ไม่งั้น manifest prompt ทำให้ skip push ใน non-interactive mode)
  run('clasp push -f');

  // Step 2: deploy
  const desc = 'Smart Billing ' + new Date().toLocaleString('th-TH');
  let deploymentId = fs.existsSync(ID_FILE) ? fs.readFileSync(ID_FILE, 'utf8').trim() : '';

  if (deploymentId) {
    console.log('\n♻️  Redeploy existing deployment: ' + deploymentId);
    run(`clasp redeploy ${deploymentId} -d "${desc}"`);
  } else {
    console.log('\n🚀 First deploy — creating new deployment...');
    const out = run('clasp deploy -d "' + desc + '"', true);
    // clasp output เช่น "- AKfycbxxxx @1." หรือมี deployment id ยาว
    let m = out.match(/-\s+(\S+)\s+@/);
    if (!m) m = out.match(/(AKfycb[\w-]+)/);
    if (!m) m = out.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/);
    deploymentId = m ? m[1] : '';
    if (deploymentId) {
      fs.writeFileSync(ID_FILE, deploymentId);
      console.log('💾 Saved deployment-id → .deployment-id');
    }
  }

  if (deploymentId) {
    console.log('\n========================================');
    console.log('🌐 Web App URL:');
    console.log('   https://script.google.com/macros/s/' + deploymentId + '/exec');
    console.log('========================================\n');
  } else {
    console.log('⚠️  ไม่พบ deployment id — ตรวจสอบด้วยคำสั่ง: npx clasp deployments');
  }

  // Step 3: auto-sync ขึ้น GitHub (git error ไม่ทำให้ deploy ล้ม)
  try {
    console.log('🔀 Auto-sync to GitHub...');
    run('node scripts/git-sync.js "Deploy: ' + desc + '"');
  } catch (e2) {
    console.error('⚠️  Git sync failed (แต่ deploy สำเร็จแล้ว): ' + e2.message);
  }
} catch (e) {
  console.error('❌ Deploy failed:', e.message);
  process.exit(1);
}
