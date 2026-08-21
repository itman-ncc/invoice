/*************************************************************
 * SMART BILLING SYSTEM — Google Apps Script Backend
 * Sheet ID : 1Nf7ud48rwtGbEmOecxXhh3YyraC_kteEXFh4Ec-8NP8
 * ชีต: Users, Customers, Products, Settings, Documents,
 *      DocumentItems, Counters, AuditLogs, Sessions
 *************************************************************/

const SHEET_ID = '1Nf7ud48rwtGbEmOecxXhh3YyraC_kteEXFh4Ec-8NP8';
const TZ = 'Asia/Bangkok';
const SESSION_HOURS = 6;
const ITEMS_PER_PAGE = 20;

/* โฟลเดอร์ Google Drive เก็บไฟล์ PDF (override ได้ที่ชีต Settings → key: pdfFolderId) */
const DEFAULT_PDF_FOLDER_ID = '1VVel-aMXeY7As0OjjS1mIrRinY63k2o_';

/* ============================================================
 * SECTION 1: WEB APP ENTRY
 * ============================================================ */
function doGet() {
  ensureSetup_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Smart Billing System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* สร้างฐานข้อมูลอัตโนมัติครั้งแรกที่เปิด Web App + migration + ข้อมูลตัวอย่าง */
function ensureSetup_() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const users = ss.getSheetByName('Users');
    if (!users || users.getLastRow() < 2) {
      setupSystem();
    }
    migrateSettings_();
    migrateSchema_();
    autoImportSamples_();
  } catch (e) {
    // ถ้ายัง authorize ไม่ได้ ให้ผู้ใช้รัน setupSystem() จาก Editor เอง
  }
}

/* นำเข้าข้อมูลตัวอย่างอัตโนมัติครั้งแรก (เก็บ flag: sampleImported ในชีต Settings) */
function autoImportSamples_() {
  try {
    if (String(getSettings_().sampleImported) === 'yes') return;
    importSampleData(true);
  } catch (e) {}
}

/* เพิ่ม key ใหม่ลงชีต Settings อัตโนมัติถ้ายังไม่มี */
function migrateSettings_() {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
    if (!sh) return;
    const last = sh.getLastRow();
    let hasFolderId = false;
    if (last > 1) {
      const keys = sh.getRange(2, 1, last - 1, 1).getValues();
      hasFolderId = keys.some(function(r) { return String(r[0]) === 'pdfFolderId'; });
    }
    if (!hasFolderId) sh.appendRow(['pdfFolderId', DEFAULT_PDF_FOLDER_ID]);
  } catch (e) {}
}

/* เพิ่มคอลัมน์ใหม่ให้ชีตเดิมโดยไม่ลบข้อมูล:
 * Users → 'ชื่อ-สกุล' | Documents → 'จัดทำโดย (ชื่อ-สกุล)' + 'แก้ไขโดย (ชื่อ-สกุล)' */
function migrateSchema_() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const ush = ss.getSheetByName('Users');
    if (ush && String(ush.getRange(1, Math.max(ush.getLastColumn(), 1)).getValue()) !== 'ชื่อ-สกุล') {
      ush.getRange(1, 6).setValue('ชื่อ-สกุล');
      ush.setColumnWidth(6, 200);
      const last = ush.getLastRow();
      if (last > 1) {
        const rows = ush.getRange(2, 1, last - 1, 2).getValues();
        /* เติมค่าเริ่มต้น = ชื่อผู้ใช้ (ให้ Admin แก้เป็นชื่อจริงภายหลังได้) */
        ush.getRange(2, 6, last - 1, 1).setValues(rows.map(function(r) { return [String(r[1])]; }));
      }
    }

    const dsh = ss.getSheetByName('Documents');
    if (dsh && String(dsh.getRange(1, Math.max(dsh.getLastColumn(), 1)).getValue()) !== 'แก้ไขโดย (ชื่อ-สกุล)') {
      dsh.getRange(1, 20).setValue('จัดทำโดย (ชื่อ-สกุล)');
      dsh.getRange(1, 21).setValue('แก้ไขโดย (ชื่อ-สกุล)');
      dsh.setColumnWidth(20, 150);
      dsh.setColumnWidth(21, 150);
      /* backfill: เติมชื่อ-สกุลผู้จัดทำย้อนหลังจากชีต Users */
      const nameMap = {};
      if (ush && ush.getLastRow() > 1) {
        ush.getRange(2, 1, ush.getLastRow() - 1, 6).getValues().forEach(function(r) {
          const fn = String(r[5] || '').trim() || String(r[1]);
          nameMap[String(r[0])] = fn;
          nameMap[String(r[1])] = fn;
        });
      }
      const last = dsh.getLastRow();
      if (last > 1) {
        const creators = dsh.getRange(2, 17, last - 1, 1).getValues();
        dsh.getRange(2, 20, last - 1, 2).setValues(
          creators.map(function(r) { const n = nameMap[String(r[0])] || String(r[0] || ''); return [n, n]; }));
      }
    }
  } catch (e) {}
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
 * SECTION 2: DATABASE SETUP (รันครั้งเดียว)
 * ============================================================ */
const DB_SCHEMA = [
  { name: 'Users',         headers: ['รหัสผู้ใช้','ชื่อผู้ใช้','รหัสผ่าน (SHA-256)','ระดับการใช้งาน','Active','ชื่อ-สกุล'],
    widths: [90,160,340,130,80,200] },
  { name: 'Customers',     headers: ['รหัสลูกค้า','ชื่อลูกค้า','ที่อยู่','เบอร์โทรศัพท์','อีเมล','เลขที่ภาษีลูกค้า'],
    widths: [90,220,300,110,180,140] },
  { name: 'Products',      headers: ['รหัสสินค้า','ชื่อสินค้า','หมวดหมู่สินค้า','หน่วยนับ','ราคาต่อหน่วย','จำนวนในสต๊อก','จำนวนคงคลัง','จุดสั่งซื้อ','สถานะสินค้า'],
    widths: [90,260,120,90,110,110,110,100,110], moneyCols: ['E'] },
  { name: 'Settings',      headers: ['Key','Value'], widths: [160,420] },
  { name: 'Documents',     headers: ['เลขที่เอกสาร','ประเภท','วันที่เอกสาร','แสดงวันที่บนเอกสาร','รหัสลูกค้า','ชื่อลูกค้า','เรื่อง','กำหนดส่งมอบ (วัน)','กำหนดยืนราคา (วัน)','อ้างอิงเอกสาร','ยอดรวมก่อนภาษี','ภาษีมูลค่าเพิ่ม','ยอดรวมทั้งสิ้น','สถานะ','เหตุผลการยกเลิก','หมายเหตุ','สร้างโดย','สร้างเมื่อ','แก้ไขล่าสุด','จัดทำโดย (ชื่อ-สกุล)','แก้ไขโดย (ชื่อ-สกุล)'],
    widths: [110,70,100,130,90,190,240,110,110,110,120,110,120,90,180,180,90,140,140,150,150], moneyCols: ['K','L','M'] },
  { name: 'DocumentItems', headers: ['เลขที่เอกสาร','ลำดับ','รหัสสินค้า','ชื่อสินค้า','หน่วยนับ','จำนวน','ราคาต่อหน่วย','ราคารวม'],
    widths: [110,60,90,260,90,80,110,110], moneyCols: ['G','H'] },
  { name: 'Counters',      headers: ['Key (ประเภท-ปี)','เลขล่าสุด'], widths: [140,110] },
  { name: 'AuditLogs',     headers: ['วัน-เวลา','ผู้ใช้','การกระทำ','เลขที่เอกสาร','รายละเอียด'],
    widths: [150,90,110,110,480] },
  { name: 'Sessions',      headers: ['Token','รหัสผู้ใช้','ชื่อผู้ใช้','ระดับ','สร้างเมื่อ','หมดอายุ'],
    widths: [360,90,120,90,150,150] },
  { name: 'Payments',      headers: ['รหัสชำระเงิน','เลขที่เอกสาร','วันที่ชำระ','ยอดที่ชำระ','วิธีชำระเงิน','รายละเอียด (เลขเช็ค/โอน/อื่นๆ)','บันทึกโดย','บันทึกเมื่อ'],
    widths: [110,110,110,120,110,240,90,150], moneyCols: ['D'] }
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧾 Smart Billing')
    .addItem('ติดตั้งฐานข้อมูล (setupSystem)', 'setupSystem')
    .addItem('นำเข้าข้อมูลตัวอย่าง (importSampleData)', 'importSampleData')
    .addToUi();
}

function setupSystem() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const log = [];

  DB_SCHEMA.forEach(function(cfg) {
    let sh = ss.getSheetByName(cfg.name);
    /* heal: ชีตหายไป / ไม่มีหัวตาราง / ไม่ได้ freeze → จัดการใหม่ให้สมบูรณ์ */
    const needsInit = !sh ||
      String(sh.getRange(1, 1).getValue()) !== cfg.headers[0] ||
      sh.getFrozenRows() < 1;
    if (!sh) sh = ss.insertSheet(cfg.name);
    if (needsInit) {
      sh.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
      styleHeader_(sh, cfg.headers.length);
      (cfg.widths || []).forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
      (cfg.moneyCols || []).forEach(function(col) {
        sh.getRange(col + '2:' + col + '1000').setNumberFormat('#,##0.00');
      });
      sh.setFrozenRows(1);
      log.push('✅ เตรียมชีต: ' + cfg.name);
    } else {
      log.push('⏭️ มีอยู่แล้ว (ข้าม): ' + cfg.name);
    }
  });

  seedUsers_(ss, log);
  seedCustomers_(ss, log);
  seedProducts_(ss, log);
  seedSettings_(ss, log);
  seedDocuments_(ss, log);
  seedAuditLog_(ss, log);

  Logger.log(log.join('\n'));
  Logger.log('🎉 setupSystem เสร็จสมบูรณ์ — Login: admin / 1234');
}

function seedUsers_(ss, log) {
  const sh = ss.getSheetByName('Users');
  if (sh.getLastRow() > 1) return;
  sh.getRange(2, 1, 2, 6).setValues([
    ['U-001', 'admin', sha256_('1234'), 'Admin', true, 'ผู้ดูแลระบบ'],
    ['U-002', 'staff', sha256_('1234'), 'Staff', true, 'พนักงาน']
  ]);
  log.push('🌱 Users: admin/1234 (Admin), staff/1234 (Staff)');
}

function seedCustomers_(ss, log) {
  const sh = ss.getSheetByName('Customers');
  if (sh.getLastRow() > 1) return;
  sh.getRange(2, 1, 3, 6).setValues([
    ['C-0001', 'บริษัท นวัตกรรมไทย จำกัด', '99/1 อาคารสาทรทาวเวอร์ ชั้น 12 กรุงเทพฯ 10120', '02-123-4567', 'info@thainov.co.th', '0105558009991'],
    ['C-0002', 'ร้าน สยามพาณิชย์การค้า', '45/8 ถนนมิตรภาพ อ.เมือง จ.ขอนแก่น 40000', '043-222-111', 'siam_trade@gmail.com', '3101400293810'],
    ['C-0003', 'คุณสมชาย เข็มกลัด', '12 หมู่ 4 ต.บางพลี อ.บางพลี จ.สมุทรปราการ 10540', '081-987-6543', 'somchai_k@hotmail.com', '-']
  ]);
  log.push('🌱 Customers: 3 รายการ');
}

function seedProducts_(ss, log) {
  const sh = ss.getSheetByName('Products');
  if (sh.getLastRow() > 1) return;
  sh.getRange(2, 1, 5, 9).setValues([
    ['P-0001', 'บริการพัฒนาระบบ Web App (Google Apps Script)', 'บริการ', 'โครงการ', 25000, 100, 85, 5, 'ปกติ'],
    ['P-0002', 'อุปกรณ์เครื่องสแกนบาร์โค้ดไร้สาย', 'ฮาร์ดแวร์', 'เครื่อง', 2500, 20, 3, 5, 'จุดสั่งซื้อ'],
    ['P-0003', 'กระดาษความร้อน Thermal Paper 80mm', 'วัสดุสิ้นเปลือง', 'ม้วน', 45, 500, 420, 50, 'ปกติ'],
    ['P-0004', 'หมึกพิมพ์เครื่องถ่ายเอกสาร EPSON', 'วัสดุสิ้นเปลือง', 'กล่อง', 350, 10, 2, 5, 'จุดสั่งซื้อ'],
    ['P-0005', 'สัญญาดูแลระบบรายปี (MA Support)', 'บริการ', 'ปี', 12000, 50, 50, 2, 'ปกติ']
  ]);
  log.push('🌱 Products: 5 รายการ');
}

function seedSettings_(ss, log) {
  const sh = ss.getSheetByName('Settings');
  if (sh.getLastRow() > 1) return;
  sh.getRange(2, 1, 10, 2).setValues([
    ['logoUrl',        'https://via.placeholder.com/120x60.png?text=LOGO'],
    ['shopName',       'บริษัท สมาร์ท บิลลิ่ง โซลูชั่น จำกัด'],
    ['shopAddress',    '123/45 ถนนสุขุมวิท แขวงคลองเตย เขตคลองเตย กรุงเทพมหานคร 10110'],
    ['shopPhone',      '02-123-4567'],
    ['shopEmail',      'contact@smartbilling.co.th'],
    ['shopTaxId',      '0105566001234'],
    ['managerName',    'นายสมศักดิ์ มั่งมี'],
    ['signatoryName',  'นางสาววิภาดา ใจดี'],
    ['vatRate',        '0.07'],
    ['pdfFolderName',  'SmartBilling_PDF']
  ]);
  log.push('🌱 Settings: 10 ค่า');
}

function seedDocuments_(ss, log) {
  const docSh = ss.getSheetByName('Documents');
  const itemSh = ss.getSheetByName('DocumentItems');
  const cntSh = ss.getSheetByName('Counters');
  if (docSh.getLastRow() > 1) return;

  docSh.getRange(2, 1, 1, 19).setValues([[
    'QT26-0001', 'QT', '2026-08-15', true, 'C-0001', 'บริษัท นวัตกรรมไทย จำกัด',
    'เสนอราคาพัฒนาระบบเทคโนโลยีสารสนเทศ', 30, 30, '',
    54500, 3815, 58315, 'Active', '', 'ชำระเงินงวดแรก 50% ณ วันลงนามสัญญา',
    'admin', '2026-08-15 10:15:22', ''
  ]]);
  itemSh.getRange(2, 1, 2, 8).setValues([
    ['QT26-0001', 1, 'P-0001', 'บริการพัฒนาระบบ Web App (Google Apps Script)', 'โครงการ', 2, 25000, 50000],
    ['QT26-0001', 2, 'P-0003', 'กระดาษความร้อน Thermal Paper 80mm', 'ม้วน', 100, 45, 4500]
  ]);
  cntSh.getRange(2, 1, 1, 2).setValues([['QT-2026', 1]]);
  log.push('🌱 Documents: QT26-0001 + Counter QT-2026=1');
}

function seedAuditLog_(ss, log) {
  const sh = ss.getSheetByName('AuditLogs');
  if (sh.getLastRow() > 1) return;
  sh.getRange(2, 1, 1, 5).setValues([
    [new Date(), 'system', 'Setup', '-', 'ติดตั้งฐานข้อมูลและ Seed ข้อมูลตัวอย่าง']
  ]);
  log.push('🌱 AuditLogs: เรียบร้อย');
}

/* ============================================================
 * SECTION 2.5: SAMPLE DATA IMPORT (นำเข้าข้อมูลตัวอย่างชุดใหญ่)
 *  - เรียกจากเมนู "🧾 Smart Billing" ในชีต หรือรันอัตโนมัติครั้งแรก
 *    ผ่าน ensureSetup_() (flag: sampleImported ในชีต Settings)
 *  - ลูกค้า/สินค้า: ข้ามรายการที่ชื่อซ้ำ | เอกสาร: สร้างใหม่เสมอ (เลขไม่ซ้ำ)
 * ============================================================ */
function importSampleData(silent) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const log = [];
  const today = new Date();
  function daysAgo(n) { return Utilities.formatDate(new Date(today.getTime() - n * 86400000), TZ, 'yyyy-MM-dd'); }
  function stampAgo(n) {
    const d = new Date(today.getTime() - n * 86400000);
    return Utilities.formatDate(d, TZ, 'yyyy-MM-dd') + ' 09:' + pad_(10 + (n % 40), 2) + ':00';
  }

  /* ---------- 1) ลูกค้าตัวอย่าง +9 ราย ---------- */
  const SAMPLE_CUSTOMERS = [
    ['บริษัท สยามเทคโนโลยี จำกัด', '88/12 อาคารไพลินท์ ถนนสุขุมวิท แขวงคลองเตย กรุงเทพฯ 10110', '02-345-6789', 'accounting@siamtech.co.th', '0105558001234'],
    ['หจก. วิศวกรรมชลบุรี', '55/2 ถนนศรีรัชชา ต.เมือง อ.เมือง จ.ชลบุรี 20000', '038-288-999', 'chonburi.eng@gmail.com', '0325500111223'],
    ['บริษัท กรีนเอ็นเนอร์ยี่ จำกัด', '150/7 อาคารเอไอเอ สาทร แขวงสีลม กรุงเทพฯ 10120', '02-679-8899', 'contact@greenenergy.co.th', '0105561004455'],
    ['บริษัท ดิจิทัล เอเชีย ซอฟต์แวร์ จำกัด', '999 อาคารเดอะไนน์ทาวเวอร์ ชั้น 22 กรุงเทพฯ 10310', '02-100-8000', 'procurement@digitalasia.co.th', '0107559006677'],
    ['ห้างหุ้นส่วนจำกัด ประชาธิปไตยก่อสร้าง', '23 ถนนประชาธิปไตย ต.ตลาดขวัญ อ.เมือง จ.นนทบุรี 11000', '02-967-1122', 'prachathipatai.cc@gmail.com', '3542200111334'],
    ['บริษัท โลจิสติกส์ไทย ขนส่ง จำกัด', '888 หมู่ 5 นิคมอุตสาหกรรมบางปะอิน จ.พระนครศรีอยุธยา 13160', '035-377-888', 'logistic@thailogistic.co.th', '0109988007788'],
    ['บริษัท เมดิคอล ซัพพลาย จำกัด', '44/9 ถนนพญาไท เขตราชเทวี กรุงเทพฯ 10400', '02-645-3322', 'sales@medsupply.co.th', '0107777008899'],
    ['หจก. ธาราทิพย์ ฟู้ดส์', '67 ถนนเจริญเมือง ต.ศรีชมพู อ.เมือง จ.ลำปาง 52000', '054-321-555', 'tharathip.foods@gmail.com', '3881100224455'],
    ['บริษัท ออร์คิด รีสอร์ท แอนด์ สปา จำกัด', '12 หาดสุรินทร์ ต.เชิงทะเล อ.ถลาง จ.ภูเก็ต 83110', '076-316-999', 'reservation@orchidresort.co.th', '0108899009900']
  ];
  const cusSh = ss.getSheetByName('Customers');
  const haveCus = {};
  let maxCus = 0;
  if (cusSh.getLastRow() > 1) {
    cusSh.getRange(2, 1, cusSh.getLastRow() - 1, 2).getValues().forEach(function(r) {
      const m = String(r[0]).match(/^C-(\d+)$/);
      if (m) maxCus = Math.max(maxCus, parseInt(m[1], 10));
      haveCus[String(r[1]).trim()] = true;
    });
  }
  const newCusRows = [];
  SAMPLE_CUSTOMERS.forEach(function(c) {
    if (haveCus[c[0]]) return;
    maxCus++;
    newCusRows.push(['C-' + pad_(maxCus, 4)].concat(c));
  });
  if (newCusRows.length) {
    cusSh.getRange(cusSh.getLastRow() + 1, 1, newCusRows.length, 6).setValues(newCusRows);
    log.push('👥 Customers: +' + newCusRows.length + ' ราย');
  } else {
    log.push('👥 Customers: มีครบแล้ว (ข้าม)');
  }

  /* ---------- 2) สินค้าตัวอย่าง +8 ราย ---------- */
  const SAMPLE_PRODUCTS = [
    ['คีย์บอร์ดไร้สาย Logitech MX Keys', 'ฮาร์ดแวร์', 'ชุด', 3590, 15, 15, 5, 'ปกติ'],
    ['จอมอนิเตอร์ Dell UltraSharp 27 นิ้ว 4K', 'ฮาร์ดแวร์', 'ชุด', 12500, 8, 8, 3, 'ปกติ'],
    ['เครื่องพิมพ์เลเซอร์ HP LaserJet Pro M404dn', 'ฮาร์ดแวร์', 'ชุด', 8900, 6, 6, 2, 'ปกติ'],
    ['หมึกเลเซอร์ HP 26A Original', 'วัสดุสิ้นเปลือง', 'ชิ้น', 2450, 25, 25, 10, 'ปกติ'],
    ['โต๊ะทำงานไม้โอ๊ค 1.6 เมตร', 'เฟอร์นิเจอร์', 'ตัว', 4500, 10, 10, 4, 'ปกติ'],
    ['เก้าอี้ Mesh Ergonomic รุ่น Pro', 'เฟอร์นิเจอร์', 'ตัว', 2800, 18, 18, 6, 'ปกติ'],
    ['สาย LAN Cat6 ม้วน 305 เมตร', 'ฮาร์ดแวร์', 'ม้วน', 3200, 7, 7, 3, 'ปกติ'],
    ['กล้องวงจรปิด IP Camera 4K ชุดติดตั้ง', 'ฮาร์ดแวร์', 'ชุด', 5600, 9, 9, 3, 'ปกติ']
  ];
  const prodSh = ss.getSheetByName('Products');
  const haveProd = {};
  let maxProd = 0;
  if (prodSh.getLastRow() > 1) {
    prodSh.getRange(2, 1, prodSh.getLastRow() - 1, 2).getValues().forEach(function(r) {
      const m = String(r[0]).match(/^P-(\d+)$/);
      if (m) maxProd = Math.max(maxProd, parseInt(m[1], 10));
      haveProd[String(r[1]).trim()] = true;
    });
  }
  const newProdRows = [];
  SAMPLE_PRODUCTS.forEach(function(p) {
    if (haveProd[p[0]]) return;
    maxProd++;
    newProdRows.push(['P-' + pad_(maxProd, 4)].concat(p));
  });
  if (newProdRows.length) {
    prodSh.getRange(prodSh.getLastRow() + 1, 1, newProdRows.length, 9).setValues(newProdRows);
    log.push('📦 Products: +' + newProdRows.length + ' ราย');
  } else {
    log.push('📦 Products: มีครบแล้ว (ข้าม)');
  }

  /* ---------- 3) เอกสารตัวอย่าง 6 ฉบับ (QT→IV→RE ครบวงจร) ---------- */
  function idByName_(sheetName, nameColIdx, name) {
    const rows = readAll_(sheetName);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][nameColIdx]).trim() === name) return String(rows[i][0]);
    }
    return '';
  }
  const prodMap = {};
  readAll_('Products').forEach(function(r) {
    prodMap[String(r[1]).trim()] = { pid: String(r[0]), unit: r[3], price: Number(r[4]) || 0 };
  });
  const vatRate = Number(String(getSettings_().vatRate || '0.07')) || 0;
  const docSh = ss.getSheetByName('Documents');
  const itemSh = ss.getSheetByName('DocumentItems');

  function makeDoc_(o) {
    const docNo = generateDocNo_(o.type);
    let sub = 0;
    const itemRows = o.items.map(function(it, i) {
      const p = prodMap[it.name] || { pid: '', unit: it.unit || 'ชุด', price: it.price || 0 };
      const lineTotal = it.qty * p.price;
      sub += lineTotal;
      return [docNo, i + 1, p.pid, it.name, p.unit, it.qty, p.price, lineTotal];
    });
    const vat = Math.round(sub * vatRate * 100) / 100;
    docSh.appendRow(docRowOf_({
      docNo: docNo, docType: o.type, docDate: o.date, showDateOnPrint: true,
      cusId: o.cusId, cusName: o.cusName, subject: o.subject,
      sendDays: o.sendDays || 30, confirmDays: o.confirmDays || 30, refDocNo: o.ref || '',
      subTotal: sub, vatAmount: vat, grandTotal: sub + vat,
      status: 'Active', notes: o.notes || '', createdBy: 'admin', createdAt: o.stamp
    }));
    itemSh.getRange(itemSh.getLastRow() + 1, 1, itemRows.length, 8).setValues(itemRows);
    log.push('📄 ' + docNo + ' → ' + o.cusName + ' (' + sub.toFixed(2) + ' + VAT)');
    return docNo;
  }

  const cSiam = idByName_('Customers', 1, 'บริษัท สยามเทคโนโลยี จำกัด');
  const cGreen = idByName_('Customers', 1, 'บริษัท กรีนเอ็นเนอร์ยี่ จำกัด');
  const cDigi = idByName_('Customers', 1, 'บริษัท ดิจิทัล เอเชีย ซอฟต์แวร์ จำกัด');
  const cLogi = idByName_('Customers', 1, 'บริษัท โลจิสติกส์ไทย ขนส่ง จำกัด');

  const IT_BUNDLE = [
    { name: 'จอมอนิเตอร์ Dell UltraSharp 27 นิ้ว 4K', qty: 10 },
    { name: 'คีย์บอร์ดไร้สาย Logitech MX Keys', qty: 10 },
    { name: 'เก้าอี้ Mesh Ergonomic รุ่น Pro', qty: 10 }
  ];
  const qtSiam = makeDoc_({
    type: 'QT', date: daysAgo(12), stamp: stampAgo(12),
    cusId: cSiam, cusName: 'บริษัท สยามเทคโนโลยี จำกัด',
    subject: 'เสนอราคาอุปกรณ์ไอทีและเฟอร์นิเจอร์สำนักงานใหม่ 30 ตำแหน่ง',
    confirmDays: 30, notes: 'ราคานี้ไม่รวมภาษีหัก ณ ที่จ่าย', items: IT_BUNDLE
  });
  const ivSiam = makeDoc_({
    type: 'IV', date: daysAgo(8), stamp: stampAgo(8), ref: qtSiam,
    cusId: cSiam, cusName: 'บริษัท สยามเทคโนโลยี จำกัด',
    subject: 'ใบวางบิลอุปกรณ์ไอทีและเฟอร์นิเจอร์ (อ้างอิง ' + qtSiam + ')',
    sendDays: 7, notes: 'โอนเงินเข้าบัญชีธนาคารกรุงไทย', items: IT_BUNDLE
  });
  makeDoc_({
    type: 'RE', date: daysAgo(5), stamp: stampAgo(5), ref: ivSiam,
    cusId: cSiam, cusName: 'บริษัท สยามเทคโนโลยี จำกัด',
    subject: 'ใบเสร็จรับเงินค่าอุปกรณ์ไอทีและเฟอร์นิเจอร์ (ชำระเต็มจำนวน)',
    items: IT_BUNDLE
  });
  makeDoc_({
    type: 'QT', date: daysAgo(6), stamp: stampAgo(6),
    cusId: cGreen, cusName: 'บริษัท กรีนเอ็นเนอร์ยี่ จำกัด',
    subject: 'เสนอราคาเฟอร์นิเจอร์สำนักงานชั้น 8 อาคารเอไอเอ สาทร',
    confirmDays: 21, items: [
      { name: 'โต๊ะทำงานไม้โอ๊ค 1.6 เมตร', qty: 6 },
      { name: 'เก้าอี้ Mesh Ergonomic รุ่น Pro', qty: 12 },
      { name: 'เครื่องพิมพ์เลเซอร์ HP LaserJet Pro M404dn', qty: 1 }
    ]
  });
  makeDoc_({
    type: 'IV', date: daysAgo(3), stamp: stampAgo(3),
    cusId: cDigi, cusName: 'บริษัท ดิจิทัล เอเชีย ซอฟต์แวร์ จำกัด',
    subject: 'ค่าสัญญาดูแลระบบรายปีและวัสดุสิ้นเปลืองประจำไตรมาส',
    sendDays: 15, items: [
      { name: 'สัญญาดูแลระบบรายปี (MA Support)', qty: 2 },
      { name: 'หมึกเลเซอร์ HP 26A Original', qty: 4 },
      { name: 'สาย LAN Cat6 ม้วน 305 เมตร', qty: 2 }
    ]
  });
  makeDoc_({
    type: 'QT', date: daysAgo(1), stamp: stampAgo(1),
    cusId: cLogi, cusName: 'บริษัท โลจิสติกส์ไทย ขนส่ง จำกัด',
    subject: 'เสนอราคาติดตั้งกล้องวงจรปิดโกดังสินค้าบางปะอิน (8 จุด)',
    confirmDays: 14, notes: 'รวมค่าเดินสายและตั้งค่าระบบ NVR', items: [
      { name: 'กล้องวงจรปิด IP Camera 4K ชุดติดตั้ง', qty: 8 },
      { name: 'สาย LAN Cat6 ม้วน 305 เมตร', qty: 2 },
      { name: 'อุปกรณ์เครื่องสแกนบาร์โค้ดไร้สาย', qty: 2 }
    ]
  });

  /* ---------- 4) เก็บ flag + audit log ---------- */
  const setSh = ss.getSheetByName('Settings');
  const flagRow = findRowById_(setSh, 'sampleImported');
  if (flagRow > 0) setSh.getRange(flagRow, 2).setValue('yes');
  else setSh.appendRow(['sampleImported', 'yes']);

  logAudit_('system', 'ImportSamples', '-', 'นำเข้าข้อมูลตัวอย่าง: ' + log.join(' | '));

  Logger.log(log.join('\n'));
  Logger.log('🎉 นำเข้าข้อมูลตัวอย่างเสร็จสมบูรณ์');
  if (!silent) {
    try { SpreadsheetApp.getUi().alert('✅ นำเข้าข้อมูลตัวอย่างเรียบร้อย\n\n' + log.join('\n')); } catch (e) {}
  }
}

/* ============================================================
 * SECTION 3: SHEET HELPERS
 * ============================================================ */
function sheet_(name) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต "' + name + '" — กรุณารัน setupSystem() ก่อน');
  return sh;
}

function readAll_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function findRowById_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function nextId_(sheetName, prefix, padLen) {
  const rows = readAll_(sheetName);
  let max = 0;
  rows.forEach(function(r) {
    const m = String(r[0]).match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + pad_(max + 1, padLen);
}

function pad_(n, len) { return String(n).padStart(len, '0'); }

/* จัดสไตล์แถวหัวตาราง */
function styleHeader_(sh, numCols) {
  const hdr = sh.getRange(1, 1, 1, numCols);
  hdr.setBackground('#15803d')
     .setFontColor('#ffffff')
     .setFontWeight('bold')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle')
     .setWrap(true);
  sh.setRowHeight(1, 32);
}

function sha256_(str) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function fmtDateTime_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm:ss'); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStr_() { return fmtDateTime_(new Date()); }

function getSettings_() {
  const obj = {};
  readAll_('Settings').forEach(function(r) { obj[String(r[0])] = String(r[1]); });
  return obj;
}

function logAudit_(user, action, docNo, details) {
  try { sheet_('AuditLogs').appendRow([new Date(), user || '-', action || '-', docNo || '-', details || '-']); } catch (e) {}
}

/* ============================================================
 * SECTION 4: AUTHENTICATION & SESSION
 * ============================================================ */
function apiLogin(username, password) {
  const uname = String(username || '').trim().toLowerCase();
  const hash = sha256_(String(password));
  purgeExpiredSessions_();
  const rows = readAll_('Users');
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[1]).trim().toLowerCase() === uname && String(r[2]) === hash) {
      const active = r[4] === true || String(r[4]).toUpperCase() === 'TRUE';
      if (!active) return { ok: false, message: 'บัญชีนี้ถูกระงับการใช้งาน' };
      const token = Utilities.getUuid();
      const now = new Date();
      const expiry = new Date(now.getTime() + SESSION_HOURS * 3600 * 1000);
      sheet_('Sessions').appendRow([token, r[0], r[1], r[3], now, expiry]);
      logAudit_(r[1], 'Login', '-', 'เข้าสู่ระบบสำเร็จ');
      return { ok: true, token: token, user: { id: r[0], name: r[1], role: r[3], fullName: String(r[5] || '').trim() || r[1] } };
    }
  }
  return { ok: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function requireAuth_(token) {
  if (!token) throw new Error('SESSION_EXPIRED');
  const rows = readAll_('Sessions');
  const now = new Date();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(token)) {
      const expiry = rows[i][5] instanceof Date ? rows[i][5] : new Date(rows[i][5]);
      if (expiry < now) { deleteSessionRow_(i + 2); throw new Error('SESSION_EXPIRED'); }
      return { token: token, userId: rows[i][1], userName: rows[i][2], role: rows[i][3],
               fullName: getUserFullName_(rows[i][1]) || rows[i][2] };
    }
  }
  throw new Error('SESSION_EXPIRED');
}

/* ค้นชื่อ-สกุลจากรหัสผู้ใช้ (หรือชื่อผู้ใช้) — ไม่เจอคืนค่าว่าง */
function getUserFullName_(userIdOrName) {
  const key = String(userIdOrName || '');
  const hit = readAll_('Users').filter(function(r) { return String(r[0]) === key || String(r[1]) === key; })[0];
  return hit ? (String(hit[5] || '').trim() || String(hit[1])) : '';
}

function requireAdmin_(session) {
  if (String(session.role).toLowerCase() !== 'admin') {
    throw new Error('ต้องการสิทธิ์ Admin เท่านั้น');
  }
}

function deleteSessionRow_(row) {
  try { sheet_('Sessions').deleteRow(row); } catch (e) {}
}

function purgeExpiredSessions_() {
  try {
    const sh = sheet_('Sessions');
    const last = sh.getLastRow();
    if (last < 2) return;
    const now = new Date();
    const data = sh.getRange(2, 1, last - 1, 6).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const exp = data[i][5] instanceof Date ? data[i][5] : new Date(data[i][5]);
      if (exp < now) sh.deleteRow(i + 2);
    }
  } catch (e) {}
}

function apiLogout(token) {
  try {
    const s = requireAuth_(token);
    for (let i = 0; i < readAll_('Sessions').length; i++) {}
    const sh = sheet_('Sessions');
    const row = findRowById_(sh, token);
    if (row > 0) sh.deleteRow(row);
    logAudit_(s.userName, 'Logout', '-', 'ออกจากระบบ');
  } catch (e) {}
  return { ok: true };
}

/* ============================================================
 * SECTION 5: BOOTSTRAP (โหลดข้อมูลทั้งหมดครั้งเดียว)
 * ============================================================ */
function apiGetBootstrap(token) {
  requireAuth_(token);
  return {
    ok: true,
    settings: getSettings_(),
    customers: listCustomers_(),
    products: listProducts_(),
    users: listUsersSanitized_(),
    documents: listDocuments_(),
    logs: listAuditLogs_(150)
  };
}

function listCustomers_() {
  return readAll_('Customers').map(function(r) {
    return { id: String(r[0]), name: String(r[1]), address: String(r[2]), phone: String(r[3]), email: String(r[4]), taxId: String(r[5]) };
  });
}

function listProducts_() {
  return readAll_('Products').map(function(r) {
    return { id: String(r[0]), name: String(r[1]), category: String(r[2]), unit: String(r[3]),
             price: Number(r[4]) || 0, stockQty: Number(r[5]) || 0, remainingQty: Number(r[6]) || 0,
             reorderPoint: Number(r[7]) || 0, status: String(r[8] || 'ปกติ') };
  });
}

function listUsersSanitized_() {
  return readAll_('Users').map(function(r) {
    return { id: String(r[0]), username: String(r[1]), role: String(r[3]),
             active: r[4] === true || String(r[4]).toUpperCase() === 'TRUE',
             fullName: String(r[5] || '') };
  });
}

function rawDocuments_() {
  return readAll_('Documents').map(function(r) {
    return {
      docNo: String(r[0]), docType: String(r[1]), docDate: String(r[2]),
      showDateOnPrint: r[3] === true || String(r[3]).toUpperCase() === 'TRUE',
      cusId: String(r[4]), cusName: String(r[5]), subject: String(r[6]),
      sendDays: Number(r[7]) || 30, confirmDays: Number(r[8]) || 30,
      refDocNo: String(r[9] || ''), subTotal: Number(r[10]) || 0,
      vatAmount: Number(r[11]) || 0, grandTotal: Number(r[12]) || 0,
      status: String(r[13] || 'Active'), cancelReason: String(r[14] || ''),
      notes: String(r[15] || ''), createdBy: String(r[16] || ''),
      createdAt: String(r[17] || ''), updatedAt: String(r[18] || ''),
      createdByName: String(r[19] || ''), updatedByName: String(r[20] || '')
    };
  });
}

function itemsOf_(docNo) {
  return readAll_('DocumentItems')
    .filter(function(r) { return String(r[0]) === String(docNo); })
    .sort(function(a, b) { return (Number(a[1]) || 0) - (Number(b[1]) || 0); })
    .map(function(r) {
      return { prodId: String(r[2]), prodName: String(r[3]), unit: String(r[4]),
               qty: Number(r[5]) || 0, unitPrice: Number(r[6]) || 0, totalPrice: Number(r[7]) || 0 };
    });
}

function listDocuments_() {
  const docs = rawDocuments_();
  docs.forEach(function(d) { d.items = itemsOf_(d.docNo); });

  /* หา "จัดพิมพ์ล่าสุด" ต่อเอกสารจาก AuditLogs (Print / ExportPDF) */
  const prints = {};
  readAll_('AuditLogs').forEach(function(r) {
    const act = String(r[2]);
    if (act !== 'Print' && act !== 'ExportPDF') return;
    const dn = String(r[3]);
    if (!dn || dn === '-') return;
    const at = r[0] instanceof Date ? fmtDateTime_(r[0]) : String(r[0]);
    if (!prints[dn] || at > prints[dn].at) prints[dn] = { by: String(r[1]), at: at };
  });
  docs.forEach(function(d) { d.lastPrint = prints[d.docNo] || null; });

  docs.sort(function(a, b) { return b.createdAt.localeCompare(a.createdAt); });
  return docs;
}

function listAuditLogs_(limit) {
  const rows = readAll_('AuditLogs').map(function(r) {
    return { timestamp: r[0] instanceof Date ? fmtDateTime_(r[0]) : String(r[0]),
             user: String(r[1]), action: String(r[2]), docNo: String(r[3]), details: String(r[4]) };
  });
  rows.reverse();
  return limit ? rows.slice(0, limit) : rows;
}

/* ============================================================
 * SECTION 6: MASTER DATA CRUD
 * ============================================================ */
function apiSaveCustomer(token, cus) {
  const s = requireAuth_(token);
  if (!cus || !cus.name || !String(cus.name).trim()) throw new Error('กรุณาระบุชื่อลูกค้า');
  const sh = sheet_('Customers');
  const row = [cus.id || '', String(cus.name), cus.address || '', cus.phone || '', cus.email || '', cus.taxId || ''];
  if (!cus.id) {
    row[0] = nextId_('Customers', 'C-', 4);
    sh.appendRow(row);
    logAudit_(s.userName, 'AddCustomer', row[0], 'เพิ่มลูกค้า: ' + cus.name);
  } else {
    const idx = findRowById_(sh, cus.id);
    if (idx < 0) throw new Error('ไม่พบลูกค้ารหัส ' + cus.id);
    row[0] = cus.id;
    sh.getRange(idx, 1, 1, 6).setValues([row]);
    logAudit_(s.userName, 'EditCustomer', cus.id, 'แก้ไขลูกค้า: ' + cus.name);
  }
  return { ok: true, customer: { id: row[0], name: row[1], address: row[2], phone: row[3], email: row[4], taxId: row[5] } };
}

function apiDeleteCustomer(token, id) {
  const s = requireAuth_(token);
  const sh = sheet_('Customers');
  const idx = findRowById_(sh, id);
  if (idx < 0) throw new Error('ไม่พบลูกค้ารหัส ' + id);
  sh.deleteRow(idx);
  logAudit_(s.userName, 'DeleteCustomer', id, 'ลบลูกค้ารหัส ' + id);
  return { ok: true };
}

function apiSaveProduct(token, p) {
  const s = requireAuth_(token);
  if (!p || !p.name || !String(p.name).trim()) throw new Error('กรุณาระบุชื่อสินค้า');
  const sh = sheet_('Products');
  const status = p.status || ((Number(p.remainingQty) <= Number(p.reorderPoint)) ? 'จุดสั่งซื้อ' : 'ปกติ');
  const row = ['', String(p.name), p.category || 'ทั่วไป', p.unit || 'ชิ้น',
               Number(p.price) || 0, Number(p.stockQty) || 0, Number(p.remainingQty) || 0,
               Number(p.reorderPoint) || 0, status];
  if (!p.id) {
    row[0] = nextId_('Products', 'P-', 4);
    sh.appendRow(row);
    logAudit_(s.userName, 'AddProduct', row[0], 'เพิ่มสินค้า: ' + p.name);
  } else {
    const idx = findRowById_(sh, p.id);
    if (idx < 0) throw new Error('ไม่พบสินค้ารหัส ' + p.id);
    row[0] = p.id;
    sh.getRange(idx, 1, 1, 9).setValues([row]);
    logAudit_(s.userName, 'EditProduct', p.id, 'แก้ไขสินค้า: ' + p.name);
  }
  return { ok: true, product: { id: row[0], name: row[1], category: row[2], unit: row[3],
           price: row[4], stockQty: row[5], remainingQty: row[6], reorderPoint: row[7], status: row[8] } };
}

function apiDeleteProduct(token, id) {
  const s = requireAuth_(token);
  const sh = sheet_('Products');
  const idx = findRowById_(sh, id);
  if (idx < 0) throw new Error('ไม่พบสินค้ารหัส ' + id);
  sh.deleteRow(idx);
  logAudit_(s.userName, 'DeleteProduct', id, 'ลบสินค้ารหัส ' + id);
  return { ok: true };
}

function apiSaveUser(token, u) {
  const s = requireAuth_(token);
  requireAdmin_(s);
  if (!u || !u.username || !String(u.username).trim()) throw new Error('กรุณาระบุชื่อผู้ใช้');
  const sh = sheet_('Users');
  if (!u.id) {
    if (!u.password) throw new Error('กรุณาระบุรหัสผ่านสำหรับผู้ใช้ใหม่');
    const id = nextId_('Users', 'U-', 3);
    sh.appendRow([id, String(u.username), sha256_(String(u.password)), u.role || 'Staff', u.active !== false, String(u.fullName || '').trim()]);
    logAudit_(s.userName, 'AddUser', id, 'เพิ่มผู้ใช้: ' + u.username + ' (' + (u.role || 'Staff') + ')' + (u.fullName ? ' ชื่อ-สกุล: ' + u.fullName : ''));
    return { ok: true };
  }
  const idx = findRowById_(sh, u.id);
  if (idx < 0) throw new Error('ไม่พบผู้ใช้รหัส ' + u.id);
  const existingHash = sh.getRange(idx, 3).getValue();
  const newHash = u.password ? sha256_(String(u.password)) : existingHash;
  sh.getRange(idx, 1, 1, 6).setValues([[u.id, String(u.username), newHash, u.role || 'Staff', u.active !== false, String(u.fullName || '').trim()]]);
  logAudit_(s.userName, 'EditUser', u.id, 'แก้ไขผู้ใช้: ' + u.username);
  return { ok: true };
}

function apiDeleteUser(token, id) {
  const s = requireAuth_(token);
  requireAdmin_(s);
  if (String(id) === String(s.userId)) throw new Error('ไม่สามารถลบบัญชีที่กำลังใช้งานอยู่ได้');
  const sh = sheet_('Users');
  const idx = findRowById_(sh, id);
  if (idx < 0) throw new Error('ไม่พบผู้ใช้รหัส ' + id);
  sh.deleteRow(idx);
  logAudit_(s.userName, 'DeleteUser', id, 'ลบผู้ใช้รหัส ' + id);
  return { ok: true };
}

function apiSaveSettings(token, settings) {
  const s = requireAuth_(token);
  requireAdmin_(s);
  const sh = sheet_('Settings');
  Object.keys(settings).forEach(function(key) {
    const idx = findRowById_(sh, key);
    if (idx > 0) sh.getRange(idx, 2).setValue(String(settings[key]));
    else sh.appendRow([key, String(settings[key])]);
  });
  logAudit_(s.userName, 'Settings', '-', 'อัปเดตการตั้งค่าร้านค้า');
  return { ok: true, settings: getSettings_() };
}

/* ============================================================
 * SECTION 7: DOCUMENT ENGINE (QT / IV / RE / RC)
 * ============================================================ */
const DOC_TYPES = { QT: 'QT', IV: 'IV', RE: 'RE', RC: 'RC' };

function generateDocNo_(type) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const year = new Date().getFullYear();
    const key = type + '-' + year;
    const sh = sheet_('Counters');
    const data = sh.getDataRange().getValues();
    let last = 0, rowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === key) { last = Number(data[i][1]) || 0; rowIdx = i + 1; break; }
    }
    last++;
    if (rowIdx < 0) sh.appendRow([key, last]);
    else sh.getRange(rowIdx, 2).setValue(last);
    return type + String(year).slice(-2) + '-' + pad_(last, 4);
  } finally {
    lock.releaseLock();
  }
}

function validateDoc_(doc) {
  if (!doc || !DOC_TYPES[doc.docType]) throw new Error('ประเภทเอกสารไม่ถูกต้อง');
  if (!doc.items || !doc.items.length) throw new Error('เอกสารต้องมีรายการสินค้าอย่างน้อย 1 รายการ');
  if (!doc.cusId) throw new Error('กรุณาเลือกลูกค้า');
  doc.items.forEach(function(it, i) {
    if (!it.prodName || !String(it.prodName).trim()) throw new Error('รายการที่ ' + (i + 1) + ' ไม่มีชื่อสินค้า');
  });
}

function docRowOf_(doc) {
  return [doc.docNo, doc.docType, doc.docDate || todayStr_(), !!doc.showDateOnPrint,
          doc.cusId, doc.cusName || '', doc.subject || '',
          Number(doc.sendDays) || 30, Number(doc.confirmDays) || 30, doc.refDocNo || '',
          Number(doc.subTotal) || 0, Number(doc.vatAmount) || 0, Number(doc.grandTotal) || 0,
          doc.status || 'Active', doc.cancelReason || '', doc.notes || '',
          doc.createdBy || '', doc.createdAt || nowStr_(), nowStr_(),
          doc.createdByName || '', doc.updatedByName || ''];
}

function appendItems_(docNo, items) {
  const sh = sheet_('DocumentItems');
  const start = sh.getLastRow() + 1;
  const rows = items.map(function(it, i) {
    return [docNo, i + 1, it.prodId || 'CUSTOM', String(it.prodName), it.unit || '',
            Number(it.qty) || 0, Number(it.unitPrice) || 0, Number(it.totalPrice) || 0];
  });
  sh.getRange(start, 1, rows.length, 8).setValues(rows);
}

function deleteItemsOf_(docNo) {
  const sh = sheet_('DocumentItems');
  const last = sh.getLastRow();
  if (last < 2) return;
  const colA = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]) === String(docNo)) sh.deleteRow(i + 2);
  }
}

function apiSaveDocument(token, doc) {
  const s = requireAuth_(token);
  validateDoc_(doc);

  const subTotal = doc.items.reduce(function(acc, it) { return acc + (Number(it.totalPrice) || 0); }, 0);
  const vatRate = parseFloat(getSettings_().vatRate) || 0.07;
  const vatAmount = doc.vatEnabled === false ? 0 : subTotal * vatRate;
  doc.subTotal = subTotal;
  doc.vatAmount = vatAmount;
  doc.grandTotal = subTotal + vatAmount;

  const sh = sheet_('Documents');

  if (!doc.docNo) {
    doc.docNo = generateDocNo_(DOC_TYPES[doc.docType]);
    doc.status = 'Active';
    doc.createdBy = s.userName;
    doc.createdAt = nowStr_();
    doc.createdByName = s.fullName || s.userName;
    doc.updatedByName = doc.createdByName;
    sh.appendRow(docRowOf_(doc));
    appendItems_(doc.docNo, doc.items);

    /* ใบเสร็จเงินสด (RC): บันทึกการชำระเงินให้อัตโนมัติ — ข้าม QT/IV ได้ทันที */
    let payId = '';
    if (doc.docType === 'RC') {
      const method = PAY_METHODS[doc.payMethod] ? doc.payMethod : 'CASH';
      const detail = String(doc.payDetail || '').trim();
      if (method !== 'CASH' && !detail) {
        throw new Error(method === 'TRANSFER' ? 'กรุณาระบุข้อมูลการโอน (ธนาคาร/เลขบัญชี/เลขอ้างอิง)'
          : method === 'CHEQUE' ? 'กรุณาระบุเลขที่เช็ค/ธนาคาร'
          : 'กรุณาระบุรายละเอียดวิธีชำระเงิน');
      }
      payId = recordPayment_(s.userName, doc.docNo, doc.docDate, doc.grandTotal, method, detail);
    }

    logAudit_(s.userName, 'CreateDoc', doc.docNo,
      'สร้าง' + typeLabel_(doc.docType) + ' ลูกค้า: ' + (doc.cusName || doc.cusId) + ' ยอดรวม ' + doc.grandTotal.toFixed(2) + ' บาท'
      + (payId ? ' | ชำระแล้ว (' + PAY_METHODS[doc.payMethod && PAY_METHODS[doc.payMethod] ? doc.payMethod : 'CASH'] + ')' : ''));
    const saved = withItems_(doc);
    if (payId) { saved.payments = getPaymentsOf_(doc.docNo); saved.paidTotal = doc.grandTotal; }
    return { ok: true, doc: saved };
  }

  const idx = findRowById_(sh, doc.docNo);
  if (idx < 0) throw new Error('ไม่พบเอกสาร ' + doc.docNo);
  const existing = rawDocuments_().filter(function(d) { return d.docNo === doc.docNo; })[0];
  if (existing.status === 'Cancelled') throw new Error('เอกสารถูกยกเลิกแล้ว แก้ไขไม่ได้');
  doc.status = existing.status;
  doc.createdBy = existing.createdBy;
  doc.createdAt = existing.createdAt;
  doc.createdByName = existing.createdByName || getUserFullName_(existing.createdBy) || existing.createdBy;
  doc.updatedByName = s.fullName || s.userName;
  sh.getRange(idx, 1, 1, 21).setValues([docRowOf_(doc)]);
  deleteItemsOf_(doc.docNo);
  appendItems_(doc.docNo, doc.items);
  logAudit_(s.userName, 'EditDoc', doc.docNo, 'แก้ไขเอกสาร ' + doc.docNo);
  return { ok: true, doc: withItems_(doc) };
}

function withItems_(doc) {
  doc.items = itemsOf_(doc.docNo);
  return doc;
}

function typeLabel_(t) {
  return { QT: 'ใบเสนอราคา', IV: 'ใบวางบิล/ใบส่งของ', RE: 'ใบเสร็จ/ใบกำกับภาษี', RC: 'ใบเสร็จเงินสด' }[t] || 'เอกสาร';
}

function apiCancelDocument(token, docNo, reason) {
  const s = requireAuth_(token);
  if (!reason || !String(reason).trim()) throw new Error('จำเป็นต้องระบุเหตุผลในการยกเลิก');
  const sh = sheet_('Documents');
  const idx = findRowById_(sh, docNo);
  if (idx < 0) throw new Error('ไม่พบเอกสาร ' + docNo);
  const statusCell = sh.getRange(idx, 14);
  if (String(statusCell.getValue()) === 'Cancelled') throw new Error('เอกสารนี้ถูกยกเลิกไปแล้ว');
  statusCell.setValue('Cancelled');
  sh.getRange(idx, 15).setValue(String(reason).trim());
  sh.getRange(idx, 19).setValue(nowStr_());
  logAudit_(s.userName, 'CancelDoc', docNo, 'ยกเลิกเอกสาร เหตุผล: ' + String(reason).trim());
  return { ok: true };
}

function apiConvertQtToBilling(token, qtDocNo) {
  const s = requireAuth_(token);
  const all = rawDocuments_();
  const qt = all.filter(function(d) { return d.docNo === String(qtDocNo); })[0];
  if (!qt) throw new Error('ไม่พบใบเสนอราคา ' + qtDocNo);
  if (qt.docType !== 'QT') throw new Error('เอกสารนี้ไม่ใช่ใบเสนอราคา');
  if (qt.status !== 'Active') throw new Error('ยกเลิก/วางบิลใบเสนอราคานี้ไปแล้ว');

  const inv = {
    docType: 'IV',
    docDate: todayStr_(),
    showDateOnPrint: true,
    cusId: qt.cusId, cusName: qt.cusName,
    subject: qt.subject, sendDays: qt.sendDays, confirmDays: qt.confirmDays,
    refDocNo: qt.docNo,
    subTotal: qt.subTotal, vatAmount: qt.vatAmount, grandTotal: qt.grandTotal,
    notes: qt.notes, createdBy: s.userName, createdAt: nowStr_(),
    createdByName: s.fullName || s.userName, updatedByName: s.fullName || s.userName
  };
  validateDoc_({ docType: 'IV', cusId: inv.cusId, items: [{ prodName: 'temp' }] });
  inv.docNo = generateDocNo_('IV');
  inv.status = 'Active';
  sheet_('Documents').appendRow(docRowOf_(inv));
  appendItems_(inv.docNo, itemsOf_(qt.docNo));

  const sh = sheet_('Documents');
  const idx = findRowById_(sh, qt.docNo);
  sh.getRange(idx, 14).setValue('Billed');
  sh.getRange(idx, 19).setValue(nowStr_());

  logAudit_(s.userName, 'ConvertDoc', inv.docNo, 'แปลงใบเสนอราคา ' + qt.docNo + ' → ใบวางบิล ' + inv.docNo);
  return { ok: true, invoice: withItems_(inv) };
}

function apiLogPrint(token, docNo, channel) {
  const s = requireAuth_(token);
  const label = channel === 'PDF' ? 'ส่งออก PDF' : 'สั่งพิมพ์เอกสาร';
  logAudit_(s.userName, channel === 'PDF' ? 'ExportPDF' : 'Print', docNo, label + ' ' + docNo);
  return { ok: true };
}

/* ============================================================
 * SECTION 7.5: PAYMENTS (บันทึกรับชำระเงิน) + ใบเสร็จเงินสด (RC)
 * ============================================================ */
const PAY_METHODS = { CASH: 'เงินสด', TRANSFER: 'เงินโอน', CHEQUE: 'เช็ค', OTHER: 'อื่นๆ' };

function getPaidTotal_(docNo) {
  return readAll_('Payments')
    .filter(function(r) { return String(r[1]) === String(docNo); })
    .reduce(function(s, r) { return s + (Number(r[3]) || 0); }, 0);
}

function getPaymentsOf_(docNo) {
  return readAll_('Payments')
    .filter(function(r) { return String(r[1]) === String(docNo); })
    .map(function(r) {
      return { payId: String(r[0]), docNo: String(r[1]), payDate: String(r[2]),
               amount: Number(r[3]) || 0, method: String(r[4]), methodLabel: PAY_METHODS[String(r[4])] || String(r[4]),
               detail: String(r[5]), createdBy: String(r[6]), createdAt: String(r[7]) };
    });
}

function recordPayment_(username, docNo, payDate, amount, method, detail) {
  const payId = nextId_('Payments', 'PAY-', 4);
  sheet_('Payments').appendRow([payId, String(docNo), payDate || todayStr_(), Number(amount),
                                method, String(detail || ''), username, nowStr_()]);
  return payId;
}

/* บันทึกการรับชำระเงินของใบวางบิล (IV) — ต้องชำระครบก่อนออกใบเสร็จ RE */
function apiAddPayment(token, docNo, payDate, amount, method, detail) {
  const s = requireAuth_(token);
  const sh = sheet_('Documents');
  const idx = findRowById_(sh, docNo);
  if (idx < 0) throw new Error('ไม่พบเอกสาร ' + docNo);
  const vals = sh.getRange(idx, 1, 1, 19).getValues()[0];
  if (String(vals[1]) !== 'IV') throw new Error('บันทึกชำระเงินได้เฉพาะใบวางบิล (IV) เท่านั้น');
  if (String(vals[13]) === 'Cancelled') throw new Error('เอกสารถูกยกเลิกแล้ว');
  amount = Number(amount);
  if (!(amount > 0)) throw new Error('จำนวนเงินต้องมากกว่า 0');
  if (!PAY_METHODS[method]) throw new Error('กรุณาเลือกวิธีชำระเงิน');
  detail = String(detail || '').trim();
  /* เงินโอน/เช็ค/อื่นๆ → บังคับระบุข้อมูลที่เกี่ยวข้อง */
  if (method !== 'CASH' && !detail) {
    throw new Error(method === 'TRANSFER' ? 'กรุณาระบุข้อมูลการโอน (ธนาคาร/เลขบัญชี/เลขอ้างอิง)'
      : method === 'CHEQUE' ? 'กรุณาระบุเลขที่เช็ค/ธนาคาร'
      : 'กรุณาระบุรายละเอียดวิธีชำระเงิน');
  }
  const grandTotal = Number(vals[12]) || 0;
  const paidBefore = getPaidTotal_(docNo);
  if (paidBefore + amount > grandTotal + 0.01) {
    throw new Error('ยอดเกินคงค้าง! ชำระแล้ว ' + paidBefore.toFixed(2) + ' / ทั้งหมด ' + grandTotal.toFixed(2) + ' → ชำระได้อีกไม่เกิน ' + Math.max(0, grandTotal - paidBefore).toFixed(2));
  }
  const payId = recordPayment_(s.userName, docNo, payDate, amount, method, detail);
  logAudit_(s.userName, 'AddPayment', docNo,
    payId + ' | ' + PAY_METHODS[method] + (detail ? ' (' + detail + ')' : '') + ' | ' + amount.toFixed(2) + ' บาท');
  return { ok: true, payId: payId, paidTotal: paidBefore + amount, grandTotal: grandTotal };
}

/* ดึงประวัติการชำระเงินของเอกสาร */
function apiGetPayments(token, docNo) {
  requireAuth_(token);
  const sh = sheet_('Documents');
  const idx = findRowById_(sh, docNo);
  if (idx < 0) throw new Error('ไม่พบเอกสาร ' + docNo);
  const grandTotal = Number(sh.getRange(idx, 13).getValue()) || 0;
  const payments = getPaymentsOf_(docNo);
  const paidTotal = payments.reduce(function(s, p) { return s + p.amount; }, 0);
  return { ok: true, payments: payments, paidTotal: paidTotal, grandTotal: grandTotal };
}

/* IV → RE: ออกใบเสร็จ/ใบกำกับภาษีได้เมื่อชำระเงินครบแล้วเท่านั้น */
function apiConvertBillingToReceipt(token, ivDocNo) {
  const s = requireAuth_(token);
  const all = rawDocuments_();
  const iv = all.filter(function(d) { return d.docNo === String(ivDocNo); })[0];
  if (!iv) throw new Error('ไม่พบใบวางบิล ' + ivDocNo);
  if (iv.docType !== 'IV') throw new Error('เอกสารนี้ไม่ใช่ใบวางบิล');
  if (iv.status !== 'Active') throw new Error('ใบวางบิลนี้ถูกยกเลิกหรือออกใบเสร็จไปแล้ว');

  const paid = getPaidTotal_(iv.docNo);
  if (paid < iv.grandTotal - 0.01) {
    throw new Error('ยังชำระเงินไม่ครบ (ชำระแล้ว ' + paid.toFixed(2) + ' / ทั้งหมด ' + iv.grandTotal.toFixed(2) + ') — กรุณาบันทึกการรับชำระเงิน (💰) ก่อนออกใบเสร็จ');
  }

  const receipt = {
    docType: 'RE',
    docDate: todayStr_(),
    showDateOnPrint: true,
    cusId: iv.cusId, cusName: iv.cusName,
    subject: iv.subject, sendDays: iv.sendDays, confirmDays: iv.confirmDays,
    refDocNo: iv.docNo,
    subTotal: iv.subTotal, vatAmount: iv.vatAmount, grandTotal: iv.grandTotal,
    notes: iv.notes, createdBy: s.userName, createdAt: nowStr_(),
    createdByName: s.fullName || s.userName, updatedByName: s.fullName || s.userName
  };
  validateDoc_({ docType: 'RE', cusId: receipt.cusId, items: [{ prodName: 'temp' }] });
  receipt.docNo = generateDocNo_('RE');
  receipt.status = 'Active';
  sheet_('Documents').appendRow(docRowOf_(receipt));
  appendItems_(receipt.docNo, itemsOf_(iv.docNo));

  const sh = sheet_('Documents');
  const idx = findRowById_(sh, iv.docNo);
  sh.getRange(idx, 14).setValue('Paid');
  sh.getRange(idx, 19).setValue(nowStr_());

  logAudit_(s.userName, 'ConvertDoc', receipt.docNo, 'ออกใบเสร็จ ' + receipt.docNo + ' จากใบวางบิล ' + iv.docNo + ' (ชำระครบ ' + paid.toFixed(2) + ' บาท)');
  return { ok: true, receipt: withItems_(receipt) };
}

/* ============================================================
 * SECTION 8: PDF EXPORT (บันทึกลง Google Drive)
 * ============================================================ */
function getOrCreateFolder_(name) {
  /* ลำดับความสำคัญ: pdfFolderId จากชีต Settings → DEFAULT_PDF_FOLDER_ID → สร้างตามชื่อ */
  const st = getSettings_();
  const fid = st.pdfFolderId || DEFAULT_PDF_FOLDER_ID;
  if (fid) {
    try { return DriveApp.getFolderById(fid); } catch (e) { /* id ไม่ถูกต้อง → ไปทางชื่อโฟลเดอร์ */ }
  }
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function apiExportPdf(token, docNo) {
  const s = requireAuth_(token);
  const doc = rawDocuments_().filter(function(d) { return d.docNo === String(docNo); })[0];
  if (!doc) throw new Error('ไม่พบเอกสาร ' + docNo);
  const items = itemsOf_(docNo);
  const settings = getSettings_();
  const html = buildStandaloneHtml_(doc, items, settings);
  const pdfBlob = Utilities.newBlob(html, 'text/html', docNo + '.html')
                          .getAs('application/pdf').setName(docNo + '.pdf');
  const folder = getOrCreateFolder_(settings.pdfFolderName || 'SmartBilling_PDF');
  const file = folder.createFile(pdfBlob);
  logAudit_(s.userName, 'ExportPDF', docNo, 'ส่งออก PDF → ' + file.getUrl());
  return { ok: true, url: file.getUrl(), fileName: docNo + '.pdf' };
}

function esc_(v) {
  return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function thaiMoney_(n) { return Number(n || 0).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

function buildStandaloneHtml_(doc, items, st) {
  const pages = [];
  for (let i = 0; i < items.length; i += ITEMS_PER_PAGE) pages.push(items.slice(i, i + ITEMS_PER_PAGE));
  if (!pages.length) pages.push([]);
  const totalPages = pages.length;
  const dateText = doc.showDateOnPrint ? esc_(doc.docDate) : '____________________';

  const head =
    '<div style="display:flex;justify-content:space-between;border-bottom:2px solid #333;padding-bottom:8px;margin-bottom:8px">' +
    '<div><div style="font-size:18px;font-weight:bold">' + esc_(st.shopName) + '</div>' +
    '<div style="font-size:11px">ที่อยู่: ' + esc_(st.shopAddress) + '</div>' +
    '<div style="font-size:11px">โทร: ' + esc_(st.shopPhone) + ' | เลขประจำตัวผู้เสียภาษี: ' + esc_(st.shopTaxId) + '</div></div>' +
    '<div style="text-align:right"><div style="font-size:16px;font-weight:bold">' + esc_(typeLabel_(doc.docType)) + '</div>' +
    '<div style="font-size:12px;font-weight:bold">' + esc_(doc.docNo) + '</div>' +
    '<div style="font-size:11px">วันที่: ' + dateText + '</div>' +
    (doc.refDocNo ? '<div style="font-size:11px">อ้างอิง: ' + esc_(doc.refDocNo) + '</div>' : '') +
    '</div></div>';

  const custInfo =
    '<table style="width:100%;font-size:12px;margin-bottom:8px;background:#f5f5f5"><tr>' +
    '<td><b>ลูกค้า:</b> ' + esc_(doc.cusName) + '<br>' + esc_(st.shopAddress ? '' : '') + '</td>' +
    '<td style="text-align:right"><b>เรื่อง:</b> ' + esc_(doc.subject || '-') + '</td></tr></table>';

  let body = '';
  pages.forEach(function(pageItems, pi) {
    const isLast = pi === totalPages - 1;
    body += '<div style="' + (pi < totalPages - 1 ? 'page-break-after:always;' : '') + '">';
    body += '<div style="text-align:right;font-size:11px;margin-bottom:4px">หน้า ' + (pi + 1) + '/' + totalPages + '</div>';
    if (pi === 0) body += head + custInfo;
    body += '<table style="width:100%;border-collapse:collapse;font-size:12px" border="1" cellpadding="4">' +
      '<tr style="background:#e5e7eb"><th style="width:40px">ลำดับ</th><th>รายการ</th><th style="width:60px">จำนวน</th>' +
      '<th style="width:90px">ราคา/หน่วย</th><th style="width:100px">จำนวนเงิน</th></tr>';
    pageItems.forEach(function(it, ii) {
      body += '<tr><td style="text-align:center">' + (pi * ITEMS_PER_PAGE + ii + 1) + '</td>' +
        '<td>' + esc_(it.prodName) + '</td><td style="text-align:center">' + it.qty + ' ' + esc_(it.unit) + '</td>' +
        '<td style="text-align:right">' + thaiMoney_(it.unitPrice) + '</td>' +
        '<td style="text-align:right">' + thaiMoney_(it.totalPrice) + '</td></tr>';
    });
    body += '</table>';

    if (isLast) {
      body += '<table style="width:100%;margin-top:10px;font-size:12px"><tr>' +
        '<td style="vertical-align:top;width:55%"><b>หมายเหตุ:</b> ' + esc_(doc.notes || '-') +
        (doc.docType === 'QT' ? '<br>กำหนดส่งมอบ: ' + doc.sendDays + ' วัน | กำหนดยืนราคา: ' + doc.confirmDays + ' วัน' : '') + '</td>' +
        '<td style="vertical-align:top;text-align:right">' +
        'ยอดรวมก่อนภาษี: ' + thaiMoney_(doc.subTotal) + ' บาท<br>' +
        'ภาษีมูลค่าเพิ่ม: ' + thaiMoney_(doc.vatAmount) + ' บาท<br>' +
        '<b style="font-size:14px">รวมทั้งสิ้น: ' + thaiMoney_(doc.grandTotal) + ' บาท</b></td></tr></table>';
    }

    body += '<div style="margin-top:40px;font-size:12px;text-align:center">';
    if (doc.docType === 'QT') {
      body += '<span style="display:inline-block;width:220px">ในนาม ' + esc_(st.shopName) +
        '<br><br><br>________________________________' +
        '<br><b>(' + esc_(st.signatoryName || st.managerName) + ')</b><br>ผู้จัดการ / ผู้มีอำนาจลงนาม</span>';
    } else {
      body += '<table style="width:100%;font-size:12px"><tr>' +
        '<td style="text-align:center;width:33%">ได้รับสินค้าถูกต้องตามรายการ<br><br><br>_____________________<br><b>ผู้รับของ</b></td>' +
        '<td style="text-align:center;width:33%">ส่งมอบสินค้าเรียบร้อยแล้ว<br><br><br>_____________________<br><b>ผู้ส่งของ</b></td>' +
        '<td style="text-align:center;width:33%">ในนาม ' + esc_(st.shopName) + '<br><br><br>_____________________<br><b>(' + esc_(st.signatoryName || st.managerName) + ')</b><br>ผู้จัดการ / ผู้มีอำนาจลงนาม</td>' +
        '</tr></table>';
    }
    body += '</div></div>';
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="font-family:sarabun,th sarabun new,arial,sans-serif;color:#111">' + body + '</body></html>';
}

/* ============================================================
 * SECTION 9: REPORTS & DASHBOARD
 * ============================================================ */
function apiGetDashboard(token) {
  requireAuth_(token);
  const docs = rawDocuments_();
  const products = listProducts_();
  const now = new Date();
  const year = now.getFullYear();
  const thisMonth = now.getMonth();

  const monthlySales = Array(12).fill(0);
  let salesThisMonth = 0, salesPrevMonth = 0;
  let pendingQT = { count: 0, value: 0 };
  let pendingIV = { count: 0, value: 0 };
  let completedCount = 0, cancelledCount = 0;

  docs.forEach(function(d) {
    const valid = d.status !== 'Cancelled';
    const dt = new Date(d.docDate + 'T00:00:00');
    const isSale = d.docType !== 'QT';
    if (isSale && valid && dt.getFullYear() === year) {
      monthlySales[dt.getMonth()] += d.grandTotal;
      if (dt.getMonth() === thisMonth) salesThisMonth += d.grandTotal;
      else if (dt.getMonth() === thisMonth - 1) salesPrevMonth += d.grandTotal;
    }
    if (valid && d.docType === 'QT' && d.status === 'Active') { pendingQT.count++; pendingQT.value += d.grandTotal; }
    if (valid && d.docType === 'IV' && d.status === 'Active') { pendingIV.count++; pendingIV.value += d.grandTotal; }
    if (valid && (d.status === 'Billed' || d.status === 'Paid' || (isSale))) completedCount++;
    if (d.status === 'Cancelled') cancelledCount++;
  });

  const lowStock = products.filter(function(p) { return p.remainingQty <= p.reorderPoint && p.status !== 'ระงับ'; });

  return {
    ok: true,
    kpi: {
      salesThisMonth: salesThisMonth,
      salesPrevMonth: salesPrevMonth,
      pendingQT: pendingQT,
      pendingIV: pendingIV,
      lowStockCount: lowStock.length
    },
    monthlySales: monthlySales,
    docStatus: {
      qtActive: pendingQT.count,
      ivPending: pendingIV.count,
      completed: completedCount,
      cancelled: cancelledCount
    },
    lowStock: lowStock
  };
}

function apiGetReport(token, from, to) {
  requireAuth_(token);
  const docs = rawDocuments_().filter(function(d) {
    if (d.status === 'Cancelled') return false;
    if (from && d.docDate < String(from)) return false;
    if (to && d.docDate > String(to)) return false;
    return true;
  });

  const byType = {};
  const prodAgg = {};
  let totalGrand = 0;

  docs.forEach(function(d) {
    if (!byType[d.docType]) byType[d.docType] = { count: 0, subtotal: 0, vat: 0, grand: 0 };
    byType[d.docType].count++;
    byType[d.docType].subtotal += d.subTotal;
    byType[d.docType].vat += d.vatAmount;
    byType[d.docType].grand += d.grandTotal;
    totalGrand += d.grandTotal;
    itemsOf_(d.docNo).forEach(function(it) {
      const key = it.prodId + '|' + it.prodName;
      if (!prodAgg[key]) prodAgg[key] = { prodId: it.prodId, prodName: it.prodName, qty: 0, amount: 0 };
      prodAgg[key].qty += it.qty;
      prodAgg[key].amount += it.totalPrice;
    });
  });

  const topProducts = Object.keys(prodAgg).map(function(k) { return prodAgg[k]; })
    .sort(function(a, b) { return b.amount - a.amount; }).slice(0, 10);

  return { ok: true, from: from || '-', to: to || '-', byType: byType, topProducts: topProducts, totalGrand: totalGrand, count: docs.length };
}
