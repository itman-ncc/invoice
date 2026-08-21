# 🧾 Smart Billing System

ระบบเอกสารการขายออนไลน์ (ใบเสนอราคา / ใบวางบิล-ใบส่งของ / ใบเสร็จ / บิลเงินสด) + รายงานสถิติ
สถาปัตยกรรม: **Google Apps Script WebApp + Google Sheets (ฐานข้อมูล) + Tailwind CSS + SweetAlert2 + Chart.js**

---

## 🔗 ลิงก์สำคัญ

| รายการ | URL |
|---|---|
| Google Sheet (ฐานข้อมูล) | https://docs.google.com/spreadsheets/d/1Nf7ud48rwtGbEmOecxXhh3YyraC_kteEXFh4Ec-8NP8/copy |
| Apps Script Editor | - |
| Web App | ดูไฟล์ `.deployment-id` → `https://script.google.com/macros/s/<id>/exec` |

**Login เริ่มต้น:** `admin` / `1234` (Admin) หรือ `staff` / `1234` (Staff)

---

## 🚀 ติดตั้งครั้งแรก (3 ขั้นตอน)

1. **เปิด Apps Script API** — ไปที่ https://script.google.com/home/usersettings แล้วเปิด toggle "Google Apps Script API" (ทำครั้งเดียว)
2. **อัปโหลดโค้ด** — ในโฟลเดอร์นี้รัน:
   ```
   npm run push
   ```
3. **สร้างฐานข้อมูล** — เปิด Apps Script Editor → เลือกฟังก์ชัน `setupSystem` → ▶ Run → Authorize (Allow ทั้งหมด)
   ชีตทั้ง 9 (Users, Customers, Products, Settings, Documents, DocumentItems, Counters, AuditLogs, Sessions) จะถูกสร้างพร้อมข้อมูลตัวอย่าง

## 🌐 Deploy เป็น Web App

```
npm run deploy
```
- ครั้งแรก: สร้าง deployment ใหม่ + บันทึก id ลง `.deployment-id`
- ครั้งถัดไป: redeploy deployment เดิม → **URL /exec ไม่เปลี่ยน**

> ⚠️ หลัง deploy ครั้งแรก ต้องเปิด URL ผ่านบัญชีเจ้าของ 1 ครั้งเพื่อกด "Review permissions → Allow" ก่อนแชร์ให้คนอื่นใช้

---

## 🔄 Workflow การพัฒนา (Auto Pull/Push/Deploy)

| คำสั่ง | ทำอะไร |
|---|---|
| `npm run watch` | **เฝ้าไฟล์ตลอดเวลา** — บันทึกไฟล์เมื่อไหร่ push ขึ้น Apps Script ทันที |
| `npm run push` | Push ครั้งเดียว |
| `npm run pull` | ดึงโค้ดจาก Editor กลับเครื่อง |
| `npm run deploy` | Push + Redeploy (URL คงเดิม) |
| `npm run open` | เปิด Web App URL ในเบราว์เซอร์ |
| `npm run editor` | เปิด Apps Script Editor |
| `npm run logs` | ดู console.log ฝั่ง server แบบเรียลไทม์ |

**แนะนำ:** เปิด terminal ค้างไว้กับ `npm run watch` แล้วแก้โค้ดใน VS Code ได้เลย

---

## ✨ ฟีเจอร์ระบบ

### เอกสาร
- 4 ประเภท: QT (ใบเสนอราคา), IV (ใบวางบิล/ใบส่งของ), RE (ใบเสร็จ+ใบกำกับภาษี), RC (บิลเงินสด)
- เลขที่เอกสารรันอัตโนมัติแยกตามประเภท + ปีปฏิทิน (`QT26-0002`) ใช้ LockService กันเลขซ้ำ
- เลือกได้: **ลงวันที่ / ไม่ลงวันที่** (DB เก็บวันที่จริงเสมอ)
- แปลง QT → IV ได้ (QT เปลี่ยนสถานะ Billed, IV เป็นเอกสารใหม่อ้างอิงกัน)
- แก้ไขได้เฉพาะสถานะ Active / ยกเลิกได้ **พร้อมเหตุผลบังคับกรอก**
- พิมพ์ A4: >20 รายการแบ่งหน้าอัตโนมัติ + เลขหน้า + ยอดรวมหน้าสุดท้าย
- ลายเซ็น: QT = 1 ช่อง (ผู้มีอำนาจ) | IV = 3 ช่อง (ผู้รับ/ผู้ส่ง/ผู้จัดการ)
- Export PDF → บันทึกใน Drive โฟลเดอร์ `SmartBilling_PDF`
- พิมพ์/PDF ได้ไม่จำกัด — ทุกครั้งบันทึก Audit Log

### ข้อมูลหลัก
- สินค้า: CRUD + สถานะ (ปกติ/จุดสั่งซื้อ/หมด/ระงับ) + Stock Alert
- เพิ่มสินค้าใหม่ทันทีจากหน้าออกเอกสาร → ลงชีต Products อัตโนมัติ
- ลูกค้า: CRUD + เลขภาษี
- ผู้ใช้: Admin/Staff, Active on/off, รหัสผ่าน SHA-256 (เฉพาะ Admin จัดการได้)

### รายงาน (ต้อง Login)
- KPI: ยอดขายเดือนนี้ (+เทียบเดือนก่อน), QT รอวางบิล, IV รอเก็บเงิน, สินค้าใกล้หมด
- กราฟ Chart.js: ยอดขายรายเดือน + สัดส่วนสถานะเอกสาร
- รายงานช่วงวันที่ + Top 10 สินค้าขายดี

---

## 🗄️ โครงสร้างไฟล์

```
D:\invoice\
├── Code.gs           Backend ทั้งหมด (API + setupSystem)
├── Index.html        โครงหน้าเว็บ
├── Stylesheet.html   CSS + A4 print engine
├── JavaScript.html   Client logic
├── appsscript.json   Manifest (timezone Bangkok, scopes)
├── .clasp.json       เชื่อม clasp ↔ script project
├── scripts/deploy.js Auto deploy (push + redeploy)
└── backup/           ไฟล์เดิมก่อนติดตั้ง
```

## 🔧 แก้ปัญหาที่พบบ่อย

| ปัญหา | วิธีแก้ |
|---|---|
| `clasp push` error 401/403 | เปิด Apps Script API ที่ script.google.com/home/usersettings แล้ว `npx clasp login` ใหม่ |
| เปิด Web App แล้วขาว/Permission denied | เจ้าของต้องเปิด URL แล้วกด Review permissions → Allow 1 ครั้ง |
| ชีตไม่ถูกสร้าง | รัน `setupSystem()` จาก Editor (Authorize ด้วย) |
| แก้โค้ดใน Editor แล้วทับ local | รัน `npm run pull` ก่อน push เสมอ ถ้าแก้สองที่ |
