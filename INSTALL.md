# ขั้นตอนการติดตั้ง GAS Promotion Merger

## สิ่งที่ต้องเตรียม

- Google Account ปลายทางที่จะติดตั้ง
- Gemini API Key ([ขอได้ที่ Google AI Studio](https://aistudio.google.com/app/apikey))
- ไฟล์โปรเจกต์: `Index.html` และ `รหัส.gs`

---

## ขั้นตอนที่ 1 — สร้างโฟลเดอร์ใน Google Drive

ล็อกอิน Google Drive แล้วสร้างโฟลเดอร์ **4 โฟลเดอร์** ดังนี้:

| ชื่อโฟลเดอร์ (ตั้งเองได้) | หน้าที่ |
|---|---|
| `Main` | โฟลเดอร์หลัก / เก็บไฟล์ผลลัพธ์ที่ merge แล้ว |
| `Target` | เก็บไฟล์ Google Sheets ต้นฉบับโปรโมชั่น |
| `Update` | เก็บไฟล์ที่ต้องการ update เข้าไฟล์หลัก |
| `Profiles` | เก็บไฟล์ Mapping Profile (สร้างโดยอัตโนมัติ แต่ต้องสร้างโฟลเดอร์ไว้ก่อน) |

**วิธีดู Folder ID:**
เปิดโฟลเดอร์ใน Drive → URL จะเป็น `https://drive.google.com/drive/folders/`**`FOLDER_ID_HERE`**
คัดลอก ID ส่วนนั้นมาใช้ในขั้นตอนถัดไป

---

## ขั้นตอนที่ 2 — สร้างโปรเจกต์ Google Apps Script

1. ไปที่ [script.google.com](https://script.google.com) → คลิก **"New project"**
2. ตั้งชื่อโปรเจกต์ (เช่น `Promotion Merger`)
3. ลบโค้ดเริ่มต้นที่มีอยู่ออกทั้งหมด

---

## ขั้นตอนที่ 3 — คัดลอกโค้ด

### 3.1 ไฟล์ `รหัส.gs` (Code.gs)

1. คลิกที่ไฟล์ `Code.gs` ในแถบซ้าย
2. วางเนื้อหาจากไฟล์ `รหัส.gs` ทั้งหมด
3. **แก้ไข Folder ID 4 ตัว** ที่บรรทัดต้นไฟล์:

```javascript
const MAIN_FOLDER_ID   = "วาง-ID-โฟลเดอร์-Main-ที่นี่";
const TARGET_FOLDER_ID = "วาง-ID-โฟลเดอร์-Target-ที่นี่";
const UPDATE_FOLDER_ID = "วาง-ID-โฟลเดอร์-Update-ที่นี่";
```

และค้นหาบรรทัด `PROFILE_FOLDER_ID_` แล้วแก้เป็น:

```javascript
var PROFILE_FOLDER_ID_ = "วาง-ID-โฟลเดอร์-Profiles-ที่นี่";
```

### 3.2 ไฟล์ `Index.html`

1. คลิก **"+"** ข้างแถบไฟล์ด้านซ้าย → เลือก **"HTML"**
2. ตั้งชื่อว่า `Index` (ตัว I พิมพ์ใหญ่ ไม่มีนามสกุล .html)
3. วางเนื้อหาจากไฟล์ `Index.html` ทั้งหมด

---

## ขั้นตอนที่ 4 — เปิด Advanced Services

1. ในหน้า Apps Script → คลิก **"Services"** (ไอคอน + ข้างซ้าย)
2. เพิ่ม **Google Sheets API** → คลิก Add
3. เพิ่ม **Google Drive API** → คลิก Add

> ทั้งสองตัวนี้จำเป็นสำหรับการตรวจ strikethrough และการ sync ไฟล์ Excel

---

## ขั้นตอนที่ 5 — ตั้งค่า Script Properties

1. ใน Apps Script → คลิก **"Project Settings"** (ไอคอนฟันเฟือง ด้านซ้ายล่าง)
2. เลื่อนลงหา **"Script properties"** → คลิก **"Add script property"**
3. เพิ่ม 3 ค่าต่อไปนี้:

| Property Name | Value |
|---|---|
| `APP_USERNAME` | ชื่อผู้ใช้สำหรับ login (เช่น `admin`) |
| `APP_PASSWORD` | รหัสผ่าน (ตั้งให้ปลอดภัย) |
| `GEMINI_API_KEY` | API Key จาก Google AI Studio |

4. คลิก **"Save script properties"**

---

## ขั้นตอนที่ 6 — Deploy เป็น Web App

1. คลิก **"Deploy"** (มุมขวาบน) → **"New deployment"**
2. คลิกไอคอนฟันเฟืองข้าง "Select type" → เลือก **"Web app"**
3. ตั้งค่าดังนี้:

| ช่อง | ค่า |
|---|---|
| Description | `v1` (หรือใส่อะไรก็ได้) |
| Execute as | **Me** (บัญชีของคุณ) |
| Who has access | **Anyone** หรือ **Anyone within [องค์กร]** |

4. คลิก **"Deploy"**
5. คลิก **"Authorize access"** → เลือก Google Account → อนุญาตสิทธิ์ทั้งหมด
6. คัดลอก **Web app URL** เก็บไว้ใช้เข้าแอป

---

## ขั้นตอนที่ 7 — ทดสอบ

1. เปิด Web app URL ที่ได้จากขั้นตอนที่ 6
2. Login ด้วย username/password ที่ตั้งไว้ใน Script Properties
3. ทดสอบเลือกไฟล์จาก Drive → ถ้าเห็นรายการไฟล์ แสดงว่าติดตั้งสำเร็จ

---

## การอัปเดตโค้ดในอนาคต

เมื่อแก้ไขโค้ดแล้วต้องการ deploy ใหม่:
1. คลิก **"Deploy"** → **"Manage deployments"**
2. คลิกไอคอนดินสอ (Edit) ข้าง deployment ที่มีอยู่
3. เปลี่ยน Version เป็น **"New version"**
4. คลิก **"Deploy"**

> ⚠️ URL ของ Web App จะไม่เปลี่ยน เมื่อใช้วิธีนี้

---

## แก้ปัญหาที่พบบ่อย

| ปัญหา | สาเหตุ / วิธีแก้ |
|---|---|
| ไม่เห็นไฟล์ใน Drive | Folder ID ผิด หรือ โฟลเดอร์ไม่ได้แชร์กับบัญชีที่รัน script |
| AI ไม่ตอบสนอง | `GEMINI_API_KEY` ไม่ถูกต้อง หรือ quota หมด |
| Strikethrough ไม่ทำงาน | ยังไม่ได้เปิด Google Sheets API ใน Advanced Services |
| Sync Excel ไม่ได้ | ยังไม่ได้เปิด Google Drive API ใน Advanced Services |
| Login ไม่ได้ | ตรวจสอบ Script Properties ว่า `APP_USERNAME` / `APP_PASSWORD` ถูกต้อง |
| "ไม่มีสิทธิ์เข้าถึงไฟล์" | ไฟล์ต้นฉบับอยู่ผิดโฟลเดอร์ ต้องอยู่ใน Main / Target / Update เท่านั้น |
