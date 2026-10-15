# 🤖 AI Face Recognition & Voice Greeting Web App (Google Apps Script)

ระบบเว็บแอปพลิเคชันตรวจจับและจดจำใบหน้าผ่านกล้อง Webcam พร้อมระบบสังเคราะห์เสียงพูดทักทายภาษาไทยอัตโนมัติ พัฒนาบน **Google Apps Script**, **Google Sheets**, **Google Drive** และขับเคลื่อนด้วยโมเดลสมองกล **face-api.js (TensorFlow.js)**

---

## 🔗 ลิงก์ระบบที่พร้อมใช้งานทันที (Live Deployment Links)

| ระบบ / แหล่งข้อมูล | ลิงก์เข้าใช้งาน | คำอธิบาย |
| :--- | :--- | :--- |
| 🌐 **Live Web Application** | [คลิกเปิด Web App ที่นี่](https://script.google.com/macros/s/AKfycbzOO-mO-bMSzj_b_SPnwMLBiR7GIXViXFJmDVV_InuAAdMf6DmghpLEDXwYRf2inRTWZQ/exec) | หน้าเว็บหลักสำหรับสแกนใบหน้าและทักทายด้วยเสียง |
| 📊 **Google Spreadsheet (Database)** | [เปิดตาราง Google Sheets](https://docs.google.com/spreadsheets/d/15iYZLwqkoWcw4yPBOy9TysfaKd0MfrmZR_JgDKl_VEE/edit) | ตารางจัดเก็บรายชื่อ Descriptors และประวัติการสแกน |
| ⚙️ **Apps Script Project Editor** | [เปิด Script Editor](https://script.google.com/d/1YlUco-bkHMzhvU7pw0Ux4z_8wgCDADu0nP8iEVL1N6nB_7L1fO0dRKtr/edit) | หน้าจัดการโค้ดบน Google Cloud Platform |

> [!NOTE]
> เมื่อเปิด Web App ครั้งแรก Google อาจขอให้ท่านกดยืนยันสิทธิ์การเข้าถึงข้อมูล (Authorization) ให้เลือกบัญชี Google ของท่าน ➔ กด **Advanced (ขั้นสูง)** ➔ กด **Go to ... (unsafe)** เพื่อเปิดใช้งาน

---

## ✨ ฟังก์ชันเด่นของระบบ

1. **ตรวจจับและจดจำใบหน้าแบบเรียลไทม์ (Real-time Face Recognition)**:
   - สตรีมภาพผ่านกล้อง Webcam ความเร็วสูง 15-30 fps
   - คำนวณเวกเตอร์ใบหน้า 128 มิติ (128-d Descriptors) ด้วย GPU/WebGL บนฝั่งเบราว์เซอร์ผู้ใช้งาน ไม่กินโควตาฝั่งเซิร์ฟเวอร์
   - ตีกรอบสีเขียวระบุชื่อผู้ใช้เมื่อพบข้อมูลตรงกัน หรือตีกรอบสีส้มระบุ "ลูกค้าทั่วไป" เมื่อไม่พบข้อมูล

2. **ระบบสังเคราะห์เสียงทักทายภาษาไทย (Thai Text-to-Speech)**:
   - **กรณีพบข้อมูลใบหน้าตรงกัน**: ออกเสียงทักทายชื่อผู้ใช้ทันที เช่น *"สวัสดีค่ะคุณ [ชื่อ]"*
   - **กรณีเป็นบุคคลแปลกหน้า (Unknown)**: ออกเสียงทักทาย *"สวัสดีค่ะคุณลูกค้า วันนี้รับอะไรดีค่ะ"*
   - รองรับเสียงพูดภาษาไทยแท้ (`th-TH`) ผ่าน Web Speech API มาตรฐาน

3. **กลไกป้องกันการทักทายซ้ำ (Greet Once Until Leaving Frame)**:
   - ระบบจะทักทายเพียง **1 ครั้ง** ตลอดช่วงเวลาที่บุคคลนั้นยืนอยู่หน้ากล้อง
   - เมื่อบุคคลนั้น **เดินพ้นรัศมีกล้องเกิน 3 วินาที** ระบบจะตัดสถานะ (Exit Detection) เพื่อให้พร้อมสำหรับการทักทายครั้งถัดไปเมื่อเดินกลับเข้ามาใหม่

4. **ระบบลงทะเบียนใบหน้า (Face Registration)**:
   - ถ่ายภาพใบหน้าหน้าตรงจากกล้อง สกัดเวกเตอร์ข้อมูลคณิตศาสตร์ และบันทึกเข้าสู่ระบบ
   - บันทึกชื่อและเวกเตอร์ลงใน **Google Sheets**
   - อัปโหลดไฟล์รูปภาพ Snapshot ไปเก็บในโฟลเดอร์ `FaceRecognition_Photos` บน **Google Drive** อัตโนมัติ

5. **ระบบจัดการรายชื่อ (Profiles Management)**:
   - ดูรายชื่อบุคคลที่ลงทะเบียนไว้ พร้อมรูปภาพพรีวิวและวันที่ลงทะเบียน
   - สามารถกดลบข้อมูลใบหน้าและไฟล์รูปภาพใน Google Drive ออกได้ทันที

---

## 📁 โครงสร้างไฟล์ในโปรเจกต์

| ชื่อไฟล์ | ประเภท | หน้าที่ |
| :--- | :--- | :--- |
| `Code.js` | Server-side (GAS) | จัดการ Google Sheets, Google Drive, สร้างโฟลเดอร์, และ API ฟังก์ชัน |
| `Index.html` | Frontend Template | โครงสร้างหน้าเว็บหลัก (Tailwind CSS, Glassmorphism, Video & Canvas Viewport) |
| `JavaScript.html` | Client-side Script | จัดการโมเดล `face-api.js`, กล้อง Webcam, เสียงพูดภาษาไทย และ State Machine |
| `Stylesheet.html` | Stylesheet | แอนิเมชันเส้นสแกน (Scan Line), กรอบไฮไลต์มุม, และคลาสตกแต่ง UI |
| `appsscript.json` | Manifest | ตั้งค่า Runtime V8, โซนเวลา Asia/Bangkok, และสิทธิ์ OAuth Scopes |
| `.clasp.json` | Configuration | การตั้งค่าเชื่อมต่อกับโปรเจกต์ Google Apps Script |
| `test-local.html` | Standalone Test | ไฟล์สำหรับเปิดทดสอบกล้อง, เสียงพูด, และการตรวจจับใบหน้าในเครื่องทันที |

---

## 💻 การทดสอบแบบด่วนบนเครื่องของคุณ (Local Preview)

หากต้องการทดสอบการเปิดกล้อง Webcam, ทดสอบการอ่านออกเสียงภาษาไทย และทดลองสแกนใบหน้าบนเครื่องทันทีโดยไม่ต้องต่ออินเทอร์เน็ต:
1. ดับเบิลคลิกเปิดไฟล์ `test-local.html` ด้วย Google Chrome หรือ Microsoft Edge
2. กดอนุญาตให้เบราว์เซอร์เข้าถึงกล้อง Webcam
3. ทดลองกดแท็บ **"ลงทะเบียนใบหน้า"** เพื่อถ่ายรูปตนเองและบันทึกชื่อ
4. กลับมาที่แท็บ **"ตรวจจับและทักทาย"** ระบบจะตีกรอบสีเขียวและเปล่งเสียง: *"สวัสดีค่ะคุณ [ชื่อของคุณ]"* ทันที!
