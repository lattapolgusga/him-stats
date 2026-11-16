# HIM Stats — ระบบเก็บสถิติคริสตจักร H.I.M. Thailand

เว็บแอป (LINE LIFF) + Google Sheet + Apps Script + LINE OA "งานพันธกิจ H.I.M."

> repo นี้มีแค่โค้ด ไม่มีข้อมูลคริสตจักร ข้อมูลทั้งหมดอยู่ใน Google Sheet กลาง

## ติดตั้ง
1. **Google Sheet กลาง:** อัปโหลด `HIM-ทะเบียนคริสตจักร-ปรับปรุง.xlsx` ขึ้น Google Drive แล้วเปิดเป็น Google Sheet
2. **Apps Script:** ใน Sheet กด Extensions → Apps Script วางโค้ดจาก `apps-script/Code.gs`
   - Project Settings → Script properties ตั้ง `LOGIN_CHANNEL_ID`, `LINE_ACCESS_TOKEN`, `CURRENT_ROUND`
   - เลือกฟังก์ชัน `setup` แล้วกด Run และอนุญาตสิทธิ์
   - Deploy → New deployment → Web app → Execute as: **Me**, Who has access: **Anyone** → คัดลอก URL
3. **LIFF:** ไปที่ LINE Developers → HIM Stats App → LIFF → Add
   - Size: Full · Endpoint URL: `https://lattapolgusga.github.io/him-stats/` · Scope: `openid`, `profile`
   - คัดลอก LIFF ID
4. ใส่ LIFF ID และ API URL ลงใน `config.js`
5. **Webhook ของ OA:** ตั้งเป็น URL ของ Apps Script แล้วเปิด Use webhook และปิดข้อความตอบกลับอัตโนมัติ
6. **แอดมินคนแรก:** เปิดเว็บแอปแล้วส่งคำขอด้วยรหัสคริสตจักรใดก็ได้ จากนั้นแก้แท็บ "ผู้ใช้" ใน Sheet ให้บทบาท = `admin` และสถานะ = `active`
7. เปลี่ยนแชนแนล HIM Stats App จาก Developing เป็น **Published**
8. **Rich menu ใน OA:** ปุ่ม "อัปเดทสถิติ" ลิงก์ไปที่ `https://liff.line.me/<LIFF ID>`
