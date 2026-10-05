# MYSUPP – Storytelling Dashboard

Dashboard เชิงเล่าเรื่องสำหรับวิชา **Business Idea Creation** (กลุ่มอุตสาหกรรมผลิตภัณฑ์เสริมอาหาร)
MYSUPP คือแพลตฟอร์มวางแผนอาหารเสริมเฉพาะบุคคลด้วย AI จากผลตรวจสุขภาพจริง · Your Health. Personalized.

> ข้อมูลทั้งหมดเป็นข้อมูลจำลอง (synthetic) เพื่อการเรียนการสอน ไม่ใช่ตัวเลขจริงของธุรกิจ

**เปิดดูบนเว็บ:** https://67160358petnapakorn.github.io/mysupp-dashboard/

## ไฟล์หลัก (ส่งงาน)
| ไฟล์ | คืออะไร |
|---|---|
| `index.html` | หน้า Dashboard (ต้องต่ออินเทอร์เน็ตเพื่อโหลด Chart.js และฟอนต์) |
| `data.xlsx` | ชุดข้อมูลจำลองแบบ Star Schema (README, Summary ที่ใช้สูตร, fact และ dim) |

## ฟีเจอร์ของหน้าเว็บ
- ตัวกรองช่วงเวลา (12 / 6 / 3 เดือนล่าสุด) และช่องทางขาย 5 ช่องทาง ทุกกราฟและข้อความเปลี่ยนตาม
- การ์ด KPI และส่วนสัญญาณเตือน 6–7 เรื่อง (ปกติ / เฝ้าระวัง / วิกฤต) พร้อมเกณฑ์และสิ่งที่ควรทำ
- 6 บทเล่าเรื่อง: ลูกค้า → AI ทำงานอย่างไร → ผลลัพธ์ → ธุรกิจ → การเติบโต → ก้าวต่อไป
- ปุ่ม “ดูเป็นตาราง” ใต้ทุกกราฟ, โหมดสว่าง/มืด, พิมพ์เป็น PDF ได้

## โครงสร้าง
```
index.html          หน้า Dashboard
data.xlsx           ข้อมูล (Star Schema)
src/generate_data.py  สร้าง data.xlsx และ src/embed.json (random seed คงที่)
src/template.html     ต้นฉบับหน้าเว็บ
src/build.py          ประกอบ index.html จาก template + embed.json
.nojekyll
```
แก้ตัวเลข: แก้ที่ `src/generate_data.py` แล้วรัน `python src/generate_data.py` ตามด้วย `python src/build.py`

## ผู้จัดทำ
| รหัสนักศึกษา | ชื่อ |
|---|---|
| 67160358 | เพชรนภากร พลนิกร |
