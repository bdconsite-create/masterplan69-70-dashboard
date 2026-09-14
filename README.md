# Dashboard Master Plan 69-70 — v15

ตำแหน่งใช้งานล่าสุด

- Excel ต้นทาง: `D:\MCP-Test\MasterPlan69-70\masterplan ซ่อมทำ ปี69.xlsx`
- Dashboard: `D:\MCP-Test\MasterPlan69-70\Dashboard_Source_v15\Dashboard_Source_v15`

## Workflow ระดับแรก (แนะนำ)

1. แก้ไขไฟล์ Excel ตามปกติ
2. กด Save ให้เรียบร้อย
3. ดับเบิลคลิก `Update Dashboard.bat`
4. Script จะคัดลอก Excel ล่าสุดมาเป็น `masterplan.xlsx` ในโฟลเดอร์ Dashboard
5. เมื่อเปิด Dashboard ผ่าน Live Server หรือ GitHub Pages หน้าเว็บจะอ่าน `masterplan.xlsx` อัตโนมัติ แล้วอัปเดต Task, Gantt และรูปจาก Excel
6. ถ้าเครื่องมี Git และโฟลเดอร์ Dashboard เชื่อม GitHub แล้ว Script จะ `git add` → `commit` → `push` ให้อัตโนมัติ
7. ถ้ายังไม่ได้ตั้ง Git/GitHub ระบบจะอัปเดต Dashboard ในเครื่องให้ก่อน โดยข้ามขั้นตอน push

## ไฟล์สำคัญ

- `index.html` — หน้าตา Dashboard
- `app.js` — logic, filter, Gantt และตัวอ่าน Excel
- `masterplan.xlsx` — สำเนา Excel ล่าสุดที่หน้า Dashboard โหลดอัตโนมัติ
- `data.json` — fallback ถ้า `masterplan.xlsx` โหลดไม่ได้
- `photos/` — fallback รูปตั้งต้นเดิม
- `tools/update_from_excel.py` — utility เดิมสำหรับ build static data; workflow ระดับแรกปัจจุบันไม่จำเป็นต้องใช้ Python
- `Update Dashboard.bat` — ปุ่มอัปเดตแบบ one-click
- `Update-Dashboard.ps1` — workflow เบื้องหลัง

## หมายเหตุ

- Workflow นี้ไม่ต้องใช้ OneDrive API, Power Automate หรือ Python
- ไฟล์ Excel ล่าสุดเป็น source of truth
- Dashboard v15 รองรับ `mater69-70`, `Daily Photos` และ `รูปภาพรายละเอียดงาน`
- `masterplan.xlsx` มีขนาดค่อนข้างใหญ่ ดังนั้นครั้งแรกที่เปิด Dashboard ออนไลน์อาจใช้เวลารอสักครู่
