# OSRAM Sales Pipeline (Supabase edition)

ระบบกรอก Sale Pipeline พร้อมระบบ login, อนุมัติบัญชีโดย admin, dashboard รายงาน และ
นำเข้า/ส่งออกข้อมูล — เวอร์ชันนี้เก็บข้อมูลจริงบน **Supabase (Postgres)** แทน
`window.storage` ของ Claude Artifact เดิม จึงสามารถ deploy เป็นเว็บแอปแยกต่างหาก
(เช่น Vercel, Netlify) ได้

## โครงสร้างโปรเจกต์

```
osram-supabase/
├── index.html
├── package.json
├── vite.config.js
├── .env.example
├── supabase/
│   └── schema.sql        ← รันไฟล์นี้ใน Supabase SQL editor ครั้งเดียว
└── src/
    ├── main.jsx
    ├── App.jsx            ← UI + business logic ทั้งหมด
    ├── db.js              ← data-access layer (แปลง camelCase ↔ Supabase columns)
    └── supabaseClient.js  ← สร้าง Supabase client จาก env vars
```

## ขั้นตอนติดตั้ง

### 1. สร้างโปรเจกต์ Supabase
1. ไปที่ [supabase.com](https://supabase.com) → สร้างโปรเจกต์ใหม่ (ฟรีได้)
2. เข้า **SQL Editor** → New query → วางเนื้อหาทั้งหมดของ `supabase/schema.sql` → Run
   - สร้างตาราง `users`, `pipeline_entries`, `zones`
   - ใส่ข้อมูลโซนเริ่มต้น และบัญชีทดลอง `admin` / `admin123`, `demo` / `demo123`
   - เปิด Row Level Security พร้อม policy ที่จำเป็น (ดูหัวข้อ "เรื่องความปลอดภัย" ด้านล่าง)
3. ไปที่ **Project Settings → API** → คัดลอก `Project URL` และ `anon public` key

### 2. ตั้งค่าโปรเจกต์
```bash
cp .env.example .env
# แก้ .env ใส่ VITE_SUPABASE_URL และ VITE_SUPABASE_ANON_KEY ที่คัดลอกมา

npm install
npm run dev
```
เปิด `http://localhost:5173` แล้วล็อกอินด้วยบัญชีทดลอง `admin` / `admin123`

### 3. Build สำหรับ deploy
```bash
npm run build     # ได้โฟลเดอร์ dist/
npm run preview   # พรีวิว build ก่อน deploy จริง
```
นำ `dist/` ไป deploy บน Vercel, Netlify, Cloudflare Pages ฯลฯ ได้เลย (เป็น static site)
อย่าลืมตั้งค่า Environment Variables `VITE_SUPABASE_URL` และ `VITE_SUPABASE_ANON_KEY`
บนแพลตฟอร์มที่ deploy ด้วย (ไม่ใช่แค่ไฟล์ `.env` ในเครื่อง)

## ฟีเจอร์หลัก
- Login / สมัครสมาชิก พร้อม approve/reject บัญชีโดย admin
- ผู้สมัครคนแรกของระบบ (หรือถ้า table ว่างตอนแอปโหลดครั้งแรก) จะได้สิทธิ์ admin อัตโนมัติ
- กรอก/แก้ไข/ลบรายการ Pipeline พร้อมฟิลด์ครบตามไฟล์ Excel ต้นฉบับ
- เพิ่มโซนใหม่ได้จากหน้าสมัครสมาชิกหรือฟอร์มกรอกข้อมูล
- Dashboard รายงาน: กราฟตามเกรด/โซน/ความคืบหน้า, leaderboard พนักงานขาย
- Import/Export CSV และสำรอง-กู้คืนข้อมูลทั้งระบบเป็น JSON

## เรื่องความปลอดภัย (สำคัญ ควรอ่านก่อนใช้งานจริง)

แอปนี้ทำระบบ login ของตัวเอง (ไม่ได้ใช้ Supabase Auth) เพื่อให้ workflow "admin
ต้องอนุมัติบัญชีใหม่" ทำงานได้ตามที่ออกแบบไว้ ผลคือ:

1. **รหัสผ่านเก็บเป็น plain text** ในตาราง `users` — ไม่เหมาะกับข้อมูลจริงที่มีความสำคัญ
   ถ้าจะใช้งานจริง ควรเพิ่มการ hash รหัสผ่าน (เช่นผ่าน Edge Function) หรือย้ายไปใช้
   Supabase Auth แทน
2. **RLS policy เปิดกว้าง (`using (true)`)** เพราะแอปเรียก Supabase ด้วย anon key
   ตรงจาก browser โดยไม่มี Supabase session ดังนั้นใครก็ตามที่มี anon key ของโปรเจกต์
   (ซึ่งฝังอยู่ใน JS bundle ที่ deploy) จะ**อ่าน/เขียนข้อมูลทุกแถวได้** ไม่ต่างจาก
   เวอร์ชัน Claude Artifact เดิมที่ข้อมูลเป็น shared storage อยู่แล้ว แต่ตอนนี้เป็น
   ฐานข้อมูลจริงที่เข้าถึงได้จากอินเทอร์เน็ต จึงควรระวังมากขึ้น
3. สำหรับการใช้งานภายในทีมเล็ก ๆ หรือ demo เรื่องนี้ยอมรับได้ แต่ถ้าจะขยายไปใช้กับ
   ข้อมูลลูกค้า/ราคาที่เป็นความลับทางธุรกิจจริงจัง แนะนำให้:
   - ย้าย insert/update/delete ไปอยู่หลัง Supabase Edge Function หรือ backend ของคุณเอง
     ที่ตรวจสอบสิทธิ์ผู้ใช้ก่อนทุกครั้ง แทนการเรียกตารางตรงจาก browser
   - หรือย้ายระบบ login ไปใช้ Supabase Auth แล้วเขียน RLS policy ผูกกับ `auth.uid()`

## การปรับแต่งเพิ่มเติมที่แนะนำ
- เพิ่ม Supabase Realtime subscription ให้หน้า "ภาพรวมทีม" อัปเดตสดโดยไม่ต้อง refresh
- เพิ่มการ hash รหัสผ่านด้วย Edge Function ก่อนบันทึกลง `users`
- เพิ่ม pagination ให้ตาราง pipeline ถ้าข้อมูลมีจำนวนมาก
