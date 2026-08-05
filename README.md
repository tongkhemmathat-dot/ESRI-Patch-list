# ESRI Patch List

เว็บไซต์แบบ Static Site (ไม่มี backend) สำหรับค้นหา/ดาวน์โหลดแพตช์ ArcGIS Enterprise, ซอฟต์แวร์ติดตั้ง และสคริปต์ช่วยงานที่เกี่ยวข้อง — ใช้งานภายในองค์กรเท่านั้น

> Powered By Khemmathat for Internal USE ONLY!!

## Live Pages

| หน้า | ไฟล์ | หน้าที่ |
|---|---|---|
| หน้าแรก / เมนู / รายการแพตช์ / ซอฟต์แวร์ | `index.html` + `assets/app.js` | เมนูหลัก, ตารางแพตช์ (Patches), ตารางซอฟต์แวร์ (Software) |
| Scripts & Tools | `scripts.html` | การ์ดลิงก์ดาวน์โหลด/เปิดดูสคริปต์ช่วยงาน |
| Script Viewer | `view.html` | แสดงเนื้อหาไฟล์สคริปต์แบบ read-only ก่อนดาวน์โหลด |

ทั้งหมดรันได้ตรง ๆ จาก static hosting (เช่น GitHub Pages) ไม่ต้อง build step

## โครงสร้างโปรเจกต์

```
.
├── index.html                 # หน้าเมนู + ตาราง Patches / Software
├── scripts.html                # หน้ารวมสคริปต์และเครื่องมือ
├── view.html                   # ตัวแสดงเนื้อหาสคริปต์ (allowlist-gated)
├── assets/
│   ├── app.js                  # โลจิกทั้งหมดของหน้าเว็บ (fetch, filter, render)
│   └── styles.css              # ธีม light/dark และสไตล์ตาราง/การ์ด
├── data/
│   └── patches.json            # ข้อมูลแพตช์ ดึงจาก Esri อัตโนมัติทุกวัน
├── script/
│   ├── DisableArcGISProUpdates.reg
│   └── Autoinstallpatch.bat
├── patch/script/                # ไฟล์ .zip ที่ build อัตโนมัติจาก script/*
│   ├── DisableArcGISProUpdates.zip
│   ├── ArcGIS_OfflinePatch_AutoInstaller.zip
│   └── autopatch_linux.zip
└── .github/workflows/
    ├── update-patches.yml              # ดึง patches.json ใหม่ทุกวัน
    └── build-patch-scripts-zips.yml    # zip สคริปต์ใน script/ อัตโนมัติเมื่อมีการแก้ไข
```

## ฟีเจอร์หลัก

### 1) Patches (`index.html#patch`)
- โหลดข้อมูลจาก `data/patches.json` (no-store, fetch ใหม่ทุกครั้งที่เปิดแท็บ)
- รองรับข้อมูลได้หลายรูปแบบผ่าน `normalizePatchesJson()` (ทั้ง raw ของ Esri ที่เป็น `{Product:[{version, patches:[...]}]}` และ schema แบบ `{columns, rows}` ที่ normalize ไว้แล้ว)
- แยกข้อมูลเป็นชีทตามเวอร์ชัน (`v10_3`, `v11_1`, ...) บวกชีทรวม `All_Enterprise`
- ตัวกรอง: ค้นหาข้อความอิสระ, กรองตาม Component (Server/Portal/Data Store/...), กรองตาม Security (Y/N)
- คอลัมน์ Download พยายามหาไฟล์แพตช์ตรงเวอร์ชันจาก `PatchFiles[]` โดยจับ pattern `ArcGIS-<ver>-`, `/PFA-<ver>-`, `/S-<ver>-` บนโดเมน `gisupdates.esri.com` เท่านั้น — ถ้าไม่เจอจะ fallback ไปหา key อื่น (`download_url`, `qfe_url`, ฯลฯ)

### 2) Software (`index.html#software`)
- ดึงข้อมูลจาก Google Sheet ที่เผยแพร่เป็น CSV (`SOFTWARE_CSV_URL` ใน `assets/app.js`)
- เดา Component และ Version จากชื่อไฟล์/พาธด้วย regex heuristics (`inferComponent`, `inferVersion`) รองรับ ArcGIS Server/Portal/Data Store/Notebook/Web Adaptor/Pro/Desktop/Insights/Monitor/License Manager
- กรองตามเวอร์ชัน, Component, และค้นหาชื่อไฟล์

### 3) Scripts & Tools (`scripts.html`)
- การ์ดสำหรับแต่ละสคริปต์: คำอธิบาย, วิธีใช้งาน (ภาษาไทย), ปุ่ม **Open** (ไปที่ `view.html?file=...`) และ **Download ZIP** (ลิงก์ตรงไปยัง `patch/script/*.zip` บน GitHub)
- ปัจจุบันมี 2 สคริปต์ที่ผูกครบวงจร:
  - `script/DisableArcGISProUpdates.reg` — ปิดการตรวจอัปเดตอัตโนมัติของ ArcGIS Pro
  - `script/Autoinstallpatch.bat` — ติดตั้งแพตช์ `.msp` แบบออฟไลน์บน Windows

### 4) Script Viewer (`view.html`)
- อ่านพารามิเตอร์ `?file=` แล้วตรวจกับ `allowList` (hardcode ไว้ในไฟล์) ก่อน fetch เนื้อหามาแสดงแบบ read-only — ป้องกันไม่ให้เปิดไฟล์ใด ๆ ในเซิร์ฟเวอร์ได้ตามใจชอบ

## Automation (GitHub Actions)

| Workflow | Trigger | หน้าที่ |
|---|---|---|
| `update-patches.yml` | cron รายวัน 02:10 UTC (~09:10 ไทย) + manual dispatch | ดึง `patches.json` ล่าสุดจาก `https://downloads.esri.com/patch_notification/patches.json` เทียบ SHA256 กับของเดิม แล้ว commit/push เฉพาะตอนมีการเปลี่ยนแปลง |
| `build-patch-scripts-zips.yml` | push ที่แก้ `script/DisableArcGISProUpdates.reg` หรือ `script/Autoinstallpatch.bat` + manual dispatch | zip ไฟล์สคริปต์แต่ละตัวใส่ `patch/script/*.zip` แล้ว commit เฉพาะตอนไฟล์ zip เปลี่ยน |

> หมายเหตุ: ปัจจุบัน `patch/script/autopatch_linux.zip` ถูกอัปโหลดตรง ๆ (ไม่ได้ build จาก workflow) และยังไม่มี source ไฟล์ `.sh` ใน `script/`, ยังไม่ถูกผูกเข้ากับ `build-patch-scripts-zips.yml`, `scripts.html`, หรือ `allowList` ใน `view.html` — เป็นงานที่ยังค้างอยู่

## วิธีเพิ่มสคริปต์ใหม่

1. วางไฟล์ source ไว้ใน `script/` (เช่น `script/autopatch.sh`)
2. เพิ่ม path ของไฟล์เข้า `paths:` ใน `.github/workflows/build-patch-scripts-zips.yml` และเพิ่มขั้นตอน zip ให้ตรงกับสคริปต์ใหม่
3. เพิ่ม `"script/<ชื่อไฟล์>"` เข้า `allowList` ใน `view.html`
4. เพิ่มการ์ดใหม่ใน `scripts.html` (คัดลอกโครงสร้าง `.script-card` เดิม) พร้อมลิงก์ `view.html?file=...` และลิงก์ดาวน์โหลด zip

## วิธีรันดูตัวอย่างในเครื่อง

เป็น static site ล้วน ๆ เปิดผ่าน local web server ได้ทันที (ต้องใช้ server เพราะมีการ `fetch()` ไฟล์ JSON):

```bash
python -m http.server 8080
# แล้วเปิด http://localhost:8080/index.html
```

## Tech Stack

- Vanilla HTML/CSS/JS (ไม่มี framework, ไม่มี build step)
- ฟอนต์: Noto Sans Thai (Google Fonts)
- Theme: รองรับ light/dark พร้อม toggle และจำค่าไว้ใน `localStorage`
- ข้อมูล: `data/patches.json` (auto-sync จาก Esri) + Google Sheet CSV (สำหรับ Software)
