# ESRI Patch List

เว็บไซต์แบบ Static Site (ไม่มี backend) สำหรับค้นหา/ดาวน์โหลดแพตช์ ArcGIS Enterprise, ซอฟต์แวร์ติดตั้ง และสคริปต์ช่วยงานที่เกี่ยวข้อง — ใช้งานภายในองค์กรเท่านั้น

> Powered By Khemmathat for Internal USE ONLY!!

> ⚠️ **สถานะ Software (หยุดอัปเดตแล้ว):** ส่วน Software จะ**ไม่มีการอัปเดตข้อมูลอีกต่อไป** เนื่องจากบัญชีที่ใช้เก็บไฟล์ต้นทาง (Google Drive/Sheet) **ไม่มีอยู่แล้ว** ลิงก์ดาวน์โหลดเดิมที่เคยบันทึกไว้ยังใช้งานได้อยู่ แต่ตัวข้อมูลใน Google Sheet (`SOFTWARE_CSV_URL` ใน `assets/app.js`) จะไม่ถูกอัปเดตเพิ่มอีก **ห้ามแก้ไข config/URL นี้** ดูรายละเอียดในหัวข้อ [Software](#2-software-indexhtmlsoftware)

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
│   ├── Autoinstallpatch.bat
│   └── autopatch_linux.sh
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
- ตัวกรอง: ค้นหาข้อความอิสระ, กรองตาม Component (Server/Portal/Data Store/...), กรองตาม Security (Y/N), กรองตาม **OS** (All/Windows/Linux/macOS)
- คอลัมน์ **Platform** แสดงค่า OS ที่แพตช์รองรับ (มาจากฟิลด์ `Platform` ของ Esri เช่น `"Linux,Windows"`) ตัวกรอง OS จะจับคู่แบบ substring กับค่านี้ (case-insensitive) เพื่อให้แพตช์ที่รองรับหลาย OS ยังปรากฏได้ในทุกตัวกรองที่เกี่ยวข้อง
- คอลัมน์ Download พยายามหาไฟล์แพตช์ตรงเวอร์ชันจาก `PatchFiles[]` โดยจับ pattern `ArcGIS-<ver>-`, `/PFA-<ver>-`, `/S-<ver>-` บนโดเมน `gisupdates.esri.com` เท่านั้น — ถ้าไม่เจอจะ fallback ไปหา key อื่น (`download_url`, `qfe_url`, ฯลฯ)
- เมื่อตั้งตัวกรอง OS เป็น **Linux** หรือ **Windows** ปุ่ม Download จะเลือกไฟล์ที่ตรงกับ OS นั้นโดยอัตโนมัติ (สังเกตจากนามสกุลไฟล์: `.tar`/`.tar.gz` = Linux, `.msp`/`.exe` = Windows) ถ้าไม่พบไฟล์เฉพาะ OS จะ fallback กลับไปใช้ไฟล์แรกที่ตรงเวอร์ชัน
- คอลัมน์ **Version** (มาจาก `rowObj.version` ของแต่ละแพตช์) แสดงเวอร์ชัน ArcGIS ที่แพตช์นั้นตรงกับไฟล์ที่จะดาวน์โหลดจริง — สำคัญเวลาดูที่ **All versions** เพราะ Esri มักออกแพตช์ชื่อเดียวกัน (เช่น "ArcGIS Desktop TLS Patch") ซ้ำหลายแถวสำหรับแต่ละเวอร์ชัน (10.3, 10.3.1, 10.4, ...) โดยที่ Released/Component/Platform/Security เหมือนกันหมด
- **มุมมอง All versions ไม่มีปุ่ม Download**: เนื่องจากแถวชื่อซ้ำกันหลายเวอร์ชันทำให้ผู้ใช้เสี่ยงกดดาวน์โหลดผิดเวอร์ชัน คอลัมน์ Download ในมุมมอง All versions จะแสดง badge "เลือกเวอร์ชันก่อนดาวน์โหลด" (disabled) แทนปุ่มดาวน์โหลดจริง ผู้ใช้ต้องเลือกเวอร์ชันที่ต้องการจาก dropdown Version ก่อน ปุ่ม Download จึงจะปรากฏและใช้งานได้ (ดู `renderPatchActionsCell()` ใน `assets/app.js` — เช็คจาก `activeSheet === "All_Enterprise"`)

### 2) Software (`index.html#software`)
> ⚠️ **หยุดอัปเดตแล้ว** — บัญชีที่ใช้เก็บไฟล์ต้นทางไม่มีอยู่แล้ว ข้อมูลใน Google Sheet จะไม่ถูกเพิ่ม/แก้ไขอีกต่อไป **ลิงก์ดาวน์โหลดที่มีอยู่เดิมยังใช้งานได้** แต่จะไม่มีซอฟต์แวร์เวอร์ชันใหม่เพิ่มเข้ามา และ**ห้ามแก้ไข** `SOFTWARE_CSV_URL` หรือ config อื่นที่เกี่ยวข้องกับส่วนนี้
- ดึงข้อมูลจาก Google Sheet ที่เผยแพร่เป็น CSV (`SOFTWARE_CSV_URL` ใน `assets/app.js`)
- เดา Component และ Version จากชื่อไฟล์/พาธด้วย regex heuristics (`inferComponent`, `inferVersion`) รองรับ ArcGIS Server/Portal/Data Store/Notebook/Web Adaptor/Pro/Desktop/Insights/Monitor/License Manager
- กรองตามเวอร์ชัน, Component, และค้นหาชื่อไฟล์

### 3) Scripts & Tools (`scripts.html`)
- การ์ดสำหรับแต่ละสคริปต์: คำอธิบาย, วิธีใช้งาน (ภาษาไทย), ปุ่ม **Open** (ไปที่ `view.html?file=...`) และ **Download ZIP** (ลิงก์ตรงไปยัง `patch/script/*.zip` บน GitHub)
- ปัจจุบันมี 3 สคริปต์ที่ผูกครบวงจร:
  - `script/DisableArcGISProUpdates.reg` — ปิดการตรวจอัปเดตอัตโนมัติของ ArcGIS Pro
  - `script/Autoinstallpatch.bat` — ติดตั้งแพตช์ `.msp` แบบออฟไลน์บน Windows
  - `script/autopatch_linux.sh` — ติดตั้งแพตช์ `.tar` แบบออฟไลน์บน Linux (แตกไฟล์ `.tar` ทั้งหมดในโฟลเดอร์ปัจจุบันแล้วรัน `applypatch` ในแต่ละโฟลเดอร์ให้อัตโนมัติ)

### 4) Script Viewer (`view.html`)
- อ่านพารามิเตอร์ `?file=` แล้วตรวจกับ `allowList` (hardcode ไว้ในไฟล์) ก่อน fetch เนื้อหามาแสดงแบบ read-only — ป้องกันไม่ให้เปิดไฟล์ใด ๆ ในเซิร์ฟเวอร์ได้ตามใจชอบ

## Automation (GitHub Actions)

| Workflow | Trigger | หน้าที่ |
|---|---|---|
| `update-patches.yml` | cron รายวัน 02:10 UTC (~09:10 ไทย) + manual dispatch | ดึง `patches.json` ล่าสุดจาก `https://downloads.esri.com/patch_notification/patches.json` เทียบ SHA256 กับของเดิม แล้ว commit/push เฉพาะตอนมีการเปลี่ยนแปลง |
| `build-patch-scripts-zips.yml` | push ที่แก้ `script/DisableArcGISProUpdates.reg`, `script/Autoinstallpatch.bat` หรือ `script/autopatch_linux.sh` + manual dispatch | zip ไฟล์สคริปต์แต่ละตัวใส่ `patch/script/*.zip` แล้ว commit เฉพาะตอนไฟล์ zip เปลี่ยน |

> หมายเหตุ (`.gitattributes`): `script/autopatch_linux.sh` ถูกบังคับให้เก็บ line ending แบบ LF เสมอ (`*.sh text eol=lf`) เพราะ shell script ที่มี CRLF จะรันไม่ได้บน Linux (`$'\r': command not found`) — สำคัญมากถ้าแก้ไฟล์นี้บนเครื่อง Windows

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
