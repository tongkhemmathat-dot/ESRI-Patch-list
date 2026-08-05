#!/bin/bash

# หากเกิดข้อผิดพลาด ให้หยุดการทำงานทันที (เพื่อความปลอดภัย)
set -e

echo "=== เริ่มต้นกระบวนการติดตั้ง Patch ==="

# วนลูปหาไฟล์ .tar ทุกไฟล์ใน directory ปัจจุบัน
for archive in *.tar; do
    # ตรวจสอบว่ามีไฟล์ .tar อยู่จริงหรือไม่ (ป้องกันกรณีไม่มีไฟล์ .tar เลย)
    [ -e "$archive" ] || { echo "ไม่พบไฟล์ .tar ใน folder นี้ค่ะ"; exit 0; }

    echo "----------------------------------------"
    echo "กำลังจัดการไฟล์: $archive"

    # ดึงชื่อ directory ออกมาจากชื่อไฟล์ (ตัดนามสกุล .tar ออก)
    dir_name="${archive%.tar}"

    echo "1. กำลังแตกไฟล์ $archive..."
    tar -xf "$archive"

    # ตรวจสอบว่า directory ที่แตกออกมามีอยู่จริงหรือไม่
    if [ -d "$dir_name" ]; then
        echo "2. เข้าสู่โฟลเดอร์ $dir_name..."
        cd "$dir_name"

        # ตรวจสอบไฟล์ applypatch
        if [ -f "./applypatch" ]; then
            echo "3. กำหนดสิทธิ์ให้รันได้ และเริ่มติดตั้ง Patch..."
            chmod +x ./applypatch

            # รันการติดตั้ง
            ./applypatch

            echo "✓ ติดตั้ง Patch $archive สำเร็จแล้วค่ะ"
        else
            echo "❌ ไม่พบไฟล์ ./applypatch ใน $dir_name"
            cd ..
            exit 1
        fi

        # ถอยกลับออกมาที่ path หลัก เพื่อเตรียมทำงานกับไฟล์ถัดไป
        cd ..
    else
        echo "❌ ไม่พบโฟลเดอร์ $dir_name หลังแตกไฟล์ (โครงสร้างไฟล์อาจไม่ตรงตามชื่อ .tar)"
        exit 1
    fi
done

echo "----------------------------------------"
echo "🎉 ติดตั้ง Patch ทั้งหมดเรียบร้อยแล้วค่ะ!"
