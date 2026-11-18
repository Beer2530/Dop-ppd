#!/bin/bash

# 1. เริ่มต้น Git หากยังไม่มี
if [ ! -d ".git" ]; then
  git init
  git branch -M main
  git remote add origin https://github.com/Beer2530/Dop-ppd.git
fi

# 2. ตรวจสอบและตั้งค่า remote ให้ถูกต้องเสมอ
git remote set-url origin https://github.com/Beer2530/Dop-ppd.git 2>/dev/null || git remote add origin https://github.com/Beer2530/Dop-ppd.git

# 3. เพิ่มไฟล์ทั้งหมดและ Commit พร้อมประทับเวลาปัจจุบัน
git add .
git commit -m "Update web app: $(date '+%Y-%m-%d %H:%M:%S')"

# 4. ส่งโค้ดขึ้น GitHub
git push -u origin main

echo "✅ ส่งโค้ดขึ้น GitHub (Beer2530/Dop-ppd) เรียบร้อยแล้ว"