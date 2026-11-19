@echo off
chcp 65001 >nul
echo ========================================================
echo กำลังส่งโค้ดทั้งหมดขึ้น GitHub (Beer2530/Dop-ppd)...
echo ========================================================
echo.

git init
git config user.name "Beer2530"
git config user.email "Prompong2530@gmail.com"
git add .
git commit -m "Update web app with 0ms cache, dual route, and file upload progress"
git branch -M main
git remote remove origin 2>nul
git remote add origin https://github.com/Beer2530/Dop-ppd.git
git push -u origin main

echo.
echo ========================================================
echo [OK] ส่งโค้ดขึ้น GitHub เรียบร้อยสมบูรณ์แล้ว!
echo Vercel จะนำขึ้นหน้าเว็บจริงให้อัตโนมัติใน 15-20 วินาที
echo ========================================================
echo.
pause