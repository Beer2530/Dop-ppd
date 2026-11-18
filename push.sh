# 1. ตั้งค่าให้ Git บันทึกรหัสผ่าน/Token ไว้โดยอัตโนมัติ (ทำครั้งเดียวจบ)
git config --global credential.helper manager

# 2. เชื่อมต่อโปรเจกต์กับ GitHub Repository ของคุณ
git remote remove origin 2>/dev/null
git remote add origin https://github.com/Beer2530/Dop-ppd.git

# 3. เตรียมไฟล์และบันทึกประวัติการแก้ไข
git add .
git commit -m "Auto update web app configuration and files"

# 4. ตั้งค่า Branch หลักเป็น main และส่งโค้ดขึ้น GitHub
git branch -M main
git push -u origin main