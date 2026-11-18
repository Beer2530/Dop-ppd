@echo off
git init
git add .
git commit -m "Update web app"
git branch -M main
git remote add origin https://github.com/Beer2530/Dop-ppd.git 2>nul
git push -u origin main
pause