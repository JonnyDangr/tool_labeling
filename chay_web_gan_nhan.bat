@echo off
chcp 65001 >nul
title Web Gan Nhan Van Ban Phap Luat (Docx -> JSON / CSV)

echo ========================================================
echo   DANG KHOI DONG WEB GAN NHAN VAN BAN PHAP LUAT...
echo ========================================================

start "" "http://localhost:8000"
python "%~dp0web_gan_nhan\server.py"
pause

