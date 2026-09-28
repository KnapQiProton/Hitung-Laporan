@echo off
title Menjalankan Web Server Laporan Keuangan
echo =========================================================
echo   Menjalankan Server Web (Simulasi Seperti di Railway)
echo =========================================================
echo   Akses di browser: http://localhost:3000
echo   Tekan Ctrl + C untuk menghentikan server.
echo =========================================================
start http://localhost:3000
node server.js
pause
