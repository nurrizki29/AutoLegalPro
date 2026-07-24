
@echo off
title Installer Auto Legal Pro
color 0B

echo =======================================================
echo                  AUTO LEGAL PRO (ALP)
echo                    versi: 1.0.0
echo =======================================================
echo.

:CEK_WORD
:: Mendeteksi apakah Microsoft Word (winword.exe) sedang berjalan
tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
if "%ERRORLEVEL%"=="0" (
    color 0E
    echo PERHATIAN: Microsoft Word masih terbuka!
    echo Mohon SIMPAN pekerjaan Anda dan TUTUP seluruh jendela Microsoft Word.
    echo.
    echo Jika Word sudah ditutup, tekan tombol apa saja pada keyboard untuk melanjutkan...
    pause >nul
    echo.
    goto CEK_WORD
)

:CEK_EXCEL
:: Mendeteksi apakah Microsoft Excel (excel.exe) sedang berjalan
tasklist /FI "IMAGENAME eq excel.exe" 2>NUL | find /I /N "excel.exe">NUL
if "%ERRORLEVEL%"=="0" (
    color 0E
    echo PERHATIAN: Microsoft Excel masih terbuka!
    echo Mohon SIMPAN pekerjaan Anda dan TUTUP seluruh jendela Microsoft Excel.
    echo.
    echo Jika Excel sudah ditutup, tekan tombol apa saja pada keyboard untuk melanjutkan...
    pause >nul
    echo.
    goto CEK_EXCEL
)

:: Jika lolos pengecekan (Word dan Excel tertutup), ubah warna jadi hijau dan mulai instalasi
color 0A
echo Memulai Instalasi...
echo.

:: Menentukan direktori tujuan
set "WORD_STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "EXCEL_STARTUP_DIR=%APPDATA%\Microsoft\Excel\XLSTART"
set "TEMPLATES_DIR=%APPDATA%\Microsoft\Templates"

:: Membuat folder jika ternyata belum ada di komputer tujuan
if not exist "%WORD_STARTUP_DIR%" mkdir "%WORD_STARTUP_DIR%"
if not exist "%EXCEL_STARTUP_DIR%" mkdir "%EXCEL_STARTUP_DIR%"
if not exist "%TEMPLATES_DIR%" mkdir "%TEMPLATES_DIR%"

:: Pengecekan keberadaan file Add-in utama
if not exist "%~dp0AutoLegalPro.dotm" (
    color 0C
    echo ERROR: File AutoLegalPro.dotm tidak ditemukan!
    echo Pastikan file instalasi berada di folder yang sama dengan file .bat ini.
    echo.
    pause
    exit
)
if not exist "%~dp0AutoLegalPro.xlam" (
    color 0C
    echo ERROR: File AutoLegalPro.xlam tidak ditemukan!
    echo Pastikan file instalasi Excel berada di folder yang sama dengan file .bat ini.
    echo.
    pause
    exit
)

echo Menyalin Add-in Auto Legal Pro (Word)...
copy /Y "%~dp0AutoLegalPro.dotm" "%WORD_STARTUP_DIR%\AutoLegalPro.dotm" >nul

echo Menyalin Add-in Auto Legal Pro (Excel)...
copy /Y "%~dp0AutoLegalPro.xlam" "%EXCEL_STARTUP_DIR%\AutoLegalPro.xlam" >nul

echo Menyalin Template Peraturan...
copy /Y "%~dp0TemplatePeraturan.dotx" "%TEMPLATES_DIR%\TemplatePeraturan.dotx" >nul

echo Menyalin Template Keputusan...
copy /Y "%~dp0TemplateKeputusan.dotx" "%TEMPLATES_DIR%\TemplateKeputusan.dotx" >nul

echo.
echo Membuka blokir keamanan sistem (Unblock-File)...
powershell -Command "Unblock-File -Path '%WORD_STARTUP_DIR%\AutoLegalPro.dotm'"
powershell -Command "Unblock-File -Path '%EXCEL_STARTUP_DIR%\AutoLegalPro.xlam'"
powershell -Command "Unblock-File -Path '%TEMPLATES_DIR%\TemplatePeraturan.dotx'"
powershell -Command "Unblock-File -Path '%TEMPLATES_DIR%\TemplateKeputusan.dotx'"
echo.
echo =======================================================
echo INSTALASI SELESAI!
echo =======================================================
echo Silakan buka kembali Microsoft Word Anda.
echo Tab "ALP" sudah terpasang.
echo.
pause