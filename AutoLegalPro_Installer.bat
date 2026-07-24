::[Bat To Exe Converter]
::
::YAwzoRdxOk+EWAjk
::fBw5plQjdCyDJGyX8VAjFBZaWRaQAE+1EbsQ5+n//NaBo1sUV+0xNobY1dQ=
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF+5
::cxAkpRVqdFKZSDk=
::cBs/ulQjdF+5
::ZR41oxFsdFKZSDk=
::eBoioBt6dFKZSDk=
::cRo6pxp7LAbNWATEpCI=
::egkzugNsPRvcWATEpCI=
::dAsiuh18IRvcCxnZtBJQ
::cRYluBh/LU+EWAnk
::YxY4rhs+aU+IeA==
::cxY6rQJ7JhzQF1fEqQJhZkoaHGQ=
::ZQ05rAF9IBncCkqN+0xwdVs0
::ZQ05rAF9IAHYFVzEqQIRPQ9bDCiHO2q2RoUO54g=
::eg0/rx1wNQPfEVWB+kM9LVsJDGQ=
::fBEirQZwNQPfEVWB+kM9LVsJDCWuDAs=
::cRolqwZ3JBvQF1fEqQJQ
::dhA7uBVwLU+EWDk=
::YQ03rBFzNR3SWATElA==
::dhAmsQZ3MwfNWATElA==
::ZQ0/vhVqMQ3MEVWAtB9wSA==
::Zg8zqx1/OA3MEVWAtB9wSA==
::dhA7pRFwIByZRRnk
::Zh4grVQjdCuDJG6B9gIULQ1RQAuSMW60Eok9/fz0w++Ao0EtRu0DWZrP1ZiLJ+Ef2lfrSr8jxW5blMcJHlVdZhfL
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
@echo off
title Auto Legal Pro - Installer ^& Updater
color 0B
echo ==========================================
echo  AUTO LEGAL PRO - INSTALLER ^& UPDATER
echo ==========================================
echo.

:: KONFIGURASI GITHUB
set "REPO=nurrizki29/AutoLegalPro"
set "BRANCH=master"

echo [ INFO PENTING ]
echo Pastikan komputer terhubung ke internet.
echo Untuk melanjutkan instalasi/pembaruan, seluruh jendela 
echo Microsoft Word dan Excel harus ditutup.
echo.
echo 1. Simpan dokumen yang sedang Anda kerjakan.
echo 2. Tutup aplikasi Word dan Excel secara manual.
echo.
echo Menunggu Word dan Excel ditutup (Tekan CTRL+C untuk batal)...

:WAIT_LOOP
tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
set WORD_STATUS=%ERRORLEVEL%

tasklist /FI "IMAGENAME eq excel.exe" 2>NUL | find /I /N "excel.exe">NUL
set EXCEL_STATUS=%ERRORLEVEL%

if "%WORD_STATUS%"=="0" (
    timeout /t 2 >nul
    goto WAIT_LOOP
)
if "%EXCEL_STATUS%"=="0" (
    timeout /t 2 >nul
    goto WAIT_LOOP
)

echo.
color 0A
echo Aplikasi telah tertutup. Memulai proses...
echo Mengunduh file terbaru dari GitHub...

:: Membuat file script PowerShell
set "PS_SCRIPT=%TEMP%\AutoLegalSetup.ps1"

> "%PS_SCRIPT%" echo $ErrorActionPreference = 'Stop'
>> "%PS_SCRIPT%" echo [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
>> "%PS_SCRIPT%" echo $baseUrl = 'https://raw.githubusercontent.com/%REPO%/%BRANCH%'
>> "%PS_SCRIPT%" echo $wordDir = "$env:APPDATA\Microsoft\Word\STARTUP"
>> "%PS_SCRIPT%" echo $excelDir = "$env:APPDATA\Microsoft\Excel\XLSTART"
>> "%PS_SCRIPT%" echo $tplDir = "$env:APPDATA\Microsoft\Templates"
:: Membuat folder jika belum ada
>> "%PS_SCRIPT%" echo if (!(Test-Path $wordDir)) { New-Item -ItemType Directory -Force -Path $wordDir ^| Out-Null }
>> "%PS_SCRIPT%" echo if (!(Test-Path $excelDir)) { New-Item -ItemType Directory -Force -Path $excelDir ^| Out-Null }
>> "%PS_SCRIPT%" echo if (!(Test-Path $tplDir)) { New-Item -ItemType Directory -Force -Path $tplDir ^| Out-Null }
:: Mengunduh file utama
>> "%PS_SCRIPT%" echo Write-Host 'Mendownload AutoLegalPro.dotm...'
>> "%PS_SCRIPT%" echo Invoke-WebRequest -Uri "$baseUrl/AutoLegalPro.dotm" -OutFile "$wordDir\AutoLegalPro.dotm"
>> "%PS_SCRIPT%" echo Unblock-File -Path "$wordDir\AutoLegalPro.dotm"
>> "%PS_SCRIPT%" echo Write-Host 'Mendownload AutoLegalPro.xlam...'
>> "%PS_SCRIPT%" echo Invoke-WebRequest -Uri "$baseUrl/AutoLegalPro.xlam" -OutFile "$excelDir\AutoLegalPro.xlam"
>> "%PS_SCRIPT%" echo Unblock-File -Path "$excelDir\AutoLegalPro.xlam"
>> "%PS_SCRIPT%" echo Write-Host 'Mendownload TemplatePeraturan.dotx...'
>> "%PS_SCRIPT%" echo Invoke-WebRequest -Uri "$baseUrl/TemplatePeraturan.dotx" -OutFile "$tplDir\TemplatePeraturan.dotx"
>> "%PS_SCRIPT%" echo Unblock-File -Path "$tplDir\TemplatePeraturan.dotx"
>> "%PS_SCRIPT%" echo Write-Host 'Mendownload TemplateKeputusan.dotx...'
>> "%PS_SCRIPT%" echo Invoke-WebRequest -Uri "$baseUrl/TemplateKeputusan.dotx" -OutFile "$tplDir\TemplateKeputusan.dotx"
>> "%PS_SCRIPT%" echo Unblock-File -Path "$tplDir\TemplateKeputusan.dotx"
:: MENYALIN DIRI SENDIRI UNTUK MENJADI UPDATER
>> "%PS_SCRIPT%" echo Write-Host 'Menyiapkan modul Updater internal...'
>> "%PS_SCRIPT%" echo Copy-Item -Path '%~f0' -Destination "$wordDir\Updater.bat" -Force
>> "%PS_SCRIPT%" echo Write-Host 'Instalasi Berhasil!'

:: Mengeksekusi script PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
del "%PS_SCRIPT%"

echo.
echo Proses selesai! Silakan buka kembali Microsoft Word / Excel Anda.
timeout /t 5 >nul
exit