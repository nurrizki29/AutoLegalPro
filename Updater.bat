@echo off
title Auto Legal Pro Updater
color 0B
echo ==========================================
echo    AUTO LEGAL PRO - SYSTEM UPDATER
echo ==========================================
echo.

:: Ganti dengan username dan nama repository GitHub Anda
set "REPO=UsernameGitHubAnda/NamaRepositoryAnda"

echo 1. Menunggu penutupan Microsoft Word dan Excel...
echo.
echo    [ INFO PENTING ]
echo    Silakan kembali ke jendela Word / Excel Anda.
echo    Jika ada dokumen yang belum disimpan, simpan pekerjaan Anda
echo    pada kotak dialog yang muncul.
echo.
echo    Sistem ini akan diam menunggu hingga aplikasi benar-benar tertutup...
echo.

:WAIT_WORD
:: Memeriksa apakah Word masih terbuka
tasklist /FI "IMAGENAME eq winword.exe" 2>NUL | find /I /N "winword.exe">NUL
if "%ERRORLEVEL%"=="0" (
    :: Jika masih terbuka, jeda 2 detik dan cek lagi
    timeout /t 2 >nul
    goto WAIT_WORD
)

:WAIT_EXCEL
:: Memeriksa apakah Excel masih terbuka
tasklist /FI "IMAGENAME eq excel.exe" 2>NUL | find /I /N "excel.exe">NUL
if "%ERRORLEVEL%"=="0" (
    :: Jika masih terbuka, jeda 2 detik dan cek lagi
    timeout /t 2 >nul
    goto WAIT_EXCEL
)

echo Aplikasi terdeteksi telah tertutup secara aman!
echo.
echo 2. Menghubungkan ke Server GitHub...

:: Menggunakan PowerShell untuk memanggil API GitHub dan mendownload aset
powershell -NoProfile -ExecutionPolicy Bypass -Command "^
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; ^
$apiUrl = 'https://api.github.com/repos/%REPO%/releases/latest'; ^
try { $release = Invoke-RestMethod -Uri $apiUrl } catch { Write-Host 'Gagal menghubungi GitHub!'; timeout /t 5; exit }; ^
Write-Host ('Versi rilis terbaru ditemukan: ' + $release.tag_name); ^
$wordDir = \"$env:APPDATA\Microsoft\Word\STARTUP\"; ^
$excelDir = \"$env:APPDATA\Microsoft\Excel\XLSTART\"; ^
$tplDir = \"$env:APPDATA\Microsoft\Templates\"; ^
foreach ($asset in $release.assets) { ^
    $url = $asset.browser_download_url; ^
    $name = $asset.name; ^
    $dest = ''; ^
    if ($name -eq 'AutoLegalPro.dotm') { $dest = Join-Path $wordDir $name } ^
    elseif ($name -eq 'AutoLegalPro.xlam') { $dest = Join-Path $excelDir $name } ^
    elseif ($name -like 'Template*.dotx') { $dest = Join-Path $tplDir $name }; ^
    if ($dest) { ^
        Write-Host ('Mendownload ' + $name + '...'); ^
        Invoke-WebRequest -Uri $url -OutFile $dest; ^
        Unblock-File -Path $dest; ^
    } ^
} ^
Write-Host ''; ^
Write-Host 'Pembaruan Berhasil!';"

echo.
echo Proses selesai. Silakan buka kembali Microsoft Word / Excel Anda.
timeout /t 5 >nul
exit