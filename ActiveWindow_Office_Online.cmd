@echo off
color F0
mode con cols=98 lines=30
title Activate Windows and office by Key Online - https://github.com/BsNgChiThanh 
setlocal EnableExtensions EnableDelayedExpansion

:======================================================================================================================================================
:MAINMENU
cls
mode con cols=98 lines=30

echo.                                        
echo.                                       BSCK1. NGUYEN CHI THANH
echo.                                ACTIVE WINDOWS - OFFICE BY KEY ONLINE
echo.                     _________________________________________________________
echo.                    ^|                                                         ^|
Echo.                    ^|   [1] ACTIVE OFFICE                                     ^|
Echo.                    ^|                                                         ^|
Echo.                    ^|   [2] ACIVE WINDOWS                                     ^|  
Echo.                    ^|                                                         ^|
Echo.                    ^|   [3] Thoat.                                            ^|
Echo.                    ^|_________________________________________________________^|
ECHO.            
choice /C:123 /N /M ".                              Nhap lua chon cua ban [1,2,3] : "
if errorlevel 3 goto:Exit
if errorlevel 2 goto:ACTIVEWINDOWS
if errorlevel 1 goto:ACTIVEOFFICE


:===============================================================================================
:ACTIVEWINDOWS

﻿chcp 65001 >nul
@echo off
title ACTIVATE WINDOWS 7 8 8.1 10 11 By Key Online - Copyright (C) Dominic Savio. All rights reserved.
mode con: cols=115 lines=35
chcp 65001 >nul
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
	
cls
type Text\Win\WinOnline.txt
color f0
ver
@echo.
@echo 1. Dang kiem tra ban quyen
cscript //nologo %windir%\system32\slmgr.vbs /xpr |findstr "permanently" >nul
if %errorlevel%==0  (
@echo                                 === WINDOWS DA KICH HOAT BAN QUYEN VINH VIEN ===
pause >nul
exit
) else (
@echo                              === WINDOWS CHUA DUOC KICH HOAT BAN QUYEN VINH VIEN ===
@echo.
@echo Nhan phim bat ky de tiep tuc kich hoat...
pause >nul
goto begin
)

:begin
@echo.
set /p key= 2. Nhap Key : 
@echo.

@echo 3. Dang go bo key cu ra khoi may
cscript //nologo %windir%\system32\slmgr.vbs /upk

@echo 4. Dang xoa may chu KMS neu co
cscript //nologo %windir%\system32\slmgr.vbs /ckms

@echo 5. Dang xoa key trong Registry
cscript //nologo %windir%\system32\slmgr.vbs /cpky

@echo 6. Dang cai dat Key
cscript //nologo %windir%\system32\slmgr.vbs /ipk %key%

@echo 7. Dang thuc hien lenh kich hoat
cscript //nologo %windir%\system32\slmgr.vbs /ato

@echo 8. Dang kiem tra ban quyen
cscript //nologo %windir%\system32\slmgr.vbs /xpr |findstr "permanently" >nul
if %errorlevel%==0  (
@echo    === Da kich hoat ban quyen VINH VIEN ===
@echo %key% >Key_Win.txt
pause >nul
exit


) else (
@echo         === Kich hoat khong thanh cong, vui long kiem tra lai Key! ===
@echo    === Hoac Key ban nhap da het luot Online, hay thu kich hoat By Phone ===
@echo.
@echo Nhan phim bat ky de thu kich hoat lai...
pause >nul
exit

:===============================================================================================
:ACTIVEOFFICE

﻿chcp 65001 >nul
@echo off
title ACTIVATE OFFICE 2010-2013-2016-2019-2021 By Key Online - Copyright (C) Dominic Savio. All rights reserved.
mode con: cols=122 lines=40
chcp 65001 >nul
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
	
cls
@echo.
@echo.                                                  ACTIVATE AIO TOOLS
@echo.                               Kich Hoat Office 2010-2013-2016-2019-2021 Bang Key Online
@echo.
@echo.
@echo.                                 =====================================================
@echo.                                 [  4. Office 2010                 : Nhan phim so 4  ]
@echo.                                 [  5. Office 2013                 : Nhan phim so 5  ]
@echo.                                 [  6. Office 2016,2019,2021       : Nhan phim so 6  ]
@echo.                                 [  7. Office 2024                 : Nhan phim so 7  ]
@echo.                                 [  8. Thoat                       : Nhan phim so 8  ]
@echo.                                 =====================================================
color f0
@echo.
@echo ===========================
Choice /N /C 45678 /M "* Nhap lua chon cua ban [4,5,6,7,8]: "
if %errorlevel% == 8 goto:exit
if %errorlevel% == 7 ( set "xx=16" & goto vogia)
if %errorlevel% == 6 ( set "xx=16" & goto vogia)
if %errorlevel% == 5 ( set "xx=15" & goto vogia)
if %errorlevel% == 4 ( set "xx=14" & goto vogia)


:vogia
if exist "%ProgramFiles%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office%xx%"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office%xx%"
goto begin

:begin
@echo.
set /p key= 1. Nhap Key : 
@echo.
@echo 2. Dang cai dat Key moi
cscript OSPP.VBS /inpkey:%key%
@echo.
@echo 3. Dang kich hoat ban quyen
cscript OSPP.VBS /act
@echo.
@echo 4. Dang kiem tra ban quyen
cscript OSPP.VBS /dstatus |findstr "LICENSED" >nul
if %errorlevel%==0 (
@echo    === Da kich hoat ban quyen VINH VIEN ===
@echo %key% >Key_Office.txt
pause >nul
goto exit
) else (
@echo ---------------------------------------
@echo ERROR CODE = MO TA LOI TUONG UNG
@echo 0xC004C060 : Key cua ban da bi chan - Block
@echo 0xC004C020 : Key da het luot kich hoat Online. Co the chuyen sang kich hoat By Phone
@echo 0xC004C003 : Key hoac Office ban dung la khong chinh hang. Hoac Office da bi loi
@echo 0xC004C008 : Khong the su dung Key da chi dinh. Co the chuyen sang kich hoat By Phone
@echo 0xC004F069 : Khong tim thay SKU - Key va phien ban Office khong tuong thich
@echo ---------------------------------------
@echo === Xuat Hien ERROR CODE Khac. Vui Long Tham Khao ERROR DESCRIPTION Ben Tren! ===
pause >nul
exit
)


:exit
@echo.
timeout 3
start https://github.com/BsNgChiThanh
exit
