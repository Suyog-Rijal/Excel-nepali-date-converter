@echo off
setlocal enabledelayedexpansion
title Office Add-in Installer v1.0

REM Registry path for Office Add-ins
set "REGISTRY_PATH=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

:main
call :show_banner
call :show_menu
call :get_user_choice
goto :eof

:show_banner
cls
echo.
color 0B
echo  ===============================================================================
echo                            OFFICE ADD-IN INSTALLER                             
echo                               Version 1.0                                      
echo  ===============================================================================
echo.
color 0F
echo   Professional tool for managing Office Add-ins via registry entries
echo   Supports manifest-based add-in installation and removal
echo.
echo  ===============================================================================
echo.
goto :eof

:show_menu
color 0E
echo   ** MAIN MENU **
echo   ---------------
echo.
color 0A
echo   [1] Install Add-in     - Install Office Add-in from manifest file
color 0C
echo   [2] Uninstall Add-in   - Remove currently installed add-in
color 0F
echo   [0] Exit               - Close the installer
echo.
echo  ===============================================================================
echo.
goto :eof

:get_user_choice
set /p "choice=  Enter your choice (1, 2, or 0): "
echo.

if "%choice%"=="1" (
    call :install_addin
    call :wait_for_keypress
    goto main
) else if "%choice%"=="2" (
    call :uninstall_addin
    call :wait_for_keypress
    goto main
) else if "%choice%"=="0" (
    call :exit_program
) else (
    color 0C
    echo   ** INVALID CHOICE! **
    echo   ---------------------
    echo   Please enter 1, 2, or 0.
    call :wait_for_keypress
    goto main
)
goto :eof

:install_addin
color 0E
echo   ** INSTALLING ADD-IN **
echo   -----------------------
echo.
color 0F
echo   Please select your Office Add-in manifest file (.xml)
echo.

REM Create a VBS script to open file dialog
set "vbs_script=%temp%\file_picker.vbs"
(
echo Set objDialog = CreateObject("MSComDlg.CommonDialog"^)
echo objDialog.DialogTitle = "Select Office Add-in Manifest File"
echo objDialog.Filter = "XML Manifest Files (*.xml^)|*.xml|All Files (*.*^)|*.*"
echo objDialog.InitDir = CreateObject("WScript.Shell"^).SpecialFolders("Desktop"^)
echo objDialog.ShowOpen
echo If objDialog.FileName ^<^> "" Then
echo     WScript.Echo objDialog.FileName
echo Else
echo     WScript.Echo "CANCELLED"
echo End If
) > "%vbs_script%"

REM Try MSComDlg.CommonDialog first, if not available use folder picker
for /f "delims=" %%i in ('cscript //nologo "%vbs_script%" 2^>nul') do set "manifest_path=%%i"

REM If MSComDlg failed, use alternative method
if not defined manifest_path (
    call :alternative_file_picker
)

if "%manifest_path%"=="CANCELLED" (
    color 0E
    echo   ** CANCELLED **
    echo   ---------------
    echo   No file selected. Installation cancelled.
    goto :eof
)

if not defined manifest_path (
    color 0E
    echo   ** CANCELLED **
    echo   ---------------
    echo   No file selected. Installation cancelled.
    goto :eof
)

REM Check if file exists
if not exist "%manifest_path%" (
    color 0C
    echo   ** ERROR! **
    echo   ------------
    echo   Selected file does not exist.
    goto :eof
)

REM Create registry key and set manifest path
reg add "%REGISTRY_PATH%" /f >nul 2>&1
reg add "%REGISTRY_PATH%" /v "Manifest" /t REG_SZ /d "%manifest_path%" /f >nul 2>&1

if !errorlevel! equ 0 (
    color 0A
    echo   ** SUCCESS! **
    echo   --------------
    echo   Add-in installed successfully!
    echo   Location: %manifest_path%
    echo.
    color 0B
    echo   Note: Please restart your Office applications to see the add-in.
) else (
    color 0C
    echo   ** ERROR! **
    echo   -----------
    echo   Failed to install add-in. Check permissions and try again.
)

del "%vbs_script%" >nul 2>&1
color 0F
goto :eof

:alternative_file_picker
color 0B
echo   Using alternative file selection method...
echo.
color 0F
echo   Please drag and drop your manifest file here and press Enter:
set /p "manifest_path=  File path: "

REM Remove quotes if present
set "manifest_path=%manifest_path:"=%"
goto :eof

:uninstall_addin
color 0E
echo   ** UNINSTALLING ADD-IN **
echo   -------------------------
echo.

REM Check if registry key exists
reg query "%REGISTRY_PATH%" >nul 2>&1
if !errorlevel! neq 0 (
    color 0E
    echo   ** NO ADD-IN FOUND **
    echo   ---------------------
    echo   No add-in registry key found. Nothing to uninstall.
    goto :eof
)

REM Get current manifest path
for /f "tokens=3*" %%a in ('reg query "%REGISTRY_PATH%" /v "Manifest" 2^>nul ^| findstr "REG_SZ"') do (
    set "current_manifest=%%a %%b"
)
set "current_manifest=!current_manifest: =!"

if defined current_manifest (
    color 0E
    echo   Current add-in: !current_manifest!
    echo.
    color 0F
    set /p "confirm=  Are you sure you want to remove this add-in? (y/n): "
    
    if /i "!confirm!"=="y" (
        reg delete "%REGISTRY_PATH%" /f >nul 2>&1
        if !errorlevel! equ 0 (
            echo.
            color 0A
            echo   ** SUCCESS! **
            echo   --------------
            echo   Add-in uninstalled successfully!
            echo.
            color 0B
            echo   Note: Please restart your Office applications to complete removal.
        ) else (
            echo.
            color 0C
            echo   ** ERROR! **
            echo   -----------
            echo   Failed to uninstall add-in. Check permissions and try again.
        )
    ) else (
        echo.
        color 0E
        echo   ** CANCELLED **
        echo   --------------
        echo   Uninstallation cancelled.
    )
) else (
    color 0E
    echo   ** NO ADD-IN FOUND **
    echo   ---------------------
    echo   No add-in manifest found in registry.
)
color 0F
goto :eof

:wait_for_keypress
echo.
echo   Press any key to continue...
pause >nul
goto :eof

:exit_program
color 0A
echo   ** GOODBYE! **
echo   --------------
echo   Thank you for using Office Add-in Installer!
echo.
timeout /t 2 /nobreak >nul
exit /b 0