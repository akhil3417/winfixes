@echo off
setlocal enabledelayedexpansion
title OneDrive Repair Utility

:: Run as administrator
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Requesting administrator privileges...
    goto UACPrompt
) else (
    goto gotAdmin
)

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs"
    pushd "%CD%"
    CD /D "%~dp0"

:menu
cls
echo.
echo ============================================
echo          ONEDRIVE REPAIR UTILITY
echo ============================================
echo.
echo  1. Check OneDrive Status
echo  2. Disconnect OneDrive
echo  3. Reset OneDrive
echo  4. Reinstall OneDrive
echo  5. Fix Common Errors
echo  6. Repair Permissions
echo  7. Clear OneDrive Cache
echo  8. Exit
echo.
echo ============================================
echo.

set /p choice="Enter your choice (1-8): "

if "%choice%"=="1" goto check_status
if "%choice%"=="2" goto disconnect
if "%choice%"=="3" goto reset
if "%choice%"=="4" goto reinstall
if "%choice%"=="5" goto fix_common
if "%choice%"=="6" goto fix_permissions
if "%choice%"=="7" goto clear_cache
if "%choice%"=="8" goto exit_script

echo Invalid choice. Please try again.
timeout /t 2 >nul
goto menu

:check_status
cls
echo.
echo ============================================
echo        CHECKING ONEDRIVE STATUS
echo ============================================
echo.

:: Check if OneDrive is running
tasklist /fi "imagename eq OneDrive.exe" | find /i "OneDrive.exe" >nul
if %errorlevel% equ 0 (
    echo [+] OneDrive is running.
) else (
    echo [-] OneDrive is not running.
)

:: Check OneDrive installation
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    echo [+] OneDrive is installed in the default location.
) else (
    echo [-] OneDrive is not installed in the default location.
)

:: Check internet connection
ping -n 1 onedrive.live.com >nul
if %errorlevel% equ 0 (
    echo [+] Connection to OneDrive.live.com established.
) else (
    echo [-] Unable to connect to OneDrive.live.com.
)

:: Check if OneDrive starts with Windows
reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v OneDrive >nul 2>&1
if %errorlevel% equ 0 (
    echo [+] OneDrive is set to start with Windows.
) else (
    echo [-] OneDrive is not set to start with Windows.
)

echo.
pause
goto menu

:disconnect
cls
echo.
echo ============================================
echo           DISCONNECT ONEDRIVE
echo ============================================
echo.
echo Closing OneDrive...
taskkill /f /im OneDrive.exe >nul 2>&1

echo Disconnecting OneDrive...
echo This will unlink your OneDrive accounts without deleting your files.
echo.
set /p confirm="Are you sure you want to continue? (Y/N): "
if /i "%confirm%"=="Y" (
    echo.
    echo Disconnecting OneDrive...
    
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /shutdown
        "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /unlink
        echo OneDrive accounts successfully disconnected.
    ) else (
        echo OneDrive is not installed in the default location.
    )
) else (
    echo Operation canceled.
)

echo.
pause
goto menu

:reset
cls
echo.
echo ============================================
echo            RESET ONEDRIVE
echo ============================================
echo.
echo This will reset OneDrive and remove all settings.
echo Your files will not be deleted, but you will need to reconfigure OneDrive.
echo.
set /p confirm="Are you sure you want to continue? (Y/N): "
if /i "%confirm%"=="Y" (
    echo.
    echo Resetting OneDrive...
    
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /reset
        echo OneDrive has been successfully reset.
        echo Restarting OneDrive...
        start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    ) else (
        echo OneDrive is not installed in the default location.
    )
) else (
    echo Operation canceled.
)

echo.
pause
goto menu

:reinstall
cls
echo.
echo ============================================
echo         REINSTALL ONEDRIVE
echo ============================================
echo.
echo This will uninstall and then reinstall OneDrive.
echo Your files will not be deleted, but you will need to reconfigure OneDrive.
echo.
set /p confirm="Are you sure you want to continue? (Y/N): "
if /i "%confirm%"=="Y" (
    echo.
    echo Uninstalling OneDrive...
    
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    if exist "%SYSTEMROOT%\SysWOW64\OneDriveSetup.exe" (
        %SYSTEMROOT%\SysWOW64\OneDriveSetup.exe /uninstall
    ) else (
        if exist "%SYSTEMROOT%\System32\OneDriveSetup.exe" (
            %SYSTEMROOT%\System32\OneDriveSetup.exe /uninstall
        ) else (
            echo OneDrive is not installed in the default location.
            goto:reinstall_end
        )
    )
    
    echo OneDrive uninstalled successfully.
    timeout /t 10 >nul
    
    echo Reinstalling OneDrive...
    if exist "%SYSTEMROOT%\SysWOW64\OneDriveSetup.exe" (
        %SYSTEMROOT%\SysWOW64\OneDriveSetup.exe
    ) else (
        if exist "%SYSTEMROOT%\System32\OneDriveSetup.exe" (
            %SYSTEMROOT%\System32\OneDriveSetup.exe
        ) else (
            echo Unable to find OneDrive installer.
            goto:reinstall_end
        )
    )
    
    echo OneDrive reinstalled successfully.
) else (
    echo Operation canceled.
)

:reinstall_end
echo.
pause
goto menu

:fix_common
cls
echo.
echo ============================================
echo         FIX COMMON ONEDRIVE ERRORS
echo ============================================
echo.
echo Choose the error to fix:
echo.
echo  1. OneDrive does not start
echo  2. Sync errors
echo  3. File conflicts
echo  4. Connection issues
echo  5. High resource usage
echo  6. Back to main menu
echo.
echo ============================================
echo.

set /p error_choice="Enter your choice (1-6): "

if "%error_choice%"=="1" goto fix_startup
if "%error_choice%"=="2" goto fix_sync
if "%error_choice%"=="3" goto fix_conflicts
if "%error_choice%"=="4" goto fix_connection
if "%error_choice%"=="5" goto fix_resources
if "%error_choice%"=="6" goto menu

echo Invalid choice. Please try again.
timeout /t 2 >nul
goto fix_common

:fix_startup
cls
echo.
echo ============================================
echo     FIX - ONEDRIVE DOES NOT START
echo ============================================
echo.
echo Applying fixes...

taskkill /f /im OneDrive.exe >nul 2>&1

echo 1/5 - Re-registering DLLs...
regsvr32 /s "%SYSTEMROOT%\System32\shell32.dll"

echo 2/5 - Checking auto-start...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v OneDrive /t REG_SZ /d "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /f >nul 2>&1

echo 3/5 - Resetting OneDrive...
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /reset >nul 2>&1
)

echo 4/5 - Cleaning registry...
reg delete "HKCU\Software\Microsoft\OneDrive" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Policies\Microsoft\OneDrive" /f >nul 2>&1

echo 5/5 - Restarting OneDrive...
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive restarted.
) else (
    echo OneDrive is not installed in the default location.
)

echo.
echo Fixes complete. Check if OneDrive now starts correctly.
echo.
pause
goto fix_common

:fix_sync
cls
echo.
echo ============================================
echo        FIX - SYNC ERRORS
echo ============================================
echo.
echo Applying fixes...

taskkill /f /im OneDrive.exe >nul 2>&1

echo 1/4 - Clearing sync cache...
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Personal" (
    del /f /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Personal\*.*" >nul 2>&1
)
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Business1" (
    del /f /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Business1\*.*" >nul 2>&1
)

echo 2/4 - Resetting ODSP caches...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" (
    rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" >nul 2>&1
)

echo 3/4 - Checking exclusion lists...
reg delete "HKCU\SOFTWARE\Microsoft\OneDrive\Accounts\Personal\Excluded" /f >nul 2>&1
reg delete "HKCU\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Excluded" /f >nul 2>&1

echo 4/4 - Restarting OneDrive...
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive restarted.
) else (
    echo OneDrive is not installed in the default location.
)

echo.
echo Fixes complete. Check if sync issues are resolved.
echo Note: The first sync may take longer after this operation.
echo.
pause
goto fix_common

:fix_conflicts
cls
echo.
echo ============================================
echo        FIX - FILE CONFLICTS
echo ============================================
echo.
echo This will help you resolve file conflicts.
echo.

echo Checking for conflict files...
if exist "%USERPROFILE%\OneDrive\*conflict*" (
    echo Conflict files were found in your OneDrive folder.
) else (
    echo No conflict files detected.
)

echo.
echo Tips for resolving conflicts:
echo 1. Open your OneDrive folder and search for files with "conflict" in the name.
echo 2. Compare versions and keep the most up-to-date one.
echo 3. Delete conflict files once you’ve saved important changes.
echo.
echo To avoid future conflicts:
echo - Avoid editing the same file simultaneously on multiple devices.
echo - Ensure OneDrive has finished syncing before shutting down.
echo - Regularly check OneDrive sync status.
echo.

pause
goto fix_common

:fix_connection
cls
echo.
echo ============================================
echo       FIX - CONNECTION ISSUES
echo ============================================
echo.
echo Applying fixes...

echo 1/5 - Checking network connection...
ping -n 1 onedrive.live.com >nul
if %errorlevel% equ 0 (
    echo    Connected to OneDrive.live.com.
) else (
    echo    Unable to connect to OneDrive.live.com.
    echo    Check your Internet connection and try again.
    echo.
    pause
    goto fix_common
)

echo 2/5 - Resetting Winsock...
netsh winsock reset >nul 2>&1

echo 3/5 - Resetting TCP/IP stack...
netsh int ip reset >nul 2>&1

echo 4/5 - Flushing DNS cache...
ipconfig /flushdns >nul 2>&1

echo 5/5 - Restarting OneDrive...
taskkill /f /im OneDrive.exe >nul 2>&1
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive restarted.
) else (
    echo OneDrive is not installed in the default location.
)

echo.
echo Fixes complete. Check if connection issues are resolved.
echo It is recommended to restart your computer for changes to take full effect.
echo.
set /p reboot="Do you want to restart your computer now? (Y/N): "
if /i "%reboot%"=="Y" (
    shutdown /r /t 60 /c "Scheduled restart to apply network changes."
    echo Your computer will restart in 60 seconds. Save your work.
    echo To cancel, run 'shutdown /a' in a new command window.
)

echo.
pause
goto fix_common

:fix_resources
cls
echo.
echo ============================================
echo      FIX - HIGH RESOURCE USAGE
echo ============================================
echo.
echo Applying fixes...

taskkill /f /im OneDrive.exe >nul 2>&1

echo 1/3 - Adjusting bandwidth settings...
reg add "HKCU\SOFTWARE\Microsoft\OneDrive\Accounts\Personal" /v "DiskSpaceCheckThresholdMB" /t REG_DWORD /d 1024 /f >nul 2>&1
reg add "HKCU\SOFTWARE\Microsoft\OneDrive" /v "MaxBandwidthKB" /t REG_DWORD /d 512 /f >nul 2>&1

echo 2/3 - Optimizing performance settings...
reg add "HKCU\SOFTWARE\Microsoft\OneDrive" /v "ProcessResourcePriority" /t REG_DWORD /d 1 /f >nul 2>&1

echo 3/3 - Restarting OneDrive...
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive restarted with optimized settings.
) else (
    echo OneDrive is not installed in the default location.
)

echo.
echo Fixes complete. Check if resource usage has improved.
echo.
echo Additional tips:
echo - Sync fewer folders to reduce load.
echo - Exclude large or temporary files from sync.
echo - Keep OneDrive updated to the latest version.
echo.
pause
goto fix_common

:fix_permissions
cls
echo.
echo ============================================
echo          REPAIR ONEDRIVE PERMISSIONS
echo ============================================
echo.
echo This will fix OneDrive folder and registry permissions.
echo.
set /p confirm="Are you sure you want to continue? (Y/N): "
if /i "%confirm%"=="Y" (
    echo.
    echo Repairing permissions...
    
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    echo 1/3 - Resetting OneDrive folder permissions...
    icacls "%USERPROFILE%\OneDrive" /reset /T /C /Q >nul 2>&1
    
    echo 2/3 - Fixing registry permissions...
    reg add "HKCU\Software\Microsoft\OneDrive" /f >nul 2>&1
    
    echo 3/3 - Restarting OneDrive...
    timeout /t 5 >nul
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
        echo OneDrive restarted.
    ) else (
        echo OneDrive is not installed in the default location.
    )
    
    echo Permissions repair complete.
) else (
    echo Operation canceled.
)

echo.
pause
goto menu

:clear_cache
cls
echo.
echo ============================================
echo          CLEAR ONEDRIVE CACHE
echo ============================================
echo.
echo This will delete all OneDrive cache files.
echo This can fix many sync issues, but the first sync after this may take longer.
echo.
set /p confirm="Are you sure you want to continue? (Y/N): "
if /i "%confirm%"=="Y" (
    echo.
    echo Clearing cache...
    
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    echo 1/3 - Deleting cache files...
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\cache" (
        rmdir /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\cache" >nul 2>&1
    )
    
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\settings" (
        rmdir /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\settings" >nul 2>&1
    )
    
    if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" (
        rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" >nul 2>&1
    )
    
    echo 2/3 - Recreating cache folders...
    mkdir "%LOCALAPPDATA%\Microsoft\OneDrive\cache" >nul 2>&1
    mkdir "%LOCALAPPDATA%\Microsoft\OneDrive\settings" >nul 2>&1
    
    echo 3/3 - Restarting OneDrive...
    timeout /t 5 >nul
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
        echo OneDrive restarted.
    ) else (
        echo OneDrive is not installed in the default location.
    )
    
    echo Cache cleared successfully.
) else (
    echo Operation canceled.
)

echo.
pause
goto menu

:exit_script
cls
echo.
echo ============================================
echo      THANK YOU FOR USING THE UTILITY
echo ============================================
echo.
echo Goodbye!
echo.
timeout /t 3 >nul
exit


