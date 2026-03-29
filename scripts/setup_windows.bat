@echo off
echo ============================================
echo  App Tools Setup — Windows
echo ============================================
echo.

:: Check Git
where git >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Git is not installed.
    echo Download: https://git-scm.com/download/win
    pause & exit /b 1
)

:: Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed.
    echo Download: https://www.python.org/downloads/
    pause & exit /b 1
)

:: Create tools folder and clone/update repo
echo Setting up C:\AppTools...
if not exist "C:\AppTools" mkdir C:\AppTools

if exist "C:\AppTools\app-tools" (
    echo Repo already exists — pulling latest...
    git -C C:\AppTools\app-tools pull
) else (
    git clone https://github.com/qpx7mxr9/app-tools C:\AppTools\app-tools
)

:: Install Python dependencies
echo Installing dependencies...
pip install xlwings pandas
pip install -e C:\AppTools\app-tools

:: Install xlwings Excel add-in
echo Installing xlwings Excel add-in...
xlwings addin install

:: Create manual update script
echo @echo off > C:\AppTools\update.bat
echo echo Updating App Tools... >> C:\AppTools\update.bat
echo git -C C:\AppTools\app-tools pull >> C:\AppTools\update.bat
echo echo Done! Press any key to close. >> C:\AppTools\update.bat
echo pause >> C:\AppTools\update.bat

:: Schedule silent daily auto-pull at 8am
echo Scheduling daily auto-update at 8:00 AM...
schtasks /create /tn "AppTools Auto-Update" ^
  /tr "git -C C:\AppTools\app-tools pull" ^
  /sc daily /st 08:00 /f >nul

echo.
echo ============================================
echo  Setup complete!
echo.
echo  - Open your Excel workbook and click Run.
echo  - Tools update automatically every morning.
echo  - To update manually: C:\AppTools\update.bat
echo ============================================
pause
