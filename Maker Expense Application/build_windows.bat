@echo off
echo ============================================
echo Project Pricer - Build Script
echo ============================================
echo.

echo Step 1: Checking current directory...
echo Current directory: %CD%
echo.

echo Step 2: Checking for required files...
if exist "project_pricer.py" (
    echo [OK] project_pricer.py found
) else (
    echo [ERROR] project_pricer.py NOT found!
    echo Please run this script in the same folder as project_pricer.py
    pause
    exit /b 1
)

if exist "excel_export.py" (
    echo [OK] excel_export.py found
) else (
    echo [ERROR] excel_export.py NOT found!
    echo Please make sure excel_export.py is in the same folder
    pause
    exit /b 1
)
echo.

echo Step 3: Installing dependencies...
pip install openpyxl
pip install pyinstaller
echo.

echo Step 4: Building executable...
echo This may take a few minutes...
pyinstaller --onefile --windowed --name="ProjectPricer" --add-data "excel_export.py;." project_pricer.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Build failed!
    echo Check the errors above for details
    pause
    exit /b 1
)
echo.

echo Step 5: Checking if executable was created...
if exist "dist\ProjectPricer.exe" (
    echo [OK] ProjectPricer.exe created successfully
    echo File size:
    dir "dist\ProjectPricer.exe" | find "ProjectPricer.exe"
) else (
    echo [ERROR] ProjectPricer.exe was not created!
    echo.
    echo Checking dist folder contents:
    if exist "dist" (
        dir dist
    ) else (
        echo dist folder does not exist
    )
    pause
    exit /b 1
)
echo.

echo Step 6: Moving executable to Desktop...
echo Desktop path: %USERPROFILE%\Desktop
if not exist "%USERPROFILE%\Desktop" (
    echo [WARNING] Desktop folder not found at standard location
    echo Checking alternative locations...
    if exist "%USERPROFILE%\OneDrive\Desktop" (
        echo [OK] Found OneDrive Desktop
        set DESKTOP_PATH=%USERPROFILE%\OneDrive\Desktop
    ) else (
        echo [ERROR] Could not find Desktop folder
        echo The executable is in the dist folder instead
        echo Location: %CD%\dist\ProjectPricer.exe
        pause
        exit /b 0
    )
) else (
    set DESKTOP_PATH=%USERPROFILE%\Desktop
)

move /Y "dist\ProjectPricer.exe" "%DESKTOP_PATH%\"
if %ERRORLEVEL% EQU 0 (
    echo [OK] Successfully moved to Desktop!
) else (
    echo [ERROR] Failed to move file
    echo The executable is in: %CD%\dist\ProjectPricer.exe
    echo You can manually move it to your Desktop
    pause
    exit /b 0
)
echo.

echo Step 7: Verifying file on Desktop...
if exist "%DESKTOP_PATH%\ProjectPricer.exe" (
    echo [OK] ProjectPricer.exe is on your Desktop!
    echo Full path: %DESKTOP_PATH%\ProjectPricer.exe
) else (
    echo [WARNING] File not found on Desktop
    echo Please check: %CD%\dist\ProjectPricer.exe
)
echo.

echo Step 8: Cleaning up build files...
if exist build (
    rmdir /S /Q build
    echo [OK] Removed build folder
)
if exist dist (
    rmdir /S /Q dist
    echo [OK] Removed dist folder
)
if exist ProjectPricer.spec (
    del /Q ProjectPricer.spec
    echo [OK] Removed spec file
)
echo.

echo ============================================
echo BUILD COMPLETE!
echo ============================================
echo.
echo Your executable should be on the Desktop:
echo %DESKTOP_PATH%\ProjectPricer.exe
echo.
echo If you don't see it, check your Desktop or search for:
echo "ProjectPricer.exe"
echo ============================================
pause
