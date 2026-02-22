@echo off
echo Building Spotify Auto Controller executable...
echo.

set ICON_PNG=stopspoti.png
set ICON_ICO=stopspoti.ico
set MAIN_SCRIPT=stopspotiv1.py

REM Check if virtual environment exists
if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
) else (
    echo Virtual environment not found. Please run setup first.
    pause
    exit /b 1
)

REM Install PyInstaller if not already installed
echo Installing/ensuring PyInstaller is available...
pip install pyinstaller

REM Ensure Pillow for PNG-ICO conversion
echo Ensuring Pillow is available for icon conversion...
pip install Pillow

REM Convert PNG logo to ICO
if exist "%ICON_PNG%" (
    echo Converting %ICON_PNG% to %ICON_ICO% ...
    python -c "from PIL import Image; Image.open(r'%ICON_PNG%').save(r'%ICON_ICO%')" || goto :icon_fail
) else (
    echo Icon PNG not found: %ICON_PNG%
    pause
    exit /b 1
)

:icon_fail
if not exist "%ICON_ICO%" (
    echo Icon conversion failed.
    pause
    exit /b 1
)

REM Build the executable with icon
echo.
echo Building single-file executable with icon...
pyinstaller --onefile --noconsole --icon=%ICON_ICO% --exclude-module multiprocessing --name="SpotifyAutoController" %MAIN_SCRIPT%

REM Check if build was successful
if exist "dist\SpotifyAutoController.exe" (
    echo.
    echo Build completed successfully!
    echo Executable location: dist\SpotifyAutoController.exe
    echo.
    echo The executable now includes the custom icon and improved error handling.
) else (
    echo.
    echo Build failed! Check the output above for errors.
)

echo.
pause