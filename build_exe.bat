@echo off
echo Building Spotify Auto Controller executable...
echo.

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

REM Build the executable with icon
echo.
echo Building executable with icon...
pyinstaller --onedir --noconsole --icon=icon.ico --exclude-module multiprocessing --name="SpotifyAutoController" stopspotv1.py

REM Check if build was successful
if exist "dist\SpotifyAutoController\SpotifyAutoController.exe" (
    echo.
    echo Build completed successfully!
    echo Executable location: dist\SpotifyAutoController\SpotifyAutoController.exe
    echo File size: 
    dir /b dist\SpotifyAutoController\SpotifyAutoController.exe | for %%A in (?) do echo %%~zA bytes
    echo.
    echo The executable now includes the custom icon and improved error handling.
    echo Note: This creates a directory with the executable and dependencies.
) else (
    echo.
    echo Build failed! Check the output above for errors.
)

echo.
pause