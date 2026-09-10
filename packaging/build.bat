@echo off
REM ---------------------------------------------------------------------------
REM Build the Windows executables and installer.
REM
REM Run from the repository root:   packaging\build.bat
REM
REM Produces:
REM   dist\DocxFindReplace\DocxFindReplace.exe      installed build
REM   dist\DocxFindReplace-Portable.exe             single-file portable build
REM   dist\installer\DocxFindReplace-Setup-x.y.z.exe
REM
REM Requires Python 3.8+, the packages in requirements-dev.txt, and Inno Setup 6
REM (https://jrsoftware.org/isdl.php) for the installer step.
REM ---------------------------------------------------------------------------
setlocal enabledelayedexpansion

cd /d "%~dp0.."

echo === Running tests before packaging ===
python -m pytest -q
if errorlevel 1 (
    echo.
    echo Tests failed - not packaging a broken build.
    exit /b 1
)

echo.
echo === Reading version ===
for /f "delims=" %%v in ('python -c "import version; print(version.__version__)"') do set APP_VERSION=%%v
echo Version: %APP_VERSION%

echo.
echo === Generating Windows version resource ===
python packaging\make_version_file.py
if errorlevel 1 exit /b 1

echo.
echo === Building installed (one-folder) build ===
python -m PyInstaller packaging\app.spec --noconfirm --clean
if errorlevel 1 exit /b 1

echo.
echo === Building portable (one-file) build ===
python -m PyInstaller packaging\app-portable.spec --noconfirm
if errorlevel 1 exit /b 1

echo.
echo === Building installer ===
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
    echo.
    echo Inno Setup 6 was not found. Install it from https://jrsoftware.org/isdl.php
    echo The executables in dist\ are still usable; only the setup.exe was skipped.
    exit /b 1
)
"%ISCC%" /DAppVersion=%APP_VERSION% packaging\installer.iss
if errorlevel 1 exit /b 1

echo.
echo === Done ===
dir /b dist
dir /b dist\installer
endlocal
