@echo off
setlocal

cd /d "%~dp0\.."

set "TEMPLATE_PATH="
for %%F in ("%CD%\assets\*.docx") do set "TEMPLATE_PATH=%%~fF"
if "%TEMPLATE_PATH%"=="" (
  echo No .docx template found in assets.
  exit /b 1
)
set "BUILD_TEMPLATE_PATH=%CD%\build\pyinstaller\standard_template.docx"
if not exist "%CD%\build\pyinstaller" mkdir "%CD%\build\pyinstaller"
copy /Y "%TEMPLATE_PATH%" "%BUILD_TEMPLATE_PATH%" >nul

py -3 -m pip install --upgrade pip
if errorlevel 1 exit /b %errorlevel%
py -3 -m pip install -r requirements.txt pyinstaller
if errorlevel 1 exit /b %errorlevel%

py -3 -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "WordStyleStandardizer" ^
  --add-data="%BUILD_TEMPLATE_PATH%:assets" ^
  --distpath "dist\windows" ^
  --workpath "build\pyinstaller" ^
  --specpath "build\pyinstaller" ^
  "scripts\word_style_standardizer_gui.py"
if errorlevel 1 exit /b %errorlevel%

echo.
echo Generated: dist\windows\WordStyleStandardizer.exe
pause
