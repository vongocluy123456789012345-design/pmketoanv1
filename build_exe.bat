@echo off
rem Dong goi AWBook thanh ung dung Windows (thu muc release\AWBook). Chay lai file nay sau moi lan sua code.
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (echo Chua co .venv - hay chay run.bat truoc 1 lan & pause & exit /b 1)
.venv\Scripts\pip install pyinstaller -q
set AIKEY=
if exist "%~dp0app\ai_builtin.json" set AIKEY=--add-data "%~dp0app\ai_builtin.json;app"
.venv\Scripts\python -m PyInstaller desktop.py --noconfirm --clean --onedir --noconsole --name AWBook ^
  --distpath release --workpath build --specpath build --paths . ^
  --add-data "%~dp0app\templates;app\templates" --add-data "%~dp0app\providers\providers.json;app\providers" --add-data "%~dp0app\ai_guide.md;app" --add-data "%~dp0app\ai_luat.md;app" %AIKEY% ^
  --collect-submodules app --collect-all playwright --collect-all winsdk --collect-all pypdfium2 ^
  --collect-all pdfplumber --collect-all pdfminer --collect-submodules uvicorn --collect-all xhtml2pdf ^
  --collect-all reportlab --collect-all svglib --collect-all html5lib --hidden-import psutil ^
  --hidden-import uvicorn.logging --hidden-import uvicorn.loops.auto --hidden-import uvicorn.protocols.http.auto ^
  --hidden-import uvicorn.protocols.websockets.auto --hidden-import uvicorn.lifespan.on --hidden-import multipart
copy /y HUONG_DAN.txt release\AWBook\ >nul
copy /y .env.example release\AWBook\ >nul
echo Xong: release\AWBook\AWBook.exe
pause
