@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
REM =============================================================================
REM  run.bat ─ 프로그램 실행 (Windows)
REM    · Ollama 가 꺼져 있으면 켜고
REM    · 웹 서버를 띄운 뒤
REM    · 브라우저를 자동으로 엽니다  →  http://127.0.0.1:8765
REM  종료: 이 창을 닫거나 Ctrl + C
REM  (설치가 안 되어 있으면 먼저 setup.bat 또는 setup_offline.bat)
REM =============================================================================
if "%PM_PORT%"=="" (set "PORT=8765") else (set "PORT=%PM_PORT%")
set "OLLAMA_URL=http://127.0.0.1:11434"

if not exist ".venv\Scripts\python.exe" (
  echo 먼저 setup.bat ^(인터넷 있음^) 또는 setup_offline.bat ^(인터넷 없음^) 을 실행하세요.
  pause
  exit /b 1
)

set "OLLAMA="
where ollama >nul 2>&1 && set "OLLAMA=ollama"
if not defined OLLAMA if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" set "OLLAMA=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"

curl -sf %OLLAMA_URL%/api/tags >nul 2>&1 && goto ollama_ready
if not defined OLLAMA (
  echo [경고] Ollama 가 설치되어 있지 않습니다. setup.bat 을 먼저 실행하세요.
  goto ollama_ready
)
echo Ollama 를 시작합니다...
start "Ollama" /MIN "%OLLAMA%" serve
set /a tries=0
:wait_ollama
timeout /t 1 /nobreak >nul
curl -sf %OLLAMA_URL%/api/tags >nul 2>&1 && goto ollama_ready
set /a tries+=1
if %tries% lss 30 goto wait_ollama
:ollama_ready

echo.
echo   특허 마킹 도구  -  http://127.0.0.1:%PORT%
echo   (이 창을 닫으면 프로그램이 종료됩니다)
echo.
REM 3초 뒤 브라우저를 연다 (서버가 뜰 시간을 줌)
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://127.0.0.1:%PORT%"
".venv\Scripts\python.exe" -m app.main
pause
