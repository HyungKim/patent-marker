@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
REM =============================================================================
REM  mark.bat ─ 브라우저 없이 마킹하기 (Windows)
REM    분석할 PPTX 파일을 이 아이콘 위에 끌어다 놓으면 바로 마킹합니다 (여러 개 가능).
REM    결과는 원본 옆에  파일이름_특허마킹.pptx  로 저장됩니다.
REM    (원본 폴더에 쓸 수 없으면 output 폴더에 저장하고 그렇게 알려 줍니다)
REM
REM    · Ollama 가 꺼져 있으면 켜고 · app\cli.py 를 실행합니다
REM    · 브라우저 화면(run.bat)이 필요 없는 방식입니다 — 업로드가 막힌 PC 용
REM  (설치가 안 되어 있으면 먼저 setup.bat 또는 setup_offline.bat)
REM =============================================================================
set "OLLAMA_URL=http://127.0.0.1:11434"

if "%~1"=="" (
  echo.
  echo   사용법: 마킹할 PPTX 파일을 이 mark.bat 아이콘 위에 끌어다 놓으세요 ^(여러 개 가능^).
  echo   결과는 원본 옆에  파일이름_특허마킹.pptx  로 저장됩니다.
  echo.
  pause
  exit /b 1
)

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
echo   특허 마킹 도구 - 명령행 모드
echo   (첫 슬라이드는 모델을 메모리에 올리느라 1~2분 더 걸립니다)
echo.
".venv\Scripts\python.exe" -m app.cli %*
echo.
pause
