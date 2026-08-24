@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"
REM =============================================================================
REM  setup.bat ─ Windows 용 "처음 한 번" 설치 (인터넷 필요)
REM
REM  하는 일
REM    1) Python 3.11+ 확인 (없으면 winget 으로 3.12 설치)
REM    2) .venv 가상환경 만들고 라이브러리 설치
REM    3) Ollama 설치 (없으면 winget)
REM    4) Ollama 서버 기동
REM    5) 모델(qwen3:14b) 내려받기  ← 약 9GB. 가장 오래 걸리는 단계
REM
REM  사용법     이 파일을 더블클릭  (또는 터미널에서  setup.bat)
REM  모델 변경  set PM_MODEL=qwen3:8b  입력 후  setup.bat
REM  오프라인   인터넷이 없는 PC 에서는 setup_offline.bat 을 쓰세요.
REM =============================================================================
if "%PM_MODEL%"=="" (set "MODEL=qwen3:14b") else (set "MODEL=%PM_MODEL%")
set "OLLAMA_URL=http://127.0.0.1:11434"

echo.
echo ============================================================
echo   특허 마킹 도구 - 처음 한 번 설치 (인터넷 필요)
echo ============================================================

echo.
echo [1/5] Python 확인
set "PY="
py -3.12 -c "print()" >nul 2>&1 && set "PY=py -3.12"
if not defined PY ( py -3.13 -c "print()" >nul 2>&1 && set "PY=py -3.13" )
if not defined PY ( py -3.11 -c "print()" >nul 2>&1 && set "PY=py -3.11" )
if not defined PY ( python -c "import sys; raise SystemExit(0 if sys.version_info>=(3,11) else 1)" >nul 2>&1 && set "PY=python" )
if defined PY goto py_ok
echo    Python 3.12 가 없어 winget 으로 설치합니다... (몇 분 걸립니다)
winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
if errorlevel 1 (
  echo.
  echo    [실패] winget 으로 설치가 안 됩니다.
  echo           https://www.python.org/downloads/ 에서 Python 3.12 를 받아 설치하세요.
  echo           설치 화면에서 "Add python.exe to PATH" 를 반드시 체크한 뒤 이 파일을 다시 실행하세요.
  pause
  exit /b 1
)
REM 방금 설치한 Python 은 이 창의 PATH 에 아직 없으므로 직접 추가
set "PATH=%LOCALAPPDATA%\Programs\Python\Launcher;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe") else (set "PY=py -3.12")
:py_ok
echo    사용: %PY%

echo.
echo [2/5] 가상환경(.venv) 및 라이브러리 설치
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv
  if errorlevel 1 ( echo    [실패] 가상환경 생성 & pause & exit /b 1 )
)
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul 2>&1
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 ( echo    [실패] 라이브러리 설치 - 인터넷 연결을 확인하세요 & pause & exit /b 1 )
".venv\Scripts\python.exe" -c "import pptx, fastapi, uvicorn, multipart"
if errorlevel 1 ( echo    [실패] 라이브러리 확인 & pause & exit /b 1 )
echo    완료

echo.
echo [3/5] Ollama 설치 확인
set "OLLAMA="
where ollama >nul 2>&1 && set "OLLAMA=ollama"
if not defined OLLAMA if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" set "OLLAMA=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
if defined OLLAMA goto ollama_ok
echo    Ollama 를 winget 으로 설치합니다...
winget install -e --id Ollama.Ollama --accept-source-agreements --accept-package-agreements
if errorlevel 1 (
  echo    [실패] https://ollama.com/download 에서 OllamaSetup.exe 를 받아 설치한 뒤 다시 실행하세요.
  pause
  exit /b 1
)
set "OLLAMA=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
:ollama_ok
echo    사용: %OLLAMA%

echo.
echo [4/5] Ollama 서버 기동
curl -sf %OLLAMA_URL%/api/tags >nul 2>&1 && goto ollama_ready
echo    서버를 시작합니다 (최소화된 "Ollama" 창이 하나 열립니다)...
start "Ollama" /MIN "%OLLAMA%" serve
set /a tries=0
:wait_ollama
timeout /t 1 /nobreak >nul
curl -sf %OLLAMA_URL%/api/tags >nul 2>&1 && goto ollama_ready
set /a tries+=1
if %tries% lss 30 goto wait_ollama
echo    [실패] Ollama 서버가 응답하지 않습니다. 시작 메뉴에서 Ollama 를 실행한 뒤 다시 시도하세요.
pause
exit /b 1
:ollama_ready
echo    실행 중

echo.
echo [5/5] 모델 내려받기 (%MODEL%) - 약 9GB, 수 분 ~ 수십 분
"%OLLAMA%" list 2>nul | findstr /B /C:"%MODEL%" >nul
if errorlevel 1 (
  "%OLLAMA%" pull %MODEL%
  if errorlevel 1 ( echo    [실패] 모델 다운로드 - 인터넷 연결을 확인하세요 & pause & exit /b 1 )
) else (
  echo    이미 있음
)

echo.
echo ============================================================
echo   설치 완료!  이제 run.bat 을 실행하세요.
echo ============================================================
pause
