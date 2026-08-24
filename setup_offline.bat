@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"
REM =============================================================================
REM  setup_offline.bat ─ Windows 용 "인터넷 없는 PC" 설치
REM
REM  미리 인터넷 되는 컴퓨터에서  bash tools/make_offline_bundle.sh windows  로 만든
REM  offline_bundle\ 폴더가 이 폴더 안에 있어야 합니다.
REM
REM    offline_bundle\wheels\   파이썬 라이브러리
REM    offline_bundle\python\   python-3.12.x-amd64.exe   (Python 이 없을 때만 사용)
REM    offline_bundle\ollama\   OllamaSetup.exe           (Ollama 가 없을 때만 사용)
REM    offline_bundle\models\   모델 파일 (약 9GB)
REM
REM  사용법   이 파일을 더블클릭
REM =============================================================================
set "B=%~dp0offline_bundle"
if "%PM_MODEL%"=="" (set "MODEL=qwen3:14b") else (set "MODEL=%PM_MODEL%")
set "OLLAMA_URL=http://127.0.0.1:11434"

echo.
echo ============================================================
echo   특허 마킹 도구 - 오프라인 설치 (offline_bundle 사용)
echo ============================================================
if not exist "%B%\wheels" (
  echo.
  echo [실패] offline_bundle\wheels 폴더가 없습니다.
  echo        인터넷 되는 PC 에서 tools\make_offline_bundle.sh windows 를 실행해
  echo        만든 offline_bundle 폴더를 이 폴더 안에 함께 복사하세요.
  pause
  exit /b 1
)

echo.
echo [1/5] Python 확인
set "PY="
py -3.12 -c "print()" >nul 2>&1 && set "PY=py -3.12"
if not defined PY ( py -3.13 -c "print()" >nul 2>&1 && set "PY=py -3.13" )
if not defined PY ( py -3.11 -c "print()" >nul 2>&1 && set "PY=py -3.11" )
if not defined PY ( python -c "import sys; raise SystemExit(0 if sys.version_info>=(3,11) else 1)" >nul 2>&1 && set "PY=python" )
if defined PY goto py_ok
set "PYINST="
for %%F in ("%B%\python\python-*.exe") do set "PYINST=%%~fF"
if not defined PYINST (
  echo    [실패] Python 3.11+ 가 없고 offline_bundle\python 에 설치 파일도 없습니다.
  pause
  exit /b 1
)
echo    Python 을 설치합니다 (진행 창이 뜨면 끝날 때까지 기다리세요)...
"%PYINST%" /passive InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0
REM 방금 설치한 Python 은 이 창의 PATH 에 아직 없으므로 직접 추가
set "PATH=%LOCALAPPDATA%\Programs\Python\Launcher;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe") else (set "PY=py -3.12")
%PY% -c "print()" >nul 2>&1
if errorlevel 1 ( echo    [실패] Python 설치 확인 실패. PC 를 재부팅한 뒤 이 파일을 다시 실행해 보세요. & pause & exit /b 1 )
:py_ok
echo    사용: %PY%

echo.
echo [2/5] 가상환경(.venv) 및 라이브러리 (오프라인 휠)
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv
  if errorlevel 1 ( echo    [실패] 가상환경 생성 & pause & exit /b 1 )
)
".venv\Scripts\python.exe" -m pip install --no-index --find-links "%B%\wheels" -r requirements.txt
if errorlevel 1 ( echo    [실패] 휠 설치 - offline_bundle 이 Windows 64bit / Python 3.12 용으로 만들어졌는지 확인 & pause & exit /b 1 )
".venv\Scripts\python.exe" -c "import pptx, fastapi, uvicorn, multipart"
if errorlevel 1 ( echo    [실패] 라이브러리 확인 & pause & exit /b 1 )
echo    완료

echo.
echo [3/5] Ollama
set "OLLAMA="
where ollama >nul 2>&1 && set "OLLAMA=ollama"
if not defined OLLAMA if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" set "OLLAMA=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
if defined OLLAMA goto ollama_ok
if not exist "%B%\ollama\OllamaSetup.exe" (
  echo    [실패] Ollama 가 없고 offline_bundle\ollama\OllamaSetup.exe 도 없습니다.
  pause
  exit /b 1
)
echo    Ollama 설치 프로그램을 실행합니다. 설치 창에서 Install 을 누르고 끝나면 돌아옵니다...
"%B%\ollama\OllamaSetup.exe"
set "OLLAMA=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
if not exist "%OLLAMA%" ( echo    [실패] Ollama 설치 확인 실패 & pause & exit /b 1 )
:ollama_ok
echo    사용: %OLLAMA%

echo.
echo [4/5] 모델 파일 복사 (약 9GB, 몇 분 걸립니다)
if exist "%B%\models\manifests" (
  if not exist "%USERPROFILE%\.ollama\models" mkdir "%USERPROFILE%\.ollama\models"
  xcopy "%B%\models" "%USERPROFILE%\.ollama\models\" /E /I /Y /Q >nul
  echo    복사 완료 -^> %USERPROFILE%\.ollama\models
) else (
  echo    [경고] offline_bundle\models 가 없습니다. 모델을 따로 준비해야 합니다.
)

echo.
echo [5/5] Ollama 서버 기동 및 모델 확인
curl -sf %OLLAMA_URL%/api/tags >nul 2>&1 && goto ollama_ready
start "Ollama" /MIN "%OLLAMA%" serve
set /a tries=0
:wait_ollama
timeout /t 1 /nobreak >nul
curl -sf %OLLAMA_URL%/api/tags >nul 2>&1 && goto ollama_ready
set /a tries+=1
if %tries% lss 30 goto wait_ollama
echo    [경고] Ollama 서버가 아직 응답하지 않습니다. 시작 메뉴에서 Ollama 를 한 번 실행해 보세요.
:ollama_ready
"%OLLAMA%" list 2>nul | findstr /B /C:"%MODEL%" >nul
if errorlevel 1 ( echo    [경고] 모델 %MODEL% 이 목록에 없습니다. ) else ( echo    모델 %MODEL% 준비됨 )

echo.
echo ============================================================
echo   설치 완료!  이제 run.bat 을 실행하세요.
echo ============================================================
pause
