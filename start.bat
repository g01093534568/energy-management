@echo off
chcp 65001 > nul
setlocal

cd /d "%~dp0"

echo ======================================
echo   에너지 관리 시스템 시작 중...
echo ======================================
echo.

REM Node.js 경로 설정
set "NODE_PATH=C:\Program Files\nodejs"
set "PATH=%NODE_PATH%;%PATH%"

REM Node.js 설치 확인
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ❌ Node.js가 설치되어 있지 않습니다.
    echo https://nodejs.org 에서 Node.js를 설치해주세요.
    echo.
    echo 또는 Node.js 설치 경로를 확인해주세요.
    echo 현재 경로: %NODE_PATH%
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('node -v') do set NODE_VERSION=%%i
echo ✅ Node.js 버전: %NODE_VERSION%
echo.

REM 기존 서버 프로세스 종료
echo 🔍 기존 서버 프로세스 확인 중...
tasklist /FI "IMAGENAME eq node.exe" 2>NUL | find /I /N "node.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo ⚠️  기존 Node.js 프로세스 발견. 종료 중...
    taskkill /F /IM node.exe >nul 2>nul
    timeout /t 2 /nobreak > nul
)

REM 서버 시작
echo 🚀 서버를 시작합니다...
start "에너지 관리 시스템" /MIN cmd /c "node server.js"

REM 서버 시작 대기
echo ⏳ 서버 시작 대기 중...
timeout /t 3 /nobreak > nul

REM 브라우저 자동 실행
echo 🌐 브라우저를 실행합니다...
start http://localhost:3000

echo.
echo ======================================
echo   ✅ 에너지 관리 시스템이 실행되었습니다!
echo ======================================
echo.
echo 📍 주소: http://localhost:3000
echo.
echo 💡 서버는 백그라운드에서 실행 중입니다.
echo    서버를 종료하려면 작업 관리자에서 node.exe를 종료하세요.
echo.
pause

endlocal
