@echo off
REM ============================================================
REM  개인정보 검출 스캐너 - Windows EXE 빌드 스크립트
REM  더블클릭하거나 명령창에서 실행하세요.
REM ============================================================
chcp 65001 >nul
echo [1/3] 의존성 설치 중...
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error

echo.
echo [2/3] EXE 빌드 중...
pyinstaller pii_scanner.spec --noconfirm
if errorlevel 1 goto :error

echo.
echo [3/3] 완료!
echo   결과물: dist\개인정보검출스캐너.exe
echo.
pause
exit /b 0

:error
echo.
echo *** 빌드 실패. 위 메시지를 확인하세요. ***
pause
exit /b 1
