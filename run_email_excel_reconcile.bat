@echo off
chcp 65001 >nul

echo ======================================
echo  오전 이메일 백업 ↔ 엑셀 제목 누락 검사
echo ======================================
echo.
echo 상위 폴더 경로를 입력하세요. (그 안에 '오전 이메일 백업' 폴더가 있어야 합니다)
echo 예) C:\Users\USER\Desktop\접수작업
echo.

set /p TARGET_FOLDER=▶ 상위 폴더 경로: 

if "%TARGET_FOLDER%"=="" (
    echo [오류] 폴더 경로가 입력되지 않았습니다.
    pause
    exit /b 1
)

echo.
echo [실행 중] 잠시만 기다려주세요...
echo.

python "%~dp0email_excel_reconcile.py" "%TARGET_FOLDER%"

echo.
echo ===== 작업 완료 =====
pause
