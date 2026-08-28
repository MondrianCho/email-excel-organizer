@echo off
chcp 65001 >nul

echo ======================================
echo  이메일 접수 엑셀 정리 자동화
echo ======================================
echo.
echo - 엑셀 파일명을 'yyyy-mm-dd 이메일 접수_초안.xlsx' 로 변경합니다.
echo - E열 제목(2행~)을 보고 K열(업무종류), L/N/O열을 자동 입력합니다.
echo.
echo 처리할 폴더의 전체 경로를 입력하세요.
echo 예) C:\Users\Erin\Desktop\이메일접수폴더
echo.

set /p TARGET_FOLDER=▶ 폴더 경로: 

if "%TARGET_FOLDER%"=="" (
    echo [오류] 폴더 경로가 입력되지 않았습니다.
    pause
    exit /b 1
)

echo.
echo [실행 중] 잠시만 기다려주세요...
echo.

python "%~dp0email_excel_organize.py" "%TARGET_FOLDER%"

echo.
echo ===== 작업 완료 =====
pause
