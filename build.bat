@echo off
chcp 65001 >nul
echo ========================================================
echo  나이스 출결 자동 입력기 (NeisAttendanceRPA) 빌드 시작
echo ========================================================
echo.

echo [1/2] 기존 빌드 디렉터리 정리...
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"

echo.
echo [2/2] PyInstaller 패키징 실행 중 (--onedir, --noupx, --windowed)...
pyinstaller --onedir --noupx --windowed --collect-all playwright --name NeisAttendanceRPA neis_attendance_rpa.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================================
    echo  빌드 성공! 
    echo  결과 폴더: dist\NeisAttendanceRPA\NeisAttendanceRPA.exe
    echo ========================================================
) else (
    echo.
    echo [오류] 빌드 중 에러가 발생했습니다. (Error code: %ERRORLEVEL%)
)

pause
