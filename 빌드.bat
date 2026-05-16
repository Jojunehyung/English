@echo off
chcp 65001 > nul
echo.
echo  ============================================
echo   English Worksheet System  EXE 빌드
echo  ============================================
echo.

cd /d "%~dp0"

echo  [1/3] 패키지 확인 중...
pip install pyinstaller tkinterdnd2 gspread google-auth google-auth-oauthlib python-docx pyhwp pdfplumber python-pptx 2>nul

echo.
echo  [2/3] PyInstaller 빌드 중...
echo.

pyinstaller ^
  --onedir ^
  --windowed ^
  --name "English Worksheet System" ^
  --collect-data tkinterdnd2 ^
  --hidden-import pyhwp ^
  --hidden-import docx ^
  --hidden-import gspread ^
  --hidden-import google.oauth2.credentials ^
  --hidden-import google.auth.transport.requests ^
  --hidden-import google_auth_oauthlib.flow ^
  --exclude-module googleapiclient ^
  convert.py

echo.
echo  [3/3] 출력 폴더 생성 중...

set DIST=dist\English Worksheet System

if not exist "%DIST%\문제" mkdir "%DIST%\문제"
if not exist "%DIST%\결과" mkdir "%DIST%\결과"
if not exist "%DIST%\문제 서식" mkdir "%DIST%\문제 서식"

rem 서식 파일이 있으면 복사
if exist "문제 서식\standard_a.docx" copy /y "문제 서식\standard_a.docx" "%DIST%\문제 서식\" >nul

echo.
echo  ======================================================
echo   빌드 완료!
echo.
echo   결과 경로: %DIST%\
echo   이 폴더 전체를 ZIP으로 압축해서 선생님들께 배포하세요.
echo  ======================================================
echo.
pause
