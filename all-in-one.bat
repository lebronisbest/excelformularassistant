@echo off
chcp 65001 >nul

echo ====================================
echo [ Excel AI Chatbot 설치 준비 스크립트 ]
echo ====================================
echo.

:: STEP 1. Python 설치 확인
where python >nul 2>&1
if errorlevel 1 (
    echo Python이 설치되어 있지 않습니다. 자동 설치를 시작합니다...
    powershell -Command "Invoke-WebRequest https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe -OutFile python-installer.exe"
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
    del python-installer.exe
) else (
    echo Python 설치 확인 완료
)

echo.
echo Python 패키지 설치 중...
pip install -r requirements.txt

:: STEP 2. Ollama 설치 확인
where ollama >nul 2>&1
if errorlevel 1 (
    echo Ollama가 설치되어 있지 않습니다. 자동 설치를 시작합니다...
    powershell -Command "Invoke-WebRequest https://ollama.com/download/OllamaSetup.exe -OutFile ollama-installer.exe"
    start /wait ollama-installer.exe
    del ollama-installer.exe
    echo Ollama 설치 후 로그아웃 또는 재시작이 필요할 수 있습니다.
    echo Enter 키를 눌러 계속합니다.
    pause
) else (
    echo Ollama 설치 확인 완료
)

:: STEP 3. 모델 다운로드
echo.
echo LLaMA 모델 다운로드 중...
ollama pull llama3

:: STEP 4. 서버 시작
echo.
echo FastAPI 서버를 시작합니다...
start /B python app.py

:: STEP 5. 엑셀 파일 열기
echo.
echo 엑셀 파일(joo.xlsm)을 실행합니다...
start "" "joo.xlsm"

echo.
echo Excel AI Chatbot 설치 및 실행 완료!
pause
