@echo off
chcp 65001 >nul

echo ====================================
echo [ Excel AI Chatbot ��ġ �غ� ���� ]
echo ====================================
echo.

:: STEP 1. Python ��ġ Ȯ��
where python >nul 2>&1
if errorlevel 1 (
    echo Python�� ��ġ�Ǿ� ���� �ʽ��ϴ�. �ڵ� ��ġ�� �����մϴ�...
    powershell -Command "Invoke-WebRequest https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe -OutFile python-installer.exe"
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
    del python-installer.exe
) else (
    echo Python ��ġ Ȯ�� �Ϸ�
)

echo.
echo Python ������ ��ġ ��...
pip install -r requirements.txt

:: STEP 2. Ollama ��ġ Ȯ��
where ollama >nul 2>&1
if errorlevel 1 (
    echo Ollama�� ��ġ�Ǿ� ���� �ʽ��ϴ�. �ڵ� ��ġ�� �����մϴ�...
    powershell -Command "Invoke-WebRequest https://ollama.com/download/OllamaSetup.exe -OutFile ollama-installer.exe"
    start /wait ollama-installer.exe
    del ollama-installer.exe
    echo Ollama ��ġ �� �α׾ƿ� �Ǵ� ������� �ʿ��� �� �ֽ��ϴ�.
    echo Enter Ű�� ���� ����մϴ�.
    pause
) else (
    echo Ollama ��ġ Ȯ�� �Ϸ�
)

:: STEP 3. �� �ٿ�ε�
echo.
echo LLaMA �� �ٿ�ε� ��...
ollama pull llama3

:: STEP 4. ���� ����
echo.
echo FastAPI ������ �����մϴ�...
start /B python app.py

:: STEP 5. ���� ���� ����
echo.
echo ���� ����(joo.xlsm)�� �����մϴ�...
start "" "joo.xlsm"

echo.
echo Excel AI Chatbot ��ġ �� ���� �Ϸ�!
pause
