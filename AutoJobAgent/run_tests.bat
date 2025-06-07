@echo off
echo ========================================
echo   AutoJobAgent Test Runner
echo ========================================
echo.

:: Check if Python is installed
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: Python not found in PATH
    echo Please install Python 3.10 64-bit and try again.
    goto :end
)

:: Check Python version
python --version
echo.

:: Check for virtual environment
if exist .venv\Scripts\activate.bat (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
) else if exist env\Scripts\activate.bat (
    echo Activating virtual environment...
    call env\Scripts\activate.bat
)

:: Check for required packages
echo Checking required packages...
python -m pip show rich >nul 2>nul || echo WARNING: rich package not installed
python -m pip show playwright >nul 2>nul || echo WARNING: playwright package not installed
python -m pip show ollama >nul 2>nul || echo WARNING: ollama package not installed
python -m pip show fastapi >nul 2>nul || echo WARNING: fastapi package not installed
echo.

:: Run the integration tests
echo Running integration tests...
python test_integration.py

:: Check if Word is installed (for PDF conversion)
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe" >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo.
    echo WARNING: Microsoft Word not detected
    echo PDF conversion may not work properly.
)

:: Check if Ollama is running
curl -s http://localhost:11434/api/tags >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo.
    echo WARNING: Ollama service not detected at http://localhost:11434
    echo LLM features may not work properly.
    echo To install Ollama, visit: https://ollama.ai/download
)

echo.
echo ========================================
echo   Test Runner Complete
echo ========================================
echo.
echo If all tests passed, you're ready to use AutoJobAgent!
echo To start the application, run: python main.py Nirajan
echo.

:end
pause
