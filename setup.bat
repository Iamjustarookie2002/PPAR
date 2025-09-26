@echo off
REM Automated setup script for Windows
REM This script creates a virtual environment, installs dependencies, and sets up the project

echo 🚀 Setting up PPA Excel Processor...

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python is not installed or not in PATH. Please install Python 3 first.
    pause
    exit /b 1
)

REM Check if pip is installed
pip --version >nul 2>&1
if errorlevel 1 (
    echo ❌ pip is not installed. Please install pip first.
    pause
    exit /b 1
)

echo ✅ Python and pip are available

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo 📦 Creating virtual environment...
    python -m venv venv
    echo ✅ Virtual environment created
) else (
    echo ✅ Virtual environment already exists
)

REM Activate virtual environment
echo 🔧 Activating virtual environment...
call venv\Scripts\activate.bat

REM Upgrade pip
echo ⬆️ Upgrading pip...
python -m pip install --upgrade pip

REM Install requirements
echo 📚 Installing dependencies...
pip install -r requirements.txt

echo.
echo 🎉 Setup complete! Your environment is ready.
echo.
echo To run the app:
echo 1. Activate the environment: venv\Scripts\activate
echo 2. Run the app: streamlit run app.py
echo.
echo Or simply run: run_app.bat
echo.
pause
