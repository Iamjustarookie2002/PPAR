@echo off
REM Windows batch file to activate virtual environment and run Streamlit app
call venv\Scripts\activate.bat
streamlit run app.py
