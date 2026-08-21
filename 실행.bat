@echo off
cd /d "%~dp0"
echo ================================================
echo   Excel Accumulate DB - Web App
echo   Opening browser at http://localhost:8501
echo   (Close this window to stop the app)
echo ================================================
py -m streamlit run app.py
echo.
echo App stopped. Press any key to close.
pause >nul
