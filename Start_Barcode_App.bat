@echo off
echo Starting Jayshakti Farsan Mart Barcode Printer...
echo.
echo Installing requirements (if missing)...
cd /d "%~dp0"
python -m pip install -r requirements.txt >nul 2>&1

echo Starting Web Application...
echo The app will open in your default browser automatically.
echo (Keep this window open while using the app)
python -m streamlit run app.py
pause
