@echo off
cd /d "%~dp0"
echo Installation des dependances...
pip install streamlit pandas python-docx openpyxl -q
echo Lancement de l'application...
streamlit run app.py
pause
