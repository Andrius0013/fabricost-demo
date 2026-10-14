@echo off
chcp 65001 >nul
echo FabriCost PRO DEMO Web
echo =====================
echo.
echo Instaliuojamos priklausomybės (jei reikia)...
pip install flask pandas openpyxl pdfplumber pytesseract pillow python-docx ezdxf pdf2image --quiet
echo.
echo Paleidžiama serveris...
echo Atidarykite naršyklėje: http://localhost:5000
echo.
python app.py
pause
