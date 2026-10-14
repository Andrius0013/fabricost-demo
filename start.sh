#!/bin/bash
echo "FabriCost PRO DEMO Web"
echo "====================="
pip install flask pandas openpyxl pdfplumber pytesseract pillow python-docx ezdxf pdf2image --quiet
echo "Paleidžiama: http://localhost:5000"
python3 app.py
