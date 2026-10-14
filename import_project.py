import os
import re
import tempfile
import shutil
from pathlib import Path
from collections import defaultdict

import pdfplumber
import pandas as pd

PAGE_LIMIT = 12  # Labai sumažinta dėl atminties

def extract_from_pdf(path, max_pages=PAGE_LIMIT):
    items = []
    project_name = None
    try:
        with pdfplumber.open(path) as pdf:
            pages_to_read = min(len(pdf.pages), max_pages)
            for i in range(pages_to_read):
                page = pdf.pages[i]
                text = page.extract_text() or ""
                if not project_name:
                    if re.search(r'(SANDĖLIO|STOGO|KONSTRUKCIJ|TDP|SK\.|JUTA)', text, re.I):
                        project_name = "02. SANDĖLIO PASTATAS – Stogo konstrukcijos"
                for line in text.splitlines():
                    line = line.strip()
                    if len(line) < 5:
                        continue
                    if re.search(r'\bx\d+\s+[A-Z0-9*\.]+', line) or re.search(r'\b(PL|IPE|HEA|HEB|UPN|L|RHS)\d', line, re.I):
                        items.append({
                            "pozicija": "",
                            "profilis": line[:100],
                            "medžiaga": "",
                            "ilgis": "",
                            "kiekis": "",
                            "masė": "",
                            "dažomas_plotas": "",
                            "šaltinis": f"{Path(path).name} | psl. {i+1}",
                            "tipas": "PDF",
                            "needs_review": True
                        })
                tables = page.extract_tables() or []
                for table in tables[:3]:  # max 3 lentelės per puslapį
                    for row in table[:20]:  # max 20 eilučių
                        if not row or not any(row):
                            continue
                        row_str = " ".join([str(c or "") for c in row])
                        if re.search(r'(Poz|Profil|Medžiag|Kiek|Ilg|Svor|Mas)', row_str, re.I):
                            continue
                        if any(c and str(c).strip() for c in row):
                            items.append({
                                "pozicija": str(row[0] or "")[:40] if len(row) > 0 else "",
                                "profilis": str(row[1] or "")[:80] if len(row) > 1 else "",
                                "medžiaga": str(row[2] or "")[:40] if len(row) > 2 else "",
                                "ilgis": str(row[3] or "")[:30] if len(row) > 3 else "",
                                "kiekis": str(row[4] or "")[:30] if len(row) > 4 else "",
                                "masė": str(row[5] or "")[:30] if len(row) > 5 else "",
                                "dažomas_plotas": "",
                                "šaltinis": f"{Path(path).name} | psl. {i+1} | lentelė",
                                "tipas": "PDF lentelė",
                                "needs_review": False
                            })
    except Exception as e:
        items.append({
            "pozicija": "",
            "profilis": f"Klaida: {str(e)[:80]}",
            "medžiaga": "",
            "ilgis": "",
            "kiekis": "",
            "masė": "",
            "dažomas_plotas": "",
            "šaltinis": Path(path).name,
            "tipas": "Klaida",
            "needs_review": True
        })
    return items, project_name or "Nenustatytas projektas"

def process_files(file_paths):
    all_items = []
    project_name = None
    files_processed = []
    for path in file_paths:
        name = Path(path).name
        files_processed.append(name)
        ext = Path(path).suffix.lower()
        if ext == ".pdf":
            items, pname = extract_from_pdf(path)
            all_items.extend(items)
            if pname and not project_name:
                project_name = pname
        else:
            all_items.append({
                "pozicija": "",
                "profilis": f"Formatas {ext} ribotai",
                "medžiaga": "",
                "ilgis": "",
                "kiekis": "",
                "masė": "",
                "dažomas_plotas": "",
                "šaltinis": name,
                "tipas": "Kitas",
                "needs_review": True
            })
    # Konfliktų žymėjimas
    seen = defaultdict(list)
    for it in all_items:
        key = (it.get("profilis", "")[:50], it.get("kiekis", ""))
        seen[key].append(it)
    for group in seen.values():
        if len(group) > 1:
            for it in group:
                it["needs_review"]
