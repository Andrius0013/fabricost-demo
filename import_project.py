import re
from pathlib import Path
from collections import defaultdict

PAGE_LIMIT = 6

def safe_item(profilis, šaltinis, tipas="PDF", needs_review=True):
    return {
        "pozicija": "",
        "profilis": str(profilis)[:150],
        "medžiaga": "",
        "ilgis": "",
        "kiekis": "",
        "masė": "",
        "dažomas_plotas": "",
        "šaltinis": šaltinis,
        "tipas": tipas,
        "needs_review": needs_review
    }

def extract_from_pdf(path):
    items = []
    project_name = None
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for i in range(min(len(pdf.pages), PAGE_LIMIT)):
                try:
                    text = pdf.pages[i].extract_text() or ""
                    if not project_name and re.search(r'(SANDĖLIO|STOGO|KONSTRUKCIJ|TDP|SK|JUTA)', text, re.I):
                        project_name = "02. SANDĖLIO PASTATAS – Stogo konstrukcijos"
                    for line in text.splitlines():
                        line = line.strip()
                        if len(line) > 10 and re.search(r'\b(PL|IPE|HEA|HEB|UPN|L|RHS|x\d+)', line, re.I):
                            items.append(safe_item(line, f"{Path(path).name} | psl. {i+1}", "PDF"))
                except Exception:
                    continue
    except Exception as e:
        items.append(safe_item(f"PDF klaida: {str(e)[:80]}", Path(path).name, "Klaida"))
    return items, project_name

def extract_from_excel(path):
    items = []
    try:
        import pandas as pd
        xls = pd.ExcelFile(path)
        for sheet in xls.sheet_names[:3]:
            df = pd.read_excel(path, sheet_name=sheet, header=None, nrows=30)
            for _, row in df.iterrows():
                vals = [str(v) for v in row if pd.notna(v) and str(v).strip()]
                if vals:
                    line = " | ".join(vals)[:150]
                    if re.search(r'(PL|IPE|HEA|HEB|Poz|Profil|Kiek)', line, re.I):
                        items.append(safe_item(line, f"{Path(path).name} | {sheet}", "Excel", False))
    except Exception as e:
        items.append(safe_item(f"Excel klaida: {str(e)[:80]}", Path(path).name, "Klaida"))
    return items

def extract_from_word(path):
    items = []
    try:
        from docx import Document
        doc = Document(path)
        for para in doc.paragraphs[:40]:
            text = para.text.strip()
            if len(text) > 10 and re.search(r'\b(PL|IPE|HEA|HEB|UPN|Poz|Profil)', text, re.I):
                items.append(safe_item(text, Path(path).name, "Word"))
    except Exception as e:
        items.append(safe_item(f"Word klaida: {str(e)[:80]}", Path(path).name, "Klaida"))
    return items

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
        elif ext in (".xlsx", ".xls"):
            all_items.extend(extract_from_excel(path))
        elif ext in (".docx", ".doc"):
            all_items.extend(extract_from_word(path))
        else:
            all_items.append(safe_item(f"Formatas {ext} kol kas ribotas", name, "Kitas"))

    # Konfliktų žymėjimas
    seen = defaultdict(list)
    for it in all_items:
        key = (it["profilis"][:80], it.get("kiekis", ""))
        seen[key].append(it)
    for group in seen.values():
        if len(group) > 1:
            for it in group:
                it["needs_review"] = True

    return {
        "projektas": project_name or "Nenustatytas projektas",
        "medžiagos": all_items[:400],
        "total": len(all_items),
        "operacijos": ["Suvirinimas", "Dažymas"] if all_items else [],
        "failai": files_processed
    }
