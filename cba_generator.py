from io import BytesIO
from datetime import datetime
from openpyxl import Workbook

def generate_cba_xlsx(data: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "CBA Inputs"

    ws["A1"] = "Project"
    ws["B1"] = data.get("project_name", "")

    ws["A3"] = "Options"
    for i, opt in enumerate(data.get("options", []), start=3):
        ws.cell(row=i, column=2, value=opt)

    ws["A6"] = "Factors"
    row = 7
    for f in data.get("factors", []):
        ws.cell(row=row, column=1, value=f.get("name", ""))
        row += 1

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

def default_filename(project_name: str) -> str:
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    safe = "".join(ch for ch in (project_name or "CBA") if ch.isalnum() or ch in (" ", "_", "-")).strip()
    safe = safe.replace(" ", "_") or "CBA"
    return f"{safe}_{stamp}.xlsx"
