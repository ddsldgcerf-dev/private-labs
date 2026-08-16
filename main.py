from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
import sqlite3
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

BASE = Path(__file__).resolve().parent
DB = BASE / "labs.db"
templates = Jinja2Templates(directory=str(BASE / "templates"))

app = FastAPI(title="قاعدة بيانات المخابر الخاصة")

HEADERS = [
    "المديرية الجهوية", "المديرية الولائية", "اسم المخبر",
    "رقم السجل التجاري", "رقم رخصة الاستغلال", "تاريخ توقيع الرخصة",
    "عنوان النشاط", "صاحب الرخصة", "المدير التقني", "وضعية المخبر",
    "التحاليل الفيزيائية والكيميائية", "التحاليل الميكروبيولوجية",
    "تاريخ آخر تفتيش", "ملاحظات"
]

# عدّل هذه القائمة لاحقًا بالقائمة الرسمية للمديريات الجهوية والولائية.
REGIONS = {
    "المديرية الجهوية للجزائر": ["الجزائر", "البليدة", "تيبازة", "بومرداس"],
    "المديرية الجهوية لوهران": ["وهران", "سيدي بلعباس", "عين تموشنت", "مستغانم"],
    "المديرية الجهوية لقسنطينة": ["قسنطينة", "عنابة", "قالمة", "سكيكدة"],
}

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    con = db()
    con.execute("""CREATE TABLE IF NOT EXISTS labs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        region TEXT NOT NULL,
        wilaya TEXT NOT NULL,
        name TEXT NOT NULL,
        rc TEXT,
        license_no TEXT,
        license_date TEXT,
        address TEXT,
        owner TEXT,
        technical_manager TEXT,
        status TEXT,
        physico_chemical TEXT,
        microbiological TEXT,
        last_inspection TEXT,
        notes TEXT
    )""")
    con.commit()
    con.close()

init_db()

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    con = db()
    labs = con.execute("SELECT * FROM labs ORDER BY region, wilaya, name").fetchall()
    con.close()
    return templates.TemplateResponse(request=request, name="index.html", context={
        "request": request, "labs": labs, "regions": REGIONS}
    )

@app.post("/add")
def add_lab(
    region: str = Form(...), wilaya: str = Form(...), name: str = Form(...),
    rc: str = Form(""), license_no: str = Form(""), license_date: str = Form(""),
    address: str = Form(""), owner: str = Form(""), technical_manager: str = Form(""),
    status: str = Form("نشط"), physico_chemical: str = Form(""),
    microbiological: str = Form(""), last_inspection: str = Form(""),
    notes: str = Form("")
):
    con = db()
    con.execute("""INSERT INTO labs
        (region,wilaya,name,rc,license_no,license_date,address,owner,technical_manager,
         status,physico_chemical,microbiological,last_inspection,notes)
         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (region,wilaya,name,rc,license_no,license_date,address,owner,technical_manager,
         status,physico_chemical,microbiological,last_inspection,notes))
    con.commit()
    con.close()
    from fastapi.responses import RedirectResponse
    return RedirectResponse("/", status_code=303)

@app.get("/export")
def export_excel():
    con = db()
    rows = con.execute("SELECT * FROM labs ORDER BY region, wilaya, name").fetchall()
    con.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "قاعدة المخابر"
    ws.sheet_view.rightToLeft = True
    ws.append(HEADERS)

    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")

    for r in rows:
        ws.append([
            r["region"], r["wilaya"], r["name"], r["rc"], r["license_no"],
            r["license_date"], r["address"], r["owner"], r["technical_manager"],
            r["status"], r["physico_chemical"], r["microbiological"],
            r["last_inspection"], r["notes"]
        ])

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 45)

    path = BASE / "قاعدة_المخابر_الخاصة.xlsx"
    wb.save(path)
    return FileResponse(path, filename="قاعدة_المخابر_الخاصة.xlsx",
                         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
