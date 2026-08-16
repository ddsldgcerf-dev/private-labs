from fastapi import FastAPI, Request, Form, Header, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
import sqlite3
import os
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

BASE = Path(__file__).resolve().parent
DB = BASE / "labs.db"
templates = Jinja2Templates(directory=str(BASE / "templates"))
app = FastAPI(title="قاعدة بيانات المخابر الخاصة")

# التوزيع الرسمي المنشور على موقع وزارة التجارة: 9 مديريات جهوية / 58 ولاية.
REGIONS = {
    "الجزائر": ["الجزائر", "بومرداس", "تيبازة"],
    "البليدة": ["البليدة", "البويرة", "تيزي وزو", "المدية", "عين الدفلى", "الجلفة"],
    "عنابة": ["عنابة", "سكيكدة", "سوق أهراس", "الطارف", "قالمة"],
    "وهران": ["وهران", "تلمسان", "سيدي بلعباس", "عين تموشنت", "مستغانم"],
    "بشار": ["بشار", "النعامة", "تندوف", "أدرار", "البيض", "بني عباس", "تيميمون", "برج باجي مختار"],
    "سطيف": ["سطيف", "جيجل", "برج بوعريريج", "المسيلة", "ميلة", "بجاية"],
    "باتنة": ["باتنة", "قسنطينة", "أم البواقي", "تبسة", "خنشلة", "بسكرة", "أولاد جلال"],
    "سعيدة": ["سعيدة", "تيارت", "غليزان", "الشلف", "تيسمسيلت", "معسكر"],
    "ورقلة": ["ورقلة", "غرداية", "الوادي", "إليزي", "تمنراست", "الأغواط", "تقرت", "المغير", "المنيعة", "جانت", "عين صالح", "عين قزام"],
}

HEADERS = [
    "المديرية الجهوية", "المديرية الولائية", "اسم المخبر", "رقم السجل التجاري",
    "رقم رخصة الاستغلال", "تاريخ توقيع الرخصة", "عنوان النشاط", "صاحب الرخصة",
    "المدير التقني", "وضعية المخبر", "التحاليل الفيزيائية والكيميائية",
    "التحاليل الميكروبيولوجية", "تاريخ آخر تفتيش", "ملاحظات"
]


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


def check_admin(password: str | None):
    expected = os.getenv("ADMIN_PASSWORD", "")
    if not expected or password != expected:
        raise HTTPException(status_code=403, detail="غير مصرح")


@app.get("/", response_class=HTMLResponse)
def form_page(request: Request):
    # الصفحة العامة = الاستمارة فقط. لا نعرض أي سجل أو جدول مركزي.
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={"request": request, "regions": REGIONS},
    )


@app.post("/add")
def add_lab(
    region: str = Form(...), wilaya: str = Form(...), name: str = Form(...),
    rc: str = Form(""), license_no: str = Form(""), license_date: str = Form(""),
    address: str = Form(""), owner: str = Form(""), technical_manager: str = Form(""),
    status: str = Form("نشط"), physico_chemical: str = Form(""),
    microbiological: str = Form(""), last_inspection: str = Form(""), notes: str = Form("")
):
    if region not in REGIONS or wilaya not in REGIONS[region]:
        raise HTTPException(status_code=400, detail="المديرية الولائية لا تتبع المديرية الجهوية المختارة")
    con = db()
    con.execute("""INSERT INTO labs
        (region,wilaya,name,rc,license_no,license_date,address,owner,technical_manager,
         status,physico_chemical,microbiological,last_inspection,notes)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (region,wilaya,name,rc,license_no,license_date,address,owner,technical_manager,
         status,physico_chemical,microbiological,last_inspection,notes))
    con.commit()
    con.close()
    return RedirectResponse("/?sent=1", status_code=303)


@app.get("/admin", response_class=HTMLResponse)
def admin_page(request: Request, password: str | None = None):
    check_admin(password)
    con = db()
    labs = con.execute("SELECT * FROM labs ORDER BY region, wilaya, name").fetchall()
    con.close()
    return templates.TemplateResponse(
        request=request,
        name="admin.html",
        context={"request": request, "labs": labs},
    )


@app.get("/admin/export")
def export_excel(password: str | None = None):
    check_admin(password)
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
        ws.append([r[k] for k in ["region", "wilaya", "name", "rc", "license_no", "license_date", "address", "owner", "technical_manager", "status", "physico_chemical", "microbiological", "last_inspection", "notes"]])
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 45)
    path = BASE / "قاعدة_المخابر_الخاصة.xlsx"
    wb.save(path)
    return FileResponse(path, filename="قاعدة_المخابر_الخاصة.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
