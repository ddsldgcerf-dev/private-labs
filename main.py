from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

import sqlite3
import os
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment


BASE = Path(__file__).resolve().parent
DB = BASE / "labs.db"

templates = Jinja2Templates(directory=str(BASE / "templates"))

app = FastAPI(title="قاعدة بيانات المخابر الخاصة")

# =========================================================
# الجلسات
# =========================================================

SESSION_SECRET = os.getenv("SESSION_SECRET", "CHANGE_THIS_SECRET_IN_RENDER")

app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    session_cookie="labs_session",
    max_age=60 * 60 * 8,  # 8 ساعات
    same_site="lax",
    https_only=True,
)


# =========================================================
# التوزيع الرسمي للمديريات
# =========================================================

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
    "المديرية الجهوية",
    "المديرية الولائية",
    "اسم المخبر",
    "رقم السجل التجاري",
    "رقم رخصة الاستغلال",
    "تاريخ توقيع الرخصة",
    "عنوان النشاط",
    "صاحب الرخصة",
    "المدير التقني",
    "وضعية المخبر",
    "التحاليل الفيزيائية والكيميائية",
    "التحاليل الميكروبيولوجية",
    "تاريخ آخر تفتيش",
    "ملاحظات",
]


# =========================================================
# قاعدة البيانات
# =========================================================

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = db()

    con.execute("""
        CREATE TABLE IF NOT EXISTS labs (
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
        )
    """)

    con.commit()
    con.close()


init_db()


# =========================================================
# المستخدمون والصلاحيات
# =========================================================

def load_json_env(name):
    value = os.getenv(name, "").strip()

    if not value:
        return {}

    try:
        data = json.loads(value)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def get_users():
    """
    الحسابات تأتي من Render Environment Variables.

    ADMIN_PASSWORD
    REGION_PASSWORDS
    WILAYA_PASSWORDS
    """

    users = []

    admin_password = os.getenv("ADMIN_PASSWORD", "").strip()

    if admin_password:
        users.append({
            "username": "admin",
            "password": admin_password,
            "role": "admin",
            "name": "الإدارة المركزية",
        })

    region_passwords = load_json_env("REGION_PASSWORDS")

    for region, password in region_passwords.items():
        if region in REGIONS and str(password).strip():
            users.append({
                "username": f"region:{region}",
                "password": str(password),
                "role": "region",
                "name": region,
                "region": region,
            })

    wilaya_passwords = load_json_env("WILAYA_PASSWORDS")

    all_wilayas = {
        wilaya
        for values in REGIONS.values()
        for wilaya in values
    }

    for wilaya, password in wilaya_passwords.items():
        if wilaya in all_wilayas and str(password).strip():
            users.append({
                "username": f"wilaya:{wilaya}",
                "password": str(password),
                "role": "wilaya",
                "name": wilaya,
                "wilaya": wilaya,
            })

    return users


def authenticate(username, password):
    for user in get_users():
        if user["username"] == username and user["password"] == password:
            return user

    return None


def current_user(request: Request):
    return request.session.get("user")


def require_login(request: Request):
    user = current_user(request)

    if not user:
        return None

    return user


def visible_labs(user):
    con = db()

    if user["role"] == "admin":
        rows = con.execute("""
            SELECT * FROM labs
            ORDER BY region, wilaya, name
        """).fetchall()

    elif user["role"] == "region":
        rows = con.execute("""
            SELECT * FROM labs
            WHERE region = ?
            ORDER BY wilaya, name
        """, (user["region"],)).fetchall()

    elif user["role"] == "wilaya":
        rows = con.execute("""
            SELECT * FROM labs
            WHERE wilaya = ?
            ORDER BY name
        """, (user["wilaya"],)).fetchall()

    else:
        rows = []

    con.close()

    return rows


# =========================================================
# الصفحة العامة
# =========================================================

@app.get("/", response_class=HTMLResponse)
def form_page(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "request": request,
            "regions": REGIONS,
        },
    )


# =========================================================
# إضافة مخبر
# =========================================================

@app.post("/add")
def add_lab(
    region: str = Form(...),
    wilaya: str = Form(...),
    name: str = Form(...),
    rc: str = Form(""),
    license_no: str = Form(""),
    license_date: str = Form(""),
    address: str = Form(""),
    owner: str = Form(""),
    technical_manager: str = Form(""),
    status: str = Form("نشط"),
    physico_chemical: str = Form(""),
    microbiological: str = Form(""),
    last_inspection: str = Form(""),
    notes: str = Form("")
):

    if region not in REGIONS or wilaya not in REGIONS[region]:
        raise HTTPException(
            status_code=400,
            detail="المديرية الولائية لا تتبع المديرية الجهوية المختارة"
        )

    con = db()

    con.execute("""
        INSERT INTO labs
        (
            region,
            wilaya,
            name,
            rc,
            license_no,
            license_date,
            address,
            owner,
            technical_manager,
            status,
            physico_chemical,
            microbiological,
            last_inspection,
            notes
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        region,
        wilaya,
        name,
        rc,
        license_no,
        license_date,
        address,
        owner,
        technical_manager,
        status,
        physico_chemical,
        microbiological,
        last_inspection,
        notes,
    ))

    con.commit()
    con.close()

    return RedirectResponse("/?sent=1", status_code=303)


# =========================================================
# تسجيل الدخول
# =========================================================

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):

    if current_user(request):
        return RedirectResponse("/dashboard", status_code=303)

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "error": None,
        },
    )


@app.post("/login")
def login(
    request: Request,
    username: str = Form(...),
    password: str = Form(...)
):

    user = authenticate(username.strip(), password)

    if not user:
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "request": request,
                "error": "اسم المستخدم أو كلمة المرور غير صحيحة.",
            },
            status_code=401,
        )

    request.session.clear()

    request.session["user"] = {
        "username": user["username"],
        "role": user["role"],
        "name": user["name"],
    }

    if user["role"] == "region":
        request.session["user"]["region"] = user["region"]

    elif user["role"] == "wilaya":
        request.session["user"]["wilaya"] = user["wilaya"]

    return RedirectResponse("/dashboard", status_code=303)


# =========================================================
# تسجيل الخروج
# =========================================================

@app.get("/logout")
def logout(request: Request):

    request.session.clear()

    return RedirectResponse("/login", status_code=303)


# =========================================================
# لوحة البيانات حسب الصلاحية
# =========================================================

@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):

    user = require_login(request)

    if not user:
        return RedirectResponse("/login", status_code=303)

    labs = visible_labs(user)

    con = db()

    if user["role"] == "admin":
        count = con.execute(
            "SELECT COUNT(*) FROM labs"
        ).fetchone()[0]

    elif user["role"] == "region":
        count = con.execute(
            "SELECT COUNT(*) FROM labs WHERE region = ?",
            (user["region"],)
        ).fetchone()[0]

    else:
        count = con.execute(
            "SELECT COUNT(*) FROM labs WHERE wilaya = ?",
            (user["wilaya"],)
        ).fetchone()[0]

    con.close()

    return templates.TemplateResponse(
        request=request,
        name="admin.html",
        context={
            "request": request,
            "labs": labs,
            "count": count,
            "user": user,
        },
    )


# =========================================================
# الإبقاء على /admin كعنوان مباشر
# =========================================================

@app.get("/admin")
def admin_redirect(request: Request):

    return RedirectResponse("/dashboard", status_code=303)


# =========================================================
# حذف مخبر
# =========================================================

@app.post("/admin/delete/{lab_id}")
def delete_lab(request: Request, lab_id: int):

    user = require_login(request)

    if not user or user["role"] != "admin":
        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    con = db()

    con.execute(
        "DELETE FROM labs WHERE id = ?",
        (lab_id,)
    )

    con.commit()
    con.close()

    return RedirectResponse("/dashboard", status_code=303)


# =========================================================
# تصدير Excel
# =========================================================

@app.get("/admin/export")
def export_excel(request: Request):

    user = require_login(request)

    if not user:
        return RedirectResponse("/login", status_code=303)

    if user["role"] == "admin":
        rows = visible_labs(user)
    else:
        rows = visible_labs(user)

    wb = Workbook()
    ws = wb.active
    ws.title = "قاعدة المخابر"

    ws.sheet_view.rightToLeft = True

    ws.append(HEADERS)

    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    keys = [
        "region",
        "wilaya",
        "name",
        "rc",
        "license_no",
        "license_date",
        "address",
        "owner",
        "technical_manager",
        "status",
        "physico_chemical",
        "microbiological",
        "last_inspection",
        "notes",
    ]

    for r in rows:
        ws.append([r[k] for k in keys])

    for col in ws.columns:

        max_len = max(
            len(str(c.value or ""))
            for c in col
        )

        ws.column_dimensions[
            col[0].column_letter
        ].width = min(
            max(max_len + 3, 12),
            45
        )

    filename = "قاعدة_المخابر_الخاصة.xlsx"

    path = BASE / filename

    wb.save(path)

    return FileResponse(
        path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )