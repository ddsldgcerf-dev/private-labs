from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from starlette.responses import Response

import sqlite3
import os
import json
import hmac
import hashlib
import secrets
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment


# =========================================================
# BASIC CONFIGURATION
# =========================================================

BASE = Path(__file__).resolve().parent
DB = BASE / "labs.db"

templates = Jinja2Templates(
    directory=str(BASE / "templates")
)

app = FastAPI(
    title="قاعدة بيانات المخابر الخاصة"
)


# =========================================================
# REGIONS / WILAYAS
# =========================================================

REGIONS = {

    "الجزائر": [
        "الجزائر",
        "بومرداس",
        "تيبازة"
    ],

    "البليدة": [
        "البليدة",
        "البويرة",
        "تيزي وزو",
        "المدية",
        "عين الدفلى",
        "الجلفة"
    ],

    "عنابة": [
        "عنابة",
        "سكيكدة",
        "سوق أهراس",
        "الطارف",
        "قالمة"
    ],

    "وهران": [
        "وهران",
        "تلمسان",
        "سيدي بلعباس",
        "عين تموشنت",
        "مستغانم"
    ],

    "بشار": [
        "بشار",
        "النعامة",
        "تندوف",
        "أدرار",
        "البيض",
        "بني عباس",
        "تيميمون",
        "برج باجي مختار"
    ],

    "سطيف": [
        "سطيف",
        "جيجل",
        "برج بوعريريج",
        "المسيلة",
        "ميلة",
        "بجاية"
    ],

    "باتنة": [
        "باتنة",
        "قسنطينة",
        "أم البواقي",
        "تبسة",
        "خنشلة",
        "بسكرة",
        "أولاد جلال"
    ],

    "سعيدة": [
        "سعيدة",
        "تيارت",
        "غليزان",
        "الشلف",
        "تيسمسيلت",
        "معسكر"
    ],

    "ورقلة": [
        "ورقلة",
        "غرداية",
        "الوادي",
        "إليزي",
        "تمنراست",
        "الأغواط",
        "تقرت",
        "المغير",
        "المنيعة",
        "جانت",
        "عين صالح",
        "عين قزام"
    ],
}


ALL_WILAYAS = {
    wilaya: region
    for region, wilayas in REGIONS.items()
    for wilaya in wilayas
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
    "اختصاص النشاط",
    "عائلات المنتوجات",
    "التحاليل الفيزيائية والكيميائية",
    "التحاليل الميكروبيولوجية",
    "تاريخ آخر تفتيش",
    "ملاحظات",
]


# =========================================================
# DATABASE
# =========================================================

def db():

    con = sqlite3.connect(DB)

    con.row_factory = sqlite3.Row

    return con


def column_exists(con, table, column):

    rows = con.execute(
        f"PRAGMA table_info({table})"
    ).fetchall()

    return any(
        row["name"] == column
        for row in rows
    )


def init_db():

    con = db()

    # -----------------------------------------------------
    # LABS
    # -----------------------------------------------------

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

            activity_specialty TEXT,

            product_families TEXT,

            physico_chemical TEXT,

            microbiological TEXT,

            last_inspection TEXT,

            notes TEXT

        )
    """)

    # -----------------------------------------------------
    # SAFE UPGRADE OF OLD DATABASE
    # -----------------------------------------------------

    if not column_exists(
        con,
        "labs",
        "activity_specialty"
    ):

        con.execute(
            """
            ALTER TABLE labs
            ADD COLUMN activity_specialty TEXT
            """
        )

    if not column_exists(
        con,
        "labs",
        "product_families"
    ):

        con.execute(
            """
            ALTER TABLE labs
            ADD COLUMN product_families TEXT
            """
        )

    # -----------------------------------------------------
    # ACCOUNTS
    # -----------------------------------------------------

    con.execute("""
        CREATE TABLE IF NOT EXISTS accounts (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            username TEXT UNIQUE NOT NULL,

            password_hash TEXT NOT NULL,

            role TEXT NOT NULL,

            region TEXT,

            wilaya TEXT,

            active INTEGER NOT NULL DEFAULT 1,

            created_at TEXT DEFAULT CURRENT_TIMESTAMP

        )
    """)

    con.commit()

    con.close()


init_db()


# =========================================================
# PASSWORD HASHING
# =========================================================

def hash_password(password: str) -> str:

    salt = secrets.token_bytes(16)

    key = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        200_000,
    )

    return (
        "pbkdf2_sha256$"
        + salt.hex()
        + "$"
        + key.hex()
    )


def verify_password(
    password: str,
    stored: str
) -> bool:

    try:

        algorithm, salt_hex, key_hex = stored.split("$")

        if algorithm != "pbkdf2_sha256":
            return False

        salt = bytes.fromhex(salt_hex)

        key = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            200_000,
        )

        return hmac.compare_digest(
            key.hex(),
            key_hex
        )

    except Exception:

        return False


# =========================================================
# GENERATE INITIAL PASSWORD
# =========================================================

def generate_initial_password():

    return secrets.token_urlsafe(9)


# =========================================================
# SESSION
# =========================================================

SESSION_SECRET = os.getenv(
    "SESSION_SECRET",
    "CHANGE_THIS_SESSION_SECRET"
)


def sign_session(data: dict) -> str:

    payload = json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":")
    ).encode()

    encoded = payload.hex()

    signature = hmac.new(
        SESSION_SECRET.encode(),
        encoded.encode(),
        hashlib.sha256
    ).hexdigest()

    return encoded + "." + signature


def read_session(request: Request):

    token = request.cookies.get("session")

    if not token or "." not in token:

        return None

    encoded, signature = token.rsplit(
        ".",
        1
    )

    expected = hmac.new(
        SESSION_SECRET.encode(),
        encoded.encode(),
        hashlib.sha256
    ).hexdigest()

    if not hmac.compare_digest(
        signature,
        expected
    ):

        return None

    try:

        data = bytes.fromhex(
            encoded
        ).decode()

        return json.loads(data)

    except Exception:

        return None


def set_session(
    response: Response,
    account: dict
):

    token = sign_session(account)

    response.set_cookie(
        key="session",
        value=token,
        httponly=True,
        secure=True,
        samesite="lax",
        max_age=60 * 60 * 12,
    )


def clear_session(
    response: Response
):

    response.delete_cookie(
        "session"
    )


# =========================================================
# AUTHORIZATION
# =========================================================

def current_user(
    request: Request
):

    return read_session(request)


def require_login(
    request: Request
):

    return current_user(request)


def region_allows(
    user,
    region
):

    if user["role"] == "admin":

        return True

    if user["role"] == "region":

        return region == user["region"]

    if user["role"] == "wilaya":

        return (
            region
            == ALL_WILAYAS.get(
                user["wilaya"]
            )
        )

    return False


def wilaya_allows(
    user,
    wilaya
):

    if user["role"] == "admin":

        return True

    if user["role"] == "region":

        return (
            wilaya
            in REGIONS.get(
                user["region"],
                []
            )
        )

    if user["role"] == "wilaya":

        return wilaya == user["wilaya"]

    return False


# =========================================================
# FILTERED LABS
# =========================================================

def get_filtered_labs(
    user,
    region=None,
    wilaya=None,
    activity=None,
    products=None,
    search=None,
):

    con = db()

    query = """
        SELECT *
        FROM labs
        WHERE 1=1
    """

    params = []

    # -----------------------------------------------------
    # SECURITY SCOPE
    # -----------------------------------------------------

    if user["role"] == "region":

        query += " AND region = ?"

        params.append(
            user["region"]
        )

    elif user["role"] == "wilaya":

        query += " AND wilaya = ?"

        params.append(
            user["wilaya"]
        )

    # -----------------------------------------------------
    # ADMIN FILTERS
    # -----------------------------------------------------

    if (
        region
        and user["role"] == "admin"
    ):

        query += " AND region = ?"

        params.append(region)

    # -----------------------------------------------------
    # WILAYA FILTER
    # -----------------------------------------------------

    if wilaya:

        if wilaya_allows(
            user,
            wilaya
        ):

            query += " AND wilaya = ?"

            params.append(wilaya)

        else:

            query += " AND 1 = 0"

    # -----------------------------------------------------
    # ACTIVITY
    # -----------------------------------------------------

    if activity:

        query += """
            AND LOWER(
                COALESCE(
                    activity_specialty,
                    ''
                )
            )
            LIKE LOWER(?)
        """

        params.append(
            "%" + activity + "%"
        )

    # -----------------------------------------------------
    # PRODUCTS
    # -----------------------------------------------------

    if products:

        query += """
            AND LOWER(
                COALESCE(
                    product_families,
                    ''
                )
            )
            LIKE LOWER(?)
        """

        params.append(
            "%" + products + "%"
        )

    # -----------------------------------------------------
    # SEARCH
    # -----------------------------------------------------

    if search:

        query += """
            AND (
                LOWER(name)
                LIKE LOWER(?)

                OR LOWER(owner)
                LIKE LOWER(?)

                OR LOWER(rc)
                LIKE LOWER(?)

                OR LOWER(address)
                LIKE LOWER(?)
            )
        """

        value = "%" + search + "%"

        params.extend([
            value,
            value,
            value,
            value,
        ])

    query += """
        ORDER BY
            region,
            wilaya,
            name
    """

    rows = con.execute(
        query,
        params
    ).fetchall()

    # -----------------------------------------------------
    # FILTER VALUES
    # -----------------------------------------------------

    activity_values = [
        row[0]
        for row in con.execute("""
            SELECT DISTINCT
                activity_specialty
            FROM labs
            WHERE activity_specialty
                IS NOT NULL
              AND TRIM(
                    activity_specialty
                  ) != ''
            ORDER BY
                activity_specialty
        """).fetchall()
    ]

    product_values = [
        row[0]
        for row in con.execute("""
            SELECT DISTINCT
                product_families
            FROM labs
            WHERE product_families
                IS NOT NULL
              AND TRIM(
                    product_families
                  ) != ''
            ORDER BY
                product_families
        """).fetchall()
    ]

    con.close()

    return (
        rows,
        activity_values,
        product_values
    )


# =========================================================
# LOGIN PAGE
# =========================================================

@app.get(
    "/",
    response_class=HTMLResponse
)
def login_page(
    request: Request
):

    user = current_user(request)

    if user:

        if user["role"] == "admin":

            return RedirectResponse(
                "/admin",
                status_code=303
            )

        if user["role"] == "region":

            return RedirectResponse(
                "/region",
                status_code=303
            )

        if user["role"] == "wilaya":

            return RedirectResponse(
                "/wilaya",
                status_code=303
            )

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "error": None,
        },
    )


# =========================================================
# LOGIN
# =========================================================

@app.post("/login")
def login(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
):

    username = username.strip()

    # -----------------------------------------------------
    # ADMIN
    # -----------------------------------------------------

    admin_password = os.getenv(
        "ADMIN_PASSWORD",
        ""
    )

    if (
        username == "admin"
        and admin_password
    ):

        if hmac.compare_digest(
            password,
            admin_password
        ):

            response = RedirectResponse(
                "/admin",
                status_code=303
            )

            set_session(
                response,
                {
                    "username": "admin",
                    "role": "admin",
                    "region": None,
                    "wilaya": None,
                }
            )

            return response

    # -----------------------------------------------------
    # ACCOUNTS
    # -----------------------------------------------------

    con = db()

    account = con.execute(
        """
        SELECT *
        FROM accounts
        WHERE username = ?
          AND active = 1
        """,
        (username,)
    ).fetchone()

    con.close()

    if account:

        if verify_password(
            password,
            account["password_hash"]
        ):

            destination = (
                "/region"
                if account["role"] == "region"
                else "/wilaya"
            )

            response = RedirectResponse(
                destination,
                status_code=303
            )

            set_session(
                response,
                {
                    "username": account["username"],
                    "role": account["role"],
                    "region": account["region"],
                    "wilaya": account["wilaya"],
                    "account_id": account["id"],
                }
            )

            return response

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "error":
                "اسم المستخدم أو كلمة المرور غير صحيحة",
        },
        status_code=401,
    )


# =========================================================
# LOGOUT
# =========================================================

@app.get("/logout")
def logout():

    response = RedirectResponse(
        "/",
        status_code=303
    )

    clear_session(response)

    return response


# =========================================================
# COMMON FILTER DATA
# =========================================================

def filter_context(
    user,
    request
):

    region_filter = (
        request.query_params.get(
            "region"
        )
    )

    wilaya_filter = (
        request.query_params.get(
            "wilaya"
        )
    )

    activity = (
        request.query_params.get(
            "activity"
        )
    )

    products = (
        request.query_params.get(
            "products"
        )
    )

    search = (
        request.query_params.get(
            "search"
        )
    )

    rows, activity_values, product_values = (
        get_filtered_labs(
            user,
            region_filter,
            wilaya_filter,
            activity,
            products,
            search,
        )
    )

    # -----------------------------------------------------
    # AVAILABLE REGIONS / WILAYAS
    # -----------------------------------------------------

    if user["role"] == "admin":

        available_regions = list(
            REGIONS.keys()
        )

        available_wilayas = [
            w
            for ws in REGIONS.values()
            for w in ws
        ]

    elif user["role"] == "region":

        available_regions = [
            user["region"]
        ]

        available_wilayas = REGIONS.get(
            user["region"],
            []
        )

    else:

        available_regions = [
            ALL_WILAYAS.get(
                user["wilaya"]
            )
        ]

        available_wilayas = [
            user["wilaya"]
        ]

    return {

        "rows": rows,

        "count": len(rows),

        "activity_values":
            activity_values,

        "product_values":
            product_values,

        "available_regions":
            available_regions,

        "available_wilayas":
            available_wilayas,

        "selected_region":
            region_filter or "",

        "selected_wilaya":
            wilaya_filter or "",

        "selected_activity":
            activity or "",

        "selected_products":
            products or "",

        "selected_search":
            search or "",

        "user": user,
    }


# =========================================================
# ADMIN PAGE
# =========================================================

@app.get(
    "/admin",
    response_class=HTMLResponse
)
def admin_page(
    request: Request
):

    user = require_login(
        request
    )

    if not user:

        return RedirectResponse(
            "/",
            status_code=303
        )

    if user["role"] != "admin":

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    context = filter_context(
        user,
        request
    )

    return templates.TemplateResponse(
        request=request,
        name="admin.html",
        context={
            "request": request,
            **context,
        },
    )


# =========================================================
# CREATE BASIC ACCOUNTS AUTOMATICALLY
# =========================================================

@app.post(
    "/admin/accounts/bootstrap"
)
def bootstrap_accounts(
    request: Request
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    con = db()

    created_accounts = []

    existing_accounts = []

    # -----------------------------------------------------
    # REGIONAL ACCOUNTS
    # -----------------------------------------------------

    for index, region in enumerate(
        REGIONS.keys(),
        start=1
    ):

        username = f"dr_{index:02d}"

        existing = con.execute(
            """
            SELECT id, username
            FROM accounts
            WHERE username = ?
            """,
            (username,)
        ).fetchone()

        if existing:

            existing_accounts.append({
                "username": username,
                "role": "region",
                "region": region,
                "wilaya": None,
            })

            continue

        password = generate_initial_password()

        con.execute(
            """
            INSERT INTO accounts
            (
                username,
                password_hash,
                role,
                region,
                wilaya
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                username,
                hash_password(password),
                "region",
                region,
                None,
            )
        )

        created_accounts.append({
            "username": username,
            "password": password,
            "role": "region",
            "region": region,
            "wilaya": None,
        })

    # -----------------------------------------------------
    # WILAYA ACCOUNTS
    # -----------------------------------------------------

    wilaya_index = 1

    for region, wilayas in REGIONS.items():

        for wilaya in wilayas:

            username = (
                f"dw_{wilaya_index:02d}"
            )

            wilaya_index += 1

            existing = con.execute(
                """
                SELECT id, username
                FROM accounts
                WHERE username = ?
                """,
                (username,)
            ).fetchone()

            if existing:

                existing_accounts.append({
                    "username": username,
                    "role": "wilaya",
                    "region": region,
                    "wilaya": wilaya,
                })

                continue

            password = (
                generate_initial_password()
            )

            con.execute(
                """
                INSERT INTO accounts
                (
                    username,
                    password_hash,
                    role,
                    region,
                    wilaya
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    username,
                    hash_password(password),
                    "wilaya",
                    region,
                    wilaya,
                )
            )

            created_accounts.append({
                "username": username,
                "password": password,
                "role": "wilaya",
                "region": region,
                "wilaya": wilaya,
            })

    con.commit()

    con.close()

    return {
        "success": True,
        "created": created_accounts,
        "existing": existing_accounts,
        "created_count":
            len(created_accounts),
        "existing_count":
            len(existing_accounts),
        "message":
            "تم إنشاء الحسابات الأساسية بنجاح"
    }


# =========================================================
# ADD ACCOUNT MANUALLY
# =========================================================

@app.post(
    "/admin/accounts/add"
)
def add_account(
    request: Request,

    username: str = Form(...),

    password: str = Form(...),

    role: str = Form(...),

    region: str = Form(""),

    wilaya: str = Form(""),
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    username = username.strip()

    if role not in [
        "region",
        "wilaya"
    ]:

        raise HTTPException(
            status_code=400,
            detail="نوع الحساب غير صحيح"
        )

    # -----------------------------------------------------
    # REGION ACCOUNT
    # -----------------------------------------------------

    if role == "region":

        if region not in REGIONS:

            raise HTTPException(
                status_code=400,
                detail=
                "المديرية الجهوية غير صحيحة"
            )

        wilaya = None

    # -----------------------------------------------------
    # WILAYA ACCOUNT
    # -----------------------------------------------------

    if role == "wilaya":

        if wilaya not in ALL_WILAYAS:

            raise HTTPException(
                status_code=400,
                detail=
                "المديرية الولائية غير صحيحة"
            )

        region = ALL_WILAYAS[
            wilaya
        ]

    if len(password) < 6:

        raise HTTPException(
            status_code=400,
            detail=
            "كلمة المرور يجب أن تكون 6 أحرف على الأقل"
        )

    con = db()

    try:

        con.execute(
            """
            INSERT INTO accounts
            (
                username,
                password_hash,
                role,
                region,
                wilaya
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                username,
                hash_password(password),
                role,
                region,
                wilaya,
            )
        )

        con.commit()

    except sqlite3.IntegrityError:

        con.close()

        raise HTTPException(
            status_code=400,
            detail=
            "اسم المستخدم موجود مسبقًا"
        )

    con.close()

    return RedirectResponse(
        "/admin",
        status_code=303
    )


# =========================================================
# CHANGE PASSWORD
# =========================================================

@app.post(
    "/admin/accounts/{account_id}/password"
)
def change_account_password(
    account_id: int,

    request: Request,

    password: str = Form(...),
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    if len(password) < 6:

        raise HTTPException(
            status_code=400,
            detail=
            "كلمة المرور يجب أن تكون 6 أحرف على الأقل"
        )

    con = db()

    con.execute(
        """
        UPDATE accounts
        SET password_hash = ?
        WHERE id = ?
        """,
        (
            hash_password(password),
            account_id,
        )
    )

    con.commit()

    con.close()

    return RedirectResponse(
        "/admin",
        status_code=303
    )


# =========================================================
# ENABLE / DISABLE ACCOUNT
# =========================================================

@app.post(
    "/admin/accounts/{account_id}/toggle"
)
def toggle_account(
    account_id: int,

    request: Request,
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    con = db()

    con.execute(
        """
        UPDATE accounts
        SET active =
            CASE
                WHEN active = 1
                THEN 0
                ELSE 1
            END
        WHERE id = ?
        """,
        (account_id,)
    )

    con.commit()

    con.close()

    return RedirectResponse(
        "/admin",
        status_code=303
    )


# =========================================================
# DELETE ACCOUNT
# =========================================================

@app.post(
    "/admin/accounts/{account_id}/delete"
)
def delete_account(
    account_id: int,

    request: Request,
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    con = db()

    con.execute(
        """
        DELETE FROM accounts
        WHERE id = ?
        """,
        (account_id,)
    )

    con.commit()

    con.close()

    return RedirectResponse(
        "/admin",
        status_code=303
    )


# =========================================================
# ACCOUNTS LIST
# =========================================================

@app.get(
    "/admin/accounts"
)
def accounts_list(
    request: Request
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    con = db()

    accounts = con.execute(
        """
        SELECT
            id,
            username,
            role,
            region,
            wilaya,
            active,
            created_at
        FROM accounts
        ORDER BY
            role,
            region,
            wilaya,
            username
        """
    ).fetchall()

    con.close()

    return {
        "accounts": [
            dict(row)
            for row in accounts
        ]
    }


# =========================================================
# REGIONAL DASHBOARD
# =========================================================

@app.get(
    "/region",
    response_class=HTMLResponse
)
def region_page(
    request: Request
):

    user = require_login(
        request
    )

    if not user:

        return RedirectResponse(
            "/",
            status_code=303
        )

    if user["role"] != "region":

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    context = filter_context(
        user,
        request
    )

    return templates.TemplateResponse(
        request=request,
        name="region.html",
        context={
            "request": request,
            **context,
        },
    )


# =========================================================
# WILAYA PAGE
# =========================================================

@app.get(
    "/wilaya",
    response_class=HTMLResponse
)
def wilaya_page(
    request: Request
):

    user = require_login(
        request
    )

    if not user:

        return RedirectResponse(
            "/",
            status_code=303
        )

    if user["role"] != "wilaya":

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    context = filter_context(
        user,
        request
    )

    return templates.TemplateResponse(
        request=request,
        name="wilaya.html",
        context={
            "request": request,
            **context,
        },
    )


# =========================================================
# ADD LAB
# =========================================================

@app.post("/add")
def add_lab(
    request: Request,

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

    activity_specialty: str = Form(""),
    product_families: str = Form(""),

    physico_chemical: str = Form(""),
    microbiological: str = Form(""),

    last_inspection: str = Form(""),
    notes: str = Form(""),
):

    user = require_login(
        request
    )

    if not user:

        return RedirectResponse(
            "/",
            status_code=303
        )

    # -----------------------------------------------------
    # ONLY WILAYA CAN SUBMIT
    # -----------------------------------------------------

    if user["role"] != "wilaya":

        raise HTTPException(
            status_code=403,
            detail=
            "الإدارة المركزية والمديرية الجهوية لا تضيف البيانات من هذه الصفحة"
        )

    # -----------------------------------------------------
    # FORCE USER'S OWN SCOPE
    # -----------------------------------------------------

    region = user["region"]

    wilaya = user["wilaya"]

    # -----------------------------------------------------
    # VALIDATION
    # -----------------------------------------------------

    if region not in REGIONS:

        raise HTTPException(
            status_code=400,
            detail=
            "المديرية الجهوية غير صحيحة"
        )

    if wilaya not in REGIONS[region]:

        raise HTTPException(
            status_code=400,
            detail=
            "المديرية الولائية لا تتبع المديرية الجهوية"
        )

    # -----------------------------------------------------
    # INSERT
    # -----------------------------------------------------

    con = db()

    con.execute(
        """
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
            activity_specialty,
            product_families,
            physico_chemical,
            microbiological,
            last_inspection,
            notes
        )
        VALUES
        (
            ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?
        )
        """,
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
            activity_specialty,
            product_families,
            physico_chemical,
            microbiological,
            last_inspection,
            notes,
        )
    )

    con.commit()

    con.close()

    return RedirectResponse(
        "/wilaya?sent=1",
        status_code=303
    )


# =========================================================
# DELETE LAB - ADMIN ONLY
# =========================================================

@app.post(
    "/admin/labs/{lab_id}/delete"
)
def delete_lab(
    lab_id: int,

    request: Request,
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    con = db()

    con.execute(
        """
        DELETE FROM labs
        WHERE id = ?
        """,
        (lab_id,)
    )

    con.commit()

    con.close()

    return RedirectResponse(
        "/admin",
        status_code=303
    )


# =========================================================
# EXPORT EXCEL
# =========================================================

@app.get(
    "/admin/export"
)
def export_excel(
    request: Request
):

    user = require_login(
        request
    )

    if (
        not user
        or user["role"] != "admin"
    ):

        raise HTTPException(
            status_code=403,
            detail="غير مصرح"
        )

    context = filter_context(
        user,
        request
    )

    rows = context["rows"]

    wb = Workbook()

    ws = wb.active

    ws.title = "قاعدة المخابر"

    ws.sheet_view.rightToLeft = True

    ws.append(
        HEADERS
    )

    for cell in ws[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    fields = [

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

        "activity_specialty",

        "product_families",

        "physico_chemical",

        "microbiological",

        "last_inspection",

        "notes",

    ]

    for row in rows:

        ws.append([
            row[field]
            for field in fields
        ])

    for col in ws.columns:

        max_len = max(
            len(
                str(
                    cell.value
                    or ""
                )
            )
            for cell in col
        )

        ws.column_dimensions[
            col[0].column_letter
        ].width = min(
            max(
                max_len + 3,
                12
            ),
            45
        )

    path = (
        BASE
        / "قاعدة_المخابر_الخاصة.xlsx"
    )

    wb.save(path)

    return FileResponse(
        path,

        filename=
        "قاعدة_المخابر_الخاصة.xlsx",

        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )