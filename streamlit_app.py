# streamlit_app.py
# -*- coding: utf-8 -*-
import csv, re, json
from io import BytesIO
from pathlib import Path
from datetime import datetime
import pytz
import streamlit as st
import pandas as pd

# --- Google Sheets (אופציונלי)
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import (
    CellFormat, Color, TextFormat,
    ConditionalFormatRule, BooleanRule, BooleanCondition,
    GridRange, format_cell_range, get_conditional_format_rules
)

# =========================
# הגדרות כלליות
# =========================
st.set_page_config(page_title="שאלון לסטודנטים – תשפ״ו", layout="centered")
st.markdown("""
<style>
:root{ --ink:#0f172a; --muted:#475569; --ring:rgba(99,102,241,.25); --card:rgba(255,255,255,.85);}
html, body, [class*="css"] { font-family: system-ui, "Segoe UI", Arial; }
.stApp, .main, [data-testid="stSidebar"]{ direction:rtl; text-align:right; }
[data-testid="stAppViewContainer"]{
  background:
    radial-gradient(1200px 600px at 8% 8%, #e0f7fa 0%, transparent 65%),
    radial-gradient(1000px 500px at 92% 12%, #ede7f6 0%, transparent 60%),
    radial-gradient(900px 500px at 20% 90%, #fff3e0 0%, transparent 55%);
}
.block-container{ padding-top:1.1rem; }
[data-testid="stForm"]{
  background:var(--card); border:1px solid #e2e8f0; border-radius:16px; padding:18px 20px;
  box-shadow:0 8px 24px rgba(2,6,23,.06);
}
[data-testid="stWidgetLabel"] p{ text-align:right; margin-bottom:.25rem; color:var(--muted); }
[data-testid="stWidgetLabel"] p::after{ content: " :"; }
input, textarea, select{ direction:rtl; text-align:right; }
</style>
""", unsafe_allow_html=True)

# =========================
# נתיבים/סודות + התמדה
# =========================
DATA_DIR   = Path("data"); DATA_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR = DATA_DIR / "backups"; BACKUP_DIR.mkdir(parents=True, exist_ok=True)

CSV_FILE      = DATA_DIR / "שאלון_שיבוץ.csv"
CSV_LOG_FILE  = DATA_DIR / "שאלון_שיבוץ_log.csv"
STATE_FILE    = DATA_DIR / "form_state.json"

ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "rawan_0304")

# תאימות query params
try:
    query_params = st.query_params
except Exception:
    query_params = st.experimental_get_query_params()

def _qp_get(name, default="0"):
    v = query_params.get(name)
    if isinstance(v, list): return (v[0] if v else default)
    return v if v is not None else default

is_admin_mode = _qp_get("admin", "0") == "1"

# =========================
# Google Sheets (אופציונלי)
# =========================
SHEET_ID = None
try:
    SHEET_ID = st.secrets["sheets"]["spreadsheet_id"]
except Exception:
    SHEET_ID = None

scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
sheet = None
if SHEET_ID:
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        gclient = gspread.authorize(creds)
        sheet = gclient.open_by_key(SHEET_ID).sheet1
    except Exception as e:
        st.error(f"⚠ לא ניתן להתחבר ל־Google Sheets: {e}")
        sheet = None

# =========================
# קבועים
# =========================
SITES = [
    "כפר הילדים חורפיש","אנוש כרמיאל","הפוך על הפוך צפת","שירות מבחן לנוער עכו","כלא חרמון",
    "בית חולים זיו","שירותי רווחה קריית שמונה","מרכז יום לגיל השלישי","מועדונית נוער בצפת","מרפאת בריאות הנפש צפת",
]
RANK_COUNT = 3
COLUMNS_ORDER = [
    "תאריך שליחה","שם פרטי","שם משפחה","תעודת זהות","מין","שיוך חברתי",
    "שפת אם","שפות נוספות","טלפון","כתובת","אימייל",
    "שנת לימודים","מסלול לימודים",
    "הכשרה קודמת","הכשרה קודמת מקום ותחום","הכשרה קודמת מדריך ומיקום","הכשרה קודמת בן זוג",
    "תחומים מועדפים","תחום מוביל","בקשה מיוחדת",
    "ממוצע","התאמות","התאמות פרטים",
    "מוטיבציה 1","מוטיבציה 2","מוטיבציה 3",
] + [f"מקום הכשרה {i}" for i in range(1, RANK_COUNT+1)] + [f"דירוג_{s}" for s in SITES] + ["אישור הגעה להכשרה"]

# =========================
# Utils / IO
# =========================
def style_google_sheet(ws):
    header_fmt = CellFormat(backgroundColor=Color(0.6,0.4,0.8),
                            textFormat=TextFormat(bold=True, foregroundColor=Color(1,1,1)),
                            horizontalAlignment='CENTER')
    format_cell_range(ws, "1:1", header_fmt)
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range('A2:Z1000', ws)],
        booleanRule=BooleanRule(condition=BooleanCondition('CUSTOM_FORMULA', ['=ISEVEN(ROW())']),
        format=CellFormat(backgroundColor=Color(0.95,0.95,0.95)))
    )
    rules = get_conditional_format_rules(ws); rules.clear(); rules.append(rule); rules.save()
    id_fmt = CellFormat(horizontalAlignment='CENTER', backgroundColor=Color(0.9,0.9,0.9))
    format_cell_range(ws, "C2:C1000", id_fmt)

def save_master_dataframe(new_row: dict) -> None:
    df_master = pd.DataFrame([new_row])
    if CSV_FILE.exists():
        df_master = pd.concat([pd.read_csv(CSV_FILE), df_master], ignore_index=True)
    df_master.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    (DATA_DIR / "backups").mkdir(exist_ok=True, parents=True)
    df_master.to_csv(BACKUP_DIR / f"שאלון_שיבוץ_{ts}.csv", index=False, encoding="utf-8-sig")

    if sheet:
        try:
            headers = sheet.row_values(1)
            if (not headers) or (headers != COLUMNS_ORDER):
                sheet.clear()
                sheet.append_row(COLUMNS_ORDER, value_input_option="USER_ENTERED")
                style_google_sheet(sheet)
            row_values = [new_row.get(col, "") for col in COLUMNS_ORDER]
            sheet.append_row(row_values, value_input_option="USER_ENTERED")
        except Exception as e:
            st.error(f"❌ לא ניתן לשמור ב־Google Sheets: {e}")

def append_to_log(row_df: pd.DataFrame) -> None:
    file_exists = CSV_LOG_FILE.exists()
    row_df.to_csv(CSV_LOG_FILE, mode="a", header=not file_exists,
                  index=False, encoding="utf-8-sig",
                  quoting=csv.QUOTE_MINIMAL, escapechar="\\", lineterminator="\n")

def load_csv_safely(path: Path) -> pd.DataFrame:
    if not path.exists(): return pd.DataFrame()
    for kw in [
        dict(encoding="utf-8-sig"),
        dict(encoding="utf-8"),
        dict(encoding="utf-8-sig", engine="python", on_bad_lines="skip"),
        dict(encoding="utf-8", engine="python", on_bad_lines="skip"),
        dict(encoding="latin-1", engine="python", on_bad_lines="skip"),
    ]:
        try:
            df = pd.read_csv(path, **kw)
            df.columns = [c.replace("\ufeff","").strip() for c in df.columns]
            return df
        except Exception:
            pass
    return pd.DataFrame()

def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
        df.to_excel(w, sheet_name=sheet_name, index=False)
        ws = w.sheets[sheet_name]
        for i, col in enumerate(df.columns):
            width = 12
            if not df.empty:
                width = min(60, max(12, int(df[col].astype(str).map(len).max()) + 4))
            ws.set_column(i, i, width)
    bio.seek(0)
    return bio.read()

def valid_email(v: str) -> bool:  return bool(re.match(r"^[^@]+@[^@]+\.[^@]+$", v.strip()))
def valid_phone(v: str) -> bool:  return bool(re.match(r"^0\d{1,2}-?\d{6,7}$", v.strip()))
def valid_id(v: str) -> bool:     return bool(re.match(r"^\d{8,9}$", v.strip()))

def show_errors(errors: list[str]):
    if not errors: return
    st.markdown("### :red[נמצאו שגיאות:]")
    for e in errors: st.markdown(f"- :red[{e}]")

# =========================
# התמדה: טעינה/שמירה
# =========================
PERSIST_KEYS = {
    "step",
    "first_name","last_name","nat_id","gender","social_affil",
    "mother_tongue","other_mt","extra_langs","extra_langs_other",
    "phone","address","email","study_year","study_year_other","track",
    "prev_training","prev_place","prev_mentor","prev_partner",
    "chosen_domains","domains_other","top_domain",
    "special_request","avg_grade","adjustments","adjustments_other","adjustments_details",
    "m1","m2","m3","arrival_confirm","confirm",
    "rank_1","rank_2","rank_3",
    "ack_0","ack_1","ack_2","ack_3","ack_4"
}

def load_persisted_state():
    if STATE_FILE.exists():
        try:
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            for k, v in data.items():
                if k in PERSIST_KEYS and k not in st.session_state:
                    st.session_state[k] = v
        except Exception:
            pass

def save_persisted_state():
    try:
        payload = {k: st.session_state.get(k) for k in PERSIST_KEYS}
        STATE_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

load_persisted_state()

# =========================
# אתחול חד-פעמי לערכי ברירת מחדל
# =========================
def ensure_state(key, default):
    if key not in st.session_state:
        st.session_state[key] = default

ensure_state("step", 0)
for k, v in dict(
    first_name="", last_name="", nat_id="", gender="זכר", social_affil="יהודי/ה",
    mother_tongue="עברית", other_mt="", extra_langs=[], extra_langs_other="",
    phone="", address="", email="",
    study_year="תואר ראשון - שנה א", study_year_other="", track="תואר ראשון – תוכנית רגילה",
    prev_training="לא", prev_place="", prev_mentor="", prev_partner="",
    chosen_domains=[], domains_other="", top_domain="— בחר/י —",
    special_request="", avg_grade=0.0,
    adjustments=[], adjustments_other="", adjustments_details="",
    m1="", m2="", m3="",
    arrival_confirm=False, confirm=False,
    rank_1="— בחר/י —", rank_2="— בחר/י —", rank_3="— בחר/י —",
    ack_0=False, ack_1=False, ack_2=False, ack_3=False, ack_4=False
).items():
    ensure_state(k, v)

# נשמור אחרי אתחול – לא ישנה קיים
save_persisted_state()

# =========================
# עטיפות Widgets שלא מאפסות ערכים
# (לא שולחים value/index/default אם יש כבר key ב-session_state)
# =========================
def _persist_on_change(key):
    def _save():
        # נורמליזציה קטנה
        if st.session_state.get(key) is None:
            st.session_state[key] = ""
        save_persisted_state()
    return _save

def ti(label, key, **kw):
    kw["key"] = key
    if key not in st.session_state and "value" not in kw:
        kw["value"] = ""
    kw.setdefault("on_change", _persist_on_change(key))
    return st.text_input(label, **kw)

def ta(label, key, **kw):
    kw["key"] = key
    if key not in st.session_state and "value" not in kw:
        kw["value"] = ""
    kw.setdefault("on_change", _persist_on_change(key))
    return st.text_area(label, **kw)

def sb(label, options, key, **kw):
    kw["key"] = key
    # index רק אם אין state קודם
    if key not in st.session_state and "index" not in kw:
        if st.session_state.get(key) in options:
            kw["index"] = options.index(st.session_state[key])
    kw.setdefault("on_change", _persist_on_change(key))
    return st.selectbox(label, options, **kw)

def ms(label, options, key, **kw):
    kw["key"] = key
    if key not in st.session_state and "default" not in kw:
        kw["default"] = []
    kw.setdefault("on_change", _persist_on_change(key))
    return st.multiselect(label, options, **kw)

def ra(label, options, key, **kw):
    kw["key"] = key
    if key not in st.session_state and "index" not in kw:
        current = st.session_state.get(key, options[0] if options else None)
        if current in options:
            kw["index"] = options.index(current)
    kw.setdefault("on_change", _persist_on_change(key))
    return st.radio(label, options, **kw)

def ni(label, key, **kw):
    kw["key"] = key
    if key not in st.session_state and "value" not in kw:
        kw["value"] = kw.get("min_value", 0.0)
    kw.setdefault("on_change", _persist_on_change(key))
    return st.number_input(label, **kw)

def cb(label, key, **kw):
    kw["key"] = key
    if key not in st.session_state and "value" not in kw:
        kw["value"] = False
    kw.setdefault("on_change", _persist_on_change(key))
    return st.checkbox(label, **kw)

# =========================
# מצב מנהל
# =========================
if is_admin_mode:
    st.title("🔑 גישת מנהל – צפייה והורדות (מאסטר + יומן)")
    ti("סיסמת מנהל", key="admin_pwd_input", type="password")
    if st.session_state.admin_pwd_input == ADMIN_PASSWORD:
        st.success("התחברת בהצלחה ✅")
        df_master = load_csv_safely(CSV_FILE)
        df_log    = load_csv_safely(CSV_LOG_FILE)

        st.subheader("📦 קובץ ראשי (מאסטר)")
        if not df_master.empty:
            st.dataframe(df_master, use_container_width=True)
            st.download_button("⬇ הורד Excel – קובץ ראשי",
                               data=df_to_excel_bytes(df_master, "Master"),
                               file_name="שאלון_שיבוץ_master.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("אין עדיין נתונים בקובץ הראשי.")

        st.subheader("🧾 קובץ יומן (Append-Only)")
        if not df_log.empty:
            st.dataframe(df_log, use_container_width=True)
            st.download_button("⬇ הורד Excel – קובץ יומן",
                               data=df_to_excel_bytes(df_log, "Log"),
                               file_name="שאלון_שיבוץ_log.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("אין עדיין נתונים ביומן.")
    else:
        if st.session_state.admin_pwd_input:
            st.error("סיסמה שגויה")
    st.stop()

# =========================
# ניווט
# =========================
def goto(i: int):
    st.session_state.step = int(i)
    save_persisted_state()

def render_nav():
    total = 6
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.session_state.step > 0:
            st.button("קודם ⬅", on_click=goto, args=(st.session_state.step - 1,),
                      use_container_width=True, key=f"nav_prev_{st.session_state.step}")
    with c2:
        st.markdown(f"<div style='text-align:center;color:#64748b'>שלב {st.session_state.step+1} מתוך {total}</div>", unsafe_allow_html=True)
    with c3:
        if st.session_state.step < total - 1:
            # הנעילה לפי ack_i
            need_ack = st.session_state.step <= 4
            disabled = (need_ack and not st.session_state.get(f"ack_{st.session_state.step}", False))
            st.button("הבא ➡", on_click=goto, args=(st.session_state.step + 1,),
                      disabled=disabled, use_container_width=True, key=f"nav_next_{st.session_state.step}")

# =========================
# טופס — שלבים
# =========================
st.title("📋 שאלון שיבוץ סטודנטים – שנת הכשרה תשפ״ו")
st.caption("מלאו/מלאי את כל הסעיפים. השדות המסומנים ב-* הינם חובה.")
STEPS = ["סעיף 1: פרטים אישיים","סעיף 2: העדפת שיבוץ","סעיף 3: נתונים אקדמיים","סעיף 4: התאמות","סעיף 5: מוטיבציה","סעיף 6: סיכום ושליחה"]
step = st.session_state.step
st.subheader(STEPS[step])

# ===== שלב 1 =====
if step == 0:
    render_nav()
    ti("שם פרטי *", key="first_name")
    ti("שם משפחה *", key="last_name")
    ti("מספר תעודת זהות *", key="nat_id")
    ra("מין *", ["זכר","נקבה"], key="gender", horizontal=True)
    sb("שיוך חברתי *", ["יהודי/ה","מוסלמי/ת","נוצרי/ה","דרוזי/ת"], key="social_affil")

    sb("שפת אם *", ["עברית","ערבית","רוסית","אחר..."], key="mother_tongue")
    if st.session_state.mother_tongue == "אחר...":
        ti("ציין/ני שפת אם אחרת *", key="other_mt")

    ms("ציין/י שפות נוספות (ברמת שיחה) *",
       ["עברית","ערבית","רוסית","אמהרית","אנגלית","ספרדית","אחר..."],
       key="extra_langs", placeholder="בחר/י שפות נוספות")
    if "אחר..." in st.session_state.extra_langs:
        ti("ציין/י שפה נוספת (אחר) *", key="extra_langs_other")

    ti("מספר טלפון נייד * (למשל 050-1234567)", key="phone")
    ti("כתובת מלאה (כולל יישוב) *", key="address")
    ti("כתובת דוא״ל *", key="email")

    sb("שנת הלימודים *",
       ["תואר ראשון - שנה א","תואר ראשון - שנה ב","תואר ראשון - שנה ג'","תואר שני - שנה א'","תואר שני - שנה ב","אחר"],
       key="study_year")
    if st.session_state.study_year == "אחר":
        ti("פרט/י שנת לימודים *", key="study_year_other")

    sb("מסלול הלימודים / תואר *", ["תואר ראשון – תוכנית רגילה","תואר ראשון – הסבה","תואר שני"], key="track")

    st.markdown("---")
    cb("אני מצהיר/ה כי מילאתי פרטים אישיים באופן מדויק. *", key="ack_0"); save_persisted_state()
    render_nav()

# ===== שלב 2 =====
if step == 1:
    render_nav()
    sb("האם עברת הכשרה מעשית בשנה קודמת? *", ["כן","לא","אחר..."], key="prev_training")
    if st.session_state.prev_training in ["כן","אחר..."]:
        ti("אם כן, נא ציין שם מקום ותחום ההתמחות *", key="prev_place")
        ti("שם המדריך והמיקום הגיאוגרפי של ההכשרה *", key="prev_mentor")
        ti("מי היה/תה בן/בת הזוג להתמחות בשנה הקודמת? *", key="prev_partner")

    all_domains = ["רווחה","מוגבלות","זקנה","ילדים ונוער","בריאות הנפש","שיקום","משפחה","נשים","בריאות","קהילה","אחר..."]
    ms("בחרו עד 3 תחומים *", all_domains, key="chosen_domains", placeholder="בחר/י עד שלושה תחומים")
    if "אחר..." in st.session_state.chosen_domains:
        ti("פרט/י תחום אחר *", key="domains_other")

    sb("מה התחום הכי מועדף עליך, מבין שלושתם? *",
       ["— בחר/י —"] + (st.session_state.chosen_domains or []),
       key="top_domain")

    st.markdown("<div style='font-weight:700; font-size:1rem; color:#0f172a;'>הדירוג אינו מחייב את מורי השיטות.</div>", unsafe_allow_html=True)
    st.markdown("**בחר/י מוסד לכל מקום הכשרה (1 = הכי רוצים, 3 = הכי פחות). הבחירה כובלת קדימה — מוסדות שנבחרו ייעלמו מהבחירות הבאות.**")

    def options_for_rank(rank_i: int) -> list:
        current = st.session_state.get(f"rank_{rank_i}", "— בחר/י —")
        chosen_before = {st.session_state.get(f"rank_{j}") for j in range(1, rank_i)}
        return ["— בחר/י —"] + [s for s in SITES if (s not in chosen_before or s == current)]

    cols = st.columns(2)
    for i in range(1, RANK_COUNT + 1):
        with cols[(i - 1) % 2]:
            sb(f"מקום הכשרה {i} (בחר/י מוסד) *", options_for_rank(i), key=f"rank_{i}")

    ta("האם קיימת בקשה מיוחדת הקשורה למיקום או תחום ההתמחות? *", key="special_request", height=100)

    st.markdown("---")
    cb("אני מצהיר/ה כי העדפתי הוזנו במלואן. *", key="ack_1"); save_persisted_state()
    render_nav()

# ===== שלב 3 =====
if step == 2:
    render_nav()
    ni("ממוצע ציונים *", key="avg_grade", min_value=0.0, max_value=100.0, step=0.1)
    st.markdown("---")
    cb("אני מצהיר/ה כי הממוצע שהזנתי נכון. *", key="ack_2"); save_persisted_state()
    render_nav()

# ===== שלב 4 =====
if step == 3:
    render_nav()
    ms("סוגי התאמות (ניתן לבחור כמה) *",
       ["אין","הריון","מגבלה רפואית (למשל: מחלה כרונית, אוטואימונית)",
        "רגישות למרחב רפואי (למשל: לא לשיבוץ בבית חולים)","אלרגיה חמורה","נכות",
        "רקע משפחתי רגיש (למשל: בן משפחה עם פגיעה נפשית)","אחר..."],
       key="adjustments", placeholder="בחר/י אפשרויות התאמה")
    if "אחר..." in st.session_state.adjustments:
        ti("פרט/י התאמה אחרת *", key="adjustments_other")
    if "אין" not in st.session_state.adjustments:
        ta("פרט: *", key="adjustments_details", height=100)
    st.markdown("---")
    cb("אני מצהיר/ה כי מסרתי מידע מדויק על התאמות. *", key="ack_3"); save_persisted_state()
    render_nav()

# ===== שלב 5 =====
if step == 4:
    render_nav()
    likert = ["בכלל לא מסכים/ה","1","2","3","4","מסכים/ה מאוד"]
    ra("1) מוכן/ה להשקיע מאמץ נוסף להגיע למקום המועדף *", likert, key="m1", horizontal=True)
    ra("2) ההכשרה המעשית חשובה לי כהזדמנות משמעותית להתפתחות *", likert, key="m2", horizontal=True)
    ra("3) אהיה מחויב/ת להגיע בזמן ולהתמיד גם בתנאים מאתגרים *", likert, key="m3", horizontal=True)
    st.markdown("---")
    cb("אני מצהיר/ה כי עניתי בכנות על שאלות המוטיבציה. *", key="ack_4"); save_persisted_state()
    render_nav()

# ===== שלב 6 =====
submitted = False
if step == 5:
    render_nav()
    rank_to_site = {i: st.session_state.get(f"rank_{i}", "— בחר/י —") for i in range(1, RANK_COUNT + 1)}
    site_to_rank = {s: None for s in SITES}
    for i, s in rank_to_site.items():
        if s and s != "— בחר/י —": site_to_rank[s] = i

    st.markdown("### 📍 העדפות שיבוץ (1=הכי רוצים)")
    st.table(pd.DataFrame({"דירוג":[f"{rank_to_site[i]} – {i}" if rank_to_site[i]!="— בחר/י —" else f"(לא נבחר) – {i}" for i in range(1,RANK_COUNT+1)]}))
    st.markdown("### 🧑‍💻 פרטים אישיים")
    st.table(pd.DataFrame([{
        "שם פרטי": st.session_state.first_name, "שם משפחה": st.session_state.last_name, "ת״ז": st.session_state.nat_id, "מין": st.session_state.gender,
        "שיוך חברתי": st.session_state.social_affil,
        "שפת אם": (st.session_state.other_mt if st.session_state.mother_tongue=="אחר..." else st.session_state.mother_tongue),
        "שפות נוספות": "; ".join([x for x in st.session_state.extra_langs if x!="אחר..."] + ([st.session_state.extra_langs_other] if "אחר..." in st.session_state.extra_langs else [])),
        "טלפון": st.session_state.phone, "כתובת": st.session_state.address, "אימייל": st.session_state.email,
        "שנת לימודים": (st.session_state.study_year_other if st.session_state.study_year=="אחר" else st.session_state.study_year),
        "מסלול לימודים": st.session_state.track,
    }]).T.rename(columns={0:"ערך"}))
    st.markdown("### 🎓 נתונים אקדמיים")
    st.table(pd.DataFrame([{"ממוצע ציונים": st.session_state.avg_grade}]).T.rename(columns={0:"ערך"}))
    st.markdown("### 🧪 התאמות")
    st.table(pd.DataFrame([{
        "התאמות": "; ".join([a for a in st.session_state.adjustments if a!="אחר..."] + ([st.session_state.adjustments_other] if "אחר..." in st.session_state.adjustments else [])),
        "פירוט התאמות": st.session_state.adjustments_details,
    }]).T.rename(columns={0:"ערך"}))
    st.markdown("### 🔥 מוטיבציה")
    st.table(pd.DataFrame([{
        "מוכנות להשקיע מאמץ": st.session_state.m1,
        "חשיבות ההכשרה": st.session_state.m2,
        "מחויבות והתמדה": st.session_state.m3
    }]).T.rename(columns={0:"ערך"}))
    st.markdown("---")
    cb("אני מצהיר/ה שאגיע בכל דרך להכשרה המעשית שתיקבע לי. *", key="arrival_confirm")
    cb("אני מאשר/ת כי המידע שמסרתי נכון ומדויק, וידוע לי שאין התחייבות להתאמה מלאה לבחירותיי. *", key="confirm")
    submitted = st.button("שליחה ✉️"); save_persisted_state()
    render_nav()

# =========================
# ולידציה ושמירה
# =========================
if submitted:
    errors = []
    if not st.session_state.first_name.strip(): errors.append("סעיף 1: יש למלא שם פרטי.")
    if not st.session_state.last_name.strip(): errors.append("סעיף 1: יש למלא שם משפחה.")
    if not valid_id(st.session_state.nat_id): errors.append("סעיף 1: ת״ז חייבת להיות 8–9 ספרות.")
    if st.session_state.mother_tongue=="אחר..." and not st.session_state.other_mt.strip(): errors.append("סעיף 1: יש לציין שפת אם (אחר).")
    if (not st.session_state.extra_langs) or ("אחר..." in st.session_state.extra_langs and not st.session_state.extra_langs_other.strip()):
        errors.append("סעיף 1: יש לבחור שפות נוספות (ואם 'אחר' – לפרט).")
    if not valid_phone(st.session_state.phone): errors.append("סעיף 1: מספר טלפון אינו תקין.")
    if not st.session_state.address.strip(): errors.append("סעיף 1: יש למלא כתובת מלאה.")
    if not valid_email(st.session_state.email): errors.append("סעיף 1: כתובת דוא״ל אינה תקינה.")
    if st.session_state.study_year=="אחר" and not st.session_state.study_year_other.strip(): errors.append("סעיף 1: יש לפרט שנת לימודים (אחר).")
    if not st.session_state.track.strip(): errors.append("סעיף 1: יש למלא מסלול לימודים/תואר.")

    rank_to_site = {i: st.session_state.get(f"rank_{i}", "— בחר/י —") for i in range(1, RANK_COUNT + 1)}
    missing = [i for i, s in rank_to_site.items() if s=="— בחר/י —"]
    if missing: errors.append(f"סעיף 2: יש לבחור מוסד לכל מקום הכשרה. חסר/ים: {', '.join(map(str, missing))}.")
    chosen_sites = [s for s in rank_to_site.values() if s!="— בחר/י —"]
    if len(set(chosen_sites)) != len(chosen_sites): errors.append("סעיף 2: קיימת כפילות בבחירת מוסדות.")
    if st.session_state.prev_training in ["כן","אחר..."] and not st.session_state.prev_place.strip(): errors.append("סעיף 2: יש למלא מקום/תחום אם הייתה הכשרה קודמת.")
    if st.session_state.prev_training in ["כן","אחר..."] and not st.session_state.prev_mentor.strip(): errors.append("סעיף 2: יש למלא שם מדריך ומיקום.")
    if st.session_state.prev_training in ["כן","אחר..."] and not st.session_state.prev_partner.strip(): errors.append("סעיף 2: יש למלא בן/בת זוג להתמחות.")
    if not st.session_state.chosen_domains: errors.append("סעיף 2: יש לבחור עד 3 תחומים (לפחות אחד).")
    if "אחר..." in st.session_state.chosen_domains and not st.session_state.domains_other.strip(): errors.append("סעיף 2: נבחר 'אחר' – יש לפרט תחום.")
    if st.session_state.chosen_domains and (st.session_state.top_domain not in st.session_state.chosen_domains): errors.append("סעיף 2: יש לבחור תחום מוביל מתוך השלושה.")
    if any("רווחה" in d for d in st.session_state.chosen_domains) and "שנה ג'" not in st.session_state.study_year:
        errors.append("סעיף 2: תחום רווחה פתוח לשיבוץ רק לסטודנטים שנה ג׳ ומעלה.")
    if not st.session_state.special_request.strip(): errors.append("סעיף 2: יש לציין בקשה מיוחדת (אפשר 'אין').")

    if st.session_state.avg_grade is None or st.session_state.avg_grade <= 0: errors.append("סעיף 3: יש להזין ממוצע ציונים גדול מ-0.")

    adj_list = [a.strip() for a in st.session_state.adjustments]
    has_none = ("אין" in adj_list) and (len([a for a in adj_list if a != "אין"]) == 0)
    if not adj_list: errors.append("סעיף 4: יש לבחור לפחות סוג התאמה אחד (או לציין 'אין').")
    if "אחר..." in adj_list and not st.session_state.adjustments_other.strip(): errors.append("סעיף 4: נבחר 'אחר' – יש לפרט התאמה.")
    if not has_none and not st.session_state.adjustments_details.strip(): errors.append("סעיף 4: יש לפרט התייחסות להתאמות.")

    if not (st.session_state.m1 and st.session_state.m2 and st.session_state.m3): errors.append("סעיף 5: יש לענות על שלוש שאלות המוטיבציה.")
    if not st.session_state.arrival_confirm: errors.append("סעיף 6: יש לסמן את ההצהרה על הגעה להכשרה.")
    if not st.session_state.confirm: errors.append("סעיף 6: יש לאשר את הצהרת הדיוק וההתאמה.")

    if errors:
        show_errors(errors); save_persisted_state()
    else:
        site_to_rank = {s: None for s in SITES}
        for i in range(1, RANK_COUNT + 1):
            site = st.session_state.get(f"rank_{i}")
            site_to_rank[site] = i
        tz = pytz.timezone("Asia/Jerusalem")
        row = {
            "תאריך שליחה": datetime.now(tz).strftime("%d/%m/%Y %H:%M:%S"),
            "שם פרטי": st.session_state.first_name.strip(),
            "שם משפחה": st.session_state.last_name.strip(),
            "תעודת זהות": st.session_state.nat_id.strip(),
            "מין": st.session_state.gender,
            "שיוך חברתי": st.session_state.social_affil,
            "שפת אם": (st.session_state.other_mt.strip() if st.session_state.mother_tongue=="אחר..." else st.session_state.mother_tongue),
            "שפות נוספות": "; ".join([x for x in st.session_state.extra_langs if x!="אחר..."] + ([st.session_state.extra_langs_other.strip()] if "אחר..." in st.session_state.extra_langs else [])),
            "טלפון": st.session_state.phone.strip(),
            "כתובת": st.session_state.address.strip(),
            "אימייל": st.session_state.email.strip(),
            "שנת לימודים": (st.session_state.study_year_other.strip() if st.session_state.study_year=="אחר" else st.session_state.study_year),
            "מסלול לימודים": st.session_state.track.strip(),
            "הכשרה קודמת": st.session_state.prev_training,
            "הכשרה קודמת מקום ותחום": st.session_state.prev_place.strip(),
            "הכשרה קודמת מדריך ומיקום": st.session_state.prev_mentor.strip(),
            "הכשרה קודמת בן זוג": st.session_state.prev_partner.strip(),
            "תחומים מועדפים": "; ".join([d for d in st.session_state.chosen_domains if d!="אחר..."] + ([st.session_state.domains_other.strip()] if "אחר..." in st.session_state.chosen_domains else [])),
            "תחום מוביל": (st.session_state.top_domain if st.session_state.top_domain and st.session_state.top_domain!="— בחר/י —" else ""),
            "בקשה מיוחדת": st.session_state.special_request.strip(),
            "ממוצע": st.session_state.avg_grade,
            "התאמות": "; ".join([a for a in st.session_state.adjustments if a!="אחר..."] + ([st.session_state.adjustments_other.strip()] if "אחר..." in st.session_state.adjustments else [])),
            "התאמות פרטים": st.session_state.adjustments_details.strip(),
            "מוטיבציה 1": st.session_state.m1,
            "מוטיבציה 2": st.session_state.m2,
            "מוטיבציה 3": st.session_state.m3,
            "אישור הגעה להכשרה": "כן" if st.session_state.arrival_confirm else "לא",
        }
        for i in range(1, RANK_COUNT + 1): row[f"מקום הכשרה {i}"] = st.session_state.get(f"rank_{i}")
        for s in SITES: row[f"דירוג_{s}"] = site_to_rank[s]

        try:
            save_master_dataframe(row)
            append_to_log(pd.DataFrame([row]))
            # איפוס קובץ מצב לאחר שליחה מוצלחת בלבד
            if STATE_FILE.exists():
                try: STATE_FILE.unlink()
                except Exception: pass
            st.success("✅ הטופס נשלח ונשמר בהצלחה! תודה רבה.")
        except Exception as e:
            st.error(f"❌ שמירה נכשלה: {e}")
            save_persisted_state()
