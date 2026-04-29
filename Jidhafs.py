import re
import json
import base64
import hashlib
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime

import streamlit as st
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation


# =========================
# إعدادات عامة
# =========================
st.set_page_config(page_title="مركز التصحيح المركزي", layout="wide")

APP_DATA_DIR = Path.home() / "Jidhafs_Control_Center_Data"
APP_DATA_DIR.mkdir(exist_ok=True)

LOCAL_REPORT_FILE = APP_DATA_DIR / "reports.xlsx"
GRADE_TEMPLATES_FILE = APP_DATA_DIR / "grade_templates.json"

# ضعي رابط Google Apps Script هنا لاحقًا
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycby1miY6jZcCNqtYtWZzaSkDVZRSOyeKzl2eN9aGrlnLkRCEC619vL8eLzYstP9pieKo0w/exec"

# كلمة مرور الأدمن
ADMIN_PASSWORD = "Jidhafs!1825"

# كلمة مرور موحدة للمستخدمين العاديين
USER_PASSWORD = "User!1234"


# =========================
# أدوات مساعدة
# =========================
def clean_header(header):
    header = str(header or "")
    header = header.replace("\xa0", " ")
    header = " ".join(header.split())
    return header


def safe_filename(name):
    name = clean_header(name)
    name = re.sub(r'[\\/:*?"<>|]', "-", name)
    return name.strip() or "ملف"


def arabic_day_name(dt):
    days = {
        "Saturday": "السبت",
        "Sunday": "الأحد",
        "Monday": "الإثنين",
        "Tuesday": "الثلاثاء",
        "Wednesday": "الأربعاء",
        "Thursday": "الخميس",
        "Friday": "الجمعة",
    }
    return days.get(dt.strftime("%A"), dt.strftime("%A"))


def get_logo_base64():
    logo_path = Path("logo.png")
    if logo_path.exists():
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


def load_grade_templates():
    if GRADE_TEMPLATES_FILE.exists():
        try:
            with open(GRADE_TEMPLATES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_grade_templates(templates):
    with open(GRADE_TEMPLATES_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)


def get_file_signature(df):
    cols = "||".join([clean_header(c) for c in df.columns])
    return hashlib.md5(cols.encode("utf-8")).hexdigest()


# =========================
# تسجيل الدخول
# =========================
def login_screen():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.role = None
        st.session_state.username = None

    if st.session_state.logged_in:
        return True

    st.markdown("<h2 style='text-align:center;'>🔐 تسجيل الدخول</h2>", unsafe_allow_html=True)

    role = st.radio("نوع الدخول", ["مستخدم", "أدمن"], horizontal=True)

    if role == "مستخدم":
        username = st.text_input("اسم المستخدم / اسم المعلمة")
        password = st.text_input("كلمة المرور", type="password")

        if st.button("دخول"):
            if username.strip() and password == USER_PASSWORD:
                st.session_state.logged_in = True
                st.session_state.role = "user"
                st.session_state.username = username.strip()
                st.rerun()
            else:
                st.error("الاسم أو كلمة المرور غير صحيحة ❌")

    else:
        username = st.text_input("اسم الأدمن", value="Admin")
        password = st.text_input("كلمة مرور الأدمن", type="password")

        if st.button("دخول الأدمن"):
            if password == ADMIN_PASSWORD:
                st.session_state.logged_in = True
                st.session_state.role = "admin"
                st.session_state.username = username.strip() or "Admin"
                st.rerun()
            else:
                st.error("كلمة مرور الأدمن غير صحيحة ❌")

    return False


# =========================
# التقرير المحلي + Google Sheet
# =========================
def append_local_report(record):
    columns = [
        "اليوم",
        "التاريخ",
        "الساعة",
        "اسم المستخدم",
        "نوع المستخدم",
        "نوع العملية",
        "اسم الملف الأصلي",
        "اسم الملف الجديد",
        "أسماء الأجزاء",
        "عدد الأجزاء",
        "عدد الاستجابات",
        "عدد الاستجابات في كل جزء",
        "هل تم التقسيم",
        "هل تم التجميع",
        "اسم ملف التجميع",
        "الحالة",
        "ملاحظات",
    ]

    df_new = pd.DataFrame([record], columns=columns)

    if LOCAL_REPORT_FILE.exists():
        try:
            df_old = pd.read_excel(LOCAL_REPORT_FILE)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            df_all = df_new
    else:
        df_all = df_new

    df_all.to_excel(LOCAL_REPORT_FILE, index=False)


def send_report_to_google(record):
    if not GOOGLE_SCRIPT_URL or GOOGLE_SCRIPT_URL == "PUT_YOUR_GOOGLE_SCRIPT_URL_HERE":
        return False

    try:
        response = requests.post(GOOGLE_SCRIPT_URL, json=record, timeout=8)
        return response.status_code in [200, 201]
    except Exception:
        return False


def log_operation(
    operation_type,
    original_file="",
    new_file_name="",
    part_names=None,
    part_count=0,
    response_count=0,
    responses_per_part="",
    split_done="لا",
    merge_done="لا",
    merged_file_name="",
    status="تم",
    notes="",
):
    now = datetime.now()
    record = {
        "اليوم": arabic_day_name(now),
        "التاريخ": now.strftime("%Y-%m-%d"),
        "الساعة": now.strftime("%H:%M:%S"),
        "اسم المستخدم": st.session_state.get("username", ""),
        "نوع المستخدم": st.session_state.get("role", ""),
        "نوع العملية": operation_type,
        "اسم الملف الأصلي": original_file,
        "اسم الملف الجديد": new_file_name,
        "أسماء الأجزاء": "، ".join(part_names or []),
        "عدد الأجزاء": part_count,
        "عدد الاستجابات": response_count,
        "عدد الاستجابات في كل جزء": responses_per_part,
        "هل تم التقسيم": split_done,
        "هل تم التجميع": merge_done,
        "اسم ملف التجميع": merged_file_name,
        "الحالة": status,
        "ملاحظات": notes,
    }

    append_local_report(record)
    send_report_to_google(record)


# =========================
# إعدادات الأعمدة والدرجات
# =========================
hidden_headers = [
    "id",
    "start time",
    "completion time",
    "email",
    "name",
    "grade posted time",
    "last modified time",
    "الاسم الرباعي",
    "الرقم الأكاديمي",
    "الشعبة",
    "points - الاسم الرباعي",
    "points - الرقم الأكاديمي",
    "points - الشعبة",
]


def should_auto_hide(header):
    header_lower = clean_header(header).lower()
    return (
        "feedback" in header_lower
        or header_lower in hidden_headers
        or "بيانات الطالبة" in header_lower
        or "اسم الطالبة" in header_lower
        or "الرقم الأكاديمي" in header_lower
        or "رقم الأكاديمي" in header_lower
    )


def is_points_column(header):
    header_lower = clean_header(header).lower()
    return (
        "points" in header_lower
        and header_lower != "total points"
        and header_lower not in [
            "points - الاسم الرباعي",
            "points - الرقم الأكاديمي",
            "points - الشعبة",
        ]
    )


def find_related_question(points_header, all_headers):
    points_header_clean = clean_header(points_header)

    possible_question = re.sub(r"^points[ ]*-[ ]*", "", points_header_clean, flags=re.IGNORECASE).strip()
    if possible_question and possible_question != points_header_clean:
        for h in all_headers:
            if clean_header(h) == possible_question:
                return clean_header(h)
        return possible_question

    return points_header_clean


def extract_score_from_text(text):
    """يستخرج مجموع الدرجات الظاهرة من نص السؤال (مثل: 5 درجات، ٥ درجات)
    ويجمعها إذا كان في أكثر من بند."""
    text = str(text or "")

    # تحويل الأرقام العربية إلى إنجليزية
    arabic_digits = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    text = text.translate(arabic_digits)

    # استخراج كل الأرقام المرتبطة بكلمة درجة/درجات
    matches = re.findall(r"([0-9]+(?:[.][0-9]+)?)[ ]*درجات?", text)

    if matches:
        total = sum(float(x) for x in matches)
        return int(total) if total.is_integer() else total

    return None


def detect_max_scores_from_data(df):
    max_scores = {}
    unknown_points = []
    all_headers = list(df.columns)

    for col in df.columns:
        header = clean_header(col)
        if not is_points_column(header):
            continue

        values = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(values) > 0:
            max_value = float(values.max())
            if max_value > 0:
                if max_value.is_integer():
                    max_value = int(max_value)
                max_scores[header] = max_value
            else:
                unknown_points.append({
                    "عمود الدرجة": header,
                    "السؤال": find_related_question(header, all_headers),
                    "الدرجة الكبرى": None,
                })
        else:
            unknown_points.append({
                "عمود الدرجة": header,
                "السؤال": find_related_question(header, all_headers),
                "الدرجة الكبرى": None,
            })

    return max_scores, unknown_points


def update_points_headers_with_max(ws, max_scores):
    for col in ws.columns:
        cell = col[0]
        header = clean_header(cell.value)
        if header in max_scores:
            if "الدرجة من" not in header:
                cell.value = f"{header} / الدرجة من {max_scores[header]}"


def add_score_validation(ws, max_scores):
    for col in ws.columns:
        col_letter = col[0].column_letter
        header = clean_header(col[0].value)
        original_header = header.split(" / الدرجة من ")[0].strip()

        if original_header in max_scores:
            max_score = max_scores[original_header]
            dv = DataValidation(
                type="decimal",
                operator="between",
                formula1="0",
                formula2=str(max_score),
                allow_blank=True,
            )
            dv.errorTitle = "خطأ في الدرجة"
            dv.error = f"الدرجة المدخلة أكبر من الدرجة المسموح بها. الحد الأعلى هو {max_score}."
            dv.promptTitle = "تنبيه"
            dv.prompt = f"اكتبي درجة من 0 إلى {max_score} فقط."
            dv.showErrorMessage = True
            dv.showInputMessage = True
            ws.add_data_validation(dv)
            dv.add(f"{col_letter}2:{col_letter}{ws.max_row}")


# =========================
# تنسيق Excel
# =========================
def format_excel_file(
    excel_bytes,
    lock_sheet=False,
    merge_mode=False,
    extra_hidden_columns=None,
    max_scores=None,
    excluded_columns=None,
):
    if extra_hidden_columns is None:
        extra_hidden_columns = []
    if max_scores is None:
        max_scores = {}
    if excluded_columns is None:
        excluded_columns = []

    extra_hidden_columns = [clean_header(c) for c in extra_hidden_columns]
    excluded_columns = [clean_header(c) for c in excluded_columns]

    excel_bytes.seek(0)
    wb = load_workbook(excel_bytes)
    ws = wb.active

    header_fill = PatternFill("solid", fgColor="BFEFFF")
    points_fill = PatternFill("solid", fgColor="CCFFCC")
    total_fill = PatternFill("solid", fgColor="FFF2CC")

    thin_border = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="thin", color="808080"),
    )

    visible_widths = {}
    points_cols_for_total = []
    total_col_letter = None

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
            cell.font = Font(size=12)
            cell.border = thin_border
            if lock_sheet:
                cell.protection = Protection(locked=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(size=13, bold=True)

    update_points_headers_with_max(ws, max_scores)

    for col in ws.columns:
        col_letter = col[0].column_letter
        header = clean_header(col[0].value)
        original_header = header.split(" / الدرجة من ")[0].strip()
        header_lower = original_header.lower()

        ws.column_dimensions[col_letter].hidden = False

        if header_lower == "total points":
            total_col_letter = col_letter
            for cell in col:
                cell.fill = total_fill

        # أعمدة الدرجات:
        # 1) إذا العمود له درجة كبرى معتمدة: يظهر، يتلوّن أخضر، ينفتح للمصححة، ويدخل في Total points
        # 2) إذا العمود مؤكد يدويًا بأنه صفر/لا يُحسب: يختفي ويبقى مقفل، ولا يدخل في Total points
        if (
            "points" in header.lower()
            and "total points" not in header.lower()
            and "points - الاسم الرباعي" not in header.lower()
            and "points - الرقم الأكاديمي" not in header.lower()
            and "points - الشعبة" not in header.lower()
        ):
            if original_header in max_scores:
                points_cols_for_total.append(col_letter)

                for cell in col:
                    cell.fill = points_fill

                if lock_sheet:
                    for cell in col[1:]:
                        cell.protection = Protection(locked=False)

            elif original_header in excluded_columns:
                # مؤكد يدويًا أنه لا يحتاج درجة: نخفيه ونقفله
                ws.column_dimensions[col_letter].hidden = True
                for cell in col:
                    cell.protection = Protection(locked=True)
                continue

            else:
                # غير معتمد وغير مؤكد: نخفيه احتياطًا ولا يدخل في المجموع
                ws.column_dimensions[col_letter].hidden = True
                for cell in col:
                    cell.protection = Protection(locked=True)
                continue

        if merge_mode:
            if header_lower == "رقم":
                ws.column_dimensions[col_letter].hidden = True
        else:
            if original_header in extra_hidden_columns or should_auto_hide(original_header):
                ws.column_dimensions[col_letter].hidden = True
                continue

        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        if col_letter == "A":
            width = 8
        elif max_len <= 20:
            width = 18
        elif max_len <= 60:
            width = 28
        elif max_len <= 140:
            width = 38
        else:
            width = 50

        ws.column_dimensions[col_letter].width = width
        visible_widths[col_letter] = width

    if total_col_letter and points_cols_for_total:
        for row_num in range(2, ws.max_row + 1):
            formula = "+".join([f"{col}{row_num}" for col in points_cols_for_total])
            ws[f"{total_col_letter}{row_num}"] = f"={formula}"
            ws[f"{total_col_letter}{row_num}"].fill = total_fill
            if lock_sheet:
                ws[f"{total_col_letter}{row_num}"].protection = Protection(locked=True)

    add_score_validation(ws, max_scores)

    for row in ws.iter_rows():
        row_num = row[0].row
        if row_num == 1:
            ws.row_dimensions[row_num].height = 60
            continue

        needed_height = 35
        for cell in row:
            col_letter = cell.column_letter
            if ws.column_dimensions[col_letter].hidden:
                continue
            if not cell.value:
                continue

            text = str(cell.value).strip()
            if len(text) < 20:
                continue

            width = visible_widths.get(col_letter, 30)
            chars_per_line = max(int(width * 1.35), 18)
            estimated_lines = text.count("\n") + (len(text) // chars_per_line) + 1
            cell_height = estimated_lines * 18
            needed_height = max(needed_height, cell_height)

        ws.row_dimensions[row_num].height = min(max(needed_height, 35), 409)

    ws.freeze_panes = "A2"

    if lock_sheet:
        ws.protection.sheet = True
        ws.protection.password = "J1825"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================
# ترتيب ملفات التجميع
# =========================
def get_part_number(filename):
    match = re.search(r"(?:مظروف|جزء|[-_\s])(\d+)(?:\.xlsx)?$", filename)
    if match:
        return int(match.group(1))

    nums = re.findall(r"\d+", filename)
    if nums:
        return int(nums[-1])

    return 999999


def detect_base_name_from_parts(files):
    if not files:
        return "الاستجابات"

    first_name = Path(files[0].name).stem
    base = re.sub(r"[-_\s]*(\d+)$", "", first_name).strip()
    base = re.sub(r"^(مظروف|جزء)[-_\s]*", "", base).strip()
    return safe_filename(base or "الاستجابات")


# =========================
# تصميم الواجهة
# =========================
def apply_ui_style():
    logo_base64 = get_logo_base64()
    logo_html = ""
    if logo_base64:
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" class="logo-img">'

    st.markdown(f"""
    <style>
    .block-container {{
        max-width: 1250px;
        margin: auto;
        padding-top: 2rem;
    }}

    .jidhafs-title {{
        max-width: 1100px;
        margin: 25px auto 30px auto;
        padding: 30px 25px 40px 25px;
        background-color:#eaf6ff;
        border-radius:12px;
        text-align:center;
    }}

    .logo-img {{
        height: 90px;
        width: auto;
        margin-bottom: 15px;
    }}

    .jidhafs-title h1 {{
        color:#15396b;
        margin:0;
        font-size:36px;
        line-height:1.5;
        font-weight:800;
    }}

    html, body, [class*="css"] {{
        direction: rtl;
        text-align: right;
        font-size: 19px !important;
    }}

    button[data-baseweb="tab"] {{
        direction: rtl;
        text-align: right;
        font-size: 19px !important;
    }}

    label, .stTextInput, .stNumberInput, .stFileUploader {{
        font-size: 19px !important;
        text-align: right;
    }}

    .stButton button {{
        font-size: 19px !important;
        padding: 10px 20px;
    }}

    [data-testid="stMarkdownContainer"] h3 {{
        text-align: center !important;
    }}

    .signature {{
        position: fixed;
        bottom: 20px;
        left: 20px;
        font-weight: bold;
        font-size: 14px;
        z-index: 9999;
        background: rgba(255,255,255,0.8);
        padding: 4px 8px;
        border-radius: 8px;
    }}
    </style>

    <div class="jidhafs-title">
        {logo_html}
        <h1>مركز التصحيح المركزي بمدرسة جدحفص الثانوية للبنات</h1>
    </div>

    """, unsafe_allow_html=True)


# =========================
# التطبيق
# =========================
if not login_screen():
    st.stop()

apply_ui_style()

col_user, col_logout = st.columns([4, 1])
with col_user:
    st.info(f"مرحبًا، {st.session_state.username} — نوع الدخول: {st.session_state.role}")
with col_logout:
    if st.button("تسجيل خروج"):
        st.session_state.logged_in = False
        st.session_state.role = None
        st.session_state.username = None
        st.rerun()

if st.session_state.role == "admin":
    tab1, tab2, tab3 = st.tabs(["✂️ تقسيم الاستجابات", "📥 تجميع الاستجابات", "📊 تقرير الأدمن"])
else:
    tab1, tab2 = st.tabs(["✂️ تقسيم الاستجابات", "📥 تجميع الاستجابات"])
    tab3 = None


# =========================
# تبويب التقسيم
# =========================
with tab1:
    st.markdown("<h3>✂️ تقسيم ملف الاستجابات</h3>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "ارفع ملف Excel الأصلي للاستجابات",
        type=["xlsx"],
        key="split_file",
    )

    if uploaded_file:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        original_file_name = uploaded_file.name

        st.success(f"تم رفع الملف بنجاح، عدد الاستجابات: {len(df)}")

        default_new_name = safe_filename(Path(original_file_name).stem)
        new_base_name = st.text_input(
            "اكتبي اسم الملف الجديد قبل التقسيم",
            value=default_new_name,
            help="مثال: تقن106 — ستكون الملفات: تقن106-1، تقن106-2 ...",
        )
        new_base_name = safe_filename(new_base_name)

        chunk_size = st.number_input(
            "عدد الاستجابات في كل ملف",
            min_value=1,
            value=10,
            step=1,
        )

        all_columns = list(df.columns)
        auto_hidden_columns = [col for col in all_columns if should_auto_hide(col)]

        extra_hidden_columns = st.multiselect(
            "اختاري أي أعمدة إضافية تريدين إخفاءها قبل تنزيل الملفات",
            options=all_columns,
            default=[],
        )

        hidden_preview_columns = set(auto_hidden_columns + extra_hidden_columns)

        preview_chunk = df.iloc[0:chunk_size].copy()
        preview_chunk.insert(0, "رقم", range(1, len(preview_chunk) + 1))

        visible_preview_columns = [
            col for col in preview_chunk.columns
            if col not in hidden_preview_columns
        ]

        st.markdown("<h4 style='text-align:center;'>👁️ معاينة أول ملف بعد الإخفاء</h4>", unsafe_allow_html=True)
        st.dataframe(preview_chunk[visible_preview_columns], use_container_width=True)

        if auto_hidden_columns:
            st.info("الأعمدة التي سيخفيها التطبيق تلقائيًا: " + "، ".join([str(c) for c in auto_hidden_columns]))

        st.markdown("---")
        st.markdown("### 🟩 تحديد درجات الأسئلة")

        # لا نعتمد على أعلى درجة في الاستجابات كدرجة كبرى؛ لأنها قد تكون درجة طالبة فقط.
        # نحاول قراءة الدرجة الظاهرة في نص السؤال فقط، والناقص يترك 0.00 ليدخله المستخدم من الإجابة النموذجية.
        detected_scores, unknown_points = detect_max_scores_from_data(df)
        templates = load_grade_templates()
        signature = get_file_signature(df)
        saved_template = templates.get(signature, {})

        all_points_items = []
        score_sources = {}

        for col in df.columns:
            header = clean_header(col)
            if is_points_column(header):
                question_text = find_related_question(header, list(df.columns))
                saved_value = saved_template.get("max_scores", {}).get(header)
                visible_score = extract_score_from_text(question_text)
                detected_value = detected_scores.get(header)

                if saved_value is not None and float(saved_value) > 0:
                    default_value = saved_value
                    source = "قالب محفوظ"
                elif visible_score is not None:
                    default_value = visible_score
                    source = "درجة ظاهرة في السؤال"
                elif detected_value is not None and float(detected_value) > 0:
                    default_value = detected_value
                    source = "درجة تلقائية من عمود Points"
                else:
                    default_value = 0
                    source = "مدخل من المستخدم"

                score_sources[header] = source

                all_points_items.append({
                    "عمود الدرجة": header,
                    "السؤال": question_text,
                    "الدرجة الكبرى": default_value,
                    "المصدر": source,
                    "أعلى درجة موجودة في الملف": detected_value if detected_value is not None else "غير محدد",
                })

        if not all_points_items:
            st.error("ما لقيت أعمدة Points في الملف. تأكدي أن ملف الاستجابات يحتوي أعمدة درجات.")
            max_scores = {}
            can_split = False
        else:
            st.info("البرنامج يضع الدرجات التلقائية الموجودة في أعمدة Points داخل خانة الدرجة الكبرى، وأي سؤال لا توجد له درجة تلقائية يبقى 0.00 لتدخلين درجته من الإجابة النموذجية.")
            max_scores = {}
            excluded_columns = []
            zero_confirmed = []

            for idx, item in enumerate(all_points_items, start=1):
                current_value = item["الدرجة الكبرى"]
                default_score = float(current_value) if current_value is not None else 0.0

                with st.container(border=True):
                    st.markdown(f"**{idx}. عمود الدرجة:** `{item['عمود الدرجة']}`")
                    st.caption(f"السؤال المرتبط: {item['السؤال']}")
                    st.caption(f"مصدر الدرجة الحالية: {item['المصدر']}")

                    if default_score == 0:
                        confirm_zero = st.checkbox(
                            "تأكيد أن هذا العمود لا تُكتب له درجة كبرى / لا يُحسب",
                            value=False,
                            key=f"confirm_zero_{signature}_{idx}_{item['عمود الدرجة']}",
                        )
                    else:
                        confirm_zero = False

                    score = st.number_input(
                        "الدرجة الكبرى",
                        min_value=0.0,
                        value=default_score,
                        step=0.5,
                        key=f"score_{signature}_{idx}_{item['عمود الدرجة']}",
                    )

                    if score > 0:
                        final_score = int(score) if float(score).is_integer() else score
                        max_scores[item["عمود الدرجة"]] = final_score

                        if item["المصدر"] == "مدخل من المستخدم":
                            score_sources[item["عمود الدرجة"]] = "مدخل من المستخدم"
                        elif final_score != item["الدرجة الكبرى"]:
                            score_sources[item["عمود الدرجة"]] = "تم تعديله من المستخدم"
                        else:
                            score_sources[item["عمود الدرجة"]] = item["المصدر"]

                    elif confirm_zero:
                        zero_confirmed.append(item["عمود الدرجة"])
                        excluded_columns.append(item["عمود الدرجة"])
                        score_sources[item["عمود الدرجة"]] = "مؤكد يدويًا لا يُحسب / مقفل"

            missing_scores = [
                item["عمود الدرجة"]
                for item in all_points_items
                if item["عمود الدرجة"] not in max_scores
                and item["عمود الدرجة"] not in zero_confirmed
            ]

            can_split = len(missing_scores) == 0

            if st.button("💾 حفظ قالب الدرجات لهذا الملف"):
                if can_split:
                    templates[signature] = {
                        "file_name": original_file_name,
                        "new_base_name": new_base_name,
                        "max_scores": max_scores,
                        "excluded_columns": excluded_columns,
                    }
                    save_grade_templates(templates)
                    st.success("تم حفظ قالب الدرجات بنجاح ✅")
                else:
                    st.error("لازم تكتبين الدرجة الكبرى لكل أعمدة Points قبل الحفظ.")

            if not can_split:
                st.warning("باقي أعمدة بدون درجة كبرى: " + "، ".join(missing_scores))

        if max_scores:
            total_exam_score = sum(float(v) for v in max_scores.values())
            if total_exam_score.is_integer():
                total_exam_score = int(total_exam_score)
            st.success(f"📌 الدرجة النهائية للاختبار: {total_exam_score}")

        with st.expander("عرض الدرجات الكبرى المعتمدة"):
            if max_scores:
                total_exam_score = sum(float(v) for v in max_scores.values())
                if total_exam_score.is_integer():
                    total_exam_score = int(total_exam_score)

                scores_df = pd.DataFrame([
                    {
                        "عمود الدرجة": k,
                        "الدرجة المعتمدة": v,
                        "المصدر": score_sources.get(k, "مدخل من المستخدم"),
                    }
                    for k, v in max_scores.items()
                ])

                if "excluded_columns" in locals() and excluded_columns:
                    excluded_df = pd.DataFrame([
                        {
                            "عمود الدرجة": col,
                            "الدرجة المعتمدة": 0,
                            "المصدر": score_sources.get(col, "مؤكد يدويًا لا يُحسب / مقفل"),
                        }
                        for col in excluded_columns
                    ])
                    scores_df = pd.concat([scores_df, excluded_df], ignore_index=True)

                total_row = pd.DataFrame([
                    {
                        "عمود الدرجة": "المجموع الكلي / الدرجة النهائية للاختبار",
                        "الدرجة المعتمدة": total_exam_score,
                        "المصدر": "مجموع الدرجات الظاهرة والمدخلة",
                    }
                ])

                scores_df = pd.concat([scores_df, total_row], ignore_index=True)

                st.dataframe(scores_df, use_container_width=True)
                st.success(f"📌 الدرجة النهائية للاختبار: {total_exam_score}")
            else:
                st.info("لا توجد درجات محددة حتى الآن.")

        if st.button("✂️ تقسيم الاستجابات وتنزيل الملفات", disabled=not can_split):
            zip_buffer = BytesIO()
            part_names = []
            responses_per_part_list = []

            with ZipFile(zip_buffer, "w") as zip_file:
                for i in range(0, len(df), chunk_size):
                    chunk = df.iloc[i:i + chunk_size].copy()
                    chunk.insert(0, "رقم", range(1, len(chunk) + 1))

                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                        chunk.to_excel(writer, index=False, sheet_name="Responses")

                    formatted_file = format_excel_file(
                        excel_buffer,
                        lock_sheet=True,
                        merge_mode=False,
                        extra_hidden_columns=extra_hidden_columns,
                        max_scores=max_scores,
                        excluded_columns=excluded_columns,
                    )

                    file_number = (i // chunk_size) + 1
                    part_file_name = f"{new_base_name}-{file_number}.xlsx"
                    part_names.append(part_file_name)
                    responses_per_part_list.append(str(len(chunk)))

                    zip_file.writestr(part_file_name, formatted_file.getvalue())

            zip_buffer.seek(0)
            part_count = len(part_names)

            log_operation(
                operation_type="تقسيم",
                original_file=original_file_name,
                new_file_name=new_base_name,
                part_names=part_names,
                part_count=part_count,
                response_count=len(df),
                responses_per_part="، ".join(responses_per_part_list),
                split_done="نعم",
                merge_done="لا",
                status="تم",
            )

            st.success(f"تم تقسيم الملف إلى {part_count} ملف ✅")
            st.download_button(
                label="⬇️ تنزيل الملفات المقسمة ZIP",
                data=zip_buffer,
                file_name=f"{new_base_name}-split_files.zip",
                mime="application/zip",
            )


# =========================
# تبويب التجميع
# =========================
with tab2:
    st.markdown("<h3>📥 تجميع ملفات الاستجابات المقسمة</h3>", unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "ارفعي ملفات Excel المقسمة",
        type=["xlsx"],
        accept_multiple_files=True,
        key="merge_files",
    )

    if uploaded_files:
        sorted_files = sorted(uploaded_files, key=lambda f: get_part_number(f.name))
        detected_base_name = detect_base_name_from_parts(sorted_files)

        merge_base_name = st.text_input(
            "اسم ملف التجميع النهائي",
            value=detected_base_name,
            help="مثال: تقن106 — الناتج سيكون تقن106-مجمعة.xlsx",
        )
        merge_base_name = safe_filename(merge_base_name)

        all_data = []
        for file in sorted_files:
            df_part = pd.read_excel(file, engine="openpyxl")
            all_data.append(df_part)

        combined = pd.concat(all_data, ignore_index=True)

        st.success(f"تم رفع {len(uploaded_files)} ملف، وعدد الاستجابات بعد التجميع: {len(combined)}")
        st.dataframe(combined.head(), use_container_width=True)

        if st.button("📥 تجميع وتنزيل الملف"):
            output_excel = BytesIO()
            with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                combined.to_excel(writer, index=False, sheet_name="Merged")

            formatted_merged = format_excel_file(
                output_excel,
                lock_sheet=False,
                merge_mode=True,
                extra_hidden_columns=[],
                max_scores={},
            )

            merged_file_name = f"{merge_base_name}-مجمعة.xlsx"

            log_operation(
                operation_type="تجميع",
                original_file="، ".join([f.name for f in sorted_files]),
                new_file_name=merge_base_name,
                part_names=[f.name for f in sorted_files],
                part_count=len(sorted_files),
                response_count=len(combined),
                responses_per_part="",
                split_done="لا",
                merge_done="نعم",
                merged_file_name=merged_file_name,
                status="تم",
            )

            st.success("تم تجميع الملف بنجاح ✅")
            st.download_button(
                label="⬇️ تنزيل ملف التجميع النهائي",
                data=formatted_merged,
                file_name=merged_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# =========================
# تبويب تقرير الأدمن
# =========================
if tab3 is not None:
    with tab3:
        st.markdown("<h3>📊 تقرير الأدمن</h3>", unsafe_allow_html=True)

        if LOCAL_REPORT_FILE.exists():
            report_df = pd.read_excel(LOCAL_REPORT_FILE)
            st.dataframe(report_df, use_container_width=True)

            report_buffer = BytesIO()
            report_df.to_excel(report_buffer, index=False)
            report_buffer.seek(0)

            st.download_button(
                label="⬇️ تنزيل التقرير المحلي Excel",
                data=report_buffer,
                file_name="reports.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("لا يوجد تقرير محلي حتى الآن.")

        st.warning("ملاحظة: التقرير المحلي يعرض عمليات هذا الجهاز فقط. التقرير المركزي لكل الأجهزة يكون في Google Sheet بعد إضافة رابط Apps Script.")


# =========================
# فوتر أسفل الصفحة
# =========================
st.markdown("""
<style>
.footer-container {
    margin-top: 80px;
    padding: 18px 10px;
    border-top: 1px solid #d9d9d9;
    display: flex;
    justify-content: space-between;
    gap: 20px;
    font-weight: bold;
    font-size: 14px;
    color: #333;
    direction: rtl;
}
.footer-right {
    text-align: right;
    flex: 1;
}
.footer-center {
    text-align: center;
    flex: 1;
}
.footer-left {
    text-align: left;
    flex: 1;
}
@media (max-width: 700px) {
    .footer-container {
        display: block;
        text-align: center;
    }
    .footer-right, .footer-center, .footer-left {
        text-align: center;
        margin: 6px 0;
    }
}
</style>

<div class="footer-container">
    <div class="footer-right">تصميم وبرمجة: أ. عفاف حسين</div>
    <div class="footer-center">إشراف: أ. أمينة الصائغ</div>
    <div class="footer-left">رئيسة المركز: أ. خلود يعقوب بدو</div>
</div>
""", unsafe_allow_html=True)

