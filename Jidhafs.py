import re
import base64
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side, Protection


st.set_page_config(page_title="مركز التصحيح المركزي", layout="wide")


def get_logo_base64():
    logo_path = Path("logo.png")
    if logo_path.exists():
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


logo_base64 = get_logo_base64()

logo_html = ""
if logo_base64:
    logo_html = f'<img src="data:image/png;base64,{logo_base64}" class="logo-img">'
st.markdown("""
<style>

/* توسيط جميع عناوين Streamlit */
[data-testid="stMarkdownContainer"] h3 {
    text-align: center !important;
}

</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<style>
.block-container {{
    max-width: 1200px;
    margin: auto;
    padding-top: 2rem;
}}

.jidhafs-title {{
    max-width: 1100px;
    margin: 30px auto 35px auto;
    padding: 35px 25px 45px 25px;
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
    font-size:38px;
    line-height:1.5;
    font-weight:800;
}}
</style>

<div class="jidhafs-title">
    {logo_html}
    <h1>
        مركز التصحيح المركزي بمدرسة جدحفص
        الثانوية للبنات
    </h1>
</div>
""", unsafe_allow_html=True)
st.markdown("""
<style>

/* جعل التطبيق كله من اليمين */
html, body, [class*="css"]  {
    direction: rtl;
    text-align: right;
}

/* التبويبات من اليمين */
button[data-baseweb="tab"] {
    direction: rtl;
    text-align: right;
}

/* الفورم */
label, .stFileUploader, .stNumberInput {
    text-align: right;
}

/* الأزرار */
.stButton > button {
    float: right;
}

/* الجدول */
.stDataFrame {
    direction: rtl;
}

/* العنوان يبقى في النص */
.jidhafs-title {
    direction: rtl;
    text-align: center;
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>

/* تكبير النص العام */
html, body, [class*="css"] {
    font-size: 20px !important;
}

/* عناوين الأقسام */
h2, h3 {
    font-size: 26px !important;
    font-weight: bold;
}

/* التبويبات */
button[data-baseweb="tab"] {
    font-size: 20px !important;
}

/* النصوص داخل الفورم */
label, .stTextInput, .stNumberInput, .stFileUploader {
    font-size: 20px !important;
}

/* الأزرار */
.stButton button {
    font-size: 20px !important;
    padding: 10px 20px;
}

/* الجدول */
.stDataFrame {
    font-size: 20px !important;
}

</style>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["✂️ تقسيم الاستجابات", "📥 تجميع الاستجابات"])


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


def clean_header(header):
    header = str(header or "")
    header = header.replace("\xa0", " ")
    header = " ".join(header.split())
    return header


def get_part_number(filename):
    match = re.search(r"جزء[_\s-]*(\d+)", filename)
    if match:
        return int(match.group(1))
    return 999999


def format_excel_file(excel_bytes, lock_sheet=False, merge_mode=False):
    excel_bytes.seek(0)
    wb = load_workbook(excel_bytes)
    ws = wb.active

    header_fill = PatternFill("solid", fgColor="BFEFFF")
    thin_border = Border(
        left=Side(style="thin", color="808080"),
        right=Side(style="thin", color="808080"),
        top=Side(style="thin", color="808080"),
        bottom=Side(style="thin", color="808080")
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

    for col in ws.columns:
        col_letter = col[0].column_letter
        header = clean_header(col[0].value)
        header_lower = header.lower()

        ws.column_dimensions[col_letter].hidden = False

        if header_lower == "total points":
            total_col_letter = col_letter

        if (
            "points" in header_lower
            and header_lower != "total points"
            and header_lower not in [
                "points - الاسم الرباعي",
                "points - الرقم الأكاديمي",
                "points - الشعبة",
            ]
        ):
            points_cols_for_total.append(col_letter)

            if lock_sheet:
                for cell in col[1:]:
                    cell.protection = Protection(locked=False)

        if merge_mode:
            if header_lower == "رقم":
                ws.column_dimensions[col_letter].hidden = True
                continue
        else:
            if "feedback" in header_lower or header_lower in hidden_headers:
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

            if lock_sheet:
                ws[f"{total_col_letter}{row_num}"].protection = Protection(locked=True)

    for row in ws.iter_rows():
        row_num = row[0].row

        if row_num == 1:
            ws.row_dimensions[row_num].height = 55
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
        ws.protection.password = "1234"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


with tab1:
    st.subheader("✂️ تقسيم ملف الاستجابات")

    uploaded_file = st.file_uploader(
        "ارفع ملف Excel الأصلي للاستجابات",
        type=["xlsx"],
        key="split_file"
    )

    chunk_size = st.number_input(
        "عدد الاستجابات في كل ملف",
        min_value=1,
        value=25,
        step=1
    )

    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        st.success(f"تم رفع الملف بنجاح، عدد الاستجابات: {len(df)}")
        st.dataframe(df.head())

        if st.button("تقسيم الاستجابات"):
            zip_buffer = BytesIO()

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
                        merge_mode=False
                    )

                    file_number = (i // chunk_size) + 1
                    zip_file.writestr(
                        f"جزء_{file_number}.xlsx",
                        formatted_file.getvalue()
                    )

            zip_buffer.seek(0)

            st.download_button(
                label="⬇️ تنزيل الملفات المقسمة ZIP",
                data=zip_buffer,
                file_name="split_files.zip",
                mime="application/zip"
            )


with tab2:
    st.subheader("📥 تجميع ملفات الاستجابات المقسمة")

    uploaded_files = st.file_uploader(
        "ارفعي ملفات Excel المقسمة",
        type=["xlsx"],
        accept_multiple_files=True,
        key="merge_files"
    )

    if uploaded_files:
        sorted_files = sorted(uploaded_files, key=lambda f: get_part_number(f.name))

        all_data = []

        for file in sorted_files:
            df_part = pd.read_excel(file, engine="openpyxl")
            all_data.append(df_part)

        combined = pd.concat(all_data, ignore_index=True)

        st.success(f"تم رفع وتجميع {len(uploaded_files)} ملف")
        st.dataframe(combined.head())

        if st.button("تجميع وتنزيل الملف"):
            output_excel = BytesIO()

            with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                combined.to_excel(writer, index=False, sheet_name="Merged")

            formatted_merged = format_excel_file(
                output_excel,
                lock_sheet=False,
                merge_mode=True
            )

            st.download_button(
                label="⬇️ تنزيل ملف التجميع النهائي",
                data=formatted_merged,
                file_name="الاستجابات_مجمعة.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
