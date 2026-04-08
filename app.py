import io
import re
import pandas as pd
import streamlit as st
from pptx import Presentation
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PowerPoint → Excel", page_icon="📥", layout="centered")
st.title("PowerPoint → Excel")
st.caption("ارفع ملف البوربوينت المعجم وسيتم استخراج البيانات إلى إكسل.")

with st.sidebar:
    st.header("الإعدادات")
    OUTFILE_NAME = st.text_input("اسم ملف الإكسل", value="dictionary_output.xlsx")
    st.markdown("---")
    st.caption("""
**الأعمدة تُكتشف تلقائياً من عناوين الجدول:**
- Code
- Term - Definition → English Name + Definition
- Owner → Data Owner EN
- Classification → Classification EN
- Personal Data
- بيانات شخصية → Personal Data AR
- التصنيف → Classification AR
- المالك → Data Owner AR
- المصطلح وتعريفه → Arabic Name + Definition
    """)

pptx_file = st.file_uploader("📎 ارفع ملف البوربوينت (.pptx)", type=["pptx"])

# ========= Helpers =========
def get_cell_text(cell):
    try:
        return cell.text_frame.text.strip()
    except:
        return ""

def split_name_def(text):
    if not text:
        return "", ""
    lines = text.splitlines()
    if len(lines) >= 2:
        name = re.sub(r'[\s:]+$', '', lines[0]).strip()
        definition = "\n".join(lines[1:]).strip()
        return name, definition
    if ":" in lines[0]:
        parts = lines[0].split(":", 1)
        return parts[0].strip(), parts[1].strip()
    return lines[0].strip(), ""

# خريطة: عنوان العمود → اسم الحقل الداخلي
HEADER_MAP = {
    "code":               "code",
    "term - definition":  "en_name_def",
    "term-definition":    "en_name_def",
    "term":               "en_name_def",
    "owner":              "owner_en",
    "classification":     "class_en",
    "personal data":      "personal_data_en",
    "المصطلح وتعريفه":    "ar_name_def",
    "المصطلح":            "ar_name_def",
    "المالك":             "owner_ar",
    "التصنيف":            "class_ar",
    "بيانات شخصية":       "personal_data_ar",
}

def detect_col_map(table):
    col_map = {}
    log = []
    for c in range(len(table.columns)):
        raw_header = get_cell_text(table.cell(0, c))
        key = raw_header.strip().lower()
        field = HEADER_MAP.get(key) or HEADER_MAP.get(raw_header.strip())
        log.append(f"col[{c}] '{raw_header}' → {field or '(مجهول)'}")
        if field:
            col_map[c] = field
    return col_map, log

# ========= Run =========
run = st.button("استخراج البيانات ✅", type="primary", use_container_width=True)

if run:
    if pptx_file is None:
        st.error("الرجاء رفع ملف البوربوينت أولاً.")
        st.stop()

    prs = Presentation(io.BytesIO(pptx_file.read()))
    st.info(f"عدد الشرائح: {len(prs.slides)}")

    # اكتشف الخريطة من أول شريحة فيها جدول
    col_map = None
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_table:
                col_map, log = detect_col_map(sh.table)
                with st.expander("🔍 الأعمدة المكتشفة (اضغط للتفاصيل)"):
                    for line in log:
                        st.text(line)
                break
        if col_map:
            break

    if not col_map:
        st.error("لم يتم العثور على جداول في الملف.")
        st.stop()

    all_records = []

    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_table:
                continue
            table = sh.table
            for r in range(1, len(table.rows)):
                raw = {c: get_cell_text(table.cell(r, c)) for c in range(len(table.columns))}
                if all(not v.strip() for v in raw.values()):
                    continue

                record = {field: raw.get(col_idx, "") for col_idx, field in col_map.items()}

                en_name, en_def = split_name_def(record.get("en_name_def", ""))
                ar_name, ar_def = split_name_def(record.get("ar_name_def", ""))

                all_records.append({
                    "Code":               record.get("code", ""),
                    "English Name":       en_name,
                    "English Definition": en_def,
                    "Classification EN":  record.get("class_en", ""),
                    "Personal Data":      record.get("personal_data_en", ""),
                    "Arabic Name":        ar_name,
                    "Arabic Definition":  ar_def,
                    "Classification AR":  record.get("class_ar", ""),
                    "Data Owner EN":      record.get("owner_en", ""),
                    "Data Owner AR":      record.get("owner_ar", ""),
                })
            break  # جدول واحد لكل شريحة

    if not all_records:
        st.error("لم يتم العثور على بيانات.")
        st.stop()

    df_out = pd.DataFrame(all_records)
    df_out = df_out.fillna("").replace("nan", "")
    df_out = df_out[~df_out.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1)]
    df_out = df_out.reset_index(drop=True)

    st.success(f"✅ تم استخراج {len(df_out)} سجل من {len(prs.slides)} شريحة.")
    st.dataframe(df_out, use_container_width=True)

    # ========= بناء الإكسل =========
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dictionary"

    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11, name="Arial")
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right",  vertical="center", wrap_text=True)
    thin        = Side(style="thin", color="CCCCCC")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = list(df_out.columns)
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center
        cell.border    = border

    alt_fill = PatternFill("solid", start_color="EBF2FA")
    for r_idx, row in df_out.iterrows():
        fill = alt_fill if r_idx % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
        for c_idx, val in enumerate(row, 1):
            v = "" if str(val).strip() in ("", "nan") else str(val)
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=v)
            cell.font   = Font(size=11, name="Arial")
            cell.border = border
            cell.fill   = fill
            col_name = headers[c_idx - 1]
            if "Arabic" in col_name or col_name in ("Classification AR", "Data Owner AR"):
                cell.alignment = right_align
            elif col_name == "Code":
                cell.alignment = center
            else:
                cell.alignment = left_align

    col_widths = {
        "Code": 12, "English Name": 28, "English Definition": 50,
        "Classification EN": 18, "Personal Data": 16,
        "Arabic Name": 28, "Arabic Definition": 50,
        "Classification AR": 18, "Data Owner EN": 25, "Data Owner AR": 25,
    }
    for c_idx, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = col_widths.get(h, 20)

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)

    st.download_button(
        "⬇️ تحميل ملف الإكسل",
        data=out_buf,
        file_name=OUTFILE_NAME or "dictionary_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")
st.caption("يكتشف الأعمدة تلقائياً من عناوين الجدول • يفصل الاسم عن التعريف • يدعم العربي والإنجليزي")
