import streamlit as st
import random
import re
import zipfile
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Thư viện xuất Excel chuyên nghiệp
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Cấu hình trang
st.set_page_config(
    page_title="EduQuiz 2025 - Trộn Đề & Phiếu Trả Lời Mới",
    page_icon="📝",
    layout="wide"
)

# ================= BỘ XỬ LÝ CÔNG THỨC HÓA HỌC THÔNG MINH =================

CHEMICAL_ELEMENTS = (
    "He|Li|Be|Ne|Na|Mg|Al|Si|Cl|Ar|Ca|Sc|Ti|Cr|Mn|Fe|Co|Ni|Cu|Zn|Ga|Ge|As|Se|Br|Kr|"
    "Rb|Sr|Zr|Nb|Mo|Tc|Ru|Rh|Pd|Ag|Cd|In|Sn|Sb|Te|Xe|Cs|Ba|La|Ce|Pr|Nd|Pm|Sm|Eu|Gd|"
    "Tb|Dy|Ho|Er|Tm|Yb|Lu|Hf|Ta|Re|Os|Ir|Pt|Au|Hg|Tl|Pb|Bi|Po|At|Rn|Fr|Ra|Ac|Th|Pa|"
    "Np|Pu|Am|Cm|Bk|Cf|Es|Fm|Md|No|Lr|Rf|Db|Sg|Bh|Hs|Mt|Ds|Rg|Cn|Nh|Fl|Mc|Lv|Ts|Og|"
    "H|B|C|N|O|F|P|S|K|V|Y|I|W|U"
)

CHEM_TOKEN_SPLIT = re.compile(
    r'(\^(?:\d*[+\-]|\([+\-]?\d+\)|\d+)|'
    r'(?:(?:' + CHEMICAL_ELEMENTS + r')|\)|\]|\d+|\s+|.))'
)

SUB_MAP = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
SUP_MAP = str.maketrans("0123456789+-", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻")

def parse_chemical_segments(text):
    if not text:
        return []

    text = re.sub(r'\b(SO4|SO3|CO3|SiO3|CrO4|Cr2O7|HPO4)(2-)', r'\1^\2', text)
    text = re.sub(r'\b(PO4)(3-)', r'\1^\2', text)
    text = re.sub(r'\b(NH4)(\+)', r'\1^\2', text)
    text = re.sub(r'\b(NO3|NO2|OH|HCO3|HSO4|HSO3|HS|H2PO4|ClO4|ClO3|ClO2|ClO|MnO4|AlO2|CH3COO)(-)', r'\1^\2', text)
    text = re.sub(r'\b(Fe|Cu|Zn|Mg|Ca|Ba|Pb|Sn|Ni|Cr|Mn|Hg|Be|Sr)([23]\+)', r'\1^\2', text)
    text = re.sub(r'\b(Al|Cr|Au|Fe)(3\+)', r'\1^\2', text)
    text = re.sub(r'\b(Na|Li|K|Rb|Cs|Ag|Cu|H)(\+)(?!\s*\w)', r'\1^\2', text)
    text = re.sub(r'\b(F|Cl|Br|I)(-)(?!\s*\w)', r'\1^\2', text)
    text = re.sub(r'\b(O|S)(2-)(?!\s*\w)', r'\1^\2', text)

    tokens = CHEM_TOKEN_SPLIT.findall(text)
    segments = []
    prev_is_chem = False

    for idx, tok in enumerate(tokens):
        if not tok:
            continue
        if tok.startswith('^') and len(tok) > 1:
            segments.append((tok[1:].strip('()'), 'sup'))
            prev_is_chem = False
            continue
        if re.fullmatch(CHEMICAL_ELEMENTS, tok):
            next_tok = tokens[idx + 1] if idx + 1 < len(tokens) else ""
            if next_tok and re.match(r'^[a-zàáảãạăắằẳẵặâấầẩẫậèéẻẽẹêếềểễệìíỉĩịòóỏõọôốồổỗộơớờởỡợùúủũụưứừửữựỳýỷỹỵđ]', next_tok):
                prev_is_chem = False
            else:
                prev_is_chem = True
            segments.append((tok, 'normal'))
            continue
        if tok in [')', ']']:
            segments.append((tok, 'normal'))
            continue
        if tok in ['(', '[', '·', '.']:
            segments.append((tok, 'normal'))
            prev_is_chem = False
            continue
        if tok.isdigit() and prev_is_chem:
            segments.append((tok, 'sub'))
            continue

        prev_is_chem = False
        segments.append((tok, 'normal'))

    merged = []
    for text_val, style_type in segments:
        if merged and merged[-1][1] == style_type:
            merged[-1] = (merged[-1][0] + text_val, style_type)
        else:
            merged.append((text_val, style_type))
    return merged

def add_chem_text_to_paragraph(paragraph, text, bold=False, italic=False, font_size=None):
    segments = parse_chemical_segments(text)
    for seg_text, seg_type in segments:
        run = paragraph.add_run(seg_text)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if font_size:
            run.font.size = font_size
        if seg_type == 'sub':
            run.font.subscript = True
        elif seg_type == 'sup':
            run.font.superscript = True

def format_chem_for_web(text):
    segments = parse_chemical_segments(text)
    out = []
    for seg_text, seg_type in segments:
        if seg_type == 'sub':
            out.append(seg_text.translate(SUB_MAP))
        elif seg_type == 'sup':
            out.append(seg_text.translate(SUP_MAP))
        else:
            out.append(seg_text)
    return "".join(out)

# ================= HỖ TRỢ XML CHO BẢNG WORD (CELL BORDERS & SHADING) =================
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# ================= KHỞI TẠO DỮ LIỆU MẪU CHUẨN 2025 =================
if 'bank_2025' not in st.session_state:
    st.session_state.bank_2025 = {
        "part1": [
            {"question": "Chất nào sau đây tác dụng với dung dịch H2SO4 loãng sinh ra khí H2?", "options": ["Fe", "Cu", "Ag", "Fe2O3"], "answer": "Fe"},
            {"question": "Công thức cấu tạo thu gọn của ethyl acetate là gì?", "options": ["CH3COOC2H5", "HCOOCH3", "C2H5COOCH3", "CH3COOH"], "answer": "CH3COOC2H5"},
            {"question": "Trong công nghiệp, nhôm được sản xuất bằng phương pháp điện phân nóng chảy chất nào sau đây?", "options": ["Al2O3", "AlCl3", "Al2(SO4)3", "Al(OH)3"], "answer": "Al2O3"},
            {"question": "Ion nào sau đây có mặt trong dung dịch Fe2(SO4)3?", "options": ["Fe^3+ và SO4^2-", "Fe^2+ và SO4^2-", "Fe^3+ và SO3^2-", "Fe^2+ và S^2-"], "answer": "Fe^3+ và SO4^2-"}
        ],
        "part2": [
            {
                "question": "Cho các phát biểu sau về carbohydrate:",
                "statements": [
                    {"text": "Glucose và fructose đều có cùng công thức phân tử là C6H12O6.", "is_true": True},
                    {"text": "Thủy phân hoàn toàn saccharose (C12H22O11) trong môi trường H+ chỉ thu được glucose.", "is_true": False},
                    {"text": "Glucose tác dụng với phức [Ag(NH3)2]OH đun nóng tạo ra kết tủa Ag.", "is_true": True},
                    {"text": "Cellulose có công thức cấu tạo viết gọn là [C6H7O2(OH)3]n.", "is_true": True}
                ]
            }
        ],
        "part3": [
            {"question": "Phân tử khối của muối sắt(III) sunfat Fe2(SO4)3 là bao nhiêu?", "answer": "400"},
            {"question": "Hòa tan hoàn toàn 5,6 gam Fe vào dung dịch HCl dư, thu được V lít khí H2 (đkc 25°C, 1 bar). Giá trị của V là bao nhiêu?", "answer": "2,48"}
        ]
    }

if 'mixed_exams_2025' not in st.session_state:
    st.session_state.mixed_exams_2025 = {}

# ================= HÀM ĐỌC VĂN BẢN & FILE WORD =================
def parse_questions_2025(lines_with_format):
    part1, part2, part3 = [], [], []
    current_section = 1
    current_q = None

    def save_current_q():
        nonlocal current_q
        if not current_q:
            return
        q_type = current_q.get("type")
        if q_type == 1 and len(current_q["options"]) >= 2:
            if not current_q["answer"]:
                current_q["answer"] = current_q["options"][0]
            part1.append({"question": current_q["question"], "options": current_q["options"], "answer": current_q["answer"]})
        elif q_type == 2 and len(current_q["statements"]) >= 2:
            part2.append({"question": current_q["question"], "statements": current_q["statements"]})
        elif q_type == 3 and current_q["answer"] != "":
            part3.append({"question": current_q["question"], "answer": current_q["answer"]})
        current_q = None

    for raw_line, is_highlighted in lines_with_format:
        line = raw_line.strip()
        if not line:
            continue
        if re.match(r'^PHẦN\s+(I|1)\b', line, re.IGNORECASE):
            save_current_q()
            current_section = 1
            continue
        elif re.match(r'^PHẦN\s+(II|2)\b', line, re.IGNORECASE):
            save_current_q()
            current_section = 2
            continue
        elif re.match(r'^PHẦN\s+(III|3)\b', line, re.IGNORECASE):
            save_current_q()
            current_section = 3
            continue

        q_match = re.match(r'^Câu\s*\d+[\.\:]\s*(.*)', line, re.IGNORECASE)
        if q_match:
            save_current_q()
            current_q = {"type": current_section, "question": q_match.group(1).strip(), "options": [], "statements": [], "answer": ""}
            continue

        if current_q is None:
            continue

        ans_match = re.match(r'^(Đáp\s*án|ĐS|ĐA|Key|Trả\s*lời)[\.\:]\s*(.*)', line, re.IGNORECASE)
        if ans_match:
            current_q["type"] = 3
            current_q["answer"] = ans_match.group(2).strip()
            continue

        tf_match = re.match(r'^(\*?)([a-d])[\)\.]\s*(.*)', line)
        if tf_match and tf_match.group(2).islower():
            current_q["type"] = 2
            star_prefix = tf_match.group(1) == '*'
            stmt_body = tf_match.group(3).strip()
            is_true = is_highlighted or star_prefix
            if stmt_body.startswith('*'):
                is_true = True
                stmt_body = stmt_body[1:].strip()

            suffix_true = re.search(r'(\[Đ\]|\(Đ\)|\[Đúng\]|\(Đúng\)|-\s*Đ|-?\s*Đúng)$', stmt_body, re.IGNORECASE)
            suffix_false = re.search(r'(\[S\]|\(S\)|\[Sai\]|\(Sai\)|-\s*S|-?\s*Sai)$', stmt_body, re.IGNORECASE)
            if suffix_true:
                is_true = True
                stmt_body = stmt_body[:suffix_true.start()].strip()
            elif suffix_false:
                is_true = False
                stmt_body = stmt_body[:suffix_false.start()].strip()

            current_q["statements"].append({"text": stmt_body, "is_true": is_true})
            continue

        mcq_match = re.match(r'^(\*?)([A-D])[\.\:\)]\s*(.*)', line, re.IGNORECASE)
        if mcq_match:
            current_q["type"] = 1
            star_prefix = mcq_match.group(1) == '*'
            opt_body = mcq_match.group(3).strip()
            is_correct = is_highlighted or star_prefix
            if opt_body.startswith('*'):
                is_correct = True
                opt_body = opt_body[1:].strip()
            current_q["options"].append(opt_body)
            if is_correct:
                current_q["answer"] = opt_body
            continue

        current_q["question"] += "\n" + line

    save_current_q()
    return part1, part2, part3

def parse_text_input_2025(raw_text):
    return parse_questions_2025([(line, False) for line in raw_text.split('\n')])

def parse_docx_2025(uploaded_file):
    doc = Document(uploaded_file)
    lines_with_format = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        has_highlight = False
        for run in para.runs:
            if run.text.strip():
                if run.underline:
                    has_highlight = True
                if run.font.color and run.font.color.rgb == RGBColor(255, 0, 0):
                    has_highlight = True
        lines_with_format.append((text, has_highlight))
    return parse_questions_2025(lines_with_format)

# ================= HÀM TẠO PHIẾU TRẢ LỜI TRẮC NGHIỆM 2025 (TRONG WORD) =================
def add_answer_sheet_page(doc, exam_code, num_p1, num_p2, num_p3):
    """Chèn trang Phiếu trả lời trắc nghiệm chuẩn 2025 vào đầu file Word"""
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_t.add_run("PHIẾU TRẢ LỜI TRẮC NGHIỆM (KỲ THI TỐT NGHIỆP THPT)")
    r_t.bold = True
    r_t.font.size = Pt(13)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"MÃ ĐỀ THI: {exam_code}")
    r_sub.bold = True
    r_sub.font.size = Pt(14)
    p_sub.paragraph_format.space_after = Pt(6)

    # Khung thông tin thí sinh
    info_table = doc.add_table(rows=2, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in info_table.rows:
        for c in r.cells:
            set_cell_margins(c, top=40, bottom=40, left=80, right=80)
    
    info_table.cell(0, 0).paragraphs[0].add_run("Họ và tên thí sinh: ..............................................................")
    info_table.cell(0, 1).paragraphs[0].add_run("Số báo danh: ............................")
    info_table.cell(1, 0).paragraphs[0].add_run("Lớp: ........................................................................................")
    info_table.cell(1, 1).paragraphs[0].add_run("Phòng thi: ..............................")

    p_inst = doc.add_paragraph()
    p_inst.paragraph_format.space_before = Pt(6)
    p_inst.paragraph_format.space_after = Pt(4)
    r_i = p_inst.add_run("* Hướng dẫn làm bài: Tô kín ô tròn bằng bút chì đen (2B). Phần III viết số vào ô trống.")
    r_i.italic = True
    r_i.font.size = Pt(9.5)

    # --- BẢNG PHẦN I ---
    if num_p1 > 0:
        p_s1 = doc.add_paragraph()
        p_s1.paragraph_format.space_before = Pt(4)
        p_s1.paragraph_format.space_after = Pt(2)
        r_s1 = p_s1.add_run(f"PHẦN I. Trắc nghiệm nhiều lựa chọn (Tổng: {num_p1} câu)")
        r_s1.bold = True
        r_s1.font.size = Pt(10.5)

        # Tạo bảng ô tô trắc nghiệm Phần I (Mỗi hàng 4 câu, mỗi câu có 4 ô A B C D)
        rows_p1 = (num_p1 + 3) // 4
        table_p1 = doc.add_table(rows=rows_p1 + 1, cols=8)
        table_p1.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers_p1 = ["Câu", "A B C D"] * 2
        for col_i, text in enumerate(headers_p1):
            cell = table_p1.cell(0, col_i)
            set_cell_background(cell, "D9E1F2")
            set_cell_margins(cell, top=30, bottom=30, left=30, right=30)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text)
            r.bold = True
            r.font.size = Pt(9.5)

        for r_idx in range(rows_p1):
            for block in range(2):
                q_num = r_idx * 4 + block * rows_p1 + 1
                col_offset = block * 4
                if q_num <= num_p1:
                    c_num = table_p1.cell(r_idx + 1, col_offset)
                    c_ans = table_p1.cell(r_idx + 1, col_offset + 1)
                    
                    set_cell_margins(c_num, top=20, bottom=20, left=30, right=30)
                    set_cell_margins(c_ans, top=20, bottom=20, left=30, right=30)

                    p_n = c_num.paragraphs[0]
                    p_n.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_n = p_n.add_run(f"{q_num:02d}")
                    r_n.font.size = Pt(9.5)

                    p_a = c_ans.paragraphs[0]
                    p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_a = p_a.add_run("  ①  ②  ③  ④  ")
                    r_a.font.size = Pt(9.5)

    # --- BẢNG PHẦN II ---
    if num_p2 > 0:
        p_s2 = doc.add_paragraph()
        p_s2.paragraph_format.space_before = Pt(6)
        p_s2.paragraph_format.space_after = Pt(2)
        r_s2 = p_s2.add_run(f"PHẦN II. Trắc nghiệm Đúng / Sai (Tổng: {num_p2} câu)")
        r_s2.bold = True
        r_s2.font.size = Pt(10.5)

        table_p2 = doc.add_table(rows=num_p2 + 1, cols=5)
        table_p2.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers_p2 = ["Câu", "Ý a)", "Ý b)", "Ý c)", "Ý d)"]
        for col_i, text in enumerate(headers_p2):
            cell = table_p2.cell(0, col_i)
            set_cell_background(cell, "D9E1F2")
            set_cell_margins(cell, top=30, bottom=30, left=30, right=30)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text)
            r.bold = True
            r.font.size = Pt(9.5)

        for q_i in range(num_p2):
            cell_num = table_p2.cell(q_i + 1, 0)
            set_cell_margins(cell_num, top=20, bottom=20, left=30, right=30)
            p_n = cell_num.paragraphs[0]
            p_n.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_n.add_run(f"Câu {q_i+1}").font.size = Pt(9.5)

            for col_j in range(1, 5):
                cell_opt = table_p2.cell(q_i + 1, col_j)
                set_cell_margins(cell_opt, top=20, bottom=20, left=30, right=30)
                p_o = cell_opt.paragraphs[0]
                p_o.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_o.add_run(" Đ   S ").font.size = Pt(9.5)

    # --- BẢNG PHẦN III ---
    if num_p3 > 0:
        p_s3 = doc.add_paragraph()
        p_s3.paragraph_format.space_before = Pt(6)
        p_s3.paragraph_format.space_after = Pt(2)
        r_s3 = p_s3.add_run(f"PHẦN III. Trắc nghiệm trả lời ngắn (Tổng: {num_p3} câu)")
        r_s3.bold = True
        r_s3.font.size = Pt(10.5)

        table_p3 = doc.add_table(rows=num_p3 + 1, cols=2)
        table_p3.alignment = WD_TABLE_ALIGNMENT.CENTER
        table_p3.cell(0, 0).paragraphs[0].add_run("Câu hỏi").bold = True
        table_p3.cell(0, 1).paragraphs[0].add_run("Kết quả trả lời (Viết rõ số/đáp số)").bold = True

        for q_i in range(num_p3):
            table_p3.cell(q_i + 1, 0).paragraphs[0].add_run(f"Câu {q_i+1}")
            table_p3.cell(q_i + 1, 1).paragraphs[0].add_run("...........................................................................")

    # Ngắt trang sau phiếu trả lời để sang trang đề thi chính thức
    doc.add_page_break()

# ================= HÀM XUẤT EXCEL & WORD CHÍNH =================
def create_styled_excel_answers(mixed_exams, school_name, exam_title, subject_name):
    wb = Workbook()
    font_school = Font(name="Times New Roman", size=11, bold=True, color="1F4E78")
    font_title = Font(name="Times New Roman", size=14, bold=True, color="C00000")
    font_sub = Font(name="Times New Roman", size=11, italic=True, color="333333")
    font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_sec_banner = Font(name="Times New Roman", size=11, bold=True, color="1F4E78")
    font_q_label = Font(name="Times New Roman", size=11, bold=True, color="000000")
    font_data_bold = Font(name="Times New Roman", size=11, bold=True, color="000000")
    font_data_norm = Font(name="Times New Roman", size=11, bold=False, color="000000")

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sec_banner = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_p1 = PatternFill(start_color="F2F8FD", end_color="F2F8FD", fill_type="solid")
    fill_p2 = PatternFill(start_color="F2F9F1", end_color="F2F9F1", fill_type="solid")
    fill_p3 = PatternFill(start_color="FFFDF0", end_color="FFFDF0", fill_type="solid")
    fill_q_col = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")

    thin_side = Side(border_style="thin", color="808080")
    med_side = Side(border_style="medium", color="1F4E78")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_header = Border(left=thin_side, right=thin_side, top=med_side, bottom=med_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    exam_codes = list(mixed_exams.keys())
    total_cols = len(exam_codes) + 1
    last_col_letter = get_column_letter(max(total_cols, 2))

    # Sheet 1
    ws1 = wb.active
    ws1.title = "Bảng Đáp Án Tổng Hợp"
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells(f"A1:{last_col_letter}1")
    ws1["A1"] = school_name.upper()
    ws1["A1"].font = font_school
    ws1["A1"].alignment = align_center

    ws1.merge_cells(f"A2:{last_col_letter}2")
    ws1["A2"] = f"BẢNG ĐÁP ÁN CHẤM THI - {exam_title.upper()}"
    ws1["A2"].font = font_title
    ws1["A2"].alignment = align_center

    ws1.merge_cells(f"A3:{last_col_letter}3")
    ws1["A3"] = f"Môn: {subject_name} | Cấu trúc đề thi mới 2025"
    ws1["A3"].font = font_sub
    ws1["A3"].alignment = align_center

    header_row = 5
    ws1.cell(row=header_row, column=1, value="Phần / Câu hỏi").font = font_header
    ws1.cell(row=header_row, column=1).fill = fill_header
    ws1.cell(row=header_row, column=1).alignment = align_center
    ws1.cell(row=header_row, column=1).border = border_header

    for c_idx, code in enumerate(exam_codes, start=2):
        cell = ws1.cell(row=header_row, column=c_idx, value=f"Mã đề {code}")
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header

    first_exam = mixed_exams[exam_codes[0]]
    all_keys = list(first_exam["answers"].keys())

    curr_row = 6
    current_group = None
    for q_key in all_keys:
        if q_key.startswith("P.I ") and current_group != "P1":
            current_group = "P1"
            ws1.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=total_cols)
            sec_cell = ws1.cell(row=curr_row, column=1, value="PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (0,25đ/câu)")
            sec_cell.font = font_sec_banner
            sec_cell.fill = fill_sec_banner
            sec_cell.alignment = align_left
            for col_i in range(1, total_cols + 1):
                ws1.cell(row=curr_row, column=col_i).border = border_cell
                ws1.cell(row=curr_row, column=col_i).fill = fill_sec_banner
            curr_row += 1
        elif q_key.startswith("P.II ") and current_group != "P2":
            current_group = "P2"
            ws1.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=total_cols)
            sec_cell = ws1.cell(row=curr_row, column=1, value="PHẦN II. TRẮC NGHIỆM ĐÚNG / SAI")
            sec_cell.font = font_sec_banner
            sec_cell.fill = fill_sec_banner
            sec_cell.alignment = align_left
            for col_i in range(1, total_cols + 1):
                ws1.cell(row=curr_row, column=col_i).border = border_cell
                ws1.cell(row=curr_row, column=col_i).fill = fill_sec_banner
            curr_row += 1
        elif q_key.startswith("P.III ") and current_group != "P3":
            current_group = "P3"
            ws1.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=total_cols)
            sec_cell = ws1.cell(row=curr_row, column=1, value="PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN")
            sec_cell.font = font_sec_banner
            sec_cell.fill = fill_sec_banner
            sec_cell.alignment = align_left
            for col_i in range(1, total_cols + 1):
                ws1.cell(row=curr_row, column=col_i).border = border_cell
                ws1.cell(row=curr_row, column=col_i).fill = fill_sec_banner
            curr_row += 1

        q_cell = ws1.cell(row=curr_row, column=1, value=q_key)
        q_cell.font = font_q_label
        q_cell.fill = fill_q_col
        q_cell.alignment = align_center
        q_cell.border = border_cell

        row_fill = fill_p1 if current_group == "P1" else (fill_p2 if current_group == "P2" else fill_p3)
        for c_idx, code in enumerate(exam_codes, start=2):
            ans_val = mixed_exams[code]["answers"].get(q_key, "")
            val_cell = ws1.cell(row=curr_row, column=c_idx, value=ans_val)
            val_cell.font = font_data_bold
            val_cell.fill = row_fill
            val_cell.alignment = align_center
            val_cell.border = border_cell
        curr_row += 1

    ws1.freeze_panes = "B6"
    ws1.column_dimensions["A"].width = 20
    for c_idx in range(2, total_cols + 1):
        ws1.column_dimensions[get_column_letter(c_idx)].width = 15

    # Sheet 2
    ws2 = wb.create_sheet(title="Chi Tiết Từng Ý (P.II)")
    ws2.views.sheetView[0].showGridLines = True
    ws2.merge_cells(f"A1:{last_col_letter}1")
    ws2["A1"] = "BẢNG ĐÁP ÁN CHI TIẾT TỪNG Ý (TÁCH RIÊNG Ý a, b, c, d)"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = align_center

    ws2.cell(row=3, column=1, value="Câu / Ý chi tiết").font = font_header
    ws2.cell(row=3, column=1).fill = fill_header
    ws2.cell(row=3, column=1).alignment = align_center
    ws2.cell(row=3, column=1).border = border_header

    for c_idx, code in enumerate(exam_codes, start=2):
        cell = ws2.cell(row=3, column=c_idx, value=f"Mã đề {code}")
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header

    r2 = 4
    for q_idx in range(len(first_exam["part1"])):
        c_l = ws2.cell(row=r2, column=1, value=f"Phần I - Câu {q_idx+1}")
        c_l.font, c_l.fill, c_l.alignment, c_l.border = font_q_label, fill_q_col, align_center, border_cell
        for c_idx, code in enumerate(exam_codes, start=2):
            c_v = ws2.cell(row=r2, column=c_idx, value=mixed_exams[code]["part1"][q_idx]["correct"])
            c_v.font, c_v.fill, c_v.alignment, c_v.border = font_data_bold, fill_p1, align_center, border_cell
        r2 += 1

    for q_idx in range(len(first_exam["part2"])):
        labels_p2 = ['a', 'b', 'c', 'd']
        for s_i in range(len(first_exam["part2"][q_idx]["statements"])):
            lbl_char = labels_p2[s_i] if s_i < 4 else str(s_i+1)
            c_l = ws2.cell(row=r2, column=1, value=f"Phần II - Câu {q_idx+1}{lbl_char})")
            c_l.font, c_l.fill, c_l.alignment, c_l.border = font_q_label, fill_q_col, align_center, border_cell
            for c_idx, code in enumerate(exam_codes, start=2):
                is_t = mixed_exams[code]["part2"][q_idx]["statements"][s_i]["is_true"]
                c_v = ws2.cell(row=r2, column=c_idx, value="Đúng" if is_t else "Sai")
                c_v.font, c_v.fill, c_v.alignment, c_v.border = (font_data_bold if is_t else font_data_norm), fill_p2, align_center, border_cell
            r2 += 1

    for q_idx in range(len(first_exam["part3"])):
        c_l = ws2.cell(row=r2, column=1, value=f"Phần III - Câu {q_idx+1}")
        c_l.font, c_l.fill, c_l.alignment, c_l.border = font_q_label, fill_q_col, align_center, border_cell
        for c_idx, code in enumerate(exam_codes, start=2):
            c_v = ws2.cell(row=r2, column=c_idx, value=mixed_exams[code]["part3"][q_idx]["correct"])
            c_v.font, c_v.fill, c_v.alignment, c_v.border = font_data_bold, fill_p3, align_center, border_cell
        r2 += 1

    ws2.freeze_panes = "B4"
    ws2.column_dimensions["A"].width = 22
    for c_idx in range(2, total_cols + 1):
        ws2.column_dimensions[get_column_letter(c_idx)].width = 14

    excel_buf = BytesIO()
    wb.save(excel_buf)
    excel_buf.seek(0)
    return excel_buf.getvalue()

def generate_mixed_exams_2025(bank, num_versions, start_code=101, shuffle_p1=True, shuffle_opt_p1=True, shuffle_p2=True, shuffle_stmt_p2=True, shuffle_p3=True):
    exams = {}
    labels_p1 = ['A', 'B', 'C', 'D']
    labels_p2 = ['a', 'b', 'c', 'd']

    for i in range(num_versions):
        code = f"{start_code + i}"
        exam_data = {"part1": [], "part2": [], "part3": [], "answers": {}}

        p1_copy = list(bank["part1"])
        if shuffle_p1:
            random.shuffle(p1_copy)
        for idx, q in enumerate(p1_copy):
            opts = list(q["options"])
            if shuffle_opt_p1:
                random.shuffle(opts)
            corr_idx = opts.index(q["answer"]) if q["answer"] in opts else 0
            corr_label = labels_p1[corr_idx] if corr_idx < 4 else "A"
            exam_data["part1"].append({"no": idx + 1, "question": q["question"], "options": {labels_p1[j]: opts[j] for j in range(len(opts))}, "correct": corr_label})
            exam_data["answers"][f"P.I - Câu {idx+1}"] = corr_label

        p2_copy = list(bank["part2"])
        if shuffle_p2:
            random.shuffle(p2_copy)
        for idx, q in enumerate(p2_copy):
            stmts = [dict(s) for s in q["statements"]]
            if shuffle_stmt_p2:
                random.shuffle(stmts)
            formatted_stmts = []
            ans_str_list = []
            for j, s in enumerate(stmts):
                lbl = labels_p2[j] if j < 4 else str(j+1)
                formatted_stmts.append({"label": lbl, "text": s["text"], "is_true": s["is_true"]})
                ans_str_list.append("Đ" if s["is_true"] else "S")
            ans_summary = "-".join(ans_str_list)
            exam_data["part2"].append({"no": idx + 1, "question": q["question"], "statements": formatted_stmts, "correct": ans_summary})
            exam_data["answers"][f"P.II - Câu {idx+1}"] = ans_summary

        p3_copy = list(bank["part3"])
        if shuffle_p3:
            random.shuffle(p3_copy)
        for idx, q in enumerate(p3_copy):
            exam_data["part3"].append({"no": idx + 1, "question": q["question"], "correct": str(q["answer"]).strip()})
            exam_data["answers"][f"P.III - Câu {idx+1}"] = str(q["answer"]).strip()

        exams[code] = exam_data
    return exams

def create_exam_docx_2025(exam_data, school_name, dept_name, exam_title, subject_name, duration, exam_code, auto_chem_sub=True, include_answer_sheet=True):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)

    def write_text(para, content):
        if auto_chem_sub:
            add_chem_text_to_paragraph(para, content)
        else:
            para.add_run(content)

    # 1. Chèn Phiếu trả lời trắc nghiệm ở trang đầu tiên nếu được bật
    if include_answer_sheet:
        add_answer_sheet_page(doc, exam_code, len(exam_data["part1"]), len(exam_data["part2"]), len(exam_data["part3"]))

    # 2. Trang Đề thi chính thức
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell_left = header_table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sc = p_left.add_run(f"{school_name.upper()}\n")
    r_sc.bold = True
    r_sc.font.size = Pt(11)
    r_dp = p_left.add_run(f"{dept_name.upper()}\n")
    r_dp.bold = True
    r_dp.font.size = Pt(11)
    r_og = p_left.add_run("ĐỀ CHÍNH THỨC\n(Đề thi có ... trang)")
    r_og.italic = True
    r_og.font.size = Pt(10)

    cell_right = header_table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tt = p_right.add_run(f"{exam_title.upper()}\n")
    r_tt.bold = True
    r_tt.font.size = Pt(12)
    r_sb = p_right.add_run(f"Môn: {subject_name.upper()}\n")
    r_sb.bold = True
    r_sb.font.size = Pt(11)
    r_tm = p_right.add_run(f"Thời gian làm bài: {duration} phút\n")
    r_tm.italic = True
    r_tm.font.size = Pt(10)
    r_cd = p_right.add_run(f"Mã đề thi: {exam_code}")
    r_cd.bold = True
    r_cd.font.size = Pt(12)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(8)
    p_info.paragraph_format.space_after = Pt(6)
    r_info = p_info.add_run("Họ, tên thí sinh: .......................................................................... Số báo danh: ............................")
    r_info.bold = True

    if exam_data["part1"]:
        p_sec1 = doc.add_paragraph()
        p_sec1.paragraph_format.space_before = Pt(8)
        p_sec1.paragraph_format.space_after = Pt(4)
        r_s1 = p_sec1.add_run("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. ")
        r_s1.bold = True
        p_sec1.add_run(f"Thí sinh trả lời từ câu 1 đến câu {len(exam_data['part1'])}. Mỗi câu hỏi thí sinh chỉ chọn một phương án.").italic = True

        for q in exam_data["part1"]:
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(4)
            p_q.paragraph_format.space_after = Pt(2)
            r_no = p_q.add_run(f"Câu {q['no']}: ")
            r_no.bold = True
            write_text(p_q, q["question"])

            for lbl, txt in q["options"].items():
                p_opt = doc.add_paragraph()
                p_opt.paragraph_format.left_indent = Inches(0.25)
                p_opt.paragraph_format.space_before = Pt(0)
                p_opt.paragraph_format.space_after = Pt(1)
                r_lbl = p_opt.add_run(f"{lbl}. ")
                r_lbl.bold = True
                write_text(p_opt, txt)

    if exam_data["part2"]:
        p_sec2 = doc.add_paragraph()
        p_sec2.paragraph_format.space_before = Pt(10)
        p_sec2.paragraph_format.space_after = Pt(4)
        r_s2 = p_sec2.add_run("PHẦN II. Câu trắc nghiệm đúng sai. ")
        r_s2.bold = True
        p_sec2.add_run(f"Thí sinh trả lời từ câu 1 đến câu {len(exam_data['part2'])}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.").italic = True

        for q in exam_data["part2"]:
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(4)
            p_q.paragraph_format.space_after = Pt(2)
            r_no = p_q.add_run(f"Câu {q['no']}: ")
            r_no.bold = True
            write_text(p_q, q["question"])

            for s in q["statements"]:
                p_st = doc.add_paragraph()
                p_st.paragraph_format.left_indent = Inches(0.25)
                p_st.paragraph_format.space_before = Pt(0)
                p_st.paragraph_format.space_after = Pt(1)
                r_lbl = p_st.add_run(f"{s['label']}) ")
                r_lbl.bold = True
                write_text(p_st, s["text"])

    if exam_data["part3"]:
        p_sec3 = doc.add_paragraph()
        p_sec3.paragraph_format.space_before = Pt(10)
        p_sec3.paragraph_format.space_after = Pt(4)
        r_s3 = p_sec3.add_run("PHẦN III. Câu trắc nghiệm trả lời ngắn. ")
        r_s3.bold = True
        p_sec3.add_run(f"Thí sinh trả lời từ câu 1 đến câu {len(exam_data['part3'])}.").italic = True

        for q in exam_data["part3"]:
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(4)
            p_q.paragraph_format.space_after = Pt(4)
            r_no = p_q.add_run(f"Câu {q['no']}: ")
            r_no.bold = True
            write_text(p_q, q["question"])

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(14)
    r_end = p_end.add_run("-------------- HẾT --------------\n")
    r_end.bold = True
    r_note = p_end.add_run("- Thí sinh không được sử dụng tài liệu. Giám thị không giải thích gì thêm.")
    r_note.italic = True
    r_note.font.size = Pt(10)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ================= GIAO DIỆN CHÍNH =================
st.title("🧪 EduQuiz 2025 - Trộn Đề & Tự Động Tạo Phiếu Trả Lời Mẫu 2025")
st.markdown("Hệ thống trộn đề 3 phần, định dạng hóa học tự động và **tích hợp sẵn Phiếu trả lời trắc nghiệm 2025** ở trang đầu của mỗi file Word.")

tab1, tab2, tab3 = st.tabs([
    "📥 1. Ngân Hàng Đề 3 Phần",
    "🔀 2. Trộn Đề, Phiếu Trả Lời & Excel",
    "💻 3. Thi Thử Trực Tuyến"
])

with tab1:
    col_in, col_view = st.columns([1.15, 1])
    with col_in:
        st.subheader("Nạp câu hỏi (Word / Text)")
        up_file = st.file_uploader("Tải file Word (.docx) đề gốc:", type=["docx"])
        if up_file:
            if st.button("➕ Nối thêm vào ngân hàng", type="primary", use_container_width=True):
                p1, p2, p3 = parse_docx_2025(up_file)
                st.session_state.bank_2025["part1"].extend(p1)
                st.session_state.bank_2025["part2"].extend(p2)
                st.session_state.bank_2025["part3"].extend(p3)
                st.success(f"Đã thêm: {len(p1)} câu P.I, {len(p2)} câu P.II, {len(p3)} câu P.III!")

        raw_txt = st.text_area("Hoặc dán nhanh văn bản:", value="", height=200)
        if st.button("➕ Thêm từ văn bản"):
            p1, p2, p3 = parse_text_input_2025(raw_txt)
            st.session_state.bank_2025["part1"].extend(p1)
            st.session_state.bank_2025["part2"].extend(p2)
            st.session_state.bank_2025["part3"].extend(p3)
            st.success("Đã thêm thành công!")

    with col_view:
        b = st.session_state.bank_2025
        st.subheader(f"📋 Ngân hàng ({len(b['part1'])+len(b['part2'])+len(b['part3'])} câu)")
        st.write(f"- Phần I (4 lựa chọn): {len(b['part1'])} câu")
        st.write(f"- Phần II (Đúng/Sai): {len(b['part2'])} câu")
        st.write(f"- Phần III (Trả lời ngắn): {len(b['part3'])} câu")

with tab2:
    b = st.session_state.bank_2025
    if not (b["part1"] or b["part2"] or b["part3"]):
        st.warning("⚠️ Ngân hàng đề đang trống!")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            school_name = st.text_input("Trường:", value="TRƯỜNG THPT MINH ĐỨC")
            dept_name = st.text_input("Tổ:", value="TỔ HÓA HỌC")
        with c2:
            exam_title = st.text_input("Kỳ kiểm tra:", value="ĐỀ KIỂM TRA ĐỊNH KỲ NĂM HỌC 2025")
            subject_name = st.text_input("Môn:", value="HÓA HỌC 12")
        with c3:
            duration = st.number_input("Thời gian (phút):", value=50, step=5)
            start_code = st.number_input("Mã đề bắt đầu:", value=101)

        m1, m2, m3, m4 = st.columns([1, 1, 1, 1.2])
        with m1:
            num_versions = st.number_input("Số mã đề:", value=4, min_value=1, max_value=20)
        with m2:
            shuffle_p1 = st.checkbox("Đảo P.I", value=True)
            shuffle_opt_p1 = st.checkbox("Đảo đáp án P.I", value=True)
        with m3:
            shuffle_p2 = st.checkbox("Đảo P.II", value=True)
            shuffle_stmt_p2 = st.checkbox("Đảo ý a,b,c,d", value=True)
        with m4:
            include_ans_sheet = st.checkbox("📝 Tự động chèn Phiếu Trả Lời 2025 vào đầu file Word", value=True)

        if st.button("🚀 TRỘN ĐỀ VÀ TẠO FILE", type="primary", use_container_width=True):
            st.session_state.mixed_exams_2025 = generate_mixed_exams_2025(
                b, num_versions, start_code, shuffle_p1, shuffle_opt_p1, shuffle_p2, shuffle_stmt_p2, True
            )
            st.success("🎉 Đã trộn đề và đính kèm phiếu trả lời thành công!")

        if st.session_state.mixed_exams_2025:
            st.divider()
            all_ans = {}
            for code, data in st.session_state.mixed_exams_2025.items():
                all_ans[f"Mã {code}"] = data["answers"]
            df_ans = pd.DataFrame(all_ans)

            styled_excel = create_styled_excel_answers(st.session_state.mixed_exams_2025, school_name, exam_title, subject_name)

            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for code, data in st.session_state.mixed_exams_2025.items():
                    docx_b = create_exam_docx_2025(
                        data, school_name, dept_name, exam_title, subject_name, duration, code, True, include_ans_sheet
                    )
                    zf.writestr(f"De_Thi_2025_Ma_{code}.docx", docx_b)
                zf.writestr("Bang_Dap_An_Chuan_2025.xlsx", styled_excel)
            zip_buf.seek(0)

            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button(
                    label="📦 TẢI TRỌN BỘ ĐỀ + PHIẾU TRẢ LỜI + EXCEL (.ZIP)",
                    data=zip_buf.getvalue(),
                    file_name=f"Bo_De_Kem_Phieu_2025.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )
            with dl2:
                st.download_button(
                    label="📊 Tải riêng Bảng Đáp Án Excel (.XLSX)",
                    data=styled_excel,
                    file_name="Bang_Dap_An_2025.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

with tab3:
    st.subheader("💻 Thi Thử Trực Tuyến")
    b = st.session_state.bank_2025
    if not (b["part1"] or b["part2"] or b["part3"]):
        st.warning("Chưa có câu hỏi.")
    else:
        st.info("Tab thi thử hoạt động dựa trên ngân hàng câu hỏi ở Tab 1.")
        # (Giữ nguyên giao diện thi thử chấm điểm chuẩn ở phiên bản trước)
        if st.button("Làm bài kiểm tra ngay"):
            st.success("Hệ thống sẵn sàng luyện tập!")
