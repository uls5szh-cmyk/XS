# -*- coding: utf-8 -*-
"""
=============================================================================
BOSCH | PCB Lesson Learn Quality Studio (Direct Table-Cell Splitting Edition)
- Real Native Cell Splitting: N images = N parallel blue-header blocks
- True Symmetrical Center Alignment: Clears cell margins, 100% centered
- Actionable High-Contrast UI & Exact New Email Template
=============================================================================
"""

import streamlit as st
import pandas as pd
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import datetime
import io
import os
import glob
import zipfile
import re
import copy
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# 尝试导入交互式图表库 Plotly
try:
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# -----------------------------------------------------------------------------
# 1. 页面基本配置与博世高端工业视觉体系 (Bosch Corporate Identity 2.0)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Bosch | PCB Lesson Learn Quality Studio",
    layout="wide",
    page_icon="🔴"
)

BOSCH_UI_STYLE = """
<style>
    :root {
        --bosch-red: #E20015;
        --bosch-blue: #005691;
        --bosch-light-blue: #007BC0;
        --bosch-cyan: #008ECF;
        --bosch-green: #78BE20;
        --bosch-dark-gray: #1C2B39;
        --bosch-gray: #525F6B;
        --bosch-bg: #EAEFF4;
        --bosch-card-bg: #F4F6F8;
        --bosch-border: #CBD5E1;
        --bosch-active-border: #005691;
    }
    
    .stApp { 
        background-color: var(--bosch-bg); 
        font-family: 'Segoe UI', 'Arial', sans-serif; 
    }
    
    .bosch-top-bar {
        height: 6px;
        background: linear-gradient(90deg, #E20015 0%, #E20015 25%, #005691 25%, #005691 65%, #007BC0 65%, #007BC0 100%);
        border-radius: 3px;
        margin-bottom: 20px;
    }
    
    .bds-step-card {
        background: var(--bosch-card-bg);
        border: 1px solid var(--bosch-border);
        border-radius: 8px;
        padding: 22px;
        margin-bottom: 20px;
    }
    
    .step-header {
        display: flex;
        align-items: center;
        margin-bottom: 14px;
        border-bottom: 1px solid #D8E0E8;
        padding-bottom: 10px;
    }
    .step-number {
        background: var(--bosch-blue);
        color: #FFFFFF;
        font-size: 0.9rem;
        font-weight: 800;
        padding: 4px 12px;
        border-radius: 4px;
        margin-right: 12px;
    }
    .step-title {
        color: #005691;
        font-size: 1.15rem;
        font-weight: 700;
        margin: 0;
    }
    
    .stTextInput input, .stSelectbox div[data-baseweb="select"] > div, .stTextArea textarea {
        background-color: #FFFFFF !important;
        border: 1.5px solid #94A3B8 !important;
        border-radius: 6px !important;
        color: #0F172A !important;
        font-size: 0.95rem !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05) !important;
    }
    
    .stTextInput input:hover, .stTextArea textarea:hover, .stSelectbox div[data-baseweb="select"] > div:hover {
        border-color: var(--bosch-light-blue) !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus, .stSelectbox div[data-baseweb="select"] > div:focus-within {
        border-color: var(--bosch-blue) !important;
        box-shadow: 0 0 0 3px rgba(0, 86, 145, 0.18) !important;
    }
    
    .stWidgetLabel p {
        font-weight: 700 !important;
        color: #1E293B !important;
        font-size: 0.92rem !important;
    }
    
    .clean-divider {
        border: 0;
        height: 1px;
        background: #CBD5E1;
        margin: 20px 0;
    }
    
    .guide-box {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-left: 4px solid var(--bosch-blue);
        border-radius: 6px;
        padding: 14px 18px;
        font-size: 0.88rem;
        line-height: 1.6;
        color: #334155;
    }
    
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid var(--bosch-border);
        border-radius: 8px;
        padding: 14px 18px;
        border-top: 4px solid var(--bosch-blue);
        box-shadow: 0 2px 6px rgba(0, 40, 80, 0.03);
    }
    .kpi-title { font-size: 0.8rem; color: #525F6B; font-weight: 600; text-transform: uppercase; }
    .kpi-value { font-size: 1.8rem; color: #005691; font-weight: 700; margin-top: 2px; }
    
    .inspector-panel {
        background: #FFFFFF;
        border: 1.5px solid #94A3B8;
        border-left: 5px solid var(--bosch-blue);
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0, 40, 80, 0.04);
    }
    
    .badge-completed {
        background-color: #E8F5E9;
        color: #2E7D32;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.85rem;
        display: inline-block;
    }
    .badge-pending {
        background-color: #FFEBEE;
        color: #C62828;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.85rem;
        display: inline-block;
    }
    
    .stButton>button {
        background-color: var(--bosch-blue) !important;
        color: white !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
        border: none !important;
        padding: 8px 20px !important;
    }
    .stButton>button:hover {
        background-color: var(--bosch-light-blue) !important;
        color: white !important;
    }
</style>
<div class="bosch-top-bar"></div>
"""
st.markdown(BOSCH_UI_STYLE, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 跨平台自动扫描与数据源匹配
# -----------------------------------------------------------------------------
BASE_G_DIR = r"G:\02_7_M-PQA-RBAC1\08_PQA_AE\09_PQA2\11_PCB\04_Lessons learn"
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
TEAMS_BOT_URL = "https://teams.microsoft.com/l/app/ffcadcc0-464f-4110-a065-0e3b4733baa9?source=bot-header-share-entrypoint"

def scan_all_master_lists():
    search_dirs = [BASE_G_DIR, CURRENT_DIR, os.getcwd()]
    found_files = []
    seen = set()
    for s_dir in search_dirs:
        if os.path.exists(s_dir):
            for ext in ["*.xlsx", "*.xlsm"]:
                pattern = os.path.join(s_dir, f"PCB Lesson Learn Master List*{ext}")
                for fpath in glob.glob(pattern):
                    real_path = os.path.abspath(fpath)
                    if real_path not in seen and not os.path.basename(real_path).startswith("~$"):
                        seen.add(real_path)
                        found_files.append(real_path)
    found_files.sort(key=lambda x: os.path.getmtime(x) if os.path.exists(x) else 0, reverse=True)
    return found_files

def resolve_exact_file(filename_list):
    search_dirs = [BASE_G_DIR, CURRENT_DIR, os.getcwd()]
    for fname in filename_list:
        for s_dir in search_dirs:
            target = os.path.join(s_dir, fname)
            if os.path.exists(target):
                return target
        if os.path.exists(fname):
            return fname
    return None

st.sidebar.markdown("### ⚙️ 系统数据源配置")

available_master_lists = scan_all_master_lists()
if available_master_lists:
    chosen_excel = st.sidebar.selectbox(
        "📊 Master List 版本:",
        options=available_master_lists,
        format_func=lambda x: f"{os.path.basename(x)} ({datetime.datetime.fromtimestamp(os.path.getmtime(x)).strftime('%m/%d %H:%M')})"
    )
    excel_path = chosen_excel
else:
    excel_path = st.sidebar.text_input("1. Master List 表格路径:", os.path.join(BASE_G_DIR, "PCB Lesson Learn Master List.xlsx"))

resolved_template = resolve_exact_file([
    "Lessons Learned Report Problem Solving.docx",
    "Blank LL Template complete version.docx",
    "LL Template complete version.docx"
])
template_path = st.sidebar.text_input("2. Word 模板路径:", resolved_template if resolved_template else os.path.join(BASE_G_DIR, "Lessons Learned Report Problem Solving.docx"))

resolved_feedback = resolve_exact_file(["LL Feedback table_Supplier version_V1.xlsx"])
feedback_path = st.sidebar.text_input("3. Feedback 表格路径:", resolved_feedback if resolved_feedback else os.path.join(BASE_G_DIR, "LL Feedback table_Supplier version_V1.xlsx"))

excel_file = excel_path if os.path.exists(excel_path) else None
template_file = template_path if os.path.exists(template_path) else None
feedback_file = feedback_path if os.path.exists(feedback_path) else None

# -----------------------------------------------------------------------------
# 3. 页面模式导航菜单定义
# -----------------------------------------------------------------------------
app_mode = st.radio(
    "👉 请选择工作模式 (Work Mode):",
    options=["📑 FEBER 报告生成与邮件协同", "📊 闭环看板 (Executive Dashboard)"],
    horizontal=True
)

st.markdown('<hr class="clean-divider">', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 4. 辅助数据处理与图片提取函数
# -----------------------------------------------------------------------------

def load_supplier_emails(file_source):
    try:
        if hasattr(file_source, 'seek'): file_source.seek(0)
        xl = pd.ExcelFile(file_source)
        if 'Vendor code' not in xl.sheet_names:
            return {}
        if hasattr(file_source, 'seek'): file_source.seek(0)
        df_vendor = pd.read_excel(file_source, sheet_name='Vendor code')
        df_vendor.columns = df_vendor.columns.astype(str).str.strip()
        
        if 'Supplier' in df_vendor.columns and 'Name' in df_vendor.columns:
            df_vendor['Supplier'] = df_vendor['Supplier'].ffill()
            supplier_dict = {}
            for _, row in df_vendor.iterrows():
                supplier = str(row['Supplier']).strip()
                name_val = str(row['Name'])
                if pd.isna(row['Supplier']) or supplier in ['nan', 'None']:
                    continue
                emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', name_val)
                if emails:
                    if supplier not in supplier_dict:
                        supplier_dict[supplier] = set()
                    for email in emails:
                        supplier_dict[supplier].add(email)
            for k in supplier_dict:
                supplier_dict[k] = list(supplier_dict[k])
            return supplier_dict
    except Exception:
        pass
    return {}

def get_multiple_images_for_row(file_source, sheet_name, header_idx, target_row_idx):
    try:
        if hasattr(file_source, 'seek'): file_source.seek(0)
        wb = openpyxl.load_workbook(file_source, data_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        
        col_ng = -1
        col_ok = -1
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=header_idx + 1, column=col_idx).value
            if val:
                val_str = str(val).strip().lower()
                if 'ng picture' in val_str or 'picture' in val_str: col_ng = col_idx - 1
                if 'ok picture' in val_str: col_ok = col_idx - 1
                
        excel_target_row = header_idx + 1 + target_row_idx
        ok_imgs = []
        ng_imgs = []
        
        for img in getattr(ws, '_images', []):
            try:
                r = img.anchor._from.row
                c = img.anchor._from.col
                if r == excel_target_row:
                    if c == col_ng or (col_ng != -1 and abs(c - col_ng) <= 1):
                        ng_imgs.append(img._data())
                    elif c == col_ok:
                        ok_imgs.append(img._data())
            except Exception:
                pass
                
        return ok_imgs, ng_imgs
    except Exception as e:
        return [], []

def load_excel_robust(file_source):
    if hasattr(file_source, 'seek'): file_source.seek(0)
    xl = pd.ExcelFile(file_source)
    sheet_names = xl.sheet_names
    
    target_sheet = next((s for s in sheet_names if "PUQ3 LL Overall list" in s), None)
    if not target_sheet:
        target_sheet = next((s for s in sheet_names if "Overall" in s or "LL" in s), sheet_names[0])
        
    if hasattr(file_source, 'seek'): file_source.seek(0)
    df_temp = pd.read_excel(file_source, sheet_name=target_sheet, nrows=10, header=None)
    header_idx = 1 
    for idx, row in df_temp.iterrows():
        row_str = [str(val).strip().lower() for val in row.tolist()]
        if any('serials' in val or 'project/part' in val or 'failure mode' in val for val in row_str):
            header_idx = idx
            break
            
    if hasattr(file_source, 'seek'): file_source.seek(0)
    df = pd.read_excel(file_source, sheet_name=target_sheet, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()
    return df, target_sheet, header_idx

def parse_bot_feber_response(bot_text):
    parsed = {
        'Abstract_Issue': '', 'Abstract_Problem': '', 'Abstract_Lessons': '',
        'Product_Process': '', 'Component': '', 'Sub_Component': '',
        'Problem': '', 'Lessons_Rows': [],
        'What_Else': '', 'Where': '', 'When': '', 'Who': ''
    }
    m_abs = re.search(r'(?:0\.\s*Abstract|Abstract)\s*([\s\S]*?)(?=1\.\s*Product|$)', bot_text, re.I)
    if m_abs:
        t = m_abs.group(1)
        i_m = re.search(r'Issue:\s*([\s\S]*?)(?=Problem:|$)', t, re.I)
        p_m = re.search(r'Problem:\s*([\s\S]*?)(?=Lessons:|$)', t, re.I)
        l_m = re.search(r'Lessons:\s*([\s\S]*?)(?=Tags:|Picture|1\.\s*Product|$)', t, re.I)
        if i_m: parsed['Abstract_Issue'] = i_m.group(1).strip()
        if p_m: parsed['Abstract_Problem'] = p_m.group(1).strip()
        if l_m: parsed['Abstract_Lessons'] = l_m.group(1).strip()
    
    m_p = re.search(r'1\.\s*Product\s*/\s*Process\s*([\s\S]*?)(?=2\.\s*Problem|$)', bot_text, re.I)
    if m_p:
        t = m_p.group(1)
        pp = re.search(r'Product\s*/\s*Process:\s*([^\n]+)', t, re.I)
        cp = re.search(r'Component:\s*([^\n]+)', t, re.I)
        sc = re.search(r'Sub-Component:\s*([^\n]+)', t, re.I)
        if pp: parsed['Product_Process'] = pp.group(1).strip()
        if cp: parsed['Component'] = cp.group(1).strip()
        if sc: parsed['Sub_Component'] = sc.group(1).strip()
        
    m_prob = re.search(r'2\.\s*Problem[^\n]*\n([\s\S]*?)(?=3\.\s*Lessons|$)', bot_text, re.I)
    if m_prob:
        parsed['Problem'] = m_prob.group(1).strip()
    
    m_less = re.search(r'3\.\s*Lessons[^\n]*\n([\s\S]*?)(?=4\.\s*Potentially|$)', bot_text, re.I)
    if m_less:
        less_text = m_less.group(1).strip()
        raw_lines = [l.strip() for l in less_text.split('\n') if l.strip()]
        for line in raw_lines:
            if "measures & sustainable solutions" in line.lower() or "---" in line or line.startswith("| :---") or line.startswith("Lessons\t"):
                continue
            if '\t' in line:
                parts = [p.strip() for p in line.split('\t')]
                if len(parts) >= 3: parsed['Lessons_Rows'].append((parts[0], parts[1], parts[2]))
                elif len(parts) == 2: parsed['Lessons_Rows'].append((parts[0], parts[1], ""))
                elif len(parts) == 1: parsed['Lessons_Rows'].append((parts[0], "", ""))
            elif '|' in line:
                parts = [p.strip() for p in line.split('|')[1:-1]]
                if len(parts) >= 3: parsed['Lessons_Rows'].append((parts[0], parts[1], parts[2]))
                elif len(parts) == 2: parsed['Lessons_Rows'].append((parts[0], parts[1], ""))
        if not parsed['Lessons_Rows']:
            parsed['Lessons_Rows'].append((less_text, "", ""))
            
    m_pot = re.search(r'4\.\s*Potentially affected[^\n]*\n([\s\S]*?)(?=5\.\s*Appendix|$)', bot_text, re.I)
    if m_pot:
        t = m_pot.group(1)
        w1 = re.search(r'What else[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        w2 = re.search(r'Where can[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        w3 = re.search(r'When can[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        w4 = re.search(r'Who else[^\n\t|]*[\t\|\n]([^\n|]+)', t, re.I)
        if w1: parsed['What_Else'] = w1.group(1).strip()
        if w2: parsed['Where'] = w2.group(1).strip()
        if w3: parsed['When'] = w3.group(1).strip()
        if w4: parsed['Who'] = w4.group(1).strip()
        
    return parsed

# -----------------------------------------------------------------------------
# 5. 精准装配 Word 模板核心函数 (原版方块蓝底、绝对居中、并列方块)
# -----------------------------------------------------------------------------

def set_cell_formatted_text(cell, text):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(str(text).strip())
    run.font.name = 'Arial'
    run.font.size = Pt(10.5)
    run.font.bold = False

def set_aligned_field_paragraph(p, label, value, indent_inches=1.0):
    p.text = ""
    p.paragraph_format.left_indent = Inches(indent_inches)
    p.paragraph_format.first_line_indent = Inches(-indent_inches)
    p.paragraph_format.tab_stops.add_tab_stop(Inches(indent_inches), WD_TAB_ALIGNMENT.LEFT)
    r_label = p.add_run(label)
    r_label.font.name = 'Arial'
    r_label.font.size = Pt(10.5)
    r_label.font.bold = True
    p.add_run("\t")
    r_val = p.add_run(str(value).strip())
    r_val.font.name = 'Arial'
    r_val.font.size = Pt(10.5)
    r_val.font.bold = False

def replace_field_value_in_doc(doc, field_label, new_value, is_abstract=False):
    if not new_value: return
    for p in doc.paragraphs:
        if field_label.lower() in p.text.lower():
            if is_abstract:
                set_aligned_field_paragraph(p, field_label, new_value, indent_inches=0.9)
            else:
                p.text = ""
                r_label = p.add_run(f"{field_label} ")
                r_label.font.name = 'Arial'
                r_label.font.size = Pt(10.5)
                r_label.font.bold = True
                r_val = p.add_run(str(new_value).strip())
                r_val.font.name = 'Arial'
                r_val.font.size = Pt(10.5)
                r_val.font.bold = False
            return

def write_problem_compact_and_clean_pagebreaks(doc, text_value):
    """紧凑写入第2章内容，彻底清除多余空行与隐藏分页符"""
    if not text_value:
        return
        
    paragraphs = doc.paragraphs
    prob_idx = -1
    for idx, p in enumerate(paragraphs):
        p_txt = p.text.strip().lower()
        if "problem" in p_txt and ("fundamental" in p_txt or p_txt.startswith("2")):
            prob_idx = idx
            break
            
    if prob_idx != -1:
        if prob_idx + 1 < len(paragraphs):
            target_p = paragraphs[prob_idx + 1]
            target_p.text = ""
            target_p.paragraph_format.space_before = Pt(2)
            target_p.paragraph_format.space_after = Pt(4)
            target_p.paragraph_format.line_spacing = 1.05
            r = target_p.add_run(str(text_value).strip())
            r.font.name = 'Arial'
            r.font.size = Pt(10.5)
            r.font.bold = False
            
            scan_idx = prob_idx + 2
            while scan_idx < len(paragraphs):
                p_curr = paragraphs[scan_idx]
                p_curr_txt = p_curr.text.strip().lower()
                if "3." in p_curr_txt or "lessons" in p_curr_txt:
                    break
                for br in p_curr._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'):
                    if br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
                        br.getparent().remove(br)
                if len(p_curr_txt) < 80:
                    p_curr.text = ""
                    p_curr.paragraph_format.space_before = Pt(0)
                    p_curr.paragraph_format.space_after = Pt(0)
                scan_idx += 1

def build_independent_blue_block(cell, block_title, img_bytes, img_width_inches):
    """
    【构建原生独立博世蓝底方块】：
    - 纯博世深蓝底色顶栏 (#005691) + 白色加粗标题 (Picture #1...)
    - 下方直接紧凑嵌入不良图片，段落左右缩进归零并严格几何居中 (Center Alignment)
    - 彻底解除向右堆挤，形成对称规整的独立框体
    """
    cell.text = ""
    
    # 彻底清除单元格内部一切边缘缩进，确保完全居中
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:left w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)
    
    # 1. 顶部深蓝标题框
    p_title = cell.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.left_indent = Inches(0)
    p_title.paragraph_format.right_indent = Inches(0)
    
    # 顶部深蓝遮罩底色
    pPr = p_title._p.get_or_add_pPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="005691"/>')
    pPr.append(shd)
    
    r_t = p_title.add_run(f" {block_title} ")
    r_t.font.name = 'Arial'
    r_t.font.size = Pt(8.5)
    r_t.font.bold = True
    r_t.font.color.rgb = RGBColor(255, 255, 255) # 纯白文字
    
    # 2. 居中置入实物不良图片
    p_img = cell.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(3)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.line_spacing = 1.0
    p_img.paragraph_format.left_indent = Inches(0)
    p_img.paragraph_format.right_indent = Inches(0)
    
    p_img.add_run().add_picture(io.BytesIO(img_bytes), width=img_width_inches)

def populate_docx_exact_tables(template_source, bot_data, raw_row, ok_imgs=None, ng_imgs=None):
    """
    【原生独立蓝底方块并列排版】
    - 如果有 1 张图片：当前蓝底方块直接居中呈现该图片；
    - 如果有 2 张及以上图片：直接在单元格内部构建并排铺满的纯净同级子矩阵，平分宽度，居中对齐，绝不向右挤压！
    """
    if hasattr(template_source, 'seek'):
        template_source.seek(0)
    doc = docx.Document(template_source)
    current_date_str = datetime.date.today().strftime('%b %d %Y')
    failure_mode_str = str(raw_row.get('Failure Mode', '')).strip()

    for p in doc.paragraphs:
        if 'May 24 2022' in p.text or 'May 24, 2022' in p.text:
            p.text = p.text.replace('May 24 2022', current_date_str).replace('May 24, 2022', current_date_str)
        if 'lesson learn' in p.text.lower() and len(p.text) < 60:
            if failure_mode_str and failure_mode_str not in p.text:
                p_clean = p.text.strip().rstrip("–-—: ").strip()
                p.text = f"{p_clean} – {failure_mode_str}"

    abs_issue = bot_data.get('Abstract_Issue') or raw_row.get('LL Brief Description', '')
    abs_prob = bot_data.get('Abstract_Problem') or raw_row.get('Failure Mode', '')
    abs_less = bot_data.get('Abstract_Lessons') or raw_row.get('Should or not to do', '')
    
    replace_field_value_in_doc(doc, "Issue:", abs_issue, is_abstract=True)
    replace_field_value_in_doc(doc, "Problem:", abs_prob, is_abstract=True)
    replace_field_value_in_doc(doc, "Lessons:", abs_less, is_abstract=True)

    pp_val = bot_data.get('Product_Process') or raw_row.get('Related Material Field / Process', '')
    comp_val = bot_data.get('Component') or raw_row.get('Project/Part name', '')
    sub_val = bot_data.get('Sub_Component') or 'Not specified'
    
    replace_field_value_in_doc(doc, "Product / Process:", pp_val)
    replace_field_value_in_doc(doc, "Component:", comp_val)
    replace_field_value_in_doc(doc, "Sub-Component:", sub_val)

    # 紧凑写入第2章内容，彻底清除多余空行与隐藏分页符
    prob_val = bot_data.get('Problem') or raw_row.get('LL Brief Description', '')
    write_problem_compact_and_clean_pagebreaks(doc, prob_val)

    # =========================================================================
    # 原生独立蓝底方块排版逻辑（消除右挤，严格几何居中）
    # =========================================================================
    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                c_txt = cell.text.lower().replace(" ", "")
                if ("picture" in c_txt or "defect" in c_txt or "not-ok" in c_txt) and "ok-part" not in c_txt:
                    if ng_imgs:
                        count = len(ng_imgs)
                        if count == 1:
                            # 1张图：原版蓝底方块居中呈现
                            build_independent_blue_block(cell, "Picture – Product – Defect", ng_imgs[0], Inches(2.2))
                        else:
                            # 2张及以上：建立 N 列并列纯净矩阵，平分列宽，彻底居中
                            cell.text = ""
                            # 移除外层单元格背景色干扰
                            tcPr = cell._tc.get_or_add_tcPr()
                            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>')
                            tcPr.append(shd)
                            
                            sub_tbl = cell.add_table(rows=1, cols=count)
                            sub_tbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            # 强制子表格居中对齐，设置零外边距
                            tblPr = sub_tbl._tbl.tblPr
                            tblPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
                            tblPr.append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:left w:w="30" w:type="dxa"/><w:right w:w="30" w:type="dxa"/></w:tblCellMar>'))
                            
                            single_col_w = 3.05 / count
                            img_render_w = Inches(single_col_w * 0.88)
                            
                            for i in range(count):
                                sub_cell = sub_tbl.cell(0, i)
                                sub_cell.width = Inches(single_col_w)
                                # 每个框体独立绘制博世蓝底顶栏与居中图片！
                                build_independent_blue_block(sub_cell, f"Picture #{i+1}", ng_imgs[i], img_render_w)

        # 3. Lessons 表格动态多行装配
        t_header = "".join(cell.text for cell in table.rows[0].cells).lower()
        if "lessons" in t_header and ("measures" in t_header or "root cause" in t_header):
            lessons_rows = bot_data.get('Lessons_Rows', [])
            if not lessons_rows:
                lessons_rows = [(
                    raw_row.get('Should or not to do', ''),
                    raw_row.get('Corrective Action', ''),
                    raw_row.get('Root Cause', '')
                )]
            while len(table.rows) > 1:
                tr = table.rows[-1]._tr
                table._tbl.remove(tr)
            for row_tuple in lessons_rows:
                new_row = table.add_row()
                for i_c in range(min(3, len(row_tuple))):
                    set_cell_formatted_text(new_row.cells[i_c], row_tuple[c_idx] if i_c < len(row_tuple) else "")

        # 4. Potentially affected 表格填充
        elif "what else" in t_header or "potentially" in t_header or len(table.rows) == 4:
            w_map = {
                0: bot_data.get('What_Else') or raw_row.get('What else could be additionally affected?') or 'Other PCB suppliers manufacturing multilayer boards using similar pattern plating processes.',
                1: bot_data.get('Where') or raw_row.get('Where can the problem additionally occur?') or 'Other production plants or lines employing similar pattern plating technologies worldwide.',
                2: bot_data.get('When') or raw_row.get('When can the problem additionally appear?') or 'Future PCB product lines or applications with comparable design/process constraints.',
                3: bot_data.get('Who') or raw_row.get('Who else can be affected?') or 'Downstream customers, assembly plants, and Bosch divisions reliant on affected PCBs.'
            }
            for r_i, row in enumerate(table.rows):
                if len(row.cells) >= 2 and r_i in w_map:
                    set_cell_formatted_text(row.cells[1], w_map[r_i])

    return doc

# -----------------------------------------------------------------------------
# 6. 生成新版邮件模板 (双附件 + 动态变量替换)
# -----------------------------------------------------------------------------

def build_new_email_dual_attachment(row_data, to_emails="", doc_bytes=None, doc_filename="LL_Template.docx", feedback_bytes=None, feedback_filename="LL Feedback table_Supplier version_V1.xlsx"):
    serial_no_val = str(row_data.get('LL Serials No', '')).strip()
    if not serial_no_val or serial_no_val in ['nan', 'None']:
        for k in row_data.keys():
            if 'serial' in str(k).lower():
                serial_no_val = str(row_data[k]).strip()
                break
    if not serial_no_val or serial_no_val in ['nan', 'None']:
        serial_no_val = "LL-Export"
        
    failure_mode_val = str(row_data.get('Failure Mode', '*****')).strip()
    if failure_mode_val in ['nan', 'None']:
        failure_mode_val = "*****"
        
    subject = f"M/PQR-AP LL | {serial_no_val} | Title {failure_mode_val}"
    
    html_body = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Segoe UI', 'Arial', sans-serif; font-size: 10.5pt; line-height: 1.6; color: #333333; }}
            .red-bold {{ color: #E20015; font-weight: bold; }}
            ul {{ margin-top: 5px; margin-bottom: 15px; padding-left: 20px; }}
            li {{ margin-bottom: 8px; }}
        </style>
    </head>
    <body>
        <p>Dear Supplier:</p>
        <p>Recently, we summarized a lesson learn about <strong>{failure_mode_val}</strong>. Please review the attached LL document(attachment 1) and complete following tasks:</p>
        <ul>
            <li>Complete feedback form ( attachment2) based on self-evaluation on your own processes and send to your responsible contactor within <span class="red-bold">one week</span>.</li>
            <li>After your self-evaluation, please close defined actions within <span class="red-bold">defined deadline</span>.</li>
            <li>Our colleague may conduct onsite verification according to the information in feedback form in <span class="red-bold">if necessary</span>.</li>
        </ul>
        <p>If you have any question about this lesson learn, please contact us freely.</p>
        <br>
        <p>Best regards,</p>
        <p><strong>Purchasing Quality Region Asia Pacific Team</strong><br>
        Robert Bosch GmbH</p>
    </body>
    </html>
    """
    
    msg = MIMEMultipart('mixed')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = 'Sunny.LIU3@cn.bosch.com'
    msg['To'] = to_emails
    msg.add_header('X-Unsent', '1')
    
    alt_part = MIMEMultipart('alternative')
    alt_part.attach(MIMEText(html_body, 'html', 'utf-8'))
    msg.attach(alt_part)
    
    if doc_bytes:
        part_doc = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
        part_doc.set_payload(doc_bytes)
        encoders.encode_base64(part_doc)
        part_doc.add_header('Content-Disposition', f'attachment; filename="{doc_filename}"')
        msg.attach(part_doc)
        
    if feedback_bytes:
        part_fb = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part_fb.set_payload(feedback_bytes)
        encoders.encode_base64(part_fb)
        part_fb.add_header('Content-Disposition', f'attachment; filename="{feedback_filename}"')
        msg.attach(part_fb)
        
    return msg.as_bytes(), subject

# -----------------------------------------------------------------------------
# 7. 主交互数据流处理与绝对物理清洗逻辑
# -----------------------------------------------------------------------------

if excel_file is not None and template_file is not None:
    try:
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        supplier_dict = load_supplier_emails(excel_file)
        
        serial_no_col = next((c for c in df.columns if 'serial' in str(c).lower()), 'LL Serials No')
        supplier_scope_col = next((c for c in df.columns if 'scope' in str(c).lower() or 'task' in str(c).lower()), 'LL Supplier Scope')
        need_col = next((c for c in df.columns if 'need or not' in str(c).lower() or 'need' in str(c).lower()), None)
        
        # 1. 业务前提一：LL Need or not 必须为 Y
        if need_col:
            df = df[df[need_col].astype(str).str.strip().str.upper() == 'Y'].copy()

        # 2. 物理级底层穿透：精准锁定 Complete or not 列
        complete_col = None
        for col in df.columns:
            col_str = str(col).strip()
            if col_str == 'Complete or not' or col_str.replace('\n', ' ') == 'Complete or not' or col_str.replace('\r', ' ') == 'Complete or not':
                complete_col = col
                break
                
        if not complete_col:
            for col in df.columns:
                c_low = str(col).lower().replace(' ', '').replace('_', '').replace('\n', '').replace('\r', '')
                if 'completeornot' in c_low or ('complete' in c_low and 'need' not in c_low):
                    complete_col = col
                    break
                    
        def parse_completion_strict(val):
            if pd.isna(val): return 'Pending'
            v_clean = str(val).replace('\n', '').replace('\r', '').replace('\t', '').replace(' ', '').strip().upper()
            if 'Y' in v_clean or 'YES' in v_clean or 'TRUE' in v_clean or v_clean == '1' or '100' in v_clean:
                return 'Completed'
            return 'Pending'
            
        if complete_col:
            df['Normalized_Status'] = df[complete_col].apply(parse_completion_strict)
        else:
            df['Normalized_Status'] = 'Pending'

        # =========================================================================
        # 模式一：FEBER 报告生成与邮件协同
        # =========================================================================
        if app_mode == "📑 FEBER 报告生成与邮件协同":
            feedback_bytes = None
            feedback_filename = "LL Feedback table_Supplier version_V1.xlsx"
            if feedback_file is not None and os.path.exists(feedback_file):
                with open(feedback_file, 'rb') as f:
                    feedback_bytes = f.read()
                    
            gen_df = df.copy()
            
            with st.expander("📖 查看 FEBER 协同与邮件分发操作规范 (Operation Guide)", expanded=False):
                st.markdown("""
                <div class="guide-box">
                    <strong>📌 闭环协同标准化作业规范：</strong><br>
                    1. <strong>STEP 01 选取事实：</strong>从过滤后的清单中选定一条失效模式记录，系统会自动抽提 100% 原始事实。<br>
                    2. <strong>STEP 02 AI 润色：</strong>点击直达 Teams M-PU Bot，发送生成的 Prompt，获取符合 FEBER 规范的润色结果。<br>
                    3. <strong>STEP 03 交付闭环：</strong>粘贴回复内容，一键生成标准 Word 报告（多图独立专属蓝底框排版）与已挂载双附件的新版邮件草稿。
                </div>
                """, unsafe_allow_html=True)
            
            # STEP 01
            st.markdown("""
            <div class="bds-step-card">
                <div class="step-header">
                    <span class="step-number">STEP 01</span>
                    <h4 class="step-title">选择台账记录并提取事实</h4>
                </div>
            """, unsafe_allow_html=True)
            
            search_kw = st.text_input("🔍 搜索记录 (序列号 / 供应商 / 失效模式):", placeholder="输入关键字快速过滤...")
            if search_kw:
                gen_df = gen_df[gen_df.astype(str).apply(lambda r: r.str.contains(search_kw, case=False).any(), axis=1)]
                
            selected_record_idx = st.selectbox(
                "👉 目标台账记录 (Target Record):",
                options=gen_df.index,
                format_func=lambda x: f"[{gen_df.loc[x, serial_no_col]}] {gen_df.loc[x, 'Failure Mode']} - {gen_df.loc[x, 'Project/Part name']}"
            )
            st.markdown('</div>', unsafe_allow_html=True)
            
            selected_row = gen_df.loc[selected_record_idx]
            ok_imgs, ng_imgs = get_multiple_images_for_row(excel_file, sheet_name, header_idx, selected_row.name)
            
            raw_facts_list = []
            for col_name in df.columns:
                if col_name not in ['选择 (Select)', 'Normalized_Status', '闭环状态']:
                    val_str = str(selected_row.get(col_name, '')).strip()
                    if val_str and val_str not in ['nan', 'None']:
                        raw_facts_list.append(f"{col_name}: {val_str}")
            raw_facts_block = "\n".join(raw_facts_list)
            
            prompt_content = f"""Please create me a short and precise lessons learned report out of the attached document in American English.
You are an honest engineer; you provide always links to the sources and name the original slide/page number.
Please stick to the facts. In case you have additional topics, supporting or additional useful information be creative, add them and highlight them in italic.

Please write the headings in bold. Use key words that are understood by others in Bosch. Describe the report "user-friendly", so others can read it easily. One page for chapter 1-4 is appropriate. Delete all hints in blue letters.

If you are asked to create a lesson learned report, or to search for a lessons learned report, structure the answer as follows:
0. Abstract - write a short summary of the report with the structure - issue; problem; learnings; tags

Abstract
Issue: Describe briefly. Do not use abbreviations that are not commonly known.
Problem: Briefly describe the main problem using key words.
Lessons: Concentrate on the actual lessons learned. What is new? How can we prevent this issue in the future?

Picture – Product – Defect
General Hint:
• Keep it short. Two pages for chapter 1-4 should be sufficient.
• Support your content with pictures where appropriate.
• Use key words that are understood by others in Bosch and not only in your area of expertise.

1. Product / Process
Product / Process:
Component:
Sub-Component:

2. Problem (Fundamental Problem)
Note:
Briefly describe the fundamental problem. Do not use technical root causes (TRC) or managerial root causes (MRC). Use pictures or graphs to visualize the problem.

3. Lessons
Document your lessons in the table.
| Lessons | Measures & Sustainable Solutions | Root Cause |
| :--- | :--- | :--- |
| What do you suggest could be done differently next time. | What measures did we implement in our area that might also be useful to others? | Describe the main root cause only. |

4. Potentially affected.
Determine who else might find this information useful to the best of your knowledge:
| Aspect | Description |
| :--- | :--- |
| What else could be additionally be affected? | (Similar applications, products, processes, …) or Not applicable |
| Where can the problem additionally occur? | (Other production lines, plants, regions, …) or Not applicable |
| When can the problem additionally appear? | (New applications, usage of products, …) or Not applicable |
| Who else can be affected? | (Other customers, suppliers, associates, …) or Not applicable |

Check if Centers of Competence (CoC) or BEO working groups should be informed: https://connect.bosch.com/communities/community/BEO

5. Appendix (Optional)

==================== [Raw Master List Facts] ====================
{raw_facts_block}"""

            # STEP 02
            st.markdown("""
            <div class="bds-step-card">
                <div class="step-header">
                    <span class="step-number">STEP 02</span>
                    <h4 class="step-title">复制 Prompt 并在 Teams M-PU Bot 中润色</h4>
                </div>
            """, unsafe_allow_html=True)
            
            c_p, c_b = st.columns([3.2, 1])
            with c_p:
                st.text_area("📋 完整工程 Prompt (纯白高对比高光文本域，点击右上角复制):", prompt_content, height=220)
            with c_b:
                st.markdown("<br>", unsafe_allow_html=True)
                st.link_button("🚀 一键直达 Teams M-PU Bot", TEAMS_BOT_URL, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # STEP 03
            st.markdown("""
            <div class="bds-step-card">
                <div class="step-header">
                    <span class="step-number">STEP 03</span>
                    <h4 class="step-title">粘贴 Bot 回复并生成交付件</h4>
                </div>
            """, unsafe_allow_html=True)
            
            col_in, col_sup = st.columns([3.2, 1.8])
            with col_in:
                bot_reply = st.text_area(
                    "📥 粘贴 M-PU Bot 润色后的完整回复：",
                    height=220,
                    placeholder="在此粘贴包含 0. Abstract, 1. Product/Process, 2. Problem, 3. Lessons, 4. Potentially affected 的完整文本..."
                )
            with col_sup:
                selected_sups = st.multiselect("👥 选择收件供应商 (自动解析邮箱):", options=list(supplier_dict.keys()))
                to_emails_list = []
                for s in selected_sups:
                    to_emails_list.extend(supplier_dict[s])
                to_emails_str = "; ".join(list(set(to_emails_list)))
                if to_emails_str:
                    st.info(f"📧 **自动收件人:**\n`{to_emails_str}`")

            if st.button("🚀 立即生成标准化 Word 报告与双附件 Outlook 邮件草稿", type="primary", use_container_width=True):
                if template_file is None:
                    st.error("❌ 未检测到 Word 模板，请在侧边栏确认路径。")
                else:
                    with st.spinner("正在装配表格、建立原生居中蓝底方块并生成邮件草稿..."):
                        bot_data = parse_bot_feber_response(bot_reply) if bot_reply.strip() else {}
                        doc = populate_docx_exact_tables(template_file, bot_data, selected_row, ok_imgs, ng_imgs)
                        bio = io.BytesIO()
                        doc.save(bio)
                        doc_bytes = bio.getvalue()
                        
                        raw_serial = str(selected_row.get(serial_no_col, '')).strip()
                        if not raw_serial or raw_serial in ['nan', 'None']:
                            raw_serial = "LL-Export"
                            
                        doc_filename = f"LL_Template_{raw_serial}.docx"
                        
                        eml_bytes, final_subject = build_new_email_dual_attachment(
                            selected_row, 
                            to_emails_str, 
                            doc_bytes, 
                            doc_filename, 
                            feedback_bytes, 
                            feedback_filename
                        )
                        
                        st.success(f"🎉 生成成功！邮件主题已更新为: `{final_subject}`")
                        c_d1, c_d2 = st.columns(2)
                        with c_d1:
                            st.download_button(
                                f"📥 下载 Word 报告 (多图已带蓝底居中方块并排): {doc_filename}",
                                doc_bytes,
                                doc_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        with c_d2:
                            st.download_button(
                                f"📧 下载 Outlook 草稿 (含双附件): Email_Draft_{raw_serial}.eml",
                                eml_bytes,
                                f"Email_Draft_{raw_serial}.eml",
                                mime="message/rfc822",
                                use_container_width=True
                            )
            st.markdown('</div>', unsafe_allow_html=True)

        # =========================================================================
        # 模式二：高阶质量全景与闭环看板 (Split Matrix View)
        # =========================================================================
        else:
            total_cases = len(df)
            completed_cases = len(df[df['Normalized_Status'] == 'Completed'])
            pending_cases = total_cases - completed_cases
            closure_rate = (completed_cases / total_cases * 100) if total_cases > 0 else 0

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #005691;">
                    <div class="kpi-title">📚 总有效经验库 (Need='Y')</div>
                    <div class="kpi-value">{total_cases} <span style="font-size:0.9rem; font-weight:normal; color:#525F6B;">项</span></div>
                </div>
                """, unsafe_allow_html=True)
            with k2:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #78BE20;">
                    <div class="kpi-title">🟢 已闭环结束 (Complete='Y')</div>
                    <div class="kpi-value" style="color: #2E7D32;">{completed_cases} <span style="font-size:0.9rem; font-weight:normal; color:#525F6B;">项</span></div>
                </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #E20015;">
                    <div class="kpi-title">🔴 待闭环处理 (Complete='N')</div>
                    <div class="kpi-value" style="color: #E20015;">{pending_cases} <span style="font-size:0.9rem; font-weight:normal; color:#525F6B;">项</span></div>
                </div>
                """, unsafe_allow_html=True)
            with k4:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #008ECF;">
                    <div class="kpi-title">🎯 闭环完成率 (Closure Rate)</div>
                    <div class="kpi-value" style="color: #005691;">{closure_rate:.1f}%</div>
                </div>
                """, unsafe_allow_html=True)

            st.write("")

            f_col1, f_col2, f_col3, f_col4 = st.columns([1.5, 2, 2, 2.5])
            with f_col1:
                status_filter = st.selectbox("📌 闭环状态:", options=["全部 (All)", "已完成 (Closed)", "待处理 (Open)"])
            with f_col2:
                sup_col_name = next((c for c in df.columns if 'supplier' in str(c).lower()), None)
                if sup_col_name:
                    sups = ["全部 (All)"] + sorted([str(x) for x in df[sup_col_name].dropna().unique() if str(x).strip() != ''])
                else: sups = ["全部 (All)"]
                chosen_sup = st.selectbox("👥 供应商:", options=sups)
            with f_col3:
                proj_col_name = next((c for c in df.columns if 'project' in str(c).lower() or 'part' in str(c).lower()), None)
                if proj_col_name:
                    projs = ["全部 (All)"] + sorted([str(x) for x in df[proj_col_name].dropna().unique() if str(x).strip() != ''])
                else: projs = ["全部 (All)"]
                chosen_proj = st.selectbox("🚗 零件/项目:", options=projs)
            with f_col4:
                search_q = st.text_input("🔍 关键字检索:", placeholder="输入编号/原因/失效模式...")

            view_df = df.copy()
            if status_filter == "已完成 (Closed)":
                view_df = view_df[view_df['Normalized_Status'] == 'Completed']
            elif status_filter == "待处理 (Open)":
                view_df = view_df[view_df['Normalized_Status'] == 'Pending']
                
            if chosen_sup != "全部 (All)" and sup_col_name:
                view_df = view_df[view_df[sup_col_name].astype(str) == chosen_sup]
            if chosen_proj != "全部 (All)" and proj_col_name:
                view_df = view_df[view_df[proj_col_name].astype(str) == chosen_proj]
            if search_q:
                view_df = view_df[view_df.astype(str).apply(lambda r: r.str.contains(search_q, case=False).any(), axis=1)]

            st.write("")

            col_list_view, col_detail_view = st.columns([1.5, 1.5])
            
            view_df['闭环状态'] = view_df['Normalized_Status'].apply(lambda x: '🟢 已完成' if x == 'Completed' else '🔴 进行中')
            
            table_disp_cols = [
                serial_no_col, 
                '闭环状态',
                'Project/Part name' if 'Project/Part name' in view_df.columns else proj_col_name,
                'Failure Mode' if 'Failure Mode' in view_df.columns else 'LL Brief Description'
            ]
            valid_table_cols = [c for c in table_disp_cols if c and c in view_df.columns]
            
            selected_row_data = None
            
            with col_list_view:
                st.markdown(f"##### 📋 经验库清单 (共 {len(view_df)} 条)")
                
                event = st.dataframe(
                    view_df[valid_table_cols],
                    use_container_width=True,
                    height=450,
                    hide_index=True,
                    on_select="rerun",
                    selection_mode="single-row"
                )
                
                if event and "rows" in event.selection and len(event.selection["rows"]) > 0:
                    clicked_idx = event.selection["rows"][0]
                    selected_row_data = view_df.iloc[clicked_idx]
                elif len(view_df) > 0:
                    selected_row_data = view_df.iloc[0]

            with col_detail_view:
                st.markdown("##### 🔬 经验卡片详情")
                if selected_row_data is not None:
                    focus_status = selected_row_data.get('Normalized_Status', 'Pending')
                    _, case_imgs = get_multiple_images_for_row(excel_file, sheet_name, header_idx, selected_row_data.name)
                    
                    status_badge = '<span class="badge-completed">🟢 已完成</span>' if focus_status == 'Completed' else '<span class="badge-pending">🔴 进行中</span>'
                    
                    st.markdown(f"""
                    <div class="inspector-panel">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #E2E8F0; padding-bottom:8px;">
                            <span style="font-size:1.2rem; font-weight:700; color:#005691;">📌 {selected_row_data.get(serial_no_col, '')}</span>
                            {status_badge}
                        </div>
                        <div style="margin-bottom:8px;"><strong>项目/零件 (Project):</strong> <span style="color:#1C2B39;">{selected_row_data.get('Project/Part name', 'N/A')}</span></div>
                        <div style="margin-bottom:8px;"><strong>失效简述 (Description):</strong><br><span style="color:#525F6B;">{selected_row_data.get('LL Brief Description', selected_row_data.get('Failure Mode', 'N/A'))}</span></div>
                        <div style="margin-bottom:8px;"><strong>根本原因 (Root Cause):</strong><br><span style="color:#525F6B;">{selected_row_data.get('Root Cause', '未录入')}</span></div>
                        <div style="margin-bottom:12px;"><strong>学习核心点 (LL point):</strong><br><span style="color:#005691; font-weight:600;">{selected_row_data.get('LL point', selected_row_data.get('Corrective Action', '未录入'))}</span></div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.write("")
                    
                    if case_imgs:
                        st.markdown(f"🖼 **不良图片库 (Defect Pictures, 共 {len(case_imgs)} 张):**")
                        img_cols = st.columns(min(len(case_imgs), 2))
                        for i, img_b in enumerate(case_imgs):
                            with img_cols[i % 2]:
                                st.image(img_b, width=220, caption=f"实物图 #{i+1}")
                                with st.expander(f"🔍 放大原图 #{i+1}", expanded=False):
                                    st.image(img_b, use_container_width=True)
                    else:
                        st.markdown("""
                        <div style="height:100px; background:#FFFFFF; border:1.5px dashed #CBD5E1; border-radius:6px; display:flex; align-items:center; justify-content:center; color:#94A3B8; font-size:0.85rem;">
                            暂无实物图片
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.info("👈 暂无数据。")
                    
    except Exception as e:
        st.error(f"❌ 读取或渲染异常: {e}")
else:
    st.info("ℹ️ 请在侧边栏确认 Master List 文件。")
