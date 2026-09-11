# -*- coding: utf-8 -*-
"""
=============================================================================
BOSCH | PCB Quality Studio & Executive Intelligence Dashboard
- Mode 1: Automated FEBER Word Report & Dual-Attachment Outlook Draft
- Mode 2: Executive Quality Intelligence Dashboard (Beyond Power BI)
- Interactive Defect Gallery, Picture Extraction & Completion Tracking
=============================================================================
"""

import streamlit as st
import pandas as pd
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
import datetime
import io
import os
import glob
import zipfile
import re
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
    page_title="Bosch | PCB Quality Studio & Dashboard",
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
        --bosch-bg: #F4F6F8;
        --bosch-card: #FFFFFF;
        --bosch-border: #DDE3EA;
    }
    .stApp { background-color: var(--bosch-bg); font-family: 'Arial', sans-serif; }
    
    .bosch-top-bar {
        height: 6px;
        background: linear-gradient(90deg, #E20015 0%, #E20015 25%, #005691 25%, #005691 65%, #007BC0 65%, #007BC0 100%);
        border-radius: 3px;
        margin-bottom: 16px;
    }
    
    .bds-card {
        background: var(--bosch-card);
        border: 1px solid var(--bosch-border);
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 16px;
        box-shadow: 0 4px 12px rgba(0, 40, 80, 0.04);
    }
    
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid var(--bosch-border);
        border-radius: 8px;
        padding: 16px 20px;
        border-top: 4px solid var(--bosch-blue);
        box-shadow: 0 4px 10px rgba(0, 40, 80, 0.03);
    }
    .kpi-title { font-size: 0.85rem; color: #525F6B; font-weight: 600; text-transform: uppercase; }
    .kpi-value { font-size: 1.9rem; color: #005691; font-weight: 700; margin-top: 4px; }
    
    .case-card {
        background: #FFFFFF;
        border: 1px solid var(--bosch-border);
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
        transition: all 0.2s ease-in-out;
    }
    .case-card:hover {
        border-color: var(--bosch-light-blue);
        box-shadow: 0 6px 18px rgba(0, 86, 145, 0.08);
    }
    
    .badge-completed {
        background-color: #E8F5E9;
        color: #2E7D32;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.75rem;
        display: inline-block;
    }
    .badge-pending {
        background-color: #FFEBEE;
        color: #C62828;
        padding: 4px 12px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.75rem;
        display: inline-block;
    }
    
    .bds-step-badge {
        display: inline-block;
        background: var(--bosch-blue);
        color: #FFFFFF;
        font-size: 0.8rem;
        font-weight: 700;
        padding: 4px 12px;
        border-radius: 20px;
        margin-bottom: 10px;
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
    """扫描所有以 PCB Lesson Learn Master List 开头的 Excel 文件"""
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
    """自适应查找具体单个文件"""
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
        "📊 Master List 版本 (已自动定位最新):",
        options=available_master_lists,
        format_func=lambda x: f"{os.path.basename(x)} ({datetime.datetime.fromtimestamp(os.path.getmtime(x)).strftime('%Y-%m-%d %H:%M')})"
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
# 3. 核心工具函数集合
# -----------------------------------------------------------------------------

def load_supplier_emails(file_source):
    """从 Vendor code 表中读取供应商与邮箱映射"""
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

def get_images_for_row(file_source, sheet_name, header_idx, target_row_idx):
    """从 Excel 指定行提取 NG 和 OK 图片二进制流"""
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
        ok_img = None
        ng_img = None
        
        for img in getattr(ws, '_images', []):
            try:
                r = img.anchor._from.row
                c = img.anchor._from.col
                if r == excel_target_row:
                    if c == col_ng or (col_ng != -1 and abs(c - col_ng) <= 1):
                        ng_img = img._data()
                    elif c == col_ok:
                        ok_img = img._data()
            except Exception:
                pass
                
        return ok_img, ng_img
    except Exception as e:
        return None, None

def load_excel_robust(file_source):
    """加载 Excel 并定位表头"""
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
    """解析 Bot 按照 FEBER 规范输出的结构化文本与 3 列表格"""
    parsed = {
        'Abstract_Issue': '',
        'Abstract_Problem': '',
        'Abstract_Lessons': '',
        'Product_Process': '',
        'Component': '',
        'Sub_Component': '',
        'Problem': '',
        'Lessons_Rows': [],
        'What_Else': '',
        'Where': '',
        'When': '',
        'Who': ''
    }
    
    # 0. Abstract
    m_abs = re.search(r'(?:0\.\s*Abstract|Abstract)\s*([\s\S]*?)(?=1\.\s*Product|$)', bot_text, re.I)
    if m_abs:
        t = m_abs.group(1)
        i_m = re.search(r'Issue:\s*([\s\S]*?)(?=Problem:|$)', t, re.I)
        p_m = re.search(r'Problem:\s*([\s\S]*?)(?=Lessons:|$)', t, re.I)
        l_m = re.search(r'Lessons:\s*([\s\S]*?)(?=Tags:|Picture|1\.\s*Product|$)', t, re.I)
        if i_m: parsed['Abstract_Issue'] = i_m.group(1).strip()
        if p_m: parsed['Abstract_Problem'] = p_m.group(1).strip()
        if l_m: parsed['Abstract_Lessons'] = l_m.group(1).strip()
    
    # 1. Product / Process
    m_p = re.search(r'1\.\s*Product\s*/\s*Process\s*([\s\S]*?)(?=2\.\s*Problem|$)', bot_text, re.I)
    if m_p:
        t = m_p.group(1)
        pp = re.search(r'Product\s*/\s*Process:\s*([^\n]+)', t, re.I)
        cp = re.search(r'Component:\s*([^\n]+)', t, re.I)
        sc = re.search(r'Sub-Component:\s*([^\n]+)', t, re.I)
        if pp: parsed['Product_Process'] = pp.group(1).strip()
        if cp: parsed['Component'] = cp.group(1).strip()
        if sc: parsed['Sub_Component'] = sc.group(1).strip()
        
    # 2. Problem
    m_prob = re.search(r'2\.\s*Problem[^\n]*\n([\s\S]*?)(?=3\.\s*Lessons|$)', bot_text, re.I)
    if m_prob:
        parsed['Problem'] = m_prob.group(1).strip()
    
    # 3. Lessons (支持制表符 \t 与 Markdown | 双模式)
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
            
    # 4. Potentially affected
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
# 5. 精准装配 Word 模板核心函数 (彻底攻克图片置入与格式)
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

def insert_content_under_heading(doc, heading_kw, text_value):
    if not text_value: return
    for idx, p in enumerate(doc.paragraphs):
        p_txt = p.text.strip().lower()
        if heading_kw.lower() in p_txt and len(p_txt) < 50:
            new_p_elem = OxmlElement('w:p')
            p._element.addnext(new_p_elem)
            new_p = docx.text.paragraph.Paragraph(new_p_elem, doc)
            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = new_p.add_run(str(text_value).strip())
            r.font.name = 'Arial'
            r.font.size = Pt(10.5)
            r.font.bold = False
            return

def populate_docx_exact_tables(template_source, bot_data, raw_row, ok_img=None, ng_img=None):
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

    prob_val = bot_data.get('Problem') or raw_row.get('LL Brief Description', '')
    insert_content_under_heading(doc, "Problem (Fundamental Problem)", prob_val)

    picture_inserted = False
    for table in doc.tables:
        t_header = "".join(cell.text for cell in table.rows[0].cells).lower()
        
        for row in table.rows:
            for cell in row.cells:
                c_txt = cell.text.lower().replace(" ", "")
                if ("picture" in c_txt or "defect" in c_txt or "not-ok" in c_txt) and "ok-part" not in c_txt:
                    if ng_img:
                        cell.text = ""
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.line_spacing = 1.0
                        p.add_run().add_picture(io.BytesIO(ng_img), width=Inches(1.85))
                        picture_inserted = True
                elif "ok-part" in c_txt:
                    if ok_img:
                        cell.text = ""
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.line_spacing = 1.0
                        p.add_run().add_picture(io.BytesIO(ok_img), width=Inches(1.85))

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
                for c_idx in range(min(3, len(row_tuple))):
                    set_cell_formatted_text(new_row.cells[c_idx], row_tuple[c_idx])

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

    if not picture_inserted and ng_img:
        for p in doc.paragraphs:
            p_txt_clean = p.text.lower().replace(" ", "")
            if "picture" in p_txt_clean and len(p_txt_clean) < 40:
                p.text = ""
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(io.BytesIO(ng_img), width=Inches(1.85))
                break

    return doc

def generate_eml_file_dual_attachment(row_data, to_emails="", doc_bytes=None, doc_filename="LL_Template.docx", feedback_bytes=None, feedback_filename="LL Feedback table_Supplier version_V1.xlsx"):
    """【双附件 Outlook 邮件生成引擎 (EML 格式)】"""
    serial_no = str(row_data.get('LL Serials No', 'LL-xxxx-xx')).strip()
    failure_mode = str(row_data.get('Failure Mode', '*****')).strip()
    subject = f"M/PQR-AP LL | {serial_no} | Title {failure_mode}"
    
    html_body = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Arial', sans-serif; font-size: 10.5pt; line-height: 1.6; color: #333333; }}
            .red-bold {{ color: #E20015; font-weight: bold; }}
            ul {{ margin-top: 5px; margin-bottom: 15px; padding-left: 20px; }}
            li {{ margin-bottom: 8px; }}
        </style>
    </head>
    <body>
        <p>Dear Supplier:</p>
        <p>Recently, we summarized a lesson learn about <strong>{failure_mode}</strong>. Please review the attached LL document and complete following tasks:</p>
        <ul>
            <li>Complete feedback form based on self-evaluation on your own processes and send to your responsible PQR and PUQ-PQA (ME) within <span class="red-bold">one week.</span></li>
            <li>After your self-evaluation, please close defined actions within <span class="red-bold">three weeks.</span></li>
            <li>Our PQR or PUQ-PQA colleague may conduct onsite verification according to the information in feedback form in <span class="red-bold">one month.</span></li>
        </ul>
        <p>If you have any question about this lesson learn, please contact with your responsible PQR and PUQ-PQA (ME).</p>
        <p>Best regards,</p>
        <p><strong>Purchasing Quality Region Asia Pacific Team</strong></p>
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
        
    return msg.as_bytes()

# -----------------------------------------------------------------------------
# 6. 系统导航：双模式交互架构
# -----------------------------------------------------------------------------
app_mode = st.radio(
    "👉 请选择工作模式 (Work Mode):",
    options=["📑 模式一：FEBER 报告生成与邮件协同", "📊 模式二：高阶质量全景与闭环看板 (Executive Dashboard)"],
    horizontal=True
)

st.write("---")

if excel_file is not None:
    try:
        df, sheet_name, header_idx = load_excel_robust(excel_file)
        supplier_dict = load_supplier_emails(excel_file)
        
        serial_no_col = next((c for c in df.columns if 'serial' in str(c).lower()), 'LL Serials No')
        supplier_scope_col = next((c for c in df.columns if 'scope' in str(c).lower() or 'task' in str(c).lower()), 'LL Supplier Scope')
        
        complete_col = None
        for col in df.columns:
            c_low = str(col).lower()
            if 'complete' in c_low or 'closure' in c_low or 'status' in c_low or 'tracking' in c_low:
                complete_col = col
                break
        if not complete_col:
            complete_col = 'Complete Status'
            df[complete_col] = 'Open'
            
        def get_clean_status(val):
            s = str(val).strip().lower()
            if s in ['y', 'yes', 'completed', 'complete', 'closed', 'done', '100%', '100', '1', 'true']:
                return 'Completed'
            return 'Pending'
            
        df['Normalized_Status'] = df[complete_col].apply(get_clean_status)

        # =========================================================================
        # 模式一：FEBER 模板生成与邮件协同
        # =========================================================================
        if app_mode == "📑 模式一：FEBER 报告生成与邮件协同":
            feedback_bytes = None
            feedback_filename = "LL Feedback table_Supplier version_V1.xlsx"
            if feedback_file is not None and os.path.exists(feedback_file):
                with open(feedback_file, 'rb') as f:
                    feedback_bytes = f.read()
                    
            ll_need_col = next((c for c in df.columns if 'need or not' in str(c).lower()), None)
            gen_df = df.copy()
            if ll_need_col:
                gen_df = gen_df[gen_df[ll_need_col].astype(str).str.strip().str.upper() == 'Y']
                
            st.markdown('<div class="bds-card">', unsafe_allow_html=True)
            st.markdown('<span class="bds-step-badge">STEP 1</span> <h4 style="display:inline; margin-left:8px; color:#005691;">选择台账记录并提取事实</h4>', unsafe_allow_html=True)
            
            search_kw = st.text_input("🔍 搜索记录 (序列号/供应商/失效模式):", placeholder="输入关键字实时过滤...")
            if search_kw:
                gen_df = gen_df[gen_df.astype(str).apply(lambda r: r.str.contains(search_kw, case=False).any(), axis=1)]
                
            selected_record_idx = st.selectbox(
                "👉 请选择台账记录：",
                options=gen_df.index,
                format_func=lambda x: f"[{gen_df.loc[x, serial_no_col]}] {gen_df.loc[x, 'Failure Mode']} - {gen_df.loc[x, 'Project/Part name']}"
            )
            st.markdown('</div>', unsafe_allow_html=True)
            
            selected_row = gen_df.loc[selected_record_idx]
            ok_img, ng_img = get_images_for_row(excel_file, sheet_name, header_idx, selected_row.name)
            
            raw_facts_list = []
            for col_name in df.columns:
                if col_name not in ['选择 (Select)', 'Normalized_Status']:
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

            st.markdown('<div class="bds-card">', unsafe_allow_html=True)
            st.markdown('<span class="bds-step-badge">STEP 2</span> <h4 style="display:inline; margin-left:8px; color:#005691;">一键复制 Prompt 并在 Teams M-PU Bot 润色</h4>', unsafe_allow_html=True)
            c_p, c_b = st.columns([3, 1])
            with c_p:
                st.text_area("📋 已完整内嵌 3 列表格原型的工程 Prompt (点击右上角图标复制):", prompt_content, height=220)
            with c_b:
                st.markdown("<br>", unsafe_allow_html=True)
                st.link_button("🚀 一键直达 Teams M-PU Bot", TEAMS_BOT_URL, use_container_width=True)
                st.caption("💡 操作提示：复制左侧带有 3 列表格的完整 Prompt，在 Teams 窗口中发送给 Bot。")
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="bds-card">', unsafe_allow_html=True)
            st.markdown('<span class="bds-step-badge">STEP 3</span> <h4 style="display:inline; margin-left:8px; color:#005691;">粘贴 Bot 回复并一键生成最终交付包</h4>', unsafe_allow_html=True)
            col_in, col_sup = st.columns([3, 2])
            with col_in:
                bot_reply = st.text_area(
                    "📥 在此粘贴 M-PU Bot 润色后的完整回复：",
                    height=220,
                    placeholder="粘贴 Bot 输出的包含 0. Abstract, 1. Product/Process, 2. Problem, 3. Lessons (包含3列表格), 4. Potentially affected 的完整文本..."
                )
            with col_sup:
                selected_sups = st.multiselect("👥 选择收件供应商 (自动读取 Vendor code 邮箱):", options=list(supplier_dict.keys()))
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
                    with st.spinner("正在定向装配表格、紧凑插入不良图片并生成双附件邮件草稿..."):
                        bot_data = parse_bot_feber_response(bot_reply) if bot_reply.strip() else {}
                        doc = populate_docx_exact_tables(template_file, bot_data, selected_row, ok_img, ng_img)
                        bio = io.BytesIO()
                        doc.save(bio)
                        doc_bytes = bio.getvalue()
                        
                        serial_str = str(selected_row.get(serial_no_col, 'LL-Export'))
                        doc_filename = f"LL_Template_{serial_str}.docx"
                        eml_bytes = generate_eml_file_dual_attachment(
                            selected_row, 
                            to_emails_str, 
                            doc_bytes, 
                            doc_filename, 
                            feedback_bytes, 
                            feedback_filename
                        )
                        
                        st.success("🎉 生成成功！邮件已包含【Word报告 + Excel反馈表】双附件，图片已紧凑居中嵌入。")
                        c_d1, c_d2 = st.columns(2)
                        with c_d1:
                            st.download_button(
                                f"📥 下载 Word 报告: {doc_filename}",
                                doc_bytes,
                                doc_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        with c_d2:
                            st.download_button(
                                f"📧 下载 Outlook 草稿 (含双附件): Email_Draft_{serial_str}.eml",
                                eml_bytes,
                                f"Email_Draft_{serial_str}.eml",
                                mime="message/rfc822",
                                use_container_width=True
                            )
            st.markdown('</div>', unsafe_allow_html=True)

        # =========================================================================
        # 模式二：高阶质量全景与闭环看板 (已修复混合类型排序报错)
        # =========================================================================
        else:
            st.markdown("""
            <div style="margin-bottom: 20px;">
                <h3 style="color:#005691; margin:0;">📊 PCB Quality Intelligence & Lessons Learned Dashboard</h3>
                <p style="color:#525F6B; font-size:0.9rem; margin-top:4px;">质量经验全景穿透 · 闭环追踪健康度分析 · 案例图文数字档案</p>
            </div>
            """, unsafe_allow_html=True)

            # 1. 顶部交互过滤器（彻底消除 int/str 比较报错）
            st.markdown('<div class="bds-card" style="padding:15px;">', unsafe_allow_html=True)
            f_col1, f_col2, f_col3, f_col4 = st.columns([1.5, 2, 2, 2.5])
            
            with f_col1:
                status_filter = st.selectbox("📌 闭环状态过滤:", options=["全部 (All)", "已完成 (Completed)", "未完成 (Pending)"])
            with f_col2:
                sup_col_name = next((c for c in df.columns if 'supplier' in str(c).lower()), None)
                if sup_col_name:
                    unique_sups = [str(x) for x in df[sup_col_name].dropna().unique() if str(x).strip() != '']
                    all_sups = ["全部 (All)"] + sorted(unique_sups)
                else:
                    all_sups = ["全部 (All)"]
                chosen_sup_filter = st.selectbox("👥 供应商筛选:", options=all_sups)
            with f_col3:
                proj_col_name = next((c for c in df.columns if 'project' in str(c).lower() or 'part' in str(c).lower()), None)
                if proj_col_name:
                    unique_projs = [str(x) for x in df[proj_col_name].dropna().unique() if str(x).strip() != '']
                    all_projs = ["全部 (All)"] + sorted(unique_projs)
                else:
                    all_projs = ["全部 (All)"]
                chosen_proj_filter = st.selectbox("🚗 零件/项目筛选:", options=all_projs)
            with f_col4:
                search_text = st.text_input("🔍 全文检索 (序列号/失效模式/根本原因):", placeholder="输入任意关键字...")
                
            st.markdown('</div>', unsafe_allow_html=True)

            # 执行过滤
            filtered_dash_df = df.copy()
            if status_filter == "已完成 (Completed)":
                filtered_dash_df = filtered_dash_df[filtered_dash_df['Normalized_Status'] == 'Completed']
            elif status_filter == "未完成 (Pending)":
                filtered_dash_df = filtered_dash_df[filtered_dash_df['Normalized_Status'] == 'Pending']
                
            if chosen_sup_filter != "全部 (All)" and sup_col_name:
                filtered_dash_df = filtered_dash_df[filtered_dash_df[sup_col_name].astype(str) == chosen_sup_filter]
                
            if chosen_proj_filter != "全部 (All)" and proj_col_name:
                filtered_dash_df = filtered_dash_df[filtered_dash_df[proj_col_name].astype(str) == chosen_proj_filter]
                
            if search_text:
                filtered_dash_df = filtered_dash_df[filtered_dash_df.astype(str).apply(lambda r: r.str.contains(search_text, case=False).any(), axis=1)]

            # 2. 核心 KPI 动态指标栏
            total_cases = len(df)
            completed_cases = len(df[df['Normalized_Status'] == 'Completed'])
            pending_cases = total_cases - completed_cases
            closure_rate = (completed_cases / total_cases * 100) if total_cases > 0 else 0

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #005691;">
                    <div class="kpi-title">📚 总 Lessons Learned 案例</div>
                    <div class="kpi-value">{total_cases} <span style="font-size:1rem; font-weight:normal; color:#525F6B;">条</span></div>
                </div>
                """, unsafe_allow_html=True)
            with k2:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #78BE20;">
                    <div class="kpi-title">✅ 已完成闭环 (Completed)</div>
                    <div class="kpi-value" style="color: #2E7D32;">{completed_cases} <span style="font-size:1rem; font-weight:normal; color:#525F6B;">条</span></div>
                </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #E20015;">
                    <div class="kpi-title">⏳ 待处理/进行中 (Pending)</div>
                    <div class="kpi-value" style="color: #E20015;">{pending_cases} <span style="font-size:1rem; font-weight:normal; color:#525F6B;">条</span></div>
                </div>
                """, unsafe_allow_html=True)
            with k4:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color: #008ECF;">
                    <div class="kpi-title">🎯 闭环健康度 (Closure Rate)</div>
                    <div class="kpi-value" style="color: #005691;">{closure_rate:.1f}%</div>
                </div>
                """, unsafe_allow_html=True)

            st.write("")

            # 3. 交互式可视化图表区
            if HAS_PLOTLY and len(filtered_dash_df) > 0:
                c_chart1, c_chart2 = st.columns([1, 2])
                with c_chart1:
                    st.markdown('<div class="bds-card">', unsafe_allow_html=True)
                    st.markdown("<h5 style='color:#005691; margin-bottom:10px;'>📊 闭环达成率分布</h5>", unsafe_allow_html=True)
                    status_counts = filtered_dash_df['Normalized_Status'].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Count']
                    fig_donut = px.pie(
                        status_counts, 
                        names='Status', 
                        values='Count',
                        hole=0.6,
                        color='Status',
                        color_discrete_map={'Completed': '#78BE20', 'Pending': '#E20015'}
                    )
                    fig_donut.update_traces(textposition='inside', textinfo='percent+label')
                    fig_donut.update_layout(
                        showlegend=False, 
                        margin=dict(t=10, b=10, l=10, r=10),
                        height=260,
                        annotations=[dict(text=f"{closure_rate:.1f}%", x=0.5, y=0.5, font_size=24, font_color="#005691", showarrow=False)]
                    )
                    st.plotly_chart(fig_donut, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                with c_chart2:
                    st.markdown('<div class="bds-card">', unsafe_allow_html=True)
                    st.markdown("<h5 style='color:#005691; margin-bottom:10px;'>📈 供应商缺陷与闭环分布</h5>", unsafe_allow_html=True)
                    if sup_col_name:
                        chart_df = filtered_dash_df.copy()
                        chart_df[sup_col_name] = chart_df[sup_col_name].astype(str)
                        sup_summary = chart_df.groupby([sup_col_name, 'Normalized_Status']).size().reset_index(name='Count')
                        fig_bar = px.bar(
                            sup_summary, 
                            x=sup_col_name, 
                            y='Count', 
                            color='Normalized_Status',
                            barmode='stack',
                            color_discrete_map={'Completed': '#005691', 'Pending': '#E20015'},
                            labels={sup_col_name: '供应商 (Supplier)', 'Count': '案例数', 'Normalized_Status': '闭环状态'}
                        )
                        fig_bar.update_layout(margin=dict(t=10, b=20, l=20, r=10), height=260, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                        st.plotly_chart(fig_bar, use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)

            # 4. 图文并茂的案例详情画廊
            st.markdown(f"#### 🔎 案例图文全景画廊 (共筛选出 {len(filtered_dash_df)} 条记录)")
            
            if len(filtered_dash_df) == 0:
                st.warning("⚠️ 当前过滤条件下未检索到相关案例。")
            else:
                for idx, row in filtered_dash_df.iterrows():
                    serial_val = row.get(serial_no_col, 'N/A')
                    desc_val = row.get('LL Brief Description', row.get('Failure Mode', 'N/A'))
                    proj_val = row.get('Project/Part name', 'N/A')
                    rc_val = row.get('Root Cause', '未录入根本原因')
                    ll_point_val = row.get('LL point', row.get('Corrective Action', '未录入建议'))
                    status_val = row.get('Normalized_Status', 'Pending')
                    
                    _, card_ng_img = get_images_for_row(excel_file, sheet_name, header_idx, row.name)
                    
                    st.markdown(f"""
                    <div class="case-card">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                            <div>
                                <strong style="font-size:1.15rem; color:#005691;">📌 {serial_val}</strong>
                                <span style="margin-left:12px; color:#525F6B; font-weight:600;">项目/零件: {proj_val}</span>
                            </div>
                            <div>
                                <span class="{ 'badge-completed' if status_val == 'Completed' else 'badge-pending' }">
                                    {'🟢 已闭环 Completed' if status_val == 'Completed' else '🔴 待处理 Pending'}
                                </span>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    card_left, card_right = st.columns([3, 1])
                    with card_left:
                        st.markdown(f"**📝 失效描述 (Description):**\n{desc_val}")
                        st.markdown(f"**🔬 根本原因 (Root Cause):**\n{rc_val}")
                        st.markdown(f"**💡 经验学习核心 (LL Point / Action):**\n{ll_point_val}")
                        
                    with card_right:
                        if card_ng_img:
                            st.image(card_ng_img, caption="不良图片 (Defect Picture)", use_column_width=True)
                        else:
                            st.markdown("""
                            <div style="height:120px; background:#F8FAFC; border:1px dashed #D0DCE5; border-radius:6px; display:flex; align-items:center; justify-content:center; color:#94A3B8; font-size:0.85rem;">
                                暂无图片 (No Picture)
                            </div>
                            """, unsafe_allow_html=True)
                            
                    st.markdown("</div>", unsafe_allow_html=True)
                    
    except Exception as e:
        st.error(f"❌ 读取或渲染异常: {e}")
else:
    st.info("ℹ️ 请在侧边栏确认 Master List (Excel) 文件的有效路径。")
