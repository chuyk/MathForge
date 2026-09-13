import os
import re
import lxml.etree as ET
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from latex2mathml.converter import convert as latex_to_mathml

# 尋找 MML2OMML.XSL 樣式轉換檔路徑
_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_LOCAL_XSL = os.path.join(_CURRENT_DIR, "MML2OMML.XSL")

_OFFICE_XSL_CANDIDATES = [
    _LOCAL_XSL,
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office15\MML2OMML.XSL",
]

_XSLT_TRANSFORM = None

def get_xslt_transform():
    """快取並返回 MML2OMML 的 XSLT 轉換物件。"""
    global _XSLT_TRANSFORM
    if _XSLT_TRANSFORM is not None:
        return _XSLT_TRANSFORM
    
    xsl_path = None
    for path in _OFFICE_XSL_CANDIDATES:
        if os.path.exists(path):
            xsl_path = path
            break
            
    if not xsl_path:
        raise FileNotFoundError(
            "找不到 MML2OMML.XSL 轉換檔！請確保微軟 Office 已安裝或 _tools/MML2OMML.XSL 存在。"
        )
        
    xslt_doc = ET.parse(xsl_path)
    _XSLT_TRANSFORM = ET.XSLT(xslt_doc)
    return _XSLT_TRANSFORM

def _rgb_to_hex(rgb):
    """將 (r, g, b) 或 hex 字串轉為 6 位元大寫 HEX 字串。"""
    if isinstance(rgb, str):
        return rgb.lstrip("#").upper()
    if isinstance(rgb, (tuple, list)) and len(rgb) >= 3:
        return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    return None

def latex_to_omml(latex_code: str, color=None):
    r"""
    將 LaTeX 數學式字串轉為 Word 原生 OMML (<m:oMath>) XML 元素。
    
    :param latex_code: LaTeX 語法字串，例如 r'8\frac{3}{4}' 或 r'\overline{AB}'
    :param color: 可選顏色，支援 (r, g, b) 數組或 HEX 字串（例如 'B40000'）
    :return: docx.oxml.OxmlElement
    """
    transform = get_xslt_transform()
    
    # 預處理台灣國中常用幾何與數學符號替換
    latex_clean = latex_code.strip()
    
    # 轉換 LaTeX -> MathML
    mathml_str = latex_to_mathml(latex_clean)
    
    # MathML -> OMML
    dom = ET.fromstring(mathml_str)
    omml_dom = transform(dom)
    
    ns = {
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    
    # 將幾何線段的 overline (<m:acc> 帶有橫槓) 轉為微軟標準頂端線條 <m:bar><m:barPr><m:pos m:val="top"/>
    for acc in omml_dom.xpath('.//m:acc', namespaces=ns):
        chr_el = acc.find('m:accPr/m:chr', namespaces=ns)
        if chr_el is not None and chr_el.get(f'{{{ns["m"]}}}val') in ('―', '\u2015', '\u00af', '¯'):
            e_el = acc.find('m:e', namespaces=ns)
            if e_el is not None:
                bar_el = ET.Element(f'{{{ns["m"]}}}bar')
                bar_pr = ET.SubElement(bar_el, f'{{{ns["m"]}}}barPr')
                pos_el = ET.SubElement(bar_pr, f'{{{ns["m"]}}}pos')
                pos_el.set(f'{{{ns["m"]}}}val', 'top')
                bar_el.append(e_el)
                parent = acc.getparent()
                if parent is not None:
                    parent.replace(acc, bar_el)
    
    # 若有指定文字顏色（例如詳解紅色 B40000），為 OMML 內所有文字 run 加入顏色
    hex_color = _rgb_to_hex(color)
    if hex_color:
        for r in omml_dom.xpath('.//m:r', namespaces=ns):
            rPr = r.find('m:rPr', namespaces=ns)
            if rPr is None:
                rPr = ET.Element('{http://schemas.openxmlformats.org/officeDocument/2006/math}rPr')
                r.insert(0, rPr)
            color_el = rPr.find('w:color', namespaces=ns)
            if color_el is None:
                color_el = ET.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
            color_el.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', hex_color)

    omml_xml_bytes = ET.tostring(omml_dom)
    return parse_xml(omml_xml_bytes)

def set_run_font(run, font_name="標楷體", ascii_font="Times New Roman", size_pt=13, bold=False, italic=False, superscript=False, color_rgb=(0,0,0)):
    """為 docx run 設定中英文字型、字級、粗斜體與顏色。"""
    run.font.name = ascii_font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if superscript:
        run.font.superscript = True
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
        
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)

def sanitize_text(text: str) -> str:
    """自動修正微軟 Symbol 字型私用區 (PUA) 缺字亂碼，例如 \uf0de -> ⇒。"""
    if not text or not isinstance(text, str):
        return text
    text = text.replace('\uf0de', '⇒')
    text = text.replace('\uf0e0', '⇒')
    text = text.replace('\uf0d8', '←')
    text = text.replace('\uf0da', '→')
    text = text.replace('\uf0db', '↔')
    return text

def add_item_to_paragraph(paragraph, item, default_font="標楷體", ascii_font="Times New Roman", default_size=13, force_color=None):
    """
    將單一文字項目或數學公式物件加入 Word 段落。
    
    支援格式：
    1. 數學公式字典：{"math": r"8\frac{3}{4}", "color": (180, 0, 0)}
    2. 一般 run 元組：("文字", bold, italic, size, superscript, color)
    3. 純文字字串（若內含 $...$ 會自動拆解為文字與公式）："已知長度為 $8\frac{3}{4}$ 公分"
    """
    # 情況 1: 字典指定數學式
    if isinstance(item, dict) and "math" in item:
        color = force_color or item.get("color", (0, 0, 0))
        omml_elem = latex_to_omml(item["math"], color=color)
        paragraph._p.append(omml_elem)
        return

    # 情況 2: 元組格式 ("math", r"...") 或 ("$", r"...")
    if isinstance(item, (tuple, list)) and len(item) >= 2 and item[0] in ("math", "$"):
        color = force_color or (item[2] if len(item) > 2 else (0, 0, 0))
        omml_elem = latex_to_omml(item[1], color=color)
        paragraph._p.append(omml_elem)
        return

    # 情況 3: 一般 run 元組
    if isinstance(item, (tuple, list)):
        text = sanitize_text(str(item[0]))
        bold = item[1] if len(item) > 1 else False
        italic = item[2] if len(item) > 2 else False
        size = item[3] if len(item) > 3 and item[3] is not None else default_size
        sup = False
        color = (0, 0, 0)
        if len(item) > 4:
            if isinstance(item[4], bool):
                sup = item[4]
                if len(item) > 5 and isinstance(item[5], (tuple, list)):
                    color = item[5]
            elif isinstance(item[4], (tuple, list)):
                color = item[4]

        if force_color is not None:
            color = force_color

        # 若文字中包含 $...$ 公式語法，進行智慧拆分處理
        if "$" in text and not sup:
            tokens = re.split(r'(\$.*?\$)', text)
            for token in tokens:
                if not token:
                    continue
                if token.startswith("$") and token.endswith("$") and len(token) >= 2:
                    latex_str = token[1:-1]
                    omml_el = latex_to_omml(latex_str, color=color)
                    paragraph._p.append(omml_el)
                else:
                    r = paragraph.add_run(token)
                    set_run_font(r, font_name=default_font, ascii_font=ascii_font, size_pt=size, bold=bold, italic=italic, superscript=sup, color_rgb=color)
        else:
            r = paragraph.add_run(text)
            set_run_font(r, font_name=default_font, ascii_font=ascii_font, size_pt=size, bold=bold, italic=italic, superscript=sup, color_rgb=color)
        return

    # 情況 4: 純字串
    if isinstance(item, str):
        item = sanitize_text(item)
        color = force_color or (0, 0, 0)
        if "$" in item:
            tokens = re.split(r'(\$.*?\$)', item)
            for token in tokens:
                if not token:
                    continue
                if token.startswith("$") and token.endswith("$") and len(token) >= 2:
                    latex_str = token[1:-1]
                    omml_el = latex_to_omml(latex_str, color=color)
                    paragraph._p.append(omml_el)
                else:
                    r = paragraph.add_run(token)
                    set_run_font(r, font_name=default_font, ascii_font=ascii_font, size_pt=default_size, color_rgb=color)
        else:
            r = paragraph.add_run(item)
            set_run_font(r, font_name=default_font, ascii_font=ascii_font, size_pt=default_size, color_rgb=color)

def add_exam_runs(paragraph, items_list, default_font="標楷體", ascii_font="Times New Roman", default_size=13, force_color=None):
    """將一系列項目（文字、公式、帶有格式之元組）依序加入 Word 段落。"""
    for item in items_list:
        add_item_to_paragraph(paragraph, item, default_font=default_font, ascii_font=ascii_font, default_size=default_size, force_color=force_color)
