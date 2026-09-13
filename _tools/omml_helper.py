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

def latex_to_omml(latex_code: str, color=None, size_pt=13):
    r"""
    將 LaTeX 數學式字串轉為 Word 原生 OMML (<m:oMath>) XML 元素。
    
    :param latex_code: LaTeX 語法字串，例如 r'8\frac{3}{4}' 或 r'\overline{AB}'
    :param color: 可選顏色，支援 (r, g, b) 數組或 HEX 字串（例如 'B40000'）
    :param size_pt: 方程式字級大小（點數），預設為大考標準 13pt（對應 OOXML sz=26）
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
    
    # 計算半點數 (Half-points)：13pt -> 26
    sz_val = str(int(size_pt * 2)) if size_pt else "26"
    hex_color = _rgb_to_hex(color)

    # 為 OMML 內所有文字 run (m:r) 加入 w:rPr 字級大小 (13pt) 與顏色設定
    for r in omml_dom.xpath('.//m:r', namespaces=ns):
        w_rPr = r.find('w:rPr', namespaces=ns)
        if w_rPr is None:
            w_rPr = ET.Element(f'{{{ns["w"]}}}rPr')
            r.insert(0, w_rPr)
            
        # 設定方程式字級 w:sz 與 w:szCs (確保在 Word 中與內文 13pt 完美一致)
        sz_el = w_rPr.find('w:sz', namespaces=ns)
        if sz_el is None:
            sz_el = ET.SubElement(w_rPr, f'{{{ns["w"]}}}sz')
        sz_el.set(f'{{{ns["w"]}}}val', sz_val)

        szCs_el = w_rPr.find('w:szCs', namespaces=ns)
        if szCs_el is None:
            szCs_el = ET.SubElement(w_rPr, f'{{{ns["w"]}}}szCs')
        szCs_el.set(f'{{{ns["w"]}}}val', sz_val)
        
        # 若有指定文字顏色（例如詳解紅色 B40000），設定 w:color
        if hex_color:
            color_el = w_rPr.find('w:color', namespaces=ns)
            if color_el is None:
                color_el = ET.SubElement(w_rPr, f'{{{ns["w"]}}}color')
            color_el.set(f'{{{ns["w"]}}}val', hex_color)

    omml_xml_bytes = ET.tostring(omml_dom)
    return parse_xml(omml_xml_bytes)

def set_run_font(run, font_name="標楷體", ascii_font="Times New Roman", size_pt=13, bold=False, italic=False, superscript=False, color_rgb=(0,0,0)):
    """為 docx run 設定中英文字型、字級、粗斜體與顏色。"""
    run.font.name = ascii_font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if isinstance(color_rgb, RGBColor):
        run.font.color.rgb = color_rgb
    elif isinstance(color_rgb, (tuple, list)):
        run.font.color.rgb = RGBColor(*color_rgb)
    
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)
    rFonts.set(qn('w:cs'), ascii_font)

def configure_document_normal_style(doc, size_pt=13):
    """
    將 Word 文件的「Normal (內文)」預設樣式字級設定為 13pt（半磅值 26），
    並設定中文字型為標楷體、英數為 Times New Roman。
    
    【核心機制】：
    Word 的方程式編輯器 (<m:oMath>) 的容器、架構字元（如分數線、根號、括號、底線等）
    高度依賴文件 Normal 樣式的字級繼承。若 Normal 樣式未改為 13pt，Word 方程式編輯器
    選取時仍會顯示 Normal 預設之 11pt。此函數從根本解決樣式層級字級繼承問題。
    """
    try:
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(size_pt)
        
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:eastAsia'), '標楷體')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')
        
        half_pts = str(int(round(size_pt * 2)))
        
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = parse_xml(f'<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{half_pts}"/>')
            rPr.append(sz)
        else:
            sz.set(qn('w:val'), half_pts)
            
        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = parse_xml(f'<w:szCs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{half_pts}"/>')
            rPr.append(szCs)
        else:
            szCs.set(qn('w:val'), half_pts)

        # 同步更新 styles.xml 的 w:docDefaults，徹底消除預設 theme 的 minorEastAsia (新細明體)
        styles_elem = doc.styles.element
        doc_defaults = styles_elem.find(qn('w:docDefaults'))
        if doc_defaults is not None:
            rPrDefault = doc_defaults.find(qn('w:rPrDefault'))
            if rPrDefault is not None:
                rPr_def = rPrDefault.find(qn('w:rPr'))
                if rPr_def is not None:
                    rFonts_def = rPr_def.find(qn('w:rFonts'))
                    if rFonts_def is not None:
                        for attr in ['asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme']:
                            if qn(f'w:{attr}') in rFonts_def.attrib:
                                del rFonts_def.attrib[qn(f'w:{attr}')]
                        rFonts_def.set(qn('w:eastAsia'), '標楷體')
                        rFonts_def.set(qn('w:ascii'), 'Times New Roman')
                        rFonts_def.set(qn('w:hAnsi'), 'Times New Roman')
                        rFonts_def.set(qn('w:cs'), 'Times New Roman')
                    sz_def = rPr_def.find(qn('w:sz'))
                    if sz_def is not None:
                        sz_def.set(qn('w:val'), half_pts)
                    szCs_def = rPr_def.find(qn('w:szCs'))
                    if szCs_def is not None:
                        szCs_def.set(qn('w:val'), half_pts)
    except Exception:
        pass

def upgrade_to_modern_word_mode(doc):
    """
    將 Word 文件升級為最新現代模式（Word 2013/2016/2019/2021/365），徹底移除「相容模式」限制。
    確保 Word 原生向量 SVG 圖形、OMML 數學公式與現代繪圖物件能以 100% 完整功能渲染，
    並原生解鎖「圖形格式」與「轉換為圖形 (Convert to Shape)」可編輯功能。
    """
    try:
        settings = doc.settings.element
        compat_nodes = settings.xpath('.//w:compatSetting[@w:name="compatibilityMode"]')
        if compat_nodes:
            for node in compat_nodes:
                node.set(qn('w:val'), '15')
        else:
            compat = settings.find(qn('w:compat'))
            if compat is None:
                compat = parse_xml('<w:compat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                settings.append(compat)
            compat.append(parse_xml('<w:compatSetting xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'))
    except Exception:
        pass

def ensure_document_font_consistency(doc, default_font="標楷體", ascii_font="Times New Roman", size_pt=13):
    """
    全卷字型與字級一致性深度固化引擎：
    1. 針對全文件 100% 的段落標記 (¶) 注入 <w:pPr><w:rPr>，顯式鎖定標楷體與 Times New Roman 13pt。
       徹底消除 Word 在跨文件「全選剪下／複製貼上」時，因段落標記繼承目標文件樣式而自動退回「新細明體」的微軟預設陷阱！
    2. 檢查全卷所有文字 Run，確保中文字型為標楷體、英數為 Times New Roman。
    3. 遍歷包含一般段落與所有表格儲存格內的段落。
    """
    half_pts = str(int(round(size_pt * 2)))

    def patch_paragraph(p):
        pPr = p._p.get_or_add_pPr()
        rPr = pPr.find(qn('w:rPr'))
        if rPr is None:
            rPr = parse_xml(
                f'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:rFonts w:ascii="{ascii_font}" w:eastAsia="{default_font}" w:hAnsi="{ascii_font}" w:cs="{ascii_font}"/>'
                f'<w:sz w:val="{half_pts}"/><w:szCs w:val="{half_pts}"/>'
                f'</w:rPr>'
            )
            pPr.append(rPr)
        else:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = parse_xml(
                    f'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    f'w:ascii="{ascii_font}" w:eastAsia="{default_font}" w:hAnsi="{ascii_font}" w:cs="{ascii_font}"/>'
                )
                rPr.append(rFonts)
            else:
                rFonts.set(qn('w:eastAsia'), default_font)
                if not rFonts.get(qn('w:ascii')):
                    rFonts.set(qn('w:ascii'), ascii_font)
                if not rFonts.get(qn('w:hAnsi')):
                    rFonts.set(qn('w:hAnsi'), ascii_font)
                if not rFonts.get(qn('w:cs')):
                    rFonts.set(qn('w:cs'), ascii_font)
            sz = rPr.find(qn('w:sz'))
            if sz is None:
                sz = parse_xml(f'<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{half_pts}"/>')
                rPr.append(sz)
            szCs = rPr.find(qn('w:szCs'))
            if szCs is None:
                szCs = parse_xml(f'<w:szCs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{half_pts}"/>')
                rPr.append(szCs)

        # 遍歷文字 Run
        for r in p.runs:
            if not r.text:
                continue
            run_rPr = r._r.get_or_add_rPr()
            rf = run_rPr.find(qn('w:rFonts'))
            if rf is None:
                rf = parse_xml(
                    f'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    f'w:ascii="{ascii_font}" w:eastAsia="{default_font}" w:hAnsi="{ascii_font}" w:cs="{ascii_font}"/>'
                )
                run_rPr.append(rf)
            else:
                if not rf.get(qn('w:eastAsia')) or rf.get(qn('w:eastAsia')) in ('新細明體', '細明體', 'SimSun', 'PMingLiU'):
                    rf.set(qn('w:eastAsia'), default_font)
                if not rf.get(qn('w:ascii')):
                    rf.set(qn('w:ascii'), ascii_font)
                if not rf.get(qn('w:hAnsi')):
                    rf.set(qn('w:hAnsi'), ascii_font)

    # 1. 處理主內文段落
    for p in doc.paragraphs:
        patch_paragraph(p)

    # 2. 處理表格內段落
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    patch_paragraph(p)

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

# 真正需要 2D 複合排版的結構（僅分數、根號、幾何線段橫槓、上下標與多行聯立才需要重型 OMML 原生方程式）
COMPLEX_MATH_PATTERNS = [
    r'\\frac', r'\\dfrac', r'\\tfrac',   # 分數
    r'\\sqrt',                          # 根號
    r'\\overline', r'\\overleftrightarrow', r'\\overrightarrow', # 幾何線段、直線、射線
    r'\^',                              # 上標 / 次方 (如 x^2, 3^2, (x-3)^2)
    r'_',                               # 下標 (如 x_1, y_2, a_n)
    r'\\begin', r'\\end',               # 矩陣或多行聯立式
    r'\\sum', r'\\int', r'\\lim',       # 高階微積分/級數
]
COMPLEX_MATH_REGEX = re.compile('|'.join(COMPLEX_MATH_PATTERNS))

def is_complex_math(latex_str: str) -> bool:
    """
    判斷 LaTeX 內容是否包含真正需要 2D 複合排版的結構。
    純數字、單一變數、純文字、平面等式（含 ×, ÷, ±, ≤, ≥, ≠, ∠, △ 等）回傳 False，轉為輕量級文字 run。
    """
    if not latex_str or not isinstance(latex_str, str):
        return False
    return bool(COMPLEX_MATH_REGEX.search(latex_str))

def add_simple_math_run(paragraph, text: str, default_font="標楷體", ascii_font="Times New Roman", size_pt=13, bold=False, color_rgb=(0,0,0)):
    """
    將純數字、單一代數變數、簡易等式、坐標與選項代號脫殼為輕量級 Word Run。
    - 英文字母變數（如 x, y, a, b, A, B）自動設為 Times New Roman 斜體 (Italic)
    - 數字、逗號、括號、等號、正負號設為 Times New Roman 正體 (Upright)
    - 徹底避免生成數百個 OMML 物件導致 Word 複製貼上單核 100% 卡死！
    """
    s = text.strip()
    s = s.replace(r'\,', ' ').replace(r'\;', ' ').replace(r'\quad', '  ').replace(r'\qquad', '   ')
    s = re.sub(r'\\text\{([^}]*)\}', r'\1', s)
    s = re.sub(r'\\rm\{([^}]*)\}', r'\1', s)
    s = s.replace(r'\degree', '°').replace(r'^\circ', '°')
    s = s.replace(r'\times', '×').replace(r'\div', '÷')
    s = s.replace(r'\pm', '±').replace(r'\mp', '∓')
    s = s.replace(r'\le', '≤').replace(r'\ge', '≥')
    s = s.replace(r'\neq', '≠').replace(r'\ne', '≠')
    s = s.replace(r'\approx', '≈').replace(r'\equiv', '≡')
    s = s.replace(r'\cdot', '·')
    s = s.replace(r'\perp', '⊥').replace(r'\parallel', '∥')
    s = s.replace(r'\sim', '∼').replace(r'\cong', '≅')
    s = s.replace(r'\angle', '∠').replace(r'\triangle', '△')
    s = s.replace(r'\in', '∈')
    s = s.replace(r'\pi', 'π').replace(r'\alpha', 'α').replace(r'\beta', 'β').replace(r'\theta', 'θ').replace(r'\lambda', 'λ')
    
    parts = re.split(r'([a-zA-Z]+)', s)
    for part in parts:
        if not part:
            continue
        is_alpha = part.isalpha()
        r = paragraph.add_run(part)
        set_run_font(
            r,
            font_name=default_font,
            ascii_font=ascii_font,
            size_pt=size_pt,
            bold=bold,
            italic=is_alpha,
            color_rgb=color_rgb
        )

def optimize_math_markdown(text: str, enable_optimization: bool = True) -> tuple[str, int, int]:
    """
    文字層級智慧算式分流演算法：
    1. 複合結構（分數、根號、幾何線段、上下標等）保留為 $...$。
    2. 簡易純數值、單一字母、簡易等式脫殼為普通純文字。
    """
    if not enable_optimization or not text or "$" not in text:
        return text, 0, text.count('$') // 2

    simplified_count = 0
    preserved_count = 0

    def inline_replacer(match):
        nonlocal simplified_count, preserved_count
        content = match.group(1).strip()
        if not content:
            return ""
        if is_complex_math(content):
            preserved_count += 1
            return f"${content}$"
        
        # 簡易算式脫殼
        s = content
        s = s.replace(r'\,', ' ').replace(r'\;', ' ').replace(r'\quad', '  ')
        s = re.sub(r'\\text\{([^}]*)\}', r'\1', s)
        s = re.sub(r'\\rm\{([^}]*)\}', r'\1', s)
        s = s.replace(r'\degree', '°').replace(r'^\circ', '°')
        simplified_count += 1
        return s

    processed_text = re.sub(r'(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)', inline_replacer, text)
    return processed_text, simplified_count, preserved_count

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
        latex_str = item["math"]
        size_val = item.get("size", default_size)
        if is_complex_math(latex_str):
            omml_elem = latex_to_omml(latex_str, color=color, size_pt=size_val)
            paragraph._p.append(omml_elem)
        else:
            add_simple_math_run(paragraph, latex_str, default_font=default_font, ascii_font=ascii_font, size_pt=size_val, color_rgb=color)
        return

    # 情況 2: 元組格式 ("math", r"...") 或 ("$", r"...")
    if isinstance(item, (tuple, list)) and len(item) >= 2 and item[0] in ("math", "$"):
        color = force_color or (item[2] if len(item) > 2 else (0, 0, 0))
        latex_str = item[1]
        size_val = item[3] if len(item) > 3 and item[3] is not None else default_size
        if is_complex_math(latex_str):
            omml_elem = latex_to_omml(latex_str, color=color, size_pt=size_val)
            paragraph._p.append(omml_elem)
        else:
            add_simple_math_run(paragraph, latex_str, default_font=default_font, ascii_font=ascii_font, size_pt=size_val, color_rgb=color)
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
                    if is_complex_math(latex_str):
                        omml_el = latex_to_omml(latex_str, color=color, size_pt=size)
                        paragraph._p.append(omml_el)
                    else:
                        add_simple_math_run(paragraph, latex_str, default_font=default_font, ascii_font=ascii_font, size_pt=size, bold=bold, color_rgb=color)
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
                    if is_complex_math(latex_str):
                        omml_el = latex_to_omml(latex_str, color=color, size_pt=default_size)
                        paragraph._p.append(omml_el)
                    else:
                        add_simple_math_run(paragraph, latex_str, default_font=default_font, ascii_font=ascii_font, size_pt=default_size, color_rgb=color)
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
