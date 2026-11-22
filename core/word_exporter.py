"""
Word Exporter — Chuyển LaTeX thành Word (.docx) với equations native
Pipeline: LaTeX → MathML → OMML → Word
Tốc độ: ~5-50ms per equation

Quy chuẩn định dạng:
- Font: Times New Roman, Cỡ 12pt
- Lề: Trên 2cm, Dưới 2cm, Trái 3cm, Phải 2cm
- Giãn dòng: Single (1.0), Space after: 4pt, Space before: 0pt
- Tự động tách đoạn mới trước mỗi câu hỏi (Câu 1., Câu 2...) và ý trắc nghiệm (A., B., C., D. hoặc a), b)...)
- Loại bỏ triệt để ký tự rác / ô vuông ẩn (như □, ⬚, control characters)
- Chèn trực tiếp hình ảnh / bảng biến thiên đã cắt vào văn bản
- Xử lý công thức OMML bên trong các ghi chú / placeholder (không để lộ dấu $ thô)
"""

import re
import os
import copy
from pathlib import Path
from PIL import Image as PILImage
from lxml import etree
import latex2mathml.converter
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from core.omml_converter import sanitize_plain_text_vectors, ensure_math_spacing_bridge


# ═══════════════════════════════════════════
# XSL TRANSFORM: MathML → OMML
# ═══════════════════════════════════════════

def find_mml2omml_xsl() -> str | None:
    """Tìm file MML2OMML.XSL từ mã nguồn dự án (chạy đám mây/Linux) hoặc Office installation."""
    # 1. Ưu tiên tìm ngay trong thư mục core/ hoặc thư mục gốc của dự án (hỗ trợ 100% Linux/Cloud/Docker)
    local_candidates = [
        Path(__file__).parent / "MML2OMML.XSL",
        Path(__file__).parent.parent / "MML2OMML.XSL",
        Path("core") / "MML2OMML.XSL",
    ]
    for lp in local_candidates:
        if lp.is_file():
            return str(lp.resolve())

    # 2. Tìm trong cài đặt Microsoft Office trên Windows (fallback)
    candidate_paths = [
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
        r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
        r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
        r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\root\Office15\MML2OMML.XSL",
    ]
    for p in candidate_paths:
        if os.path.isfile(p):
            return p
    return None


_xsl_transform = None

def get_xsl_transform():
    """Load and cache the XSLT transform."""
    global _xsl_transform
    if _xsl_transform is not None:
        return _xsl_transform

    xsl_path = find_mml2omml_xsl()
    if not xsl_path:
        raise FileNotFoundError(
            "Không tìm thấy MML2OMML.XSL! Cần Microsoft Office được cài đặt.\n"
            "Thường ở: C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\MML2OMML.XSL"
        )

    with open(xsl_path, 'rb') as f:
        xsl_doc = etree.parse(f)
    _xsl_transform = etree.XSLT(xsl_doc)
    return _xsl_transform


# ═══════════════════════════════════════════
# LATEX → OMML CONVERSION
# ═══════════════════════════════════════════

def wrap_negative_numbers(s: str) -> str:
    """
    Bọc số âm hoặc biểu thức âm đơn nguyên dạng {-2}, {-\\infty}
    khi đứng ở đầu công thức hoặc sau các ký tự (, [, {, ;, ,, =, <, >
    để MathML/OMML không nhận nhầm thành phép trừ hai ngôi sinh ô vuông rỗng.
    """
    # 1. Ở đầu chuỗi: -2, -x, -\infty...
    s = re.sub(r'^\s*([+\-])\s*([0-9]+(?:\.[0-9]+)?|[a-zA-Z]|\\[a-zA-Z]+(?:\{[^}]*\})*)', r'{\1\2}', s)
    # 2. Sau các ký tự (, [, {, ;, ,, =, <, >
    s = re.sub(r'([\(\[\{;,\=\<\>])\s*([+\-])\s*([0-9]+(?:\.[0-9]+)?|[a-zA-Z]|\\[a-zA-Z]+(?:\{[^}]*\})*)', r'\1{\2\3}', s)
    # Tránh lồng {{...}}
    s = re.sub(r'\{\{([^{}]+)\}\}', r'{\1}', s)
    return s


def clean_latex_math(latex_str: str) -> str:
    """Làm sạch và chuẩn hóa công thức toán trước khi convert."""
    s = latex_str.strip()
    # Xóa các ký tự ô vuông rác trong công thức (kể cả ô vuông đứt nét ⬚ \u2b1a)
    s = re.sub(r'[\u25a0-\u25ff\u2b12-\u2b2f\u2500-\u257f\ufffd\u200b-\u200f\ufeff]', '', s)
    # Chuẩn hóa vector: chuyển \overrightarrow{AB} thành \vec{AB} (cho OMML)
    s = re.sub(r'\\overrightarrow\{([^}]+)\}', r'\\vec{\1}', s)
    # Chuẩn hóa vector biến thể B'D'→, u→, B'D'\vec, u\vec (CHỈ CHẤP NHẬN CHỮ CÁI, CẤM SỐ VÀ CẤM ngoặc {)
    vec_unit = r"(?:[A-Z][']*(?:\s*[A-Z][']*)+|[a-z][']*)"
    s = re.sub(r"(?<![\\a-zA-Z0-9])(" + vec_unit + r")\s*[\u2192→]", r"\\vec{\1}", s)
    s = re.sub(r"(?<![\\a-zA-Z0-9])(" + vec_unit + r")\\vec(?!\{)(?![a-zA-Z])", r"\\vec{\1}", s)
    s = re.sub(r"\\vec\s+(" + vec_unit + r")", r"\\vec{\1}", s)

    # 🎯 Chuẩn hóa hệ phương trình & tuyển nghiệm dùng aligned/align* sang cases hoặc matrix:
    # 1. Hệ phương trình có ngoặc nhọn: \left\{ \begin{aligned} ... \end{aligned} \right. -> \begin{cases} ... \end{cases}
    s = re.sub(r'\\left\\\{\s*\\begin\{(?:aligned|align\*?)\}([\s\S]*?)\\end\{(?:aligned|align\*?)\}\s*(?:\\right\.?|(?=[^a-zA-Z0-9\\]|$))', r'\\begin{cases}\1\\end{cases}', s)
    # 2. Tuyển nghiệm có ngoặc vuông: \left[ \begin{aligned} ... \end{aligned} \right. -> \left[ \begin{matrix} ... \end{matrix} \right.
    s = re.sub(r'\\left\[\s*\\begin\{(?:aligned|align\*?)\}([\s\S]*?)\\end\{(?:aligned|align\*?)\}\s*(?:\\right\.?|(?=[^a-zA-Z0-9\\]|$))', r'\\left[ \\begin{matrix}\1\\end{matrix} \\right.', s)
    # 3. Các môi trường aligned/align* còn lại: chuyển thành matrix
    s = re.sub(r'\\begin\{(?:aligned|align\*?)\}([\s\S]*?)\\end\{(?:aligned|align\*?)\}', r'\\begin{matrix}\1\\end{matrix}', s)
    # 4. Gọt bỏ ký tự & căn lề gây lỗi <mi>&</mi> trong môi trường matrix/cases
    s = re.sub(r'(\\begin\{(?:matrix|cases|pmatrix|bmatrix)\}\s*)&', r'\1', s)
    s = re.sub(r'(\\\\\s*)&', r'\1', s)

    # 🎯 Chuẩn hóa phân số thiếu ngoặc nhọn dạng \frac{1}2 hoặc \frac12 -> \frac{1}{2}
    s = re.sub(r'\\frac\{([^{}]+)\}([a-zA-Z0-9])', r'\\frac{\1}{\2}', s)
    s = re.sub(r'\\frac([a-zA-Z0-9])([a-zA-Z0-9])', r'\\frac{\1}{\2}', s)

    # Bọc số âm/dấu trừ đơn nguyên tránh lỗi ô vuông đứt nét
    s = wrap_negative_numbers(s)
    return s


def latex_to_mathml(latex_str: str) -> str:
    """Convert LaTeX math string to MathML với ép form=prefix cho dấu âm."""
    clean_s = clean_latex_math(latex_str)
    try:
        mathml = latex2mathml.converter.convert(clean_s)
        # 🛡️ Lá chắn kép: Khử triệt để ký tự & trần không nằm trong thực thể XML (&amp;, &lt;, &gt;, &#...;)
        # ngăn chặn 100% lỗi xmlParseEntityRef: no name làm vỡ chuyển đổi OMML
        mathml = re.sub(r'&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)', '&amp;', mathml)

        # Ép form="prefix" cho các toán tử -, + đứng đầu hoặc sau ngoặc/phân cách
        mml_doc = etree.fromstring(mathml.encode('utf-8'))
        for mo in mml_doc.xpath('.//*[local-name()="mo"]'):
            if mo.text in ['-', '+', '\u2212']:
                parent = mo.getparent()
                idx = parent.index(mo) if parent is not None else 0
                is_prefix = False
                if idx == 0:
                    is_prefix = True
                elif idx > 0 and parent[idx - 1].text in ['(', '[', '{', ';', ',', '=']:
                    is_prefix = True
                if is_prefix:
                    mo.attrib['form'] = 'prefix'
        return etree.tostring(mml_doc, encoding='utf-8').decode('utf-8')
    except Exception:
        escaped = clean_s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return f'<math xmlns="http://www.w3.org/1998/Math/MathML"><mtext>{escaped}</mtext></math>'


def mathml_to_omml(mathml_str: str) -> etree._Element:
    """Convert MathML string to OMML element using XSL transform và lọc bỏ <m:e> rỗng."""
    transform = get_xsl_transform()
    mathml_doc = etree.fromstring(mathml_str.encode('utf-8'))

    if mathml_doc.tag != '{http://www.w3.org/1998/Math/MathML}math':
        wrapper = etree.Element('{http://www.w3.org/1998/Math/MathML}math')
        wrapper.append(mathml_doc)
        mathml_doc = wrapper

    tree = etree.ElementTree(mathml_doc)
    omml_tree = transform(tree)
    omml_root = omml_tree.getroot()

    # Lọc sạch triệt để mọi phần tử <m:e> rỗng không có nội dung chữ
    for e in omml_root.xpath('.//*[local-name()="e"]'):
        t_nodes = e.xpath('.//*[local-name()="t"]')
        has_text = any(t.text and t.text.strip() for t in t_nodes)
        if not has_text:
            p = e.getparent()
            if p is not None:
                p.remove(e)

    return omml_root


def latex_to_omml(latex_str: str) -> etree._Element:
    """Convert LaTeX math to OMML element."""
    mathml = latex_to_mathml(latex_str)
    omml = mathml_to_omml(mathml)
    return omml


# ═══════════════════════════════════════════
# CLEANING & PARSING VĂN BẢN
# ═══════════════════════════════════════════

def remove_page_footers(text: str) -> str:
    """
    Lọc bỏ triệt để các chuỗi Footer chân trang đề thi scan:
    - 'Trang 1/4', 'Trang 2/4', 'Trang 3/4', 'Trang 4/4'
    - 'Đề kiểm tra cuối kỳ 1, môn Toán 12, Trang 3/4 - Mã đề thi 132'
    - 'Trang 4/4 - Mã đề 132'
    - 'Mã đề thi 132' (khi đứng ở chân trang)
    - '--- HẾT ---', 'Cán bộ coi thi không giải thích gì thêm...'
    - 'Thí sinh không được sử dụng tài liệu...'
    Bảo vệ tuyệt đối mã đề trong thẻ [TIÊU_ĐỀ_ĐỀ_THI: ...] ở đầu trang 1 không bị xóa nhầm.
    """
    if not text:
        return ""

    # Bảo vệ khối tiêu đề đề thi nếu có ở đầu đề
    hdr_map = {}
    def mask_hdr(m):
        t_id = f"@@__EXAM_HDR_FOOTER_PROTECT_{len(hdr_map)}__@@"
        hdr_map[t_id] = m.group(0)
        return t_id

    text = re.sub(r'\[TIÊU_ĐỀ_ĐỀ_THI:[\s\S]*?\]', mask_hdr, text, flags=re.IGNORECASE)

    # Các dòng chân trang scan thường gặp
    footer_patterns = [
        # Dạng đầy đủ: [Đề kiểm tra...] Trang X/Y [- Mã đề...]
        r'(?i)(?:\$\s*)?(?:Đề\s+kiểm\s+tra[^\n]*?)?Trang\s+\d+\s*(?:/\s*\d+)?\s*(?:[-–—]\s*Mã\s+đề(?:\s+thi)?\s*[:=]?\s*\d+)?(?:\s*\$)?',
        # Dạng Trang X/Y đứng độc lập
        r'(?i)(?:\$\s*)?\bTrang\s+\d+\s*/\s*\d+\b(?:\s*\$)?',
        # Dạng Mã đề thi XXX đứng ở cuối dòng hoặc một mình
        r'(?i)(?:^|[\n\r])\s*(?:\$\s*)?Mã\s+đề(?:\s+thi)?\s*[:=]?\s*\d+(?:\s*\$)?(?=\s*[\n\r]|$)',
        # Dạng --- HẾT ---
        r'(?i)[-–—\s]*\bHẾT\b[-–—\s]*',
        # Dòng cán bộ coi thi
        r'(?i)Cán\s+bộ\s+coi\s+thi\s+không\s+giải\s+thích\s+gì\s+thêm[^\n]*',
        # Dòng thí sinh không được sử dụng tài liệu
        r'(?i)Thí\s+sinh\s+không\s+được\s+sử\s+dụng\s+tài\s+liệu[^\n]*'
    ]

    for pat in footer_patterns:
        text = re.sub(pat, '', text)

    # Khôi phục khối tiêu đề đề thi
    for t_id, orig_hdr in hdr_map.items():
        text = text.replace(t_id, orig_hdr)

    # Thu gọn các dòng trống liên tiếp do xóa footer sinh ra
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def clean_exam_text(text: str) -> str:
    """
    Chuẩn hóa các nhãn câu hỏi và điểm số:
    1. Gỡ bỏ dấu $ bao quanh số thứ tự câu hỏi: Câu $1$: -> Câu 1:, Câu $12$: -> Câu 12:, Bài $1$: -> Bài 1:
    2. Tách hoàn toàn điểm số tự luận ra ngoài dấu $...$:
       - $(0,75\\text{ điểm})$ hoặc $(0,75\\text{ điểm}):$ -> (0,75 điểm):
       - ($0,75$ điểm) hoặc ($1,5$ điểm) -> (1,5 điểm)
       - $(0,75 điểm)$ -> (0,75 điểm)
    """
    if not text:
        return ""

    # 0. Lọc bỏ triệt để chân trang scan
    text = remove_page_footers(text)

    # 1. Gỡ bỏ dấu $ bao quanh số thứ tự câu hỏi
    text = re.sub(r'\b(Câu|Bài|Đề|Ví dụ)\s*\$(\d+)\$', r'\1 \2', text)
    text = re.sub(r'\b(Câu|Bài|Đề|Ví dụ)\s+(\d+)\s*\$', r'\1 \2', text)
    text = re.sub(r'\$(Câu|Bài|Đề|Ví dụ)\s*(\d+)\$?', r'\1 \2', text)

    # 2. Chuẩn hóa điểm số tự luận: tách hoàn toàn ra ngoài dấu $...$
    def repl_score(m):
        num = m.group('num').replace('.', ',')
        colon = ":" if ":" in m.group(0) else ""
        return f" ({num} điểm){colon}"

    score_pattern = re.compile(
        r'\$?\s*\(\s*\$?\s*(?P<num>\d+(?:[,\.]\d+)?)\s*\$?\s*(?:\\(?:text|rm|mathrm)\s*\{)?\s*điểm\}?\s*\$?\s*\)\s*[:]?\s*\$?',
        re.IGNORECASE
    )
    text = score_pattern.sub(repl_score, text)

    # Dọn dẹp khoảng cách thừa quanh nhãn câu hỏi và điểm số, đưa về chuẩn "Câu 1 (0,75 điểm): "
    text = re.sub(r'\b((?:Câu|Bài|Đề|Ví dụ)\s+\d+)\s*:\s*\((\d+(?:,\d+)?\s+điểm)\)', r'\1 (\2):', text)
    text = re.sub(r'\b((?:Câu|Bài|Đề|Ví dụ)\s+\d+)\s+\((\d+(?:,\d+)?\s+điểm)\)\s*:', r'\1 (\2):', text)
    text = re.sub(r'(\(\d+(?:,\d+)?\s+điểm\):?)([^\s\n])', r'\1 \2', text)
    text = re.sub(r'\b((?:Câu|Bài|Đề|Ví dụ)\s+\d+)\s*:', r'\1:', text)
    text = re.sub(r'(\b(?:Câu|Bài|Đề|Ví dụ)\s+\d+(?:\s*\([^\)\n]*\))?[:\.])([^\s\n])', r'\1 \2', text)

    return text


def mask_math_blocks(text: str) -> tuple[str, dict[str, str]]:
    """Tạm thời thay thế các khối $$...$$ và $...$ bằng token ẩn để bảo vệ công thức toán học."""
    if not text:
        return "", {}
    math_map = {}
    counter = 0
    math_pat = re.compile(r'(\$\$[\s\S]*?\$\$|\$(?:[^\$\\]|\\.)*\$)')

    def repl(m):
        nonlocal counter
        tok = f"@@__MATH_BLOCK_{counter}__@@"
        math_map[tok] = m.group(0)
        counter += 1
        return tok

    masked_text = math_pat.sub(repl, text)
    return masked_text, math_map


def unmask_math_blocks(text: str, math_map: dict[str, str]) -> str:
    """Khôi phục lại các khối Math nguyên vẹn."""
    if not text or not math_map:
        return text
    for tok, orig in math_map.items():
        text = text.replace(tok, orig)
    return text


# Regex nhận diện toàn bộ thẻ placeholder [HÌNH_VẼ...], [BẢNG...] kể cả có [...] lồng nhau hoặc thiếu ] cuối dòng
PLACEHOLDER_REGEX = re.compile(
    r'\[\s*(?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE)\b(?:[^\[\]]|\[[^\]]*\])*\]?',
    re.IGNORECASE
)


def mask_placeholders(text: str) -> tuple[str, dict[str, str]]:
    """Tạm thời thay thế các thẻ placeholder bằng token an toàn để tránh bị các regex làm sạch phá hỏng."""
    if not text:
        return "", {}
    mapping = {}
    counter = 0

    def repl(m):
        nonlocal counter
        token = f"@@__PLACEHOLDER_IMG_TAG_{counter}__@@"
        raw_tag = m.group(0).strip()
        # Đảm bảo thẻ luôn có ] đóng ở cuối
        if not raw_tag.endswith(']'):
            raw_tag += ']'
        mapping[token] = raw_tag
        counter += 1
        return token

    masked_text = PLACEHOLDER_REGEX.sub(repl, text)
    return masked_text, mapping


def unmask_placeholders(text: str, mapping: dict[str, str]) -> str:
    """Khôi phục lại các thẻ placeholder nguyên vẹn."""
    if not text or not mapping:
        return text
    for token, raw_tag in mapping.items():
        text = text.replace(token, raw_tag)
    return text


def sanitize_text_aesthetics(text: str) -> str:
    """
    Chuẩn hóa thẩm mỹ và xử lý rác OCR:
    - Sửa số thập phân dính vạch đứng: 13|5 -> 13,5
    - Xóa khoảng trắng trong tên hình hộp: ABCD. A'B'C'D' -> ABCD.A'B'C'D'
    - Xóa khoảng trắng trước các dấu ngắt câu: Oxyz , -> Oxyz,, gốc O , -> gốc O,
    - Xóa khoảng trắng giữa công thức toán và dấu ngắt câu: (C) $ . -> (C)$.
    - Chuẩn hóa khoảng cách sát dấu $: $ A' $ -> $A'$
    - Tách rời các từ tiếng Việt toán học bị dính nhau (do OCR hoặc LaTeX math mode artifact)
    """
    if not text:
        return ""

    # 1. Nắn số thập phân dính vạch đứng (OCR đọc nhầm dấu phẩy thành |)
    text = re.sub(r'(\d+)\|(\d+)', r'\1,\2', text)

    # 2. Xóa khoảng trắng trong tên hình hộp/lăng trụ (ABCD. A'B'C'D' -> ABCD.A'B'C'D')
    text = re.sub(r'\b([A-Z]{3,8}\.)\s+([A-Z\']{3,8})\b', r'\1\2', text)

    # 3. Tách rời các từ tiếng Việt toán học dính nhau phổ biến từ OCR / Math mode
    stuck_words_map = [
        (r'\bđiquađiểm\b', 'đi qua điểm'),
        (r'\bđiqua\b', 'đi qua'),
        (r'\blàgiaobiểmcủacácđườngtiệmcậncủađồthị\b', 'là giao điểm của các đường tiệm cận của đồ thị'),
        (r'\blàgiaobiểm\b', 'là giao điểm'),
        (r'\bgiaobiểmcủacácđườngtiệmcậncủađồthị\b', 'giao điểm của các đường tiệm cận của đồ thị'),
        (r'\bgiaobiểmcủa\b', 'giao điểm của'),
        (r'\bgiaobiểm\b', 'giao điểm'),
        (r'\btiệmcậncủa\b', 'tiệm cận của'),
        (r'\btiệmcậnđứng\b', 'tiệm cận đứng'),
        (r'\btiệmcậnngang\b', 'tiệm cận ngang'),
        (r'\btiệmcận\b', 'tiệm cận'),
        (r'\bcủacácđườngtiệmcậncủađồthị\b', 'của các đường tiệm cận của đồ thị'),
        (r'\bcủacácđườngtiệmcận\b', 'của các đường tiệm cận'),
        (r'\bcủacácđường\b', 'của các đường'),
        (r'\bcủađồthị\b', 'của đồ thị'),
        (r'\bcóđồthịlà\b', 'có đồ thị là'),
        (r'\bcóđồthị\b', 'có đồ thị'),
        (r'\bđồthị\b', 'đồ thị'),
        (r'\bhàmsố\b', 'hàm số'),
        (r'\bđồngbiến\b', 'đồng biến'),
        (r'\bnghịchbiến\b', 'nghịch biến'),
        (r'\bcựctrị\b', 'cực trị'),
        (r'\bcựcđại\b', 'cực đại'),
        (r'\bcựctiểu\b', 'cực tiểu'),
        (r'\btọađộcủa\b', 'tọa độ của'),
        (r'\btọađộ\b', 'tọa độ'),
        (r'\btamgiác\b', 'tam giác'),
        (r'\bhìnhhộp\b', 'hình hộp'),
        (r'\bphươngtrình\b', 'phương trình'),
        (r'\bmặtphẳng\b', 'mặt phẳng'),
        (r'\bđườngthẳng\b', 'đường thẳng'),
        (r'\bchuvi\b', 'chu vi'),
        (r'\bdiệntích\b', 'diện tích'),
        (r'\bthểtích\b', 'thể tích'),
    ]
    for pat, rep in stuck_words_map:
        text = re.sub(pat, rep, text, flags=re.IGNORECASE)

    # 4. Xóa khoảng trắng trước các dấu ngắt câu , ; . ? ! :
    text = re.sub(r'[ \t]+([,;\.\?!:])', r'\1', text)

    # 5. Xóa khoảng trắng giữa $ và dấu ngắt câu kế tiếp: $ . -> $. hoặc $ , -> $,
    text = re.sub(r'(\$)[ \t]+([,;\.\?!:])', r'\1\2', text)

    # 6. Chuẩn hóa khoảng trắng sát dấu $ bên trong công thức: $ expr $ -> $expr$
    def _trim_math_token(m):
        inner = m.group(1).strip()
        return f"${inner}$"
    text = re.sub(r'(?<!\$)\$(?!\$)([^\$\n]+?)(?<!\$)\$(?!\$)', _trim_math_token, text)

    # 7. Xóa dấu $ rỗng
    text = re.sub(r'\$[ \t]*\$', '', text)

    # 8. Xóa khoảng trắng kép (collapsing multiple spaces to single space)
    text = re.sub(r'[ \t]{2,}', ' ', text)

    return text


def is_real_math_content(content: str) -> bool:
    """
    Kiểm tra xem nội dung trong $...$ có phải là biểu thức toán học thực thụ hay không.
    Nếu là toán học thực thụ (chứa lệnh LaTeX, chỉ số dưới, số mũ, dấu so sánh,
    ký hiệu viết tắt toán học như x_{CĐ}, x_{CT}, S_{đáy}, V_{chóp}...),
    tuyệt đối BẢO VỆ NGUYÊN VẸN, KHÔNG ĐƯỢC gỡ bỏ dấu $.
    """
    if not content or not content.strip():
        return False

    clean = content.strip()

    # Bóc tách phần text bên ngoài các khối lệnh {...} và \text{...}
    # để kiểm tra xem có chứa câu văn tiếng Việt cấu trúc không
    outer_text = re.sub(r'\\[a-zA-Z]+\{[^{}]*\}|\{[^{}]*\}', ' ', clean)

    # Nếu phần text bên ngoài chứa các từ nối/hư từ cấu trúc câu tiếng Việt:
    # (Ví dụ: "Hàm số đồng biến trên khoảng", "Đi qua điểm", "với mọi giá trị")
    # -> Đây là câu tiếng Việt bị bọc nhầm trong $!
    sentence_vn_pattern = re.compile(
        r'\b(hàm\s+số|đồng\s+biến|nghịch\s+biến|đi\s+qua|giao\s+điểm|tiệm\s+cận|'
        r'mệnh\s+đề|khẳng\s+định|phương\s+trình|nghiệm|với\s+mọi|tồn\s+tại|'
        r'thỏa\s+mãn|liên\s+tục|xác\s+định|tập\s+xác\s+định|giá\s+trị\s+lớn\s+nhất|'
        r'giá\s+trị\s+nhỏ\s+nhất|trên\s+khoảng|trên\s+đoạn|trong\s+khoảng)\b',
        re.IGNORECASE
    )
    if sentence_vn_pattern.search(outer_text):
        return False

    # 1. Chứa các lệnh LaTeX toán học đặc thù
    if re.search(r'\\(?:frac|dfrac|tfrac|sqrt|vec|overrightarrow|int|sum|prod|lim|infty|alpha|beta|gamma|delta|Delta|pi|theta|le|ge|ne|approx|equiv|pm|mp|cdot|times|in|notin|subset|cup|cap|partial|forall|exists|log|ln|sin|cos|tan|cot|left|right|begin|end|matrix|cases)\b', clean):
        return True

    # 2. Chứa chỉ số dưới hoặc số mũ (subscript / superscript)
    # Ví dụ: x_{CĐ}, x_{CT}, y_{CĐ}, S_{đáy}, V_{chóp}, x_1, a^2, 4x^2, x_{1,2}
    if re.search(r'[\^_]\{[^{}]+\}|[\^_][a-zA-Z0-9à-ỹÀ-Ỹ]', clean):
        return True

    # 3. Chứa đẳng thức hoặc bất đẳng thức toán học (=, <, >, <=, >=, !=, \le, \ge, \ne)
    # Ví dụ: x_{CĐ} = 4, y = 2x - 1, m > 0, -1 <= x <= 2, f(x) = 0
    if re.search(r'[-+*/]?\s*(?:=|<=|>=|!=|<|>|\\le\b|\\ge\b|\\ne\b)\s*[-+*/]?', clean):
        if re.search(r'[a-zA-Z0-9\(\)\{\}\^_\'\"]\s*(?:=|<=|>=|!=|<|>|\\le\b|\\ge\b|\\ne\b)\s*[-+]?[a-zA-Z0-9\(\)\{\}\^_\'\"]', clean):
            if not re.search(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]', outer_text, re.IGNORECASE):
                return True

    # 4. Chứa tọa độ điểm hoặc bộ số đơn lập thuần túy
    # Ví dụ: (1; 2), (-1; 3; -2), A(0; 2)
    if re.match(r'^(?:[A-Z][\'\d_]*\s*)?\(\s*[-+]?[a-zA-Z0-9_\.]+(?:\s*[;,]\s*[-+]?[a-zA-Z0-9_\.]+)+\s*\)$', clean):
        return True

    # 5. Chứa phép toán đại số cơ bản với biến số (+, -, *, /)
    # Ví dụ: 2x + 1, a - b, 3x^2 - 4x + 1
    if re.search(r'\b[a-zA-Z]\s*[-+*/]\s*[a-zA-Z0-9]|\b\d+\s*[-+*/]\s*[a-zA-Z]', clean):
        if not re.search(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]', outer_text, re.IGNORECASE):
            return True

    return False


def unwrap_vietnamese_math(text: str) -> str:
    """
    Gỡ bỏ dấu $...$ bao bọc cả câu tiếng Việt (kể cả khi rách dòng qua nhiều \\n)
    trước khi tách paragraph để text không bị dính chữ (như 'điquađiểm', 'làgiaobiểm...').
    Bỏ qua display math $$...$$.
    BẢO TOÀN TUYỆT ĐỐI 100% CÁC CÔNG THỨC TOÁN THỰC THỤ (chứa x_{CĐ}, x_{CT}, S_{đáy}, v.v.).
    """
    if not text:
        return ""

    vn_char_pattern = re.compile(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]', re.IGNORECASE)
    vn_word_pattern = re.compile(
        r'\b(và|hoặc|với|là|khi|thì|nếu|của|trong|trên|dưới|điểm|đồ\s+thị|véc[ -]tơ|vectơ|giao\s+điểm|tiệm\s+cận|đồng\s+biến|nghịch\s+biến|cực\s+trị|có|đi\s+qua|tọa\s+độ|mặt\s+phẳng|đường\s+thẳng|tam\s+giác|hình\s+hộp|chu\s+vi|diện\s+tích|thể\s+tích|tâm|trọng\s+tâm|đỉnh|tiếp\s+tuyến|nghiệm|mệnh\s+đề|khẳng\s+định|phương\s+trình|hàm\s+số|lớn\s+hơn|nhỏ\s+hơn|bằng|thuộc|không\s+thuộc)\b',
        re.IGNORECASE
    )

    def repl_vn(m):
        full = m.group(0)
        content = m.group(1)
        if full.startswith('$$'):
            return full
        # NẾU LÀ TOÁN HỌC THỰC THỤ: BẢO VỆ NGUYÊN VẸN 100%, TUYỆT ĐỐI KHÔNG UNWRAP!
        if is_real_math_content(content):
            return full
        if vn_char_pattern.search(content) or vn_word_pattern.search(content):
            clean_content = content.strip()
            return f" {clean_content} "
        return full

    parts = re.split(r'(\$\$[\s\S]*?\$\$)', text)
    out_parts = []
    for part in parts:
        if part.startswith('$$'):
            out_parts.append(part)
        else:
            unwrapped = re.sub(r'(?<!\$)\$([^\$]{1,800}?)\$(?!\$)', repl_vn, part, flags=re.DOTALL)
            out_parts.append(unwrapped)

    res = "".join(out_parts)
    return sanitize_text_aesthetics(res)


def heal_broken_math_and_coords(text: str) -> str:
    """
    Hàn gắn triệt để các tọa độ và công thức toán bị ngắt dòng ngang từ OCR/Gemini:
    - Tọa độ 2D/3D bị ngắt dòng: (0;\n2), (-1; 3; -\n2), (1; 2;\n1), A'(0;0;\n4), (2;3;\n4)
    - Tọa độ vectơ: \\vec{a} = (-1; 3; -\n2), \\vec{u} = (1; 2;\n1)
    - Đồ thị (C) bị rách dòng: có đồ thị là $\\n(C)$ hoặc đồ thị $\\n(C)$
    - Cặp dấu $...$ bị rách qua dòng
    - Câu chỉ dẫn Phần II bị rách dòng: 'Trong mỗi ý\\na), b), c), d) ở mỗi câu...'
    """
    if not text:
        return ""

    # Dọn dẹp thẩm mỹ trước
    text = sanitize_text_aesthetics(text)
    # Gỡ bỏ $...$ ôm câu tiếng Việt trước
    text = unwrap_vietnamese_math(text)

    # 1. Hàn gắn tọa độ có dấu âm hoặc phân cách bị ngắt dòng ngang:
    for _ in range(5):
        text = re.sub(r'(\([^\)\n]*?[-+])\s*\n\s*([0-9a-zA-Z_\.]+\s*\))', r'\1\2', text)
        text = re.sub(r'(\([^\)\n]*?[;,])\s*\n\s*([-+]?\s*[0-9a-zA-Z_\.]+\s*\))', r'\1 \2', text)
        text = re.sub(r'(\([^\)\n]*?[;,])\s*\n\s*([-+]?\s*[0-9a-zA-Z_\.]+\s*[;,])', r'\1 \2', text)
        text = re.sub(r'(\b[A-Z][\'\d_]*\s*\([^\)\n]*?[;,])\s*\n\s*([-+]?\s*[0-9a-zA-Z_\.]+\s*\))', r'\1 \2', text)
        text = re.sub(r'(\b[A-Z][\'\d_]*\s*\([^\)\n]*?[;,])\s*\n\s*([-+]?\s*[0-9a-zA-Z_\.]+\s*[;,])', r'\1 \2', text)
        text = re.sub(r'(\$\s*[A-Z]?[\'\d_]*\s*\([^\)\n]*?[;,])\s*\n\s*([-+]?\s*[0-9a-zA-Z_\.]+\s*\)\s*\$)', r'\1 \2', text)

    # 2. Hàn gắn đồ thị (C) bị ngắt dòng sau 'đồ thị', 'đồ thị là', 'của đồ thị':
    for _ in range(3):
        text = re.sub(
            r'(?i)\b((?:có\s+)?đồ\s+thị(?:\s+là)?|của\s+đồ\s+thị|đường\s+cong|đường\s+tròn|mặt\s+cầu|mặt\s+phẳng)\s*\$?([ \t]*\n[ \t]*)\$?\(?([A-Za-z0-9_]+)\)?\s*\$?',
            r'\1 $(\3)$',
            text
        )

    # 3. Hàn gắn cặp dấu $...$ bị rách qua dòng:
    for _ in range(4):
        text = re.sub(r'(?<!\$)\$([^\$\n]{1,400}?)\n\s*([^\$\n]{1,400}?)\$(?!\$)', r'$\1 \2$', text)

    # 4. Hàn gắn câu chỉ dẫn Phần II bị rách dòng:
    text = re.sub(r'(?i)(Trong\s+mỗi\s+ý)\s*\n\s*([a-d]\s*[\)\/].*?đúng\s+hoặc\s+sai)', r'\1 \2', text)
    text = re.sub(r'(?i)(Thí\s+sinh\s+trả\s+lời\s+từ\s+câu\s+\d+\s+đến\s+câu\s+\d+\.)\s*\n\s*(Trong\s+mỗi\s+ý)', r'\1 \2', text)

    # 5. Chuẩn hóa khoảng cách sau dấu chấm phẩy trong tọa độ (x; y) hoặc (x; y; z)
    for _ in range(4):
        text = re.sub(r'(\([^\)\n]*?);(?!\s)', r'\1; ', text)

    # 6. Dọn dẹp thẩm mỹ lần nữa
    text = sanitize_text_aesthetics(text)
    return text


def join_broken_ocr_lines(text: str) -> str:
    """
    Nối lại từ bị ngắt dòng sai từ OCR (ví dụ: 'hàng phần chụ\\nc)' -> 'hàng phần chục)').
    Ghép các dòng bị ngắt ngang giữa từ nếu:
    - Dòng trước không kết thúc bằng dấu chấm câu (., ?, !, …) hoặc thẻ đóng/display math
    - Dòng sau không bắt đầu bằng Câu, Bài, Đề, Ví dụ hoặc A., B., C., D. hoặc thẻ [HÌNH_VẼ / [BẢNG / $$
    - Bắt buộc ghép nếu dòng trước có ngoặc đơn '(' chưa đóng hoặc dấu '$' lẻ hoặc kết thúc bằng toán tử/từ nối.
    """
    if not text:
        return ""

    # Bảo vệ thẻ placeholder trước khi ghép dòng
    text, tag_map = mask_placeholders(text)
    # Chuẩn hóa nhãn câu hỏi (gỡ $) và tách điểm số trước khi ghép dòng
    text = clean_exam_text(text)
    # Hàn gắn triệt để tọa độ và công thức toán học bị ngắt dòng ngang
    text = heal_broken_math_and_coords(text)

    # 1. Ghép từ bị gãy có dấu gạch nối cuối dòng: từ-\nphần -> từphần
    text = re.sub(r'([a-zA-Zà-ỹÀ-Ỹ0-9_]+)-\s*\n\s*([a-zA-Zà-ỹÀ-Ỹ0-9_]+)', r'\1\2', text)

    # 2. Ghép trường hợp đặc thù phụ âm cuối hoặc đuôi ngoặc bị rớt xuống dòng:
    # ví dụ: "chụ\nc)" -> "chục)", "chụ\n c)" -> "chục)"
    text = re.sub(r'([a-zA-Zà-ỹÀ-Ỹ]+)\s*\n\s*([cmnpt]|ng|ch|nh)(\)|\]|[\s,;:\.]|$)', r'\1\2\3', text)
    # Và trường hợp từ bị ngắt trước dấu ngoặc đóng:
    text = re.sub(r'([a-zA-Zà-ỹÀ-Ỹ]+)\s*\n\s*([a-zA-Zà-ỹÀ-Ỹ]+[\)\]])', r'\1\2', text)

    # 3. Quét dòng và ghép nếu dòng trước không kết thúc bằng dấu chấm câu
    # và dòng sau không bắt đầu bằng Câu, Bài, Đề, Ví dụ hoặc A., B., C., D.
    lines = text.split('\n')
    merged_lines = []

    # Regex nhận diện dấu kết thúc câu ở cuối dòng trước (đã loại bỏ ';' và ':' để không ngắt đôi tọa độ)
    sentence_end_pattern = re.compile(r'[\.\?!…]\s*$|(?:\$\$|\])\s*$|@@__PLACEHOLDER_IMG_TAG_\d+__@@\s*$')

    # Regex nhận diện dòng sau bắt đầu khối mới độc lập (chỉ nhận diện [a-d] thường, không nhận số 1-9 hay [B-D] viết hoa)
    # TUYỆT ĐỐI không dùng re.IGNORECASE cho PHẦN để tránh nhận nhầm 'phần trăm', 'phần mềm', 'từng phần'...
    new_block_pattern = re.compile(
        r'^(?:(?:Câu|Bài|Đề|Ví dụ)\s+(?:\$?\d+\$?)(?:\s*\([^\)\n]*\))?\s*[\.:]|(?:[IVXLCDM]+\.|\d+\.|\b[A-Z]\.)?\s*(?:PHẦN|Phần)\s+(?:[IVXLCDM]+|\d+|[A-D]\b|TRẮC|TỰ|MỘT|HAI|BA|BỐN|NĂM|THỨ)|[A-D]\.|\([a-d]\)|[a-d][\)\/]|\$\$|@@__PLACEHOLDER_IMG_TAG_\d+__@@|\[\s*(?:TIÊU[ _]ĐỀ|HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE))'
    )

    for line in lines:
        line_str = line.strip()
        if not line_str:
            merged_lines.append("")
            continue

        if not merged_lines:
            merged_lines.append(line_str)
            continue

        prev_line = merged_lines[-1].strip()

        # Bảo vệ tuyệt đối các dòng bảng Markdown (| Cột 1 | Cột 2 |) không bị ghép lẫn vào nhau
        if line_str.startswith('|') or prev_line.startswith('|'):
            merged_lines.append(line_str)
            continue

        # Nếu dòng trước rỗng
        if not prev_line:
            # Kiểm tra xem dòng hiện tại có phải là câu bị ngắt dòng ngang sang đoạn mới không (bắt đầu bằng chữ thường tiếng Việt)
            last_non_empty = ""
            for l in reversed(merged_lines[:-1]):
                if l.strip():
                    last_non_empty = l.strip()
                    break
            if last_non_empty and re.match(r'^[a-zà-ỹ]', line_str) and not sentence_end_pattern.search(last_non_empty):
                while merged_lines and not merged_lines[-1].strip():
                    merged_lines.pop()
                merged_lines[-1] = merged_lines[-1] + " " + line_str
                continue
            else:
                merged_lines[-1] = line_str
                continue

        # Kiểm tra xem dòng trước có kết thúc bằng dấu chấm câu không (hoặc là tiêu đề PHẦN)
        is_prev_sec = bool(re.search(r'^(?:[IVXLCDM]+\.|\d+\.|\b[A-Z]\.)?\s*(?:PHẦN|Phần)\s+(?:[IVXLCDM]+|\d+|[A-D]\b|TRẮC|TỰ|MỘT|HAI|THỨ)', prev_line))
        has_punct_end = bool(sentence_end_pattern.search(prev_line)) or is_prev_sec
        # Kiểm tra xem dòng sau có bắt đầu khối mới không
        is_new_block = bool(new_block_pattern.match(line_str))

        # Kiểm tra ngoặc chưa đóng, $ lẻ, hoặc từ nối ở cuối dòng trước
        open_paren = prev_line.count('(') > prev_line.count(')')
        open_dollar = (prev_line.count('$') % 2 == 1)
        ends_connector = bool(re.search(r'(?:[;,+\-=\/]|\bvà|\bhoặc|\blà|\bđồ\s+thị(?:\s+là)?|\bvectơ|\bvéc-tơ|\btại)\s*$', prev_line, re.IGNORECASE))

        # Nếu có ngoặc chưa đóng, $ lẻ hoặc từ nối, ép buộc KHÔNG xem là kết thúc câu hay khối mới
        if open_paren or open_dollar or ends_connector:
            has_punct_end = False
            is_new_block = False

        if not has_punct_end and not is_new_block:
            # Ghép dòng: kiểm tra nếu là đuôi từ dính liền (như c) sau chụ)
            if re.match(r'^(?:[cmnpt]|ng|ch|nh)[\)\]]', line_str) or (prev_line and prev_line[-1].isalpha() and re.match(r'^[a-zA-Zà-ỹÀ-Ỹ][\)\]]', line_str)):
                merged_lines[-1] = prev_line + line_str
            else:
                merged_lines[-1] = prev_line + " " + line_str
        else:
            merged_lines.append(line_str)

    res = '\n'.join(merged_lines)
    return unmask_placeholders(res, tag_map)


def clean_raw_text(text: str) -> str:
    """
    Làm sạch sâu văn bản thô:
    - Loại bỏ triệt để chân trang scan
    - Bảo vệ tuyệt đối các thẻ [HÌNH_VẼ...] và [BẢNG...] không bị biến dạng
    - Nối lại các từ bị ngắt dòng sai từ OCR
    - Xóa các ký tự ô vuông, ký tự ẩn không in được (như ô vuông trước dấu trừ trong [⬚-1; 3] hay (⬚-2; 0))
    - Gỡ bỏ định dạng **...** bao quanh cả câu hỏi (chỉ giữ in đậm cho từ khóa Câu 1:...)
    - Chuẩn hóa khoảng cách quanh công thức toán
    """
    if not text:
        return ""

    # Lọc bỏ chân trang scan
    text = remove_page_footers(text)

    # 0. Bảo vệ tuyệt đối các thẻ placeholder trước khi xử lý
    text, tag_map = mask_placeholders(text)

    # Nối lại từ bị ngắt dòng sai từ OCR
    text = join_broken_ocr_lines(text)

    # 1. Xóa ký tự điều khiển, ký tự ô vuông Unicode (Geometric Shapes, Dotted Squares, Box Drawing, Replacement)
    # Bao gồm \u25a0-\u25ff (□, ■), \u2b12-\u2b2f (⬚ \u2b1a), \u2500-\u257f, \ufffd, \u200b-\u200f (zero-width), \ufeff
    text = re.sub(r'[\u25a0-\u25ff\u2b12-\u2b2f\u2500-\u257f\ufffd\u200b-\u200f\ufeff\u202a-\u202e]', '', text)

    # 2. Xóa các khối includegraphics rác nếu còn sót
    text = re.sub(
        r'\\begin\{center\}\s*\\includegraphics\[.*?\]\{.*?\}\s*(?:%.*?)?\\end\{center\}',
        lambda m: '[HÌNH_VẼ: Đồ thị / Hình vẽ minh họa]',
        text,
        flags=re.DOTALL
    )
    text = re.sub(
        r'\\includegraphics(?:\[.*?\])?\{.*?\}',
        '[HÌNH_VẼ: Đồ thị / Hình vẽ minh họa]',
        text
    )

    # Chỉ xóa sạch các dòng chú thích ảnh thô ĐỨNG ĐỘC LẬP trên một dòng riêng biệt (tránh chém vào thẻ)
    text = re.sub(r'(?m)^\s*\|\s*(?:Đồ thị|Hình vẽ|Bảng biến thiên|Bảng số liệu|Hình ảnh|Minh họa|Sơ đồ|mô phỏng)[^\n]*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'(?m)^\s*\[(?:Đồ thị|Hình vẽ|Bảng biến thiên|Bảng số liệu|Hình ảnh|Minh họa|Sơ đồ)[^\]\n]*\]\s*$', '', text, flags=re.IGNORECASE)

    # 3. Gỡ bỏ dấu ** bao bọc cả câu hỏi nếu Gemini vô tình sinh ra (VD: **Câu 1: Cho hàm số...** -> Câu 1: Cho hàm số...)
    # Chỉ gỡ khi bắt đầu bằng Câu/Bài/Đề/Ví dụ và có dấu : hoặc . (không gỡ dấu ** của tiêu đề PHẦN)
    text = re.sub(r'\*\*\s*((?:Câu|Bài|Đề|Ví dụ)\s+(?:\$?\d+\$?)(?:\s*\([^\)\n]*\))?\s*[\.:])\s*(.*?)\*\*', r'\1 \2', text)

    # 4. Chuẩn hóa thẩm mỹ và khoảng cách dấu câu
    text = sanitize_text_aesthetics(text)

    # Khôi phục lại các thẻ placeholder nguyên vẹn
    text = unmask_placeholders(text, tag_map)
    return text


def auto_wrap_math_symbols(text: str) -> str:
    """
    Tự động chuẩn hóa và nhúng các ký hiệu toán học phổ biến vào equation ($...$):
    - Tên điểm: A, B, C, S, M, N, P, G, O...
    - Trục, hệ tọa độ, mặt phẳng: Ox, Oy, Oz, Oxyz, Oxy, (Oxy), (Oyz), (Oxz), (P), (Q)...
    - Tọa độ điểm / vectơ: G(1; 2; 3), A(0; 2), I(-2; 3), M(3; 3; -2), A'(0; 0; 4)...
    - Tọa độ bộ số đơn lập: (0; 2), (-1; -3; 5), (2; 3; 4)...
    - Biểu thức vectơ: \\vec{a} = (-1; 3; -2), \\vec{u} = \\vec{i} - 2\\vec{j} - \\vec{k}, \\vec{b} - \\vec{a}...
    - Biểu thức phân số: y = \\frac{3x+2}{x+2}, \\frac{a}{b}...
    - Đồ thị (C), (P), đường cong (C), đường tròn (C)...
    - Tập số: \\mathbb{R}, \\mathbb{N}...
    - Tách bỏ các từ tiếng Việt bị AI/OCR gom nhầm vào trong $...$.
    """
    if not text:
        return ""

    # 0. Gỡ bỏ cặp $ bao bọc toàn bộ câu tiếng Việt (do OCR gom nhầm)
    text = sanitize_text_aesthetics(text)
    text = unwrap_vietnamese_math(text)

    # Chuẩn hóa các biến thể vectơ ngoài văn bản và phân số thường
    text = sanitize_plain_text_vectors(text)
    text = re.sub(r"([yfx]\s*=\s*)?\(([0-9a-zA-Z\+\-\*]+)\)\s*/\s*\(([0-9a-zA-Z\+\-\*]+)\)",
                   lambda m: f"${(m.group(1) or '')}\\frac{{{m.group(2).strip()}}}{{{m.group(3).strip()}}}$", text)

    # 1. Hợp nhất triệt để các công thức toán bị gãy dòng ở giữa dấu $...$
    text = re.sub(r'\$([^$]+?)\$', lambda m: '$' + re.sub(r'\s*\n\s*', ' ', m.group(1)) + '$', text)

    # 1b. Hợp nhất các khối công thức toán học đứng kề nhau được nối bởi toán tử toán học (=, +, -, \cdot...)
    # Ví dụ: $\vec{a}$ = $(-1; 3; -2)$ -> $\vec{a} = (-1; 3; -2)$, $\vec{b}$ - $\vec{a}$ -> $\vec{b} - \vec{a}$
    for _ in range(3):
        prev = text
        text = re.sub(r'\$([^\$\n]+?)\$\s*([=+\-]|<|>|\\le\b|\\ge\b|\\cdot\b)\s*\$([^\$\n]+?)\$', r'$\1 \2 \3$', text)
        if text == prev:
            break

    # 2. Tokenize tất cả math block ($$...$$ và $...$) và placeholder có sẵn
    math_tokens = {}
    tok_counter = 0

    def mask_token(expr: str) -> str:
        nonlocal tok_counter
        clean_expr = expr.strip()
        # Chuẩn hóa khoảng cách sát dấu $ trong token
        clean_expr = re.sub(r'^\$\s+', '$', clean_expr)
        clean_expr = re.sub(r'\s+\$$', '$', clean_expr)
        clean_expr = re.sub(r'^\$\$\s+', '$$', clean_expr)
        clean_expr = re.sub(r'\s+\$\$$', '$$', clean_expr)
        tok = f"@@__TEMP_MATH_WRAP_{tok_counter}__@@"
        math_tokens[tok] = clean_expr
        tok_counter += 1
        return tok

    token_pattern = r'(\$\$[\s\S]*?\$\$|\$(?:[^\$\\]|\\.)*\$|\[\s*(?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE)\s*[:=]?\s*(?:[^\[\]]|\[(?:[^\[\]]|\[[^\]]*\])*\])*\])'
    text = re.sub(token_pattern, lambda m: mask_token(m.group(0)), text, flags=re.IGNORECASE)

    # 2b. Biểu thức biến có chỉ số dưới/trên và đẳng thức: x_{CĐ} = 4, x_{CT} = 0, y_{CĐ} = -1, S_{đáy} = a^2...
    sub_eq_pattern = r'(?<![\$\w])([a-zA-Z][\^_]\{[^{}]+\}\s*=\s*[-+]?[a-zA-Z0-9\.\/]+(?:\s*[,;]\s*[a-zA-Z][\^_]\{[^{}]+\}\s*=\s*[-+]?[a-zA-Z0-9\.\/]+)*)(?![\$\w])'
    text = re.sub(sub_eq_pattern, lambda m: mask_token(f"${m.group(1)}$"), text)

    # 2c. Biểu thức chỉ số dưới đơn lẻ: x_{CĐ}, x_{CT}, y_{CT}, S_{đáy}, V_{chóp}...
    sub_single_pattern = r'(?<![\$\w\\])([a-zA-Z][\^_]\{[^{}]+\})(?![\$\w])'
    text = re.sub(sub_single_pattern, lambda m: mask_token(f"${m.group(1)}$"), text)

    # 3. Biểu thức hàm số / phân số: y = \frac{3x+2}{x+2} hoặc \frac{a}{b}
    text = re.sub(r'((?:[yfx]\s*=\s*)?\\frac\{[^{}]+\}\{[^{}]+\})', lambda m: mask_token(f"${m.group(1)}$"), text)

    # 4. Biểu thức vectơ có tọa độ: \vec{a} = (-1; 3; -2), \vec{u} = (1; 2; 1), \vec{AC'} = (2; 3; 4)
    text = re.sub(r'(\\(?:vec|overrightarrow)\{[^}]+\}\s*=\s*\([-+]?[a-zA-Z0-9_\.]+(?:\s*[;,]\s*[-+]?[a-zA-Z0-9_\.]+){1,3}\))',
                   lambda m: mask_token(f"${m.group(1)}$"), text)

    # Biểu thức vectơ tuyến tính: \vec{u} = \vec{i} - 2\vec{j} - \vec{k}
    text = re.sub(r'(\\(?:vec|overrightarrow)\{[^}]+\}\s*=\s*\\(?:vec|overrightarrow)\{[^}]+\}(?:\s*[-+]\s*[-+]?[a-zA-Z0-9_\.]*\s*\\(?:vec|overrightarrow)\{[^}]+\})*)',
                   lambda m: mask_token(f"${m.group(1)}$"), text)

    # Hiệu / tổng / tích vectơ: \vec{b} - \vec{a}, \vec{AC'} . \vec{CD'}
    text = re.sub(r'(\\(?:vec|overrightarrow)\{[^}]+\}\s*(?:[\+\-\*]|\\cdot|\.)\s*\\(?:vec|overrightarrow)\{[^}]+\})',
                   lambda m: mask_token(f"${m.group(1)}$"), text)

    # Vectơ đơn lẻ còn sót: \vec{a}, \vec{b}, \vec{u}, \vec{AC'}
    text = re.sub(r'(\\(?:vec|overrightarrow)\{[^}]+\})',
                   lambda m: mask_token(f"${m.group(1)}$"), text)

    # 5. Tọa độ điểm dạng A(0; 2), I(-2; 3), A'(0; 0; 4), B(2; 0; 0)
    def wrap_coord(m):
        pt_name = m.group(1)
        coords = m.group(2)
        clean_coords = re.sub(r'\s*;\s*', '; ', re.sub(r'\s*,\s*', '; ', coords.strip()))
        return mask_token(f"${pt_name}({clean_coords})$")
    text = re.sub(r'\b([A-Z][\'\d_]*)\s*\(\s*([-+]?[a-zA-Z0-9_\.]+(?:\s*[;,]\s*[-+]?[a-zA-Z0-9_\.]+){1,3})\s*\)', wrap_coord, text)

    # 6. Tọa độ bộ số đơn lập có dấu chấm phẩy: (0; 2), (2; 3; 4), (-1; -3; 5)
    text = re.sub(r'(?<![A-Za-z0-9_\$])\(\s*([-+]?[a-zA-Z0-9_\.]+(?:\s*;\s*[-+]?[a-zA-Z0-9_\.]+)+)\s*\)',
                  lambda m: mask_token(f"$({re.sub(r'\s*;\s*', '; ', m.group(1).strip())})$"), text)

    # 7. Đồ thị (C), (P)...; có đồ thị là (C); của đồ thị (C)
    text = re.sub(r'(?i)\b((?:có\s+)?đồ\s+thị(?:\s+là)?|của\s+đồ\s+thị|đường\s+cong|đường\s+tròn|mặt\s+cầu|mặt\s+phẳng)\s+\(([A-Za-z0-9_]+)\)',
                  lambda m: f"{m.group(1)} {mask_token(f'$({m.group(2)})$')}", text)
    text = re.sub(r'(?i)\bđồ\s+thị\b([^;\.\?\!\n]{1,35})\(([A-Z])\)',
                  lambda m: f"đồ thị{m.group(1)}{mask_token(f'$({m.group(2)})$')}", text)

    # 8. Không gian Oxyz, Oxy, trục Ox, Oy, Oz, gốc O
    text = re.sub(r'(?i)\b(không\s+gian|hệ\s+trục\s+tọa\s+độ|hệ\s+tọa\s+độ)\s+(Oxyz|Oxy)\b',
                  lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}", text)
    text = re.sub(r'(?i)\b(gốc)\s+(O)\b', lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}", text)
    text = re.sub(r'(?i)\btrục\s+(Ox|Oy|Oz)\s*,\s*(Ox|Oy|Oz)(?:\s*(?:,|và)\s*(Ox|Oy|Oz))?\b',
                  lambda m: f"trục {mask_token(f'${m.group(1)}$')}, {mask_token(f'${m.group(2)}$')}" + (f" và {mask_token(f'${m.group(3)}$')}" if m.group(3) else ""), text)
    text = re.sub(r'(?i)\btrục\s+(Ox|Oy|Oz)\b', lambda m: f"trục {mask_token(f'${m.group(1)}$')}", text)
    text = re.sub(r'(?i)\bmặt\s+phẳng\s+\(?(Oxy|Oyz|Oxz)\)?', lambda m: f"mặt phẳng {mask_token(f'$({m.group(1)})$')}", text)
    text = re.sub(r'(?<![A-Za-z0-9_\$])(Oxyz|Oxy)(?![A-Za-z0-9_\$])', lambda m: mask_token(f"${m.group(1)}$"), text)

    # 9. Tập số: \mathbb{R}, \mathbb{N}... hoặc R đứng sau trên/thuộc/trong
    text = re.sub(r'\\mathbb\{([RNZQ])\}', lambda m: mask_token(f"$\\mathbb{{{m.group(1)}}}$"), text)
    text = re.sub(r'(?i)\b(trên|thuộc|trong)\s+([R])\b', lambda m: f"{m.group(1)} {mask_token(f'$\\mathbb{{{m.group(2)}}}$')}", text)

    # 10. Điểm hình học đơn lẻ
    text = re.sub(r'(?i)\b(hai\s+điểm)\s+([A-Z])\s+và\s+([A-Z])\b',
                  lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')} và {mask_token(f'${m.group(3)}$')}", text)
    text = re.sub(r'(?i)\b(ba\s+điểm)\s+([A-Z])\s*,\s*([A-Z])\s+và\s+([A-Z])\b',
                  lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}, {mask_token(f'${m.group(3)}$')} và {mask_token(f'${m.group(4)}$')}", text)
    text = re.sub(r'(?i)\b(các\s+điểm)\s+([A-Z])\s*,\s*([A-Z])\b',
                  lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}, {mask_token(f'${m.group(3)}$')}", text)
    text = re.sub(r'(?i)\b(điểm|đỉnh|tâm|trọng\s+tâm|trực\s+tâm)[ \t]+([A-Z])(?!\.\s*(?:[a-zA-Z0-9\$\-\+\(]|\\|\/))\b',
                  lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}", text)
    text = re.sub(r'(?i)\btam\s+giác\s+([A-Z\'][A-Z0-9_\']{2,8})(?![a-zA-Z\'])',
                  lambda m: f"tam giác {mask_token(f'${m.group(1)}$')}", text)
    text = re.sub(r'(?i)\bhình\s+hộp(?:\s+chữ\s+nhật)?\s+([A-Z\'][A-Z0-9_\']{2,8}(?:\.[A-Z\'][A-Z0-9_\']{2,8})?)(?![a-zA-Z\'])',
                  lambda m: f"{m.group(0)[:m.start(1)-m.start(0)]}{mask_token(f'${m.group(1)}$')}", text)
    text = re.sub(r'(?i)\bhình\s+chóp\s+([A-Z]\.[A-Z0-9_\']{3,8})(?![a-zA-Z\'])',
                  lambda m: f"hình chóp {mask_token(f'${m.group(1)}$')}", text)

    # 11. Các phép tính, tham số
    text = re.sub(r'(?i)\b(tham\s+số|ẩn\s+số|ẩn|nghiệm)\s+([a-z])\b', lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}", text)
    text = re.sub(r'(?i)\b(phép\s+tính)\s+(\([+\-]?\d+\)\s*[+\-\*\/:]\s*[-+]?\d+)', lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}", text)
    text = re.sub(r'(?i)\b(số\s+(?:nguyên|thực|hữu\s+tỉ)?)\s+([+\-]\d+)\b', lambda m: f"{m.group(1)} {mask_token(f'${m.group(2)}$')}", text)

    # Khôi phục tokens
    for tok, expr in math_tokens.items():
        text = text.replace(tok, expr)

    # Hợp nhất các khối công thức toán học đứng kề nhau được nối bởi toán tử toán học (=, +, -, \cdot...)
    for _ in range(3):
        prev = text
        text = re.sub(r'\$([^\$\n]+?)\$\s*([=+\-]|<|>|\\le\b|\\ge\b|\\cdot\b)\s*\$([^\$\n]+?)\$', r'$\1 \2 \3$', text)
        if text == prev:
            break

    # Dọn dẹp khoảng cách thừa trước dấu câu và chuẩn hóa thẩm mỹ
    text = sanitize_text_aesthetics(text)

    # Smart Spacing Bridge: Đảm bảo tuyệt đối không dính chữ giữa công thức toán $ và chữ cái tiếng Việt
    text = ensure_math_spacing_bridge(text)

    return text


def parse_inline_formatting(text_fragment: str) -> list[dict]:
    """Parse đoạn text tìm **in đậm** hoặc text thường."""
    results = []
    pattern = r'\*\*(.*?)\*\*'
    last_end = 0

    for match in re.finditer(pattern, text_fragment):
        start, end = match.span()
        if start > last_end:
            results.append({'text': text_fragment[last_end:start], 'bold': False})
        results.append({'text': match.group(1), 'bold': True})
        last_end = end

    if last_end < len(text_fragment):
        results.append({'text': text_fragment[last_end:], 'bold': False})

    return results if results else [{'text': text_fragment, 'bold': False}]


def parse_tag_content(content: str):
    """Trích xuất tọa độ box, tên file ảnh và mô tả từ nội dung bên trong thẻ (hỗ trợ cả có và không có question=, không phân biệt thứ tự)."""
    box = None
    if not content:
        return None, None, "", None, None

    # Làm sạch nếu content truyền vào chứa cả thẻ đầy đủ [HÌNH_VẼ: ...] hoặc HÌNH_VẼ: ...
    cleaned_content = content.strip()
    m_tag_wrap = re.search(r'^\s*\[?\s*(?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE)\s*[:=]?\s*(.*?)\s*\]?$', cleaned_content, re.IGNORECASE | re.DOTALL)
    if m_tag_wrap:
        cleaned_content = m_tag_wrap.group(1).strip()
        cleaned_content = re.sub(r'\]+$', '', cleaned_content).strip()

    desc = cleaned_content

    # 0. Trích xuất question=label nếu có và làm sạch metadata
    question_label = None
    m_q = re.search(r'question\s*[:=]\s*["\']([^"\']+)["\']', desc, re.IGNORECASE)
    if not m_q:
        m_q = re.search(r'question\s*[:=]\s*([^\s\|\]]+)', desc, re.IGNORECASE)
    if m_q:
        question_label = m_q.group(1).strip()
        desc = desc[:m_q.start()] + " " + desc[m_q.end():]

    # 1. Trích xuất image=filename hoặc filename ảnh độc lập (bất kể thứ tự trước hay sau box)
    image_name = None
    m_img = re.search(r'image\s*[:=]\s*["\']?([^\s\|\]\'"]+)["\']?', desc, re.IGNORECASE)
    if not m_img:
        m_img = re.search(r'image\s*[:=]\s*["\']?([^\|\]\'"]+?)["\']?\s*(?:\||\]|$)', desc, re.IGNORECASE)
    if m_img:
        raw_img = m_img.group(1).strip()
        desc = desc[:m_img.start()] + " " + desc[m_img.end():]
        raw_img = re.sub(r'question\s*[:=]\s*(?:"[^"]*"|\'[^\']*\'|[^\s\|\]]+)\s*\|*', '', raw_img, flags=re.IGNORECASE)
        m_file = re.search(r'((?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg)|[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg))', raw_img, re.IGNORECASE)
        if m_file:
            image_name = m_file.group(1)
        else:
            image_name = re.sub(r'^[\|\s,;:\'"]+|[\|\s,;:\'"]+$', '', raw_img).strip()

    # Quét thêm nếu chưa có image_name: tìm trực tiếp tên file ảnh dạng crop_xxx.png hoặc autocrop_xxx.png
    if not image_name:
        m_file_crop = re.search(r'\b((?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg))\b', desc, re.IGNORECASE)
        if m_file_crop:
            image_name = m_file_crop.group(1)
            desc = desc[:m_file_crop.start()] + " " + desc[m_file_crop.end():]

    # 2. Kiểm tra box=[ymin, xmin, ymax, xmax] hoặc box: [...]
    m_box = re.search(r'box\s*[:=]\s*\[?\s*(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)\s*\]?', desc, re.IGNORECASE)
    if m_box:
        box = [float(m_box.group(i)) for i in range(1, 5)]
        desc = desc[:m_box.start()] + " " + desc[m_box.end():]

    # 3. Kiểm tra [ymin, xmin, ymax, xmax] độc lập
    if not box:
        m_box2 = re.search(r'\[\s*(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)\s*\]', desc)
        if m_box2:
            box = [float(m_box2.group(i)) for i in range(1, 5)]
            desc = desc[:m_box2.start()] + " " + desc[m_box2.end():]

    # 4. Kiểm tra ymin=... xmin=...
    if not box:
        m_box3 = re.search(r'ymin\s*[:=]\s*(\d+(?:\.\d+)?).*?xmin\s*[:=]\s*(\d+(?:\.\d+)?).*?ymax\s*[:=]\s*(\d+(?:\.\d+)?).*?xmax\s*[:=]\s*(\d+(?:\.\d+)?)', desc, re.IGNORECASE)
        if m_box3:
            box = [float(m_box3.group(i)) for i in range(1, 5)]
            desc = desc[:m_box3.start()] + " " + desc[m_box3.end():]

    # 5. Nếu chưa tìm thấy image_name, quét trực tiếp tên file ảnh trong desc
    if not image_name:
        m_file_direct = re.search(r'\b((?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg)|[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg))\b', desc, re.IGNORECASE)
        if m_file_direct:
            image_name = m_file_direct.group(1)
            desc = desc[:m_file_direct.start()] + " " + desc[m_file_direct.end():]

    # Tách phương án A. B. C. D. nếu vô tình bị dính vào trong mô tả ảnh
    trailing_choice = None
    m_ch_inside = re.search(r'(?i)\b([A-D]\.\s*.*)$', desc)
    if m_ch_inside:
        trailing_choice = m_ch_inside.group(1).strip()
        desc = desc[:m_ch_inside.start()].strip()

    desc = re.sub(r'\|+', '|', desc)
    desc = re.sub(r'^[\|\s,;:]+|[\|\s,;:]+$', '', desc).strip()
    return box, image_name, desc, question_label, trailing_choice


def is_invalid_or_fake_figure(box: list[float] | None, desc: str | None) -> bool:
    """
    Kiểm tra và loại bỏ các thẻ hình vẽ rác / ảo giác do AI sinh ra:
    1. Chú thích chứa từ khóa báo hiệu không có hình (như: 'không có hình', 'scan đề', 'toàn trang'...)
    2. Box khoanh trùm toàn bộ trang giấy A4 (diện tích > 55% hoặc bao trùm gần hết trang).
    """
    desc_str = (desc or "").strip().lower()

    # 1. Kiểm tra từ khóa rác báo hiệu không có hình hoặc chụp scan toàn trang
    fake_keywords = [
        "không có hình", "không có đồ thị", "không có bảng", "không hình vẽ",
        "scan đề", "bản scan", "trang scan", "toàn trang", "toàn bộ trang",
        "ảnh chụp đề", "đề bài scan", "không có hình ảnh", "chỉ có chữ", "chỉ gồm chữ"
    ]
    for kw in fake_keywords:
        if kw in desc_str:
            return True

    # 2. Kiểm tra kích thước bounding box
    if box and len(box) == 4:
        ymin, xmin, ymax, xmax = [float(v) for v in box]
        if ymax <= 1.0 and xmax <= 1.0:
            ymin, xmin, ymax, xmax = ymin * 1000, xmin * 1000, ymax * 1000, xmax * 1000

        h = abs(ymax - ymin)
        w = abs(xmax - xmin)

        # Nếu box bao trùm gần như toàn bộ trang (trên 750x750) hoặc diện tích > 550,000 (55% trang)
        if (h >= 750 and w >= 750) or (h * w >= 550000):
            return True
        # Nếu box ôm sát cả 4 mép trang A4 (ymin <= 80 và ymax >= 880 và xmin <= 80 và xmax >= 880)
        if ymin <= 80 and ymax >= 880 and xmin <= 80 and xmax >= 880:
            return True

    return False


def remove_fake_figure_tags(text: str) -> str:
    """Xóa sạch các thẻ hình vẽ giả mạo / chụp toàn trang / không có hình khỏi văn bản."""
    if not text:
        return ""
    def _clean_fake(m):
        raw = m.group(0).strip()
        if not raw.endswith(']'):
            raw += ']'
        m_inner = re.search(r'^\s*\[\s*((?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE))\s*[:=]?\s*(.*?)\s*\]$', raw, re.IGNORECASE | re.DOTALL)
        if m_inner:
            content = m_inner.group(2).strip()
            box, _, desc, _, _ = parse_tag_content(content)
            if is_invalid_or_fake_figure(box, desc):
                return ""
        return m.group(0)
    return PLACEHOLDER_REGEX.sub(_clean_fake, text)


def parse_placeholder_tag(text: str) -> dict | None:
    """
    Nhận diện thẻ [HÌNH_VẼ: ...] hoặc [BẢNG: ...] với mọi biến thể cú pháp:
    - [HÌNH_VẼ: question="Câu 2" | image=crop_123.png | box=[ymin, xmin, ymax, xmax] | mô tả]
    - [HÌNH_VẼ: image=crop_123.png | box=[ymin, xmin, ymax, xmax] | mô tả]
    - [HÌNH_VẼ: image=crop_123.png | box=[ymin, xmin, ymax, xmax]] (hoặc thiếu ] ngoài cùng)
    - [HÌNH VẼ: box=[ymin, xmin, ymax, xmax] | mô tả]
    - [BẢNG BIẾN THIÊN: [ymin, xmin, ymax, xmax] | mô tả]
    - [ĐỒ THỊ: mô tả]
    """
    if not text:
        return None

    tag_pattern = r'\[\s*((?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE))\s*[:=]?\s*((?:[^\[\]]|\[(?:[^\[\]]|\[[^\]]*\])*\])*)\]'
    match = re.search(tag_pattern, text, re.IGNORECASE)
    if not match:
        tag_pattern_flexible = r'\[\s*((?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE))\s*[:=]?\s*(.*?)\s*\](?:\s*$|\s+(?=[A-D]\.|\n|$))'
        match = re.search(tag_pattern_flexible, text, re.IGNORECASE | re.DOTALL)
    if not match:
        # Fallback bao dung lỗi: bắt được thẻ kể cả khi thiếu dấu ] đóng ngoài cùng
        tag_pattern_tolerant = r'\[\s*((?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE))\s*[:=]?\s*([^\n\r]+?)(?:\]\s*$|\s*$)'
        match = re.search(tag_pattern_tolerant, text, re.IGNORECASE)
    if not match:
        return None

    raw_type = match.group(1).upper()
    subtype = 'table' if any(k in raw_type for k in ('BẢNG', 'TABLE')) else 'image'
    content = match.group(2).strip()

    box, image_name, desc, question_label, trailing_choice = parse_tag_content(content)
    if is_invalid_or_fake_figure(box, desc):
        return None

    if not desc:
        desc = 'Bảng biến thiên / Bảng số liệu' if subtype == 'table' else 'Hình vẽ / Đồ thị'

    return {
        'type': 'placeholder',
        'subtype': subtype,
        'image': image_name,
        'box': [int(x) for x in box] if box else None,
        'desc': desc,
        'question': question_label,
        'trailing_choice': trailing_choice
    }


def parse_latex_content(text: str) -> list[dict]:
    """Parse nội dung thành các segment (placeholder, display math, inline math, text_runs)."""
    if not text:
        return []

    text = remove_fake_figure_tags(text)
    # Loại bỏ thẻ tiêu đề đề thi nếu còn sót
    _, text = extract_exam_header(text)

    # 1. Kiểm tra placeholder trước tiên:
    # Nếu cả dòng là một thẻ placeholder (hoặc bắt đầu bằng placeholder)
    tag_info = parse_placeholder_tag(text)
    if tag_info and text.strip().startswith('['):
        segments = [tag_info]
        if tag_info.get('trailing_choice'):
            segments.extend(parse_latex_content(tag_info['trailing_choice']))
        return segments

    # Nếu thẻ placeholder nằm ở giữa hoặc có text xung quanh:
    m_tag = PLACEHOLDER_REGEX.search(text)
    if m_tag:
        tag_info = parse_placeholder_tag(m_tag.group(0))
        if tag_info:
            before_text = text[:m_tag.start()].strip()
            after_text = text[m_tag.end():].strip()
            # Xóa sạch chú thích ảnh thô ngay sau thẻ nếu có
            after_text = re.sub(r'^\s*\|[^\]\n]*\]?\s*', '', after_text, flags=re.IGNORECASE).strip()
            after_text = re.sub(r'^\s*\[(?:Đồ thị|Hình vẽ|Bảng biến thiên|Bảng số liệu|Hình ảnh|Minh họa|Sơ đồ)[^\]\n]*\]\s*', '', after_text, flags=re.IGNORECASE).strip()
            segments = []
            if before_text:
                segments.extend(parse_latex_content(before_text))
            segments.append(tag_info)
            if tag_info.get('trailing_choice'):
                segments.extend(parse_latex_content(tag_info['trailing_choice']))
            if after_text:
                segments.extend(parse_latex_content(after_text))
            return segments

    text = clean_raw_text(text)
    segments = []

    # 2. Xử lý Display math ($$...$$)
    parts = re.split(r'(\$\$[\s\S]*?\$\$)', text)
    for part in parts:
        if part.startswith('$$') and part.endswith('$$'):
            math_content = part[2:-2].strip()
            if math_content:
                segments.append({'type': 'display_math', 'content': math_content})
        else:
            # Inline math ($...$)
            inline_parts = re.split(r'(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)', part)
            for i, ip in enumerate(inline_parts):
                if i % 2 == 0:
                    if ip:
                        runs = parse_inline_formatting(ip)
                        segments.append({'type': 'text_runs', 'runs': runs})
                else:
                    if ip.strip():
                        # Kiểm tra nếu inline math bị dính nhãn A., B., C., D. ở đầu
                        m_opt = re.match(r'^\s*([A-D]\.)\s*(.*)$', ip.strip())
                        if m_opt:
                            segments.append({'type': 'text_runs', 'runs': [{'text': m_opt.group(1) + ' ', 'bold': True}]})
                            rest_math = m_opt.group(2).strip()
                            if rest_math:
                                segments.append({'type': 'inline_math', 'content': rest_math})
                        else:
                            segments.append({'type': 'inline_math', 'content': ip.strip()})

    return segments


def clean_choice_expr(expr: str) -> str:
    """Làm sạch biểu thức của phương án trắc nghiệm: gọt bỏ \\quad, \\qquad, khoảng trắng thừa."""
    if not expr:
        return ""
    s = expr.strip()
    s = re.sub(r'(?:\\quad|\\qquad|\\enspace|\\;|\\,|\\!|~|\s)+$', '', s).strip()
    s = re.sub(r'^(?:\\quad|\\qquad|\\enspace|\\;|\\,|\\!|~|\s)+', '', s).strip()
    return s


def normalize_choice_structure(text: str) -> str:
    """
    Chuẩn hóa toàn diện cấu trúc trắc nghiệm (A., B., C., D. hoặc a), b), c), d)):
    1. Bóc tách triệt để nếu các phương án trắc nghiệm bị Gemini bọc chung trong 1 khối math ($...$ hoặc $$...$$):
       Ví dụ: 'A. $1 \\in A B. 3 \\in A C. 5 \\notin A D. 9 \\in A$'
       hoặc '$A. 1 \\in A B. 3 \\in A C. 5 \\notin A D. 9 \\in A$'
       hoặc '$$A. 1 \\in A B. 3 \\in A C. 5 \\notin A D. 9 \\in A$$'
    2. Bóc tách nhãn phương án ($A.$, $B.$, $A. ...$, **A.**) ra ngoài Math/Bold.
    3. Tự động ngắt dòng \\n\\n trước mỗi phương án A., B., C., D. (kể cả khi nằm cùng dòng).
    4. Bảo vệ an toàn tuyệt đối các ký hiệu điểm hình học A, B, C, D trong câu hỏi.
    """
    if not text:
        return ""

    # Bước 1: Gỡ in đậm quanh nhãn phương án
    text = re.sub(r'\*\*\s*([A-D]\.)\s*\*\*', r'\1 ', text)
    text = re.sub(r'\*\*\s*([A-D])\s*\*\*\.', r'\1. ', text)
    text = re.sub(r'\*\*\s*([A-D])\.\s*\*\*', r'\1. ', text)

    # Bước 2: Tokenize Math blocks và Plain text để xử lý an toàn
    token_pattern = r'(\$\$[\s\S]*?\$\$|\$(?:[^\$\\]|\\.)*\$)'
    parts = re.split(token_pattern, text)

    new_parts = []
    for part in parts:
        if not part:
            continue

        # Display Math ($$...$$) hoặc Inline Math ($...$)
        is_display = part.startswith('$$') and part.endswith('$$') and len(part) >= 4
        is_inline = part.startswith('$') and part.endswith('$') and len(part) >= 2 and not is_display

        if is_display or is_inline:
            inner = part[2:-2].strip() if is_display else part[1:-1].strip()

            # Chuẩn hóa \\text{A.}, \\mathbf{A.}, \\rm{A.} thành A. bên trong math
            inner = re.sub(r'\\(?:text|mathbf|rm|textbf)\s*\{\s*([A-D]\.)\s*\}', r'\1', inner)

            # Kiểm tra xem bên trong math có chứa các nhãn trắc nghiệm không
            has_choices = bool(re.search(r'(?:^|\s+|\\quad|\\qquad|~)[B-D]\.\s+', inner)) or bool(re.search(r'^\s*A\.\s+', inner) and re.search(r'[B-D]\.', inner))

            if has_choices:
                # Tách theo nhãn phương án
                sub_opts = re.split(r'(?=(?:^|\s+|\\quad|\\qquad|~)[A-D]\.\s+)', inner)
                expanded = []
                for s_opt in sub_opts:
                    s_opt = s_opt.strip()
                    s_opt = re.sub(r'^(?:\\quad|\\qquad|\\enspace|~|\s)+', '', s_opt).strip()
                    if not s_opt:
                        continue
                    m_lbl = re.match(r'^([A-D]\.)\s*(.*)$', s_opt, re.DOTALL)
                    if m_lbl:
                        lbl = m_lbl.group(1)
                        expr = clean_choice_expr(m_lbl.group(2))
                        if expr:
                            expanded.append(f"\n\n{lbl} ${expr}$")
                        else:
                            expanded.append(f"\n\n{lbl}")
                    else:
                        expr = clean_choice_expr(s_opt)
                        if expr:
                            expanded.append(f"${expr}$")
                new_parts.append(" ".join(expanded))
            elif re.match(r'^\s*([A-D]\.)\s*(.*)$', inner, re.DOTALL):
                # Chỉ có 1 nhãn đơn lẻ bị bọc trong math: $A. 2x+1$ hoặc $$A. 2x+1$$
                m_single = re.match(r'^\s*([A-D]\.)\s*(.*)$', inner, re.DOTALL)
                lbl = m_single.group(1)
                expr = clean_choice_expr(m_single.group(2))
                delim = "$$" if is_display else "$"
                if expr:
                    new_parts.append(f"\n\n{lbl} {delim}{expr}{delim}\n\n" if is_display else f"\n\n{lbl} {delim}{expr}{delim}")
                else:
                    new_parts.append(f"\n\n{lbl} ")
            else:
                new_parts.append(part)
        else:
            p_text = part
            def repl_choice_stuck(m):
                prefix = m.group(1)
                lbl = m.group(2)
                start = m.start()
                before = p_text[max(0, start - 25):start].lower()
                # TUYỆT ĐỐI không tách nếu từ trước là đồ thị, đồ thị là, của đồ thị
                if re.search(r'(?:đồ\s+thị(?:\s+là)?|của\s+đồ\s+thị)\s*$', before):
                    return m.group(0)
                return f"{prefix}\n\n{lbl}. "

            p_text = re.sub(
                r'(\b(?:tại|là|bằng))\s*\$([A-D])\$\.\s+(?=(?:(?!Câu\s+\d+)[\s\S])*?\b[B-D]\.\s+)',
                repl_choice_stuck,
                p_text,
                flags=re.IGNORECASE
            )
            new_parts.append(p_text)

    text = "".join(new_parts)

    # 🎯 Tách dòng độc lập cho Khối Lời giải / Hướng dẫn giải / Chọn [A-D] nếu bị dính chùm ngay sau phương án trắc nghiệm
    sol_inline_pattern = re.compile(
        r'(\b[A-D]\.\s+[^\n]+?)\s*(\b(?:Lời\s+giải|Hướng\s+dẫn\s+giải|Giải\s+chi\s+tiết|HƯỚNG\s+DẪN\s+GIẢI|LỜI\s+GIẢI|GIẢI\s+CHI\s+TIẾT|HDG)\s*[:\.\-]?|\b(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b)',
        re.IGNORECASE
    )
    text = sol_inline_pattern.sub(r'\1\n\n\2', text)

    # Bước 3: Tách dòng cho các phương án trắc nghiệm trong text thường
    choice_pattern = re.compile(
        r'(?<!\n)(?:(?<=\.)|(?<=\?)|(?<=\:)|(?<=[a-zA-Zà-ỹÀ-Ỹ0-9\$\)\]]))\s+(\bA\.\s+)(?=(?:(?!Câu\s+\d+)[\s\S])*?\bB\.\s+(?:(?!Câu\s+\d+)[\s\S])*?\bC\.\s+(?:(?!Câu\s+\d+)[\s\S])*?\bD\.\s+)',
        re.IGNORECASE
    )
    text = choice_pattern.sub(r'\n\n\1', text)

    # Tách dòng trước B., C., D. khi đã có A. (bảo vệ tuyệt đối không tách nhầm Chọn D., Đáp án D.)
    def _split_choice_bcd(m):
        full = m.group(0)
        lbl = m.group(1)
        start = m.start()
        prefix = text[max(0, start - 20):start].lower()
        if re.search(r'\b(?:chọn|đáp\s*án|phương\s*án)\s*[:\.]?\s*$', prefix):
            return full
        return f"\n\n{lbl}"

    text = re.sub(r'(?<!\n)\s*(\b[B-D]\.\s+)', _split_choice_bcd, text)

    # Đảm bảo ngắt dòng cho bất kỳ A. nào đứng độc lập nếu theo sau có B. (bảo vệ Chọn A., Đáp án A.)
    def _split_choice_a(m):
        full = m.group(0)
        lbl = m.group(1)
        start = m.start()
        prefix = text[max(0, start - 20):start].lower()
        if re.search(r'\b(?:chọn|đáp\s*án|phương\s*án)\s*[:\.]?\s*$', prefix):
            return full
        return f"\n\n{lbl}"

    text = re.sub(r'(?<!\n)\s*(\bA\.\s+)(?=(?:(?!Câu\s+\d+)[\s\S])*?\bB\.\s+)', _split_choice_a, text)

    # Dọn dẹp $ rỗng do việc ghép sát label
    text = re.sub(r'(\b[A-D]\.)\s*\$\s*\$', r'\1 ', text)
    text = re.sub(r'(\b[A-D]\.)\s*\$\s*([A-D]\.)', r'\1\n\n\2', text)

    # Tách lại lần 2 nếu có cụm Lời giải / Chọn [A-D] còn sót
    text = sol_inline_pattern.sub(r'\1\n\n\2', text)

    return text


def split_into_paragraphs(text: str) -> list[str]:
    """
    Tách văn bản thành các đoạn văn độc lập một cách thông minh:
    - Gỡ bỏ $...$ ôm câu tiếng Việt (tránh dính chữ điquađiểm)
    - Hàn gắn tọa độ và công thức toán bị ngắt dòng
    - Bảo vệ tuyệt đối thẻ [HÌNH_VẼ...] và [BẢNG...] bằng MASK toàn bộ chu trình
    - Chuẩn hóa cấu trúc trắc nghiệm và bóc tách A., B., C., D. ra ngoài Equation
    - Bắt buộc chuẩn hóa nhúng ký hiệu toán học A, B, C, Ox, Oy, G(1,2,3)... vào equation
    - Bắt buộc tách đoạn mới trước mỗi tiêu đề câu hỏi: Câu 1., Câu 2., Bài 1...
    - Bắt buộc tách đoạn mới trước mỗi phương án trắc nghiệm: A., B., C., D. hoặc a), b), c), d)
    - Tách đoạn trước các thẻ placeholder [HÌNH_VẼ...] và [BẢNG...]
    """
    if not text:
        return []

    # 0. Gỡ bỏ dấu $...$ ôm cả câu tiếng Việt & hàn gắn toán / tọa độ gãy dòng trước
    text = unwrap_vietnamese_math(text)
    text = heal_broken_math_and_coords(text)
    text = remove_fake_figure_tags(text)
    # Loại bỏ thẻ tiêu đề đề thi nếu còn sót
    _, text = extract_exam_header(text)

    # 1. Chuẩn hóa & làm sạch metadata question="..." bên trong các thẻ hình vẽ/bảng
    def clean_tag_meta(m):
        raw = m.group(0).strip()
        if not raw.endswith(']'):
            raw += ']'
        m_inner = re.search(r'^\s*\[\s*((?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE))\s*[:=]?\s*(.*?)\s*\]$', raw, re.IGNORECASE | re.DOTALL)
        if m_inner:
            tag_type = m_inner.group(1).upper()
            tag_body = m_inner.group(2).strip()
            # Gỡ bỏ question=... để không xung đột với regex nhận diện Câu 1, Câu 2
            tag_body = re.sub(r'question\s*[:=]\s*(?:"[^"]*"|\'[^\']*\'|[^\s\|\]]+)\s*\|*', '', tag_body, flags=re.IGNORECASE)
            tag_body = re.sub(r'\|+', '|', tag_body).strip()
            tag_body = re.sub(r'^[\|\s,;:]+|[\|\s,;:]+$', '', tag_body).strip()
            return f"[{tag_type}: {tag_body}]"
        return raw

    text = PLACEHOLDER_REGEX.sub(clean_tag_meta, text)

    # 🛡️ 2. BẢO VỆ TUYỆT ĐỐI CÁC THẺ PLACEHOLDER:
    # MASK thẻ hình vẽ/bảng ngay từ đầu và GIỮ MASK trong suốt toàn bộ quá trình tách dòng/trắc nghiệm/câu hỏi/toán học
    text, tag_map = mask_placeholders(text)

    # Nối lại từ bị ngắt dòng sai từ OCR trước khi tách đoạn
    text = join_broken_ocr_lines(text)

    # Chuẩn hóa nhãn câu hỏi (gỡ $) và tách điểm số tự luận ra ngoài equation
    text = clean_exam_text(text)

    # Chuẩn hóa cấu trúc trắc nghiệm trước khi xử lý ký hiệu toán
    text = normalize_choice_structure(text)

    # Bọc ký hiệu toán học
    text = auto_wrap_math_symbols(text)

    # Dọn dẹp ký tự thừa
    text = clean_raw_text(text)

    # Đảm bảo mỗi token placeholder đứng trên dòng riêng biệt
    text = re.sub(r'\s*(@@__PLACEHOLDER_IMG_TAG_\d+__@@)\s*', r'\n\n\1\n\n', text)
    # Tách nếu sau placeholder có đáp án A-D
    text = re.sub(r'(@@__PLACEHOLDER_IMG_TAG_\d+__@@)\s*([A-D]\.)', r'\1\n\n\2', text)

    # 🛡️ 3. MASKING MATH BLOCKS FIRST để bảo vệ tuyệt đối tọa độ Oxy/Oxyz và công thức toán học
    # không bị bất kỳ regex tách ý tự luận hay ngắt dòng nào làm đứt gãy
    text, math_map = mask_math_blocks(text)

    # Tách dòng trước mỗi tiêu đề Câu / Bài / Đề / Ví dụ
    # Bắt buộc: Câu/Bài viết hoa, theo sau số là :, . hoặc điểm số trong ngoặc.
    # Luôn ngắt paragraph độc lập trước mỗi Câu \d+: hoặc Bài \d+:, bất kể câu trước kết thúc bằng gì.
    # TUYỆT ĐỐI không ngắt dòng khi trước nó là "từ câu", "đến câu", "ở mỗi câu", "trong câu"...
    q_split_pattern = re.compile(r'(\b(?:Câu|Bài|Đề|Ví dụ)\s+\d+(?:\s*\([^\)\n]*\))?\s*[\.:])')
    def _split_q_repl(m):
        start = m.start()
        prefix = text[max(0, start - 30):start].lower()
        if re.search(r'(?:từ|đến|ở|mỗi|các|trong|theo)\s+$', prefix):
            return m.group(0)
        if start > 0 and text[start - 1] == '\n':
            return m.group(0)
        return '\n\n' + m.group(0)
    text = q_split_pattern.sub(_split_q_repl, text)

    # Tách dòng trước các tiêu đề PHẦN (PHẦN I, PHẦN 1, I. PHẦN, PHẦN CÂU TRẮC NGHIỆM, PHẦN TỰ LUẬN...)
    # Giữ nguyên toàn bộ câu chỉ dẫn của tiêu đề phần trên 1 dòng hoàn chỉnh cho đến hết dòng
    # TUYỆT ĐỐI không dùng re.IGNORECASE để tránh nhận nhầm các từ 'phần trăm', 'phần mềm', 'từng phần'...
    sec_pattern = re.compile(
        r'(?<!\n)\s*(\b(?:[IVXLCDM]+\.|\d+\.|\b[A-Z]\.)?\s*(?:PHẦN|Phần)\s+(?:[IVXLCDM]+|\d+|[A-D]\b|TRẮC|TỰ|MỘT|HAI|BA|BỐN|NĂM|THỨ)[^\n]*)'
    )
    # Masking các tiêu đề PHẦN để bảo vệ tuyệt đối không bị các regex sau (trắc nghiệm, ý tự luận a/b/c/d) cắt ngang
    sec_map = {}
    sec_counter = 0
    def _mask_sec_repl(m):
        nonlocal sec_counter
        start = m.start()
        prefix = text[max(0, start - 20):start].lower()
        if re.search(r'(?:hàng|đến|ở|mỗi|các|trong|từng|với|về)\s+$', prefix):
            return m.group(0)
        tok = f"@@__SECTION_HEADER_{sec_counter}__@@"
        sec_map[tok] = m.group(1).strip()
        sec_counter += 1
        return f"\n\n{tok}\n\n"

    text = sec_pattern.sub(_mask_sec_repl, text)

    # Chuẩn hóa & tách dòng trước mỗi phương án trắc nghiệm A., B., C., D.
    text = normalize_choice_structure(text)
    text = re.sub(r'(?<!\n)\s*(\b[A-D]\.\s*)', r'\n\n\1', text)

    # TÁCH DÒNG Ý TỰ LUẬN & CÂU ĐÚNG/SAI a), b), c), d)
    # - Tách ý đầu tiên a) khi đi sau dấu kết thúc câu hoặc điểm số
    text = re.sub(r'([\.\?!:;\)]|\bđiểm\))\s+([a-d][\)\/]|\([a-d]\))\s+', r'\1\n\n\2 ', text)
    # - Tách ý a) đứng độc lập nếu theo sau có b)
    text = re.sub(r'(?<!\n)\s*([a-d][\)\/]|\([a-d]\))\s+(?=(?:(?!Câu|Bài)[\s\S])*?[b-d][\)\/])', r'\n\n\1 ', text)

    # - Tách các ý tiếp theo b) c) d) hoặc (b) (c) (d):
    # CHỈ khớp chữ thường [b-d] (tuyệt đối không khớp (C) của đồ thị hay B-D trắc nghiệm).
    # TUYỆT ĐỐI không khớp nếu phía trước có dấu phẩy ',' (như a), b), c), d) trong câu chỉ dẫn)
    # hoặc chữ 'ý', 'các ý', 'mỗi ý', 'trong ý', 'và', 'hoặc'.
    def _split_choice_bcd_repl(m):
        full = m.group(0)
        start = m.start()
        prefix = text[max(0, start - 25):start].lower()
        if re.search(r'[,;:\-\+]\s*$', prefix) or re.search(r'\b(?:ý|các\s+ý|mỗi\s+ý|trong\s+ý|và|hoặc|đồ\s+thị|đường\s+cong|đường\s+tròn)\s*$', prefix):
            return full
        lbl = m.group(1)
        return f"\n\n{lbl} "

    bcd_pattern = re.compile(r'(?<!\n)\s*([b-d][\)\/]|\([b-d]\))\s+')
    text = bcd_pattern.sub(_split_choice_bcd_repl, text)

    # - Tách ý đánh số 1), 2), 3) hoặc 1/, 2/, 3/ khi đi sau dấu kết thúc câu (loại trừ ';')
    def _split_num_item_repl(m):
        full = m.group(0)
        start = m.start()
        prefix = text[max(0, start - 30):start]
        # Nếu có dấu mở ngoặc '(' chưa đóng trong phạm vi 30 ký tự -> đây là tọa độ, KHÔNG TÁCH
        if '(' in prefix and ')' not in prefix[prefix.rfind('('):]:
            return full
        if re.search(r'[,;]\s*$', prefix):
            return full
        punct = m.group(1)
        item_lbl = m.group(2)
        return f"{punct}\n\n{item_lbl} "

    text = re.sub(r'([\.\?!:]|\bđiểm\))\s+([1-9][\)\/])\s+(?=[A-ZÀ-Ỹa-zà-ỹ0-9\$\@])', _split_num_item_repl, text)

    # 🎯 TÁCH DÒNG ĐỘC LẬP CHO KHỐI LỜI GIẢI / HƯỚNG DẪN GIẢI / CHỌN [A-D]
    sol_block_pattern = re.compile(
        r'(?<!\n)\s*(\b(?:Lời\s+giải|Hướng\s+dẫn\s+giải|Giải\s+chi\s+tiết|HƯỚNG\s+DẪN\s+GIẢI|LỜI\s+GIẢI|GIẢI\s+CHI\s+TIẾT|HDG)\s*[:\.\-]?)',
        re.IGNORECASE
    )
    text = sol_block_pattern.sub(r'\n\n\1', text)

    sol_choice_pattern = re.compile(
        r'(?<!\n)(?:(?<=[\.\?!:;\n])|\s{2,})\s*(\b(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b(?!\s*để|\s*cho))',
        re.IGNORECASE
    )
    text = sol_choice_pattern.sub(r'\n\n\1', text)

    # Khôi phục lại các tiêu đề PHẦN nguyên vẹn
    for tok, orig in sec_map.items():
        text = text.replace(tok, orig)

    # 🛡️ Khôi phục lại toàn bộ khối Math nguyên vẹn
    text = unmask_math_blocks(text, math_map)

    # 🛡️ Khôi phục lại các thẻ placeholder sau khi toàn bộ ngắt dòng trắc nghiệm đã hoàn tất
    text = unmask_placeholders(text, tag_map)
    text = sanitize_text_aesthetics(text)

    # 6. Phân tách theo dòng và gom nhóm hợp lý
    raw_lines = text.split('\n')
    paragraphs = []
    current_acc = []

    for line in raw_lines:
        line_str = line.strip()
        if not line_str:
            continue

        # Bỏ qua hoàn toàn các dòng chú thích ảnh thô kiểu | Đồ thị hàm số...] hoặc [Đồ thị...] hoặc | mô phỏng...]
        if re.match(r'^\s*\|[^\]\n]*\]?\s*$', line_str):
            continue
        if re.match(r'^(?:\|\s*|\[)?\s*(?:Đồ thị|Hình vẽ|Bảng biến thiên|Bảng số liệu|Hình ảnh|Minh họa|Sơ đồ|mô phỏng)\b[^\]\n]*\]?$', line_str, re.IGNORECASE):
            continue

        is_placeholder = bool(re.match(r'^\[\s*(?:HÌNH|BẢNG|ĐỒ|IMAGE|TABLE|FIGURE)', line_str, re.IGNORECASE))
        is_solution_block = bool(
            re.match(r'^(?:Lời\s+giải|Hướng\s+dẫn\s+giải|Giải\s+chi\s+tiết|HƯỚNG\s+DẪN\s+GIẢI|LỜI\s+GIẢI|GIẢI\s+CHI\s+TIẾT|HDG)\b', line_str, re.IGNORECASE) or
            re.match(r'^(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b', line_str, re.IGNORECASE)
        )
        is_new_block = (
            re.match(r'^(?:[IVXLCDM]+\.|\d+\.|\b[A-Z]\.)?\s*(?:PHẦN|Phần)\s+(?:[IVXLCDM]+|\d+|[A-D]\b|TRẮC|TỰ|MỘT|HAI|THỨ)', line_str) or
            re.match(r'^(?:Câu|Bài|Đề|Ví dụ)\s+(?:\$?\d+\$?)(?:\s*\([^\)\n]*\))?\s*[\.:]', line_str) or
            re.match(r'^[A-D]\.\s*', line_str) or
            re.match(r'^(?:[a-d][\)\/]|\([a-d]\))\s+', line_str) or
            is_placeholder or
            is_solution_block or
            line_str.startswith('$$')
        )

        if is_placeholder:
            # Placeholder luôn đứng riêng 1 paragraph, không cho dòng sau gộp vào
            if current_acc:
                paragraphs.append(' '.join(current_acc))
                current_acc = []
            paragraphs.append(line_str)
            continue

        if is_new_block:
            if current_acc:
                paragraphs.append(' '.join(current_acc))
            current_acc = [line_str]
        else:
            if current_acc:
                current_acc.append(line_str)
            else:
                current_acc = [line_str]

    if current_acc:
        paragraphs.append(' '.join(current_acc))

    return [p for p in paragraphs if p.strip()]


# ═══════════════════════════════════════════
# WORD DOCUMENT BUILDER
# ═══════════════════════════════════════════

def add_math_to_paragraph(paragraph, latex_str: str, is_display: bool = False):
    """Thêm công thức OMML vào paragraph của Word."""
    try:
        omml_element = latex_to_omml(latex_str)

        omath_elements = omml_element.findall('.//' + qn('m:oMath'))
        if not omath_elements:
            omath_elements = omml_element.findall(qn('m:oMath'))

        if not omath_elements:
            if omml_element.tag == qn('m:oMath') or 'oMath' in omml_element.tag:
                omath_elements = [omml_element]

        if omath_elements:
            for omath in omath_elements:
                paragraph._element.append(copy.deepcopy(omath))
        else:
            run = paragraph.add_run(f' {latex_str} ')
            run.font.italic = True

    except Exception:
        run = paragraph.add_run(f' {latex_str} ')
        run.font.italic = True


def add_text_and_math_runs(paragraph, text: str, italic: bool = False, bold: bool = False, font_size=Pt(12), color=None, math_format: str = "omml"):
    """
    Thêm đoạn văn bản vào paragraph, tự động nhận diện nếu có $...$
    để chèn công thức OMML chuẩn hoặc giữ định dạng LaTeX text ($...$).
    """
    is_latex_mode = (str(math_format).lower() in ("latex", "raw", "text"))
    parts = re.split(r'(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)', text)
    for i, part in enumerate(parts):
        if not part:
            continue
        if i % 2 == 1:
            # Inline math
            if is_latex_mode:
                run = paragraph.add_run(f"${part.strip()}$")
                run.italic = italic
                run.bold = bold
                if font_size:
                    run.font.size = font_size
                if color:
                    run.font.color.rgb = color
            else:
                add_math_to_paragraph(paragraph, part.strip(), is_display=False)
        else:
            run = paragraph.add_run(part)
            run.italic = italic
            run.bold = bold
            if font_size:
                run.font.size = font_size
            if color:
                run.font.color.rgb = color


# ═══════════════════════════════════════════
# EXAM HEADER BUILDER (CHUẨN BỘ GD&ĐT / NGHỊ ĐỊNH 30)
# ═══════════════════════════════════════════

def set_table_borders_none(table):
    """Xóa sạch mọi đường viền của bảng (tạo bảng không viền - borderless table)."""
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        for b in tblPr[0].xpath('w:tblBorders'):
            tblPr[0].remove(b)
        borders = OxmlElement('w:tblBorders')
        for b_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            b = OxmlElement(f'w:{b_name}')
            b.set(qn('w:val'), 'none')
            borders.append(b)
        tblPr[0].append(borders)


def set_cell_box_border(cell):
    """Tạo đường viền hộp đơn (box border) bao quanh ô (dùng cho Mã đề thi)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for b in tcPr.xpath('w:tcBorders'):
        tcPr.remove(b)
    tcBorders = OxmlElement('w:tcBorders')
    for b_name in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{b_name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '8')  # 1pt
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        tcBorders.append(b)
    tcPr.append(tcBorders)


def parse_tagged_exam_header(text: str) -> tuple[dict | None, str]:
    """Trích xuất khối tiêu đề đề thi nếu được bọc trong thẻ [TIÊU_ĐỀ_ĐỀ_THI: ...]."""
    if not text:
        return None, text

    m_tag = re.search(r'\[\s*(?:TIÊU[ _]ĐỀ(?:[ _]ĐỀ[ _]THI)?|HEADER(?:[ _]EXAM)?)\s*:\s*(.*?)\s*\]', text, re.IGNORECASE | re.DOTALL)
    if not m_tag:
        return None, text

    content = m_tag.group(1).strip()
    rem_text = (text[:m_tag.start()] + text[m_tag.end():]).strip()

    left_m = re.search(r'(?:BÊN[ _]TRÁI|LEFT)\s*:\s*(.*?)(?=(?:BÊN[ _]PHẢI|RIGHT|THÔNG[ _]TIN|INFO)\s*:|$)', content, re.IGNORECASE | re.DOTALL)
    right_m = re.search(r'(?:BÊN[ _]PHẢI|RIGHT)\s*:\s*(.*?)(?=(?:BÊN[ _]TRÁI|LEFT|THÔNG[ _]TIN|INFO)\s*:|$)', content, re.IGNORECASE | re.DOTALL)
    info_m = re.search(r'(?:THÔNG[ _]TIN|INFO)\s*:\s*(.*?)(?=(?:BÊN[ _]TRÁI|LEFT|BÊN[ _]PHẢI|RIGHT)\s*:|$)', content, re.IGNORECASE | re.DOTALL)

    left_lines = [l.strip() for l in left_m.group(1).split('\n') if l.strip()] if left_m else []
    right_lines = [l.strip() for l in right_m.group(1).split('\n') if l.strip()] if right_m else []
    info_str = info_m.group(1).strip() if info_m else ''

    student_info = info_str
    exam_code = ''
    if '|' in info_str:
        parts = [p.strip() for p in info_str.split('|')]
        s_parts = []
        for p in parts:
            m_c = re.search(r'\b(Mã\s+đề(?:\s+thi)?\s*[:=]?\s*[\w\d]+)', p, re.IGNORECASE)
            if m_c and not exam_code:
                exam_code = m_c.group(1).strip()
                rem = (p[:m_c.start()] + p[m_c.end():]).strip(' ,;:')
                if rem:
                    s_parts.append(rem)
            else:
                s_parts.append(p)
        student_info = ', '.join(s_parts)
    else:
        m_c = re.search(r'\b(Mã\s+đề(?:\s+thi)?\s*[:=]?\s*[\w\d]+)', info_str, re.IGNORECASE)
        if m_c:
            exam_code = m_c.group(1).strip()
            student_info = (info_str[:m_c.start()] + info_str[m_c.end():]).strip(' ,;:')

    return {
        'left': left_lines,
        'right': right_lines,
        'student_info': student_info,
        'exam_code': exam_code
    }, rem_text


def parse_fallback_exam_header(text: str) -> tuple[dict | None, str]:
    """
    Cơ chế nhận diện thông minh (Fallback):
    Nếu văn bản bắt đầu bằng các dòng Sở GD&ĐT, Bộ GD&ĐT, Trường THPT... dù AI chưa bọc thẻ,
    hệ thống vẫn tự động trích xuất và dựng bảng 2 cột chuẩn.
    """
    if not text:
        return None, text

    lines = text.strip().split('\n')
    left_pat = re.compile(r'^(?:SỞ|BỘ|PHÒNG)\s+GD|TRƯỜNG\s+THPT|TRƯỜNG\s+THCS|TRƯỜNG\s+PHỔ\s+THÔNG|\(Đề(?:\s+thi)?\s+(?:gồm|có)|^\(?Đề\s+chính\s+thức\)?', re.IGNORECASE)
    right_pat = re.compile(r'ĐỀ\s+KIỂM\s+TRA|ĐỀ\s+THI|KỲ\s+THI|KIỂM\s+TRA\s+(?:CUỐI|GIỮA|HỌC|ĐỊNH)|MÔN\s*:|BÀI\s+THI\s*:|Thời\s+gian\s+làm\s+bài', re.IGNORECASE)
    info_pat = re.compile(r'Họ\s+(?:và\s+)?tên|Số\s+báo\s+danh|Lớp\s*:|Mã\s+đề', re.IGNORECASE)
    stop_pat = re.compile(r'^(?:[IVXLCDM]+\.|\d+\.|\b[A-Z]\.)?\s*(?:PHẦN|Phần)\b|^(?:Câu|Bài|Ví dụ)\s+(?:\$?\d+\$?)(?:\s*\([^\)\n]*\))?\s*[\.:]', re.IGNORECASE)

    left_lines = []
    right_lines = []
    info_lines = []

    idx = 0
    while idx < len(lines) and idx < 12:
        l = lines[idx].strip()
        if not l:
            idx += 1
            continue
        if stop_pat.search(l):
            break

        if info_pat.search(l):
            info_lines.append(l)
            idx += 1
            continue
        elif left_pat.search(l):
            left_lines.append(l)
            idx += 1
            continue
        elif right_pat.search(l):
            right_lines.append(l)
            idx += 1
            continue
        else:
            if left_lines or right_lines:
                if 'trang' in l.lower():
                    left_lines.append(l)
                elif 'thời gian' in l.lower() or 'phút' in l.lower() or 'môn' in l.lower():
                    right_lines.append(l)
                else:
                    break
            else:
                break
            idx += 1

    if (len(left_lines) + len(right_lines) >= 2) or (left_lines and right_lines):
        rest_text = '\n'.join(lines[idx:]).strip()
        info_str = ' | '.join(info_lines)
        student_info = info_str
        exam_code = ''
        if '|' in info_str:
            parts = [p.strip() for p in info_str.split('|')]
            s_parts = []
            for p in parts:
                m_c = re.search(r'\b(Mã\s+đề(?:\s+thi)?\s*[:=]?\s*[\w\d]+)', p, re.IGNORECASE)
                if m_c and not exam_code:
                    exam_code = m_c.group(1).strip()
                    rem = (p[:m_c.start()] + p[m_c.end():]).strip(' ,;:')
                    if rem:
                        s_parts.append(rem)
                else:
                    s_parts.append(p)
            student_info = ', '.join(s_parts)
        else:
            m_c = re.search(r'\b(Mã\s+đề(?:\s+thi)?\s*[:=]?\s*[\w\d]+)', info_str, re.IGNORECASE)
            if m_c:
                exam_code = m_c.group(1).strip()
                student_info = (info_str[:m_c.start()] + info_str[m_c.end():]).strip(' ,;:')

        return {
            'left': left_lines,
            'right': right_lines,
            'student_info': student_info,
            'exam_code': exam_code
        }, rest_text

    return None, text


def parse_inline_exam_header(text: str) -> tuple[dict | None, str]:
    """
    Nhận diện và bóc tách khối tiêu đề đề thi bị dính liền thành 1 dòng đơn hoặc 1 đoạn liên tục:
    Ví dụ: 'SỞ GD&ĐT THÀNH PHỐ HUẾ TRƯỜNG THPT PHÚ LỘC (Đề thi gồm có 04 trang) ĐỀ KIỂM TRA CUỐI KỲ I, NĂM HỌC 2025-2026 MÔN: TOÁN 12 Thời gian làm bài 90 phút, không kể thời gian phát đề Họ và tên học sinh: ... Mã đề 132 PHẦN I...'
    """
    if not text:
        return None, text

    has_so = bool(re.search(r'\b(?:SỞ|BỘ|PHÒNG)\s+(?:GD&ĐT|GIÁO\s+DỤC)\b', text, re.IGNORECASE))
    has_truong = bool(re.search(r'\bTRƯỜNG\s+(?:THPT|THCS|PHỔ\s+THÔNG)\b', text, re.IGNORECASE))
    has_kythi = bool(re.search(r'\b(?:ĐỀ\s+KIỂM\s+TRA|KỲ\s+THI|KIỂM\s+TRA\s+CUỐI|KIỂM\s+TRA\s+GIỮA|KIỂM\s+TRA\s+HỌC|ĐỀ\s+THI)\b', text, re.IGNORECASE))
    has_mon = bool(re.search(r'\b(?:MÔN|Bài\s+thi)\s*:', text, re.IGNORECASE))

    # Phải có ít nhất 2 trong các dấu hiệu trên nằm ở đầu văn bản
    if not ((has_so or has_truong) and (has_kythi or has_mon)):
        return None, text

    boundary_pat = re.compile(
        r'\b(?:(?:[IVXLCDM]+\.|\d+\.|\b[A-Z]\.)?\s*(?:PHẦN|Phần)\b|(?:Câu|Bài|Ví dụ)\s+(?:\$?\d+\$?)(?:\s*\([^\)\n]*\))?\s*[\.:])'
    )
    m_boundary = boundary_pat.search(text)

    if m_boundary:
        header_text = text[:m_boundary.start()].strip()
        rem_text = text[m_boundary.start():].strip()
    else:
        # Nếu chưa tìm thấy PHẦN hoặc Câu 1:, tìm ranh giới sau Mã đề \d+
        m_made_end = re.search(r'\bMã\s+đề(?:\s+thi)?\s*[:=]?\s*[\w\d]+', text, re.IGNORECASE)
        if m_made_end:
            header_text = text[:m_made_end.end()].strip()
            rem_text = text[m_made_end.end():].strip()
        else:
            return None, text

    p_so = re.search(r'\b((?:SỞ|BỘ|PHÒNG)\s+(?:GD&ĐT|GIÁO\s+DỤC)[^\(\n]*?)(?=\bTRƯỜNG\b|\(Đề\b|\bĐỀ\b|\bKỲ\b|$)', header_text, re.IGNORECASE)
    p_truong = re.search(r'\b(TRƯỜNG\s+(?:THPT|THCS|PHỔ\s+THÔNG)[^\(\n]*?)(?=\(Đề\b|\bĐỀ\b|\bKỲ\b|\bMÔN\b|$)', header_text, re.IGNORECASE)
    p_trang = re.search(r'(\(Đề(?:\s+thi)?\s+(?:gồm\s+có|có)\s+\d+\s+trang\))', header_text, re.IGNORECASE)

    p_kythi = re.search(r'\b((?:ĐỀ\s+KIỂM\s+TRA|KỲ\s+THI|KIỂM\s+TRA|ĐỀ\s+THI(?!\s+(?:gồm|có)))[^\n]*?)(?=\bMÔN\b|\bBÀI\s+THI\b|\bThời\s+gian\b|$)', header_text, re.IGNORECASE)
    p_mon = re.search(r'\b((?:MÔN|Bài\s+thi)\s*:[^\n]*?)(?=\bThời\s+gian\b|\bHọ\s+và\s+tên\b|\bHọ\s+tên\b|$)', header_text, re.IGNORECASE)
    p_tg = re.search(r'\b(Thời\s+gian\s+làm\s+bài[^\n]*?)(?=\bHọ\s+(?:và\s+)?tên\b|\bSố\s+báo\s+danh\b|\bLớp\b|\bMã\s+đề\b|$)', header_text, re.IGNORECASE)

    p_hs = re.search(r'\b(Họ\s+(?:và\s+)?tên[^\n]*?)(?=\bMã\s+đề\b|$)', header_text, re.IGNORECASE)
    p_made = re.search(r'\b(Mã\s+đề(?:\s+thi)?\s*[:=]?\s*[\w\d]+)', header_text, re.IGNORECASE)

    left = []
    if p_so: left.append(p_so.group(1).strip())
    if p_truong: left.append(p_truong.group(1).strip())
    if p_trang: left.append(p_trang.group(1).strip())

    right = []
    if p_kythi: right.append(p_kythi.group(1).strip())
    if p_mon: right.append(p_mon.group(1).strip())
    if p_tg: right.append(p_tg.group(1).strip())

    student_info = p_hs.group(1).strip(' ,;') if p_hs else ''
    exam_code = p_made.group(1).strip() if p_made else ''

    return {
        'left': left,
        'right': right,
        'student_info': student_info,
        'exam_code': exam_code
    }, rem_text


def extract_exam_header(text: str) -> tuple[dict | None, str]:
    """Trích xuất khối tiêu đề đề thi (thử cả cú pháp thẻ, nhận diện đa dòng, và nhận diện chuỗi dính liền inline)."""
    header_info, rem_text = parse_tagged_exam_header(text)
    if header_info:
        return header_info, rem_text
    header_info, rem_text = parse_fallback_exam_header(text)
    if header_info:
        return header_info, rem_text
    return parse_inline_exam_header(text)


def insert_exam_header(doc, header_info: dict, math_format: str = "omml"):
    """
    Tạo khối tiêu đề đề thi chuẩn thể thức Bộ GD&ĐT / Nghị định 30:
    - Bảng 2 cột không viền (borderless table):
      + Cột 1 (trái, căn giữa): Cơ quan chủ quản (chữ in hoa), Tên trường (chữ in hoa đậm), Số trang (chữ nghiêng).
      + Cột 2 (phải, căn giữa): Tên kỳ thi (chữ in hoa đậm), Môn thi (đậm), Thời gian làm bài (chữ nghiêng).
    - Hàng tiếp theo hoặc đoạn dưới: Họ tên, Lớp (căn trái) và Mã đề (đóng khung hoặc in đậm bên phải).
    """
    if not header_info:
        return

    left_lines = header_info.get('left', [])
    right_lines = header_info.get('right', [])
    student_info = header_info.get('student_info', '').strip()
    exam_code = header_info.get('exam_code', '').strip()

    # 1. Bảng 2 cột không viền cho thông tin trường & kỳ thi
    tbl_header = doc.add_table(rows=1, cols=2)
    set_table_borders_none(tbl_header)
    tbl_header.columns[0].width = Cm(7.0)
    tbl_header.columns[1].width = Cm(9.0)

    cell_l = tbl_header.cell(0, 0)
    cell_r = tbl_header.cell(0, 1)
    cell_l.width = Cm(7.0)
    cell_r.width = Cm(9.0)

    # Đổ nội dung Cột Trái (Sở / Trường / Số trang)
    for i, line in enumerate(left_lines):
        p = cell_l.paragraphs[0] if i == 0 else cell_l.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0

        is_page_count = bool(re.search(r'trang\b|chính thức', line, re.IGNORECASE))
        is_school = bool(re.search(r'\bTRƯỜNG\b', line, re.IGNORECASE)) or (i == 1 and not is_page_count)

        run = p.add_run(line)
        run.font.name = 'Times New Roman'
        if is_page_count:
            run.font.size = Pt(10)
            run.italic = True
        elif is_school:
            run.font.size = Pt(11)
            run.bold = True
        else:
            run.font.size = Pt(11)
            if i == 0:
                run.bold = False

    # Đổ nội dung Cột Phải (Kỳ thi / Môn / Thời gian)
    for i, line in enumerate(right_lines):
        p = cell_r.paragraphs[0] if i == 0 else cell_r.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0

        is_time = bool(re.search(r'thời\s+gian|phút\b', line, re.IGNORECASE))
        is_subject = bool(re.search(r'môn\b|bài\s+thi\b', line, re.IGNORECASE))
        is_exam_name = bool(re.search(r'đề\s+kiểm\s+tra|đề\s+thi|kỳ\s+thi', line, re.IGNORECASE)) or (i == 0)

        run = p.add_run(line)
        run.font.name = 'Times New Roman'
        if is_time:
            run.font.size = Pt(10)
            run.italic = True
        elif is_subject:
            run.font.size = Pt(11)
            run.bold = True
        elif is_exam_name:
            run.font.size = Pt(11)
            run.bold = True
        else:
            run.font.size = Pt(11)

    # 2. Dòng Họ tên học sinh & Mã đề thi
    if student_info and not exam_code:
        p_inf = doc.add_paragraph()
        p_inf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_inf.paragraph_format.space_before = Pt(5)
        p_inf.paragraph_format.space_after = Pt(4)
        p_inf.paragraph_format.line_spacing = 1.0
        r_inf = p_inf.add_run(student_info)
        r_inf.font.name = 'Times New Roman'
        r_inf.font.size = Pt(11)
        r_inf.italic = True
    elif student_info or exam_code:
        tbl_info = doc.add_table(rows=1, cols=2)
        set_table_borders_none(tbl_info)
        tbl_info.columns[0].width = Cm(12.5)
        tbl_info.columns[1].width = Cm(3.5)

        c_info = tbl_info.cell(0, 0)
        c_code = tbl_info.cell(0, 1)
        c_info.width = Cm(12.5)
        c_code.width = Cm(3.5)

        if student_info:
            p_inf = c_info.paragraphs[0]
            p_inf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_inf.paragraph_format.space_before = Pt(5)
            p_inf.paragraph_format.space_after = Pt(4)
            p_inf.paragraph_format.line_spacing = 1.0
            r_inf = p_inf.add_run(student_info)
            r_inf.font.name = 'Times New Roman'
            r_inf.font.size = Pt(11)
            r_inf.italic = True

        if exam_code:
            p_cod = c_code.paragraphs[0]
            p_cod.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cod.paragraph_format.space_before = Pt(3)
            p_cod.paragraph_format.space_after = Pt(3)
            p_cod.paragraph_format.line_spacing = 1.0
            set_cell_box_border(c_code)
            r_cod = p_cod.add_run(exam_code)
            r_cod.bold = True
            r_cod.font.name = 'Times New Roman'
            r_cod.font.size = Pt(11)


def insert_placeholder_box(doc, subtype: str, desc: str, math_format: str = "omml"):
    """
    Tạo ô thông báo vị trí Bảng / Hình ảnh khi chưa có file ảnh cắt.
    Tự động render công thức bên trong desc nếu có $...$, không để lộ $ thô.
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell = table.cell(0, 0)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    icon = "📊" if subtype == "table" else "🖼️"
    title = "BẢNG BIẾN THIÊN / BẢNG SỐ LIỆU" if subtype == "table" else "HÌNH VẼ / ĐỒ THỊ"

    run1 = p.add_run(f"{icon} [{title}]\n")
    run1.bold = True
    run1.font.size = Pt(11)
    run1.font.color.rgb = RGBColor(99, 102, 241)

    # Chèn mô tả có công thức OMML chuẩn hoặc LaTeX text
    add_text_and_math_runs(p, desc + "\n", italic=True, font_size=Pt(10), color=RGBColor(100, 116, 139), math_format=math_format)

    run3 = p.add_run("(Dán hoặc cắt ảnh chụp màn hình chèn vào vị trí này)")
    run3.font.size = Pt(9)
    run3.font.color.rgb = RGBColor(148, 163, 184)


def insert_image_into_doc(doc, img_path: Path | str, desc: str = "", math_format: str = "omml"):
    """
    Chèn hình ảnh trực tiếp vào file Word với kích thước chuẩn:
    - Bề rộng tối đa trong khoảng in an toàn của trang A4 (<= 15.5 cm)
    - Tự động điều chỉnh kích thước theo aspect ratio gốc, không làm méo ảnh
    - Đối với ảnh bảng/đồ thị bè ngang (aspect >= 1.8), khống chế width <= 15.0 cm,
      tuyệt đối không ép min height 3.0 cm để tránh làm phình width lên 24-29 cm
    - Cố định cả width và height khi chèn vào paragraph
    - Tuyệt đối không in dòng mô tả / chú thích ảnh bên dưới
    """
    img_path = Path(img_path)
    if not img_path.exists():
        return False

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(6)

    run = p.add_run()
    try:
        with PILImage.open(img_path) as im:
            w_px, h_px = im.size

        aspect = (w_px / h_px) if h_px > 0 else 1.0

        # Giới hạn bề rộng in an toàn tối đa của trang A4 (21cm - 3cm trái - 2cm phải = 16cm)
        max_w_cm = 15.5

        if aspect >= 1.8:
            # Ảnh bảng số liệu / bảng biến thiên / đồ thị bè ngang:
            # Khống chế bề rộng vừa vặn khổ in (tối đa 15.0 cm), chiều cao tính theo aspect
            # Tuyệt đối không ép min height = 3.0 cm để tránh bung bề ngang lên 24-29.5 cm
            target_w_cm = min(max_w_cm, 15.0)
            target_h_cm = target_w_cm / aspect
            # Giới hạn chiều cao an toàn tối đa 5.0 cm
            if target_h_cm > 5.0:
                target_h_cm = 5.0
                target_w_cm = target_h_cm * aspect
        else:
            # Ảnh thông thường (vuông, chữ nhật đứng hoặc ngang vừa)
            if aspect >= 1.3:
                target_h_cm = 3.8
            elif aspect >= 0.8:
                target_h_cm = 4.0
            else:
                target_h_cm = 4.3

            target_w_cm = target_h_cm * aspect
            if target_w_cm > max_w_cm:
                target_w_cm = max_w_cm
                target_h_cm = target_w_cm / aspect
            else:
                target_h_cm = max(2.5, min(5.0, target_h_cm))
                target_w_cm = target_h_cm * aspect
                if target_w_cm > max_w_cm:
                    target_w_cm = max_w_cm
                    target_h_cm = target_w_cm / aspect

        # Cố định cả 2 chiều width và height để đảm bảo 100% không bao giờ tràn lề A4
        run.add_picture(
            str(img_path),
            width=Cm(round(target_w_cm, 2)),
            height=Cm(round(target_h_cm, 2))
        )
        return True
    except Exception as e:
        print(f"Lỗi khi chèn ảnh {img_path}: {e}")
        return False


def export_latex_to_docx(
    pages_data: list[dict],
    output_path: str,
    title: str = "Tài liệu Toán học",
    session_dir: str | Path = None,
    math_format: str = "omml"
) -> str:
    """
    Xuất danh sách các trang LaTeX sang file Word chuẩn:
    - Font Times New Roman 12pt
    - Lề: Top 2cm, Bottom 2cm, Left 3cm, Right 2cm
    - Giãn dòng: Single (1.0), Space after 4pt
    - Tách đoạn trước mỗi Câu X và phương án A. B. C. D.
    - Chèn ảnh trực tiếp nếu có file cắt ảnh (không chèn caption thừa bên dưới)
    - Hỗ trợ 2 chế độ công thức: OMML (Word native equation) hoặc LaTeX text ($...$, $$...$$)
    - Không chèn bảng báo cáo rà soát hay checklist gây thừa
    """
    doc = Document()
    is_latex_mode = (str(math_format).lower() in ("latex", "raw", "text"))

    # Cài đặt lề chuẩn trang A4 theo Nghị định 30/2020/NĐ-CP
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

    # Cài đặt Style Normal chuẩn
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.0  # Single
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.space_before = Pt(0)

    # ── Bóc tách và chèn Header Đề thi 2 cột chuẩn thể thức Bộ GD&ĐT / Nghị định 30 ──
    header_info = None
    for page_info in pages_data:
        l_txt = page_info.get('latex', '')
        h_info, rem_l_txt = extract_exam_header(l_txt)
        if h_info:
            header_info = h_info
            page_info['latex'] = rem_l_txt
            break

    has_exam_header = (header_info is not None)
    if has_exam_header:
        insert_exam_header(doc, header_info, math_format=math_format)
    else:
        # Xóa bỏ hoàn toàn tiêu đề mặc định thừa (TAI_LIEU_TOAN_HOC) khi không cần thiết.
        # Chỉ in tiêu đề nếu người dùng chủ động truyền một tiêu đề tùy chỉnh khác rỗng.
        default_titles = {"", "TÀI LIỆU TOÁN HỌC", "TAI_LIEU_TOAN_HOC", "TAI LIEU TOAN HOC"}
        if title and title.strip().upper() not in default_titles:
            title_p = doc.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_p.paragraph_format.space_after = Pt(14)
            t_run = title_p.add_run(title.upper())
            t_run.bold = True
            t_run.font.size = Pt(14)
            t_run.font.color.rgb = RGBColor(30, 41, 59)

    # Chuẩn bị đường dẫn thư mục tạm để tìm ảnh
    sess_path = Path(session_dir) if session_dir else None
    used_image_names = set()

    for page_info in pages_data:
        latex_text = page_info.get('latex', '')
        if not latex_text.strip():
            continue

        latex_text = remove_page_footers(latex_text)
        latex_text = sanitize_text_aesthetics(latex_text)
        latex_text = unwrap_vietnamese_math(latex_text)
        latex_text = heal_broken_math_and_coords(latex_text)
        latex_text = join_broken_ocr_lines(latex_text)
        latex_text = sanitize_text_aesthetics(latex_text)
        paragraphs = split_into_paragraphs(latex_text)

        for para_text in paragraphs:
            segments = parse_latex_content(para_text)
            if not segments:
                continue

            # ── 1. Kiểm tra nếu có placeholder (Bảng / Hình vẽ) ──
            placeholder_segs = [s for s in segments if s.get('type') == 'placeholder']
            non_placeholder_segs = [s for s in segments if s.get('type') != 'placeholder']

            if placeholder_segs:
                for info in placeholder_segs:
                    img_name = info.get('image')
                    inserted = False

                    # Nếu chưa có img_name nhưng có box và có session_dir
                    if not img_name and info.get('box') and sess_path and sess_path.exists():
                        existing_crops = sorted(list(sess_path.glob("autocrop_*.png")) + list(sess_path.glob("crop_*.png")), key=lambda f: f.stat().st_mtime)
                        for c_file in existing_crops:
                            if c_file.name not in used_image_names:
                                img_name = c_file.name
                                break

                    # Tìm file ảnh đã cắt linh hoạt (làm sạch nếu bị dính metadata question="..." ||)
                    if img_name:
                        img_name = re.sub(r'question\s*=\s*(?:"[^"]*"|\'[^\']*\'|[^\s\|\]]+)\s*\|*', '', img_name, flags=re.IGNORECASE)
                        m_file = re.search(r'((?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg)|[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg))', img_name, re.IGNORECASE)
                        if m_file:
                            img_name = m_file.group(1)
                        else:
                            img_name = img_name.strip(' "\'|')

                        possible_img = None
                        if sess_path and (sess_path / img_name).exists():
                            possible_img = sess_path / img_name
                        elif (Path("temp_output") / img_name).exists():
                            possible_img = Path("temp_output") / img_name
                        elif sess_path and sess_path.exists():
                            matches = list(sess_path.rglob(img_name))
                            if matches:
                                possible_img = matches[0]
                        elif Path("temp_output").exists():
                            matches = list(Path("temp_output").rglob(img_name))
                            if matches:
                                possible_img = matches[0]
                        elif Path(img_name).exists():
                            possible_img = Path(img_name)
                        else:
                            matches = list(Path(".").rglob(img_name))
                            if matches:
                                possible_img = matches[0]

                        if possible_img and possible_img.exists():
                            inserted = insert_image_into_doc(doc, possible_img, info.get('desc', ''), math_format=math_format)
                            if inserted:
                                used_image_names.add(possible_img.name)

                    # Nếu vẫn chưa có ảnh cắt nhưng có crop files trong session_dir chưa dùng
                    if not inserted and sess_path and sess_path.exists():
                        existing_crops = sorted(list(sess_path.glob("autocrop_*.png")) + list(sess_path.glob("crop_*.png")), key=lambda f: f.stat().st_mtime)
                        for c_file in existing_crops:
                            if c_file.name not in used_image_names:
                                inserted = insert_image_into_doc(doc, c_file, info.get('desc', ''), math_format=math_format)
                                if inserted:
                                    used_image_names.add(c_file.name)
                                    break

                    # Nếu vẫn chưa tìm thấy nhưng trong temp_output có crop files phù hợp
                    if not inserted and Path("temp_output").exists():
                        all_crops = sorted(list(Path("temp_output").rglob("autocrop_*.png")) + list(Path("temp_output").rglob("crop_*.png")), key=lambda f: f.stat().st_mtime)
                        for c_file in all_crops:
                            if c_file.name not in used_image_names:
                                inserted = insert_image_into_doc(doc, c_file, info.get('desc', ''), math_format=math_format)
                                if inserted:
                                    used_image_names.add(c_file.name)
                                    break

                    # Nếu không có ảnh cắt nào, hiển thị khung placeholder sạch
                    if not inserted:
                        insert_placeholder_box(doc, info['subtype'], info.get('desc', ''), math_format=math_format)

                if not non_placeholder_segs:
                    continue
                segments = non_placeholder_segs

            # ── 2. Kiểm tra nếu đoạn chỉ gồm 1 công thức display ($$...$$) ──
            if len(segments) == 1 and segments[0]['type'] == 'display_math':
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(6)
                if is_latex_mode:
                    p.add_run(f"$${segments[0]['content']}$$")
                else:
                    add_math_to_paragraph(p, segments[0]['content'], is_display=True)
                continue

            # Tái tạo chuỗi văn bản cho đoạn để phân loại trắc nghiệm hoặc câu hỏi
            rem_text = ""
            for seg in segments:
                if seg['type'] == 'text_runs':
                    rem_text += "".join(r['text'] for r in seg['runs'])
                elif seg['type'] == 'inline_math':
                    rem_text += f"${seg['content']}$"
                elif seg['type'] == 'display_math':
                    rem_text += f"$${seg['content']}$$"
            rem_text = rem_text.strip()

            # Bỏ qua hoàn toàn nếu đoạn chỉ là chú thích ảnh thô sót lại
            if re.match(r'^\s*\|[^\]\n]*\]?\s*$', rem_text):
                continue
            if re.match(r'^(?:\|\s*|\[)?\s*(?:Đồ thị|Hình vẽ|Bảng biến thiên|Bảng số liệu|Hình ảnh|Minh họa|Sơ đồ|mô phỏng)\b', rem_text, re.IGNORECASE):
                continue

            # ── 3. Phân loại & xử lý TIÊU ĐỀ PHẦN (PHẦN 1, PHẦN I, I. PHẦN, PHẦN CÂU TRẮC NGHIỆM,...) ──
            m_sec = re.match(r'^\s*((?:[IVXLCDM]+\.|\d+\.|\b[A-Z]\.)?\s*(?:PHẦN|Phần)\b.*)$', rem_text, re.DOTALL | re.IGNORECASE)
            if m_sec:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(4)

                sec_title = m_sec.group(1).strip()
                add_text_and_math_runs(p, sec_title, bold=True, font_size=Pt(12), math_format=math_format)
                continue

            # ── 4. Phân loại & xử lý phương án trắc nghiệm A., B., C., D. hoặc ý tự luận a), b), c) / a/, b/, c/ / 1), 2) ──
            m_choice = re.match(r'^\s*([A-D]\.|\([A-Da-d]\)|[a-d][\)\/]|\([1-9]\)|[1-9][\)\/])\s*(.*)$', rem_text, re.DOTALL)
            if m_choice:
                label_candidate = m_choice.group(1).strip()
                body_candidate = m_choice.group(2).strip()

                # Kiểm tra nếu nhãn là số 1-9 (1), 2), 3)...):
                # Phải là mục tự luận thực sự (phía sau có nội dung văn bản/toán có nghĩa, không chỉ là dấu chấm/phẩy/ngoặc/đô la/rỗng)
                is_valid_choice = True
                if re.match(r'^[1-9][\)\/]$', label_candidate):
                    if not body_candidate or re.match(r'^[\.,;:\$\s\)]*$', body_candidate):
                        is_valid_choice = False

                if is_valid_choice:
                    p = doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.left_indent = Inches(0.2)

                    label_str = label_candidate
                    r_lbl = p.add_run(label_str + " ")
                    r_lbl.bold = True
                    r_lbl.italic = False
                    r_lbl.font.name = 'Times New Roman'
                    r_lbl.font.size = Pt(12)

                    body_text = sanitize_text_aesthetics(body_candidate)

                # Tự động đưa biểu thức toán học phía sau vào Equation (OMML) nếu chưa được bọc $...$
                if not is_latex_mode and '$' not in body_text:
                    m_dot = re.match(r'^(.*?)([\.]?)$', body_text)
                    expr = m_dot.group(1).strip() if m_dot else body_text
                    trailing_dot = m_dot.group(2) if m_dot else ""
                    # Nhận diện: số đơn lẻ (0, 1, 21, -4...), tập hợp, biến số, khoảng/đoạn, phép tính
                    is_math_expr = (
                        bool(re.match(r'^[-+]?\d+(?:[\.,]\d+)?$', expr)) or  # Số đơn lẻ: 0, 1, 2, 5, 21, -4...
                        bool(re.match(r'^\{[-+0-9a-zA-Z;,\s\.\\]+\}$', expr)) or  # Tập hợp: {-4; 1; 0}
                        bool(re.match(r'^[a-zA-Z]$', expr)) or  # Ký hiệu chữ đơn lẻ: x, y, m
                        bool(re.search(r'[=<>≤≥≠\+\-\*\/_^\\]|(?:\b[a-zA-Z]\b.*?\d)|\d+\s*[a-zA-Z]|[\(\[][-+]?\d+', expr))
                    ) and not bool(re.search(r'\b(hàm\s+số|đồng\s+biến|nghịch\s+biến|liên\s+tục|có|không|với\s+mọi|thỏa\s+mãn|biểu\s+diễn)\b', expr, re.IGNORECASE))
                    if is_math_expr:
                        body_text = f"${expr}${trailing_dot}"

                body_segments = parse_latex_content(body_text)
                for b_seg in body_segments:
                    if b_seg['type'] == 'text_runs':
                        for r_info in b_seg['runs']:
                            txt = r_info['text']
                            if not txt:
                                continue
                            run = p.add_run(txt)
                            run.bold = r_info.get('bold', False)
                    elif b_seg['type'] == 'inline_math':
                        if is_latex_mode:
                            p.add_run(f"${b_seg['content']}$")
                        else:
                            add_math_to_paragraph(p, b_seg['content'], is_display=False)
                    elif b_seg['type'] == 'display_math':
                        if is_latex_mode:
                            p.add_run(f"$${b_seg['content']}$$")
                        else:
                            add_math_to_paragraph(p, b_seg['content'], is_display=True)
                continue

            # ── 5. Phân loại & xử lý câu hỏi (Câu 1:, Câu 2., Bài 1 (1,5 điểm)...) ──
            m_question = re.match(r'^(\s*(?:Câu|Bài|Đề|Ví dụ)\s+(?:\$?\d+\$?)(?:\s*\([^\)\n]*\))?\s*[\.:]\s*)(.*)$', rem_text, re.DOTALL)
            if m_question:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(3)

                label_str = m_question.group(1).strip()
                # 🎯 TASK 4: Tách điểm số tự luận ra ngoài công thức toán, hiển thị in nghiêng sau nhãn câu hỏi in đậm
                # Ví dụ: **Câu 1** *(0,75 điểm)*: ... hoặc **Bài 2** *(1,5 điểm)*: ...
                m_score = re.match(r'^(.*?)(\s*\([^\)\n]*\))\s*([:\.])$', label_str)
                if m_score:
                    main_lbl = m_score.group(1).strip()
                    score_str = m_score.group(2).strip()
                    punct = m_score.group(3).strip()

                    r1 = p.add_run(main_lbl)
                    r1.bold = True
                    r1.font.name = 'Times New Roman'
                    r1.font.size = Pt(12)

                    r2 = p.add_run(f" {score_str}")
                    r2.bold = False
                    r2.italic = True
                    r2.font.name = 'Times New Roman'
                    r2.font.size = Pt(11)

                    r3 = p.add_run(f"{punct} ")
                    r3.bold = True
                    r3.font.name = 'Times New Roman'
                    r3.font.size = Pt(12)
                else:
                    r_lbl = p.add_run(label_str + " ")
                    r_lbl.bold = True
                    r_lbl.font.name = 'Times New Roman'
                    r_lbl.font.size = Pt(12)

                body_text = m_question.group(2).strip()
                body_segments = parse_latex_content(body_text)
                for b_seg in body_segments:
                    if b_seg['type'] == 'text_runs':
                        for r_info in b_seg['runs']:
                            txt = r_info['text']
                            if not txt:
                                continue
                            run = p.add_run(txt)
                            run.bold = False
                    elif b_seg['type'] == 'inline_math':
                        if is_latex_mode:
                            p.add_run(f"${b_seg['content']}$")
                        else:
                            add_math_to_paragraph(p, b_seg['content'], is_display=False)
                    elif b_seg['type'] == 'display_math':
                        if is_latex_mode:
                            p_disp = doc.add_paragraph()
                            p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_disp.paragraph_format.space_before = Pt(4)
                            p_disp.paragraph_format.space_after = Pt(6)
                            p_disp.add_run(f"$${b_seg['content']}$$")
                        else:
                            p_disp = doc.add_paragraph()
                            p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_disp.paragraph_format.space_before = Pt(4)
                            p_disp.paragraph_format.space_after = Pt(6)
                            add_math_to_paragraph(p_disp, b_seg['content'], is_display=True)
                continue

            # ── 6. Phân loại & xử lý KHỐI LỜI GIẢI / HƯỚNG DẪN GIẢI ──
            m_sol = re.match(r'^\s*(\b(?:Lời\s+giải|Hướng\s+dẫn\s+giải|Giải\s+chi\s+tiết|HƯỚNG\s+DẪN\s+GIẢI|LỜI\s+GIẢI|GIẢI\s+CHI\s+TIẾT|HDG)\s*[:\.\-]?)(.*)$', rem_text, re.DOTALL | re.IGNORECASE)
            if m_sol:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(2)
                sol_lbl = m_sol.group(1).strip()
                r_sol = p.add_run(sol_lbl + " ")
                r_sol.bold = True
                r_sol.italic = False
                r_sol.font.name = 'Times New Roman'
                r_sol.font.size = Pt(12)
                r_sol.font.color.rgb = RGBColor(15, 23, 42)

                body_text = m_sol.group(2).strip()
                if body_text:
                    body_segments = parse_latex_content(body_text)
                    for b_seg in body_segments:
                        if b_seg['type'] == 'text_runs':
                            for r_info in b_seg['runs']:
                                txt = r_info['text']
                                if not txt:
                                    continue
                                if re.search(r'\b(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b', txt, re.IGNORECASE):
                                    parts_sel = re.split(r'(\b(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b[\.:]?)', txt, flags=re.IGNORECASE)
                                    for p_sel in parts_sel:
                                        if re.match(r'^\b(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b[\.:]?$', p_sel, re.IGNORECASE):
                                            run = p.add_run(p_sel)
                                            run.bold = True
                                            run.font.color.rgb = RGBColor(16, 120, 60)
                                        else:
                                            run = p.add_run(p_sel)
                                            run.bold = r_info.get('bold', False)
                                else:
                                    run = p.add_run(txt)
                                    run.bold = r_info.get('bold', False)
                        elif b_seg['type'] == 'inline_math':
                            if is_latex_mode:
                                p.add_run(f"${b_seg['content']}$")
                            else:
                                add_math_to_paragraph(p, b_seg['content'], is_display=False)
                        elif b_seg['type'] == 'display_math':
                            if is_latex_mode:
                                p_disp = doc.add_paragraph()
                                p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p_disp.paragraph_format.space_before = Pt(4)
                                p_disp.paragraph_format.space_after = Pt(6)
                                p_disp.add_run(f"$${b_seg['content']}$$")
                            else:
                                p_disp = doc.add_paragraph()
                                p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p_disp.paragraph_format.space_before = Pt(4)
                                p_disp.paragraph_format.space_after = Pt(6)
                                add_math_to_paragraph(p_disp, b_seg['content'], is_display=True)
                continue

            # ── 7. Phân loại & xử lý dòng kết luận độc lập: Chọn A., Đáp án: B... ──
            m_sel_only = re.match(r'^\s*(\b(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b[\.:]?)(.*)$', rem_text, re.DOTALL | re.IGNORECASE)
            if m_sel_only:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                sel_lbl = m_sel_only.group(1).strip()
                r_sel = p.add_run(sel_lbl + " ")
                r_sel.bold = True
                r_sel.font.name = 'Times New Roman'
                r_sel.font.size = Pt(12)
                r_sel.font.color.rgb = RGBColor(16, 120, 60)

                body_text = m_sel_only.group(2).strip()
                if body_text:
                    body_segments = parse_latex_content(body_text)
                    for b_seg in body_segments:
                        if b_seg['type'] == 'text_runs':
                            for r_info in b_seg['runs']:
                                txt = r_info['text']
                                if not txt:
                                    continue
                                run = p.add_run(txt)
                                run.bold = r_info.get('bold', False)
                        elif b_seg['type'] == 'inline_math':
                            if is_latex_mode:
                                p.add_run(f"${b_seg['content']}$")
                            else:
                                add_math_to_paragraph(p, b_seg['content'], is_display=False)
                        elif b_seg['type'] == 'display_math':
                            if is_latex_mode:
                                p_disp = doc.add_paragraph()
                                p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p_disp.paragraph_format.space_before = Pt(4)
                                p_disp.paragraph_format.space_after = Pt(6)
                                p_disp.add_run(f"$${b_seg['content']}$$")
                            else:
                                p_disp = doc.add_paragraph()
                                p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p_disp.paragraph_format.space_before = Pt(4)
                                p_disp.paragraph_format.space_after = Pt(6)
                                add_math_to_paragraph(p_disp, b_seg['content'], is_display=True)
                continue

            # ── 8. Đoạn văn bản kết hợp công thức thông thường ──
            # Lưới an toàn phòng thủ kép: nếu đoạn văn bản còn sót thẻ placeholder chưa bóc tách
            if re.search(r'\[\s*(?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE)\b', rem_text, re.IGNORECASE) or re.search(r'\b(?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg)\b', rem_text, re.IGNORECASE):
                tag_fallback = parse_placeholder_tag(rem_text)
                if tag_fallback:
                    fb_img_name = tag_fallback.get('image')
                    fb_inserted = False
                    if fb_img_name:
                        possible_fb = None
                        if sess_path and (sess_path / fb_img_name).exists():
                            possible_fb = sess_path / fb_img_name
                        elif (Path("temp_output") / fb_img_name).exists():
                            possible_fb = Path("temp_output") / fb_img_name
                        elif sess_path and sess_path.exists():
                            m_fb = list(sess_path.rglob(fb_img_name))
                            if m_fb: possible_fb = m_fb[0]
                        elif Path("temp_output").exists():
                            m_fb = list(Path("temp_output").rglob(fb_img_name))
                            if m_fb: possible_fb = m_fb[0]
                        elif Path(fb_img_name).exists():
                            possible_fb = Path(fb_img_name)
                        else:
                            m_fb = list(Path(".").rglob(fb_img_name))
                            if m_fb: possible_fb = m_fb[0]

                        if possible_fb and possible_fb.exists():
                            fb_inserted = insert_image_into_doc(doc, possible_fb, tag_fallback.get('desc', ''), math_format=math_format)
                            if fb_inserted:
                                used_image_names.add(possible_fb.name)

                    if not fb_inserted and sess_path and sess_path.exists():
                        existing_crops = sorted(list(sess_path.glob("autocrop_*.png")) + list(sess_path.glob("crop_*.png")), key=lambda f: f.stat().st_mtime)
                        for c_file in existing_crops:
                            if c_file.name not in used_image_names:
                                fb_inserted = insert_image_into_doc(doc, c_file, tag_fallback.get('desc', ''), math_format=math_format)
                                if fb_inserted:
                                    used_image_names.add(c_file.name)
                                    break

                    if not fb_inserted:
                        insert_placeholder_box(doc, tag_fallback['subtype'], tag_fallback.get('desc', ''), math_format=math_format)
                    continue

            p = doc.add_paragraph()
            for seg in segments:
                if seg['type'] == 'text_runs':
                    for r_info in seg['runs']:
                        txt = r_info['text']
                        if not txt:
                            continue
                        run = p.add_run(txt)
                        run.bold = r_info.get('bold', False)
                elif seg['type'] == 'inline_math':
                    if is_latex_mode:
                        p.add_run(f"${seg['content']}$")
                    else:
                        add_math_to_paragraph(p, seg['content'], is_display=False)
                elif seg['type'] == 'display_math':
                    if is_latex_mode:
                        p_disp = doc.add_paragraph()
                        p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_disp.paragraph_format.space_before = Pt(4)
                        p_disp.paragraph_format.space_after = Pt(6)
                        p_disp.add_run(f"$${seg['content']}$$")
                    else:
                        p_disp = doc.add_paragraph()
                        p_disp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_disp.paragraph_format.space_before = Pt(4)
                        p_disp.paragraph_format.space_after = Pt(6)
                        add_math_to_paragraph(p_disp, seg['content'], is_display=True)

    doc.save(output_path)
    return output_path
