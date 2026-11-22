"""
OMML (Office Math Markup Language) to LaTeX Converter
Hỗ trợ chuyển đổi toàn diện công thức Word Equation sang LaTeX cho trình chiếu KaTeX:
- Phân số: <m:f> -> \\frac{tử}{mẫu}
- Số mũ: <m:sSup> -> {cơ số}^{số mũ}
- Chỉ số dưới: <m:sSub> -> {cơ số}_{chỉ số}
- Vừa trên vừa dưới: <m:sSubSup> -> {cơ số}_{dưới}^{trên}
- Căn thức: <m:rad> -> \\sqrt{...} hoặc \\sqrt[n]{...}
- Dấu vectơ / Accent: <m:acc> -> \\vec{...}
- Dấu ngoặc: <m:d> -> \\left( ... \\right)
- Tích phân, Tổng: <m:nary> -> \\int, \\sum
- Ma trận: <m:m> -> \\begin{matrix}...\\end{matrix}
- Làm sạch vectơ dạng chữ: B'D'→, u→, A'A→ -> \\vec{...}
"""

import re
from lxml import etree

OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MATH_TAG = lambda name: f"{{{OMML_NS}}}{name}"


def normalize_primes(text: str) -> str:
    """
    Quy đổi toàn bộ biến thể Unicode prime về dấu nháy đơn ' chuẩn:
    - \u2032 (PRIME ′) -> '
    - \u2033 (DOUBLE PRIME ″) -> ''
    - \u2034 (TRIPLE PRIME ‴) -> '''
    - \u2019 (RIGHT SINGLE QUOTATION ’) -> '
    - \u02b9 (MODIFIER LETTER PRIME ʹ) -> '
    - \u2035 (REVERSED PRIME ‵) -> '
    - \u02bc (MODIFIER LETTER APOSTROPHE ʼ) -> '
    """
    if not text:
        return ""
    text = text.replace("\u2034", "'''").replace("\u2033", "''")
    return re.sub(r"[\u2032\u2019\u02b9\u2035\u02bc]", "'", text)


def extract_clean_vector_base(base_str: str) -> str:
    """
    Thu gom chuỗi ký tự gốc cho vectơ, dọn sạch khoảng trắng/xuống dòng bên trong
    và unwrap ngoặc nhọn thừa (ví dụ: {B'D'} -> B'D', {u} -> u, { B' } -> B').
    """
    if not base_str:
        return ""
    s = normalize_primes(base_str)
    # Dọn sạch toàn bộ khoảng trắng, tab, xuống dòng bên trong tên vector
    s = re.sub(r"\s+", "", s)
    # Gỡ bỏ các cặp ngoặc nhọn bao toàn bộ nếu có
    while s.startswith("{") and s.endswith("}"):
        depth = 0
        is_matched = True
        for idx, ch in enumerate(s):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0 and idx < len(s) - 1:
                    is_matched = False
                    break
        if is_matched and depth == 0:
            s = s[1:-1]
        else:
            break
    # Gỡ bỏ các ngoặc nhọn đơn lẻ bên trong nếu có (ví dụ: {B'}{D'} -> B'D')
    s = re.sub(r"\{([A-Za-z0-9_']+)\}", r"\1", s)
    return s



def make_vector_latex(clean_base: str) -> str:
    r"""
    Tạo cú pháp LaTeX chuẩn cho vectơ:
    - Nếu là 1 chữ cái đơn lẻ không có prime (u, v, a, b, x, n): dùng \vec{u}.
    - Nếu là 2 ký tự trở lên (AB, AC', A'B', MN, C'A...) hoặc có dấu prime ('):
      BẮT BUỘC dùng \overrightarrow{...} để KaTeX và LaTeX tự động co giãn (stretch)
      thanh mũi tên phủ kín từ ký tự đầu đến ký tự cuối, không bao giờ bị lệch sang trái/phải/giữa!
    """
    if not clean_base:
        return ""
    clean = re.sub(r"\s+", "", clean_base)
    letters_only = re.sub(r"[^a-zA-Z0-9]", "", clean)
    if len(letters_only) >= 2 or "'" in clean:
        return f"\\overrightarrow{{{clean}}}"
    return f"\\vec{{{clean}}}"

def omml_element_to_latex(element) -> str:
    """Chuyển đổi đệ quy một phần tử OMML sang chuỗi LaTeX chuẩn."""
    if element is None:
        return ""

    tag = etree.QName(element).localname if isinstance(element.tag, str) else ""

    # 1. Text run
    if tag == "t":
        text = element.text or ""
        # Chuẩn hóa dấu trừ unicode
        text = text.replace("\u2212", "-").replace("−", "-")
        # Chuẩn hóa dấu nhân
        text = text.replace("\u22c5", " \\cdot ").replace("·", " \\cdot ")
        # Chuẩn hóa các biến thể dấu phẩy đỉnh (prime) sang dấu nháy đơn chuẩn
        text = normalize_primes(text)
        return text

    # 2. Container oMath hoặc oMathPara
    if tag in ("oMath", "oMathPara"):
        res = [omml_element_to_latex(c) for c in element]
        return "".join(res)

    # 3. Phân số <m:f>
    if tag == "f":
        num_el = element.find(MATH_TAG("num"))
        den_el = element.find(MATH_TAG("den"))
        num_str = omml_element_to_latex(num_el).strip() if num_el is not None else ""
        den_str = omml_element_to_latex(den_el).strip() if den_el is not None else ""
        return f"\\frac{{{num_str}}}{{{den_str}}}"

    # 4. Số mũ <m:sSup>
    if tag == "sSup":
        e_el = element.find(MATH_TAG("e"))
        sup_el = element.find(MATH_TAG("sup"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        exp_str = omml_element_to_latex(sup_el).strip() if sup_el is not None else ""
        
        # Chuẩn hóa prime cho số mũ
        clean_exp = normalize_primes(exp_str).strip()
        # Nếu số mũ là dấu phẩy đỉnh (1 hoặc nhiều ký tự ' hoặc ′ \u2032):
        # Trả về TRỰC TIẾP base' (ví dụ: A', B', D''), TUYỆT ĐỐI KHÔNG sinh ra {base}^{'} hay dấu mũ ^
        exp_pure = re.sub(r"[\s\{\}]", "", clean_exp)
        if exp_pure and all(ch == "'" for ch in exp_pure):
            base_clean = re.sub(r"^\{([^{}]+)\}$", r"\1", base_str)
            return f"{base_clean}{exp_pure}"
        return f"{{{base_str}}}^{{{exp_str}}}"

    # 5. Chỉ số dưới <m:sSub>
    if tag == "sSub":
        e_el = element.find(MATH_TAG("e"))
        sub_el = element.find(MATH_TAG("sub"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        sub_str = omml_element_to_latex(sub_el).strip() if sub_el is not None else ""
        return f"{{{base_str}}}_{{{sub_str}}}"

    # 6. Cả chỉ số dưới và trên <m:sSubSup>
    if tag == "sSubSup":
        e_el = element.find(MATH_TAG("e"))
        sub_el = element.find(MATH_TAG("sub"))
        sup_el = element.find(MATH_TAG("sup"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        sub_str = omml_element_to_latex(sub_el).strip() if sub_el is not None else ""
        exp_str = omml_element_to_latex(sup_el).strip() if sup_el is not None else ""
        clean_exp = normalize_primes(exp_str).strip()
        exp_pure = re.sub(r"[\s\{\}]", "", clean_exp)
        if exp_pure and all(ch == "'" for ch in exp_pure):
            base_clean = re.sub(r"^\{([^{}]+)\}$", r"\1", base_str)
            return f"{{{base_clean}{exp_pure}}}_{{{sub_str}}}"
        return f"{{{base_str}}}_{{{sub_str}}}^{{{exp_str}}}"

    # 7. Căn thức <m:rad>
    if tag == "rad":
        deg_el = element.find(f"{MATH_TAG('radPr')}/{MATH_TAG('deg')}")
        e_el = element.find(MATH_TAG("e"))
        inner_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        deg_str = omml_element_to_latex(deg_el).strip() if deg_el is not None else ""
        if deg_str:
            return f"\\sqrt[{deg_str}]{{{inner_str}}}"
        return f"\\sqrt{{{inner_str}}}"

    # 8. Dấu vectơ / Accent <m:acc>
    if tag == "acc":
        chr_el = element.find(f"{MATH_TAG('accPr')}/{MATH_TAG('chr')}")
        e_el = element.find(MATH_TAG("e"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        val = chr_el.get(MATH_TAG("val")) if chr_el is not None else "→"
        clean_base = extract_clean_vector_base(base_str)
        # Mũi tên vectơ
        if val in ("→", "\u2192", "\u20d7", "&#8594;", "vec"):
            return f"\\vec{{{clean_base}}}"
        elif val in ("¯", "‾", "\u203e", "\u00af", "bar"):
            return f"\\overline{{{clean_base}}}"
        elif val in ("^", "hat"):
            return f"\\hat{{{clean_base}}}"
        return f"\\vec{{{clean_base}}}"

    # 9. Dấu ngoặc <m:d>
    if tag == "d":
        beg_el = element.find(f"{MATH_TAG('dPr')}/{MATH_TAG('begChr')}")
        end_el = element.find(f"{MATH_TAG('dPr')}/{MATH_TAG('endChr')}")
        beg = beg_el.get(MATH_TAG("val"), "(") if beg_el is not None else "("
        end = end_el.get(MATH_TAG("val"), ")") if end_el is not None else ")"
        # Xử lý các thành phần con <m:e>
        e_els = element.findall(MATH_TAG("e"))
        if e_els:
            inner_str = "".join(omml_element_to_latex(e) for e in e_els)
        else:
            inner_str = "".join(omml_element_to_latex(c) for c in element if etree.QName(c).localname not in ("dPr",))
        
        # Thoát các dấu ngoặc nhọn nếu cần
        if beg == "{": beg = "\\{"
        if end == "}": end = "\\}"
        if beg == "" and end == "":
            return inner_str
        return f"\\left{beg}{inner_str}\\right{end}"

    # 10. Giới hạn dưới <m:limLow>
    if tag == "limLow":
        e_el = element.find(MATH_TAG("e"))
        lim_el = element.find(MATH_TAG("lim"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        lim_str = omml_element_to_latex(lim_el).strip() if lim_el is not None else ""
        return f"{base_str}_{{{lim_str}}}"

    # 10b. Giới hạn trên <m:limUpp> (RẤT PHỔ BIẾN TRONG WORD ĐỂ BIỂU DIỄN VÉC-TƠ NHƯ B'D'→, A'C'→ HOẶC ĐOẠN THẲNG GẠCH ĐẦU)
    if tag == "limUpp":
        e_el = element.find(MATH_TAG("e"))
        lim_el = element.find(MATH_TAG("lim"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        lim_str = omml_element_to_latex(lim_el).strip() if lim_el is not None else ""
        clean_base = extract_clean_vector_base(base_str)
        # Kiểm tra nếu lim là mũi tên vectơ
        if any(arr in lim_str for arr in ("→", "\u2192", "\u20d7", "\\rightarrow", "\\to", "\\longrightarrow", "\\rightharpoonup", "\\Rightarrow", "vec")):
            return f"\\vec{{{clean_base}}}"
        elif any(b in lim_str for b in ("¯", "‾", "\u203e", "\u00af", "bar", "\\bar", "\\overline", "-")):
            return f"\\overline{{{clean_base}}}"
        elif any(h in lim_str for h in ("^", "hat", "\\hat")):
            return f"\\hat{{{clean_base}}}"
        return f"\\overset{{{lim_str}}}{{{clean_base}}}"

    # 10c. Nhóm ký tự / Group Character <m:groupChr> (mũi tên hoặc gạch đầu bao bọc nhóm đỉnh)
    if tag == "groupChr":
        chr_el = element.find(f"{MATH_TAG('groupChrPr')}/{MATH_TAG('chr')}")
        pos_el = element.find(f"{MATH_TAG('groupChrPr')}/{MATH_TAG('pos')}")
        e_el = element.find(MATH_TAG("e"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        val = chr_el.get(MATH_TAG("val"), "→") if chr_el is not None else "→"
        pos = pos_el.get(MATH_TAG("val"), "top") if pos_el is not None else "top"
        clean_base = extract_clean_vector_base(base_str)
        if pos == "top":
            if any(arr in val for arr in ("→", "\u2192", "\u20d7", "\\rightarrow", "\\to", "vec")):
                return f"\\vec{{{clean_base}}}"
            elif any(b in val for b in ("¯", "‾", "\u203e", "\u00af", "bar", "\\bar", "\\overline", "-")):
                return f"\\overline{{{clean_base}}}"
            return f"\\vec{{{clean_base}}}"
        else:
            if any(arr in val for arr in ("→", "\u2192", "\u20d7", "\\rightarrow", "\\to")):
                return f"\\underrightarrow{{{clean_base}}}"
            return f"\\underline{{{clean_base}}}"

    # 10d. Hộp đóng khung <m:borderBox>
    if tag == "borderBox":
        e_el = element.find(MATH_TAG("e"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        return base_str

    # 11. Tích phân / Tổng <m:nary>
    if tag == "nary":
        chr_el = element.find(f"{MATH_TAG('naryPr')}/{MATH_TAG('chr')}")
        sub_el = element.find(MATH_TAG("sub"))
        sup_el = element.find(MATH_TAG("sup"))
        e_el = element.find(MATH_TAG("e"))
        op_char = chr_el.get(MATH_TAG("val"), "∫") if chr_el is not None else "∫"
        
        latex_op = "\\int" if op_char in ("∫", "\u222b") else ("\\sum" if op_char in ("∑", "\u2211") else op_char)
        sub_str = f"_{{{omml_element_to_latex(sub_el).strip()}}}" if sub_el is not None else ""
        sup_str = f"^{{{omml_element_to_latex(sup_el).strip()}}}" if sup_el is not None else ""
        e_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        return f"{latex_op}{sub_str}{sup_str} {e_str}"

    # 12. Gạch đầu <m:bar>
    if tag == "bar":
        e_el = element.find(MATH_TAG("e"))
        base_str = omml_element_to_latex(e_el).strip() if e_el is not None else ""
        clean_base = extract_clean_vector_base(base_str)
        return f"\\overline{{{clean_base}}}"

    # Mặc định: duyệt các phần tử con
    parts = []
    if element.text:
        parts.append(element.text)
    for c in element:
        parts.append(omml_element_to_latex(c))
        if c.tail:
            parts.append(c.tail)
    return "".join(parts)


def clean_math_latex(latex_str: str) -> str:
    """Làm sạch và chuẩn hóa toàn diện chuỗi LaTeX xuất từ OMML."""
    if not latex_str:
        return ""
    s = latex_str.strip()

    # 0. Chuẩn hóa tất cả các biến thể dấu phẩy đỉnh (prime) về dấu nháy đơn chuẩn
    s = normalize_primes(s)

    # 1. Khắc phục triệt để lỗi dấu phẩy đỉnh bị văng dấu mũ (ví dụ: A'^B'^C'^D'^, A'^B^\vec{'}, C'^\vecA, B^\vec{'})
    # a. Thay ^\{'+\} hoặc ^'+ thành ' hoặc '' tương ứng
    s = re.sub(r"\^\s*\{([']+)\}", r"\1", s)
    s = re.sub(r"([A-Za-z0-9_'])\s*\^\s*\{([']+)\}", r"\1\2", s)
    s = re.sub(r"([A-Za-z0-9_'])\s*\^\s*([']+)", r"\1\2", s)
    # b. Sửa ^\vec{'} -> ' (ví dụ B^\vec{'} -> B', A'B^\vec{'} -> A'B')
    s = re.sub(r"\^\s*\\vec\{([']+)\}", r"\1", s)
    s = re.sub(r"([A-Za-z0-9_'])\s*\^\s*\\vec\{([']+)\}", r"\1\2", s)
    # c. Sửa ^\vec{A} -> \vec{A} (ví dụ C'^\vecA -> C'\vec{A} hoặc \vec{C'A})
    s = re.sub(r"\^\s*\\vec\{([A-Za-z0-9_']+)\}", r"\\vec{\1}", s)
    s = re.sub(r"\^\s*\\vec([A-Za-z0-9_']+)", r"\\vec{\1}", s)
    # d. Loại bỏ dấu ^ văng ra sau dấu phẩy đỉnh nếu không có số mũ hợp lệ đằng sau (ví dụ: ABCD.A'^B'^C'^D'^ -> ABCD.A'B'C'D')
    s = re.sub(r"([']+)\^+(?=[A-Za-z\s\$\.,;\)\]\}]|$)", r"\1", s)

    # 2. Khắc phục lỗi vectơ bị vỡ mất ngoặc nhọn: \vecDB, \vecBD, \vecu, \vecB'D' -> bọc lại \vec{...}
    s = re.sub(r"\\vec(?![a-zA-Z]*\{)([A-Za-z_\']+)", r"\\vec{\1}", s)

    # 3. Gỡ bỏ các cặp ngoặc nhọn bọc lẻ biến số hoặc đỉnh không thuộc lệnh LaTeX (ví dụ: {B'}{D'} -> B'D', {A}' -> A', {B} -> B)
    def _unwrap_single_braces(m):
        cmd = m.group(1)
        content = m.group(2)
        prime = m.group(3) or ""
        if cmd:
            return f"{cmd}{{{content}}}{prime}"
        return f"{content}{prime}"

    s = re.sub(r"(\\[a-zA-Z]+)?\{([A-Za-z0-9_']+)\}([']*)", _unwrap_single_braces, s)

    # 4. Dọn dẹp khoảng trắng / xuống dòng bên trong dấu \vec{...}
    s = re.sub(r"\\vec\{([^}]+)\}", lambda m: f"\\vec{{{re.sub(r'\s+', '', m.group(1))}}}", s)

    # 5. Tách rời mũi tên nếu dính chữ tiếng Việt hoặc dính $ (ví dụ: \rightarrowbằng -> \rightarrow bằng, \rightarrow$là -> \rightarrow là)
    s = re.sub(r"(\\rightarrow|\\to|[\u2192→])\$?\s*([a-zA-Zà-ỹÀ-Ỹ])", r"\1 \2", s)

    # 6a. Sửa trường hợp Word Equation accent chỉ bọc chữ cái đầu (ví dụ: \vec{D}B -> \vec{DB}, \vec{B}D -> \vec{BD}, \vec{A'}C' -> \vec{A'C'})
    s = re.sub(r"\\(?:vec|overrightarrow)\{([A-Z][']*)\}\s*([A-Z][']*)", lambda m: make_vector_latex(f"{m.group(1)}{m.group(2)}"), s)
    # 6b. Sửa trường hợp Word Equation accent chỉ bọc chữ cái sau (ví dụ: A' \vec{A} -> \vec{A'A}, C' \vec{A} -> \vec{C'A}, A' \vec{C} -> \vec{A'C}, B\vec{D} -> \vec{BD})
    s = re.sub(r"([A-Z][']*)\s*\\(?:vec|overrightarrow)\{([A-Z][']*)\}", lambda m: make_vector_latex(f"{m.group(1)}{m.group(2)}"), s)

    def _vec_repl(m):
        target = m.group(1).replace(" ", "")
        return make_vector_latex(target)

    # Chuẩn hóa tích vô hướng dấu chấm giữa 2 vectơ: \vec{AC'}.\vec{CD'} -> \overrightarrow{AC'} \cdot \overrightarrow{CD'}
    s = re.sub(r"(\\(?:vec|overrightarrow)\{[^}]+\})\s*\.\s*(\\(?:vec|overrightarrow)\{[^}]+\})", r"\1 \\cdot \2", s)

    # Sửa các dạng vector bị rơi mũi tên ra sau (CHỈ CHẤP NHẬN CHỮ CÁI, CẤM CHỮ SỐ 0-9 VÀ CẤM \vec{...}):
    vec_unit = r"(?:[A-Z][']*(?:\s*[A-Z][']*)+|[a-z][']*)"
    suffix_arrow = r"(?:\\rightarrow|\\to|[\u2192→]|\\vec(?!\{)|\\overrightarrow(?!\{))(?![a-zA-Z\{])"
    s = re.sub(rf"(?<![\\a-zA-Z0-9])({vec_unit})\s*{suffix_arrow}", _vec_repl, s)
    s = re.sub(rf"(?:\\rightarrow|\\to|[\u2192→])\s*({vec_unit})", _vec_repl, s)
    s = re.sub(rf"\\vec\s+({vec_unit})", _vec_repl, s)
    # 10. \overrightarrow{AB} -> \vec{AB}
    # Nâng cấp các vectơ đa ký tự hoặc có prime thành \\overrightarrow để KaTeX/Word phủ kín thanh mũi tên
    s = re.sub(r"\\(?:vec|overrightarrow)\{([^}]+)\}", lambda m: make_vector_latex(m.group(1)), s)

    # Chuẩn hóa khoảng trắng quanh toán tử (BẮT BUỘC dùng \b cho \le, \ge, \cdot để KHÔNG làm rách \left, \right)
    s = re.sub(r"\s*([=+\-]|<|>|\\le\b|\\ge\b|\\cdot\b)\s*", r" \1 ", s)
    # Thu gọn khoảng trắng thừa
    s = re.sub(r"[ \t]+", " ", s).strip()
    return s


def mask_math_tokens(text: str) -> tuple[str, dict]:
    """
    Đóng két bảo vệ công thức (Math Tokenization - Chuẩn Azota):
    Mask toàn bộ công thức toán vào token ẩn @@__MATH_TOKEN_i__@@
    trước khi chạy các bộ lọc cấu trúc đề thi. Tuyệt đối không cho Regex văn bản chạy đè lên cú pháp LaTeX.
    """
    if not text:
        return "", {}
    tokens = {}
    counter = 0

    # Match block math $$...$$ trước, sau đó inline math $...$
    math_pattern = re.compile(r'(\$\$[\s\S]*?\$\$|(?<!\$)\$(?!\$)[^\$\n]+?(?<!\$)\$(?!\$))')

    def _mask(m):
        nonlocal counter
        tok = f"@@__MATH_TOKEN_{counter}__@@"
        tokens[tok] = m.group(0)
        counter += 1
        return tok

    masked_text = math_pattern.sub(_mask, text)
    return masked_text, tokens


def unmask_math_tokens(text: str, tokens: dict) -> str:
    """Hoàn nguyên các token công thức toán học về chuỗi LaTeX nguyên bản 100%."""
    if not text or not tokens:
        return text
    for tok, orig in tokens.items():
        text = text.replace(tok, orig)
    return text


MEDIA_TAG_REGEX = re.compile(
    r'(\[\s*(?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|BẢNG(?:[ _]SỐ[ _]LIỆU)?|IMAGE|FIGURE|TABLE)\s*[:=]?\s*(?:[^\[\]]|\[(?:[^\[\]]|\[[^\]]*\])*\])*\])',
    re.IGNORECASE
)


def mask_media_tags(text: str) -> tuple[str, dict]:
    """Bảo vệ tuyệt đối 100% các thẻ media/hình ảnh/bảng biểu chứa base64 để không bị các hàm xử lý văn bản làm hỏng."""
    if not text:
        return "", {}
    tokens = {}
    counter = 0

    def _mask(m):
        nonlocal counter
        tok = f"@@__PROTECTED_MEDIA_SAFE_TOKEN_{counter}__@@"
        tokens[tok] = m.group(0)
        counter += 1
        return tok

    masked_text = MEDIA_TAG_REGEX.sub(_mask, text)
    return masked_text, tokens


def unmask_media_tags(text: str, tokens: dict) -> str:
    """Hoàn nguyên các thẻ media/hình ảnh về dữ liệu gốc 100% không suy suyển 1 byte."""
    if not text or not tokens:
        return text
    for tok, orig in tokens.items():
        text = text.replace(tok, orig)
    return text


def separate_stuck_math_and_words(text: str) -> str:
    """
    Tách rời thông minh và tổng quát các từ tiếng Việt bị dính với ký hiệu toán, số hoặc đơn vị đo:
    Ví dụ:
    - 50bình -> 50 bình, 10người -> 10 người, 30triệu -> 30 triệu, 6km -> 6 km, kmdưới -> km dưới
    - Atrên bờ -> A trên bờ, Btrên -> B trên, Clà -> C là, Mlà -> M là
    - BCvuông góc -> BC vuông góc, ACsao cho -> AC sao cho, AMBthì -> AMB thì
    - \rightarrowbằng -> \rightarrow bằng, \rightarrow$là -> \rightarrow là
    - $formula$là -> $formula$ là, $formula$và -> $formula$ và
    - (-1; 3; -2)và -> (-1; 3; -2) và, (-2; 0; 3)là -> (-2; 0; 3) là
    """
    if not text:
        return ""

    # KHIÊN BẢO VỆ: Cất an toàn toàn bộ thẻ hình ảnh Base64
    text, media_tokens = mask_media_tags(text)

    # 0. Chuẩn hóa tất cả các biến thể dấu phẩy đỉnh (prime)
    text = re.sub(r"[\u2032\u2019\u02b9\u2035]", "'", text)

    # 1. Tách mũi tên dính từ tiếng Việt (ví dụ: \rightarrowbằng -> \rightarrow bằng, →bằng -> → bằng, \rightarrow$là -> \rightarrow$ là)
    text = re.sub(r"(\\rightarrow|\\to|[\u2192→])(\$?)\s*([a-zA-Zà-ỹÀ-Ỹ])", r"\1\2 \3", text)

    # 1c. Tách dấu đóng ngoặc ) (của tọa độ điểm, bộ số hoặc biểu thức) dính với chữ cái tiếng Việt
    text = re.sub(r"(?<=[0-9a-zA-Z'\)])\)\s*([a-zA-Zà-ỹÀ-Ỹ])", r") \1", text)

    # 2. QUY TẮC TỔNG QUÁT: Tách mọi số đứng liền trước chữ cái tiếng Việt (50bình -> 50 bình, 10người -> 10 người, 30triệu -> 30 triệu...)
    # Bỏ qua nếu là số mũ toán học có ^ hoặc sau _
    text = re.sub(r"(?<![\^_])(\d+)([a-zA-Zà-ỹÀ-Ỹ]+)", r"\1 \2", text)

    # 3. Tách đơn vị đo dính với từ tiếng Việt (kmdưới -> km dưới, mtrên -> m trên, cmtại -> cm tại)
    text = re.sub(r"\b(km|cm|mm|dm|m|kg|g|l|ml)(dưới|trên|trong|ngoài|vuông|bằng|nước|đất|biển|sông|đến|tại)\b", r"\1 \2", text, flags=re.IGNORECASE)

    # 4. Tách chữ hoa toán học (tên điểm, đoạn thẳng, tam giác: A, B, C, M, BC, AC, AMB...) dính với các từ tiếng Việt phổ biến
    vn_common_words = (
        r"trên|dưới|trong|ngoài|là|sao|thì|vuông|cắt|tiếp|đồng|nghịch|điểm|đoạn|đường|"
        r"mặt|góc|tại|với|và|thuộc|không|khi|cho|bằng|nằm|có|qua|đến|từ|song|chứng|minh|"
        r"thỏa|thỏa\s+mãn|giao|chiếu|hướng|khoảng|cách"
    )
    # Tách khi có tên điểm hoa A, B, C, M hoặc cụm điểm BC, AC, AMB dính từ tiếng Việt (KHÔNG dùng IGNORECASE)
    text = re.sub(rf"\b([A-Z]{{1,6}}[']*)({vn_common_words})\b", r"\1 \2", text)

    # 5. QUY TẮC TỔNG QUÁT: Tách cụm 2 chữ cái hoa toán học trở lên đứng trước bất kỳ từ thường tiếng Việt (BCvuông -> BC vuông, ACsao -> AC sao, AMBthì -> AMB thì)
    text = re.sub(r"\b([A-Z]{2,6}[']*)([a-zà-ỹ]{2,})\b", r"\1 \2", text)

    # 6. Tách từ tiếng Việt kết thúc dính với chữ cái hoa toán học (điểmA -> điểm A, tam giácABC -> tam giác ABC, đoạnAC -> đoạn AC)
    text = re.sub(r"\b(điểm|đỉnh|đoạn|tam\s+giác|mặt\s+phẳng|đường\s+thẳng|tại|từ|đến)([A-Z][']*)\b", r"\1 \2", text, flags=re.IGNORECASE)

    # 7. Tách từ nối tiếng Việt phổ biến khi đứng sát sau ngoặc hoặc $
    vn_connectors = r"là|và|với|bằng|khi|thì|nếu|thuộc|không|tại|trong|qua|cho|sao"
    text = re.sub(rf"(\))\s*({vn_connectors})\b", r"\1 \2", text, flags=re.IGNORECASE)

    # Thu gọn khoảng trắng thừa
    text = re.sub(r"[ \t]+", " ", text)

    # HOÀN NGUYÊN: Trả lại thẻ hình ảnh Base64 nguyên vẹn 100%
    text = unmask_media_tags(text, media_tokens)

    # Áp dụng Smart Spacing Bridge đảm bảo ranh giới math và text luôn có khoảng trắng chuẩn
    text = ensure_math_spacing_bridge(text)
    return text


def ensure_math_spacing_bridge(text: str) -> str:
    """Đảm bảo tuyệt đối không dính chữ giữa công thức toán học $...$ hoặc $$...$$ và văn bản tiếng Việt xung quanh."""
    if not text:
        return ""
    text, media_tokens = mask_media_tags(text)

    # Tách các khối math $$...$$ hoặc $...$
    parts = []
    tokens = re.split(r'(\$\$[\s\S]*?\$\$|\$[^\$]+?\$)', text)
    for tok in tokens:
        if (tok.startswith('$$') and tok.endswith('$$') and len(tok) >= 4) or (tok.startswith('$') and tok.endswith('$') and len(tok) >= 2):
            parts.append(tok)
        else:
            parts.append(tok)

    def _is_math(s: str) -> bool:
        return (s.startswith('$$') and s.endswith('$$') and len(s) >= 4) or (s.startswith('$') and s.endswith('$') and len(s) >= 2)

    safe_parts = []
    n = len(parts)
    for i in range(n):
        curr = parts[i]
        if _is_math(curr):
            safe_parts.append(curr)
        else:
            # Nếu trước nó là math token và curr bắt đầu bằng chữ cái tiếng Việt/số
            if i > 0 and _is_math(parts[i - 1]):
                if curr and re.match(r'^[a-zA-Zà-ỹÀ-Ỹ0-9]', curr):
                    curr = " " + curr
            # Nếu sau nó là math token và curr kết thúc bằng chữ cái tiếng Việt/số
            if i < n - 1 and _is_math(parts[i + 1]):
                if curr and re.search(r'[a-zA-Zà-ỹÀ-Ỹ0-9]$', curr):
                    curr = curr + " "
            safe_parts.append(curr)

    res = "".join(safe_parts)
    # Tách dấu ngoặc ) đóng của tọa độ / số dính chữ cái tiếng Việt
    res = re.sub(r'(?<=[0-9a-zA-Z\)])\)\s*([a-zA-Zà-ỹÀ-Ỹ])', r') \1', res)
    res = re.sub(r'[ \t]{2,}', ' ', res)
    res = unmask_media_tags(res, media_tokens)
    return res


def sanitize_plain_text_vectors(text: str) -> str:
    """
    Tìm và chuẩn hóa các vectơ viết dạng thường ngoài Equation trong văn bản:
    Ví dụ:
    - B'D'→, A'C' \rightarrow, A'B' \rightarrow -> $\\vec{B'D'}$, $\\vec{A'C'}$, $\\vec{A'B'}$
    - u→ -> $\\vec{u}$
    - A'A→ -> $\\vec{A'A}$
    - A'\vec{A} -> $\\vec{A'A}$
    - u→=A'A→+A'B'→+A'D'→ -> $\\vec{u} = \\vec{A'A} + \\vec{A'B'} + \\vec{A'D'}$
    - B'D'\\vec -> $\\vec{B'D'}$
    """
    if not text:
        return ""

    # KHIÊN BẢO VỆ: Cất an toàn toàn bộ thẻ hình ảnh Base64
    text, media_tokens = mask_media_tags(text)

    # Tách trước các từ dính
    text = separate_stuck_math_and_words(text)

    # Sửa trường hợp accent Word bị lệch cả 2 chiều (ví dụ: \vec{D}B -> \vec{DB}, B\vec{D} -> \vec{BD}, A'\vec{A} -> \vec{A'A})
    text = re.sub(r"\\(?:vec|overrightarrow)\{([A-Z][']*)\}\s*([A-Z][']*)", lambda m: make_vector_latex(f"{m.group(1)}{m.group(2)}"), text)
    text = re.sub(r"([A-Z][']*)\s*\\(?:vec|overrightarrow)\{([A-Z][']*)\}", lambda m: make_vector_latex(f"{m.group(1)}{m.group(2)}"), text)

    # Chuẩn hóa tích vô hướng dấu chấm giữa 2 vectơ ngoài Equation hoặc trong Equation
    text = re.sub(r"(\\(?:vec|overrightarrow)\{[^}]+\})\s*\.\s*(\\(?:vec|overrightarrow)\{[^}]+\})", r"\1 \\cdot \2", text)

    vec_unit_name = r"(?:[A-Z][']*(?:\s*[A-Z][']*)+|[a-z][']*)"
    suffix_arrow = r"(?:\\rightarrow|\\to|[\u2192→]|\\vec(?!\{)|\\overrightarrow(?!\{))(?![a-zA-Z\{])"
    vec_pattern = rf"(?<![\\a-zA-Z0-9])({vec_unit_name})\s*{suffix_arrow}"

    # Hàm thay thế an toàn cho từng đoạn ngoài dấu $ ... $
    parts = []
    tokens = re.split(r"(\$[^$]+\$)", text)
    for token in tokens:
        if token.startswith("$") and token.endswith("$"):
            # Bên trong công thức: chuẩn hóa vectơ
            inner = token[1:-1]
            inner = clean_math_latex(inner)
            parts.append(f"${inner}$")
        else:
            # Bên ngoài công thức:
            # 1. Bọc các \vec{...} chưa có dấu $ bên ngoài
            t = re.sub(r"\\(?:vec|overrightarrow)\{([^}]+)\}", lambda m: f"${make_vector_latex(m.group(1))}$", token)

            # 2. Phát hiện biểu thức vectơ liên hoàn như u→=A'A→+A'B'→+A'D'→
            def _repl_vec_expr(m):
                expr = m.group(0)
                expr = re.sub(vec_pattern, lambda x: make_vector_latex(x.group(1).replace(' ', '')), expr)
                expr = re.sub(r"\s*([=+\-])\s*", r" \1 ", expr)
                return f" ${expr.strip()}$ "

            # Nhận diện chuỗi có chứa vectơ và dấu mũi tên
            t = re.sub(rf"(?:(?<![\\a-zA-Z0-9]){vec_unit_name}\s*{suffix_arrow})(?:\s*[=+\-]\s*(?:(?<![\\a-zA-Z0-9]){vec_unit_name}\s*{suffix_arrow}))*", _repl_vec_expr, t)

            # Thay thế đơn lẻ còn sót lại
            t = re.sub(vec_pattern, lambda x: f"${make_vector_latex(x.group(1).replace(' ', ''))}$", t)
            parts.append(t)

    result = "".join(parts)
    # Áp dụng Smart Spacing Bridge đảm bảo ranh giới công thức không bao giờ dính chữ
    result = ensure_math_spacing_bridge(result)

    # HOÀN NGUYÊN: Trả lại thẻ hình ảnh Base64 nguyên vẹn 100%
    result = unmask_media_tags(result, media_tokens)
    return result
