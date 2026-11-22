"""
Module Giải Mã Công Thức MathType Native (MTEF & TeX Stream)
Dành cho đề thi Toán Việt Nam:
1. Trích xuất trực tiếp chuỗi TeX Input Language từ ảnh WMF (nếu có).
2. Giải mã luồng nhị phân MTEF (MathType Equation Format v3/v4/v5) từ OLE Object (.bin).
3. Chuẩn hóa công thức sang chuẩn KaTeX sắc nét ($...$).
"""

import re
from typing import Optional

try:
    from core.mtef.mtef import MTEF
except Exception:
    MTEF = None


def unwrap_outer_braces(tex: str) -> str:
    """Unwrap outer matching curly braces if balanced across the entire string."""
    if not tex:
        return ""
    while tex.startswith('{') and tex.endswith('}'):
        depth = 0
        is_matched = True
        for idx, ch in enumerate(tex):
            # Bỏ qua escaped \{ và \}
            if ch == '\\' and idx + 1 < len(tex) and tex[idx + 1] in ('{', '}'):
                continue
            if ch == '{' and (idx == 0 or tex[idx - 1] != '\\'):
                depth += 1
            elif ch == '}' and (idx == 0 or tex[idx - 1] != '\\'):
                depth -= 1
                if depth == 0 and idx < len(tex) - 1:
                    is_matched = False
                    break
        if is_matched and depth == 0:
            tex = tex[1:-1].strip()
        else:
            break
    return tex


def balance_curly_braces(tex: str) -> str:
    """
    Đảm bảo số lượng ngoặc nhọn mở và đóng unescaped luôn cân bằng 100%:
    Nếu thừa '{' ở đầu hoặc thiếu '}' ở cuối, tự động cân bằng chống lỗi KaTeX đỏ.
    """
    if not tex:
        return ""
    tex = unwrap_outer_braces(tex)
    # Loại bỏ escaped \{ và \} khi đếm
    clean = re.sub(r'\\[{}]', '', tex)
    open_c = clean.count('{')
    close_c = clean.count('}')
    if open_c > close_c:
        if tex.startswith('{') and open_c - close_c == 1:
            tex = tex[1:].strip()
        else:
            tex = tex + ('}' * (open_c - close_c))
    elif close_c > open_c:
        if tex.endswith('}') and close_c - open_c == 1:
            tex = tex[:-1].strip()
        else:
            tex = ('{' * (close_c - open_c)) + tex
    return tex


def extract_tex_from_wmf(wmf_bytes: bytes) -> str:
    """Tìm chuỗi TeX Input Language nhúng trong ảnh WMF của MathType nếu có"""
    if not wmf_bytes:
        return ""
    # MathType TeX stream có thể chứa nhiều dòng \r\n, kết thúc bằng byte \x00
    m = re.search(rb'(?:%?TeX Input Language)\x00([^\x00]+)', wmf_bytes)
    if m:
        tex = m.group(1).decode('utf-8', errors='ignore').strip()
        tex = re.sub(r'[\r\n]+', ' ', tex).strip()
        tex = unwrap_outer_braces(tex)
        tex = balance_curly_braces(tex)
        return tex
    return ""


def clean_mathtype_latex(tex: str) -> str:
    """Chuẩn hóa ký hiệu đặc thù của MathType/MTEF sang cú pháp LaTeX chuẩn"""
    if not tex:
        return ""
    # 0. Loại bỏ toàn bộ ký tự rác Private Use Area (PUA 0xE000-0xF8FF) của MathType căn lề
    tex = re.sub(r'[\ue000-\uf8ff]', '', tex)
    # Chuẩn hóa chữ cái Latin Eth / Đ tiếng Việt bị MathType mã hóa lệch
    tex = tex.replace('\u00d0', 'Đ').replace('Ð', 'Đ').replace('ð', 'đ')

    # Chuẩn hóa prime
    tex = tex.replace('\u2032', "'").replace('^{\\prime}', "'").replace('^\\prime', "'").replace('\\prime', "'")
    tex = tex.replace('∫', '\\int ')
    tex = tex.replace('≈', ' \\approx ')
    tex = tex.replace('≤', ' \\le ')
    tex = tex.replace('≥', ' \\ge ')
    tex = tex.replace('≠', ' \\ne ')
    tex = tex.replace('±', ' \\pm ')
    tex = tex.replace('×', ' \\times ')
    tex = tex.replace('÷', ' \\div ')
    tex = tex.replace('⇔', ' \\Leftrightarrow ')
    tex = tex.replace('⇒', ' \\Rightarrow ')
    tex = tex.replace('→', ' \\rightarrow ')
    tex = tex.replace('∈', ' \\in ')
    tex = tex.replace('∉', ' \\notin ')
    tex = tex.replace('∞', ' \\infty ')
    tex = tex.replace('°', '^\\circ ')
    tex = re.sub(r'\{\s*~([a-zA-Z]+)\s*\}', r'\\text{ \1}', tex)
    tex = re.sub(r'\{\s*\\rm\{\s*([a-zA-Z]+)\s*\}\s*\}', r'\\mathrm{\1}', tex)
    tex = re.sub(r'\\rm\{\s*([a-zA-Z]+)\s*\}', r'\\mathrm{\1}', tex)

    # Học từ MathRender: Chuyển hệ phương trình \left\{\begin{array}{l} sang \begin{cases} chuẩn KaTeX
    tex = re.sub(r'\\left\\\{\s*\\begin\{array\}\s*\{?\s*l\s*\}?', r'\\begin{cases}', tex)
    if r'\begin{cases}' in tex and r'\end{array}' in tex:
        cases_count = tex.count(r'\begin{cases}')
        array_end_count = tex.count(r'\end{array}')
        if cases_count == array_end_count:
            tex = tex.replace(r'\end{array}', r'\end{cases}')
            tex = re.sub(r'\\right\.', '', tex)

    tex = balance_curly_braces(tex)
    tex = re.sub(r'\s+', ' ', tex).strip()
    return tex


def extract_mathtype_formula(ole_bytes: Optional[bytes], wmf_bytes: Optional[bytes] = None) -> str:
    """
    Giải mã đa tầng:
    Tầng 1: Đọc chuỗi TeX nguyên bản từ WMF (nét 100%).
    Tầng 2: Giải mã byte stream MTEF từ oleObject.bin qua MTEF Parser.
    """
    # 1. Ưu tiên lấy từ chuỗi TeX trong WMF nếu có
    if wmf_bytes:
        tex = extract_tex_from_wmf(wmf_bytes)
        if tex:
            return clean_mathtype_latex(tex)

    # 2. Giải mã bằng MTEF parser
    if ole_bytes and MTEF:
        try:
            mtef_obj, err = MTEF.OpenBytes(ole_bytes)
            if not err and mtef_obj:
                tex = mtef_obj.Translate().strip()
                if tex.startswith('$') and tex.endswith('$'):
                    tex = tex[1:-1].strip()
                return clean_mathtype_latex(tex)
        except Exception:
            pass

    return ""


def extract_mathtype_from_element(container_elem, doc) -> str:
    """
    Kiểm tra xem phần tử (Run hoặc Object) có chứa đối tượng MathType OLE không.
    Nếu có, trích xuất và trả về chuỗi công thức LaTeX.
    Nếu không phải MathType, trả về chuỗi rỗng "".
    """
    objs = [e for e in container_elem.iter() if e.tag.endswith('object')]
    if not objs:
        # Nếu chính container_elem là object
        if container_elem.tag.endswith('object'):
            objs = [container_elem]
        else:
            return ""

    obj = objs[0]
    ole_bytes = None
    wmf_bytes = None
    is_mathtype = False

    for elem in obj.iter():
        if elem.tag.endswith('OLEObject'):
            prog_id = elem.get('ProgID', '')
            if any(k in prog_id for k in ['Equation', 'MathType', 'DSMT']):
                is_mathtype = True
            rId = elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rId and rId in doc.part.rels:
                ole_bytes = doc.part.rels[rId].target_part.blob
                if b'MathType' in ole_bytes or b'DSMT' in ole_bytes or b'Equation' in ole_bytes:
                    is_mathtype = True
        elif elem.tag.endswith('imagedata'):
            rId = elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rId and rId in doc.part.rels:
                wmf_bytes = doc.part.rels[rId].target_part.blob
                if b'MathType' in wmf_bytes or b'DSMT' in wmf_bytes or b'Design Science' in wmf_bytes:
                    is_mathtype = True

    if is_mathtype or ole_bytes or wmf_bytes:
        return extract_mathtype_formula(ole_bytes, wmf_bytes)

    return ""
