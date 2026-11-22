"""
Math OCR Tool - Phase 1 MVP
Chuyển PDF/Ảnh → Gemini OCR → LaTeX → Preview
"""

import io
import os
import re
import json
import uuid
import shutil
import base64
import asyncio
import zipfile
from pathlib import Path
from dotenv import load_dotenv

import sys
if sys.platform == 'win32':
    try:
        from asyncio.proactor_events import _ProactorBasePipeTransport
        _orig_call_connection_lost = _ProactorBasePipeTransport._call_connection_lost
        def _silent_call_connection_lost(self, exc):
            try:
                _orig_call_connection_lost(self, exc)
            except (ConnectionResetError, OSError):
                pass
        _ProactorBasePipeTransport._call_connection_lost = _silent_call_connection_lost
    except Exception:
        pass

import pymupdf as fitz
from PIL import Image
from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Form, Request, Response
from fastapi.responses import HTMLResponse, StreamingResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from google import genai
from core.word_exporter import (
    export_latex_to_docx,
    join_broken_ocr_lines,
    heal_broken_math_and_coords,
    unwrap_vietnamese_math,
    sanitize_text_aesthetics,
    remove_page_footers,
    auto_wrap_math_symbols,
    split_sub_items_to_lines
)
from core.omml_converter import (
    omml_element_to_latex,
    clean_math_latex,
    sanitize_plain_text_vectors,
    separate_stuck_math_and_words,
    ensure_math_spacing_bridge,
    mask_media_tags,
    unmask_media_tags,
    mask_math_tokens,
    unmask_math_tokens,
    normalize_primes
)
from core.pdf_native_extractor import (
    extract_native_images_from_doc,
    find_matching_native_image,
    is_black_or_corrupted_image
)
from core.mathtype_extractor import extract_mathtype_from_element, extract_mathtype_formula

# Load .env
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Quản lý danh sách API Keys (hỗ trợ phân tách bằng dấu phẩy, chấm phẩy hoặc xuống dòng)
DEFAULT_API_KEYS = []
if GEMINI_API_KEY:
    for k in re.split(r'[,;\n\r]+', GEMINI_API_KEY):
        k = k.strip()
        if k and not k.startswith("#"):
            DEFAULT_API_KEYS.append(k)

# Cache các client theo api key để tái sử dụng
_client_cache = {}

def get_gemini_client(api_key: str):
    """Lấy hoặc khởi tạo genai.Client cho từng API key"""
    if not api_key:
        return None
    if api_key not in _client_cache:
        _client_cache[api_key] = genai.Client(api_key=api_key)
    return _client_cache[api_key]

# Initialize default client
ai_client = get_gemini_client(DEFAULT_API_KEYS[0]) if DEFAULT_API_KEYS else None

app = FastAPI(title="Math OCR Tool")

# Thư mục tạm
TEMP_DIR = Path("temp_output")
TEMP_DIR.mkdir(exist_ok=True)

PROMPTS_DIR = Path("prompts")
PROMPTS_DIR.mkdir(exist_ok=True)

def load_prompt_file(filename: str, default_text: str) -> str:
    """Đọc prompt từ thư mục prompts/ nếu có, nếu chưa có thì dùng default_text."""
    p_path = PROMPTS_DIR / filename
    if p_path.exists():
        try:
            with open(p_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    return content
        except Exception as e:
            print(f"Lỗi đọc prompt {filename}: {e}")
    return default_text


# ══════════════════════════════════════════════
# PROMPT PRESETS
# ══════════════════════════════════════════════

DEFAULT_QUIZ_PROMPT = r"""Bạn là chuyên gia OCR tài liệu và đề thi toán học tiếng Việt THPT. Hãy chuyển đổi CHÍNH XÁC nội dung trong ảnh thành văn bản sạch, tuân thủ các quy tắc:

1. TUYỆT ĐỐI KHÔNG dùng dấu sao ** để in đậm (viết 'Câu 1.' hoặc 'Câu 1:' thay vì '**Câu 1.**'). TUYỆT ĐỐI KHÔNG bọc số thứ tự câu vào dấu $ (viết 'Câu 1:', 'Câu 2:' thay vì 'Câu $1$:', 'Câu $2$:').
2. BẮT BUỘC XUỐNG DÒNG RÕ RÀNG TRƯỚC CÁC MỤC:
   - Các tiêu đề PHẦN (ví dụ: 'PHẦN I', 'PHẦN II', 'PHẦN TỰ LUẬN') BẮT BUỘC nằm trên MỘT DÒNG RIÊNG BIỆT (xuống dòng trước và sau tiêu đề Phần).
   - Trước mỗi 'Câu X:' hoặc 'Bài X (điểm):' BẮT BUỘC phải xuống dòng riêng biệt. TUYỆT ĐỐI KHÔNG viết dính liền '... D. 12 Câu 9: ...' hoặc 'PHẦN TỰ LUẬN Bài 1:...'.
   - Mỗi phương án trắc nghiệm A., B., C., D. BẮT BUỘC nằm trên MỘT DÒNG RIÊNG BIỆT.
   - Mỗi ý tự luận a), b), c), d) hoặc 1), 2)... BẮT BUỘC nằm trên MỘT DÒNG RIÊNG BIỆT (xuống dòng trước a), b), c)...).
3. BẮT BUỘC NHÚNG MỌI KÝ HIỆU TOÁN HỌC VÀO CẶP DẤU $...$:
   - Tên điểm, đỉnh hình học: $A$, $B$, $C$, $S$, $M$, $N$, $P$, $G$, $O$, $A'$, $B'$... (KHÔNG viết chữ A, B trần)
   - Tọa độ điểm, tọa độ vectơ: $G(1; 2; 3)$, $A(0; 2)$, $I(-2; 3)$, $M(3; 3; -2)$, $A'(0; 0; 4)$, $\vec{a} = (-1; 3; -2)$, $\vec{b} = (-2; 0; 3)$, $\vec{u}(x; y; z)$, $(2; 3; 4)$, $(a; b; c)$
   - Trục và hệ tọa độ, mặt phẳng: $Ox$, $Oy$, $Oz$, $Oxyz$, $Oxy$, $(Oxy)$, $(Oyz)$, $(Oxz)$, $(P)$, $(Q)$, $(\alpha)$, đồ thị $(C)$
   - Tên hình hình học: tam giác $ABC$, tam giác $OAB$, hình chóp $S.ABCD$, hình hộp $ABCD.A'B'C'D'$
   - Tham số, biến số, hàm số: $x$, $y$, $z$, $m$, $t$, $f(x)$, $y = f(x)$
   - Tập hợp và quan hệ phần tử: tập hợp $A$, tập hợp $X$, tập hợp số $\mathbb{N}$, $\mathbb{Z}$, $\mathbb{Q}$, $\mathbb{R}$; tập hợp liệt kê phần tử: $\{-4; 1; 0; 8; -3; 7\}$; các ký hiệu: $\in$, $\notin$, $\subset$, $\cup$, $\cap$, $\emptyset$ (TUYỆT ĐỐI KHÔNG để trần như '{-4; 1; 0}' hay '1 in A').
   - Các phép tính, biểu thức và số học: phép tính $(-36) + 15$, $60 : 5$, $2^2 \cdot 15$, $4 \cdot 6^2 - 5 \cdot 2^3$; số nguyên âm $-4$, $-21$...
   - TUYỆT ĐỐI KHÔNG ngắt dòng \n ở giữa cặp dấu $...$ (ví dụ: cấm viết '$G(a; b;\n c)$' hay '$(-1; 3; -\n2)$'). Tọa độ điểm và vectơ bắt buộc viết liền mạch.
   - TUYỆT ĐỐI KHÔNG bọc cả câu văn bản tiếng Việt vào trong cặp dấu $...$ (ví dụ: CẤM viết 'Đồ thị $ (C) \text{ đi qua điểm } A(0; 2) $' hay '$ I(-2; 3) \text{ là giao điểm...} $'; CHỈ bọc đúng ký hiệu toán: 'Đồ thị $(C)$ đi qua điểm $A(0; 2)$', 'Điểm $I(-2; 3)$ là giao điểm của các đường tiệm cận của đồ thị $(C)$').
4. LUÔN CÓ DẤU CÁCH giữa chữ tiếng Việt và công thức toán $...$ (ví dụ: 'Cho hàm số $y = f(x)$ có đồ thị', KHÔNG viết dính chữ 'số$y=f(x)$có').
5. Mọi công thức toán PHẢI viết bằng LaTeX chuẩn:
   - Inline trong dòng: $...$
   - Công thức riêng dòng: $$...$$
   - Vector: dùng \vec{AB} (ví dụ: $\vec{AB} + \vec{AD}$)
   - Phân số: \frac{a}{b}, Căn: \sqrt{x}, Lũy thừa: x^2, Chỉ số: x_1
   - Tích phân, Giới hạn, Lượng giác: \int, \lim, \sin, \cos, \tan
6. ĐỐI VỚI HÌNH VẼ, ĐỒ THỊ, HÌNH HỌC KHÔNG GIAN:
   - TUYỆT ĐỐI KHÔNG tự sinh mã \includegraphics hay \begin{center}.
   - TUYỆT ĐỐI KHÔNG xuất thẻ [HÌNH_VẼ] cho toàn bộ trang scan, tiêu đề trường/sở, phần đề bài văn bản, hoặc khi trang chỉ có chữ và công thức.
   - NẾU TRANG KHÔNG CÓ HÌNH VẼ MINH HỌA: CẤM TUYỆT ĐỐI việc sinh thẻ [HÌNH_VẼ] với các mô tả như "Không có hình", "Scan đề thi", "Toàn trang", "Đề bài". Chỉ trả về văn bản đề bài sạch.
   - CHỈ XUẤT THẺ [HÌNH_VẼ] khi có hình vẽ hình học thực sự (tam giác, hình không gian...), đồ thị hàm số, hình vẽ mô phỏng của một câu hỏi cụ thể.
   - Tọa độ `box=[ymin, xmin, ymax, xmax]` CHỈ ĐƯỢC khoanh gọn gàng đúng phần hình vẽ / đồ thị của riêng câu đó (ví dụ: chỉ khoanh hình vẽ của Câu 12), TUYỆT ĐỐI KHÔNG khoanh trùm lên chữ của đề bài, câu hỏi hay đáp án A, B, C, D và TUYỆT ĐỐI KHÔNG khoanh toàn bộ trang A4.
   - Hãy xuất thẻ kèm tọa độ bounding box trên ảnh gốc theo thang 0 đến 1000:
     [HÌNH_VẼ: box=[ymin, xmin, ymax, xmax] | mô tả ngắn]
     (Ví dụ: [HÌNH_VẼ: box=[300, 150, 680, 580] | Đồ thị hàm số bậc ba])
   - BẮT BUỘC MỞ RỘNG KHUNG BAO VÀ CHỐNG CẮT LẸM NHÁNH TRÁI:
     + Trục tung $Oy$ thường ở giữa hình: CẤM NHẦM trục $Oy$ hoặc gốc $O$ là mép trái! `xmin` BẮT BUỘC phải mở rộng hết sang trái để bao trọn toàn bộ nhánh đồ thị bên trái (nhánh âm $x < 0$), các điểm cực đại/cực tiểu bên trái, nét đứt gióng tọa độ và các số âm (như $-1, -2$) trên trục $Ox$.
     + Bao trọn cả 4 phía: đỉnh cao nhất (`ymin`), đáy thấp nhất (`ymax`), nhánh ngoài cùng bên phải (`xmax`), chữ cái $x, y$ và mũi tên hai đầu trục.
     + Thêm đệm an toàn: Nới rộng ra ngoài 15-25 đơn vị trên thang 1000 để không cắt sát mép.
7. ĐỐI VỚI BẢNG BIẾN THIÊN HOẶC BẢNG SỐ LIỆU GHÉP NHÓM:
   - ĐỐI VỚI BẢNG BIẾN THIÊN:
     + KHÔNG dùng môi trường tabular vẽ thô vì sẽ vỡ giao diện.
     + Hãy xuất thẻ kèm tọa độ bounding box trên ảnh gốc theo thang 0 đến 1000:
       [BẢNG: box=[ymin, xmin, ymax, xmax] | mô tả ngắn]
       (Ví dụ: [BẢNG: box=[180, 100, 480, 850] | Bảng biến thiên hàm số f(x)])
     + BẮT BUỘC MỞ RỘNG KHUNG BAO: `xmin` phải bao trọn cột tiêu đề ngoài cùng bên trái (chứa các ký hiệu $x$, $y'$, $y$ hoặc $f'(x)$, $f(x)$) và $-\infty$. `xmax`, `ymin`, `ymax` bao trọn toàn bộ viền bảng, các mũi tên và cực trị. Nới rộng đệm an toàn 15-20 đơn vị.
   - ĐỐI VỚI BẢNG SỐ LIỆU GHÉP NHÓM / BẢNG THỐNG KÊ TẦN SỐ:
     + Có thể xuất thẻ [BẢNG: box=[ymin, xmin, ymax, xmax] | Bảng số liệu ghép nhóm] để cắt trực tiếp từ ảnh gốc.
     + HOẶC xuất bảng Markdown đầy đủ: BẮT BUỘC mỗi hàng nằm trên MỘT DÒNG RIÊNG BIỆT, có hàng phân cách |:---|:---|, các khoảng ghép nhóm như $[145; 150)$ phải bọc trong cặp dấu $...$. TUYỆT ĐỐI KHÔNG viết tất cả các hàng dính liền trên một dòng đơn.
8. ĐỐI VỚI CÁC ĐÁP ÁN TRẮC NGHIỆM A, B, C, D:
   - BẮT BUỘC viết mỗi phương án trên MỘT DÒNG RIÊNG BIỆT (xuống dòng trước A., B., C., D.).
   - MỌI GIÁ TRỊ TOÁN HỌC TRONG ĐÁP ÁN PHẢI BỌC TRONG CẶP DẤU $...$: Kể cả khi đáp án chỉ là một con số đơn giản, số nguyên âm hay phân số (viết 'A. $0$', 'B. $1$', 'C. $-4$', 'D. $21$'). TUYỆT ĐỐI KHÔNG viết số trần như 'A. 0', 'B. 1'.
   - TUYỆT ĐỐI KHÔNG bọc chữ cái nhãn A., B., C., D. vào trong cặp dấu $...$ (viết 'A. $2x + 1$' thay vì '$A. 2x + 1$' hay '$A.$ $2x+1$').
   - TUYỆT ĐỐI KHÔNG bọc chung tất cả đáp án A, B, C, D trong cùng một cặp dấu $...$ hay $$...$$. Mỗi đáp án phải có cặp dấu $...$ riêng biệt.
9. ĐỐI VỚI CÂU HỎI ĐÚNG / SAI (PHẦN II) VÀ CÂU HỎI TỰ LUẬN (PHẦN IV):
   - BẮT BUỘC viết mỗi ý trên MỘT DÒNG RIÊNG BIỆT (xuống dòng trước a), b), c), d) hoặc a/, b/, c/, d/ hoặc 1), 2)).
   - TUYỆT ĐỐI KHÔNG viết các ý a), b), c), d) dính liền trên cùng một hàng.
   - BẮT BUỘC nhúng mọi công thức toán học và biểu thức trong từng ý vào cặp dấu $...$.
   - ĐỐI VỚI ĐIỂM SỐ TỰ LUẬN (ví dụ: 0,75 điểm, 1,0 điểm, 1,5 điểm):
     + Viết điểm số trong ngoặc đơn sau nhãn câu hỏi (ví dụ: 'Câu 1 (0,75 điểm):' hoặc 'Bài 1 (1,5 điểm):').
     + TUYỆT ĐỐI KHÔNG bọc điểm số vào trong dấu $...$ (CẤM viết 'Câu 1 $(0,75\text{ điểm})$:' hay 'Câu 1 ($0,75$ điểm):'). Điểm số luôn để dạng văn bản thường ngoài equation.
10. ĐỐI VỚI KHỐI TIÊU ĐỀ ĐỀ THI (NẾU CÓ Ở ĐẦU TRANG ĐẦU TIÊN):
   - Nhận diện khối tiêu đề thông tin trường/sở, kỳ thi, môn học, họ tên, mã đề và xuất theo thẻ cấu trúc:
     [TIÊU_ĐỀ_ĐỀ_THI:
     BÊN_TRÁI:
     SỞ GD&ĐT THÀNH PHỐ HUẾ
     TRƯỜNG THPT PHÚ LỘC
     (Đề thi gồm có 04 trang)
     BÊN_PHẢI:
     ĐỀ KIỂM TRA CUỐI KỲ I, NĂM HỌC 2025-2026
     MÔN: TOÁN 12
     Thời gian làm bài 90 phút, không kể thời gian phát đề
     THÔNG_TIN:
     Họ và tên học sinh: ......................................., Lớp: 12B... | Mã đề 132
     ]
   - Nếu đề thi không có đủ thông tin, chỉ điền những thông tin thực tế có trên đề.
   - Các trang tiếp theo (trang 2, 3...) không lặp lại tiêu đề thì KHÔNG xuất thẻ này.
11. ĐỐI VỚI TIÊU ĐỀ CÁC PHẦN (PHẦN I, PHẦN II, PHẦN III, PHẦN IV...):
   - BẮT BUỘC giữ nguyên toàn bộ câu hướng dẫn của phần trên MỘT DÒNG HOÀN CHỈNH, ví dụ:
     PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.
     PHẦN II. Câu trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.
     PHẦN III. Câu trắc nghiệm trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu 6.
     PHẦN IV. TỰ LUẬN
   - TUYỆT ĐỐI KHÔNG ngắt dòng giữa các câu chỉ dẫn đề thi.
12. KHÔNG thêm bớt nội dung, KHÔNG giải thích."""

PROMPT_PRESETS = {
    "quiz": {
        "name": "🎯 Đề thi & Trắc nghiệm THPT",
        "description": "Chuẩn hóa đề thi toán, câu hỏi, đáp án A B C D, tự động nhận diện box hình vẽ/bảng",
        "prompt": load_prompt_file("quiz_prompt.txt", DEFAULT_QUIZ_PROMPT)
    },
    "text_and_formula": {
        "name": "📖 Văn bản + Công thức (SGK / Tài liệu)",
        "description": "Giữ nguyên text sạch, chuyển công thức thành LaTeX, nhận diện box ảnh",
        "prompt": """Bạn là chuyên gia OCR tài liệu toán học tiếng Việt. Hãy chuyển đổi nội dung ảnh thành văn bản chuẩn:

1. TUYỆT ĐỐI KHÔNG dùng dấu sao ** để in đậm.
2. BẮT BUỘC XUỐNG DÒNG RÕ RÀNG TRƯỚC CÁC MỤC:
   - Các tiêu đề PHẦN (ví dụ: 'I. PHẦN TRẮC NGHIỆM', 'II. PHẦN TỰ LUẬN') BẮT BUỘC nằm trên MỘT DÒNG RIÊNG BIỆT.
   - Trước mỗi 'Câu X:' hoặc 'Bài X (điểm):' BẮT BUỘC phải xuống dòng riêng biệt, không viết dính liền nhau.
   - Các ý tự luận a), b), c) hoặc a/, b/, c/ hoặc 1), 2): BẮT BUỘC mỗi ý nằm trên một dòng riêng biệt.
3. BẮT BUỘC nhúng TẤT CẢ ký hiệu toán học vào $...$: điểm ($A, B, C, M, N$), trục ($Ox, Oy, Oz$), hệ tọa độ ($Oxyz, Oxy$), mặt phẳng ($Oxy, (P)$), tọa độ ($G(1; 2; 3)$), tham số ($m, x, y$), tập hợp ($A$, $\\mathbb{N}$, $\\mathbb{Z}$, $\\{-4; 1; 0\\}$), phép tính ($(-36) + 15$, $2^2 \\cdot 15$), quan hệ ($\\in, \\notin, \\subset$).
4. LUÔN CÓ DẤU CÁCH giữa chữ tiếng Việt và công thức toán (ví dụ: 'hình hộp $ABCD.A\\'B\\'C\\'D\\'$').
5. Mọi công thức toán viết bằng LaTeX: inline $...$, display $$...$$. TUYỆT ĐỐI KHÔNG ngắt dòng giữa $...$.
   - Vector: dùng \\vec{AB}
   - Phân số: \\frac{tử}{mẫu}, Căn: \\sqrt{x}
6. Gặp HÌNH VẼ / ĐỒ THỊ: CHỈ xuất thẻ khi có hình vẽ/đồ thị thực sự của câu hỏi, tuyệt đối KHÔNG xuất thẻ cho toàn bộ trang scan hay khi không có hình. Tọa độ box chỉ khoanh đúng vùng hình vẽ của câu đó (ví dụ hình Câu 12), viết [HÌNH_VẼ: box=[ymin, xmin, ymax, xmax] | mô tả ngắn] theo thang 0-1000. CẤM NHẦM trục Oy là mép trái, xmin bắt buộc phải mở rộng hết sang trái bao trọn toàn bộ nhánh đồ thị bên âm (x < 0), nét đứt và số âm. Nới rộng khung bao cả 4 phía và thêm đệm an toàn 15-25 đơn vị.
7. Gặp BẢNG BIẾN THIÊN / BẢNG PHỨC TẠP: viết [BẢNG: box=[ymin, xmin, ymax, xmax] | mô tả ngắn] theo thang 0-1000. xmin bắt buộc bao trọn cột tiêu đề x, y', y bên trái, bao trọn mũi tên và cực trị. Nới rộng đệm an toàn 15-20 đơn vị.
8. KHÔNG thêm lời giải thích, chỉ trả về nội dung đã chuyển đổi."""
    },
    "formula_only": {
        "name": "🧮 Chỉ công thức",
        "description": "Chỉ trích xuất công thức toán, bỏ qua text thường",
        "prompt": """Hãy trích xuất TẤT CẢ các công thức toán học trong ảnh và viết bằng LaTeX.

Quy tắc:
1. Mỗi công thức trên một dòng, dùng $$...$$ 
2. Phân số: \\frac{}{}, Căn: \\sqrt{}, Vector: \\vec{}
3. KHÔNG viết text giải thích, CHỈ viết công thức LaTeX"""
    },
    "solution": {
        "name": "📝 Bài giải chi tiết",
        "description": "Giữ cả bước giải, lời giải, đánh số bước",
        "prompt": """Bạn là chuyên gia OCR bài giải toán tiếng Việt. Chuyển đổi nội dung ảnh:

1. Giữ nguyên văn bản tiếng Việt sạch, KHÔNG dùng **
2. Luôn có khoảng cách giữa chữ và công thức toán $...$
3. Công thức inline: $...$, công thức display: $$...$$
4. Hệ phương trình: dùng \\begin{cases}...\\end{cases}
5. Gặp hình vẽ hoặc bảng biến thiên: CHỈ xuất thẻ khi có hình vẽ/đồ thị thực sự, tuyệt đối KHÔNG xuất thẻ cho cả trang scan hay khi không có hình. Xuất [HÌNH_VẼ: box=[ymin, xmin, ymax, xmax] | mô tả ngắn] hoặc [BẢNG: box=[ymin, xmin, ymax, xmax] | mô tả ngắn] theo thang 0-1000. CẤM NHẦM trục Oy là mép trái, xmin bắt buộc phải mở rộng hết sang trái bao trọn toàn bộ nhánh đồ thị bên trái (nhánh âm x < 0), nét đứt, đỉnh/đáy và số âm. Nới rộng khung bao cả 4 phía và thêm đệm an toàn 15-25 đơn vị."""
    },
    "table": {
        "name": "📊 Bảng biểu toán học",
        "description": "Bảng số liệu đơn giản",
        "prompt": """Chuyển đổi bảng biểu có công thức toán trong ảnh:
1. Nếu là bảng số liệu thông thường: dùng markdown table | Cột 1 | Cột 2 |
2. Nếu là BẢNG BIẾN THIÊN phức tạp có mũi tên: viết [BẢNG: Bảng biến thiên của hàm số] để cắt ảnh thay thế."""
    }
}


# ══════════════════════════════════════════════
# UTILITY: Parse page ranges like "1-5, 7, 10-15"
# ══════════════════════════════════════════════

def parse_page_range(page_range_str: str, max_pages: int) -> list[int]:
    """Parse page range string like '1-5, 7, 10-15' into list of 0-indexed page numbers."""
    if not page_range_str or page_range_str.strip().lower() in ("all", ""):
        return list(range(max_pages))

    pages = set()
    parts = page_range_str.replace(" ", "").split(",")
    for part in parts:
        if "-" in part:
            start, end = part.split("-", 1)
            start = max(1, int(start))
            end = min(max_pages, int(end))
            pages.update(range(start - 1, end))  # Convert to 0-indexed
        else:
            p = int(part)
            if 1 <= p <= max_pages:
                pages.add(p - 1)
    return sorted(pages)


# ══════════════════════════════════════════════
# API ENDPOINTS
# ══════════════════════════════════════════════

@app.get("/", response_class=HTMLResponse)
async def home():
    html_path = Path(__file__).parent / "templates" / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.post("/api/convert")
async def convert_pdf(
    file: UploadFile = File(...),
    format: str = Query("png", pattern="^(png|jpeg)$"),
    dpi: int = Query(300, ge=72, le=600),
    pages: str = Query("all"),
):
    """Chuyển PDF thành ảnh với tùy chọn trang."""
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Chỉ chấp nhận file PDF!")

    pdf_bytes = await file.read()
    if not pdf_bytes:
        raise HTTPException(400, "File PDF rỗng!")

    session_id = str(uuid.uuid4())
    session_dir = TEMP_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        raise HTTPException(400, "Không thể đọc file PDF!")

    # Lưu lại file PDF gốc vào session để tái sử dụng khi cần
    try:
        (session_dir / "original.pdf").write_bytes(pdf_bytes)
    except Exception:
        pass

    # 🌟 Luồng 1 (Digital PDF Bản đẹp): Tự động trích xuất trước toàn bộ ảnh gốc nét 100%
    native_summary = {}
    try:
        native_summary = extract_native_images_from_doc(doc, session_dir)
        if native_summary.get("total_images", 0) > 0:
            print(f"[NATIVE_PDF] Đã bóc tách {native_summary['total_images']} ảnh gốc nét 100% từ PDF.")
    except Exception as e:
        print(f"[NATIVE_PDF] Trích xuất ảnh gốc thất bại (sẽ fallback sang scan): {e}")

    total_pages = len(doc)
    selected_pages = parse_page_range(pages, total_pages)

    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)
    images_info = []

    for page_num in selected_pages:
        page = doc[page_num]
        pix = page.get_pixmap(matrix=matrix)
        ext = format
        filename = f"page_{page_num + 1}.{ext}"
        filepath = session_dir / filename

        if format == "jpeg":
            pix.save(str(filepath), output="jpeg", jpg_quality=95)
        else:
            pix.save(str(filepath))

        images_info.append({
            "page": page_num + 1,
            "filename": filename,
            "width": pix.width,
            "height": pix.height,
            "url": f"/api/image/{session_id}/{filename}",
        })

    doc.close()

    return {
        "session_id": session_id,
        "total_pages": total_pages,
        "selected_pages": len(selected_pages),
        "format": format,
        "dpi": dpi,
        "images": images_info,
    }


@app.post("/api/upload-images")
async def upload_images(
    files: list[UploadFile] = File(...),
    session_id: str = Form(None),
):
    """Upload hoặc dán ảnh trực tiếp (JPG/PNG/Clipboard)."""
    if not session_id or not (TEMP_DIR / session_id).exists():
        session_id = str(uuid.uuid4())
        session_dir = TEMP_DIR / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        existing_count = 0
    else:
        session_dir = TEMP_DIR / session_id
        existing_count = len(list(session_dir.iterdir()))

    images_info = []
    for i, file in enumerate(files):
        ext = file.filename.split(".")[-1].lower() if "." in file.filename else "png"
        if ext not in ("png", "jpg", "jpeg", "bmp", "webp"):
            ext = "png"

        img_bytes = await file.read()
        filename = f"image_{existing_count + i + 1}.{ext}"
        filepath = session_dir / filename
        filepath.write_bytes(img_bytes)

        img = Image.open(io.BytesIO(img_bytes))
        images_info.append({
            "page": existing_count + i + 1,
            "filename": filename,
            "width": img.width,
            "height": img.height,
            "url": f"/api/image/{session_id}/{filename}",
        })

    if not images_info:
        raise HTTPException(400, "Không có ảnh hợp lệ!")

    return {
        "session_id": session_id,
        "images": images_info,
    }


@app.get("/api/image/{session_id}/{filename}")
async def get_image(session_id: str, filename: str):
    filepath = TEMP_DIR / session_id / filename
    if not filepath.exists():
        raise HTTPException(404, "Ảnh không tồn tại!")
    media_type = "image/png" if filename.endswith(".png") else "image/jpeg"
    return FileResponse(str(filepath), media_type=media_type, filename=filename)


def smart_trim_figure(img: Image.Image, thresh: int = 240, margin: int = 5, is_table: bool = False) -> Image.Image:
    """
    Tự động tìm vùng bao nội dung thực tế:
    - Với mọi thẻ [HÌNH_VẼ] (is_table=False):
      Phục hồi 100% thuật toán trim gốc: chỉ xén khoảng trắng ngoài cùng 4 phía và lề 5px,
      tuyệt đối KHÔNG cắt vào các đỉnh A, B, C, D, A', B', C', D' của hình không gian hay trục tọa độ của đồ thị.
    - Với thẻ [BẢNG] (is_table=True):
      Nhận diện đường viền kẻ ngang thực sự của bảng (đoạn kẻ đen liên tục dài nhất >= 30% W).
      Cắt thẳng từ viền trên first_grid_y trở xuống, loại bỏ sạch 100% dòng chữ đề bài phía trên mép bảng.
      Cắt thẳng tại viền dưới last_grid_y, loại bỏ sạch 100% phương án phía dưới mép bảng.
    """
    if img.mode != 'RGB':
        img = img.convert('RGB')

    gray = img.convert('L')
    W, H = img.size

    # 1. Tính mật độ điểm tối và độ dài đoạn nét đen liên tục trên từng hàng
    row_dark_counts = []
    row_max_runs = []
    for y in range(H):
        row_pixels = [gray.getpixel((x, y)) < thresh for x in range(W)]
        dark_cnt = sum(row_pixels)
        row_dark_counts.append(dark_cnt)

        max_run = 0
        curr_run = 0
        for val in row_pixels:
            if val:
                curr_run += 1
                if curr_run > max_run:
                    max_run = curr_run
            else:
                curr_run = 0
        row_max_runs.append(max_run)

    min_content_dark = max(2, int(W * 0.005))

    start_y = 0
    end_y = H - 1
    margin_top = margin
    margin_bottom = margin

    # CHỈ kích hoạt chế độ cắt viền bảng KHI VÀ CHỈ KHI is_table is True (từ thẻ [BẢNG: ...])
    # Tuyệt đối KHÔNG tự đoán bảng bằng len(grid_lines) >= 2 để tránh cắt cụt hình hộp 3D và đồ thị
    if is_table:
        # Đường kẻ ngang thực sự của bảng: đoạn kẻ đen liên tục >= 30% W (chữ đề bài không bao giờ đạt mức này)
        grid_lines = [y for y in range(H) if row_max_runs[y] >= W * 0.30]

        if len(grid_lines) >= 2:
            first_grid_y = grid_lines[0]
            last_grid_y = grid_lines[-1]

            # Kiểm tra xem phía trên first_grid_y có chữ đề bài không
            has_dark_above = any(row_dark_counts[k] >= min_content_dark for k in range(first_grid_y))
            if has_dark_above:
                # Cắt thẳng từ viền trên first_grid_y trở xuống, lề trên = 0 để loại sạch 100% chữ đề bài
                start_y = first_grid_y
                margin_top = 0
            else:
                start_y = first_grid_y
                margin_top = min(margin, first_grid_y)

            # Kiểm tra xem phía dưới last_grid_y có chữ đáp án không
            has_dark_below = any(row_dark_counts[k] >= min_content_dark for k in range(last_grid_y + 1, H))
            if has_dark_below:
                # Cắt thẳng tại viền dưới last_grid_y, lề dưới = 0 để loại sạch 100% chữ đáp án
                end_y = last_grid_y
                margin_bottom = 0
            else:
                end_y = last_grid_y
                margin_bottom = min(margin, H - 1 - last_grid_y)
        else:
            for y in range(H):
                if row_dark_counts[y] >= min_content_dark:
                    start_y = y
                    break
            for y in range(H - 1, -1, -1):
                if row_dark_counts[y] >= min_content_dark:
                    end_y = y
                    break
    else:
        # THUẬT TOÁN TRIM GỐC CHO [HÌNH_VẼ] / ĐỒ THỊ / HÌNH HỌC KHÔNG GIAN:
        # Chỉ xén khoảng trắng ngoài cùng 4 phía và lề 5px, tuyệt đối KHÔNG cắt vào các đỉnh A, B, C, D, A', B', C', D'
        for y in range(H):
            if row_dark_counts[y] >= min_content_dark:
                start_y = y
                break

        for y in range(H - 1, -1, -1):
            if row_dark_counts[y] >= min_content_dark:
                end_y = y
                break

    if end_y <= start_y:
        start_y = 0
        end_y = H - 1

    # 2. Quét cột trái và cột phải trong khoảng [start_y, end_y]
    col_dark_counts = []
    for x in range(W):
        dark_cnt = sum(1 for y in range(start_y, end_y + 1) if gray.getpixel((x, y)) < thresh)
        col_dark_counts.append(dark_cnt)

    start_x = 0
    for x in range(W):
        if col_dark_counts[x] >= 2:
            start_x = x
            break

    end_x = W - 1
    for x in range(W - 1, -1, -1):
        if col_dark_counts[x] >= 2:
            end_x = x
            break

    if end_x <= start_x:
        start_x = 0
        end_x = W - 1

    # Thêm lề an toàn
    c_left = max(0, start_x - margin)
    c_top = max(0, start_y - margin_top)
    c_right = min(W, end_x + 1 + margin)
    c_bottom = min(H, end_y + 1 + margin_bottom)

    if (c_right - c_left >= 20) and (c_bottom - c_top >= 20):
        return img.crop((c_left, c_top, c_right, c_bottom))
    return img


@app.post("/api/crop")
async def crop_image(
    session_id: str = Form(...),
    filename: str = Form(...),
    x: int = Form(...),
    y: int = Form(...),
    width: int = Form(...),
    height: int = Form(...),
):
    """Cắt vùng trên ảnh, tự động trim khoảng trắng thừa và làm sạch viền chữ dính."""
    filepath = TEMP_DIR / session_id / filename
    if not filepath.exists():
        raise HTTPException(404, "Ảnh không tồn tại!")

    img = Image.open(filepath)
    cropped = img.crop((x, y, x + width, y + height))
    cropped = smart_trim_figure(cropped, thresh=240, margin=4)

    # Save cropped version
    crop_filename = f"crop_{uuid.uuid4().hex[:8]}.png"
    crop_path = TEMP_DIR / session_id / crop_filename
    cropped.save(str(crop_path))

    return {
        "filename": crop_filename,
        "width": cropped.width,
        "height": cropped.height,
        "url": f"/api/image/{session_id}/{crop_filename}",
    }


@app.get("/api/download-figures/{session_id}")
async def download_all_figures(session_id: str):
    """Nén và tải về toàn bộ hình ảnh/đồ thị/bảng đã cắt trong session dưới dạng file ZIP."""
    session_dir = TEMP_DIR / session_id
    if not session_dir.exists():
        raise HTTPException(404, "Phiên làm việc không tồn tại!")

    crop_files = sorted(
        list(session_dir.glob("autocrop_*.png")) + list(session_dir.glob("crop_*.png")),
        key=lambda f: f.stat().st_mtime
    )
    if not crop_files:
        raise HTTPException(404, "Chưa có hình ảnh nào được cắt trong phiên làm việc này!")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for idx, f in enumerate(crop_files, 1):
            zip_file.write(f, arcname=f"Hinh_{idx}_{f.name}")

    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=hinh_anh_{session_id[:8]}.zip"}
    )


@app.get("/api/figures/{session_id}")
async def get_all_figures(session_id: str):
    """Lấy danh sách thông tin tất cả hình ảnh đã cắt trong session."""
    session_dir = TEMP_DIR / session_id
    if not session_dir.exists():
        return {"session_id": session_id, "figures": []}

    crop_files = sorted(
        list(session_dir.glob("autocrop_*.png")) + list(session_dir.glob("crop_*.png")),
        key=lambda f: f.stat().st_mtime
    )
    figures = []
    for f in crop_files:
        try:
            with Image.open(f) as im:
                w, h = im.size
        except Exception:
            w, h = 0, 0
        figures.append({
            "filename": f.name,
            "url": f"/api/image/{session_id}/{f.name}",
            "width": w,
            "height": h,
            "size_bytes": f.stat().st_size
        })
    return {"session_id": session_id, "figures": figures}


@app.get("/api/prompts")
async def get_prompts():
    """Lấy danh sách prompt presets."""
    return {
        key: {"name": v["name"], "description": v["description"], "prompt": v["prompt"]}
        for key, v in PROMPT_PRESETS.items()
    }


@app.get("/api/presets")
async def get_presets_alias():
    """Alias cho /api/prompts."""
    return await get_prompts()


@app.get("/api/models")
async def get_models_alias():
    """Danh sách các model khả dụng và cấu hình pool keys."""
    return {
        "models": AVAILABLE_FALLBACKS,
        "default_keys_count": len(DEFAULT_API_KEYS)
    }


@app.get("/api/config-keys")
async def get_server_keys():
    """Trả về số lượng và danh sách key từ server (dành cho client đồng bộ tự động)."""
    return {
        "count": len(DEFAULT_API_KEYS),
        "keys": DEFAULT_API_KEYS
    }


def parse_tag_content(content: str):
    """Trích xuất tọa độ bounding box, tên file ảnh và mô tả từ nội dung bên trong thẻ (hỗ trợ cả có và không có question=)."""
    box = None
    desc = content

    # 0. Trích xuất question=label trước nếu có (hỗ trợ cả có và không có question=)
    question_label = None
    m_q = re.search(r'question\s*=\s*["\']([^"\']+)["\']', desc, re.IGNORECASE)
    if not m_q:
        m_q = re.search(r'question\s*=\s*([^\s\|\]]+)', desc, re.IGNORECASE)
    if m_q:
        question_label = m_q.group(1).strip()
        desc = desc[:m_q.start()] + desc[m_q.end():]

    desc = re.sub(r'\|+', '|', desc).strip()

    # 1. Kiểm tra box=[ymin, xmin, ymax, xmax] hoặc box: [...]
    m_box = re.search(r'box\s*[:=]\s*\[?\s*(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)\s*\]?', desc, re.IGNORECASE)
    if m_box:
        box = [float(m_box.group(i)) for i in range(1, 5)]
        desc = desc[:m_box.start()] + desc[m_box.end():]

    # 2. Kiểm tra [ymin, xmin, ymax, xmax] độc lập
    if not box:
        m_box2 = re.search(r'\[\s*(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)[,\s]+(\d+(?:\.\d+)?)\s*\]', desc)
        if m_box2:
            box = [float(m_box2.group(i)) for i in range(1, 5)]
            desc = desc[:m_box2.start()] + desc[m_box2.end():]

    # 3. Kiểm tra ymin=... xmin=...
    if not box:
        m_box3 = re.search(r'ymin\s*=\s*(\d+(?:\.\d+)?).*?xmin\s*=\s*(\d+(?:\.\d+)?).*?ymax\s*=\s*(\d+(?:\.\d+)?).*?xmax\s*=\s*(\d+(?:\.\d+)?)', desc, re.IGNORECASE)
        if m_box3:
            box = [float(m_box3.group(i)) for i in range(1, 5)]
            desc = desc[:m_box3.start()] + desc[m_box3.end():]

    # 4. Kiểm tra image=filename (hoặc filename ảnh độc lập)
    image_name = None
    m_img = re.search(r'image\s*=\s*["\']?([^\s\|\]\'"]+)["\']?', desc, re.IGNORECASE)
    if m_img:
        raw_img = m_img.group(1).strip()
        desc = desc[:m_img.start()] + desc[m_img.end():]
        raw_img = re.sub(r'question\s*=\s*(?:"[^"]*"|\'[^\']*\'|[^\s\|\]]+)\s*\|*', '', raw_img, flags=re.IGNORECASE)
        m_file = re.search(r'((?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg)|[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg))', raw_img, re.IGNORECASE)
        if m_file:
            image_name = m_file.group(1)
        else:
            image_name = re.sub(r'^[\|\s,;:\'"]+|[\|\s,;:\'"]+$', '', raw_img).strip()

    if not image_name:
        m_file_direct = re.search(r'\b((?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg)|[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg))\b', desc, re.IGNORECASE)
        if m_file_direct:
            image_name = m_file_direct.group(1)
            desc = desc[:m_file_direct.start()] + desc[m_file_direct.end():]

    desc = re.sub(r'\|+', '|', desc)
    desc = re.sub(r'^[\|\s,;:]+|[\|\s,;:]+$', '', desc).strip()
    return box, image_name, desc, question_label


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
    tag_pattern = r'\[((?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|IMAGE|FIGURE|TABLE)):\s*((?:[^\[\]]|\[[^\]]*\])*)\]'
    def _clean_fake(m):
        content = m.group(2).strip()
        box, _, desc, _ = parse_tag_content(content)
        if is_invalid_or_fake_figure(box, desc):
            return ""
        return m.group(0)
    return re.sub(tag_pattern, _clean_fake, text, flags=re.IGNORECASE)


def auto_crop_figures(session_id: str, orig_filename: str, latex_text: str) -> tuple[str, list[dict]]:
    """
    Tự động tìm kiếm các thẻ [HÌNH_VẼ: box=[ymin, xmin, ymax, xmax] | ...]
    hoặc [BẢNG: box=[...] | ...] để cắt từ ảnh gốc và cập nhật lại thẻ.
    Hỗ trợ cả hai định dạng thẻ: có question="..." và không có question=.
    Trích xuất ngữ cảnh câu hỏi (Câu 1, Câu 3...), số trang để hiển thị rõ ràng cho người dùng.
    """
    if not latex_text:
        return latex_text, []

    # Lọc bỏ chân trang scan
    latex_text = remove_page_footers(latex_text)
    # Nối lại các từ bị ngắt sai dòng trước khi xử lý
    latex_text = join_broken_ocr_lines(latex_text)
    # Lọc bỏ trước các thẻ hình vẽ giả mạo / chụp toàn trang / không có hình
    latex_text = remove_fake_figure_tags(latex_text)

    orig_path = TEMP_DIR / session_id / orig_filename
    if not orig_path.exists():
        return latex_text, []

    # Trích xuất số trang
    m_pg = re.search(r'page_(\d+)', orig_filename)
    page_num = int(m_pg.group(1)) if m_pg else 1

    figures = []
    # Khớp mọi biến thể thẻ có hỗ trợ lồng ngoặc vuông [...]
    tag_pattern = r'\[((?:HÌNH[ _]VẼ|HÌNH|ĐỒ[ _]THỊ|BẢNG(?:[ _]BIẾN[ _]THIÊN)?|IMAGE|FIGURE|TABLE)):\s*((?:[^\[\]]|\[[^\]]*\])*)\]'

    try:
        with Image.open(orig_path) as im:
            W, H = im.size

            def replace_match(match):
                raw_tag = match.group(1).upper()
                content = match.group(2).strip()

                tag_name = "BẢNG" if any(k in raw_tag for k in ("BẢNG", "TABLE")) else "HÌNH_VẼ"
                box, existing_img, desc, q_label_in_tag = parse_tag_content(content)

                # Lọc bỏ ngay nếu là thẻ giả mạo / chụp toàn trang / không có hình vẽ
                if is_invalid_or_fake_figure(box, desc):
                    return ""

                if not desc:
                    desc = "Bảng biến thiên / Bảng số liệu" if tag_name == "BẢNG" else "Hình vẽ / Đồ thị"

                # Trích xuất câu hỏi gần nhất phía trước (hoặc phía sau nếu thẻ nằm đầu câu)
                preceding = latex_text[:match.start()]
                q_matches = list(re.finditer(r'(?i)\b((?:Câu|Bài|Đề|Ví dụ)\s+\d+[\.:]?)(.*?)(?=(?:Câu|Bài|Đề|Ví dụ)\s+\d+[\.:]|$)', preceding, re.DOTALL))
                if q_label_in_tag:
                    q_label = q_label_in_tag
                    q_snippet = ""
                elif q_matches:
                    last_q = q_matches[-1]
                    q_label = last_q.group(1).strip()
                    q_snippet = re.sub(r'\s+', ' ', last_q.group(2).strip())[:90]
                else:
                    following = latex_text[match.end():]
                    m_next_q = re.search(r'(?i)\b((?:Câu|Bài|Đề|Ví dụ)\s+\d+[\.:]?)', following)
                    if m_next_q:
                        q_label = m_next_q.group(1).strip()
                    else:
                        q_label = f"Hình {len(figures) + 1}"
                    q_snippet = ""

                # 1. Đã có ảnh cắt trước đó -> giữ nguyên
                if existing_img:
                    clean_existing = re.sub(r'question\s*=\s*(?:"[^"]*"|\'[^\']*\'|[^\s\|\]]+)\s*\|*', '', existing_img, flags=re.IGNORECASE)
                    m_f = re.search(r'((?:auto)?crop_[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg)|[a-zA-Z0-9_\-]+\.(?:png|jpg|jpeg))', clean_existing, re.IGNORECASE)
                    if m_f:
                        clean_existing = m_f.group(1)
                    else:
                        clean_existing = re.sub(r'^[\|\s,;:\'"]+|[\|\s,;:\'"]+$', '', clean_existing).strip()

                    fig_info = {
                        "index": len(figures) + 1,
                        "tag": tag_name,
                        "image": clean_existing,
                        "url": f"/api/image/{session_id}/{clean_existing}",
                        "desc": desc,
                        "box": [int(x) for x in box] if box else None,
                        "orig_filename": orig_filename,
                        "page_num": page_num,
                        "question_label": q_label,
                        "context_text": q_snippet,
                        "needs_review": True
                    }
                    figures.append(fig_info)
                    box_str = f" | box=[{int(box[0])}, {int(box[1])}, {int(box[2])}, {int(box[3])}]" if box else ""
                    return f"[{tag_name}: question=\"{q_label}\" | image={clean_existing}{box_str} | {desc}]"

                # 2. Có tọa độ box -> tiến hành cắt ảnh tự động
                if box:
                    ymin, xmin, ymax, xmax = box
                    box_int = [int(ymin), int(xmin), int(ymax), int(xmax)]

                    # 🌟 LUỒNG 1: Ưu tiên số 1 - Kiểm tra xem có ảnh gốc nhúng (Native PDF Image) khớp với câu hỏi không
                    matched_native = find_matching_native_image(TEMP_DIR / session_id, page_num, box_int)
                    if matched_native:
                        native_src = Path(matched_native["full_path"])
                        crop_fname = f"autocrop_{uuid.uuid4().hex[:8]}.png"
                        crop_path = TEMP_DIR / session_id / crop_fname
                        try:
                            with Image.open(native_src) as native_im:
                                # Quy tắc 20: Chốt chặn an toàn - nếu ảnh gốc bị đen xì (>70% pixel đen hoặc đơn sắc)
                                # do mất smask hoặc lỗi XObject, lập tức hủy bỏ và fallback 100% sang smart_trim_figure
                                if is_black_or_corrupted_image(native_im):
                                    print(f"[NATIVE_PDF] CẢNH BÁO: Ảnh gốc {matched_native['filename']} cho {q_label} bị đen xì (>70% pixel đen hoặc đơn sắc hỏng), tự động fallback sang smart_trim_figure.")
                                    matched_native = None
                                else:
                                    native_im.save(crop_path, "PNG")

                            if matched_native:
                                is_vector = matched_native.get("origin") == "vector"
                                desc_label = f"{desc} (Hình vẽ vector PDF)" if is_vector else f"{desc} (Ảnh gốc PDF)"

                                fig_info = {
                                    "index": len(figures) + 1,
                                    "tag": tag_name,
                                    "image": crop_fname,
                                    "url": f"/api/image/{session_id}/{crop_fname}",
                                    "desc": desc_label,
                                    "box": box_int,
                                    "pixel_box": None,
                                    "orig_filename": orig_filename,
                                    "page_num": page_num,
                                    "question_label": q_label,
                                    "context_text": q_snippet,
                                    "is_native": True,
                                    "origin": matched_native.get("origin", "bitmap"),
                                    "needs_review": False
                                }
                                figures.append(fig_info)
                                print(f"[NATIVE_PDF] Khớp thành công {'hình vẽ vector' if is_vector else 'ảnh gốc'} {matched_native['filename']} cho {q_label} (điểm={matched_native.get('score')})")
                                return f"[{tag_name}: question=\"{q_label}\" | image={crop_fname} | box=[{box_int[0]}, {box_int[1]}, {box_int[2]}, {box_int[3]}] | {desc}]"
                        except Exception as e:
                            print(f"[NATIVE_PDF] Lỗi sao chép ảnh gốc {native_src}: {e}, fallback sang smart_trim")
                            matched_native = None

                    # 🌟 LUỒNG 2: Fallback sang cắt ảnh từ trang scan/render (smart_trim_figure)
                    # Chuẩn hóa tọa độ sang pixel
                    if ymax <= 1.0 and xmax <= 1.0:
                        top = int(ymin * H)
                        left = int(xmin * W)
                        bottom = int(ymax * H)
                        right = int(xmax * W)
                    elif ymax <= 1000 and xmax <= 1000:
                        top = int((ymin / 1000.0) * H)
                        left = int((xmin / 1000.0) * W)
                        bottom = int((ymax / 1000.0) * H)
                        right = int((xmax / 1000.0) * W)
                    else:
                        top = int(ymin)
                        left = int(xmin)
                        bottom = int(ymax)
                        right = int(xmax)

                    # Thêm margin 5px (trong khoảng 4px-6px)
                    margin = 5
                    top = max(0, top - margin)
                    left = max(0, left - margin)
                    bottom = min(H, bottom + margin)
                    right = min(W, right + margin)

                    if right > left + 15 and bottom > top + 15:
                        cropped = im.crop((left, top, right, bottom))
                        # Tự động trim khoảng trắng thừa và làm sạch viền chữ dính (sát viền nếu là Bảng)
                        cropped = smart_trim_figure(cropped, thresh=240, margin=5, is_table=(tag_name == "BẢNG"))

                        crop_fname = f"autocrop_{uuid.uuid4().hex[:8]}.png"
                        crop_path = TEMP_DIR / session_id / crop_fname
                        cropped.save(crop_path)

                        box_int = [int(ymin), int(xmin), int(ymax), int(xmax)]
                        fig_info = {
                            "index": len(figures) + 1,
                            "tag": tag_name,
                            "image": crop_fname,
                            "url": f"/api/image/{session_id}/{crop_fname}",
                            "desc": desc,
                            "box": box_int,
                            "pixel_box": [left, top, right, bottom],
                            "orig_filename": orig_filename,
                            "page_num": page_num,
                            "question_label": q_label,
                            "context_text": q_snippet,
                            "needs_review": True
                        }
                        figures.append(fig_info)

                        return f"[{tag_name}: question=\"{q_label}\" | image={crop_fname} | box=[{box_int[0]}, {box_int[1]}, {box_int[2]}, {box_int[3]}] | {desc}]"

                # 3. Không có tọa độ -> chuẩn hóa thẻ placeholder để frontend hỗ trợ cắt thủ công 1-click
                fig_info = {
                    "index": len(figures) + 1,
                    "tag": tag_name,
                    "image": None,
                    "url": None,
                    "desc": desc,
                    "box": None,
                    "orig_filename": orig_filename,
                    "page_num": page_num,
                    "question_label": q_label,
                    "context_text": q_snippet,
                    "needs_review": True
                }
                figures.append(fig_info)
                return f"[{tag_name}: question=\"{q_label}\" | {desc}]"

            updated_text = re.sub(tag_pattern, replace_match, latex_text, flags=re.IGNORECASE)
            return updated_text, figures
    except Exception as e:
        print(f"Lỗi auto crop: {e}")
        return latex_text, figures
    except Exception as e:
        print(f"Lỗi auto crop: {e}")
        return latex_text, figures


def prepare_image_for_gemini(img: Image.Image, max_dim: int = 1800, quality: int = 88) -> Image.Image:
    """
    Tối ưu hóa ảnh gửi lên Gemini API:
    - Resize giữ nguyên tỉ lệ nếu cạnh lớn nhất vượt quá max_dim (mặc định 1800px).
    - Giảm dung lượng từ 5-6MB xuống ~300-400KB giúp upload và xử lý siêu tốc.
    - Lưu ý: Thao tác này chỉ dùng cho payload gửi AI, ảnh gốc trên đĩa vẫn giữ nguyên để cắt crop nét 100%.
    """
    try:
        w, h = img.size
        if max(w, h) > max_dim:
            scale = max_dim / max(w, h)
            new_w = max(1, int(w * scale))
            new_h = max(1, int(h * scale))
            img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        else:
            img_resized = img

        if img_resized.mode != 'RGB':
            img_resized = img_resized.convert('RGB')

        buf = io.BytesIO()
        img_resized.save(buf, format='JPEG', quality=quality, optimize=True)
        buf.seek(0)
        return Image.open(buf)
    except Exception as e:
        print(f"[Optimize Image Warning] Không thể tối ưu ảnh, dùng ảnh gốc: {e}")
        return img


DEPRECATED_MODEL_MAP = {
    "gemini-1.5-flash": "gemini-3.1-flash-lite",
    "gemini-1.5-flash-8b": "gemini-3.1-flash-lite",
    "gemini-1.5-pro": "gemini-3.5-flash-lite",
    "gemini-2.0-flash": "gemini-3.5-flash-lite",
    "gemini-2.0-flash-lite": "gemini-3.1-flash-lite",
    "gemini-2.5-flash": "gemini-3.5-flash-lite",
    "gemini-2.5-flash-lite": "gemini-3.1-flash-lite",
    "gemini-flash-latest": "gemini-3.1-flash-lite",
    "gemini-pro-latest": "gemini-3.5-flash-lite",
}

# Danh sách fallback ưu tiên độ ổn định và phản hồi siêu nhanh (~2-4s)
AVAILABLE_FALLBACKS = [
    "gemini-3.1-flash-lite",
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
    "gemini-3.5-flash",
]


async def _generate_content_with_retry(img, prompt_text: str, model_name: str = "auto", api_keys: list = None, timeout_seconds: float = 16.0) -> tuple:
    """
    Gọi Gemini với cơ chế bất đồng bộ non-blocking (asyncio.to_thread),
    tự động nén tối ưu ảnh, áp dụng Timeout Fail-Fast và fallback sang model/key khác khi gặp lỗi.
    """
    keys_to_use = [k for k in (api_keys or DEFAULT_API_KEYS) if k]
    if not keys_to_use:
        raise RuntimeError("Không có API Key nào khả dụng! Vui lòng kiểm tra cấu hình.")

    # Tối ưu kích thước và dung lượng ảnh gửi lên Gemini nếu có ảnh
    optimized_img = prepare_image_for_gemini(img, max_dim=1800, quality=88) if img is not None else None
    gemini_contents = [optimized_img, prompt_text] if optimized_img is not None else [prompt_text]

    clean_model = DEPRECATED_MODEL_MAP.get(model_name, model_name)
    if not clean_model or clean_model == "auto":
        models_to_try = list(AVAILABLE_FALLBACKS)
    else:
        models_to_try = [clean_model]
        for fb in AVAILABLE_FALLBACKS:
            if fb not in models_to_try:
                models_to_try.append(fb)

    last_err = None
    # Thử qua từng API key (nếu key 1 hết quota hoặc lỗi, tự động chuyển key 2)
    for key_idx, key in enumerate(keys_to_use):
        client = get_gemini_client(key)
        if not client:
            continue

        for current_model in models_to_try:
            try:
                print(f"[Gemini OCR] Đang gọi model: {current_model} (Key #{key_idx + 1}: ...{key[-6:]}, timeout={timeout_seconds}s)...")
                # Bọc lệnh gọi qua asyncio.wait_for với timeout để chống treo vô hạn khi model quá tải
                response = await asyncio.wait_for(
                    asyncio.to_thread(
                        client.models.generate_content,
                        model=current_model,
                        contents=gemini_contents
                    ),
                    timeout=timeout_seconds
                )
                if response and response.text:
                    if current_model != model_name and model_name != "auto":
                        print(f"[Gemini OCR] Tự động chuyển đổi thành công từ '{model_name}' sang '{current_model}'!")
                    else:
                        print(f"[Gemini OCR] Thành công với model '{current_model}' (Key #{key_idx + 1})")
                    return response.text, current_model
                return "", current_model
            except asyncio.TimeoutError:
                last_err = f"Model {current_model} phản hồi quá thời gian quy định ({timeout_seconds}s)"
                print(f"[Gemini OCR Warning] Model {current_model} bị timeout ({timeout_seconds}s), lập tức chuyển model tiếp theo...")
                continue
            except Exception as e:
                last_err = e
                err_str = str(e)
                print(f"[Gemini OCR Warning] Key #{key_idx + 1} - Model {current_model} lỗi ({err_str[:90]})...")
                # Nếu bị Rate Limit (429) hoặc Quota Exceeded, ngưng model này trên key này và thử key tiếp theo
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str or "quota" in err_str.lower():
                    print(f"[Gemini OCR] Key #{key_idx + 1} hết quota/bị hạn chế, chuyển sang Key tiếp theo...")
                    break
                # Nếu model 404/NOT_FOUND hoặc 503 OVERLOADED thì chuyển model tiếp theo trong cùng key
                continue

    err_msg = str(last_err)
    raise RuntimeError(f"Tất cả các API Key và Model AI đều không phản hồi: {err_msg}")


@app.post("/api/ocr")
async def ocr_image(
    session_id: str = Form(...),
    filename: str = Form(...),
    preset: str = Form("quiz"),
    custom_prompt: str = Form(None),
    model_name: str = Form("auto"),
    api_key: str = Form(None),
    api_keys: str = Form(None),
):
    # Tổng hợp danh sách keys từ request hoặc từ default
    active_keys = []
    if api_key and api_key.strip():
        active_keys.append(api_key.strip())
    if api_keys:
        try:
            parsed = json.loads(api_keys) if api_keys.startswith('[') else re.split(r'[,;\n\r]+', api_keys)
            for k in parsed:
                k = str(k).strip()
                if k and k not in active_keys:
                    active_keys.append(k)
        except Exception:
            pass
    if not active_keys:
        active_keys = DEFAULT_API_KEYS

    if not active_keys:
        raise HTTPException(400, "Chưa cấu hình Gemini API Key! Hãy nhập vào ô Quản lý API Key hoặc file .env")

    filepath = TEMP_DIR / session_id / filename
    if not filepath.exists():
        raise HTTPException(404, "Ảnh không tồn tại!")

    # Chọn prompt
    if custom_prompt:
        prompt_text = custom_prompt
    elif preset in PROMPT_PRESETS:
        prompt_text = PROMPT_PRESETS[preset]["prompt"]
    else:
        prompt_text = PROMPT_PRESETS["quiz"]["prompt"]

    try:
        img = Image.open(filepath)
        raw_text, used_model = await _generate_content_with_retry(img, prompt_text, model_name, active_keys)
        # Làm sạch toàn diện: lọc footer chân trang, gỡ $...$ ôm tiếng Việt & hàn gắn toán / tọa độ gãy dòng
        raw_text = remove_page_footers(raw_text)
        raw_text = unwrap_vietnamese_math(raw_text)
        raw_text = heal_broken_math_and_coords(raw_text)
        # Nối từ bị ngắt dòng sai từ OCR
        raw_text = join_broken_ocr_lines(raw_text)
        # Tách dòng độc lập cho các ý Đúng/Sai a,b,c,d, ý tự luận a-h và 1,2,3
        raw_text = split_sub_items_to_lines(raw_text)
        # Tự động cắt hình ảnh/bảng biến thiên từ ảnh gốc
        updated_text, figures = auto_crop_figures(session_id, filename, raw_text)
    except Exception as e:
        raise HTTPException(500, f"Lỗi Gemini API ({model_name}): {str(e)}")

    return {
        "filename": filename,
        "preset": preset,
        "latex": updated_text,
        "figures": figures,
        "used_model": used_model
    }


@app.post("/api/ocr-batch")
async def ocr_batch(
    session_id: str = Form(...),
    filenames: str = Form(...),
    preset: str = Form("quiz"),
    custom_prompt: str = Form(None),
    model_name: str = Form("auto"),
    api_keys: str = Form(None),
):
    """OCR nhiều ảnh cùng lúc và tự động crop hình/bảng."""
    active_keys = []
    if api_keys:
        try:
            parsed = json.loads(api_keys) if api_keys.startswith('[') else re.split(r'[,;\n\r]+', api_keys)
            for k in parsed:
                k = str(k).strip()
                if k and k not in active_keys:
                    active_keys.append(k)
        except Exception:
            pass
    if not active_keys:
        active_keys = DEFAULT_API_KEYS

    if not active_keys:
        raise HTTPException(400, "Chưa cấu hình Gemini API Key!")

    file_list = json.loads(filenames)
    results = []

    if custom_prompt:
        prompt_text = custom_prompt
    elif preset in PROMPT_PRESETS:
        prompt_text = PROMPT_PRESETS[preset]["prompt"]
    else:
        prompt_text = PROMPT_PRESETS["quiz"]["prompt"]

    for fname in file_list:
        filepath = TEMP_DIR / session_id / fname
        if not filepath.exists():
            results.append({"filename": fname, "error": "File not found"})
            continue

        try:
            img = Image.open(filepath)
            raw_text, used_model = await _generate_content_with_retry(img, prompt_text, model_name, active_keys)
            # Làm sạch toàn diện: lọc footer chân trang, gỡ $...$ ôm tiếng Việt & hàn gắn toán / tọa độ gãy dòng
            raw_text = remove_page_footers(raw_text)
            raw_text = unwrap_vietnamese_math(raw_text)
            raw_text = heal_broken_math_and_coords(raw_text)
            # Nối từ bị ngắt dòng sai từ OCR
            raw_text = join_broken_ocr_lines(raw_text)
            # Tách dòng độc lập cho các ý Đúng/Sai a,b,c,d, ý tự luận a-h và 1,2,3
            raw_text = split_sub_items_to_lines(raw_text)
            updated_text, figures = auto_crop_figures(session_id, fname, raw_text)
            results.append({
                "filename": fname,
                "latex": updated_text,
                "figures": figures,
                "used_model": used_model
            })
        except Exception as e:
            results.append({"filename": fname, "error": str(e)})

        await asyncio.sleep(0.1)

    return {"session_id": session_id, "results": results}


def clean_exam_metadata_and_headers(text: str) -> str:
    """
    Tự động nhận diện và gọt bỏ các thông tin thừa thãi trong file đề thi Word:
    - Khối tiêu đề đầu đề thi: Sở/Phòng GD&ĐT, Trường THPT, Kỳ thi, Năm học, Họ tên thí sinh, SBD, Mã đề...
    - Hướng dẫn thi: Thời gian làm bài, Cán bộ coi thi không giải thích gì thêm, Đề gồm có XX trang...
    - Dấu kết thúc và số trang: Trang 1/4, --- HẾT ---
    """
    if not text:
        return ""

    # KHIÊN BẢO VỆ MEDIA: Không chạm vào các thẻ ảnh Base64
    text, media_tokens = mask_media_tags(text)

    # 1. Xóa các dòng số trang ở chân trang (Trang 1/4, Trang 2 / 5, Trang 3...)
    text = re.sub(r'(?mi)^\s*Trang\s+\d+\s*(?:/\s*\d+)?\s*$', '', text)

    # 2. Xóa các dòng kết thúc đề thi (--- HẾT ---, HẾT, Cán bộ coi thi...)
    text = re.sub(r'(?mi)^\s*[-—–=_*]{2,}\s*HẾT\s*[-—–=_*]{2,}\s*$', '', text)
    text = re.sub(r'(?mi)^\s*(?:Cán bộ coi thi|Giám thị|Thí sinh)\s+không\s+(?:giải thích|được sử dụng).*$', '', text)

    # 3. Tìm vị trí câu hỏi đầu tiên (Câu 1, Bài 1, Phần I, PHẦN 1...)
    m_first_q = re.search(r'(?m)^(?:(?:PHẦN|Phần)\s+[I|V|X\d]+|(?:Câu|Bài|Ví dụ|Question)\s+1[\.:]?)', text)
    if m_first_q:
        header_part = text[:m_first_q.start()]
        # Kiểm tra xem header_part có chứa các từ khóa hành chính đề thi không
        admin_keywords = [
            'SỞ GIÁO DỤC', 'SỞ GD', 'PHÒNG GD', 'TRƯỜNG THPT', 'TRƯỜNG THCS',
            'ĐỀ THI', 'ĐỀ KIỂM TRA', 'HỌC KỲ', 'NĂM HỌC', 'HỌ VÀ TÊN',
            'HỌ TÊN', 'SỐ BÁO DANH', 'SBD', 'MÃ ĐỀ', 'THỜI GIAN LÀM BÀI',
            'ĐỀ CHÍNH THỨC', 'ĐỀ DỰ BỊ', 'ĐỀ GỒM', 'KHẢO SÁT CHẤT LƯỢNG'
        ]
        has_admin = any(kw in header_part.upper() for kw in admin_keywords)
        if has_admin:
            # Gọt bỏ toàn bộ header hành chính thừa thãi, chỉ giữ từ câu hỏi đầu tiên
            text = text[m_first_q.start():]
        else:
            # Nếu chỉ có vài dòng hành chính lẻ tẻ, xóa từng dòng
            lines = text.split('\n')
            cleaned_lines = []
            found_first_q = False
            for line in lines:
                if not found_first_q:
                    if re.match(r'^(?:(?:PHẦN|Phần)\s+[I|V|X\d]+|(?:Câu|Bài|Ví dụ|Question)\s+1[\.:]?)', line.strip()):
                        found_first_q = True
                        cleaned_lines.append(line)
                    elif any(kw in line.upper() for kw in admin_keywords):
                        continue
                    else:
                        cleaned_lines.append(line)
                else:
                    cleaned_lines.append(line)
            text = '\n'.join(cleaned_lines)

    text = unmask_media_tags(text.strip(), media_tokens)
    return text


# Bảng ánh xạ font Symbol Word (w:sym) sang Unicode Toán học chuẩn (Học hỏi từ dự án MathRender)
WORD_SYMBOL_MAP = {
    0x22: '∀', 0x24: '∃', 0x27: '∋', 0x2a: '∗', 0x2d: '−', 0x40: '≅', 0x41: 'Α', 0x42: 'Β', 0x43: 'Χ', 0x44: 'Δ',
    0x45: 'Ε', 0x46: 'Φ', 0x47: 'Γ', 0x48: 'Η', 0x49: 'Ι', 0x4b: 'Κ', 0x4c: 'Λ', 0x4d: 'Μ', 0x4e: 'Ν', 0x4f: 'Ο',
    0x50: 'Π', 0x51: 'Θ', 0x52: 'Ρ', 0x53: 'Σ', 0x54: 'Τ', 0x55: 'Υ', 0x57: 'Ω', 0x58: 'Ξ', 0x59: 'Ψ', 0x5a: 'Ζ',
    0x5c: '∴', 0x5e: '⊥', 0x61: 'α', 0x62: 'β', 0x63: 'χ', 0x64: 'δ', 0x65: 'ε', 0x66: 'φ', 0x67: 'γ', 0x68: 'η',
    0x69: 'ι', 0x6a: 'ϕ', 0x6b: 'κ', 0x6c: 'λ', 0x6d: 'μ', 0x6e: 'ν', 0x6f: 'ο', 0x70: 'π', 0x71: 'θ', 0x72: 'ρ',
    0x73: 'σ', 0x74: 'τ', 0x75: 'υ', 0x77: 'ω', 0x78: 'ξ', 0x79: 'ψ', 0x7a: 'ζ', 0x7e: '∼', 0xa2: '′', 0xa3: '≤',
    0xa5: '∞', 0xae: '→', 0xb0: '°', 0xb1: '±', 0xb2: '″', 0xb3: '≥', 0xb4: '×', 0xb6: '∂', 0xb8: '÷', 0xb9: '≠',
    0xba: '≡', 0xbb: '≈', 0xbc: '…', 0xc6: '∅', 0xc7: '∩', 0xc8: '∪', 0xc9: '⊃', 0xcc: '⊂', 0xcd: '⊆', 0xce: '∈',
    0xcf: '∉', 0xd0: '∠', 0xd1: '∇', 0xd6: '√', 0xd7: '⋅', 0xdb: '⇔', 0xde: '⇒', 0xe5: '∑', 0xf2: '∫'
}

def decode_word_symbol(font: str, char_hex: str) -> str:
    try:
        code = int(char_hex, 16)
        low = code - 0xf000 if code >= 0xf000 else code
        if 'symbol' in (font or '').lower() and low in WORD_SYMBOL_MAP:
            return WORD_SYMBOL_MAP[low]
        if low >= 0x20:
            return chr(low)
    except Exception:
        pass
    return ""


def process_blip_formula_or_image(b, doc, pieces, para_images, prev_piece_type):
    """
    Xử lý phần tử blip hình ảnh trong Word:
    1. Ưu tiên kiểm tra xem có phải công thức MathType bị dán/flattened dưới dạng WMF/ảnh không.
       Nếu có, giải mã trực tiếp thành chuẩn KaTeX ($...$).
    2. Nếu là hình vẽ/đồ thị thực sự, đưa vào para_images để chiếu 2 cột (Zero-Scroll).
    3. Tuyệt đối không để rò rỉ mã HTML <img src="data:image/x-wmf..."> ra ngoài text.
    """
    rId = b.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
    if not (rId and rId in doc.part.rels):
        return prev_piece_type, 0

    img_part = doc.part.rels[rId].target_part
    blob = img_part.blob
    mime = getattr(img_part, 'content_type', 'image/png') or 'image/png'

    # 1. Thử giải mã MathType (cả chuỗi TeX và luồng MTEF nhúng trong WMF)
    mt_tex = extract_mathtype_formula(None, blob)
    if mt_tex:
        mt_tex = clean_math_latex(mt_tex)
        if mt_tex:
            trailing_punct = ""
            m_punct = re.search(r'([.,;:?!])$', mt_tex)
            if m_punct and not re.search(r'\\(?:dots|ldots|cdots)\s*$', mt_tex):
                trailing_punct = m_punct.group(1)
                mt_tex = mt_tex[:-1].strip()
            if prev_piece_type == 'text' and pieces:
                last_txt = pieces[-1]
                if last_txt and not last_txt.endswith((' ', '\t', '\n')):
                    if re.search(r'[a-zA-Zà-ỹÀ-Ỹ0-9\)\}\]]$', last_txt):
                        pieces.append(" ")
            pieces.append(f"${mt_tex}${trailing_punct}")
            return ('punct' if trailing_punct else 'math'), 0

    # 2. Nếu là hình ảnh thông thường
    b64 = base64.b64encode(blob).decode('utf-8')
    cx, cy = 0.0, 0.0
    parent = b.getparent()
    while parent is not None and not (cx > 0 and cy > 0):
        for ext in parent.xpath('.//*[local-name()="extent"]'):
            cx = int(ext.get('cx', 0)) / 914400 * 2.54
            cy = int(ext.get('cy', 0)) / 914400 * 2.54
            if cx > 0:
                break
        parent = parent.getparent()
        if parent is not None and parent.tag.endswith(('p', 'body')):
            break

    # Nếu là đồ thị / hình vẽ thực sự
    if cx >= 2.8 or cy >= 1.8 or (cx == 0.0 and cy == 0.0):
        para_images.append(f"data:{mime};base64,{b64}")
        return prev_piece_type, 1
    elif 'wmf' not in mime.lower():
        para_images.append(f"data:{mime};base64,{b64}")
        return prev_piece_type, 1

    return prev_piece_type, 0


@app.post("/api/parse-docx-file")
async def parse_docx_file(file: UploadFile = File(...)):
    """
    Đọc toàn diện file Word (.docx):
    1. Trích xuất văn bản và công thức toán học (OMML) theo ĐÚNG THỨ TỰ TỰ NHIÊN (Paragraph & Table).
    2. Tự động sửa lỗi dính chữ tổng quát (Smart Spacing Bridge) giữa số/toán và tiếng Việt.
    3. Chuẩn hóa triệt để véc-tơ (A'A→, AC'→, A'\vec{A}...).
    4. Bóc tách 100% hình ảnh (đồ thị, hình không gian, bảng biến thiên) với KHIÊN BẢO VỆ BASE64,
       chuyển sang Base64 Data URL và nhúng trực tiếp vào câu hỏi tương ứng để chiếu slide!
    5. Tự động nhận diện và loại bỏ rác/tiêu đề hành chính thừa thãi (Sở, Trường, Mã đề...).
    """
    try:
        content = await file.read()
        import io, base64, docx
        doc = docx.Document(io.BytesIO(content))
        lines = []
        img_count = 0

        vn_alpha_set = set('àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ'
                           'ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ')

        # Duyệt tất cả các phần tử theo đúng thứ tự dòng thời gian của file Word (paragraph và table)
        body_elements = doc._element.body
        for elem in body_elements:
            elem_tag = elem.tag

            # 1. Đoạn văn bản (Paragraph)
            if elem_tag.endswith('p'):
                p = docx.text.paragraph.Paragraph(elem, doc)
                pieces = []
                para_images = []
                prev_piece_type = None  # 'math', 'text', or None

                for child in p._p:
                    tag = child.tag
                    # Run văn bản thông thường
                    if tag.endswith('r'):
                        # 0. Kiểm tra MathType OLE Object nhúng bên trong Run
                        mt_tex = extract_mathtype_from_element(child, doc)
                        if mt_tex:
                            mt_tex = clean_math_latex(mt_tex)
                            if mt_tex:
                                trailing_punct = ""
                                m_punct = re.search(r'([.,;:?!])$', mt_tex)
                                if m_punct and not re.search(r'\\(?:dots|ldots|cdots)\s*$', mt_tex):
                                    trailing_punct = m_punct.group(1)
                                    mt_tex = mt_tex[:-1].strip()
                                if prev_piece_type == 'text' and pieces:
                                    last_txt = pieces[-1]
                                    if last_txt and not last_txt.endswith((' ', '\t', '\n')):
                                        if re.search(r'[a-zA-Zà-ỹÀ-Ỹ0-9\)\}\]]$', last_txt):
                                            pieces.append(" ")
                                pieces.append(f"${mt_tex}${trailing_punct}")
                                prev_piece_type = 'punct' if trailing_punct else 'math'
                            continue

                        r_text = ""
                        for t in child.xpath('.//*[local-name()="t"]'):
                            if t.text:
                                r_text += t.text
                        # Giải mã ký tự Symbol Word (w:sym) như ≤, ≥, ∈, ±, ≠...
                        for sym in child.xpath('.//*[local-name()="sym"]'):
                            font = sym.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}font') or sym.get('font') or ''
                            code_str = sym.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}char') or sym.get('char') or ''
                            ch = decode_word_symbol(font, code_str)
                            if ch:
                                r_text += ch
                        if child.xpath('.//*[local-name()="tab"]'):
                            r_text += " "

                        r_text = normalize_primes(r_text)

                        # Nhận diện subscript / superscript thông thường từ Word (Ctrl+= hoặc Ctrl+Shift+=)
                        vert = child.xpath('.//*[local-name()="vertAlign"]')
                        if vert:
                            v_val = vert[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or vert[0].get('val')
                            if v_val == 'subscript' and r_text:
                                r_text = f"_{{{r_text}}}"
                            elif v_val == 'superscript' and r_text:
                                r_text = f"^{{{r_text}}}"

                        if r_text:
                            # 1. Ranh giới Math -> Text: Nếu trước đó là <m:oMath> hoặc MathType ($formula$) và Run văn bản kế tiếp bắt đầu bằng chữ cái tiếng Việt/số
                            if prev_piece_type == 'math':
                                if not r_text.startswith((' ', '\t', '\n')) and re.match(r'^[a-zA-Zà-ỹÀ-Ỹ0-9]', r_text):
                                    pieces.append(" ")
                            # 2. Ranh giới Text -> Text dính nhau giữa 2 Run (ví dụ: (-1; 3; -2)và -> (-1; 3; -2) và, 50bình -> 50 bình)
                            elif prev_piece_type == 'text' and pieces:
                                last_txt = pieces[-1]
                                if last_txt and not last_txt.endswith((' ', '\t', '\n')) and not r_text.startswith((' ', '\t', '\n')):
                                    prev_char = last_txt[-1]
                                    curr_char = r_text[0]
                                    # Nếu trước đó là dấu đóng ngoặc ) hoặc chữ cái/số và sau đó là chữ cái/số
                                    if prev_char in (')', ']', '}') or prev_char.isalnum() or prev_char in vn_alpha_set:
                                        if curr_char.isalnum() or curr_char in vn_alpha_set:
                                            pieces.append(" ")

                            pieces.append(r_text)
                            prev_piece_type = 'text'

                        # Bóc tách ảnh từ w:r (chỉ lấy drawing thực sự, giải mã nếu là công thức MathType)
                        for b in child.xpath('.//*[local-name()="blip"]'):
                            prev_piece_type, img_added = process_blip_formula_or_image(
                                b, doc, pieces, para_images, prev_piece_type
                            )
                            img_count += img_added

                    # Run công thức toán học chuẩn OMML (Phân số, số mũ, căn, vectơ, ngoặc...)
                    elif tag.endswith('oMath') or tag.endswith('oMathPara'):
                        m_str = omml_element_to_latex(child)
                        m_str = clean_math_latex(m_str)
                        if m_str:
                            trailing_punct = ""
                            m_punct = re.search(r'([.,;:?!])$', m_str)
                            if m_punct and not re.search(r'\\(?:dots|ldots|cdots)\s*$', m_str):
                                trailing_punct = m_punct.group(1)
                                m_str = m_str[:-1].strip()
                            # Ranh giới Text -> Math: Nếu trước đó là văn bản thường kết thúc bằng chữ cái/số (hoặc ngoặc), chèn space
                            if prev_piece_type == 'text' and pieces:
                                last_txt = pieces[-1]
                                if last_txt and not last_txt.endswith((' ', '\t', '\n')):
                                    if re.search(r'[a-zA-Zà-ỹÀ-Ỹ0-9\)\}\]]$', last_txt):
                                        pieces.append(" ")
                            pieces.append(f"${m_str}${trailing_punct}")
                            prev_piece_type = 'punct' if trailing_punct else 'math'

                    # Khung chứa MathType OLE Object hoặc hình ảnh OLE
                    elif tag.endswith('object'):
                        mt_tex = extract_mathtype_from_element(child, doc)
                        if mt_tex:
                            mt_tex = clean_math_latex(mt_tex)
                            if mt_tex:
                                trailing_punct = ""
                                m_punct = re.search(r'([.,;:?!])$', mt_tex)
                                if m_punct and not re.search(r'\\(?:dots|ldots|cdots)\s*$', mt_tex):
                                    trailing_punct = m_punct.group(1)
                                    mt_tex = mt_tex[:-1].strip()
                                if prev_piece_type == 'text' and pieces:
                                    last_txt = pieces[-1]
                                    if last_txt and not last_txt.endswith((' ', '\t', '\n')):
                                        if re.search(r'[a-zA-Zà-ỹÀ-Ỹ0-9\)\}\]]$', last_txt):
                                            pieces.append(" ")
                                pieces.append(f"${mt_tex}${trailing_punct}")
                                prev_piece_type = 'punct' if trailing_punct else 'math'
                        else:
                            for b in child.xpath('.//*[local-name()="blip"]'):
                                prev_piece_type, img_added = process_blip_formula_or_image(
                                    b, doc, pieces, para_images, prev_piece_type
                                )
                                img_count += img_added
                            for b in child.xpath('.//*[local-name()="imagedata"]'):
                                rId = b.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                                if rId and rId in doc.part.rels:
                                    img_part = doc.part.rels[rId].target_part
                                    b64 = base64.b64encode(img_part.blob).decode('utf-8')
                                    mime = getattr(img_part, 'content_type', 'image/png') or 'image/png'
                                    para_images.append(f"data:{mime};base64,{b64}")
                                    img_count += 1

                    # Khung vẽ / Hình ảnh (Drawing) trực tiếp
                    elif tag.endswith('drawing'):
                        for b in child.xpath('.//*[local-name()="blip"]'):
                            prev_piece_type, img_added = process_blip_formula_or_image(
                                b, doc, pieces, para_images, prev_piece_type
                            )
                            img_count += img_added

                    # Xử lý AlternateContent (Choice/Fallback trong Word)
                    elif tag.endswith('AlternateContent'):
                        choice = child.xpath('.//*[local-name()="Choice"]')
                        target_node = choice[0] if choice else (child.xpath('.//*[local-name()="Fallback"]') or [None])[0]
                        if target_node is not None:
                            mt_tex = extract_mathtype_from_element(target_node, doc)
                            if mt_tex:
                                mt_tex = clean_math_latex(mt_tex)
                                if mt_tex:
                                    trailing_punct = ""
                                    m_punct = re.search(r'([.,;:?!])$', mt_tex)
                                    if m_punct and not re.search(r'\\(?:dots|ldots|cdots)\s*$', mt_tex):
                                        trailing_punct = m_punct.group(1)
                                        mt_tex = mt_tex[:-1].strip()
                                    if prev_piece_type == 'text' and pieces:
                                        last_txt = pieces[-1]
                                        if last_txt and not last_txt.endswith((' ', '\t', '\n')):
                                            if re.search(r'[a-zA-Zà-ỹÀ-Ỹ0-9\)\}\]]$', last_txt):
                                                pieces.append(" ")
                                    pieces.append(f"${mt_tex}${trailing_punct}")
                                    prev_piece_type = 'punct' if trailing_punct else 'math'
                            else:
                                for b in target_node.xpath('.//*[local-name()="blip"]'):
                                    prev_piece_type, img_added = process_blip_formula_or_image(
                                        b, doc, pieces, para_images, prev_piece_type
                                    )
                                    img_count += img_added

                p_text = ''.join(pieces).strip()
                if p_text:
                    if re.match(r'^(?:\|\s*|\[)?\s*(?:(?:auto)?crop_[^\s\|\]]+\.(?:png|jpg)\s*\|?\s*)?(?:box=\[[^\]]*\]\s*\|?\s*)?(?:Biểu đồ|Đồ thị|Hình vẽ|Bảng biến thiên|Bảng số liệu|Hình ảnh|Minh họa|Sơ đồ|mô phỏng)\b[^\]\n]*\]?$', p_text, re.IGNORECASE):
                        p_text = ""
                    else:
                        p_text = re.sub(r'^(?:\s*(?:\|\s*|\[)?\s*(?:(?:auto)?crop_[^\s\|\]]+\.(?:png|jpg)\s*\|?\s*)?(?:box=\[[^\]]*\]\s*\|?\s*)?(?:Biểu đồ|Đồ thị|Hình vẽ|Bảng biến thiên|Bảng số liệu|Hình ảnh|Minh họa|Sơ đồ|mô phỏng)\b[^\]\n]*\]?\s*)+', '', p_text, flags=re.IGNORECASE).strip()

                if p_text:
                    lines.append(p_text)

                for img_data in para_images:
                    lines.append(f"[HÌNH_VẼ: image_data={img_data} | Hình minh họa từ file Word]")

            # 2. Bảng biểu (Table) xuất hiện đúng dòng thời gian tự nhiên
            elif elem_tag.endswith('tbl'):
                tbl = docx.table.Table(elem, doc)
                table_lines = []
                for row in tbl.rows:
                    row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_cells:
                        row_text = " | ".join(row_cells)
                        table_lines.append(row_text)
                if table_lines:
                    lines.append("\n".join(table_lines))

        full_text = "\n\n".join(lines)

        # ══════════════════════════════════════════════════════════════════════════
        # KIẾN TRÚC 1 CHIỀU CHUẨN CÔNG NGHIỆP AZOTA:
        # BƯỚC 1: KHIÊN BẢO VỆ MEDIA (BASE64)
        # BƯỚC 2: ĐÓNG KÉT BẢO VỆ CÔNG THỨC TOÁN (MATH TOKENIZATION)
        # BƯỚC 3: XỬ LÝ LỌC CẤU TRÚC ĐỀ THI & VĂN BẢN (KHÔNG CHẠM VÀO LATEX)
        # BƯỚC 4: MỞ KÉT HOÀN NGUYÊN CÔNG THỨC TOÁN
        # BƯỚC 5: ĐẢM BẢO KHOẢNG CÁCH RANH GIỚI MATH (SPACING BRIDGE)
        # BƯỚC 6: HOÀN NGUYÊN MEDIA DATA URL
        # ══════════════════════════════════════════════════════════════════════════

        # 1. KHIÊN BẢO VỆ TUYỆT ĐỐI CHO TẤT CẢ THẺ MEDIA (BASE64)
        full_text, media_tokens = mask_media_tags(full_text)

        # 2. ĐÓNG KÉT BẢO VỆ TOÀN BỘ CÔNG THỨC TOÁN HỌC (MATH TOKENIZATION)
        full_text, math_tokens = mask_math_tokens(full_text)

        # 3. CHẠY CÁC BỘ LỌC CẤU TRÚC VĂN BẢN TRÊN TEXT ĐÃ ĐƯỢC BẢO VỆ AN TOÀN
        # Nối lại các câu bị ngắt dòng ngang sang đoạn mới nếu bắt đầu bằng chữ thường
        full_text = re.sub(r'([^\.\?!…:;\]\)\$])\s*\n+\s*([a-zà-ỹ0-9][^\n]*)', r'\1 \2', full_text)
        full_text = re.sub(r'\n{3,}', '\n\n', full_text)
        # Gọt bỏ tiêu đề hành chính trường/sở thừa thãi
        full_text = clean_exam_metadata_and_headers(full_text)
        # Tách dính chữ giữa các từ tiếng Việt và số ngoài toán
        full_text = separate_stuck_math_and_words(full_text)
        # Tách dòng độc lập cho khối Lời giải / Hướng dẫn giải / Chọn [A-D] nếu dính sau phương án D
        full_text = re.sub(
            r'(\b[A-D]\.\s+[^\n]+?)\s*(\b(?:Lời\s+giải|Hướng\s+dẫn\s+giải|Giải\s+chi\s+tiết|HƯỚNG\s+DẪN\s+GIẢI|LỜI\s+GIẢI|GIẢI\s+CHI\s+TIẾT|HDG)\s*[:\.\-]?|\b(?:Chọn|Đáp\s*án)\s*[:\.]?\s*[A-D]\b)',
            r'\1\n\n\2',
            full_text,
            flags=re.IGNORECASE
        )

        # 4. HOÀN NGUYÊN TOÀN BỘ CÔNG THỨC TOÁN HỌC NGUYÊN BẢN 100%
        full_text = unmask_math_tokens(full_text, math_tokens)

        # 4b. CHUẨN HÓA VECTƠ TOÀN DIỆN (CẢ TRONG MATH VÀ TEXT NGOÀI)
        full_text = sanitize_plain_text_vectors(full_text)
        # Tách dòng độc lập cho các ý Đúng/Sai a,b,c,d, ý tự luận a-h và 1,2,3
        full_text = split_sub_items_to_lines(full_text)

        # 4c. TỰ ĐỘNG BỌC CÔNG THỨC TOÁN HỌC CÒN THIẾU DẤU $ (x_{CĐ} = 4, x_{CT} = 0, y = 2x - 1...)
        full_text = auto_wrap_math_symbols(full_text)

        # 4d. CHUẨN HÓA BẢNG MÃ UNICODE DỰNG SẴN NFC (CHỐNG PHÂN RÃ CHỮ TIẾNG VIỆT)
        import unicodedata
        full_text = unicodedata.normalize('NFC', full_text)

        # 5. SMART SPACING BRIDGE: Đảm bảo ranh giới quanh công thức $...$ không bao giờ dính chữ
        full_text = ensure_math_spacing_bridge(full_text)

        # 6. HOÀN NGUYÊN THẺ MEDIA NGUYÊN BẢN 100%
        full_text = unmask_media_tags(full_text, media_tokens)

        return {"text": full_text, "filename": file.filename, "images_count": img_count}
    except Exception as e:
        raise HTTPException(500, f"Lỗi đọc file Word: {str(e)}")


@app.post("/api/export-slide-offline")
async def export_slide_offline(request: Request):
    """
    Đóng gói toàn bộ Slide Đề Toán thành 1 file HTML tự chạy Offline độc lập:
    - Nhúng toàn bộ câu hỏi, công thức KaTeX, bố cục 2 cột và hình ảnh Base64.
    - Sử dụng template chuẩn mực offline_presentation.html, không lỗi cú pháp regex.
    """
    try:
        body = await request.json()
        slides = body.get('slides', [])
        if not slides:
            raise HTTPException(400, "Danh sách slide rỗng!")

        template_path = Path("templates") / "offline_presentation.html"
        if not template_path.exists():
            raise HTTPException(500, "Không tìm thấy file template templates/offline_presentation.html!")

        html_content = template_path.read_text(encoding="utf-8")

        import json
        slides_json = json.dumps(slides, ensure_ascii=False)
        slides_json = slides_json.replace("</script>", "<\\/script>")

        # Thay thế placeholder dữ liệu slide
        html_content = html_content.replace("/*__SLIDES_DATA__*/[]", slides_json)

        return Response(
            content=html_content.encode("utf-8"),
            media_type="text/html; charset=utf-8",
            headers={
                "Content-Disposition": 'attachment; filename="[TrinhChieu]_De_Thi_Toan.html"'
            }
        )
    except Exception as e:
        if isinstance(e, HTTPException):
            raise e
        raise HTTPException(500, f"Lỗi đóng gói Slide Offline: {str(e)}")


def safe_parse_gemini_json(s: str, fallback_question: str = "", fallback_choices: list = None) -> dict:
    """
    Trích xuất an toàn câu hỏi và phương án từ phản hồi của Gemini:
    Hỗ trợ cả định dạng JSON (kể cả có escape LaTeX), định dạng thẻ [ĐỀ BÀI] / [PHƯƠNG ÁN],
    hoặc văn bản thuần A. B. C. D. Không bao giờ quăng lỗi làm sập request!
    """
    if fallback_choices is None:
        fallback_choices = []
    if not s or not s.strip():
        return {"question": fallback_question, "choices": fallback_choices}

    s = s.strip()

    # 1. Thử bóc tách nếu có khối JSON
    m_json = re.search(r'\{[\s\S]*\}', s)
    if m_json:
        j_block = m_json.group(0).strip()
        try:
            data = json.loads(j_block)
            if isinstance(data, dict):
                return {
                    "question": data.get("question", fallback_question),
                    "choices": data.get("choices", fallback_choices)
                }
        except Exception:
            pass

        try:
            fixed = re.sub(r'\\(?![/"\\bfnrt]|u[0-9a-fA-F]{4})', r'\\\\', j_block)
            fixed = re.sub(r',\s*([\]\}])', r'\1', fixed)
            data = json.loads(fixed)
            if isinstance(data, dict):
                return {
                    "question": data.get("question", fallback_question),
                    "choices": data.get("choices", fallback_choices)
                }
        except Exception:
            pass

    # 2. Thử bóc tách theo khối văn bản tự nhiên [ĐỀ BÀI] / [CÂU HỎI]
    q_text = ""
    choices = []
    m_q = re.search(r'\[(?:ĐỀ\s*BÀI|CÂU\s*HỎI|NỘI\s*DUNG)\]\s*:?\s*([\s\S]*?)(?=\[(?:PHƯƠNG\s*ÁN|CÁC\s*PHƯƠNG\s*ÁN|LỰA\s*CHỌN)\]|\n\s*[A-D]\.\s+|\Z)', s, re.I)
    if m_q and m_q.group(1).strip():
        q_text = m_q.group(1).strip()

    # Bóc tách các phương án A, B, C, D từ các dòng
    lines = s.splitlines()
    for line in lines:
        line_str = line.strip()
        m_line = re.match(r'^([A-D])\.\s*(.+)$', line_str)
        if m_line:
            choices.append({"label": f"{m_line.group(1)}.", "text": m_line.group(2).strip()})

    if not choices:
        for m in re.finditer(r'(?:^|\n|\s{2,})([A-D])\.\s*([^\n\r]+)', s):
            choices.append({"label": f"{m.group(1)}.", "text": m.group(2).strip()})
        seen = set()
        clean_c = []
        for c in choices:
            if c['label'] not in seen:
                seen.add(c['label'])
                clean_c.append(c)
        choices = clean_c if len(clean_c) >= 2 else []

    if not q_text:
        m_before_a = re.split(r'\n\s*A\.\s+', s, maxsplit=1)
        if len(m_before_a) > 1 and m_before_a[0].strip():
            q_text = m_before_a[0].strip()
        else:
            q_text = s

    q_text = re.sub(r'^\[(?:ĐỀ\s*BÀI|CÂU\s*HỎI|NỘI\s*DUNG)\]\s*:?\s*', '', q_text, flags=re.I).strip()
    q_text = re.sub(r'^```[a-z]*\s*', '', q_text, flags=re.I)
    q_text = re.sub(r'\s*```$', '', q_text).strip()

    return {
        "question": q_text or fallback_question,
        "choices": choices if choices else fallback_choices
    }


@app.post("/api/ai-ocr-question")
async def ai_ocr_question(request: Request):
    """
    Dùng Gemini AI đọc lại và chuẩn hóa một câu hỏi toán học (OCR / Math Fallback):
    - Đọc lại công thức từ văn bản bị lỗi hoặc từ hình ảnh câu hỏi.
    - Chuẩn hóa toàn bộ công thức sang định dạng KaTeX ($...$).
    - Trả về câu hỏi và danh sách đáp án A, B, C, D đã được làm sạch để slide cập nhật tức thì.
    """
    try:
        data = await request.json()
        question_text = data.get("question_text", "").strip()
        choices = data.get("choices", [])
        images = data.get("images", [])
        api_keys_str = data.get("api_keys", "")
        model_name = data.get("model_name", "auto")

        if not question_text and not images:
            raise HTTPException(400, "Thiếu dữ liệu câu hỏi hoặc hình ảnh để đọc lại!")

        # Gom danh sách API key
        active_keys = list(DEFAULT_API_KEYS)
        if api_keys_str:
            try:
                parsed = json.loads(api_keys_str) if api_keys_str.startswith('[') else re.split(r'[,;\n\r]+', api_keys_str)
                for k in parsed:
                    k = str(k).strip()
                    if k and k not in active_keys:
                        active_keys.insert(0, k)
            except Exception:
                pass

        choices_formatted = []
        for c in choices:
            label = c.get("label", "")
            text = c.get("text", "")
            choices_formatted.append(f"{label} {text}")
        choices_text = "\n".join(choices_formatted) if choices_formatted else ""

        prompt = f"""Bạn là chuyên gia số hóa đề thi Toán học THPT Việt Nam.
Nhiệm vụ của bạn: Hãy đọc lại, sửa lỗi và chuẩn hóa nội dung câu hỏi toán học sau đây (đặc biệt là các công thức toán bị lỗi cú pháp, bị dính chữ, hoặc bị dán nhầm mã HTML):

[NỘI DUNG CÂU HỎI HIỆN TẠI]:
{question_text}
"""
        if choices_text:
            prompt += f"""
[CÁC PHƯƠNG ÁN LỰA CHỌN]:
{choices_text}
"""

        prompt += """
QUY TẮC BẮT BUỘC:
1. Mọi công thức toán học, biểu thức, biến số, phân số, vectơ, tọa độ, hệ phương trình BẮT BUỘC viết dưới dạng chuẩn KaTeX đặt trong cặp dấu $...$ (ví dụ: $y = \\frac{x-1}{x+1}$, $[0; 3]$, $\\vec{a}$, $\\min\\limits_{[0; 3]} y = -3$).
2. Giữ nguyên toàn bộ nội dung đề bài và các phương án, chỉ sửa lỗi công thức, loại bỏ thẻ HTML rác (như <img...>), và chuẩn hóa định dạng.
3. Trình bày rõ ràng theo định dạng sau:
[ĐỀ BÀI]:
Nội dung câu hỏi đã chuẩn hóa KaTeX

[PHƯƠNG ÁN]:
A. Nội dung phương án A
B. Nội dung phương án B
C. Nội dung phương án C
D. Nội dung phương án D
(Nếu không có phương án trắc nghiệm thì chỉ ghi phần [ĐỀ BÀI]).
"""
        pil_img = None
        if images and isinstance(images, list):
            for img_item in images:
                if not isinstance(img_item, str):
                    continue
                if img_item.startswith('data:image'):
                    try:
                        header, b64_data = img_item.split(',', 1)
                        if 'wmf' not in header.lower():
                            img_bytes = base64.b64decode(b64_data)
                            pil_img = Image.open(io.BytesIO(img_bytes))
                            break
                    except Exception:
                        pass
                elif img_item.startswith('/api/image/'):
                    try:
                        rel = img_item.replace('/api/image/', '', 1).strip('/')
                        fpath = TEMP_DIR / rel
                        if fpath.exists():
                            pil_img = Image.open(fpath)
                            break
                    except Exception:
                        pass

        raw_output, used_model = await _generate_content_with_retry(
            img=pil_img,
            prompt_text=prompt,
            model_name=model_name or "auto",
            api_keys=active_keys,
            timeout_seconds=25.0
        )

        result_data = safe_parse_gemini_json(raw_output, fallback_question=question_text, fallback_choices=choices)
        return {
            "success": True,
            "question": result_data.get("question", question_text),
            "choices": result_data.get("choices", choices),
            "used_model": used_model
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(500, f"Lỗi Gemini AI đọc lại câu hỏi: {str(e)}")


@app.post("/api/ai-solve-question")
async def ai_solve_question(request: Request):
    """
    Dùng Gemini AI giải chi tiết câu hỏi toán học và tự động xác định đáp án đúng (A, B, C, D) cho slide trình chiếu.
    """
    try:
        data = await request.json()
        question_text = data.get("question_text", "").strip()
        choices = data.get("choices", [])
        api_keys_str = data.get("api_keys", "")
        model_name = data.get("model_name", "auto")

        if not question_text:
            raise HTTPException(400, "Thiếu nội dung câu hỏi để giải!")

        # Gom danh sách API key
        active_keys = list(DEFAULT_API_KEYS)
        if api_keys_str:
            try:
                parsed = json.loads(api_keys_str) if api_keys_str.startswith('[') else re.split(r'[,;\n\r]+', api_keys_str)
                for k in parsed:
                    k = str(k).strip()
                    if k and k not in active_keys:
                        active_keys.insert(0, k)
            except Exception:
                pass

        choices_formatted = []
        for c in choices:
            label = c.get("label", "")
            text = c.get("text", "")
            choices_formatted.append(f"{label} {text}")
        choices_text = "\n".join(choices_formatted) if choices_formatted else "Không có phương án trắc nghiệm (Tự luận / Trả lời ngắn)"

        prompt = f"""Bạn là chuyên gia sư phạm Toán học THPT Việt Nam.
Hãy giải chi tiết câu hỏi toán học sau đây một cách chính xác, ngắn gọn, súc tích và mạch lạc:

[ĐỀ BÀI CÂU HỎI]:
{question_text}

[CÁC PHƯƠNG ÁN LỰA CHỌN]:
{choices_text}

QUY TẮC BẮT BUỘC:
1. Trình bày lời giải từng bước rõ ràng, ngắn gọn, sư phạm cho học sinh THPT.
2. Mọi ký hiệu, biểu thức, biến số toán học BẮT BUỘC viết bằng chuẩn LaTeX đặt trong cặp dấu $...$ (hoặc $$...$$ nếu công thức riêng dòng).
3. Nếu là câu hỏi trắc nghiệm nhiều phương án, ở DÒNG CUỐI CÙNG của lời giải BẮT BUỘC ghi rõ kết luận: 'Chọn A.' (hoặc 'Chọn B.', 'Chọn C.', 'Chọn D.').
"""
        raw_solution, used_model = await _generate_content_with_retry(
            img=None,
            prompt_text=prompt,
            model_name=model_name,
            api_keys=active_keys,
            timeout_seconds=22.0
        )

        # Trích xuất đáp án đúng nếu có
        correct_label = None
        match = re.search(r'(?:Chọn|Đáp\s*án|Phương\s*án)(?:[\s:]*(?:đúng|chính\s*xác|là|đáp\s*án|phương\s*án))*[\s:\.]*([A-D])\b', raw_solution, re.IGNORECASE)
        if match:
            correct_label = match.group(1).upper()

        return {
            "solution": raw_solution.strip(),
            "correct_label": correct_label,
            "used_model": used_model
        }
    except Exception as e:
        raise HTTPException(500, f"Lỗi Gemini AI giải đề: {str(e)}")


@app.post("/api/export-docx")
async def export_docx(
    title: str = Form("Tài liệu Toán học"),
    pages_json: str = Form(...),  # JSON string of list of dicts: [{"page": 1, "latex": "..."}]
    session_id: str = Form(None),
    math_format: str = Form("omml"),  # "omml" (Word native Equation) hoặc "latex" ($...$ và $$...$$)
):
    """Xuất kết quả LaTeX sang file Word (.docx) chứa công thức native OMML hoặc văn bản LaTeX và hình ảnh tự động nhúng."""
    try:
        pages_data = json.loads(pages_json)
        if not pages_data:
            raise HTTPException(400, "Không có nội dung để xuất!")

        session_dir = (TEMP_DIR / session_id) if session_id else None

        # 🎯 Đồng bộ hóa toàn diện cho từng trang trước khi export:
        for p_data in pages_data:
            p_latex = p_data.get('latex', '')
            if not p_latex:
                continue
            # 0. Lọc bỏ triệt để chân trang scan
            p_latex = remove_page_footers(p_latex)
            # 1. Làm sạch thẩm mỹ & sửa lỗi ký tự OCR (13|5, ABCD. A'B'C'D', khoảng trắng dấu câu)
            p_latex = sanitize_text_aesthetics(p_latex)
            # 2. Gỡ bỏ dấu $...$ ôm cả câu tiếng Việt trước khi tách paragraph (tránh dính chữ điquađiểm)
            p_latex = unwrap_vietnamese_math(p_latex)
            # 3. Hàn gắn tọa độ, nối $...$ bị rách dòng qua nhiều \n
            p_latex = heal_broken_math_and_coords(p_latex)
            # 4. Nối từ bị ngắt dòng sai từ OCR
            p_latex = join_broken_ocr_lines(p_latex)
            # 5. Chuẩn hóa vectơ dạng chữ ngoài văn bản sang LaTeX Equation
            p_latex = sanitize_plain_text_vectors(p_latex)
            # 6. Dọn dẹp thẩm mỹ lần cuối
            p_latex = sanitize_text_aesthetics(p_latex)

            # Tự động cắt ảnh nếu session_dir có sẵn và thẻ chỉ có box mà chưa có image
            if session_dir and session_dir.exists():
                p_num = p_data.get('page', 1)
                if re.search(r'\[(?:HÌNH|BẢNG|ĐỒ|IMAGE|TABLE|FIGURE)[^\]]*box=', p_latex, re.I) and not re.search(r'\[(?:HÌNH|BẢNG|ĐỒ|IMAGE|TABLE|FIGURE)[^\]]*image=', p_latex, re.I):
                    cand_files = [f"page_{p_num}.png", f"page_{p_num}.jpg", f"image_{p_num}.png", f"image_{p_num}.jpg", "image_1.jpg", "image_1.png", "page_1.png"]
                    orig_file = None
                    for cand in cand_files:
                        if (session_dir / cand).exists():
                            orig_file = cand
                            break
                    if orig_file:
                        p_latex, _ = auto_crop_figures(session_id, orig_file, p_latex)
            p_data['latex'] = p_latex

        doc_id = uuid.uuid4().hex[:8]
        out_filename = f"math_doc_{doc_id}.docx"
        out_path = TEMP_DIR / out_filename

        export_latex_to_docx(
            pages_data,
            str(out_path),
            title=title,
            session_dir=session_dir,
            math_format=math_format
        )

        return FileResponse(
            str(out_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"{title}.docx"
        )
    except Exception as e:
        raise HTTPException(500, f"Lỗi tạo file Word: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    import sys
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    port = int(os.environ.get("PORT", 8000))
    print("\n" + "=" * 55)
    print(f"  📐 Math OCR Tool — Port: {port}")
    print(f"  🌐 Mở trình duyệt: http://localhost:{port}")
    print("  🔑 API Key:", "✅ Đã cấu hình" if GEMINI_API_KEY else "❌ Chưa có")
    print("=" * 55 + "\n")
    uvicorn.run(app, host="0.0.0.0", port=port)
