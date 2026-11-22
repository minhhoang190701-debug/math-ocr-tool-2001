"""
Module: pdf_native_extractor.py
Trích xuất ảnh gốc nguyên bản từ file PDF số hóa (Digital PDF)
và khớp nối thông minh với câu hỏi đề thi (thay thế cho việc crop lại từ ảnh render).
Hỗ trợ cơ chế Hybrid Waterfall Thông Minh:
- Ưu tiên 1A: Tự động bóc ảnh bitmap nhúng (XObject PNG/JPEG gốc nét 100%).
- Ưu tiên 1B: Tự động gom cụm nét vẽ Vector Drawing (Hình học, Đồ thị, Bảng biểu) và kết xuất 300 DPI nguyên bản.
- Luồng Fallback an toàn: Nếu là bản scan toàn trang hoặc không bóc tách được -> Fallback 100% sang AI Vision + smart_trim_figure.
"""

import io
import json
import logging
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any

import pymupdf as fitz
from PIL import Image

logger = logging.getLogger("pdf_native_extractor")


def is_black_or_corrupted_image(
    img: Image.Image,
    black_thresh: int = 40,
    max_dark_ratio: float = 0.70
) -> bool:
    """
    Kiểm tra xem ảnh có bị đen xì (do mất kênh smask trong PDF) hoặc là một khối đơn sắc hỏng không.
    - black_pixels > 70% (pixel có giá trị độ sáng < black_thresh)
    - Hoặc toàn bộ ảnh chỉ có 1 màu duy nhất (colors count == 1)
    """
    try:
        if img.mode != "RGB":
            img_rgb = img.convert("RGB")
        else:
            img_rgb = img

        colors = img_rgb.getcolors(maxcolors=10)
        if colors and len(colors) == 1:
            return True

        thumb = img_rgb.resize((100, 100))
        gray = thumb.convert("L")
        hist = gray.histogram()
        dark_pixels = sum(hist[:black_thresh])
        if (dark_pixels / 10000.0) >= max_dark_ratio:
            return True

        return False
    except Exception as e:
        logger.warning(f"Lỗi kiểm tra tính toàn vẹn của ảnh: {e}")
        return False


def _cluster_vector_drawings(
    page: fitz.Page,
    gap: float = 14.0,
    min_dim: float = 35.0,
    max_area_ratio: float = 0.50
) -> List[Dict[str, Any]]:
    """
    Gom cụm các nét vẽ vector drawing gần nhau thành các bounding box hình vẽ logic.
    Loại trừ các đường kẻ phân cách toàn trang hoặc viền lề A4.
    """
    pw, ph = page.rect.width, page.rect.height
    page_area = pw * ph if (pw > 0 and ph > 0) else 1.0

    try:
        dwgs = page.get_drawings()
    except Exception as e:
        logger.warning(f"Không thể đọc danh sách nét vẽ vector: {e}")
        return []

    if not dwgs:
        return []

    clusters: List[Dict[str, Any]] = []
    for d in dwgs:
        r = d.get("rect")
        if not r:
            continue
        # Bỏ qua nét quá nhỏ (dưới 2.5pt) hoặc đường kẻ trải dài > 88% kích thước trang
        if r.width < 2.5 and r.height < 2.5:
            continue
        if r.width >= pw * 0.88 or r.height >= ph * 0.88:
            continue

        items_cnt = len(d.get("items", [])) or 1
        r_expanded = fitz.Rect(r.x0 - gap, r.y0 - gap, r.x1 + gap, r.y1 + gap)
        merged = False
        for c in clusters:
            if c["expanded"].intersects(r_expanded):
                c["rect"] = c["rect"] | r
                c["expanded"] = fitz.Rect(
                    c["rect"].x0 - gap, c["rect"].y0 - gap,
                    c["rect"].x1 + gap, c["rect"].y1 + gap
                )
                c["count"] += 1
                c["items_count"] += items_cnt
                merged = True
                break
        if not merged:
            clusters.append({
                "rect": fitz.Rect(r),
                "expanded": r_expanded,
                "count": 1,
                "items_count": items_cnt
            })

    # Lọc các cụm nét vẽ đạt chuẩn hình vẽ
    valid_clusters = []
    for c in clusters:
        r = c["rect"]
        r_area = r.width * r.height
        # Cụm phải có ít nhất 2 nét vẽ (hoặc 1 hình phức hợp >= 2 items), kích thước tối thiểu min_dim pt và không chiếm quá 50% trang
        if (
            r.width >= min_dim and
            r.height >= min_dim and
            (c["count"] >= 2 or c["items_count"] >= 2) and
            (r_area / page_area) <= max_area_ratio
        ):
            valid_clusters.append(c)

    return valid_clusters


def extract_native_images_from_doc(
    doc: fitz.Document,
    session_dir: Path,
    min_dim: int = 35,
    max_area_ratio: float = 0.75
) -> Dict[str, Any]:
    """
    Quét toàn bộ trang trong PDF và trích xuất:
    1. Tất cả ảnh nhúng độc lập (PNG/JPG XObject) nét 100%.
    2. Tất cả cụm nét vẽ Vector Drawing kết xuất tại độ phân giải 300 DPI.
    Lưu ảnh vào `session_dir / native_images /` và lập chỉ mục `native_images_index.json`.

    Args:
        doc: Đối tượng fitz.Document đang mở.
        session_dir: Thư mục phiên làm việc của session.
        min_dim: Kích thước pixel tối thiểu.
        max_area_ratio: Tỷ lệ diện tích tối đa so với trang (nếu vượt quá -> là ảnh scan cả trang).

    Returns:
        Dict tổng hợp số lượng ảnh trích xuất được và danh sách chi tiết.
    """
    native_dir = session_dir / "native_images"
    native_dir.mkdir(parents=True, exist_ok=True)

    index_data: Dict[str, List[Dict[str, Any]]] = {}
    total_images_extracted = 0

    for page_num in range(len(doc)):
        page = doc[page_num]
        pw, ph = page.rect.width, page.rect.height
        page_area = pw * ph if (pw > 0 and ph > 0) else 1.0

        page_key = f"page_{page_num + 1}"
        index_data[page_key] = []

        try:
            img_list = page.get_images(full=True)
        except Exception as e:
            logger.warning(f"Không thể đọc danh sách ảnh trang {page_num + 1}: {e}")
            img_list = []

        is_scanned_page = False
        # Nếu trang chỉ có đúng 1 ảnh duy nhất, kiểm tra xem có phải là bản scan toàn trang không
        if len(img_list) == 1:
            xref = img_list[0][0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            if rects:
                r = rects[0]
                area = r.width * r.height
                if (area / page_area > 0.40) or (r.width >= pw * 0.75):
                    is_scanned_page = True
                    logger.info(f"Trang {page_num + 1}: Là bản scan ({area/page_area:.1%}), kích hoạt luồng scan fallback.")

        # NẾU LÀ TRANG SCAN TOÀN TRANG -> Không trích xuất native image, để dành cho smart_trim_figure
        if is_scanned_page:
            continue

        extracted_xrefs = set()

        # ═══════════ NHÁNH 1A: ẢNH BITMAP NHÚNG GỐC (XObject) ═══════════
        for img_idx, img_info in enumerate(img_list):
            xref = img_info[0]
            if xref in extracted_xrefs:
                continue

            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []

            if not rects:
                continue

            try:
                base_img = doc.extract_image(xref)
            except Exception as e:
                logger.warning(f"Lỗi trích xuất ảnh xref={xref}: {e}")
                continue

            if not base_img or "image" not in base_img:
                continue

            img_bytes = base_img["image"]
            ext = base_img.get("ext", "png").lower()
            if ext == "jpx":
                ext = "png"
            elif ext not in ("png", "jpg", "jpeg", "webp"):
                ext = "png"

            w = base_img.get("width", 0)
            h = base_img.get("height", 0)

            # Bỏ qua icon hoặc đường kẻ quá nhỏ
            if w < min_dim or h < min_dim:
                continue

            for r_idx, r in enumerate(rects):
                if r.width < 15 or r.height < 15:
                    continue

                r_area = r.width * r.height
                if (r_area / page_area > 0.40) or (r.width >= pw * 0.75):
                    continue

                # Chuẩn hóa tọa độ rect về thang 0-1000
                ymin = int(round((r.y0 / ph) * 1000))
                xmin = int(round((r.x0 / pw) * 1000))
                ymax = int(round((r.y1 / ph) * 1000))
                xmax = int(round((r.x1 / pw) * 1000))

                ymin = max(0, min(1000, ymin))
                xmin = max(0, min(1000, xmin))
                ymax = max(0, min(1000, ymax))
                xmax = max(0, min(1000, xmax))

                if ymax <= ymin or xmax <= xmin:
                    continue

                clip_rect = fitz.Rect(
                    max(0.0, r.x0),
                    max(0.0, r.y0),
                    min(pw, r.x1),
                    min(ph, r.y1)
                )
                if clip_rect.is_empty or clip_rect.width < 5 or clip_rect.height < 5:
                    continue

                fname = f"native_p{page_num + 1}_{xref}_{r_idx}.png"
                save_path = native_dir / fname

                # Quy tắc 20: Thay vì dump byte thô extract_image bị mất smask (làm ảnh đen thui),
                # kết xuất trực tiếp clip từ trang PDF ở độ phân giải in ấn >= 300 DPI (bảo toàn smask & nền trắng A4).
                scale = max(300.0 / 72.0, min(600.0 / 72.0, w / r.width)) if r.width > 0 else (300.0 / 72.0)
                matrix = fitz.Matrix(scale, scale)

                rendered_ok = False
                try:
                    pix = page.get_pixmap(matrix=matrix, clip=clip_rect, alpha=False)
                    if pix.width >= 15 and pix.height >= 15:
                        pix.save(str(save_path))
                        rendered_ok = True
                except Exception as e:
                    logger.warning(f"Không thể render clip trang cho xref={xref}: {e}")

                if not rendered_ok:
                    # Fallback ghi byte thô nếu render clip bất thành
                    try:
                        save_path.write_bytes(img_bytes)
                    except Exception as e:
                        logger.error(f"Không thể ghi file ảnh {save_path}: {e}")
                        continue

                # Chốt chặn an toàn: Kiểm tra ảnh có bị đen xì hoặc đơn sắc hỏng không
                try:
                    with Image.open(save_path) as check_im:
                        if is_black_or_corrupted_image(check_im):
                            logger.warning(f"Ảnh {fname} bị đen xì (>70% pixel đen hoặc đơn sắc hỏng), loại bỏ khỏi native index.")
                            save_path.unlink(missing_ok=True)
                            continue
                        w_rendered, h_rendered = check_im.size
                except Exception as e:
                    logger.warning(f"Lỗi kiểm tra ảnh vừa lưu {save_path}: {e}")
                    continue

                item = {
                    "id": f"p{page_num + 1}_{xref}_{r_idx}",
                    "filename": fname,
                    "rel_path": f"native_images/{fname}",
                    "page_num": page_num + 1,
                    "box": [ymin, xmin, ymax, xmax],
                    "width": w_rendered,
                    "height": h_rendered,
                    "aspect_ratio": round(w_rendered / h_rendered, 2) if h_rendered > 0 else 1.0,
                    "ext": "png",
                    "xref": xref,
                    "origin": "bitmap",
                    "desc": "Ảnh gốc PDF"
                }
                index_data[page_key].append(item)
                total_images_extracted += 1

            extracted_xrefs.add(xref)

        # ═══════════ NHÁNH 1B: CỤM NÉT VẼ VECTOR DRAWING (300 DPI) ═══════════
        # Chỉ quét vector drawing nếu trang không phải là bản scan
        vector_clusters = _cluster_vector_drawings(page, gap=14.0, min_dim=min_dim)
        for v_idx, cluster in enumerate(vector_clusters):
            r = cluster["rect"]
            # Đệm an toàn 5pt xung quanh cụm vector để không lẹm chữ nhãn A, B, C, O, x, y
            clip_rect = fitz.Rect(
                max(0.0, r.x0 - 5.0),
                max(0.0, r.y0 - 5.0),
                min(pw, r.x1 + 5.0),
                min(ph, r.y1 + 5.0)
            )

            # Chuẩn hóa tọa độ rect về thang 0-1000
            ymin = int(round((clip_rect.y0 / ph) * 1000))
            xmin = int(round((clip_rect.x0 / pw) * 1000))
            ymax = int(round((clip_rect.y1 / ph) * 1000))
            xmax = int(round((clip_rect.x1 / pw) * 1000))

            ymin = max(0, min(1000, ymin))
            xmin = max(0, min(1000, xmin))
            ymax = max(0, min(1000, ymax))
            xmax = max(0, min(1000, xmax))

            if ymax <= ymin or xmax <= xmin:
                continue

            # Kiểm tra xem cụm vector này có trùng lặp với ảnh bitmap đã trích xuất không
            overlap_with_bitmap = False
            for existing in index_data[page_key]:
                if existing.get("origin") == "bitmap":
                    eymin, exmin, eymax, exmax = existing["box"]
                    # Nếu tâm cụm vector nằm hoàn toàn trong bitmap -> bỏ qua
                    center_y = (ymin + ymax) / 2.0
                    center_x = (xmin + xmax) / 2.0
                    if eymin <= center_y <= eymax and exmin <= center_x <= exmax:
                        overlap_with_bitmap = True
                        break

            if overlap_with_bitmap:
                continue

            # Render trực tiếp vùng vector ở độ phân giải 300 DPI
            try:
                matrix = fitz.Matrix(300.0 / 72.0, 300.0 / 72.0)
                pix = page.get_pixmap(matrix=matrix, clip=clip_rect, alpha=False)
                if pix.width < 25 or pix.height < 25:
                    continue

                fname = f"native_vec_p{page_num + 1}_{v_idx}.png"
                save_path = native_dir / fname
                pix.save(str(save_path))

                # Kiểm tra chốt chặn an toàn: Cụm vector có bị đen xì hoặc đơn sắc hỏng không
                try:
                    with Image.open(save_path) as check_im:
                        if is_black_or_corrupted_image(check_im):
                            logger.warning(f"Cụm vector {fname} bị đen xì hoặc hỏng, loại bỏ khỏi native index.")
                            save_path.unlink(missing_ok=True)
                            continue
                except Exception:
                    pass

                item = {
                    "id": f"vec_p{page_num + 1}_{v_idx}",
                    "filename": fname,
                    "rel_path": f"native_images/{fname}",
                    "page_num": page_num + 1,
                    "box": [ymin, xmin, ymax, xmax],
                    "width": pix.width,
                    "height": pix.height,
                    "aspect_ratio": round(pix.width / pix.height, 2) if pix.height > 0 else 1.0,
                    "ext": "png",
                    "origin": "vector",
                    "desc": "Hình vẽ vector PDF (300 DPI)"
                }
                index_data[page_key].append(item)
                total_images_extracted += 1
            except Exception as e:
                logger.warning(f"Không thể render cụm vector trang {page_num + 1} ({v_idx}): {e}")

    # Ghi file chỉ mục index
    index_file = session_dir / "native_images_index.json"
    try:
        index_file.write_text(json.dumps(index_data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        logger.error(f"Không thể ghi chỉ mục native images: {e}")

    return {
        "total_images": total_images_extracted,
        "index_file": str(index_file),
        "pages_with_images": sum(1 for k, v in index_data.items() if len(v) > 0),
        "data": index_data
    }


def find_matching_native_image(
    session_dir: Path,
    page_num: int,
    gemini_box: List[int],
    iou_thresh: float = 0.20,
    coverage_thresh: float = 0.40
) -> Optional[Dict[str, Any]]:
    """
    So khớp Bounding Box của Gemini với danh sách ảnh gốc (cả Bitmap và Vector) trong `native_images_index.json`.

    Args:
        session_dir: Thư mục session.
        page_num: Số thứ tự trang (1-based).
        gemini_box: [ymin, xmin, ymax, xmax] trên thang 0-1000.
        iou_thresh: Ngưỡng IoU tối thiểu để coi là trùng khớp.
        coverage_thresh: Ngưỡng diện tích ảnh gốc được bao phủ bởi Gemini box.

    Returns:
        Dict thông tin ảnh gốc khớp nhất, hoặc None nếu không khớp (sẽ fallback an toàn sang AI crop).
    """
    index_file = session_dir / "native_images_index.json"
    if not index_file.exists():
        return None

    try:
        index_data = json.loads(index_file.read_text(encoding="utf-8"))
    except Exception:
        return None

    page_key = f"page_{page_num}"
    native_candidates = index_data.get(page_key, [])
    if not native_candidates:
        return None

    g_ymin, g_xmin, g_ymax, g_xmax = gemini_box
    g_area = max(1, (g_ymax - g_ymin) * (g_xmax - g_xmin))

    best_match = None
    best_score = 0.0

    for cand in native_candidates:
        n_ymin, n_xmin, n_ymax, n_xmax = cand["box"]
        n_area = max(1, (n_ymax - n_ymin) * (n_xmax - n_xmin))

        # Tính phần giao
        inter_ymin = max(g_ymin, n_ymin)
        inter_xmin = max(g_xmin, n_xmin)
        inter_ymax = min(g_ymax, n_ymax)
        inter_xmax = min(g_xmax, n_xmax)

        if inter_ymax <= inter_ymin or inter_xmax <= inter_xmin:
            continue

        inter_area = (inter_ymax - inter_ymin) * (inter_xmax - inter_xmin)
        union_area = g_area + n_area - inter_area
        iou = inter_area / union_area if union_area > 0 else 0.0
        coverage = inter_area / n_area if n_area > 0 else 0.0

        # Khoảng cách tâm theo trục Y (trục dọc câu hỏi)
        g_center_y = (g_ymin + g_ymax) / 2.0
        n_center_y = (n_ymin + n_ymax) / 2.0
        y_dist = abs(g_center_y - n_center_y)

        # Trọng số chấm điểm kết hợp: coverage (độ bao phủ) + iou
        # Ưu tiên nhẹ cho ảnh bitmap nếu điểm tương đương
        origin_bonus = 0.05 if cand.get("origin") == "bitmap" else 0.0
        score = (coverage * 0.7 + iou * 0.3) + origin_bonus

        if (coverage >= coverage_thresh or iou >= iou_thresh) and y_dist <= 260:
            if score > best_score:
                best_score = score
                best_match = cand

    if best_match:
        real_file = session_dir / best_match["rel_path"]
        if real_file.exists():
            matched_dict = dict(best_match)
            matched_dict["full_path"] = str(real_file)
            matched_dict["score"] = round(best_score, 3)
            return matched_dict

    return None
