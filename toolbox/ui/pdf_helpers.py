# -*- coding: utf-8 -*-
import os


def stem_matches_row_code(stem_lower: str, code: str) -> bool:
    s, c = (stem_lower or "").strip().lower(), (code or "").strip().lower()
    if not c or not s:
        return False
    if s == c:
        return True
    if not s.startswith(c):
        return False
    if len(s) == len(c):
        return True
    tail = s[len(c):]
    if tail.startswith("-"):
        return True
    return bool(tail) and (not tail[0].isdigit())


def list_cert_stems_for_row(stems_lower: set[str]) -> list[str]:
    suffix = "-1"
    certs = []
    for item in stems_lower:
        if (item + suffix) not in stems_lower:
            continue
        if any((x + suffix) == item for x in stems_lower):
            continue
        certs.append(item)
    return sorted(set(certs))


def strip_code_prefix(orig_stem: str, code: str) -> str:
    o, c = (orig_stem or "").strip(), (code or "").strip()
    if not o or not c:
        return ""
    if o.lower() == c.lower():
        return ""
    if o.lower().startswith(c.lower()):
        return o[len(c):].lstrip("-_").strip()
    return ""


def build_tag_jobs(code: str, pdf_lower_to_orig: dict) -> list[tuple[str, str, str]]:
    code = str(code).strip()
    if not code or code.lower() in ("nan", "none"):
        return []
    stems = {k for k in pdf_lower_to_orig if stem_matches_row_code(k, code)}
    if not stems:
        return []
    certs = list_cert_stems_for_row(stems)
    jobs = []
    for cert in certs:
        wash = cert + "-1"
        if wash not in stems:
            continue
        main_f = pdf_lower_to_orig[cert]
        var_f = pdf_lower_to_orig[wash]
        orig_main = os.path.splitext(main_f)[0]
        suffix = strip_code_prefix(orig_main, code)
        out_base = f"{suffix}826" if suffix else "826"
        jobs.append((main_f, var_f, out_base))
    return jobs


def images_to_pdf_files(image_paths: list, out_pdf_path: str):
    """多图写入一个 PDF（依赖 PyMuPDF）。"""
    import fitz

    doc = fitz.open()
    try:
        for p in image_paths:
            if not os.path.isfile(p):
                continue
            with fitz.open(p) as imgdoc:
                pdfb = imgdoc.convert_to_pdf()
            with fitz.open("pdf", pdfb) as imgpdf:
                doc.insert_pdf(imgpdf)
        os.makedirs(os.path.dirname(out_pdf_path) or ".", exist_ok=True)
        doc.save(out_pdf_path)
    finally:
        doc.close()


def tag_font(size):
    from PIL import ImageFont

    for path in ("C:/Windows/Fonts/SIMHEI.TTF", "C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/msyhbd.ttc"):
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def pdf_to_image_safe(pdf_path):
    import fitz
    from PIL import Image

    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72))
        w, h = pix.width, pix.height
        samples = bytes(pix.samples)
        return Image.frombytes("RGB", [w, h], samples)
    finally:
        doc.close()


def resize_image_static(img, max_w, max_h):
    from PIL import Image

    ratio = min(max_w / img.width, max_h / img.height)
    return img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)


def compose_tag826_jpg(main_pdf_path, variant_pdf_path, price_str, out_jpg_path):
    """合格证 + 水洗唛 PDF → 吊牌图；价格在左侧条码区下方与左图居中对齐。"""
    from PIL import Image, ImageDraw

    img_main = pdf_to_image_safe(main_pdf_path)
    img_variant = pdf_to_image_safe(variant_pdf_path)
    canvas = Image.new("RGB", (750, 800), "white")
    draw = ImageDraw.Draw(canvas)
    img_main_r = resize_image_static(img_main, 340, 500)
    img_variant_r = resize_image_static(img_variant, 340, 500)
    x1 = (370 - img_main_r.width) // 2 + 20
    y1 = 350 - img_main_r.height // 2
    x2 = 380 + (370 - img_variant_r.width) // 2 - 20
    canvas.paste(img_main_r, (x1, y1))
    canvas.paste(img_variant_r, (x2, y1))
    font = tag_font(16)
    price_text = f"全国统一零售价：{price_str if str(price_str).strip() else '0'}"
    bbox = draw.textbbox((0, 0), price_text, font=font)
    tw = bbox[2] - bbox[0]
    cx = x1 + img_main_r.width // 2
    text_x = int(cx - tw // 2)
    text_y = y1 + img_main_r.height + 12
    draw.text((text_x, text_y), price_text, fill="black", font=font)
    os.makedirs(os.path.dirname(out_jpg_path), exist_ok=True)
    canvas.save(out_jpg_path, "JPEG", quality=95)
