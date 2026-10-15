import os
import re
import tempfile

import fitz
import numpy as np
import win32com.client
PP_SAVE_AS_PDF = 32

def parse_pages(page_spec, total_pages):
    """
    "3,5-7" -> [2, 4, 5, 6]
    使用者输入页码从 1 开始。
    """
    if page_spec is None or str(page_spec).lower() == "all":
        return list(range(total_pages))

    result = []

    for part in str(page_spec).split(","):
        part = part.strip()

        if "-" in part:
            start, end = part.split("-", 1)
            start = int(start)
            end = int(end)

            result.extend(range(start - 1, end))
        else:
            result.append(int(part) - 1)

    result = sorted(set(result))

    for p in result:
        if p < 0 or p >= total_pages:
            raise ValueError(
                f"页码 {p + 1} 超出范围，PPT 共 {total_pages} 页"
            )

    return result

def ppt_to_pdf(ppt_path, pdf_path):
    """
    使用 Microsoft PowerPoint 原生导出 PDF。
    """
    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    powerpoint.Visible = 1

    presentation = None

    try:
        presentation = powerpoint.Presentations.Open(
            os.path.abspath(ppt_path),
            WithWindow=False
        )

        presentation.SaveAs(
            os.path.abspath(pdf_path),
            PP_SAVE_AS_PDF
        )

        return presentation.Slides.Count

    finally:
        if presentation is not None:
            presentation.Close()

        powerpoint.Quit()

def detect_content_rect(
    page,
    dpi=200,
    white_threshold=250,
    margin_pt=6
):
    """
    将 PDF 页面临时 rasterize，
    找到非白色像素 bounding box。

    最终裁剪仍使用原始 PDF，
    不会把 PDF 栅格化。
    """

    scale = dpi / 72.0

    pix = page.get_pixmap(
        matrix=fitz.Matrix(scale, scale),
        alpha=False
    )

    img = np.frombuffer(
        pix.samples,
        dtype=np.uint8
    ).reshape(pix.height, pix.width, pix.n)

    rgb = img[:, :, :3]

    # 任意 RGB 通道低于 threshold
    # 都认为属于有效内容
    mask = np.any(rgb < white_threshold, axis=2)

    ys, xs = np.where(mask)

    # 完全空白页
    if len(xs) == 0:
        return page.rect

    x0 = xs.min() / scale
    y0 = ys.min() / scale
    x1 = (xs.max() + 1) / scale
    y1 = (ys.max() + 1) / scale

    rect = fitz.Rect(
        x0 - margin_pt,
        y0 - margin_pt,
        x1 + margin_pt,
        y1 + margin_pt
    )

    # 防止超出原页面
    rect &= page.rect

    return rect

def crop_pdf_pages(
    input_pdf,
    output_pdf,
    pages,
    margin_pt=6,
    crop_mode="per_page",
    dpi=200,
    white_threshold=250
):
    src = fitz.open(input_pdf)

    crop_rects = {}

    for page_index in pages:
        page = src[page_index]

        crop_rects[page_index] = detect_content_rect(
            page,
            dpi=dpi,
            white_threshold=white_threshold,
            margin_pt=margin_pt
        )

    # 多页统一 CropBox
    if crop_mode == "union":
        x0 = min(r.x0 for r in crop_rects.values())
        y0 = min(r.y0 for r in crop_rects.values())
        x1 = max(r.x1 for r in crop_rects.values())
        y1 = max(r.y1 for r in crop_rects.values())

        common_rect = fitz.Rect(x0, y0, x1, y1)

        for page_index in pages:
            crop_rects[page_index] = common_rect

    elif crop_mode != "per_page":
        raise ValueError(
            "crop_mode 只能是 per_page 或 union"
        )

    output = fitz.open()

    for page_index in pages:
        crop = crop_rects[page_index]

        # 创建真正符合裁剪后尺寸的新 PDF 页面
        new_page = output.new_page(
            width=crop.width,
            height=crop.height
        )

        # 将原始 PDF 的矢量内容 clip 到新页面
        new_page.show_pdf_page(
            new_page.rect,
            src,
            page_index,
            clip=crop
        )

    output.save(
        output_pdf,
        garbage=4,
        deflate=True
    )

    output.close()
    src.close()


def export_ppt_to_cropped_pdf(
    ppt_path,
    pages="all",
    output_path=None,
    margin=6,
    crop_mode="per_page",
    dpi=200
):
    ppt_path = os.path.abspath(ppt_path)

    if not os.path.exists(ppt_path):
        raise FileNotFoundError(ppt_path)

    if output_path is None:
        base = os.path.splitext(ppt_path)[0]
        output_path = base + "_cropped.pdf"

    output_path = os.path.abspath(output_path)

    with tempfile.TemporaryDirectory() as tmpdir:

        temp_pdf = os.path.join(
            tmpdir,
            "full_presentation.pdf"
        )

        total_pages = ppt_to_pdf(
            ppt_path,
            temp_pdf
        )

        selected_pages = parse_pages(
            pages,
            total_pages
        )

        crop_pdf_pages(
            temp_pdf,
            output_path,
            selected_pages,
            margin_pt=margin,
            crop_mode=crop_mode,
            dpi=dpi
        )

    return {
        "output": output_path,
        "pages": [p + 1 for p in selected_pages],
        "crop_mode": crop_mode,
        "margin_pt": margin
    }


if __name__ == "__main__":
    # runing example
    result = export_ppt_to_cropped_pdf(
        ppt_path=r"你的ppt相对或者绝对地址",
        pages="28",
        output_path=r"输出地址，为空时默认为ppt输入地址",
        margin=1,
        crop_mode="per_page"
    )

    print(result)