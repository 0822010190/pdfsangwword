import io
import os
import json
from typing import List, Tuple, Optional

import streamlit as st
import requests
from PIL import Image
import fitz  # PyMuPDF
import pytesseract

from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.shared import Pt

from streamlit_paste_button import paste_image_button as paste_btn


# =========================
# SambaNova (OpenAI-compatible)
# =========================
SAMBANOVA_BASE_URL = "https://api.sambanova.ai/v1"
DEFAULT_MODEL = "Meta-Llama-3.3-70B-Instruct"


def sambanova_chat(api_key: str, model: str, messages: List[dict], timeout: int = 90) -> str:
    if not api_key:
        raise RuntimeError("Chưa nhập SambaNova API Key.")

    url = f"{SAMBANOVA_BASE_URL}/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": 0.1, "max_tokens": 2048}

    r = requests.post(url, headers=headers, data=json.dumps(payload), timeout=timeout)
    if r.status_code != 200:
        raise RuntimeError(f"SambaNova API lỗi {r.status_code}: {r.text}")

    data = r.json()
    return data["choices"][0]["message"]["content"]


SYSTEM_RULES = """
Bạn là công cụ chuẩn hoá văn bản từ PDF/OCR để đưa vào Word.

YÊU CẦU NGHIÊM NGẶT:
1) GIỮ NGUYÊN NỘI DUNG: không thêm/bớt/suy diễn.
2) GIỮ NGUYÊN XUỐNG DÒNG: không tự gộp dòng, không tự ngắt dòng lại.
3) MỌI BIỂU THỨC/CT TOÁN PHẢI NẰM TRONG $...$.
   - Nếu đã là LaTeX thì vẫn phải bọc $...$ (nếu chưa bọc).
   - Không dùng \\( \\) hoặc \\[ \\] hoặc $$ $$.
4) Văn bản thường KHÔNG đặt trong $...$.
5) Chỉ trả về TEXT THUẦN, không markdown, không thêm tiêu đề.
""".strip()


def normalize_with_ai(api_key: str, model: str, raw_text: str) -> str:
    messages = [
        {"role": "system", "content": SYSTEM_RULES},
        {"role": "user", "content": raw_text},
    ]
    return sambanova_chat(api_key, model, messages).strip()


# =========================
# OCR (FAIL-SOFT)
# =========================
def ocr_image_pil(img: Image.Image) -> str:
    """OCR bằng Tesseract. Nếu thiếu Tesseract (Streamlit Cloud), trả về '' và KHÔNG làm app chết."""
    try:
        return pytesseract.image_to_string(img) or ""
    except Exception as e:
        st.warning(
            "⚠️ OCR không chạy được (thiếu Tesseract trên môi trường deploy). "
            "App sẽ bỏ qua OCR để không bị lỗi.\n"
            f"Chi tiết: {e}"
        )
        return ""


# =========================
# PDF extraction: "giống bản đầu"
# - Ưu tiên text layer
# - Trang không có text: render ảnh; OCR chỉ khi bật
# =========================
def render_page_image(page: fitz.Page, dpi: int = 200) -> Image.Image:
    zoom = dpi / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def extract_pdf_pages(
    pdf_bytes: bytes,
    dpi: int = 200,
    use_ocr: bool = False,
    text_min_chars: int = 30,
) -> List[dict]:
    """
    Trả về list page items:
      {
        "page_index": int,
        "text": str,               # text layer hoặc OCR (nếu bật)
        "has_text": bool,          # có text layer đủ ngưỡng
        "image": PIL.Image         # ảnh trang để đính kèm
      }
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = []

    for i in range(len(doc)):
        page = doc.load_page(i)

        text_layer = (page.get_text("text") or "").replace("\r\n", "\n").replace("\r", "\n").strip()
        has_text = len(text_layer) >= text_min_chars

        img = render_page_image(page, dpi=dpi)

        if has_text:
            text = text_layer
        else:
            text = ocr_image_pil(img).strip() if use_ocr else ""

        out.append({"page_index": i, "text": text, "has_text": has_text, "image": img})

    doc.close()
    return out


# =========================
# Images & clipboard
# =========================
def bytes_to_pil(b: bytes) -> Image.Image:
    return Image.open(io.BytesIO(b)).convert("RGB")


# =========================
# Word export
# =========================
def set_doc_default_font(doc: Document, font_name: str = "Times New Roman", font_size_pt: int = 13):
    style = doc.styles["Normal"]
    font = style.font
    font.name = font_name
    font.size = Pt(font_size_pt)
    # Ensure East Asia font as well
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def add_text_preserve_lines(doc: Document, text: str):
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    for line in text.split("\n"):
        doc.add_paragraph(line)


def build_docx_from_pdf_pages(
    title: str,
    pages: List[dict],
    normalized_texts: Optional[List[str]],
    embed_images: bool,
    page_image_width_in: float = 6.5,
) -> bytes:
    """
    Nếu normalized_texts != None: dùng text đã chuẩn hoá theo AI theo từng trang (cùng số lượng pages).
    """
    doc = Document()
    set_doc_default_font(doc, "Times New Roman", 13)

    if title.strip():
        doc.add_heading(title.strip(), level=1)

    for idx, p in enumerate(pages):
        page_no = p["page_index"] + 1

        doc.add_paragraph(f"--- Trang {page_no} ---")

        # Đính kèm ảnh trang (không crop; add_picture chỉ resize theo width)
        if embed_images and p.get("image") is not None:
            buf = io.BytesIO()
            p["image"].save(buf, format="PNG")
            buf.seek(0)
            doc.add_picture(buf, width=Inches(page_image_width_in))

        # Text cho trang
        text_to_write = ""
        if normalized_texts is not None:
            text_to_write = normalized_texts[idx] or ""
        else:
            text_to_write = p.get("text", "") or ""

        if text_to_write.strip():
            add_text_preserve_lines(doc, text_to_write)
        else:
            # Không có text => ghi chú rõ ràng để thầy thấy "không bị cắt", chỉ là trang scan chưa OCR
            if p.get("has_text", False):
                doc.add_paragraph("(Trang này có text layer nhưng không trích xuất được nội dung.)")
            else:
                doc.add_paragraph("(Trang dạng ảnh/scan: chưa có text. Bật OCR nếu muốn thử đọc chữ.)")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def build_docx_from_images(
    title: str,
    images: List[Image.Image],
    use_ocr: bool,
    use_ai: bool,
    api_key: str,
    model: str,
    embed_images: bool,
    image_width_in: float = 6.5,
) -> Tuple[bytes, str]:
    """
    Dùng cho ảnh upload/paste: OCR (nếu bật) -> AI (nếu bật).
    Trả về (docx_bytes, preview_text).
    """
    raw_texts = []
    for img in images:
        raw_texts.append(ocr_image_pil(img).strip() if use_ocr else "")

    raw_text = "\n\n".join([t for t in raw_texts if t]).strip()

    final_text = raw_text
    if use_ai and raw_text.strip():
        final_text = normalize_with_ai(api_key, model, raw_text)

    doc = Document()
    set_doc_default_font(doc, "Times New Roman", 13)

    if title.strip():
        doc.add_heading(title.strip(), level=1)

    if embed_images:
        doc.add_paragraph("Ảnh đính kèm:")
        for im in images:
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
            doc.add_picture(buf, width=Inches(image_width_in))

    doc.add_paragraph("Nội dung trích xuất:")
    if final_text.strip():
        add_text_preserve_lines(doc, final_text)
    else:
        doc.add_paragraph("(Chưa có text. Bật OCR để thử trích xuất chữ từ ảnh.)")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), final_text


# =========================
# UI
# =========================
st.set_page_config(page_title="PDF/Ảnh → Word (bản giống bản đầu + an toàn)", layout="wide")
st.title("📄 PDF/Ảnh → Word (.docx) — giống bản đầu nhưng an toàn")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    api_key = st.text_input("SambaNova API Key", type="password", value=os.getenv("SAMBANOVA_API_KEY", ""))
    model = st.text_input("Model", value=DEFAULT_MODEL)

    use_ai = st.checkbox("Dùng AI để ép công thức vào $...$ (khuyến nghị)", value=True)
    use_ocr = st.checkbox("Dùng OCR (cần Tesseract, Streamlit Cloud thường KHÔNG có)", value=False)

    embed_images = st.checkbox("Đính kèm ảnh trang vào Word", value=True)
    dpi = st.slider("DPI render PDF (để ảnh rõ hơn)", 120, 300, 200, 10)

    st.caption(
        "Logic giống bản đầu: ưu tiên text layer. Trang scan sẽ không OCR nếu tắt OCR, "
        "nhưng vẫn đính kèm ảnh + ghi chú để không mất trang."
    )

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("1) Nhập dữ liệu")
    st.markdown("**A) Dán ảnh (Ctrl+V)**: bấm nút rồi Ctrl+V.")
    paste = paste_btn("📋 Paste ảnh (Ctrl+V)")
    pasted_images: List[Image.Image] = []
    if paste.image_data is not None:
        pasted_images.append(paste.image_data)
        st.image(paste.image_data, caption="Ảnh dán", use_container_width=True)

    st.markdown("**B) Upload PDF / ảnh**")
    uploads = st.file_uploader("Chọn PDF hoặc ảnh", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

with col2:
    st.subheader("2) Xuất Word")
    title = st.text_input("Tiêu đề trong Word (tuỳ chọn)", value="")

    run = st.button("🚀 Chuyển đổi", type="primary")

    if run:
        if (not uploads) and (not pasted_images):
            st.error("Chưa có file/ảnh.")
            st.stop()

        pdf_pages_all: List[dict] = []
        img_only: List[Image.Image] = []

        # Collect uploads
        if uploads:
            for f in uploads:
                data = f.read()
                if f.name.lower().endswith(".pdf"):
                    with st.spinner(f"Đang xử lý PDF: {f.name} ..."):
                        pages = extract_pdf_pages(
                            data,
                            dpi=dpi,
                            use_ocr=use_ocr,      # an toàn: fail-soft nếu thiếu tesseract
                            text_min_chars=30     # "giống bản đầu": chỉ coi là có text khi đủ ngưỡng
                        )
                        pdf_pages_all.extend(pages)
                else:
                    img_only.append(bytes_to_pil(data))

        # Collect pasted images
        img_only.extend(pasted_images)

        # ===== PDF → Word =====
        if pdf_pages_all:
            # Chuẩn hoá AI theo từng trang (chỉ những trang có text)
            normalized_per_page: Optional[List[str]] = None
            if use_ai and api_key.strip():
                with st.spinner("Đang chuẩn hoá $...$ bằng SambaNova (chỉ trên phần text trích xuất) ..."):
                    normalized_per_page = []
                    for p in pdf_pages_all:
                        t = p.get("text", "") or ""
                        if t.strip():
                            normalized_per_page.append(normalize_with_ai(api_key, model, t))
                        else:
                            normalized_per_page.append("")  # trang scan chưa OCR => để trống
            else:
                normalized_per_page = None

            docx_bytes = build_docx_from_pdf_pages(
                title=title,
                pages=pdf_pages_all,
                normalized_texts=normalized_per_page,
                embed_images=embed_images,
                page_image_width_in=6.5,
            )

            # Preview nhanh: ghép text (để thầy thấy không chỉ có ảnh)
            preview_text = "\n\n".join([(normalized_per_page[i] if normalized_per_page else p["text"]) for i, p in enumerate(pdf_pages_all)]).strip()
            st.markdown("### Xem trước (text trích xuất từ PDF)")
            st.text_area("Preview", preview_text if preview_text else "(Không có text — PDF dạng scan. Bật OCR nếu muốn thử.)", height=260)

            st.success("Xong PDF → Word.")
            st.download_button(
                "⬇️ Tải Word từ PDF",
                data=docx_bytes,
                file_name="pdf_to_word.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        # ===== Images → Word =====
        if img_only:
            with st.spinner("Đang tạo Word từ ảnh ..."):
                docx_bytes2, preview2 = build_docx_from_images(
                    title=title if not pdf_pages_all else (title + " (Ảnh)"),
                    images=img_only,
                    use_ocr=use_ocr,
                    use_ai=use_ai and bool(api_key.strip()),
                    api_key=api_key,
                    model=model,
                    embed_images=embed_images,
                    image_width_in=6.5,
                )

            st.markdown("### Xem trước (text trích xuất từ ảnh)")
            st.text_area("Preview ảnh", preview2 if preview2 else "(Không có text — bật OCR để thử.)", height=220)

            st.success("Xong Ảnh → Word.")
            st.download_button(
                "⬇️ Tải Word từ Ảnh",
                data=docx_bytes2,
                file_name="images_to_word.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

st.divider()
st.caption(
    "Bản này cố ý giống bản đầu: ưu tiên text layer. "
    "Trang scan sẽ không làm app chết; nếu không OCR thì vẫn đính kèm ảnh và ghi chú."
)
