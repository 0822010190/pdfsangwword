import io
import os
import json
from typing import List, Optional, Tuple

import streamlit as st
import requests
import fitz  # PyMuPDF
from PIL import Image
import pytesseract

from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

from streamlit_paste_button import paste_image_button as paste_btn


# ======================================================
# SambaNova (OpenAI-compatible)
# ======================================================
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
    return r.json()["choices"][0]["message"]["content"]


SYSTEM_RULES = """
Bạn là công cụ chuẩn hoá văn bản để đưa vào Word.

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
    messages = [{"role": "system", "content": SYSTEM_RULES}, {"role": "user", "content": raw_text}]
    return sambanova_chat(api_key, model, messages).strip()


# ======================================================
# OCR (fail-soft, không làm app chết)
# ======================================================
def ocr_image_pil(img: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(img) or ""
    except Exception as e:
        st.warning(
            "⚠️ OCR không chạy được (thiếu Tesseract trên môi trường deploy). "
            "Bỏ qua OCR để app không bị lỗi.\n"
            f"Chi tiết: {e}"
        )
        return ""


# ======================================================
# PDF: "phiên bản đầu tiên" = ưu tiên TEXT LAYER
# - Không render trang thành ảnh trừ khi user chọn đính kèm trang scan
# ======================================================
def render_page_image(page: fitz.Page, dpi: int = 200) -> Image.Image:
    zoom = dpi / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def extract_pdf_text_first(
    pdf_bytes: bytes,
    scan_handling: str = "note",   # "note" | "embed" | "ocr"
    dpi: int = 200,
) -> Tuple[str, List[Image.Image]]:
    """
    Trả về (text_all, scan_images_to_embed)

    scan_handling:
      - "note": không OCR, không nhét ảnh; chỉ ghi chú trang scan.
      - "embed": không OCR; chỉ đính kèm ảnh trang scan vào Word.
      - "ocr": OCR trang scan (nếu có Tesseract), nếu không thì note.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out_lines: List[str] = []
    scan_images: List[Image.Image] = []

    for i in range(len(doc)):
        page = doc.load_page(i)
        page_no = i + 1

        txt = (page.get_text("text") or "").replace("\r\n", "\n").replace("\r", "\n").strip()

        if txt:
            # Giống bản đầu: có text layer thì lấy thẳng
            out_lines.append(f"--- Trang {page_no} ---")
            out_lines.append(txt)
        else:
            # Trang scan/ảnh
            out_lines.append(f"--- Trang {page_no} ---")
            if scan_handling == "embed":
                img = render_page_image(page, dpi=dpi)
                scan_images.append(img)
                out_lines.append("(Trang dạng ảnh/scan: không có text layer. Đính kèm ảnh trang ở dưới.)")
            elif scan_handling == "ocr":
                img = render_page_image(page, dpi=dpi)
                ocr_txt = ocr_image_pil(img).strip()
                if ocr_txt:
                    out_lines.append(ocr_txt)
                else:
                    out_lines.append("(Trang dạng ảnh/scan: OCR không chạy hoặc không đọc được.)")
            else:
                out_lines.append("(Trang dạng ảnh/scan: không có text layer. Bật OCR hoặc Đính kèm ảnh trang nếu cần.)")

        out_lines.append("")  # dòng trống giữa các trang

    doc.close()
    return "\n".join(out_lines).strip(), scan_images


# ======================================================
# Word export
# ======================================================
def set_doc_default_font(doc: Document, font_name: str = "Times New Roman", font_size_pt: int = 13):
    style = doc.styles["Normal"]
    font = style.font
    font.name = font_name
    font.size = Pt(font_size_pt)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def add_text_preserve_lines(doc: Document, text: str):
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    for line in text.split("\n"):
        doc.add_paragraph(line)


def build_docx(
    title: str,
    final_text: str,
    scan_images: List[Image.Image],
    embed_scan_images: bool,
) -> bytes:
    doc = Document()
    set_doc_default_font(doc, "Times New Roman", 13)

    if title.strip():
        doc.add_heading(title.strip(), level=1)

    # Text trước (giống bản đầu)
    add_text_preserve_lines(doc, final_text)

    # Chỉ kèm ảnh nếu user chọn chế độ embed scan
    if embed_scan_images and scan_images:
        doc.add_page_break()
        doc.add_paragraph("Ảnh các trang scan/không có text layer:")
        for im in scan_images:
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
            doc.add_picture(buf, width=Inches(6.5))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ======================================================
# Streamlit UI
# ======================================================
st.set_page_config(page_title="PDF/Ảnh → Word (bản đầu tiên)", layout="wide")
st.title("📄 PDF/Ảnh → Word (.docx) — bản giống phiên bản đầu tiên (text-first)")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    api_key = st.text_input("SambaNova API Key", type="password", value=os.getenv("SAMBANOVA_API_KEY", ""))
    model = st.text_input("Model", value=DEFAULT_MODEL)

    use_ai = st.checkbox("Dùng AI ép công thức vào $...$", value=True)

    st.subheader("PDF scan/ảnh xử lý thế nào?")
    scan_mode = st.radio(
        "Chọn 1",
        options=[
            ("note", "Chỉ ghi chú (không OCR, không ảnh) — giống bản đầu nhất"),
            ("embed", "Đính kèm ảnh trang scan (không OCR)"),
            ("ocr", "OCR trang scan (cần Tesseract; nếu thiếu sẽ tự bỏ qua)"),
        ],
        index=0,
        format_func=lambda x: x[1],
    )[0]

    dpi = st.slider("DPI render (chỉ dùng khi scan_mode=embed/ocr)", 120, 300, 200, 10)

    st.caption("Mặc định: KHÔNG nhét ảnh vào Word, chỉ lấy TEXT layer như bản đầu.")

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("1) Upload PDF / dán ảnh")

    # Paste ảnh (tuỳ chọn)
    st.markdown("**Dán ảnh (Ctrl+V)**: bấm nút rồi Ctrl+V (tuỳ máy).")
    paste = paste_btn("📋 Paste ảnh (Ctrl+V)")
    pasted_img: Optional[Image.Image] = None
    if paste.image_data is not None:
        pasted_img = paste.image_data
        st.image(pasted_img, caption="Ảnh dán", use_container_width=True)

    st.markdown("**Upload PDF** (ưu tiên loại có text layer):")
    pdf_file = st.file_uploader("Chọn file PDF", type=["pdf"])

with col2:
    st.subheader("2) Chuyển đổi & tải Word")
    title = st.text_input("Tiêu đề (tuỳ chọn)", value="")

    if st.button("🚀 Chuyển PDF → Word", type="primary"):
        if not pdf_file:
            st.error("Chưa chọn PDF.")
            st.stop()

        pdf_bytes = pdf_file.read()

        with st.spinner("Đang trích xuất TEXT layer từ PDF (giống bản đầu) ..."):
            raw_text, scan_images = extract_pdf_text_first(pdf_bytes, scan_handling=scan_mode, dpi=dpi)

        final_text = raw_text
        if use_ai and raw_text.strip():
            if not api_key.strip():
                st.warning("Chưa nhập API key nên bỏ qua AI, xuất text thô.")
            else:
                with st.spinner("Đang chuẩn hoá $...$ bằng SambaNova ..."):
                    final_text = normalize_with_ai(api_key, model, raw_text)

        st.markdown("### Xem trước (text)")
        st.text_area("Preview", final_text, height=360)

        docx_bytes = build_docx(
            title=title,
            final_text=final_text,
            scan_images=scan_images,
            embed_scan_images=(scan_mode == "embed"),
        )

        st.success("Xong! Tải file Word:")
        st.download_button(
            "⬇️ Download .docx",
            data=docx_bytes,
            file_name="pdf_to_word.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.divider()
    st.caption("Nếu PDF là scan (không có text layer) thì bản 'giống bản đầu' sẽ không thể ra chữ nếu không OCR.")
