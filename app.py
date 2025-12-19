import io
import os
import json
from typing import List, Optional, Tuple

import streamlit as st
import requests
from PIL import Image
import fitz  # PyMuPDF
import pytesseract

from docx import Document
from docx.shared import Inches

from streamlit_paste_button import paste_image_button as paste_btn


# ======================================================
# CONFIG
# ======================================================
SAMBANOVA_BASE_URL = "https://api.sambanova.ai/v1"
DEFAULT_MODEL = "Meta-Llama-3.3-70B-Instruct"


# ======================================================
# SambaNova API (OpenAI-compatible)
# ======================================================
def sambanova_chat(api_key: str, model: str, messages: List[dict]) -> str:
    if not api_key:
        raise RuntimeError("Chưa nhập SambaNova API Key")

    url = f"{SAMBANOVA_BASE_URL}/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0.1,
        "max_tokens": 2048
    }

    r = requests.post(url, headers=headers, data=json.dumps(payload), timeout=90)
    if r.status_code != 200:
        raise RuntimeError(f"SambaNova API lỗi {r.status_code}: {r.text}")

    return r.json()["choices"][0]["message"]["content"]


# ======================================================
# OCR (FAIL-SOFT – KHÔNG LÀM APP CHẾT)
# ======================================================
def ocr_image_pil(img: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(img) or ""
    except Exception as e:
        st.warning(
            "⚠️ OCR không chạy được (thiếu Tesseract trên môi trường deploy).\n"
            "→ Bỏ qua OCR, vẫn xuất Word kèm ảnh/PDF.\n\n"
            f"Chi tiết: {e}"
        )
        return ""


def extract_pdf_text_or_ocr(
    pdf_bytes: bytes,
    dpi: int = 200,
    use_ocr: bool = False
) -> Tuple[str, List[Image.Image]]:

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texts, images = [], []

    for i in range(len(doc)):
        page = doc.load_page(i)
        text = (page.get_text("text") or "").strip()

        zoom = dpi / 72
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        images.append(img)

        if len(text) >= 50:
            texts.append(text)
        else:
            texts.append(ocr_image_pil(img) if use_ocr else "")

    doc.close()
    return "\n\n".join(texts).strip(), images


# ======================================================
# PROMPT ÉP $...$ + GIỮ NGUYÊN DÒNG
# ======================================================
SYSTEM_RULES = """
Bạn là công cụ chuyển OCR/PDF sang văn bản Word.

YÊU CẦU BẮT BUỘC:
1. GIỮ NGUYÊN NỘI DUNG – không thêm bớt.
2. GIỮ NGUYÊN XUỐNG DÒNG – không gộp dòng.
3. MỌI CÔNG THỨC TOÁN PHẢI NẰM TRONG $...$
   - Phân số, căn, mũ, phương trình, ký hiệu toán học.
   - Nếu đã là LaTeX thì vẫn phải bọc $...$.
4. Văn bản thường KHÔNG đặt trong $...$.
5. Không markdown. Chỉ trả về TEXT THUẦN.
""".strip()


def normalize_with_ai(api_key: str, model: str, text: str) -> str:
    messages = [
        {"role": "system", "content": SYSTEM_RULES},
        {"role": "user", "content": text}
    ]
    return sambanova_chat(api_key, model, messages).strip()


# ======================================================
# WORD EXPORT
# ======================================================
def build_docx(
    title: str,
    text: str,
    images: List[Image.Image],
    embed_images: bool
) -> bytes:

    doc = Document()

    if title.strip():
        doc.add_heading(title, level=1)

    if embed_images and images:
        doc.add_paragraph("Ảnh / Trang PDF:")
        for img in images:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            doc.add_picture(buf, width=Inches(6.5))

    doc.add_paragraph("Nội dung trích xuất:")

    for line in text.replace("\r", "").split("\n"):
        doc.add_paragraph(line)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ======================================================
# STREAMLIT UI
# ======================================================
st.set_page_config(page_title="PDF / Ảnh → Word ($...$)", layout="wide")
st.title("📄 PDF / Ảnh → Word (.docx) bằng SambaNova")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    api_key = st.text_input(
        "SambaNova API Key",
        type="password",
        value=os.getenv("SAMBANOVA_API_KEY", "")
    )
    model = st.text_input("Model", value=DEFAULT_MODEL)

    use_ai = st.checkbox("Dùng AI ép công thức $...$", value=True)
    use_ocr = st.checkbox("Dùng OCR (cần Tesseract)", value=False)
    embed_images = st.checkbox("Đính kèm ảnh/PDF vào Word", value=True)

    st.caption("⚠️ Streamlit Cloud KHÔNG có Tesseract → OCR nên để TẮT")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Dán ảnh / Upload PDF")

    paste = paste_btn("📋 Paste ảnh (Ctrl+V)")
    pasted_images = []

    if paste.image_data is not None:
        pasted_images.append(paste.image_data)
        st.image(paste.image_data, caption="Ảnh dán", use_container_width=True)

    uploads = st.file_uploader(
        "Chọn ảnh hoặc PDF",
        type=["png", "jpg", "jpeg", "pdf"],
        accept_multiple_files=True
    )

with col2:
    st.subheader("2️⃣ Xử lý & Xuất Word")
    title = st.text_input("Tiêu đề Word (tuỳ chọn)", "")

    if st.button("🚀 CHUYỂN ĐỔI", type="primary"):

        raw_texts = []
        images = []

        if uploads:
            for f in uploads:
                data = f.read()
                if f.name.lower().endswith(".pdf"):
                    text, imgs = extract_pdf_text_or_ocr(
                        data, use_ocr=use_ocr
                    )
                    raw_texts.append(text)
                    images.extend(imgs)
                else:
                    img = Image.open(io.BytesIO(data)).convert("RGB")
                    images.append(img)
                    raw_texts.append(
                        ocr_image_pil(img) if use_ocr else ""
                    )

        for img in pasted_images:
            images.append(img)
            raw_texts.append(
                ocr_image_pil(img) if use_ocr else ""
            )

        raw_text = "\n\n".join(raw_texts).strip()

        if use_ai and raw_text:
            with st.spinner("Đang chuẩn hoá bằng SambaNova..."):
                raw_text = normalize_with_ai(api_key, model, raw_text)

        st.text_area("📄 Kết quả text", raw_text, height=350)

        docx_bytes = build_docx(
            title, raw_text, images, embed_images
        )

        st.download_button(
            "⬇️ Tải file Word",
            docx_bytes,
            file_name="output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.caption("© App tối ưu cho giáo viên Toán – công thức luôn nằm trong $...$")
