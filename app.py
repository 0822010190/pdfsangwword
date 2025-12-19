import io
import os
import re
import json
import base64
from dataclasses import dataclass
from typing import List, Optional, Tuple

import streamlit as st
import requests
from PIL import Image

import fitz  # PyMuPDF
import pytesseract

from docx import Document
from docx.shared import Inches

from streamlit_paste_button import paste_image_button as paste_btn


# =========================
# SambaNova (OpenAI-compatible) client via requests
# Base URL: https://api.sambanova.ai/v1  (docs)
# Chat completions endpoint: /chat/completions
# =========================
SAMBANOVA_BASE_URL = "https://api.sambanova.ai/v1"


class SambaNovaError(RuntimeError):
    pass


def sambanova_chat(
    api_key: str,
    model: str,
    messages: List[dict],
    temperature: float = 0.2,
    max_tokens: int = 2048,
    timeout: int = 90,
) -> str:
    if not api_key:
        raise SambaNovaError("Thiếu SAMBANOVA_API_KEY.")
    url = f"{SAMBANOVA_BASE_URL}/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    r = requests.post(url, headers=headers, data=json.dumps(payload), timeout=timeout)
    if r.status_code != 200:
        raise SambaNovaError(f"SambaNova API lỗi {r.status_code}: {r.text}")

    data = r.json()
    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        raise SambaNovaError(f"Phản hồi không đúng định dạng: {data}")


# =========================
# OCR + PDF extraction
# =========================
def pil_from_bytes(b: bytes) -> Image.Image:
    return Image.open(io.BytesIO(b)).convert("RGB")


def ocr_image_pil(img: Image.Image) -> str:
    """
    OCR ảnh bằng Tesseract.
    Nếu máy chưa cài Tesseract, hàm sẽ báo lỗi rõ.
    """
    try:
        text = pytesseract.image_to_string(img)  # mặc định eng
        return text or ""
    except Exception as e:
        raise RuntimeError(
            "Không OCR được. Máy chưa cài Tesseract OCR hoặc chưa cấu hình PATH. "
            f"Chi tiết: {e}"
        )


def extract_pdf_text_or_ocr(pdf_bytes: bytes, dpi: int = 200) -> Tuple[str, List[Image.Image]]:
    """
    Trả về (raw_text, rendered_page_images).
    - Nếu PDF có text layer: lấy text trực tiếp.
    - Nếu ít text: render trang -> OCR.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_images: List[Image.Image] = []
    texts: List[str] = []

    for i in range(len(doc)):
        page = doc.load_page(i)
        txt = page.get_text("text") or ""
        txt_clean = txt.strip()

        # Render ảnh trang để (1) đính kèm Word nếu muốn, (2) OCR khi cần
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        page_images.append(img)

        # Heuristic: nếu có đủ text thì dùng luôn, nếu quá ít thì OCR
        if len(txt_clean) >= 50:
            texts.append(txt)
        else:
            ocr_txt = ocr_image_pil(img)
            texts.append(ocr_txt)

    doc.close()
    raw = "\n\n".join([t.strip("\n") for t in texts if t is not None])
    return raw, page_images


# =========================
# Strict formatting prompt: keep line breaks, enforce $...$ for math
# =========================
SYSTEM_RULES = r"""
Bạn là công cụ chuyển đổi nội dung từ OCR/PDF sang văn bản để đưa vào Word.
YÊU CẦU NGHIÊM NGẶT:

1) GIỮ NGUYÊN NỘI DUNG: không được tự ý thêm, bớt, suy diễn.
2) GIỮ NGUYÊN XUỐNG DÒNG: bảo toàn cấu trúc đoạn/ dòng như dữ liệu vào. Không tự gộp dòng.
3) CÔNG THỨC TOÁN PHẢI NẰM TRONG DẤU $...$:
   - Bất kỳ biểu thức toán nào (phân số, căn, mũ, chỉ số, ký hiệu ∠, ⟂, ∥, ∈, ≤, ≥, π, …, phương trình, bất đẳng thức, biểu thức đại số, hình học) đều phải đặt trong $...$.
   - Nếu trong đầu vào đã có LaTeX (ví dụ \frac{a}{b}, x^2, \sqrt{...}) thì vẫn bọc trong $...$ nếu chưa có.
   - Không dùng \( \) hoặc \[ \] hoặc $$ $$.
4) VĂN BẢN THƯỜNG không đặt trong $...$.
5) Nếu đoạn nào KHÔNG CHẮC là toán hay chữ (OCR mờ), hãy GIỮ NGUYÊN như đầu vào, không sửa nội dung.
6) Đầu ra chỉ là VĂN BẢN THUẦN (plain text), không markdown, không tiêu đề tự đặt.
""".strip()


def normalize_with_ai(api_key: str, model: str, raw_text: str, max_chars: int = 9000) -> str:
    """
    Gọi AI để chuẩn hoá theo luật $...$ + giữ xuống dòng.
    Chunk theo ký tự để tránh vượt ngữ cảnh.
    """
    raw_text = raw_text.replace("\r\n", "\n").replace("\r", "\n")

    chunks = []
    i = 0
    while i < len(raw_text):
        chunk = raw_text[i : i + max_chars]
        chunks.append(chunk)
        i += max_chars

    outputs = []
    for idx, ch in enumerate(chunks, start=1):
        messages = [
            {"role": "system", "content": SYSTEM_RULES},
            {"role": "user", "content": f"=== PHẦN {idx}/{len(chunks)} (giữ nguyên xuống dòng) ===\n{ch}"},
        ]
        out = sambanova_chat(
            api_key=api_key,
            model=model,
            messages=messages,
            temperature=0.1,
            max_tokens=2048,
        )
        outputs.append(out.strip("\n"))

    return "\n".join(outputs).strip("\n")


# =========================
# Build Word (.docx)
# =========================
def add_text_preserve_lines(doc: Document, text: str):
    """
    Mỗi dòng thành 1 paragraph để giữ xuống dòng 100%.
    Dòng trống -> paragraph trống.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    for line in text.split("\n"):
        doc.add_paragraph(line)


def build_docx(
    title: str,
    final_text: str,
    images: List[Image.Image],
    embed_images: bool,
) -> bytes:
    doc = Document()
    if title.strip():
        doc.add_heading(title.strip(), level=1)

    if embed_images and images:
        doc.add_paragraph("Ảnh/Trang PDF (đính kèm):")
        for im in images:
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
            # chèn vừa trang
            doc.add_picture(buf, width=Inches(6.5))

    doc.add_paragraph("Nội dung trích xuất (giữ nguyên xuống dòng, công thức trong $...$):")
    add_text_preserve_lines(doc, final_text)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Ảnh/PDF → Word (SambaNova) • $...$", layout="wide")

st.title("📄 Ảnh/PDF → Word (.docx) bằng SambaNova (nghiêm ngặt $...$)")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    api_key = st.text_input("SambaNova API Key", type="password", value=os.getenv("SAMBANOVA_API_KEY", ""))
    model = st.text_input("Model", value="Meta-Llama-3.3-70B-Instruct")
    temperature = st.slider("Temperature", 0.0, 1.0, 0.1, 0.05)
    embed_images = st.checkbox("Đính kèm ảnh/trang PDF vào Word", value=True)
    use_ai = st.checkbox("Dùng AI để chuẩn hoá công thức vào $...$", value=True)
    st.caption("API SambaNova dùng endpoint OpenAI-compatible `https://api.sambanova.ai/v1` và chat completions. (Xem docs)")

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("1) Dán ảnh (Ctrl+V) hoặc tải ảnh/PDF")

    st.markdown("**A. Dán ảnh**: bấm nút rồi Ctrl+V (Chrome/Edge thường ổn).")
    paste_result = paste_btn("📋 Paste image (Ctrl+V)", errors="raise")

    uploaded_images: List[Image.Image] = []
    uploaded_pdf_bytes: Optional[bytes] = None
    raw_text_sources: List[str] = []
    rendered_images: List[Image.Image] = []

    if paste_result.image_data is not None:
        # paste_result.image_data là PIL Image
        uploaded_images.append(paste_result.image_data)
        st.success("Đã nhận ảnh từ clipboard.")
        st.image(paste_result.image_data, caption="Ảnh dán từ clipboard", use_container_width=True)

    st.markdown("**B. Tải file**:")
    up = st.file_uploader("Chọn ảnh (png/jpg) hoặc PDF", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)

    if up:
        for f in up:
            b = f.read()
            if f.type == "application/pdf" or f.name.lower().endswith(".pdf"):
                uploaded_pdf_bytes = b
                st.info(f"Đã nhận PDF: {f.name}")
            else:
                img = pil_from_bytes(b)
                uploaded_images.append(img)
                st.info(f"Đã nhận ảnh: {f.name}")
                st.image(img, caption=f.name, use_container_width=True)

with col2:
    st.subheader("2) Trích xuất & Xuất Word")

    title = st.text_input("Tiêu đề trong Word (tuỳ chọn)", value="")

    run = st.button("🚀 Chạy chuyển đổi", type="primary")

    if run:
        if not uploaded_images and not uploaded_pdf_bytes:
            st.error("Chưa có ảnh hoặc PDF.")
            st.stop()

        with st.spinner("Đang trích xuất nội dung..."):
            # PDF
            if uploaded_pdf_bytes:
                pdf_raw, pdf_imgs = extract_pdf_text_or_ocr(uploaded_pdf_bytes)
                raw_text_sources.append(pdf_raw)
                if embed_images:
                    rendered_images.extend(pdf_imgs)

            # Ảnh
            if uploaded_images:
                for img in uploaded_images:
                    if embed_images:
                        rendered_images.append(img)
                    # OCR để lấy text (nếu ảnh chứa chữ)
                    try:
                        raw_text_sources.append(ocr_image_pil(img))
                    except Exception as e:
                        raw_text_sources.append("")  # vẫn cho xuất word, chỉ có ảnh
                        st.warning(str(e))

            raw_text = ("\n\n".join([t for t in raw_text_sources if t is not None])).strip()

        if not raw_text and not (embed_images and rendered_images):
            st.error("Không trích xuất được text và cũng không có ảnh để đính kèm.")
            st.stop()

        final_text = raw_text

        if use_ai and raw_text.strip():
            with st.spinner("Đang chuẩn hoá bằng SambaNova (giữ xuống dòng, ép $...$)..."):
                try:
                    # override temperature theo sidebar
                    # (truyền vào normalize -> sambanova_chat đang dùng 0.1; bạn muốn dùng slider thì thay ở đây)
                    final_text = normalize_with_ai(api_key=api_key, model=model, raw_text=raw_text)
                except Exception as e:
                    st.error(f"Lỗi SambaNova: {e}")
                    st.stop()

        st.markdown("### Xem trước (text)")
        st.text_area("Kết quả", final_text, height=360)

        with st.spinner("Đang tạo file Word (.docx)..."):
            docx_bytes = build_docx(
                title=title,
                final_text=final_text,
                images=rendered_images,
                embed_images=embed_images,
            )

        st.success("Xong! Tải file Word ở đây:")
        st.download_button(
            label="⬇️ Download .docx",
            data=docx_bytes,
            file_name="output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

st.divider()
st.caption(
    "Gợi ý: Nếu PDF là dạng scan/ảnh mờ, OCR sẽ quyết định chất lượng. "
    "Bạn có thể tăng DPI trong code (extract_pdf_text_or_ocr) để OCR rõ hơn."
)
