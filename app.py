import io
import re
import base64
from dataclasses import dataclass
from typing import List, Tuple

import streamlit as st
from openai import OpenAI
import fitz  # PyMuPDF
from PIL import Image

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn


# =========================
# Settings & Helpers
# =========================

DEFAULT_BASE_URL = "https://api.sambanova.ai/v1"
DEFAULT_VISION_MODEL = "Llama-4-Maverick-17B-128E-Instruct"  # theo ví dụ vision docs
DEFAULT_TEXT_MODEL = "Meta-Llama-3.3-70B-Instruct"

SYSTEM_PROMPT = """Bạn là trợ lý chuyển đổi tài liệu Toán sang văn bản gõ lại.
YÊU CẦU NGHIÊM NGẶT:
1) Mọi công thức/toán học phải nằm trong dấu $...$ (inline math).
2) TUYỆT ĐỐI KHÔNG có ký tự xuống dòng bên trong $...$.
3) Không tự ý sắp xếp lại thứ tự, không đổi số câu, không gộp/tách câu.
4) Giữ xuống dòng lời giải hợp lí (giống bố cục gốc), nhưng không đưa tab \\t.
5) Nếu không chắc một ký hiệu/toán tử, hãy giữ nguyên như nhìn thấy.
6) Đầu ra chỉ là NỘI DUNG (plain text), không thêm tiêu đề/giải thích ngoài lề.
"""

VISION_USER_INSTRUCTION = """Hãy đọc chính xác nội dung trong ảnh và gõ lại.
- Giữ nguyên thứ tự dòng/ý/câu.
- Với mọi biểu thức toán học: bọc vào $...$ và đảm bảo không có xuống dòng trong $...$.
- Không dùng \\(\\), \\[\\], $$...$$; chỉ dùng $...$.
- Không dùng tab.
Trả về đúng nội dung đã gõ lại (plain text)."""

TEXT_CLEANUP_INSTRUCTION = """Bạn hãy chuẩn hóa lại văn bản sau cho đúng yêu cầu:
- Mọi công thức/toán học phải nằm trong $...$.
- Không có xuống dòng trong $...$.
- Không thêm/bớt ý, không đổi thứ tự.
- Không dùng tab.
Chỉ trả về văn bản đã chuẩn hóa."""


def make_client(api_key: str, base_url: str) -> OpenAI:
    return OpenAI(api_key=api_key, base_url=base_url)


def encode_image_bytes(img_bytes: bytes, mime: str) -> str:
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def strip_tabs(text: str) -> str:
    return text.replace("\t", " ").replace("\u000b", " ")


def collapse_newlines_inside_dollars(text: str) -> str:
    """
    Remove any newline characters inside $...$ blocks.
    If there are multiple math blocks, handle all.
    """
    def _fix_block(m: re.Match) -> str:
        inner = m.group(1)
        inner = inner.replace("\r", " ").replace("\n", " ")
        inner = re.sub(r"\s{2,}", " ", inner).strip()
        return f"${inner}$"

    # non-greedy match for $...$
    return re.sub(r"\$(.*?)\$", _fix_block, text, flags=re.DOTALL)


def final_sanitize(text: str) -> str:
    text = strip_tabs(text)
    text = collapse_newlines_inside_dollars(text)
    # tránh khoảng trắng thừa quá nhiều
    text = re.sub(r"[ \u00A0]{3,}", "  ", text)
    return text.strip()


def pil_to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def pdf_to_page_images(pdf_bytes: bytes, dpi: int = 220) -> List[Image.Image]:
    """
    Render PDF pages to PIL images using PyMuPDF (no poppler needed).
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images: List[Image.Image] = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for i in range(len(doc)):
        page = doc[i]
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        images.append(img)
    return images


def call_vision_transcribe(client: OpenAI, model: str, image_png_bytes: bytes) -> str:
    data_url = encode_image_bytes(image_png_bytes, "image/png")
    resp = client.chat.completions.create(
        model=model,
        temperature=0.2,
        max_tokens=3000,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": VISION_USER_INSTRUCTION},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            },
        ],
    )
    out = resp.choices[0].message.content or ""
    return final_sanitize(out)


def call_text_cleanup(client: OpenAI, model: str, raw_text: str) -> str:
    resp = client.chat.completions.create(
        model=model,
        temperature=0.1,
        max_tokens=3000,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": TEXT_CLEANUP_INSTRUCTION + "\n\n---\n\n" + raw_text},
        ],
    )
    out = resp.choices[0].message.content or ""
    return final_sanitize(out)


def build_docx(all_sections: List[Tuple[str, str]]) -> bytes:
    """
    all_sections: list of (title, content)
    Create a .docx with Times New Roman size 13.
    """
    doc = Document()

    # set default font = Times New Roman, size 13
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(13)
    # ensure East Asia font also set
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    for idx, (title, content) in enumerate(all_sections, start=1):
        if title:
            p = doc.add_paragraph()
            run = p.add_run(title)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

        # add content paragraph-by-paragraph
        # preserve line breaks: each line becomes its own paragraph
        lines = content.splitlines() if content else []
        if not lines:
            doc.add_paragraph("")
        else:
            for line in lines:
                # keep empty lines as blank paragraphs
                doc.add_paragraph(line)

        if idx != len(all_sections):
            doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Ảnh/PDF → Word (SambaNova)", layout="wide")

st.title("📄 Ảnh / PDF → Word (.docx) bằng SambaNova")
st.caption("Nghiêm ngặt: công thức toán nằm trong $...$ và không xuống dòng bên trong $...$.")

with st.sidebar:
    st.header("Cấu hình API")
    api_key = st.text_input("SambaNova API Key", type="password", placeholder="Nhập key của bạn…")
    base_url = st.text_input("Base URL", value=DEFAULT_BASE_URL)
    vision_model = st.text_input("Vision model", value=DEFAULT_VISION_MODEL)
    text_model = st.text_input("Text model (cleanup)", value=DEFAULT_TEXT_MODEL)
    dpi = st.slider("DPI render PDF", 120, 300, 220, 10)

st.subheader("Tải tệp")
uploads = st.file_uploader(
    "Chọn 1 hoặc nhiều tệp (PDF/PNG/JPG/JPEG)",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

do_cleanup = st.toggle("Chạy bước chuẩn hoá lại văn bản (khuyến nghị)", value=True)

if st.button("🚀 Chuyển sang Word", type="primary", disabled=(not uploads or not api_key)):
    client = make_client(api_key, base_url)

    all_sections: List[Tuple[str, str]] = []
    progress = st.progress(0)
    total_steps = sum([1 for _ in uploads])  # rough; we'll update with pages too
    done = 0

    for up in uploads:
        filename = up.name
        data = up.read()

        if filename.lower().endswith(".pdf"):
            st.write(f"### 📎 PDF: {filename}")
            pages = pdf_to_page_images(data, dpi=dpi)
            st.write(f"- Số trang: {len(pages)}")

            # transcribe each page
            page_texts: List[str] = []
            for i, img in enumerate(pages, start=1):
                with st.spinner(f"Đang đọc trang {i}/{len(pages)}…"):
                    png_bytes = pil_to_png_bytes(img)
                    t = call_vision_transcribe(client, vision_model, png_bytes)
                    page_texts.append(t)

            merged = "\n\n".join(page_texts).strip()
            if do_cleanup and merged:
                with st.spinner("Đang chuẩn hoá văn bản…"):
                    merged = call_text_cleanup(client, text_model, merged)

            all_sections.append((f"{filename}", merged))

        else:
            st.write(f"### 🖼️ Ảnh: {filename}")
            img = Image.open(io.BytesIO(data)).convert("RGB")
            png_bytes = pil_to_png_bytes(img)

            with st.spinner("Đang đọc ảnh…"):
                text = call_vision_transcribe(client, vision_model, png_bytes)

            if do_cleanup and text:
                with st.spinner("Đang chuẩn hoá văn bản…"):
                    text = call_text_cleanup(client, text_model, text)

            all_sections.append((f"{filename}", text))

        done += 1
        progress.progress(min(1.0, done / max(1, total_steps)))

    # build docx
    with st.spinner("Đang tạo file Word…"):
        docx_bytes = build_docx(all_sections)

    st.success("Xong! Tải file Word bên dưới.")
    st.download_button(
        "⬇️ Tải Word (.docx)",
        data=docx_bytes,
        file_name="output.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

elif not api_key:
    st.info("Nhập SambaNova API Key ở thanh bên trái để bắt đầu.")
