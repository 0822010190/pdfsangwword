import os
import re
import io
import json
import base64
import requests
import streamlit as st
from PIL import Image
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches
from streamlit_paste_button import paste_image_button


# =========================
# SambaNova (OpenAI-compatible)
# =========================
SAMBANOVA_BASE_URL = "https://api.sambanova.ai/v1"  # SambaNova Cloud base URL :contentReference[oaicite:2]{index=2}
CHAT_COMPLETIONS_URL = f"{SAMBANOVA_BASE_URL}/chat/completions"

DEFAULT_MODEL = "Llama-4-Maverick-17B-128E-Instruct"  # có thể đổi theo model bạn thấy trong portal


# =========================
# Prompt nghiêm ngặt: LaTeX trong $...$
# =========================
SYSTEM_PROMPT = """Bạn là hệ thống OCR + chuyển đổi tài liệu Toán học sang văn bản tiếng Việt để đưa vào Microsoft Word.

RÀNG BUỘC BẮT BUỘC (KHÔNG ĐƯỢC VI PHẠM):
1) Mọi công thức toán học PHẢI đặt trong dấu $...$ (inline), không dùng \\(...\\), \\[...\\], $$...$$.
2) Giữ nguyên xuống dòng theo bố cục hợp lý của bài toán/lời giải. Không gộp dòng bừa bãi.
3) Không tự ý đánh lại số thứ tự câu nếu ảnh có số thứ tự.
4) Trả về DUY NHẤT JSON hợp lệ theo schema:
{
  "pages": [
    {
      "page_index": 1,
      "content": "văn bản đã OCR, có công thức trong $...$"
    }
  ]
}
5) Không thêm lời dẫn, không thêm markdown, không thêm giải thích ngoài JSON.
"""

USER_TASK = """Hãy đọc ảnh (có thể là đề Toán, có công thức, ký hiệu, hình/biểu thức).
- OCR chính xác tối đa.
- Với ký hiệu toán: chuyển sang LaTeX và bắt buộc đặt trong $...$.
- Văn bản tiếng Việt đúng chính tả (nếu nhìn thấy).
- Kết quả trả về theo JSON đã quy định.
"""


# =========================
# Helpers
# =========================
def get_api_key() -> str:
    return (st.session_state.get("SAMBANOVA_API_KEY") or os.getenv("SAMBANOVA_API_KEY") or "").strip()


def image_bytes_to_data_url(img_bytes: bytes, mime: str = "image/png") -> str:
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def pil_to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def render_pdf_to_images(pdf_bytes: bytes, dpi: int = 200) -> list[bytes]:
    """
    Render PDF pages -> list of PNG bytes using PyMuPDF.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for i in range(doc.page_count):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out.append(pix.tobytes("png"))
    doc.close()
    return out


def call_sambanova_vision(image_png_bytes: bytes, model: str, api_key: str, temperature: float = 0.0) -> dict:
    """
    OpenAI multimodal format (text + image_url base64 data URL) :contentReference[oaicite:3]{index=3}
    """
    data_url = image_bytes_to_data_url(image_png_bytes, mime="image/png")

    payload = {
        "model": model,
        "temperature": temperature,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": USER_TASK},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            },
        ],
        "max_tokens": 3000,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    resp = requests.post(CHAT_COMPLETIONS_URL, headers=headers, data=json.dumps(payload), timeout=120)
    if resp.status_code != 200:
        raise RuntimeError(f"SambaNova API lỗi {resp.status_code}: {resp.text}")

    return resp.json()


def extract_json_from_model_text(text: str) -> dict:
    """
    Model được yêu cầu trả JSON thuần. Nhưng để chắc ăn:
    - tìm khối JSON lớn nhất
    """
    text = text.strip()
    # nếu đã là JSON
    try:
        return json.loads(text)
    except Exception:
        pass

    # tìm đoạn {...} lớn nhất
    m = re.search(r"\{[\s\S]*\}\s*$", text)
    if not m:
        raise ValueError("Không tìm thấy JSON trong phản hồi model.")
    return json.loads(m.group(0))


def enforce_math_dollars(s: str) -> str:
    """
    Hậu kiểm đơn giản:
    - đổi \\( ... \\) -> $...$
    - đổi \\[ ... \\] -> $...$
    - đổi $$...$$ -> $...$
    (Không “render”, chỉ chuẩn hoá dấu)
    """
    s = re.sub(r"\\\(([\s\S]*?)\\\)", r"$\1$", s)
    s = re.sub(r"\\\[([\s\S]*?)\\\]", r"$\1$", s)
    s = re.sub(r"\$\$([\s\S]*?)\$\$", r"$\1$", s)
    return s


def build_docx(pages: list[dict], images_per_page: list[bytes] | None = None, title: str = "Chuyển đổi") -> bytes:
    doc = Document()

    # Set default font Times New Roman size 13
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(13)

    doc.add_paragraph(title)

    for idx, page in enumerate(pages):
        page_index = page.get("page_index", idx + 1)
        content = page.get("content", "")

        content = enforce_math_dollars(content)

        doc.add_paragraph(f"\n--- Trang {page_index} ---\n")

        # giữ xuống dòng: mỗi dòng -> 1 paragraph
        for line in content.splitlines():
            # giữ dòng trống
            if line.strip() == "":
                doc.add_paragraph("")
            else:
                doc.add_paragraph(line)

        # chèn ảnh trang (tuỳ chọn)
        if images_per_page and idx < len(images_per_page):
            doc.add_paragraph("")
            try:
                doc.add_picture(io.BytesIO(images_per_page[idx]), width=Inches(6.2))
            except Exception:
                # nếu ảnh quá lớn/ lỗi thì bỏ qua
                pass

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# UI
# =========================
st.set_page_config(page_title="Ảnh/PDF → Word (SambaNova)", layout="wide")
st.title("📄 Ảnh / PDF → Word (.docx) bằng SambaNova (OCR + Toán trong $...$)")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    st.session_state["SAMBANOVA_API_KEY"] = st.text_input(
        "SambaNova API Key",
        value=st.session_state.get("SAMBANOVA_API_KEY", os.getenv("SAMBANOVA_API_KEY", "")),
        type="password",
        help="API key dùng ở server. Không nên hardcode.",
    )
    model = st.text_input("Model", value=DEFAULT_MODEL)
    temperature = st.slider("Temperature", 0.0, 1.0, 0.0, 0.1)
    dpi = st.slider("PDF render DPI", 120, 300, 200, 10)
    include_page_images = st.checkbox("Chèn ảnh gốc vào Word (mỗi trang)", value=False)

st.subheader("1) Dán ảnh bằng Ctrl+V hoặc tải file")
col1, col2 = st.columns(2)

with col1:
    pasted = paste_image_button("📋 Dán ảnh từ Clipboard (Ctrl+V)")
    pasted_img_bytes = None
    if pasted is not None:
        # pasted là PIL image
        pasted_img_bytes = pil_to_png_bytes(pasted)
        st.image(pasted, caption="Ảnh đã dán", use_container_width=True)

with col2:
    up = st.file_uploader("Tải lên ảnh hoặc PDF", type=["png", "jpg", "jpeg", "webp", "pdf"])
    uploaded_bytes = up.read() if up is not None else None

st.divider()
st.subheader("2) Chuyển đổi")

api_key = get_api_key()
if not api_key:
    st.warning("Bạn chưa nhập SambaNova API Key (ở sidebar).")

convert_btn = st.button("🚀 Chuyển sang Word", type="primary", disabled=not api_key)

if convert_btn and api_key:
    try:
        images = []
        if pasted_img_bytes:
            images = [pasted_img_bytes]
        elif uploaded_bytes and up is not None:
            if up.type == "application/pdf" or up.name.lower().endswith(".pdf"):
                images = render_pdf_to_images(uploaded_bytes, dpi=dpi)
            else:
                # ảnh thường
                img = Image.open(io.BytesIO(uploaded_bytes)).convert("RGB")
                images = [pil_to_png_bytes(img)]
        else:
            st.error("Hãy dán ảnh (Ctrl+V) hoặc tải file lên.")
            st.stop()

        st.info(f"Số trang/ảnh cần xử lý: {len(images)}")

        pages_out = []
        for i, img_bytes in enumerate(images, start=1):
            with st.spinner(f"Đang OCR + hiểu nội dung trang {i}..."):
                resp = call_sambanova_vision(img_bytes, model=model, api_key=api_key, temperature=temperature)
                # OpenAI-compatible: resp['choices'][0]['message']['content']
                content_text = resp["choices"][0]["message"]["content"]
                data = extract_json_from_model_text(content_text)

                # kỳ vọng data["pages"] có 1 page; nếu model trả nhiều, vẫn gom
                if "pages" in data and isinstance(data["pages"], list) and len(data["pages"]) > 0:
                    # nếu có nhiều pages, gán lại page_index hợp lệ
                    for p in data["pages"]:
                        if "page_index" not in p:
                            p["page_index"] = i
                        pages_out.append(p)
                else:
                    # fallback
                    pages_out.append({"page_index": i, "content": enforce_math_dollars(str(data))})

        # Sort theo page_index để ổn định
        pages_out.sort(key=lambda x: x.get("page_index", 0))

        st.success("Xử lý xong. Xem preview bên dưới.")
        for p in pages_out:
            st.markdown(f"### Trang {p.get('page_index')}")
            st.text(p.get("content", ""))

        docx_bytes = build_docx(
            pages_out,
            images_per_page=(images if include_page_images else None),
            title="Kết quả chuyển đổi (SambaNova OCR)",
        )

        st.download_button(
            "⬇️ Tải Word (.docx)",
            data=docx_bytes,
            file_name="ket-qua-chuyen-doi.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(f"Lỗi: {e}")
