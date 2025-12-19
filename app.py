# app.py
# Ảnh / PDF -> Word (.docx) bằng SambaNova (OCR + Toán trong $...$)
# - Ctrl+V dán ảnh (streamlit-paste-button) (sửa lỗi PasteResult)
# - Upload ảnh/PDF, render PDF -> ảnh
# - Retry khi 429 rate_limit_exceeded + backoff
# - Rate limit: sleep giữa các trang
# - Cho phép xử lý theo lô + resume (bắt đầu từ trang)
# - Cache theo hash ảnh để tránh gọi API lại
# - Xuất Word Times New Roman size 13

import os
import re
import io
import json
import base64
import time
import hashlib
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
SAMBANOVA_BASE_URL = "https://api.sambanova.ai/v1"
CHAT_COMPLETIONS_URL = f"{SAMBANOVA_BASE_URL}/chat/completions"

DEFAULT_MODEL = "Llama-4-Maverick-17B-128E-Instruct"


# =========================
# Prompt nghiêm ngặt: LaTeX trong $...$
# =========================
SYSTEM_PROMPT = r"""Bạn là hệ thống OCR + chuyển đổi tài liệu Toán học sang văn bản tiếng Việt để đưa vào Microsoft Word.

RÀNG BUỘC BẮT BUỘC (KHÔNG ĐƯỢC VI PHẠM):
1) Mọi công thức toán học PHẢI đặt trong dấu $...$ (inline). Tuyệt đối KHÔNG dùng \(...\), \[...\], $$...$$.
2) Giữ nguyên xuống dòng theo bố cục hợp lý của bài toán/lời giải. Không gộp dòng bừa bãi.
3) Không tự ý thay đổi / sắp lại số thứ tự câu nếu ảnh có số thứ tự.
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

USER_TASK = r"""Hãy đọc ảnh (có thể là đề Toán, có công thức, ký hiệu, hình/biểu thức).
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


def sha1_bytes(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()


def image_bytes_to_data_url(img_bytes: bytes, mime: str = "image/png") -> str:
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def pil_to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def render_pdf_to_images(pdf_bytes: bytes, dpi: int = 200) -> list[bytes]:
    """Render PDF pages -> list of PNG bytes using PyMuPDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out: list[bytes] = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for i in range(doc.page_count):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out.append(pix.tobytes("png"))

    doc.close()
    return out


def extract_json_from_model_text(text: str) -> dict:
    """Model được yêu cầu trả JSON thuần. Nhưng để chắc ăn: thử parse trực tiếp, nếu fail thì tìm khối JSON lớn nhất."""
    text = (text or "").strip()
    try:
        return json.loads(text)
    except Exception:
        pass

    m = re.search(r"\{[\s\S]*\}\s*$", text)
    if not m:
        raise ValueError("Không tìm thấy JSON trong phản hồi model.")
    return json.loads(m.group(0))


def enforce_math_dollars(s: str) -> str:
    """Chuẩn hoá dấu toán: \(..\), \[..], $$..$$ -> $..$"""
    if not s:
        return s
    s = re.sub(r"\\\(([\s\S]*?)\\\)", r"$\1$", s)
    s = re.sub(r"\\\[([\s\S]*?)\\\]", r"$\1$", s)
    s = re.sub(r"\$\$([\s\S]*?)\$\$", r"$\1$", s)
    return s


def build_docx(pages: list[dict], images_per_page: list[bytes] | None, title: str) -> bytes:
    doc = Document()

    # Default font Times New Roman size 13
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(13)

    doc.add_paragraph(title)

    for idx, page in enumerate(pages):
        page_index = page.get("page_index", idx + 1)
        content = enforce_math_dollars(page.get("content", "") or "")

        doc.add_paragraph("")
        doc.add_paragraph(f"--- Trang {page_index} ---")
        doc.add_paragraph("")

        for line in content.splitlines():
            if line.strip() == "":
                doc.add_paragraph("")
            else:
                doc.add_paragraph(line)

        if images_per_page and idx < len(images_per_page):
            doc.add_paragraph("")
            try:
                doc.add_picture(io.BytesIO(images_per_page[idx]), width=Inches(6.2))
            except Exception:
                pass

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def paste_result_to_pil(pasted) -> Image.Image | None:
    """streamlit-paste-button thường trả PasteResult, không phải PIL."""
    if pasted is None:
        return None

    if isinstance(pasted, Image.Image):
        return pasted.convert("RGB")

    if hasattr(pasted, "image") and getattr(pasted, "image") is not None:
        img = getattr(pasted, "image")
        if isinstance(img, Image.Image):
            return img.convert("RGB")

    if hasattr(pasted, "bytes") and getattr(pasted, "bytes"):
        b = getattr(pasted, "bytes")
        try:
            return Image.open(io.BytesIO(b)).convert("RGB")
        except Exception:
            pass

    if hasattr(pasted, "data") and getattr(pasted, "data"):
        b = getattr(pasted, "data")
        try:
            return Image.open(io.BytesIO(b)).convert("RGB")
        except Exception:
            pass

    return None


def call_sambanova_vision_with_retry(
    image_png_bytes: bytes,
    model: str,
    api_key: str,
    temperature: float,
    max_tokens: int,
    max_retries: int,
    base_sleep: float,
) -> dict:
    """
    Gọi SambaNova chat/completions (multimodal) + retry khi 429/5xx.
    - Backoff: base_sleep * 2^attempt + jitter nhỏ
    """
    data_url = image_bytes_to_data_url(image_png_bytes, mime="image/png")

    payload = {
        "model": model,
        "temperature": float(temperature),
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
        "max_tokens": int(max_tokens),
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    last_err = None
    for attempt in range(max_retries + 1):
        try:
            resp = requests.post(
                CHAT_COMPLETIONS_URL,
                headers=headers,
                data=json.dumps(payload),
                timeout=180,
            )

            if resp.status_code == 200:
                return resp.json()

            # 429 rate limit or 5xx
            if resp.status_code in (429, 500, 502, 503, 504):
                last_err = RuntimeError(f"SambaNova API lỗi {resp.status_code}: {resp.text}")
                if attempt < max_retries:
                    sleep_s = base_sleep * (2 ** attempt) + (0.1 * attempt)
                    time.sleep(sleep_s)
                    continue
                raise last_err

            # các lỗi khác: fail luôn
            raise RuntimeError(f"SambaNova API lỗi {resp.status_code}: {resp.text}")

        except requests.RequestException as e:
            last_err = e
            if attempt < max_retries:
                sleep_s = base_sleep * (2 ** attempt) + (0.1 * attempt)
                time.sleep(sleep_s)
                continue
            raise RuntimeError(f"Lỗi mạng khi gọi SambaNova: {e}")

    raise RuntimeError(f"Không gọi được SambaNova sau retry: {last_err}")


# =========================
# Streamlit State (cache)
# =========================
if "ocr_cache" not in st.session_state:
    # key: sha1(image_bytes) -> parsed page dict (content)
    st.session_state["ocr_cache"] = {}


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
        help="Không nên hardcode key. Nên đặt biến môi trường SAMBANOVA_API_KEY.",
    )

    model = st.text_input("Model", value=DEFAULT_MODEL)
    temperature = st.slider("Temperature", 0.0, 1.0, 0.0, 0.1)

    st.subheader("PDF")
    dpi = st.slider("PDF render DPI", 120, 300, 180, 10)

    st.subheader("Giới hạn tốc độ / Retry (để tránh 429)")
    per_page_sleep = st.slider("Sleep giữa các trang (giây)", 0.0, 5.0, 1.0, 0.1)
    max_retries = st.slider("Số lần retry khi 429", 0, 8, 5, 1)
    base_sleep = st.slider("Base sleep backoff (giây)", 0.5, 5.0, 1.0, 0.5)
    max_tokens = st.slider("Max tokens", 800, 6000, 2500, 100)

    st.subheader("Chạy theo lô / Resume")
    start_page = st.number_input("Bắt đầu từ trang số (1 = đầu)", min_value=1, value=1, step=1)
    max_pages = st.number_input("Xử lý tối đa N trang (0 = tất cả)", min_value=0, value=0, step=1)

    include_page_images = st.checkbox("Chèn ảnh gốc vào Word (mỗi trang)", value=False)

    if st.button("🧹 Xoá cache OCR"):
        st.session_state["ocr_cache"] = {}
        st.success("Đã xoá cache.")


st.subheader("1) Dán ảnh bằng Ctrl+V hoặc tải file")

col1, col2 = st.columns(2)

with col1:
    pasted = paste_image_button("📋 Dán ảnh từ Clipboard (Ctrl+V)")
    pasted_img_bytes = None

    if pasted is not None:
        img = paste_result_to_pil(pasted)
        if img is None:
            st.error("Không lấy được ảnh từ Clipboard. Hãy thử dán lại hoặc tải file lên.")
        else:
            pasted_img_bytes = pil_to_png_bytes(img)
            st.image(img, caption="Ảnh đã dán", use_container_width=True)

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
        images: list[bytes] = []

        # Ưu tiên ảnh dán
        if pasted_img_bytes:
            images = [pasted_img_bytes]
        elif uploaded_bytes and up is not None:
            if (up.type == "application/pdf") or up.name.lower().endswith(".pdf"):
                images = render_pdf_to_images(uploaded_bytes, dpi=int(dpi))
            else:
                img = Image.open(io.BytesIO(uploaded_bytes)).convert("RGB")
                images = [pil_to_png_bytes(img)]
        else:
            st.error("Hãy dán ảnh (Ctrl+V) hoặc tải file lên.")
            st.stop()

        total = len(images)

        # cắt theo start_page / max_pages
        sp = int(start_page)
        if sp < 1:
            sp = 1
        start_idx = sp - 1
        if start_idx >= total:
            st.error(f"Bắt đầu từ trang {sp} nhưng tài liệu chỉ có {total} trang.")
            st.stop()

        end_idx = total
        if int(max_pages) > 0:
            end_idx = min(total, start_idx + int(max_pages))

        images_slice = images[start_idx:end_idx]
        st.info(f"Số trang/ảnh cần xử lý: {len(images_slice)} (từ trang {start_idx+1} đến {end_idx})")

        progress = st.progress(0)
        status = st.empty()

        pages_out: list[dict] = []
        images_for_doc = []  # ảnh tương ứng pages_out nếu bật chèn ảnh

        for local_i, img_bytes in enumerate(images_slice, start=1):
            real_page_index = start_idx + local_i  # 1-based
            status.write(f"Đang xử lý trang {real_page_index}/{total}...")

            # cache theo hash ảnh
            key = sha1_bytes(img_bytes)
            if key in st.session_state["ocr_cache"]:
                page_dict = st.session_state["ocr_cache"][key]
                # đảm bảo page_index đúng theo vị trí hiện tại
                page_dict = dict(page_dict)
                page_dict["page_index"] = real_page_index
                pages_out.append(page_dict)
                if include_page_images:
                    images_for_doc.append(img_bytes)
            else:
                resp = call_sambanova_vision_with_retry(
                    image_png_bytes=img_bytes,
                    model=model,
                    api_key=api_key,
                    temperature=float(temperature),
                    max_tokens=int(max_tokens),
                    max_retries=int(max_retries),
                    base_sleep=float(base_sleep),
                )
                content_text = resp["choices"][0]["message"]["content"]
                data = extract_json_from_model_text(content_text)

                # gom output
                if "pages" in data and isinstance(data["pages"], list) and len(data["pages"]) > 0:
                    # lấy page đầu tiên cho ảnh này
                    p0 = data["pages"][0]
                    content = enforce_math_dollars(p0.get("content", "") or "")
                else:
                    content = enforce_math_dollars(str(data))

                page_dict = {"page_index": real_page_index, "content": content}
                st.session_state["ocr_cache"][key] = dict(page_dict)
                pages_out.append(page_dict)
                if include_page_images:
                    images_for_doc.append(img_bytes)

            # sleep để hạn chế 429
            if float(per_page_sleep) > 0:
                time.sleep(float(per_page_sleep))

            progress.progress(int((local_i / len(images_slice)) * 100))

        pages_out.sort(key=lambda x: x.get("page_index", 0))

        st.success("Xử lý xong. Preview nội dung (mỗi trang):")
        for p in pages_out:
            st.markdown(f"### Trang {p.get('page_index')}")
            st.text(p.get("content", ""))

        docx_bytes = build_docx(
            pages_out,
            images_per_page=(images_for_doc if include_page_images else None),
            title="Kết quả chuyển đổi (SambaNova OCR)",
        )

        st.download_button(
            "⬇️ Tải Word (.docx)",
            data=docx_bytes,
            file_name="ket-qua-chuyen-doi.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.info(
            "Nếu gặp 429: hãy tăng 'Sleep giữa các trang' (ví dụ 1.5–3s), "
            "giảm DPI (180→150), hoặc xử lý theo lô (ví dụ 2–3 trang/lần)."
        )

    except Exception as e:
        # Thông báo “dễ hiểu” khi rate limit
        msg = str(e)
        if "429" in msg or "rate_limit" in msg:
            st.error(
                "Bị giới hạn tốc độ (429 rate_limit_exceeded). "
                "Hãy tăng Sleep giữa các trang, giảm DPI, hoặc chạy ít trang hơn mỗi lần."
                f"\n\nChi tiết: {e}"
            )
        else:
            st.error(f"Lỗi: {e}")
