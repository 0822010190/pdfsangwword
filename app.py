import streamlit as st
import base64
import io
from PIL import Image
from pdf2image import convert_from_bytes
from docx import Document
from openai import OpenAI

# --- Cấu hình trang ---
st.set_page_config(page_title="Vision2Word - SambaNova", layout="centered")
st.title("📄 Image/PDF to Word (SambaNova)")
st.markdown("Chuyển đổi tài liệu chứa công thức toán học sang Word với chuẩn LaTeX $...$.")

# --- Nhập API Key ---
with st.sidebar:
    api_key = st.text_input("Nhập SambaNova API Key:", type="password")
    model_choice = "Llama-3.2-11B-Vision-Instruct" # Model hỗ trợ Vision của SambaNova

# --- Khởi tạo Client SambaNova (Dùng chung chuẩn OpenAI) ---
client = None
if api_key:
    client = OpenAI(
        base_url="https://api.sambanova.ai/v1",
        api_key=api_key
    )

def image_to_base64(image):
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode('utf-8')

def process_with_sambanova(base64_image):
    """Gửi ảnh đến SambaNova và yêu cầu trích xuất văn bản + LaTeX"""
    prompt = """Trích xuất toàn bộ văn bản từ hình ảnh này. 
    YÊU CẦU NGHIÊM NGẶT: 
    1. Mọi công thức toán học, ký hiệu toán học, biến số (ví dụ: x, y, delta) PHẢI được đặt trong dấu $...$ (ví dụ: $E=mc^2$).
    2. Giữ nguyên định dạng đoạn văn.
    3. Chỉ trả về văn bản trích xuất, không thêm lời dẫn."""
    
    response = client.chat.completions.create(
        model=model_choice,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                    },
                ],
            }
        ],
        temperature=0.1,
    )
    return response.choices[0].message.content

# --- Giao diện chính ---
uploaded_file = st.file_uploader("Chọn ảnh hoặc file PDF (Hỗ trợ dán ảnh từ Clipboard)", type=["png", "jpg", "jpeg", "pdf"])

# Hỗ trợ Ctrl+V: Streamlit file_uploader mặc định cho phép dán file ảnh từ clipboard 
# khi bạn click vào nó và nhấn Ctrl+V.

if uploaded_file is not None:
    images = []
    
    # Xử lý file đầu vào
    if uploaded_file.type == "application/pdf":
        pdf_pages = convert_from_bytes(uploaded_file.read())
        images.extend(pdf_pages)
    else:
        images.append(Image.open(uploaded_file))

    st.success(f"Đã tải lên {len(images)} trang.")
    
    if st.button("Bắt đầu chuyển đổi") and client:
        full_text = ""
        progress_bar = st.progress(0)
        
        for i, img in enumerate(images):
            with st.spinner(f"Đang xử lý trang {i+1}..."):
                b64_img = image_to_base64(img)
                extracted_text = process_with_sambanova(b64_img)
                full_text += extracted_text + "\n\n"
            progress_bar.progress((i + 1) / len(images))

        # Hiển thị kết quả tạm thời
        st.subheader("Văn bản đã trích xuất:")
        st.markdown(full_text)

        # Xuất file Word
        doc = Document()
        for line in full_text.split('\n'):
            doc.add_paragraph(line)
        
        bio = io.BytesIO()
        doc.save(bio)
        
        st.download_button(
            label="📥 Tải xuống file Word (.docx)",
            data=bio.getvalue(),
            file_name="converted_document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
elif not api_key:
    st.warning("Vui lòng nhập API Key ở thanh bên để bắt đầu.")
