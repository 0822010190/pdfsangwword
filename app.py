import streamlit as st
import base64
import io
from PIL import Image
from pdf2image import convert_from_bytes
from docx import Document
from openai import OpenAI

# --- Cấu hình ---
st.set_page_config(page_title="Vision2Word Math", layout="wide")
st.title("📄 Chuyển đổi Ảnh/PDF sang Word (SambaNova)")
st.info("Hỗ trợ: Công thức toán $...$, Dán ảnh trực tiếp (Ctrl+V), Xử lý file PDF.")

# --- Sidebar: Cấu hình API ---
with st.sidebar:
    st.header("Cấu hình")
    api_key = st.text_input("SambaNova API Key:", type="password")
    model_name = "Llama-3.2-11B-Vision-Instruct"

# --- Khởi tạo Client ---
client = None
if api_key:
    client = OpenAI(base_url="https://api.sambanova.ai/v1", api_key=api_key)

def process_image(img):
    """Gửi ảnh sang SambaNova và nhận văn bản"""
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()

    prompt = "Trích xuất văn bản và công thức toán học. BẮT BUỘC để các ký hiệu/công thức toán vào trong dấu $...$ (ví dụ $x^2 + y = 0$). Không giải thích thêm."
    
    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_str}"}}
                ]
            }],
            temperature=0.1
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Lỗi: {str(e)}"

# --- Giao diện chính ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("Đầu vào")
    # Tính năng dán ảnh (Ctrl+V)
    pasted_image = st.paste("Nhấn vào đây rồi Ctrl+V để dán ảnh")
    
    # Tính năng upload file
    uploaded_file = st.file_uploader("Hoặc tải lên file (Ảnh/PDF)", type=["png", "jpg", "jpeg", "pdf"])

images_to_process = []

if pasted_image:
    images_to_process.append(pasted_image)
    st.image(pasted_image, caption="Ảnh đã dán", use_container_width=True)

if uploaded_file:
    if uploaded_file.type == "application/pdf":
        pdf_images = convert_from_bytes(uploaded_file.read())
        images_to_process.extend(pdf_images)
        st.write(f"Đã nhận PDF: {len(pdf_images)} trang.")
    else:
        img = Image.open(uploaded_file)
        images_to_process.append(img)
        st.image(img, caption="Ảnh đã tải lên", use_container_width=True)

with col2:
    st.subheader("Kết quả (Word)")
    if st.button("🚀 Bắt đầu chuyển đổi") and client:
        if not images_to_process:
            st.error("Vui lòng dán ảnh hoặc tải file lên!")
        else:
            full_text = ""
            progress = st.progress(0)
            
            for i, img in enumerate(images_to_process):
                text = process_image(img)
                full_text += text + "\n\n"
                progress.progress((i + 1) / len(images_to_process))
            
            st.markdown(full_text)
            
            # Tạo file Word
            doc = Document()
            for line in full_text.split('\n'):
                doc.add_paragraph(line)
            
            word_io = io.BytesIO()
            doc.save(word_io)
            
            st.download_button(
                label="📥 Tải xuống file .docx",
                data=word_io.getvalue(),
                file_name="ket_qua_toan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    elif not api_key:
        st.warning("Vui lòng nhập API Key để tiếp tục.")
