import streamlit as st
import pandas as pd
import google.generativeai as genai
from docx import Document
from io import BytesIO

# 🔑 ใส่ API Key ของ Google Gemini
API_KEY = "AIzaSyCd9GZQo30cjtd0fFOkFAZorJ5V_La-Nic"
genai.configure(api_key=API_KEY)  # ✅ ตั้งค่า API Key

# 📄 ฟังก์ชันอ่านไฟล์ Word / TXT และเก็บข้อความเป็น dictionary
def read_file(uploaded_files):
    texts = {}  # ✅ ใช้ dictionary เก็บข้อมูลแต่ละไฟล์
    for file in uploaded_files:
        if file.type == "text/plain":
            text = file.getvalue().decode("utf-8", errors="ignore")  # ป้องกัน error เมื่อมี character ที่ไม่สามารถอ่านได้
        elif file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            doc = Document(file)
            text = "\n".join([p.text for p in doc.paragraphs])
        else:
            text = ""
        
        texts[file.name] = text  # ✅ เก็บข้อมูลแยกตามชื่อไฟล์
    return texts

# ⚡ ฟังก์ชันเปรียบเทียบโดยใช้ Google Gemini API
def compare_features_gemini(texts, category):
    prompt = f"""
    เปรียบเทียบ '{category}' ของเครื่องมือแพทย์จากข้อความต่อไปนี้ 
    โดยแสดงผลในรูปแบบตารางเปรียบเทียบ และจัดให้หัวข้อย่อยในหมวดหมู่เดียวกัน:
    """

    for filename, text in texts.items():
        prompt += f"--- {filename} ---\n{text[:3000]}\n\n"  # ✅ จำกัดข้อความไม่เกิน 3000 ตัวอักษรต่อไฟล์

    try:
        # เรียกใช้ Google Gemini API ผ่าน client
        model = genai.GenerativeModel("gemini-2.0-flash")  # ✅ กำหนดโมเดลให้ถูกต้อง
        response = model.generate_content(prompt)
        if response and hasattr(response, "text"):
            return response.text
        else:
            return "⚠️ ไม่สามารถดึงข้อมูลจาก Gemini API ได้"
    except Exception as e:
        return f"⚠️ เกิดข้อผิดพลาด: {str(e)}"

# 🎨 Streamlit UI
st.title("🔬 แอปเปรียบเทียบคุณลักษณะเครื่องมือแพทย์ (ฺBME-Nonthavej)")

uploaded_files = st.file_uploader("📂 อัปโหลดไฟล์ Word หรือ TXT (2-3 ไฟล์)", 
                                  type=["docx", "txt"], accept_multiple_files=True)

category = st.selectbox("📊 เลือกหมวดหมู่การเปรียบเทียบ", ["คุณสมบัติทั่วไป", "คุณสมบัติเฉพาะ"])

if st.button("🔍 เปรียบเทียบ"):
    if 2 <= len(uploaded_files) <= 3:
        texts = read_file(uploaded_files)  # ✅ อ่านไฟล์และเก็บข้อมูลแยกตามชื่อไฟล์
        with st.spinner("กำลังเปรียบเทียบข้อมูล..."):
            result = compare_features_gemini(texts, category)

        # 📌 แปลงผลลัพธ์เป็น DataFrame
        cleaned_result = result.replace("&nbsp;", " ")  # ✅ ลบคำว่า &nbsp;
        data = [line.split("|")[1:-1] for line in cleaned_result.split("\n") if "|" in line]
        if data:
            df = pd.DataFrame(data[1:], columns=data[0])
            st.dataframe(df)

            # 📂 สร้างไฟล์ Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Comparison')

            # 📥 ปุ่มดาวน์โหลดผลลัพธ์
            st.download_button(
                label="📥 ดาวน์โหลดผลลัพธ์เป็น Excel",
                data=output.getvalue(),
                file_name="comparison_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ ไม่สามารถสร้างตารางเปรียบเทียบได้ กรุณาลองใหม่อีกครั้ง")
    else:
        st.warning("⚠️ โปรดอัปโหลดไฟล์จำนวน 2-3 ไฟล์เท่านั้น")