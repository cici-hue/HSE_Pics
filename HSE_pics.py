import streamlit as st
import pdfplumber
from pdf2image import convert_from_bytes
import tempfile
import os
from PIL import Image
import zipfile
import json
import re

st.set_page_config(page_title="PDF缺陷提取器", layout="wide")
st.title("📄 PDF缺陷图片提取器（替代方案）")

st.markdown("""
这个版本使用pdfplumber和pdf2image库，不需要PyMuPDF。
功能：提取PDF中的文本和图片。
""")

uploaded_files = st.file_uploader(
    "上传PDF文件（支持多文件）",
    type=["pdf"],
    accept_multiple_files=True
)

def extract_text_near_image(page_text, search_radius=500):
    """在文本中查找缺陷信息"""
    lines = page_text.split('\n')
    
    for i, line in enumerate(lines):
        line_lower = line.lower()
        
        # 查找包含"defect code"的行
        if "defect code" in line_lower:
            # 提取缺陷代码
            code_match = re.search(r'defect code\s*[:=]?\s*(\d+)', line, re.IGNORECASE)
            if code_match:
                defect_code = code_match.group(1)
                
                # 查找原因（在后续行中）
                reason = "Unknown"
                for j in range(i+1, min(i+5, len(lines))):
                    if "defect" in lines[j].lower():
                        # 提取缺陷原因
                        reason_match = re.search(r'(.+?)\s+defect', lines[j], re.IGNORECASE)
                        if reason_match:
                            reason = reason_match.group(1).strip()
                        break
                
                return {
                    "defect_code": defect_code,
                    "reason": reason
                }
    
    return None

if uploaded_files:
    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("处理中..."):
            all_results = []
            
            for uploaded_file in uploaded_files:
                st.write(f"处理文件: {uploaded_file.name}")
                
                try:
                    # 使用pdfplumber提取文本
                    with pdfplumber.open(uploaded_file) as pdf:
                        for page_num, page in enumerate(pdf.pages):
                            # 提取文本
                            text = page.extract_text()
                            
                            if text:
                                # 查找缺陷信息
                                defect_info = extract_text_near_image(text)
                                
                                if defect_info:
                                    # 使用pdf2image转换当前页为图片
                                    images = convert_from_bytes(
                                        uploaded_file.getvalue(),
                                        first_page=page_num+1,
                                        last_page=page_num+1,
                                        dpi=150
                                    )
                                    
                                    if images:
                                        all_results.append({
                                            "file": uploaded_file.name,
                                            "page": page_num + 1,
                                            "defect_code": defect_info["defect_code"],
                                            "reason": defect_info["reason"],
                                            "image": images[0]  # 第一张图片
                                        })
                    
                    st.success(f"✓ {uploaded_file.name}: 处理完成")
                    
                except Exception as e:
                    st.error(f"❌ 处理 {uploaded_file.name} 时出错: {e}")
            
            # 显示结果
            if all_results:
                st.success(f"✅ 共找到 {len(all_results)} 个缺陷")
                
                # 创建ZIP文件
                with tempfile.TemporaryDirectory() as tmpdir:
                    # 按缺陷原因组织文件夹
                    for result in all_results:
                        folder_name = result["reason"].replace("/", "_").replace("\\", "_")[:50]
                        folder_path = os.path.join(tmpdir, folder_name)
                        os.makedirs(folder_path, exist_ok=True)
                        
                        # 保存图片
                        img_path = os.path.join(
                            folder_path,
                            f"{result['file']}_page{result['page']}_code{result['defect_code']}.jpg"
                        )
                        result["image"].save(img_path, "JPEG")
                    
                    # 创建ZIP
                    zip_path = os.path.join(tmpdir, "缺陷提取结果.zip")
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for root, dirs, files in os.walk(tmpdir):
                            for file in files:
                                if file.endswith('.zip'):
                                    continue
                                file_path = os.path.join(root, file)
                                arcname = os.path.relpath(file_path, tmpdir)
                                zipf.write(file_path, arcname)
                    
                    # 提供下载
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            "📦 下载所有提取结果",
                            f.read(),
                            file_name="缺陷提取结果.zip",
                            mime="application/zip"
                        )
                
                # 显示统计
                st.subheader("📊 提取结果")
                for result in all_results:
                    st.write(f"- **{result['reason']}** (代码: {result['defect_code']}) - {result['file']} 第{result['page']}页")
            else:
                st.warning("未找到缺陷信息")

else:
    st.info("请上传PDF文件开始处理")
