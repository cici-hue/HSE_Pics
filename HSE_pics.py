import streamlit as st
import os
import tempfile
import json
import zipfile
import shutil
import sys
from pathlib import Path

# 尝试导入 PyMuPDF
try:
    import fitz
    FITZ_AVAILABLE = True
    st.success("✅ PyMuPDF 导入成功")
except ImportError:
    st.error("❌ PyMuPDF 导入失败")
    st.info("尝试安装中...")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pymupdf"])
        import fitz
        FITZ_AVAILABLE = True
        st.success("✅ PyMuPDF 安装成功")
    except:
        FITZ_AVAILABLE = False
        st.error("无法安装 PyMuPDF，请检查依赖")

st.set_page_config(
    page_title="PDF缺陷提取器",
    page_icon="📄",
    layout="wide"
)

st.title("📄 PDF缺陷图片提取器")
st.markdown("""
上传PDF文档，自动提取缺陷图片并按缺陷原因分类保存。
""")

# 显示环境信息
with st.expander("环境信息"):
    st.write(f"Python版本: {sys.version}")
    st.write(f"PyMuPDF可用: {FITZ_AVAILABLE}")

if not FITZ_AVAILABLE:
    st.error("应用无法启动，因为缺少必需的依赖包。")
    st.stop()

def analyze_text_blocks(blocks, start_index):
    """分析文本块寻找缺陷信息"""
    try:
        # 收集后面的文本块
        text_blocks = []
        current_index = start_index + 1
        
        while len(text_blocks) < 6 and current_index < len(blocks):
            block = blocks[current_index]
            if block.get("type") == 0:  # 文本块
                # 提取文本
                text = ""
                if "lines" in block:
                    for line in block["lines"]:
                        if "spans" in line:
                            for span in line["spans"]:
                                text += span.get("text", "") + " "
                if text.strip():
                    text_blocks.append(text.strip())
            current_index += 1
        
        if len(text_blocks) < 6:
            return None
        
        # 检查第5个文本块
        import re
        fifth_text = text_blocks[4].lower()
        if "defect code" in fifth_text:
            # 提取缺陷代码
            code_match = re.search(r'defect code\s*[:=]?\s*(\d+)', text_blocks[4], re.IGNORECASE)
            if code_match:
                defect_code = code_match.group(1)
                
                # 检查第6个文本块
                sixth_text = text_blocks[5]
                reason = "Unknown Defect"
                
                # 尝试提取原因
                reason_match = re.search(r'(.+?)\s+defect', sixth_text, re.IGNORECASE)
                if reason_match:
                    reason = reason_match.group(1).strip()
                elif "defect" in sixth_text.lower():
                    parts = re.split(r'\s+defect', sixth_text, flags=re.IGNORECASE)
                    if parts and parts[0].strip():
                        reason = parts[0].strip()
                
                # 清理原因字符串
                reason = reason.replace("/", "_").replace("\\", "_").replace(":", "_")
                reason = reason[:50]  # 限制长度
                
                return {
                    "defect_code": defect_code,
                    "reason": reason
                }
    except Exception as e:
        st.warning(f"分析文本块时出错: {e}")
    
    return None

def extract_defects_from_pdf(pdf_path):
    """从PDF提取缺陷"""
    results = []
    
    try:
        doc = fitz.open(pdf_path)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            blocks = page.get_text("dict")["blocks"]
            
            # 找到所有图片块
            image_blocks = [(i, block) for i, block in enumerate(blocks) 
                           if block.get("type") == 1]
            
            # 跳过第一张图片，处理后面的
            for img_idx, (block_idx, block) in enumerate(image_blocks):
                if img_idx == 0:
                    continue  # 跳过第一张
                
                # 分析文本块
                defect_info = analyze_text_blocks(blocks, block_idx)
                
                if defect_info:
                    # 提取图片
                    if block.get("images"):
                        xref = block["images"][0][0]
                        base_image = doc.extract_image(xref)
                        
                        results.append({
                            "page": page_num + 1,
                            "defect_code": defect_info["defect_code"],
                            "reason": defect_info["reason"],
                            "image_data": base_image["image"],
                            "image_ext": base_image["ext"]
                        })
        
        doc.close()
        return results
        
    except Exception as e:
        st.error(f"处理PDF时出错: {e}")
        return []

def main():
    # 文件上传
    uploaded_files = st.file_uploader(
        "选择PDF文件",
        type=["pdf"],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        if st.button("🚀 开始处理", type="primary"):
            with st.spinner("处理中..."):
                # 创建临时目录
                with tempfile.TemporaryDirectory() as temp_dir:
                    all_results = []
                    
                    for uploaded_file in uploaded_files:
                        st.write(f"处理: {uploaded_file.name}")
                        
                        # 保存PDF到临时文件
                        pdf_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(pdf_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        
                        # 提取缺陷
                        results = extract_defects_from_pdf(pdf_path)
                        
                        if results:
                            all_results.extend(results)
                            st.success(f"✓ 找到 {len(results)} 个缺陷")
                        else:
                            st.warning(f"未找到符合规则的缺陷")
                    
                    # 如果有结果，组织并打包
                    if all_results:
                        # 按原因创建文件夹
                        output_dir = os.path.join(temp_dir, "缺陷提取结果")
                        os.makedirs(output_dir, exist_ok=True)
                        
                        for i, result in enumerate(all_results):
                            # 创建文件夹
                            folder_name = f"{result['reason']}_代码{result['defect_code']}"
                            # 清理文件夹名
                            folder_name = "".join(c for c in folder_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                            folder_path = os.path.join(output_dir, folder_name)
                            os.makedirs(folder_path, exist_ok=True)
                            
                            # 保存图片
                            img_filename = f"page{result['page']}_code{result['defect_code']}.{result['image_ext']}"
                            img_path = os.path.join(folder_path, img_filename)
                            
                            with open(img_path, "wb") as f:
                                f.write(result['image_data'])
                        
                        # 创建ZIP文件
                        zip_path = os.path.join(temp_dir, "defect_images.zip")
                        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                            for root, dirs, files in os.walk(output_dir):
                                for file in files:
                                    file_path = os.path.join(root, file)
                                    arcname = os.path.relpath(file_path, output_dir)
                                    zipf.write(file_path, arcname)
                        
                        # 提供下载
                        with open(zip_path, "rb") as f:
                            st.download_button(
                                "📦 下载提取结果",
                                f.read(),
                                file_name="缺陷提取结果.zip",
                                mime="application/zip",
                                type="primary"
                            )
                        
                        # 显示统计
                        st.subheader("📊 提取统计")
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("总文件数", len(uploaded_files))
                        with col2:
                            st.metric("总缺陷数", len(all_results))
                        with col3:
                            reasons = len(set(r['reason'] for r in all_results))
                            st.metric("缺陷类型", reasons)
                        
                        # 显示详情
                        with st.expander("查看提取详情"):
                            for result in all_results:
                                st.write(f"**Page {result['page']}** - Code {result['defect_code']}: {result['reason']}")
                    else:
                        st.warning("⚠️ 没有找到任何缺陷")

if __name__ == "__main__":
    main()
