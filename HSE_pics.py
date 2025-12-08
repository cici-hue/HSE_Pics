import streamlit as st
import re
import os
import json
import tempfile
import zipfile
from pathlib import Path
import shutil
import traceback

# 尝试导入 fitz，如果失败则显示更友好的错误信息
try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError as e:
    st.error(f"❌ 导入 PyMuPDF 失败: {e}")
    st.info("""
    ⚠️ **依赖包问题解决步骤：**
    
    1. 请确保 requirements.txt 中包含 `PyMuPDF==1.23.8`
    2. 如果是本地运行：`pip install PyMuPDF`
    3. 如果问题持续，可能需要先安装系统依赖：
       - Ubuntu/Debian: `sudo apt-get install libmupdf-dev`
       - macOS: `brew install mupdf`
    4. 或者使用以下替代命令安装：
       ```
       pip install pymupdf
       ```
    """)
    FITZ_AVAILABLE = False

# 设置页面配置
st.set_page_config(
    page_title="PDF缺陷图片提取器",
    page_icon="📄",
    layout="wide"
)

def extract_defect_images(pdf_path, output_dir):
    """提取缺陷图片和原因"""
    if not FITZ_AVAILABLE:
        raise ImportError("PyMuPDF 不可用")
    
    doc = fitz.open(pdf_path)
    extracted_items = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        
        # 查找图片块
        image_blocks = [(i, block) for i, block in enumerate(blocks) if block.get("type") == 1]
        
        # 处理每个图片块（跳过第一个）
        for img_idx, (block_index, block) in enumerate(image_blocks):
            if img_idx == 0:  # 跳过第一张图片
                continue
            
            # 分析后面的6个文本块
            result = analyze_six_text_blocks(blocks, block_index)
            
            if result and "defect_code" in result:
                # 提取图片
                if block.get("images"):
                    try:
                        xref = block["images"][0][0]
                        base_image = doc.extract_image(xref)
                        
                        # 使用缺陷原因作为文件夹名
                        reason = result.get("reason", "Unknown")
                        # 清理文件夹名中的非法字符
                        folder_name = re.sub(r'[<>:"/\\|?*]', '_', reason)[:100]  # 限制长度
                        if not folder_name.strip():
                            folder_name = "Unknown_Defect"
                        
                        folder_path = os.path.join(output_dir, folder_name)
                        os.makedirs(folder_path, exist_ok=True)
                        
                        # 保存图片
                        img_filename = f"defect_p{page_num+1}_code{result['defect_code']}.{base_image['ext']}"
                        img_path = os.path.join(folder_path, img_filename)
                        
                        with open(img_path, "wb") as f:
                            f.write(base_image["image"])
                        
                        # 保存提取的信息
                        item = {
                            "page": page_num + 1,
                            "image_path": img_path,
                            "defect_code": result.get("defect_code", ""),
                            "reason": reason,
                            "folder": folder_name
                        }
                        
                        extracted_items.append(item)
                        
                    except Exception as e:
                        st.warning(f"提取图片失败: {str(e)[:100]}")
    
    doc.close()
    
    # 保存提取结果
    if extracted_items:
        json_path = os.path.join(output_dir, "extraction_results.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(extracted_items, f, ensure_ascii=False, indent=2)
    
    return extracted_items

def analyze_six_text_blocks(blocks, start_index):
    """分析图片块后面的6个文本块"""
    # 收集从图片块后面开始的连续6个文本块
    text_blocks = []
    current_index = start_index + 1
    
    while len(text_blocks) < 6 and current_index < len(blocks):
        block = blocks[current_index]
        if block.get("type") == 0:  # 文本块
            text = extract_text_from_block(block)
            if text.strip():
                text_blocks.append((current_index, text))
        current_index += 1
    
    if len(text_blocks) < 6:
        return None
    
    # 检查第5个文本块是否是Defect Code
    fifth_block_index, fifth_text = text_blocks[4]
    if "defect code" not in fifth_text.lower():
        return None
    
    # 提取缺陷代码
    code_match = re.search(r'defect code\s*[:=]?\s*(\d+)', fifth_text, re.IGNORECASE)
    if not code_match:
        return None
    
    defect_code = code_match.group(1)
    result = {"defect_code": defect_code}
    
    # 检查第6个文本块并提取原因
    sixth_block_index, sixth_text = text_blocks[5]
    
    # 提取"Defect"之前的字符串作为原因
    reason_match = re.search(r'(.+?)\s+defect', sixth_text, re.IGNORECASE)
    if reason_match:
        reason = reason_match.group(1).strip()
        result["reason"] = reason
    elif "defect" in sixth_text.lower():
        parts = re.split(r'\s+defect', sixth_text, flags=re.IGNORECASE)
        if parts and parts[0].strip():
            result["reason"] = parts[0].strip()
        else:
            return None
    else:
        return None
    
    return result

def extract_text_from_block(block):
    """从文本块中提取文本"""
    text = ""
    if "lines" in block:
        for line in block["lines"]:
            if "spans" in line:
                for span in line["spans"]:
                    text += span.get("text", "") + " "
    return text.strip()

def create_zip_folder(output_dir):
    """创建ZIP文件"""
    zip_path = output_dir + ".zip"
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, output_dir)
                zipf.write(file_path, arcname)
    return zip_path

def main():
    st.title("📄 PDF缺陷图片提取器")
    st.markdown("""
    上传PDF文档，自动提取缺陷图片并按缺陷原因分类保存。
    
    **提取规则：**
    1. 跳过每页的第一张图片
    2. 从第二张图片开始，分析后面6个文本块
    3. 第5个文本块必须是"Defect Code"且有数字
    4. 第6个文本块的"Defect"之前的内容作为缺陷原因
    """)
    
    if not FITZ_AVAILABLE:
        st.stop()
    
    # 文件上传
    uploaded_files = st.file_uploader(
        "上传PDF文件（支持多文件）", 
        type=['pdf'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        # 显示上传的文件信息
        st.success(f"已上传 {len(uploaded_files)} 个PDF文件")
        
        # 处理按钮
        if st.button("🚀 开始处理", type="primary"):
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                all_results = []
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, uploaded_file in enumerate(uploaded_files):
                    status_text.text(f"正在处理文件 {idx+1}/{len(uploaded_files)}: {uploaded_file.name}")
                    
                    # 保存上传的PDF到临时文件
                    pdf_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # 为每个文件创建输出目录
                    file_output_dir = os.path.join(temp_dir, f"output_{uploaded_file.name}")
                    
                    try:
                        results = extract_defect_images(pdf_path, file_output_dir)
                        
                        if results:
                            all_results.extend(results)
                            st.success(f"✓ {uploaded_file.name}: 提取到 {len(results)} 个缺陷")
                        else:
                            st.warning(f"⚠️ {uploaded_file.name}: 未找到符合规则的缺陷项目")
                        
                    except Exception as e:
                        st.error(f"❌ {uploaded_file.name}: 处理失败 - {str(e)}")
                    
                    # 更新进度
                    progress_bar.progress((idx + 1) / len(uploaded_files))
                
                # 如果有结果，提供下载
                if all_results:
                    st.divider()
                    st.subheader("📊 处理结果汇总")
                    
                    # 显示统计信息
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("总文件数", len(uploaded_files))
                    with col2:
                        st.metric("总缺陷数", len(all_results))
                    with col3:
                        unique_folders = len(set([r['folder'] for r in all_results]))
                        st.metric("缺陷类别", unique_folders)
                    
                    # 显示缺陷原因统计
                    if all_results:
                        st.subheader("缺陷分类详情：")
                        reason_stats = {}
                        for item in all_results:
                            reason = item['reason']
                            reason_stats[reason] = reason_stats.get(reason, 0) + 1
                        
                        for reason, count in sorted(reason_stats.items(), key=lambda x: x[1], reverse=True):
                            st.write(f"**📁 {reason}** - {count} 张图片")
                    
                    # 创建主输出目录
                    main_output_dir = os.path.join(temp_dir, "所有缺陷")
                    os.makedirs(main_output_dir, exist_ok=True)
                    
                    # 合并所有结果
                    for result in all_results:
                        src_path = result['image_path']
                        if os.path.exists(src_path):
                            dst_folder = os.path.join(main_output_dir, result['folder'])
                            os.makedirs(dst_folder, exist_ok=True)
                            shutil.copy2(src_path, dst_folder)
                    
                    # 创建ZIP文件
                    try:
                        zip_path = create_zip_folder(main_output_dir)
                        
                        # 提供下载按钮
                        with open(zip_path, "rb") as f:
                            zip_data = f.read()
                        
                        st.download_button(
                            label="📦 下载所有提取结果（ZIP格式）",
                            data=zip_data,
                            file_name="defect_images.zip",
                            mime="application/zip",
                            type="primary"
                        )
                        
                    except Exception as e:
                        st.error(f"创建下载包失败: {e}")
                else:
                    st.warning("⚠️ 所有文件均未找到符合规则的缺陷项目")
                
                status_text.text("✅ 处理完成！")

if __name__ == "__main__":
    main()
