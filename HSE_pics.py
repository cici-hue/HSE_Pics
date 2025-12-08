import fitz  # PyMuPDF
import re
import os
import json
import tempfile
import zipfile
import streamlit as st
from pathlib import Path
import shutil

def extract_defect_images(pdf_path, output_dir):
    """提取缺陷图片和原因"""
    os.makedirs(output_dir, exist_ok=True)
    
    doc = fitz.open(pdf_path)
    extracted_items = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        
        # 查找图片块
        image_blocks = [(i, block) for i, block in enumerate(blocks) if block["type"] == 1]
        
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
                        folder_name = re.sub(r'[<>:"/\\|?*]', '_', reason)
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
                        st.warning(f"提取图片失败: {e}")
    
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
        if block["type"] == 0:  # 文本块
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
    code_match = re.search(r'defect code\s+(\d+)', fifth_text, re.IGNORECASE)
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
    st.set_page_config(page_title="PDF缺陷提取器", layout="wide")
    
    st.title("📄 PDF缺陷图片提取器")
    st.markdown("""
    上传PDF文档，自动提取缺陷图片并按缺陷原因分类保存。
    **提取规则：**
    1. 跳过每页的第一张图片
    2. 从第二张图片开始，分析后面6个文本块
    3. 第5个文本块必须是"Defect Code"且有数字
    4. 第6个文本块的"Defect"之前的内容作为缺陷原因
    """)
    
    # 文件上传
    uploaded_files = st.file_uploader(
        "上传PDF文件（支持多文件）", 
        type=['pdf'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            all_results = []
            
            for uploaded_file in uploaded_files:
                st.subheader(f"处理文件: {uploaded_file.name}")
                
                # 保存上传的PDF到临时文件
                pdf_path = os.path.join(temp_dir, uploaded_file.name)
                with open(pdf_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # 为每个文件创建输出目录
                file_output_dir = os.path.join(temp_dir, f"output_{uploaded_file.name}")
                
                with st.spinner(f"正在处理 {uploaded_file.name}..."):
                    results = extract_defect_images(pdf_path, file_output_dir)
                    
                    if results:
                        all_results.extend(results)
                        
                        # 显示提取结果
                        st.success(f"✓ 提取完成！共找到 {len(results)} 个缺陷")
                        
                        # 显示文件夹结构
                        st.subheader("生成的文件夹结构：")
                        folders = set([r['folder'] for r in results])
                        for folder in folders:
                            folder_images = [r for r in results if r['folder'] == folder]
                            st.markdown(f"**📁 {folder}** - {len(folder_images)} 张图片")
                    else:
                        st.warning(f"未找到符合规则的缺陷项目")
            
            # 如果有结果，提供下载
            if all_results:
                st.divider()
                st.subheader("📥 下载提取结果")
                
                # 创建主输出目录
                main_output_dir = os.path.join(temp_dir, "所有缺陷")
                os.makedirs(main_output_dir, exist_ok=True)
                
                # 合并所有结果
                for result in all_results:
                    src_path = result['image_path']
                    dst_folder = os.path.join(main_output_dir, result['folder'])
                    os.makedirs(dst_folder, exist_ok=True)
                    
                    if os.path.exists(src_path):
                        shutil.copy2(src_path, dst_folder)
                
                # 创建ZIP文件
                zip_path = create_zip_folder(main_output_dir)
                
                # 提供下载按钮
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📦 下载所有文件（ZIP格式）",
                        data=f,
                        file_name="defect_images.zip",
                        mime="application/zip"
                    )
                
                # 显示统计信息
                st.info(f"**总计提取了 {len(all_results)} 个缺陷图片**")
                
                # 显示缺陷原因统计
                if all_results:
                    st.subheader("缺陷分类统计：")
                    reason_stats = {}
                    for item in all_results:
                        reason = item['reason']
                        reason_stats[reason] = reason_stats.get(reason, 0) + 1
                    
                    for reason, count in reason_stats.items():
                        st.write(f"- {reason}: {count} 个")

if __name__ == "__main__":
    main()