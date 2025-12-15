import streamlit as st
import fitz  # PyMuPDF
# import PyMuPDF  # PyMuPDF
import re
import zipfile
import tempfile
import os
from pathlib import Path
from collections import defaultdict, OrderedDict
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import hashlib
import io

# Set page configuration
st.set_page_config(
    page_title="HSE Defect Image Extraction and PPT Generation System",
    page_icon="📊",
    layout="wide"
)

class PDFDefectExtractor:
    """PDF Defect Extractor Class"""
    def __init__(self):
        self.extracted_items = []
    
    def extract_defects_from_pdf(self, pdf_file, filename):
        """Extract defect information from a single PDF file"""
        extracted_items = []
        
        try:
            pdf_bytes = pdf_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                blocks = page.get_text("dict")["blocks"]
                image_list = page.get_images(full=True)
                
                # Find all image blocks
                image_blocks = []
                for i, block in enumerate(blocks):
                    if block["type"] == 1:  # Image block
                        image_blocks.append({
                            "index": i,
                            "bbox": block["bbox"],
                            "y_position": block["bbox"][1]
                        })
                
                # Sort by y-coordinate (top to bottom)
                image_blocks.sort(key=lambda x: x["y_position"])
                
                # Process each image block (skip the first one)
                for block_idx, block_info in enumerate(image_blocks):
                    if block_idx == 0:  # Skip the first image
                        continue
                    
                    result = self._analyze_text_blocks(blocks, block_info["index"])
                    
                    if result and "defect_code" in result:
                        # Find the closest image based on block position
                        bbox = block_info["bbox"]
                        matched_image_idx = self._find_matching_image(page, bbox, image_list)
                        
                        if matched_image_idx is not None:
                            try:
                                xref = image_list[matched_image_idx][0]
                                base_image = doc.extract_image(xref)
                                
                                # Clean defect reason for filename
                                reason = result.get("reason", f"defect_{result['defect_code']}")
                                clean_reason = self._sanitize_filename(reason)
                                
                                if not clean_reason or clean_reason == "_":
                                    clean_reason = f"defect_{result['defect_code']}"
                                
                                extracted_items.append({
                                    "pdf_name": filename,
                                    "page": page_num + 1,
                                    "defect_code": result.get("defect_code", ""),
                                    "reason": reason,
                                    "clean_reason": clean_reason,
                                    "image_data": base_image["image"],
                                    "image_ext": base_image["ext"]
                                })
                                
                            except Exception as e:
                                st.warning(f"Failed to extract image: {e}")
                                continue
            
            doc.close()
        except Exception as e:
            st.error(f"Error processing PDF file {filename}: {str(e)}")
        
        return extracted_items
    
    def _analyze_text_blocks(self, blocks, start_index):
        """Analyze the 6 text blocks following an image block"""
        result = {}
        text_blocks = []
        current_index = start_index + 1
        
        while len(text_blocks) < 6 and current_index < len(blocks):
            block = blocks[current_index]
            if block["type"] == 0:  # Text block
                text = self._extract_text_from_block(block)
                if text.strip():
                    text_blocks.append(text)
            current_index += 1
        
        if len(text_blocks) < 6:
            return None
        
        # Check the 5th text block
        if "defect code" not in text_blocks[4].lower():
            return None
        
        # Extract defect code
        code_match = re.search(r'defect code\s+(\d+)', text_blocks[4], re.IGNORECASE)
        if not code_match:
            return None
        
        result["defect_code"] = code_match.group(1)
        
        # Extract reason
        if "defect" in text_blocks[5].lower():
            parts = re.split(r'\s+defect', text_blocks[5], flags=re.IGNORECASE)
            if parts and parts[0].strip():
                result["reason"] = parts[0].strip()
            else:
                return None
        else:
            return None
        
        return result
    
    def _find_matching_image(self, page, bbox, image_list):
        """Find matching image"""
        block_center_x = (bbox[0] + bbox[2]) / 2
        block_center_y = (bbox[1] + bbox[3]) / 2
        
        best_match_idx = None
        min_distance = float('inf')
        
        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            img_rects = page.get_image_rects(xref)
            
            if img_rects:
                img_rect = img_rects[0]
                img_center_x = (img_rect.x0 + img_rect.x1) / 2
                img_center_y = (img_rect.y0 + img_rect.y1) / 2
                
                distance = ((img_center_x - block_center_x) ** 2 + 
                           (img_center_y - block_center_y) ** 2) ** 0.5
                
                if distance < min_distance:
                    min_distance = distance
                    best_match_idx = img_idx
        
        return best_match_idx
    
    def _extract_text_from_block(self, block):
        """Extract text from text block"""
        text = ""
        if "lines" in block:
            for line in block["lines"]:
                if "spans" in line:
                    for span in line["spans"]:
                        text += span.get("text", "") + " "
        return text.strip()
    
    def _sanitize_filename(self, filename):
        """Clean filename"""
        if not filename:
            return "unknown"
        
        # Remove special characters
        illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*', '\n', '\r', '\t']
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        
        # Replace multiple underscores with single
        filename = re.sub(r'_{2,}', '_', filename)
        
        # Limit length
        if len(filename) > 100:
            filename = filename[:100]
        
        return filename.strip()

class PPTCreator:
    """PPT Generator Class"""
    def __init__(self):
        pass
    
    def create_ppt_from_images(self, all_defects, ppt_name="Defect_Report.pptx"):
        """Create PPT from extracted images"""
        if not all_defects:
            return None
        
        # Categorize images by defect reason
        defects_by_reason = OrderedDict()
        file_counter = defaultdict(int)
        
        for defect in all_defects:
            reason = defect['reason']
            clean_reason = self._sanitize_filename(reason)
            
            if not clean_reason or clean_reason == "_":
                clean_reason = f"defect_{defect.get('defect_code', 'unknown')}"
            
            # Handle duplicate filenames
            file_counter[clean_reason] += 1
            count = file_counter[clean_reason]
            
            if count > 1:
                clean_reason = f"{clean_reason}_{count}"
            
            if reason not in defects_by_reason:
                defects_by_reason[reason] = []
            
            defects_by_reason[reason].append({
                'order_number': defect.get('pdf_name', 'unknown').replace('.pdf', ''),
                'image_data': defect['image_data'],
                'image_ext': defect['image_ext'],
                'clean_name': clean_reason
            })
        
        # Sort by defect reason name
        defects_by_reason = OrderedDict(sorted(defects_by_reason.items()))
        
        # Create PPT
        return self._create_pptx_by_defect_reason(defects_by_reason, ppt_name)
    
    def _create_pptx_by_defect_reason(self, defects_by_reason, ppt_name):
        """Create PPT categorized by defect reason"""
        try:
            # Create PPT object
            prs = Presentation()
            
            # Set slide size (16:9)
            prs.slide_width = Inches(16)
            prs.slide_height = Inches(9)
            
            # Add title page
            self._add_title_page(prs, len(defects_by_reason), 
                               sum(len(images) for images in defects_by_reason.values()))
            
            # Add table of contents page
            self._add_table_of_contents(prs, defects_by_reason)
            
            # Create content for each defect type
            for defect_index, (defect_reason, images) in enumerate(defects_by_reason.items(), 1):
                # Add defect type title page
                self._add_defect_title_page(prs, defect_reason, defect_index, len(defects_by_reason))
                
                # Group images, 3 per group
                for i in range(0, len(images), 3):
                    img_group = images[i:i+3]
                    group_number = i // 3 + 1
                    total_groups = (len(images) - 1) // 3 + 1
                    
                    # Add image page
                    self._add_defect_images_page(prs, defect_reason, img_group, group_number, total_groups)
            
            # Add ending page
            self._add_ending_page(prs)
            
            # Save to memory
            ppt_buffer = io.BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)
            
            return ppt_buffer
            
        except Exception as e:
            st.error(f"Failed to create PPT: {str(e)}")
            return None
    
    def _add_title_page(self, prs, defect_types_count, total_images_count):
        """Add title page"""
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "Quality Defect Report"
        subtitle.text = f"By Defect Reason\n\n" \
                        f"Total Defect Types: {defect_types_count}\n" \
                        f"Total Images: {total_images_count}\n" \
                        f"Generated: {self._get_current_date()}"
        
        # Adjust subtitle font size
        for paragraph in subtitle.text_frame.paragraphs:
            paragraph.font.size = Pt(20)
    
    def _add_table_of_contents(self, prs, defects_by_reason):
        """Add table of contents page"""
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Add title
        left = Inches(0.5)
        top = Inches(0.5)
        width = Inches(15)
        height = Inches(1)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = "Table of Contents"
        title_frame.paragraphs[0].font.size = Pt(32)
        title_frame.paragraphs[0].font.bold = True
        
        # Add table of contents content
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(14)
        height = Inches(6)
        
        content_box = slide.shapes.add_textbox(left, top, width, height)
        content_frame = content_box.text_frame
        
        for i, (defect_reason, images) in enumerate(defects_by_reason.items(), 1):
            p = content_frame.add_paragraph()
            p.text = f"{i}. {defect_reason} ({len(images)} images)"
            p.font.size = Pt(20)
            p.level = 0
            p.space_after = Pt(5)
    
    def _add_defect_title_page(self, prs, defect_reason, defect_index, total_defects):
        """Add defect type title page"""
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Add title
        left = Inches(1)
        top = Inches(2)
        width = Inches(14)
        height = Inches(3)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        
        # Add defect type title
        p = title_frame.paragraphs[0]
        p.text = defect_reason
        p.font.size = Pt(44)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        
        # Add page number info
        p = title_frame.add_paragraph()
        p.text = f"Defect Type {defect_index} of {total_defects}"
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(100, 100, 100)
        p.alignment = PP_ALIGN.CENTER
    
    def _add_defect_images_page(self, prs, defect_reason, img_group, group_number, total_groups):
        """Add defect image page"""
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Add defect reason header (top)
        self._add_defect_header(slide, defect_reason, group_number, total_groups)
        
        # Add images
        self._add_images_with_order_numbers(slide, img_group)
        
        # Add page number
        self._add_page_number(slide, group_number, total_groups)
    
    def _add_defect_header(self, slide, defect_reason, group_number, total_groups):
        """Add header: defect reason"""
        left = Inches(0.5)
        top = Inches(0.2)
        width = Inches(15)
        height = Inches(0.8)
        
        header_box = slide.shapes.add_textbox(left, top, width, height)
        header_frame = header_box.text_frame
        
        # Add defect reason
        p = header_frame.paragraphs[0]
        p.text = f"Defect Reason: {defect_reason}"
        p.font.size = Pt(28)
        p.font.bold = True
        
        # Add group info (if needed)
        if total_groups > 1:
            p = header_frame.add_paragraph()
            p.text = f"Group {group_number} of {total_groups}"
            p.font.size = Pt(16)
            p.font.color.rgb = RGBColor(100, 100, 100)
    
    def _add_images_with_order_numbers(self, slide, img_group):
        """Add images and order numbers"""
        img_count = len(img_group)
        if img_count == 0:
            return
        
        # Set different layouts based on image count
        if img_count == 1:
            # 1 image: centered
            width = Inches(8)
            height = Inches(5.38)
            left = (Inches(16) - width) / 2
            top = Inches(1.8)
            
            positions = [(left, top, width, height)]
            
        elif img_count == 2:
            # 2 images: side by side
            width = Inches(6)
            height = Inches(5.38)
            total_width = 2 * width + Inches(1)
            start_left = (Inches(16) - total_width) / 2
            top = Inches(1.8)
            
            positions = [
                (start_left, top, width, height),
                (start_left + width + Inches(1), top, width, height)
            ]
            
        else:  # img_count == 3
            # 3 images: horizontal side by side, using new dimensions
            width = Inches(4.78)
            height = Inches(5.38)
            
            total_width = 3 * width + Inches(2 * 0.3)
            start_left = (Inches(16) - total_width) / 2
            top = Inches(1.8)
            
            positions = [
                (start_left, top, width, height),
                (start_left + width + Inches(0.3), top, width, height),
                (start_left + 2 * (width + Inches(0.3)), top, width, height)
            ]
        
        # Add images and order numbers
        for i, (img_info, (left, top, width, height)) in enumerate(zip(img_group, positions)):
            try:
                # Save image to temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{img_info['image_ext']}") as tmp_file:
                    tmp_file.write(img_info['image_data'])
                    tmp_file_path = tmp_file.name
                
                # Add order number (above image)
                self._add_order_number(slide, img_info['order_number'], left, top - Inches(0.4), width)
                
                # Add image
                slide.shapes.add_picture(tmp_file_path, left, top, width=width, height=height)
                
                # Delete temporary file
                os.unlink(tmp_file_path)
                
            except Exception as e:
                st.warning(f"Failed to add image: {e}")
    
    def _add_order_number(self, slide, order_number, left, top, width):
        """Add order number label"""
        height = Inches(0.3)
        
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        
        text_frame.text = f"Order No: {order_number}"
        text_frame.paragraphs[0].font.size = Pt(20)
        text_frame.paragraphs[0].font.bold = True
        text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 139)
        text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    def _add_page_number(self, slide, current_group, total_groups):
        """Add page number"""
        left = Inches(14.5)
        top = Inches(8.2)
        width = Inches(1)
        height = Inches(0.5)
        
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        
        text_frame.text = f"{current_group}/{total_groups}"
        text_frame.paragraphs[0].font.size = Pt(12)
        text_frame.paragraphs[0].font.color.rgb = RGBColor(150, 150, 150)
        text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    
    def _add_ending_page(self, prs):
        """Add ending page"""
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Add closing text
        left = Inches(2)
        top = Inches(3)
        width = Inches(12)
        height = Inches(3)
        
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        
        p = text_frame.paragraphs[0]
        p.text = "End of Report"
        p.font.size = Pt(36)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        
        p = text_frame.add_paragraph()
        p.text = "Quality Control Department"
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(100, 100, 100)
        p.alignment = PP_ALIGN.CENTER
    
    def _get_current_date(self):
        """Get current date"""
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d")
    
    def _sanitize_filename(self, filename):
        """Clean filename"""
        if not filename:
            return "unknown"
        
        illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*', '\n', '\r', '\t']
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        
        filename = re.sub(r'_{2,}', '_', filename)
        
        if len(filename) > 100:
            filename = filename[:100]
        
        return filename.strip()

def main():
    """Main application function"""
    st.title("📊 HSE Defect Pics Extraction and PPT Generation System")
    st.markdown("""
    ### Features:
    1. **Upload PDF files**: Upload HSE claim report PDF documents containing defect images
    2. **Automatic defect image extraction**: System automatically identifies and extracts defect images
    3. **Generate PPT report**: Automatically generates PPT report categorized by defect reason
    4. **Download results**: Download extracted images and generated PPT
    """)
    
    # Create two main function tabs
    tab1, tab2 = st.tabs(["📄 PDF Defect Extraction", "📊 PPT Generation"])
    
    with tab1:
        st.header("PDF Defect Image Extraction")
        uploaded_files = st.file_uploader(
            "Select PDF files (multiple allowed)",
            type="pdf",
            accept_multiple_files=True,
            key="pdf_uploader"
        )
        
        if uploaded_files:
            extractor = PDFDefectExtractor()
            all_defects = []
            
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            with st.spinner("Processing PDF files..."):
                for i, uploaded_file in enumerate(uploaded_files):
                    # Update progress
                    progress = (i) / len(uploaded_files)
                    progress_bar.progress(progress)
                    status_text.text(f"Processing: {uploaded_file.name} ({i+1}/{len(uploaded_files)})")
                    
                    # Extract defects
                    defects = extractor.extract_defects_from_pdf(uploaded_file, uploaded_file.name)
                    for defect in defects:
                        defect['pdf_file'] = uploaded_file.name
                        all_defects.append(defect)
                
                progress_bar.progress(1.0)
                status_text.text("Processing complete!")
            
            if all_defects:
                st.success(f"✅ Extraction complete! Found {len(all_defects)} defects")
                
                # Display statistics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("PDF Files", len(uploaded_files))
                with col2:
                    st.metric("Total Defects", len(all_defects))
                with col3:
                    # Count defect types
                    defect_types = len(set(d['reason'] for d in all_defects))
                    st.metric("Defect Types", defect_types)
                
                # Display defect details table
                st.subheader("📋 Defect Details")
                display_data = []
                for i, defect in enumerate(all_defects[:50], 1):  # Show max 50 entries
                    display_data.append({
                        "No.": i,
                        "PDF File": defect['pdf_name'],
                        "Page": defect['page'],
                        "Defect Code": defect.get('defect_code', 'N/A'),
                        "Defect Reason": defect['reason']
                    })
                
                st.dataframe(display_data, use_container_width=True)
                
                # Create ZIP file for download
                st.subheader("📥 Download Extracted Images")
                
                with tempfile.TemporaryDirectory() as tmpdir:
                    # Create ZIP file
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        # Create folders by PDF file
                        file_counter = defaultdict(int)
                        
                        for defect in all_defects:
                            pdf_name = Path(defect['pdf_name']).stem
                            reason = defect['clean_reason']
                            
                            # Handle duplicate filenames
                            file_counter[(pdf_name, reason)] += 1
                            count = file_counter[(pdf_name, reason)]
                            
                            if count == 1:
                                filename = f"{reason}.{defect['image_ext']}"
                            else:
                                filename = f"{reason}_{count}.{defect['image_ext']}"
                            
                            # Full ZIP path
                            zip_path = f"{pdf_name}/{filename}"
                            
                            # Add to ZIP
                            zip_file.writestr(zip_path, defect['image_data'])
                    
                    # Create download button
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📦 Download All Images (ZIP Format)",
                        data=zip_buffer,
                        file_name="extracted_defect_images.zip",
                        mime="application/zip",
                        help="Click to download ZIP file containing all extracted images"
                    )
                
                # Preview some images
                st.subheader("🖼️ Image Preview")
                preview_cols = st.columns(4)
                
                for idx, defect in enumerate(all_defects[:8]):  # Preview max 8 images
                    col_idx = idx % 4
                    with preview_cols[col_idx]:
                        # Display image
                        st.image(
                            defect['image_data'],
                            caption=f"{defect['reason']} (Page {defect['page']})",
                            use_container_width=True
                        )
                
                # Save extraction results to session state
                st.session_state.extracted_defects = all_defects
                st.success("✅ Extraction results saved, you can switch to the PPT Generation tab")
                
            else:
                st.warning("⚠️ No defect information found")
    
    with tab2:
        st.header("PPT Report Generation")
        
        if 'extracted_defects' not in st.session_state or not st.session_state.extracted_defects:
            st.info("👈 Please upload and extract PDF files in the left tab first")
        else:
            st.success(f"✅ Loaded {len(st.session_state.extracted_defects)} defects")
            
            # PPT options
            col1, col2 = st.columns(2)
            with col1:
                ppt_name = st.text_input("PPT File Name", "Defect_Report.pptx")
            with col2:
                ppt_layout = st.selectbox(
                    "PPT Layout",
                    ["3 images per page", "2 images per page", "1 image per page"],
                    index=0
                )
            
            # Generate PPT
            if st.button("🚀 Generate PPT Report", type="primary"):
                with st.spinner("Generating PPT..."):
                    ppt_creator = PPTCreator()
                    ppt_buffer = ppt_creator.create_ppt_from_images(
                        st.session_state.extracted_defects,
                        ppt_name
                    )
                
                if ppt_buffer:
                    st.success("✅ PPT generated successfully!")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download PPT File",
                        data=ppt_buffer,
                        file_name=ppt_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        help="Click to download the generated PPT report"
                    )
                    
                    # Display PPT statistics
                    st.subheader("📊 PPT Report Statistics")
                    
                    # Count defect types
                    defects_by_reason = defaultdict(list)
                    for defect in st.session_state.extracted_defects:
                        defects_by_reason[defect['reason']].append(defect)
                    
                    stats_data = []
                    for reason, defects in sorted(defects_by_reason.items()):
                        stats_data.append({
                            "Defect Reason": reason,
                            "Image Count": len(defects),
                            "Involved PDF Files": len(set(d['pdf_name'] for d in defects))
                        })
                    
                    st.dataframe(stats_data, use_container_width=True)
                else:
                    st.error("❌ PPT generation failed")

# Sidebar information
with st.sidebar:
    st.header("ℹ️ Instructions")
    st.markdown("""
    ### Steps:
    1. **Upload HSE claim report PDF files**:
       - Click "Browse files" or drag and drop PDF files
       - Support multiple simultaneous uploads
    
    2. **Extract defect images**:
       - System automatically identifies defect images in PDFs
       - Automatically extracts defect reasons and codes
       - Generates image previews and statistics
    
    3. **Generate PPT report**:
       - Switch to PPT Generation tab
       - Set PPT filename and layout
       - Click generate button to create PPT
    
    4. **Download results**:
       - Download extracted images (ZIP format)
       - Download generated PPT report
    """)
    
    st.header("📈 System Information")
    st.markdown("""
    - **Version**: 1.0.0
    - **Update Date**: 2024-01-20
    - **Supported Formats**: PDF files
    - **Output Formats**: JPEG images + PPT report
    """)

if __name__ == "__main__":
    main()
