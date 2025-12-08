import streamlit as st
import sys
import subprocess
import os

# 尝试导入PyMuPDF，如果失败则自动安装
def install_pymupdf():
    """安装PyMuPDF包"""
    try:
        st.info("正在安装PyMuPDF...")
        # 使用pip安装
        subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF==1.23.8"])
        st.success("PyMuPDF安装成功！")
        return True
    except Exception as e:
        st.error(f"安装失败: {e}")
        return False

# 尝试导入fitz
try:
    import fitz
    st.success("✅ PyMuPDF导入成功！")
    FITZ_AVAILABLE = True
except ImportError:
    st.warning("❌ PyMuPDF未安装")
    if st.button("点击安装PyMuPDF"):
        if install_pymupdf():
            # 重新加载模块
            import importlib
            import fitz
            FITZ_AVAILABLE = True
            st.rerun()  # 重新运行应用
        else:
            st.error("安装失败，请检查日志")
            FITZ_AVAILABLE = False
    else:
        FITZ_AVAILABLE = False

# 设置页面
st.set_page_config(page_title="PDF测试", layout="wide")
st.title("📄 PDF缺陷提取器")

# 显示环境信息
st.write(f"Python版本: {sys.version}")
st.write(f"当前目录: {os.getcwd()}")
st.write(f"PyMuPDF可用: {FITZ_AVAILABLE}")

# 如果PyMuPDF可用，显示上传功能
if FITZ_AVAILABLE:
    uploaded_file = st.file_uploader("上传PDF文件", type=["pdf"])
    
    if uploaded_file:
        st.success(f"已上传文件: {uploaded_file.name}")
        
        # 临时保存文件
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        
        try:
            # 打开PDF
            doc = fitz.open(tmp_path)
            st.success(f"✅ PDF打开成功！共 {len(doc)} 页")
            
            # 显示一些信息
            col1, col2 = st.columns(2)
            with col1:
                st.metric("总页数", len(doc))
            
            # 提取第一页的文本（测试）
            if st.button("提取第一页文本"):
                page = doc[0]
                text = page.get_text()
                st.text_area("第一页文本", text[:500] + "..." if len(text) > 500 else text, height=200)
            
            doc.close()
            
        except Exception as e:
            st.error(f"处理PDF时出错: {e}")
        
        finally:
            # 清理临时文件
            try:
                os.unlink(tmp_path)
            except:
                pass
else:
    st.error("请先安装PyMuPDF依赖")
