import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
import io
from PIL import Image

# --- 页面配置 ---
st.set_page_config(page_title="PDF 转 PPT 提取工具", page_icon="📊")

st.title("📄 PDF 图表提取与布局工具")
st.markdown("上传 PDF 文件，自动提取其中的图片并按 **16:9** 尺寸居中排版生成 PPT。")

# --- 侧边栏设置 ---
st.sidebar.header("⚙️ 参数设置")
min_px = st.sidebar.slider("忽略小于此像素的图片", 50, 500, 100, help="用于过滤掉图标、Logo等小图片")
layout_mode = st.sidebar.radio("布局模式", ["居中适应 (Contain)", "拉伸铺满 (Stretch)"], index=0)

# --- 文件上传 ---
uploaded_file = st.file_uploader("请拖入或选择 PDF 文件", type="pdf")

if uploaded_file is not None:
    # 显示文件信息
    st.info(f"文件名: {uploaded_file.name} | 大小: {uploaded_file.size / 1024:.2f} KB")
    
    if st.button("🚀 开始转换", type="primary"):
        try:
            # 1. 读取 PDF
            # 注意：Streamlit 的 uploaded_file 是 BytesIO，PyMuPDF 需要 bytes
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            
            # 2. 初始化 PPT
            prs = Presentation()
            prs.slide_width = Inches(13.333) # 16:9 宽度
            prs.slide_height = Inches(7.5)   # 16:9 高度
            
            img_count = 0
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            total_pages = len(doc)
            
            # 3. 遍历处理
            for page_index, page in enumerate(doc):
                status_text.text(f"正在处理第 {page_index + 1}/{total_pages} 页...")
                progress_bar.progress((page_index + 1) / total_pages)
                
                image_list = page.get_images(full=True)
                
                for img in image_list:
                    xref = img[0]
                    base = doc.extract_image(xref)
                    image_bytes = base["image"]
                    
                    try:
                        # 图片预处理与过滤
                        image_stream = io.BytesIO(image_bytes)
                        pil_img = Image.open(image_stream)
                        w, h = pil_img.size
                        
                        if w < min_px or h < min_px:
                            continue
                            
                        # 新建幻灯片 (空白版式)
                        slide = prs.slides.add_slide(prs.slide_layouts[6])
                        
                        # PPT 尺寸 (Emu 单位)
                        ppt_w = prs.slide_width
                        ppt_h = prs.slide_height
                        
                        # 计算位置与尺寸
                        if layout_mode == "居中适应 (Contain)":
                            # 保持比例缩放
                            aspect_ratio = w / h
                            target_w = ppt_w
                            target_h = target_w / aspect_ratio
                            
                            if target_h > ppt_h:
                                target_h = ppt_h
                                target_w = target_h * aspect_ratio
                            
                            left = (ppt_w - target_w) / 2
                            top = (ppt_h - target_h) / 2
                            slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width=target_w, height=target_h)
                            
                        else: 
                            # 拉伸 (不推荐，但作为选项)
                            slide.shapes.add_picture(io.BytesIO(image_bytes), 0, 0, width=ppt_w, height=ppt_h)
                            
                        img_count += 1
                        
                    except Exception as e:
                        print(f"Skipped image due to error: {e}")
            
            # 4. 导出结果
            output_ppt = io.BytesIO()
            prs.save(output_ppt)
            output_ppt.seek(0)
            
            status_text.text("✅ 处理完成！")
            st.success(f"成功提取并布局了 {img_count} 张图片。")
            
            # 下载按钮
            st.download_button(
                label="📥 下载 PPTX 文件",
                data=output_ppt,
                file_name=f"converted_{uploaded_file.name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
        except Exception as e:
            st.error(f"发生错误: {str(e)}")
