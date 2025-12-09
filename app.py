import streamlit as st
import fitz  # PyMuPDF
from PIL import Image, ImageChops
import io
import re
import zipfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from streamlit_drawable_canvas import st_canvas

# --- 配置 ---
st.set_page_config(page_title="PDF 图表手动提取工具", layout="wide", page_icon="🖱️")

# --- 辅助函数 ---
def sanitize_filename(text):
    text = re.sub(r'\s+', ' ', text).strip()
    return re.sub(r'[\\/*?:"<>|]', "_", text)[:50]

def trim_white_borders(pil_image):
    """像切吐司一样切掉四周白边"""
    bg = Image.new(pil_image.mode, pil_image.size, pil_image.getpixel((0,0)))
    diff = ImageChops.difference(pil_image, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return pil_image.crop(bbox)
    return pil_image

def get_image_above_caption(page, caption_rect, page_width):
    """
    根据用户框选的图注位置，向上寻找缝隙。
    """
    c_x0, c_y0, c_x1, c_y1 = caption_rect
    
    # 1. 确定分栏（简单的左右判断）
    mid = page_width / 2
    if c_x1 < mid + 20: # 左栏
        col_x0, col_x1 = 0, mid
    elif c_x0 > mid - 20: # 右栏
        col_x0, col_x1 = mid, page_width
    else: # 通栏
        col_x0, col_x1 = 0, page_width

    # 2. 向上找天花板 (最近的文字块)
    blocks = page.get_text("blocks")
    top_limit = 50 # 默认页眉
    
    for b in blocks:
        # b: x0, y0, x1, y1, text...
        # 必须在图注上方
        if b[3] < c_y0:
            # 必须在同栏
            if not (b[2] < col_x0 or b[0] > col_x1):
                if b[3] > top_limit:
                    top_limit = b[3]
    
    # 返回图注上方的区域
    return fitz.Rect(col_x0, top_limit, col_x1, c_y0)

# --- 状态管理 ---
if 'extracted_images' not in st.session_state:
    st.session_state.extracted_images = []

# --- UI ---
st.title("🖱️ PDF 图表手动提取器 (600 DPI)")
st.markdown("""
**操作说明：**
1. 在左侧选择页码。
2. 用鼠标在图片上**框选“图注文字”**（例如：图1 某某系统）。
3. 点击“提取”按钮，程序会自动抓取**图注上方的图片**并以图注命名。
""")

with st.sidebar:
    uploaded_file = st.file_uploader("上传 PDF", type="pdf")
    
    if uploaded_file:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        total_pages = len(doc)
        page_selector = st.number_input("选择页码", min_value=1, max_value=total_pages, value=1)
        
        st.divider()
        st.write(f"当前已提取: {len(st.session_state.extracted_images)} 张")
        
        # 清空按钮
        if st.button("清空所有提取结果"):
            st.session_state.extracted_images = []
            st.rerun()

# --- 主界面 ---
if uploaded_file:
    # 1. 渲染当前页为图片供用户操作
    page_idx = page_selector - 1
    page = doc[page_idx]
    
    # 提高显示清晰度方便框选 (2倍缩放)
    display_zoom = 2.0
    pix = page.get_pixmap(matrix=fitz.Matrix(display_zoom, display_zoom))
    img_height = pix.height
    img_width = pix.width
    
    # 将 PyMuPDF 图像转为 PIL 供 Canvas 使用
    bg_image = Image.open(io.BytesIO(pix.tobytes("png")))

    col1, col2 = st.columns([3, 1])
    
    with col1:
        # 2. 创建画布组件
        canvas_result = st_canvas(
            fill_color="rgba(255, 165, 0, 0.3)",  # 填充色
            stroke_width=2,
            stroke_color="#FF0000",
            background_image=bg_image,
            update_streamlit=True,
            height=img_height,
            width=img_width,
            drawing_mode="rect", # 矩形模式
            key=f"canvas_p{page_selector}",
            display_toolbar=True,
        )

    with col2:
        st.write("### 操作面板")
        
        if canvas_result.json_data is not None:
            objects = canvas_result.json_data["objects"]
            
            if len(objects) > 0:
                # 获取最后一个画的框
                obj = objects[-1]
                
                # 3. 坐标转换 (Canvas像素 -> PDF坐标)
                # Canvas 是 2倍缩放显示的，所以要除以 2
                scale = 1 / display_zoom 
                
                rect_x = obj["left"] * scale
                rect_y = obj["top"] * scale
                rect_w = obj["width"] * scale
                rect_h = obj["height"] * scale
                
                # PDF 坐标下的图注框
                caption_rect = fitz.Rect(rect_x, rect_y, rect_x + rect_w, rect_y + rect_h)
                
                # 4. 提取文字（文件名）
                text_in_box = page.get_textbox(caption_rect).strip()
                if not text_in_box:
                    text_in_box = f"Figure_Page_{page_selector}"
                
                st.info(f"识别图名: **{text_in_box}**")
                
                if st.button("✂️ 确认提取", type="primary"):
                    # 5. 自动计算上方图片区域
                    # 逻辑：以你画的框为底，向上一直切到上一段文字
                    target_rect = get_image_above_caption(page, caption_rect, page.rect.width)
                    
                    if target_rect.height > 10:
                        # 6. 600 DPI 渲染 (72 * 8.33 ≈ 600)
                        zoom_600 = 8.33
                        hd_pix = page.get_pixmap(matrix=fitz.Matrix(zoom_600, zoom_600), clip=target_rect, alpha=False)
                        hd_img = Image.open(io.BytesIO(hd_pix.tobytes("png")))
                        
                        # 7. 自动切白边
                        final_img = trim_white_borders(hd_img)
                        
                        # 保存
                        img_byte_arr = io.BytesIO()
                        final_img.save(img_byte_arr, format='PNG')
                        
                        st.session_state.extracted_images.append({
                            "bytes": img_byte_arr.getvalue(),
                            "name": sanitize_filename(text_in_box),
                            "page": page_selector,
                            "w": final_img.width,
                            "h": final_img.height
                        })
                        st.success("已添加！")
                    else:
                        st.error("上方未检测到足够空间，请检查框选位置。")
            else:
                st.info("请在左侧图片上框选图注...")

    # --- 底部导出区域 ---
    st.divider()
    if st.session_state.extracted_images:
        st.subheader("📤 导出结果")
        
        # 预览
        with st.expander("查看已提取列表"):
            for item in st.session_state.extracted_images:
                st.write(f"P{item['page']} - {item['name']}")
        
        c1, c2 = st.columns(2)
        
        # PPT 生成
        ppt_type = st.radio("PPT 比例", ["3:4 (竖版)", "16:9 (横版)"])
        prs = Presentation()
        if ppt_type.startswith("3:4"):
            prs.slide_width = Inches(7.5); prs.slide_height = Inches(10)
        else:
            prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
            
        for item in st.session_state.extracted_images:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            
            # 布局
            pw, ph = prs.slide_width, prs.slide_height
            margin = Inches(0.5)
            
            # 图片
            img_stream = io.BytesIO(item['bytes'])
            # 简单自适应，底部留空给字
            avail_h = ph - Inches(2.0)
            avail_w = pw - margin*2
            
            ratio = item['w'] / item['h']
            w, h = avail_w, avail_w / ratio
            if h > avail_h:
                h = avail_h
                w = h * ratio
                
            left = (pw - w) / 2
            top = Inches(0.5)
            slide.shapes.add_picture(img_stream, left, top, width=w, height=h)
            
            # 图名
            tb = slide.shapes.add_textbox(margin, top + h + Inches(0.1), pw - margin*2, Inches(1.5))
            p = tb.text_frame.add_paragraph()
            p.text = item['name']
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(14)
            p.font.name = "Microsoft YaHei"
            
        ppt_out = io.BytesIO()
        prs.save(ppt_out)
        ppt_out.seek(0)
        
        c1.download_button("📥 下载 PPTX", ppt_out, "manual_extract.pptx")
        
        # ZIP 生成
        zip_out = io.BytesIO()
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, item in enumerate(st.session_state.extracted_images):
                zf.writestr(f"{i+1}_{item['name']}.png", item['bytes'])
        zip_out.seek(0)
        
        c2.download_button("📦 下载高清图包", zip_out, "manual_images.zip")
