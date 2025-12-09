import streamlit as st
import streamlit.elements.image as st_image
from PIL import Image, ImageChops, ImageDraw
import io
import re
import zipfile
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from streamlit_drawable_canvas import st_canvas

# ==========================================
# 🔥 紧急修复补丁 (Monkey Patch) 🔥
# 修复 Streamlit 新版本导致 st_canvas 报错的问题
# ==========================================
if not hasattr(st_image, 'image_to_url'):
    def local_image_to_url(image, width, clamp, channels, output_format, image_id):
        """将 PIL 图片转为 Base64 DataURL，模拟旧版 Streamlit 行为"""
        buffered = io.BytesIO()
        # 强制转为 RGB 防止 RGBA 在 JPEG 下报错
        if output_format.upper() == "JPEG" and image.mode == "RGBA":
            image = image.convert("RGB")
        image.save(buffered, format=output_format)
        img_str = base64.b64encode(buffered.getvalue()).decode()
        return (f"data:image/{output_format.lower()};base64,{img_str}",)
    
    # 强行把这个函数塞回 Streamlit 里
    st_image.image_to_url = local_image_to_url
# ==========================================

# --- 页面配置 ---
st.set_page_config(page_title="PDF 图表手动提取工具 (修复版)", layout="wide", page_icon="✂️")

# --- 核心函数 ---
def sanitize_filename(text):
    text = re.sub(r'\s+', ' ', text).strip()
    return re.sub(r'[\\/*?:"<>|]', "_", text)[:50]

def trim_white_borders(pil_image):
    bg = Image.new(pil_image.mode, pil_image.size, pil_image.getpixel((0,0)))
    diff = ImageChops.difference(pil_image, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return pil_image.crop(bbox)
    return pil_image

def process_selection(page, rect_pdf, dpi_scale=8.33):
    # 1. 提取文字
    text_dict = page.get_text("dict", clip=rect_pdf)
    extracted_text_parts = []
    text_blocks_rects = []
    
    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"].strip()
                if text:
                    extracted_text_parts.append(text)
                    text_blocks_rects.append(span["bbox"])
    
    full_caption = " ".join(extracted_text_parts)
    if not full_caption:
        full_caption = "未命名图表"
        
    # 2. 高清截图 (600 DPI)
    mat = fitz.Matrix(dpi_scale, dpi_scale)
    pix = page.get_pixmap(matrix=mat, clip=rect_pdf, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    
    # 3. 涂白文字
    draw = ImageDraw.Draw(img)
    offset_x = rect_pdf.x0
    offset_y = rect_pdf.y0
    
    for bbox in text_blocks_rects:
        x0 = (bbox[0] - offset_x) * dpi_scale
        y0 = (bbox[1] - offset_y) * dpi_scale
        x1 = (bbox[2] - offset_x) * dpi_scale
        y1 = (bbox[3] - offset_y) * dpi_scale
        draw.rectangle([x0-2, y0-2, x1+2, y1+2], fill="white")
        
    # 4. 自动修剪
    final_img = trim_white_borders(img)
    
    out_io = io.BytesIO()
    final_img.save(out_io, format="PNG")
    
    return out_io.getvalue(), full_caption, final_img.width, final_img.height

import fitz # PyMuPDF

# --- UI 逻辑 ---
if 'extracted_list' not in st.session_state:
    st.session_state.extracted_list = []

with st.sidebar:
    st.header("1. 上传文件")
    uploaded_file = st.file_uploader("PDF 文件", type="pdf")
    
    st.header("3. 导出设置")
    ppt_ratio = st.radio("PPT 比例", ["3:4 (竖版)", "16:9 (横版)"], index=0)
    
    st.divider()
    st.write(f"已提取: **{len(st.session_state.extracted_list)}** 张")
    if st.button("🗑️ 清空列表"):
        st.session_state.extracted_list = []
        st.rerun()

st.title("✂️ 框选提取工具 (已修复错误)")
st.caption("步骤：上传 PDF → 选择页码 → **框选包含图和文字的区域** → 点击提取。")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    
    col_sel, col_info = st.columns([1, 3])
    with col_sel:
        page_num = st.number_input("当前页码", min_value=1, max_value=len(doc), value=1)
    
    # 准备页面图像
    page = doc[page_num - 1]
    
    # 2倍缩放显示
    display_zoom = 2.0
    disp_pix = page.get_pixmap(matrix=fitz.Matrix(display_zoom, display_zoom))
    bg_img = Image.open(io.BytesIO(disp_pix.tobytes("png")))
    
    st.write("👇 **在下方画框 (包含图和文字)**")
    
    # 画布
    canvas_result = st_canvas(
        fill_color="rgba(255, 0, 0, 0.1)",
        stroke_width=2,
        stroke_color="#FF0000",
        background_image=bg_img, # 这里之前报错，现在补丁已修复
        update_streamlit=True,
        height=bg_img.height,
        width=bg_img.width,
        drawing_mode="rect",
        key=f"canvas_p{page_num}",
        display_toolbar=True,
    )
    
    if canvas_result.json_data is not None:
        objects = canvas_result.json_data["objects"]
        if objects:
            last_obj = objects[-1]
            if st.button("⚡ 提取选中区域", type="primary"):
                scale = 1 / display_zoom
                r_x = last_obj["left"] * scale
                r_y = last_obj["top"] * scale
                r_w = last_obj["width"] * scale
                r_h = last_obj["height"] * scale
                
                rect_pdf = fitz.Rect(r_x, r_y, r_x + r_w, r_y + r_h)
                
                try:
                    img_bytes, img_name, w, h = process_selection(page, rect_pdf)
                    
                    st.session_state.extracted_list.append({
                        "bytes": img_bytes,
                        "name": sanitize_filename(img_name),
                        "page": page_num,
                        "w": w, "h": h
                    })
                    st.success(f"提取成功: {img_name}")
                except Exception as e:
                    st.error(f"提取出错: {e}")

    # --- 导出 ---
    if st.session_state.extracted_list:
        st.divider()
        st.subheader("📥 导出")
        
        c1, c2 = st.columns(2)
        
        # PPT
        prs = Presentation()
        if ppt_ratio.startswith("3:4"):
            pr
