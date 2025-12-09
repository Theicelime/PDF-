import streamlit as st
import fitz  # PyMuPDF
from PIL import Image, ImageChops, ImageDraw
import io
import re
import zipfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from streamlit_drawable_canvas import st_canvas
import base64

# --- 页面基础设置 ---
st.set_page_config(page_title="PDF 图表手动提取工具", layout="wide", page_icon="✂️")

# --- 核心处理函数 ---
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

# --- 状态管理 ---
if 'extracted_list' not in st.session_state:
    st.session_state.extracted_list = []

# --- UI ---
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

st.title("✂️ 框选提取工具")
st.caption("步骤：上传 PDF → 选择页码 → **框选包含图和文字的区域** → 点击提取。")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    
    col_sel, col_info = st.columns([1, 3])
    with col_sel:
        page_num = st.number_input("当前页码", min_value=1, max_value=len(doc), value=1)
    
    # 准备页面图像
    page = doc[page_num - 1]
    # 使用 2.0 倍缩放显示，保证画框时能看清字
    display_zoom = 2.0
    disp_pix = page.get_pixmap(matrix=fitz.Matrix(display_zoom, display_zoom))
    bg_img = Image.open(io.BytesIO(disp_pix.tobytes("png")))
    
    st.write("👇 **在下方画框 (务必把图和下面的图注文字都框进去)**")
    
    # 画布
    # 注意：这里如果streamlit版本不对，会报错。请确保 requirements.txt 使用 streamlit==1.38.0
    canvas_result = st_canvas(
        fill_color="rgba(255, 0, 0, 0.1)",
        stroke_width=2,
        stroke_color="#FF0000",
        background_image=bg_img, # 关键点：需要 Streamlit <= 1.38.0
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
                # 坐标换算 Canvas -> PDF
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
                    st.success(f"成功提取: {img_name}")
                except Exception as e:
                    st.error(f"提取失败，请重试: {e}")

    # --- 导出 ---
    if st.session_state.extracted_list:
        st.divider()
        st.subheader("📥 导出")
        
        c1, c2 = st.columns(2)
        
        # PPT
        prs = Presentation()
        if ppt_ratio.startswith("3:4"):
            prs.slide_width = Inches(7.5); prs.slide_height = Inches(10)
        else:
            prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
            
        for item in st.session_state.extracted_list:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            pw, ph = prs.slide_width, prs.slide_height
            margin = Inches(0.5)
            
            # 图片布局
            max_h = ph - Inches(1.5)
            max_w = pw - margin * 2
            ratio = item["w"] / item["h"]
            target_w = max_w
            target_h = target_w / ratio
            if target_h > max_h:
                target_h = max_h
                target_w = target_h * ratio
                
            left = (pw - target_w) / 2
            top = Inches(0.5)
            
            slide.shapes.add_picture(io.BytesIO(item["bytes"]), left, top, width=target_w, height=target_h)
            
            tb = slide.shapes.add_textbox(margin, top + target_h + Inches(0.1), pw - margin*2, Inches(1))
            p = tb.text_frame.add_paragraph()
            p.text = item["name"]
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(14)
            p.font.name = "Microsoft YaHei"
            
        ppt_out = io.BytesIO()
        prs.save(ppt_out); ppt_out.seek(0)
        c1.download_button("📥 下载 PPTX", ppt_out, "extracted_slides.pptx")
        
        # ZIP
        zip_out = io.BytesIO()
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, item in enumerate(st.session_state.extracted_list):
                zf.writestr(f"P{item['page']}_{i+1}_{item['name']}.png", item["bytes"])
        zip_out.seek(0)
        c2.download_button("📦 下载图片包", zip_out, "extracted_images.zip")

else:
    st.info("请上传 PDF 开始")
