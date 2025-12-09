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

# --- 页面基础设置 ---
st.set_page_config(page_title="PDF 图表手动提取工具 (去图名版)", layout="wide", page_icon="✂️")

# --- 核心处理函数 ---

def sanitize_filename(text):
    """清理文件名"""
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

def process_selection(page, rect_pdf, dpi_scale=8.33):
    """
    输入：PDF页面，用户画的矩形(PDF坐标系)
    输出：处理后的图片(bytes), 提取到的图名(str)
    """
    # 1. 提取矩形内的文字（作为图名）
    # 使用 "dict" 模式可以获取文字的精确坐标，方便后续涂白
    text_dict = page.get_text("dict", clip=rect_pdf)
    
    extracted_text_parts = []
    text_blocks_rects = [] # 记录文字的区域，用于涂白
    
    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"].strip()
                if text:
                    extracted_text_parts.append(text)
                    # 记录这段文字的包围盒 (x0, y0, x1, y1)
                    text_blocks_rects.append(span["bbox"])
    
    # 拼接图名
    full_caption = " ".join(extracted_text_parts)
    if not full_caption:
        full_caption = "未命名图表"
        
    # 2. 高清截图 (包含图和字)
    # 600 DPI ≈ 8.33 倍 zoom (72 * 8.33 = 600)
    mat = fitz.Matrix(dpi_scale, dpi_scale)
    pix = page.get_pixmap(matrix=mat, clip=rect_pdf, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    
    # 3. 【关键】涂白文字区域 (去除图名)
    draw = ImageDraw.Draw(img)
    
    # PDF坐标 -> 图片像素坐标 的转换系数
    # 因为我们只截取了 rect_pdf 这一块，所以原点要移动
    offset_x = rect_pdf.x0
    offset_y = rect_pdf.y0
    
    for bbox in text_blocks_rects:
        # bbox 是全局PDF坐标
        # 我们需要转换成“相对于截图左上角”的坐标，并乘缩放倍率
        x0 = (bbox[0] - offset_x) * dpi_scale
        y0 = (bbox[1] - offset_y) * dpi_scale
        x1 = (bbox[2] - offset_x) * dpi_scale
        y1 = (bbox[3] - offset_y) * dpi_scale
        
        # 稍微画大一点点，确保覆盖干净
        margin = 2
        draw.rectangle([x0-margin, y0-margin, x1+margin, y1+margin], fill="white")
        
    # 4. 自动修剪白边 (Trim)
    # 此时图名已经被涂白了，trim 会自动把这些留白切掉
    final_img = trim_white_borders(img)
    
    # 转 bytes
    out_io = io.BytesIO()
    final_img.save(out_io, format="PNG")
    
    return out_io.getvalue(), full_caption, final_img.width, final_img.height

# --- 状态管理 ---
if 'extracted_list' not in st.session_state:
    st.session_state.extracted_list = []

# --- UI 侧边栏 ---
with st.sidebar:
    st.header("1. 上传文件")
    uploaded_file = st.file_uploader("PDF 文件", type="pdf")
    
    st.header("3. 导出设置")
    ppt_ratio = st.radio("PPT 比例", ["3:4 (竖版)", "16:9 (横版)"], index=0)
    
    st.divider()
    # 结果列表管理
    st.write(f"已提取: **{len(st.session_state.extracted_list)}** 张")
    if st.button("🗑️ 清空列表"):
        st.session_state.extracted_list = []
        st.rerun()

# --- 主区域 ---
st.title("✂️ 框选提取工具")
st.caption("步骤：上传 PDF -> 选择页码 -> **框选包含图和文字的区域** -> 点击提取。程序会自动提取字作为名字，并在图片中删除字。")

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    
    # 页码选择器
    col_sel, col_btn = st.columns([1, 3])
    with col_sel:
        page_num = st.number_input("当前页码", min_value=1, max_value=len(doc), value=1)
    
    # 准备页面图像供 Canvas 显示
    page = doc[page_num - 1]
    
    # 为了操作流畅，显示时用 2倍 缩放 (144 DPI)
    display_zoom = 2.0
    disp_pix = page.get_pixmap(matrix=fitz.Matrix(display_zoom, display_zoom))
    bg_img = Image.open(io.BytesIO(disp_pix.tobytes("png")))
    
    # 画布区域
    st.write("### 👇 在下方画框 (包含图和图注)")
    
    # 创建画布
    canvas_result = st_canvas(
        fill_color="rgba(255, 0, 0, 0.1)", # 红色半透明
        stroke_width=2,
        stroke_color="#FF0000",
        background_image=bg_img,
        update_streamlit=True,
        height=bg_img.height,
        width=bg_img.width,
        drawing_mode="rect",
        key=f"canvas_p{page_num}", # 换页重置画布
        display_toolbar=True
    )
    
    # 处理逻辑
    if canvas_result.json_data is not None:
        objects = canvas_result.json_data["objects"]
        if objects:
            # 取最后一个画的框
            last_obj = objects[-1]
            
            if st.button("⚡ 提取选中区域", type="primary"):
                # 1. 坐标换算 (Canvas -> PDF)
                scale = 1 / display_zoom
                r_x = last_obj["left"] * scale
                r_y = last_obj["top"] * scale
                r_w = last_obj["width"] * scale
                r_h = last_obj["height"] * scale
                
                rect_pdf = fitz.Rect(r_x, r_y, r_x + r_w, r_y + r_h)
                
                # 2. 调用核心处理
                img_bytes, img_name, w, h = process_selection(page, rect_pdf)
                
                # 3. 存入 session
                st.session_state.extracted_list.append({
                    "bytes": img_bytes,
                    "name": sanitize_filename(img_name),
                    "page": page_num,
                    "w": w, "h": h
                })
                st.success(f"已提取: {img_name}")
                
    
    # --- 导出区域 ---
    if st.session_state.extracted_list:
        st.divider()
        st.subheader("📥 导出与预览")
        
        # 预览
        with st.expander("点击查看已提取的图片"):
            cols = st.columns(3)
            for i, item in enumerate(st.session_state.extracted_list):
                with cols[i % 3]:
                    st.image(item["bytes"], caption=f"图名: {item['name']}")
        
        c1, c2 = st.columns(2)
        
        # 生成 PPT
        prs = Presentation()
        # 设置 PPT 尺寸
        if ppt_ratio.startswith("3:4"):
            prs.slide_width = Inches(7.5)
            prs.slide_height = Inches(10)
        else:
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)
            
        for item in st.session_state.extracted_list:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # 空白页
            
            # 布局参数
            pw = prs.slide_width
            ph = prs.slide_height
            margin = Inches(0.5)
            
            # 图片区域 (留底部给文字)
            max_h = ph - Inches(1.5)
            max_w = pw - margin * 2
            
            # 计算缩放
            ratio = item["w"] / item["h"]
            target_w = max_w
            target_h = target_w / ratio
            
            if target_h > max_h:
                target_h = max_h
                target_w = target_h * ratio
            
            # 居中
            left = (pw - target_w) / 2
            top = Inches(0.5)
            
            # 插入图片
            slide.shapes.add_picture(io.BytesIO(item["bytes"]), left, top, width=target_w, height=target_h)
            
            # 插入图名 (文本框)
            tb = slide.shapes.add_textbox(margin, top + target_h + Inches(0.1), pw - margin*2, Inches(1))
            p = tb.text_frame.add_paragraph()
            p.text = item["name"]
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(14)
            p.font.name = "Microsoft YaHei"
            
        ppt_out = io.BytesIO()
        prs.save(ppt_out)
        ppt_out.seek(0)
        c1.download_button("📥 下载 PPTX", ppt_out, "extracted_slides.pptx")
        
        # 生成 ZIP
        zip_out = io.BytesIO()
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, item in enumerate(st.session_state.extracted_list):
                # 文件名: 页码_序号_图名.png
                fname = f"P{item['page']}_{i+1}_{item['name']}.png"
                zf.writestr(fname, item["bytes"])
        zip_out.seek(0)
        c2.download_button("📦 下载图片包 (ZIP)", zip_out, "extracted_images.zip")

else:
    st.info("请在左侧上传 PDF 文件开始。")
