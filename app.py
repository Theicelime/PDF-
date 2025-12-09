import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import zipfile
from PIL import Image, ImageChops

# --- 基础设置 ---
st.set_page_config(page_title="PDF 图表暴力提取 (最终修正版)", layout="wide", page_icon="🔨")

def sanitize_filename(text):
    return re.sub(r'[\\/*?:"<>|\s]', "_", text)[:50]

def trim_white_borders(pil_image):
    """
    自动切除图片四周的白边。
    """
    bg = Image.new(pil_image.mode, pil_image.size, pil_image.getpixel((0,0)))
    diff = ImageChops.difference(pil_image, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return pil_image.crop(bbox)
    return pil_image # 全白或切不了，返回原图

def is_caption(text):
    # 匹配中文和英文图注
    return re.match(r'^\s*(图|Fig(ure)?\.?)\s*\d+', text, re.IGNORECASE) is not None

def get_column_range(x_mid, page_width):
    """根据图注的中心位置，返回它所在的栏位左右边界"""
    mid_page = page_width / 2
    if x_mid < mid_page: # 左栏
        return 0, mid_page
    else: # 右栏
        return mid_page, page_width

def extract_figures_strictly(doc, dpi_scale=4.0):
    extracted_data = []
    
    for page_idx, page in enumerate(doc):
        # 1. 获取所有文本块，关键：sort=True 保证按人类阅读顺序（从上到下，从左到右）
        blocks = page.get_text("blocks", sort=True)
        page_w = page.rect.width
        
        # 找出本页所有图注
        captions = []
        for i, b in enumerate(blocks):
            text = b[4].strip().replace('\n', ' ')
            if is_caption(text):
                captions.append((i, b, text)) # 保存索引，方便找上一个块
        
        for i, (block_idx, cap_block, cap_text) in enumerate(captions):
            # cap_block: (x0, y0, x1, y1, text, block_no, block_type)
            c_x0, c_y0, c_x1, c_y1 = cap_block[:4]
            cap_center_x = (c_x0 + c_x1) / 2
            
            # --- A. 确定左右边界 (分栏) ---
            # 如果图注宽度超过页面的 60%，认为是通栏图，否则按左右分栏处理
            if (c_x1 - c_x0) > page_w * 0.6:
                col_x0, col_x1 = 0, page_w # 通栏
            else:
                col_x0, col_x1 = get_column_range(cap_center_x, page_w)
            
            # --- B. 确定上边界 (天花板) ---
            # 默认天花板是页眉 (假设 50pt)
            top_limit = 50.0 
            
            # 倒序遍历在当前图注之前的文本块，寻找最近的一个在同一栏的文字
            # blocks 已经是排好序的，所以我们从当前图注的 index 往前找
            for prev_idx in range(block_idx - 1, -1, -1):
                p_b = blocks[prev_idx]
                p_x0, p_y0, p_x1, p_y1 = p_b[:4]
                
                # 检查是否在同一栏 (水平方向有重叠)
                # 逻辑：文本块中心点是否在栏位范围内
                p_center_x = (p_x0 + p_x1) / 2
                if col_x0 <= p_center_x <= col_x1:
                    # 找到了正上方的文字！这就是天花板
                    top_limit = p_y1 # 文字的底部作为图片的顶部
                    break # 找到最近的一个就停止，不要再往上找了
            
            # --- C. 截图 ---
            # 定义截图区域：[栏左, 天花板, 栏右, 图注顶]
            # 加上一点 padding 防止切坏
            clip_rect = fitz.Rect(col_x0, top_limit, col_x1, c_y0)
            
            # 有效性检查：如果高度是负的或者太小，说明出错了
            if clip_rect.height < 10:
                continue
                
            # 高清渲染
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale), clip=clip_rect, alpha=False)
            
            # --- D. 去白边 (关键) ---
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                img_trimmed = trim_white_borders(img)
            except:
                img_trimmed = img
            
            # 如果切完没东西了，跳过
            if img_trimmed.width < 10 or img_trimmed.height < 10:
                continue
            
            # 转回 bytes
            out_buffer = io.BytesIO()
            img_trimmed.save(out_buffer, format="PNG")
            
            extracted_data.append({
                "image_bytes": out_buffer.getvalue(),
                "name": sanitize_filename(cap_text),
                "caption": cap_text,
                "page": page_idx + 1,
                "width": img_trimmed.width, # 像素宽
                "height": img_trimmed.height # 像素高
            })
            
    return extracted_data

# --- 主界面 ---
st.title("🔨 论文图表提取工具 (强力阻断模式)")
st.markdown("如果不准，那是我的错。此模式使用物理阻断法：**图注**与**上一段文字**之间的所有像素，一律切下来。")

with st.sidebar:
    st.header("生成设置")
    # 按照你的要求，3:4 竖版
    ppt_ver = st.radio("PPT 版式", ["3:4 (竖版 A4)", "16:9 (宽屏)"])
    dpi_val = 4.0 # 默认高清

uploaded_file = st.file_uploader("请上传PDF文件", type="pdf")

if uploaded_file and st.button("开始处理", type="primary"):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    
    # 1. 执行提取
    with st.spinner("正在逐页扫描缝隙..."):
        results = extract_figures_strictly(doc, dpi_scale=dpi_val)
    
    if not results:
        st.error("未提取到图片。请确认PDF是文字版（可选中文字），而非扫描版。")
    else:
        st.success(f"成功提取 {len(results)} 张图表！")
        
        # 2. 生成 PPT
        prs = Presentation()
        # 设置版式
        if ppt_ver.startswith("3:4"):
            prs.slide_width = Inches(7.5) # A4 宽
            prs.slide_height = Inches(10) # A4 高
        else:
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)
            
        for item in results:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # 空白页
            
            # PPT 尺寸
            pw = prs.slide_width
            ph = prs.slide_height
            margin = Inches(0.5)
            
            # 布局计算：图片区域预留 80% 高度，底部留给图注
            max_img_h = ph - Inches(2.0)
            max_img_w = pw - margin * 2
            
            # 原始比例
            ratio = item["width"] / item["height"]
            
            # 目标尺寸
            final_w = max_img_w
            final_h = final_w / ratio
            
            if final_h > max_img_h:
                final_h = max_img_h
                final_w = final_h * ratio
            
            # 居中放置
            left = (pw - final_w) / 2
            top = Inches(0.5)
            
            # 插入图片
            slide.shapes.add_picture(io.BytesIO(item["image_bytes"]), left, top, width=final_w, height=final_h)
            
            # 插入图注
            tb = slide.shapes.add_textbox(margin, top + final_h + Inches(0.2), pw - margin*2, Inches(1.5))
            tf = tb.text_frame
            p = tf.add_paragraph()
            p.text = item["caption"]
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(14)
            p.font.name = "Microsoft YaHei"
        
        # 3. 下载按钮
        col1, col2 = st.columns(2)
        
        # PPT
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        col1.download_button("📥 下载 PPT", ppt_io, "figures_export.pptx")
        
        # ZIP
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in results:
                fname = f"P{item['page']}_{item['name']}.png"
                zf.writestr(fname, item['image_bytes'])
        zip_io.seek(0)
        col2.download_button("📦 下载图片包 (ZIP)", zip_io, "figures_images.zip")
        
        st.divider()
        st.write("### 提取结果核对")
        for res in results:
            st.image(res["image_bytes"], caption=res["caption"])
