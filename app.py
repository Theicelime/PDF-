import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import zipfile
from PIL import Image, ImageChops

# --- 页面设置 ---
st.set_page_config(page_title="论文图表暴力提取器", page_icon="⛏️", layout="wide")

def sanitize_filename(text):
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text[:50]

def trim(im):
    """
    自动裁剪图片四周的白边（基于像素差异）。
    如果图片是全白的，返回 None。
    """
    bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return im.crop(bbox)
    return None

def is_caption(text):
    # 匹配 "图 6", "图6", "Fig.6", "Figure 6"
    return re.match(r'^\s*(图|Fig(ure)?\.?)\s*\d+', text, re.IGNORECASE) is not None

def get_gap_crop(page, caption_block, all_blocks, page_width):
    """
    【核心逻辑：缝隙切片法】
    不找图，只找图注和上一段正文之间的缝隙。
    """
    c_x0, c_y0, c_x1, c_y1 = caption_block[:4]
    
    # 1. 强行判定栏位（以页面中线为界）
    mid_x = page_width / 2
    # 如果图注在左边
    if c_x1 < mid_x + 10: 
        col_x0, col_x1 = 0, mid_x
    # 如果图注在右边
    elif c_x0 > mid_x - 10:
        col_x0, col_x1 = mid_x, page_width
    # 否则是通栏
    else:
        col_x0, col_x1 = 0, page_width

    # 2. 寻找天花板（正上方最近的文字）
    # 默认天花板是页眉处 (70)
    top_limit = 70 
    
    # 遍历所有文本块，找到在这个栏位里，且在图注上方的块
    for b in all_blocks:
        b_x0, b_y0, b_x1, b_y1 = b[:4]
        
        # 排除自己
        if abs(b_y0 - c_y0) < 5: continue
        
        # 必须在图注上方
        if b_y1 < c_y0:
            # 必须在同一栏（水平有重叠）
            if not (b_x1 < col_x0 or b_x0 > col_x1):
                # 更新最高点：取最大的 y1（最靠下的那个文本块的底部）
                if b_y1 > top_limit:
                    top_limit = b_y1
    
    # 3. 生成切片区域
    # 宽度：直接占满整个分栏（靠后期去白边来修正）
    # 高度：从上一段文字的底部，到图注的顶部
    return fitz.Rect(col_x0, top_limit, col_x1, c_y0)

# --- 主界面 ---
st.title("⛏️ 论文图表提取 (缝隙切片版)")
st.markdown("原理：定位图注 -> 找到上一段文字 -> **暴力切取中间所有内容** -> 自动修剪白边。")

with st.sidebar:
    ppt_ratio = st.radio("PPT 尺寸", ["3:4 (竖版 A4)", "16:9 (宽屏)"])
    dpi = st.number_input("清晰度 (DPI倍率)", value=4.0, min_value=2.0, max_value=6.0)

uploaded_file = st.file_uploader("上传 PDF", type="pdf")

if uploaded_file and st.button("开始提取"):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    
    # PPT 初始化
    prs = Presentation()
    if ppt_ratio.startswith("3:4"):
        prs.slide_width = Inches(8.27); prs.slide_height = Inches(11.69)
    else:
        prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
        
    results = []
    status = st.empty()
    bar = st.progress(0)
    
    for i, page in enumerate(doc):
        status.text(f"正在切片: 第 {i+1} 页...")
        bar.progress((i+1)/len(doc))
        
        # 获取所有文本块
        blocks = page.get_text("blocks")
        
        for b in blocks:
            text = b[4].replace('\n', ' ').strip()
            
            # 1. 发现图注
            if is_caption(text):
                # 2. 计算缝隙区域
                crop_rect = get_gap_crop(page, b, blocks, page.rect.width)
                
                # 如果缝隙太小（小于10px），说明没有图，跳过
                if crop_rect.height < 10:
                    continue
                
                # 3. 高清渲染这个区域
                pix = page.get_pixmap(matrix=fitz.Matrix(dpi, dpi), clip=crop_rect, alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                
                # 4. 关键步骤：自动裁剪白边
                # 因为我们要了整个分栏的宽，左右肯定有很多白边，这里切掉
                try:
                    img_trimmed = trim(img)
                except:
                    img_trimmed = img
                    
                if img_trimmed is None or img_trimmed.height < 20:
                    continue
                
                # 转换数据
                img_byte_arr = io.BytesIO()
                img_trimmed.save(img_byte_arr, format='PNG')
                final_bytes = img_byte_arr.getvalue()
                
                results.append({
                    "bytes": final_bytes,
                    "name": sanitize_filename(text),
                    "caption": text,
                    "page": i+1
                })
                
                # --- 写入 PPT ---
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                ppt_w, ppt_h = prs.slide_width, prs.slide_height
                
                # 布局
                margin = Inches(0.5)
                max_w = ppt_w - 2 * margin
                max_h = ppt_h - Inches(2.0)
                
                w, h = img_trimmed.size
                ratio = w / h
                
                target_w = max_w
                target_h = target_w / ratio
                if target_h > max_h:
                    target_h = max_h
                    target_w = target_h * ratio
                    
                left = (ppt_w - target_w) / 2
                top = Inches(0.5)
                
                slide.shapes.add_picture(io.BytesIO(final_bytes), left, top, width=target_w, height=target_h)
                
                # 文本框
                tb = slide.shapes.add_textbox(margin, top + target_h + Inches(0.2), max_w, Inches(1.5))
                p = tb.text_frame.add_paragraph()
                p.text = text
                p.alignment = PP_ALIGN.CENTER
                p.font.size = Pt(14)
                p.font.bold = True
                p.font.name = "Microsoft YaHei"

    status.success(f"完成！共提取 {len(results)} 张图。")
    
    if results:
        col1, col2 = st.columns(2)
        
        ppt_out = io.BytesIO()
        prs.save(ppt_out); ppt_out.seek(0)
        col1.download_button("📥 下载 PPT", ppt_out, "extracted.pptx")
        
        zip_out = io.BytesIO()
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in results:
                zf.writestr(f"P{item['page']}_{item['name']}.png", item['bytes'])
        zip_out.seek(0)
        col2.download_button("📦 下载图片包 (ZIP)", zip_out, "images.zip")
        
        st.divider()
        st.write("### 结果预览")
        for res in results:
            st.image(res["bytes"], caption=res["caption"])
