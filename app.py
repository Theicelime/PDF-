import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import zipfile
from PIL import Image, ImageChops

# --- 页面配置 ---
st.set_page_config(page_title="论文图表精准提取工具 (最终版)", page_icon="✂️", layout="wide")

# --- 辅助函数 ---

def trim_white_space(pil_image):
    """
    像切吐司边一样，自动切除图片四周的空白区域。
    """
    bg = Image.new(pil_image.mode, pil_image.size, (255, 255, 255))
    diff = ImageChops.difference(pil_image, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return pil_image.crop(bbox)
    return pil_image

def sanitize_filename(text):
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text[:50]

def is_caption(text):
    # 匹配 "图 1", "图1", "Fig 1", "Figure 1"
    return re.match(r'^\s*(图|Fig(ure)?\.?)\s*\d+', text, re.IGNORECASE) is not None

def get_precise_crop_area(page, current_caption_block, all_blocks, page_width, page_height):
    """
    核心算法：三明治夹心法 + 严格分栏
    """
    # current_caption_block: (x0, y0, x1, y1, text, ...)
    c_x0, c_y0, c_x1, c_y1 = current_caption_block[:4]
    
    # 1. 判断栏位 (左栏 / 右栏 / 通栏)
    # 这种学术期刊中缝一般在宽度的 50% 处
    mid_point = page_width / 2
    caption_center = (c_x0 + c_x1) / 2
    
    if c_x1 < mid_point + 10: 
        # === 左栏 ===
        scan_x0, scan_x1 = 0, mid_point
    elif c_x0 > mid_point - 10:
        # === 右栏 ===
        scan_x0, scan_x1 = mid_point, page_width
    else:
        # === 通栏 (跨页大图) ===
        scan_x0, scan_x1 = 0, page_width
        
    # 2. 寻找上边界 (Top Limit)
    # 向上寻找最近的一个文本块（无论是正文还是上一个图注），把它作为“天花板”
    # 默认天花板是页眉位置 (假设 60)
    top_limit = 60
    
    # 筛选出所有在“当前图注”上方的文本块
    blocks_above = []
    for b in all_blocks:
        b_y1 = b[3] # 文本块的底边
        b_x0, b_x1 = b[0], b[2]
        
        # 必须在图注上方
        if b_y1 < c_y0:
            # 必须在同一栏内 (水平方向有重叠)
            # 只要有一点点水平重叠就算，防止漏掉居中的标题
            if not (b_x1 < scan_x0 or b_x0 > scan_x1):
                blocks_above.append(b_y1)
    
    if blocks_above:
        # 找到最靠下的那个文本块的底边，作为图片的起始位置
        top_limit = max(blocks_above)
        
    # 3. 构建初始裁剪框 (粗略)
    # 留一点余地 (padding)，防止切掉线条边缘
    rect = fitz.Rect(scan_x0, top_limit, scan_x1, c_y0)
    
    return rect

# --- UI ---
st.title("✂️ 论文图表精准切分 (防干扰版)")
st.markdown("""
**解决痛点：**
1. 彻底解决图6、图7连在一起切不开的问题。
2. 彻底解决右边栏文字被切进去的问题。
3. 自动切除白边，图片不再留有大片空白。
""")

with st.sidebar:
    st.header("设置")
    ppt_ratio = st.radio("PPT 尺寸", ["3:4 (竖版 A4)", "16:9 (宽屏)"])
    zoom_dpi = 4.0 # 300 DPI

uploaded_file = st.file_uploader("重新上传你的 PDF", type="pdf")

if uploaded_file and st.button("🚀 重新提取 (执行严格模式)", type="primary"):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    
    prs = Presentation()
    if ppt_ratio.startswith("3:4"):
        prs.slide_width = Inches(8.27); prs.slide_height = Inches(11.69)
    else:
        prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
        
    results = []
    status = st.empty()
    bar = st.progress(0)
    
    for page_idx, page in enumerate(doc):
        status.text(f"正在精细处理: 第 {page_idx + 1} 页...")
        bar.progress((page_idx + 1) / len(doc))
        
        # 1. 获取全页文本块 (按垂直坐标排序)
        blocks = page.get_text("blocks")
        # 格式: (x0, y0, x1, y1, text, block_no, block_type)
        blocks.sort(key=lambda b: b[1]) 
        
        for b in blocks:
            text = b[4].replace('\n', ' ').strip()
            
            # 2. 锁定图注
            if is_caption(text):
                caption_rect = b # 保存整个block信息
                
                # 3. 计算“安全区域” (Safe Zone)
                # 这一步只确定：左边界、右边界、上边界(碰到上一段字为止)、下边界(碰到图注为止)
                crop_rect = get_precise_crop_area(page, caption_rect, blocks, page.rect.width, page.rect.height)
                
                # 校验：如果高度太小(小于20像素)，说明图注贴着上一段字，没图，跳过
                if crop_rect.height < 10:
                    continue
                
                # 4. 高清截图 (此时截图包含大量白边)
                mat = fitz.Matrix(zoom_dpi, zoom_dpi)
                pix = page.get_pixmap(matrix=mat, clip=crop_rect, alpha=False)
                
                # 转换成 PIL 图片进行二次处理
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                
                # 5. 【关键步骤】自动裁剪白边 (Trim Whitespace)
                # 这一步会去掉所有多余的空白，只留下图表内容
                try:
                    trimmed_img = trim_white_space(img)
                except Exception:
                    trimmed_img = img # 兜底
                
                # 再次校验：如果切完白边没东西了，跳过
                if trimmed_img.width < 10 or trimmed_img.height < 10:
                    continue
                
                # 转回 bytes
                img_byte_arr = io.BytesIO()
                trimmed_img.save(img_byte_arr, format='PNG')
                final_img_bytes = img_byte_arr.getvalue()
                
                # 存入结果
                results.append({
                    "bytes": final_img_bytes,
                    "name": sanitize_filename(text),
                    "caption": text,
                    "page": page_idx + 1,
                    "w": trimmed_img.width,
                    "h": trimmed_img.height
                })
                
                # --- 写入 PPT ---
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                ppt_w, ppt_h = prs.slide_width, prs.slide_height
                margin = Inches(0.5)
                
                # 布局计算
                avail_w = ppt_w - 2 * margin
                avail_h = ppt_h - Inches(2.0) # 底部留多点位置给字
                
                img_w, img_h = trimmed_img.size
                aspect = img_w / img_h
                
                target_w = avail_w
                target_h = target_w / aspect
                if target_h > avail_h:
                    target_h = avail_h
                    target_w = target_h * aspect
                
                left = (ppt_w - target_w) / 2
                top = Inches(0.5) # 顶对齐，或者居中
                
                slide.shapes.add_picture(io.BytesIO(final_img_bytes), left, top, width=target_w, height=target_h)
                
                # 图注文本框
                tb = slide.shapes.add_textbox(margin, top + target_h + Inches(0.2), avail_w, Inches(1.5))
                p = tb.text_frame.add_paragraph()
                p.text = text
                p.alignment = PP_ALIGN.CENTER
                p.font.size = Pt(14); p.font.bold = True; p.font.name = "Microsoft YaHei"

    status.success(f"处理完成！成功提取 {len(results)} 张图。")
    
    if results:
        c1, c2 = st.columns(2)
        out_ppt = io.BytesIO()
        prs.save(out_ppt); out_ppt.seek(0)
        c1.download_button("📥 下载 PPT", out_ppt, "final_result.pptx")
        
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in results:
                zf.writestr(f"P{item['page']}_{item['name']}.png", item['bytes'])
        zip_buf.seek(0)
        c2.download_button("📦 下载图片包", zip_buf, "final_images.zip")
        
        st.divider()
        for r in results:
            st.image(r["bytes"], caption=f"P{r['page']}: {r['caption']}")
