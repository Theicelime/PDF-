import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import zipfile
from PIL import Image

# --- 配置 ---
st.set_page_config(page_title="论文图表智能提取器", page_icon="📑", layout="wide")

def sanitize_filename(text):
    """清理文件名，移除非法字符"""
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text.strip()[:50]  # 限制长度

def is_caption(text):
    """判断文本块是否像图注"""
    # 匹配常见的图注开头：Fig. 1, Figure 2, 图 3, Fig 4
    pattern = r'^(Fig(ure)?\.?|图)\s*\d+'
    return re.match(pattern, text, re.IGNORECASE) is not None

def get_image_area(page, caption_rect, page_width):
    """
    核心算法：根据图注位置，向上寻找图片区域。
    策略：
    1. 图注上方通常是图。
    2. 扫描图注上方的空间，直到遇到上一段文字（Text Block）或页面顶部。
    3. 为了避免截取到正文，我们检测上方最近的一个文本块的底部。
    """
    x0, y0, x1, y1 = caption_rect
    
    # 获取页面所有文本块
    blocks = page.get_text("blocks")
    
    # 找到当前图注在blocks中的索引（近似）
    current_block_idx = -1
    for i, b in enumerate(blocks):
        # b 的格式: (x0, y0, x1, y1, text, block_no, block_type)
        if abs(b[1] - y0) < 5 and abs(b[0] - x0) < 5: # 坐标匹配
            current_block_idx = i
            break
            
    # 默认顶部边界是页面顶部（或者页眉下方）
    top_boundary = 50 # 假设页眉高度
    
    # 尝试寻找图注“上方”最近的一个文本块作为边界
    # 简单的倒序遍历
    # 注意：PDF Block 顺序不一定代表物理位置，所以我们要按坐标找
    
    # 筛选出所有位于图注上方(y < y0)的文本块
    blocks_above = [b for b in blocks if b[3] < y0] # b[3]是bottom y
    
    if blocks_above:
        # 找到最靠下的那个文本块（离图注最近的上方文字）
        nearest_text_block = max(blocks_above, key=lambda b: b[3])
        top_boundary = nearest_text_block[3] + 5 # 留一点缝隙
    
    # 确定图片区域
    # 左边界和右边界：如果图注很宽，可能是通栏图；如果很窄，可能是双栏图
    # 这里做一个简单的启发式：取图注的宽度，稍微外扩，或者如果是学术论文，往往图是居中的
    
    # 策略A：激进模式，截取整行宽度（适合单栏或通栏图）
    # rect = fitz.Rect(50, top_boundary, page_width - 50, y0)
    
    # 策略B：适应性模式 (推荐)
    # 如果图注在左半边，可能是左栏；在右半边，是右栏。
    # 这里简化处理：以图注中心为轴，向两边扩充，或者直接扫描该区域内的绘图指令（Drawings）
    
    # 为了保证截取完整，我们使用 PyMuPDF 的 "drawings" 检测
    drawings = page.get_drawings()
    # 筛选出位于 top_boundary 和 y0 之间的绘图元素
    relevant_rects = []
    
    # 添加图片对象检测 (Image objects)
    images = page.get_images(full=True)
    for img in images:
        try:
            img_rect = page.get_image_bbox(img)
            if img_rect.y1 <= y0 + 10 and img_rect.y0 >= top_boundary - 50:
                 relevant_rects.append(img_rect)
        except:
            pass

    # 如果没有检测到明确对象，回退到几何切割
    if not relevant_rects:
        # 默认：宽度与图注对齐，或者扩展到版心
        # 判断是否跨栏：图注中心点
        center_x = (x0 + x1) / 2
        if page_width > 0:
            if 0.3 * page_width < center_x < 0.7 * page_width:
                 # 中间位置，假设是通栏大图
                 img_x0, img_x1 = 40, page_width - 40
            elif center_x < 0.5 * page_width:
                 # 左栏
                 img_x0, img_x1 = 40, page_width / 2
            else:
                 # 右栏
                 img_x0, img_x1 = page_width / 2, page_width - 40
            
            return fitz.Rect(img_x0, top_boundary, img_x1, y0)
    
    # 如果检测到了绘图元素，计算它们的并集包围盒
    final_rect = fitz.Rect(relevant_rects[0]) if relevant_rects else fitz.Rect(x0, top_boundary, x1, y0)
    for r in relevant_rects:
        final_rect |= r # 合并矩形
        
    # 稍微修正边界，包含图注宽度
    final_rect.x0 = min(final_rect.x0, x0)
    final_rect.x1 = max(final_rect.x1, x1)
    # 确保不越过文字边界
    final_rect.y0 = max(final_rect.y0, top_boundary)
    final_rect.y1 = y0 # 底部紧贴图注上方
    
    return final_rect


# --- UI ---
st.title("📊 论文图表提取与 PPT 生成器")
st.markdown("""
本工具专为学术论文设计：
1. **自动识别图注** (Figure X...)
2. **智能截取** 图注上方的图表区域（含矢量图、文字、组合图）
3. **高清导出** 并自动生成 PPT
""")

col1, col2 = st.columns(2)
with col1:
    ppt_ratio = st.selectbox("PPT 尺寸", ["16:9 (宽屏)", "4:3 (标准)"])
with col2:
    zoom_level = st.slider("截图清晰度 (DPI倍率)", 1.0, 4.0, 2.0, 0.5, help="2.0 相当于 144 DPI，3.0 相当于 216 DPI")

uploaded_file = st.file_uploader("上传 PDF 论文", type="pdf")

if uploaded_file:
    if st.button("🚀 开始提取分析"):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        
        # 准备 PPT
        prs = Presentation()
        if ppt_ratio == "16:9 (宽屏)":
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
        else:
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(7.5)
            
        extracted_data = [] # 存储结果: {'image': bytes, 'name': str, 'page': int}
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for page_num, page in enumerate(doc):
            status_text.text(f"正在分析第 {page_num + 1} 页...")
            progress_bar.progress((page_num + 1) / len(doc))
            
            # 1. 获取所有文本块
            blocks = page.get_text("blocks")
            blocks.sort(key=lambda b: b[1]) # 按垂直位置排序
            
            for b in blocks:
                text = b[4].strip().replace('\n', ' ')
                
                # 2. 判断是否是图注
                if is_caption(text):
                    # b: (x0, y0, x1, y1, text, block_no, block_type)
                    caption_rect = fitz.Rect(b[:4])
                    
                    # 3. 智能计算图片区域
                    # 简单的启发式：通常图在图注上方，高度不超过半页
                    # 我们尝试截取图注上方的一块区域
                    
                    # 确定裁剪框
                    clip_rect = get_image_area(page, caption_rect, page.rect.width)
                    
                    # 4. 有效性检查
                    if clip_rect.height < 20 or clip_rect.width < 20:
                        continue
                        
                    # 5. 高清渲染 (Snapshot)
                    # matrix 控制缩放，2 表示 2倍分辨率
                    mat = fitz.Matrix(zoom_level, zoom_level)
                    pix = page.get_pixmap(matrix=mat, clip=clip_rect, alpha=False)
                    img_data = pix.tobytes("png")
                    
                    # 6. 生成文件名
                    safe_name = sanitize_filename(text)
                    if not safe_name:
                        safe_name = f"Figure_Page_{page_num+1}"
                    
                    extracted_data.append({
                        "image_bytes": img_data,
                        "name": safe_name,
                        "caption": text,
                        "page": page_num + 1,
                        "width": pix.width,
                        "height": pix.height
                    })
                    
                    # --- 添加到 PPT ---
                    blank_slide_layout = prs.slide_layouts[6] 
                    slide = prs.slides.add_slide(blank_slide_layout)
                    
                    # 添加图片
                    img_stream = io.BytesIO(img_data)
                    
                    ppt_w = prs.slide_width
                    ppt_h = prs.slide_height
                    
                    # 图片布局计算 (Contain)
                    margin_top = Inches(0.5)
                    margin_bottom = Inches(1.5) # 底部留给图注
                    available_h = ppt_h - margin_top - margin_bottom
                    
                    # 原始尺寸
                    img_w_px = pix.width
                    img_h_px = pix.height
                    ratio = img_w_px / img_h_px
                    
                    # 目标尺寸
                    target_w = ppt_w
                    target_h = target_w / ratio
                    
                    if target_h > available_h:
                        target_h = available_h
                        target_w = target_h * ratio
                        
                    left = (ppt_w - target_w) / 2
                    top = (available_h - target_h) / 2 + margin_top
                    
                    slide.shapes.add_picture(img_stream, left, top, width=target_w, height=target_h)
                    
                    # 添加图注文本框
                    tx_box = slide.shapes.add_textbox(Inches(0.5), top + target_h + Inches(0.2), ppt_w - Inches(1), Inches(1))
                    tf = tx_box.text_frame
                    tf.word_wrap = True
                    p = tf.add_paragraph()
                    p.text = text
                    p.alignment = PP_ALIGN.CENTER
                    p.font.size = Pt(14)
                    p.font.bold = True

        status_text.text("✅ 处理完成！")
        
        if extracted_data:
            st.success(f"共提取到 {len(extracted_data)} 张图表。")
            
            # --- 下载区域 ---
            col_d1, col_d2 = st.columns(2)
            
            # 1. PPT 下载
            ppt_out = io.BytesIO()
            prs.save(ppt_out)
            ppt_out.seek(0)
            col_d1.download_button(
                label="📥 下载 PPTX",
                data=ppt_out,
                file_name="extracted_figures.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            
            # 2. 图片打包下载 (ZIP)
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, item in enumerate(extracted_data):
                    # 防止重名
                    file_name = f"{item['page']}_{idx}_{item['name']}.png"
                    zf.writestr(file_name, item['image_bytes'])
            
            zip_buffer.seek(0)
            col_d2.download_button(
                label="📦 下载高清图片包 (ZIP)",
                data=zip_buffer,
                file_name="figures_images.zip",
                mime="application/zip"
            )
            
            # --- 预览区域 ---
            st.divider()
            st.subheader("预览提取结果")
            for item in extracted_data:
                st.image(item['image_bytes'], caption=f"P{item['page']}: {item['caption']}")
                
        else:
            st.warning("未检测到明显的图注（Figure/Fig./图）。请确认PDF是可搜索文本的格式，而非扫描件。")
