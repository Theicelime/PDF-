import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import zipfile
from PIL import Image

# --- 页面基础配置 ---
st.set_page_config(page_title="论文图表高清提取工具", page_icon="📑", layout="wide")

# --- 核心逻辑函数 ---

def sanitize_filename(text):
    """清理文件名，去除非法字符，保留图名关键信息"""
    # 去除换行符
    text = text.replace('\n', ' ').replace('\r', '')
    # 只保留中文、字母、数字、部分符号
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    # 限制长度防止文件名过长
    return text.strip()[:80]

def is_caption(text):
    """
    判断文本块是否是图注。
    针对中文期刊优化：匹配 '图 1'、'图1'、'Fig. 1'、'Figure 1'
    """
    # 移除首尾空白
    text = text.strip()
    # 正则：以 "图" 或 "Fig" 开头，后跟数字
    # 允许 "图" 和数字之间有空格
    pattern = r'^(图|Fig(ure)?\.?)\s*\d+'
    return re.match(pattern, text, re.IGNORECASE) is not None

def get_smart_clip_rect(page, caption_rect, page_width, page_height):
    """
    智能计算截图区域 (核心算法)
    针对双栏排版优化。
    """
    x0, y0, x1, y1 = caption_rect
    caption_center_x = (x0 + x1) / 2
    
    # --- 1. 判断版式 (左栏、右栏、通栏) ---
    # 假设页面分为三部分：左(0-40%)，中(40-60%)，右(60-100%)
    # 实际上双栏的中轴线大约在 page_width / 2
    
    layout_type = "UNKNOWN"
    
    # 判定阈值
    left_boundary = page_width * 0.45
    right_boundary = page_width * 0.55
    
    if x1 < left_boundary:
        layout_type = "LEFT_COLUMN"
        search_x0, search_x1 = 0, page_width / 2
    elif x0 > right_boundary:
        layout_type = "RIGHT_COLUMN"
        search_x0, search_x1 = page_width / 2, page_width
    else:
        # 如果图注横跨了中轴线，或者位于中间，通常是通栏大图
        layout_type = "FULL_WIDTH"
        search_x0, search_x1 = 0, page_width

    # --- 2. 向上寻找视觉元素 (Images & Drawings) ---
    # 获取页面上所有的绘图指令(矢量线条)和图片
    drawings = page.get_drawings()
    images = page.get_images(full=True)
    
    # 收集所有位于图注上方、且在当前栏宽度范围内的视觉元素包围盒
    candidates = []
    
    # 设定搜索的顶部极限 (防止截到上一页的内容或者页眉)
    # 假设图表不会超过大半页，且至少在页眉(50pt)之下
    min_y_limit = 50 
    
    # 检查矢量绘图 (线条、背景色块等)
    for draw in drawings:
        r = draw["rect"] # fitz.Rect
        # 逻辑：
        # 1. 元素底部必须在图注上方 (r.y1 <= y0 + 10) (+10是容错)
        # 2. 元素顶部必须在页眉下方
        # 3. 元素水平方向必须在当前栏范围内 (有一定交集)
        if r.y1 <= y0 + 15 and r.y0 > min_y_limit:
            # 检查水平重叠
            if not (r.x1 < search_x0 or r.x0 > search_x1):
                candidates.append(r)
                
    # 检查嵌入图片
    for img in images:
        try:
            img_rect = page.get_image_bbox(img)
            if img_rect.y1 <= y0 + 15 and img_rect.y0 > min_y_limit:
                 if not (img_rect.x1 < search_x0 or img_rect.x0 > search_x1):
                    candidates.append(img_rect)
        except:
            pass

    # --- 3. 计算最终裁剪框 ---
    if not candidates:
        # 如果没找到任何矢量或图片对象（可能是扫描件或者纯文本图），回退到几何估算
        # 默认截取图注上方 1/3 页高度的区域
        fallback_height = page_height / 3
        final_top = max(min_y_limit, y0 - fallback_height)
        
        # 宽度收缩一下，避免贴边
        margin = 30
        final_rect = fitz.Rect(search_x0 + margin, final_top, search_x1 - margin, y0)
        return final_rect
    
    # 合并所有候选框
    final_rect = candidates[0]
    for r in candidates:
        final_rect |= r # 计算并集
        
    # --- 4. 边界微调 ---
    # 底部：紧贴图注上方
    final_rect.y1 = y0
    
    # 左右：如果是通栏，尽量居中；如果是分栏，确保不越界
    # 可以在检测到的物体边缘再加一点点留白(padding)
    padding = 5
    final_rect.x0 = max(0, final_rect.x0 - padding)
    final_rect.x1 = min(page_width, final_rect.x1 + padding)
    final_rect.y0 = max(min_y_limit, final_rect.y0 - padding)
    
    # 宽度校验：如果检测到的区域太窄（比如只是一个标点），可能出错了，强制扩充到图注宽度
    if final_rect.width < caption_rect.width:
        center = (final_rect.x0 + final_rect.x1) / 2
        half_w = caption_rect.width / 2
        final_rect.x0 = min(final_rect.x0, center - half_w)
        final_rect.x1 = max(final_rect.x1, center + half_w)

    return final_rect

# --- 主程序 UI ---
st.title("📑 论文智能图表提取 & PPT生成器 (Pro版)")
st.markdown("专为双栏排版中文期刊设计。自动识别“图 X”，智能裁剪，生成高清PPT。")

# 侧边栏设置
with st.sidebar:
    st.header("⚙️ 导出设置")
    ppt_ratio = st.radio("PPT 画板尺寸", ["16:9 (宽屏)", "3:4 (竖版/A4类似)", "4:3 (传统)"])
    st.info("💡 说明：\n会自动使用 **300 DPI** 超高清渲染，确保文字清晰可见。")

uploaded_file = st.file_uploader("📂 上传 PDF 文件", type="pdf")

if uploaded_file:
    # 按钮触发
    if st.button("🚀 开始高清提取"):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        
        # 1. 初始化 PPT
        prs = Presentation()
        
        # 设置尺寸
        if ppt_ratio == "16:9 (宽屏)":
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
        elif ppt_ratio == "3:4 (竖版/A4类似)":
            # 7.5英寸宽 x 10英寸高
            prs.slide_width = Inches(7.5)
            prs.slide_height = Inches(10)
        else:
            # 4:3
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(7.5)

        extracted_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_pages = len(doc)
        
        for page_idx, page in enumerate(doc):
            status_text.text(f"正在扫描第 {page_idx + 1}/{total_pages} 页...")
            progress_bar.progress((page_idx + 1) / total_pages)
            
            # 获取文本块
            blocks = page.get_text("blocks")
            # 排序：从上到下，从左到右
            blocks.sort(key=lambda b: (b[1], b[0]))
            
            for block in blocks:
                # block: (x0, y0, x1, y1, text, ...)
                text = block[4]
                
                if is_caption(text):
                    # 找到图注
                    caption_rect = fitz.Rect(block[:4])
                    clean_caption = text.strip().replace("\n", " ")
                    
                    # 智能计算图片区域
                    clip_rect = get_smart_clip_rect(page, caption_rect, page.rect.width, page.rect.height)
                    
                    # 过滤无效小区域
                    if clip_rect.width < 50 or clip_rect.height < 50:
                        continue
                        
                    # --- 高清截图 (Snapshot) ---
                    # matrix=4 表示 4倍分辨率 (约300 DPI)，保证极高清晰度
                    zoom = 4 
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat, clip=clip_rect, alpha=False)
                    img_bytes = pix.tobytes("png")
                    
                    # 文件名处理
                    file_name_clean = sanitize_filename(clean_caption)
                    if not file_name_clean:
                        file_name_clean = f"Page_{page_idx+1}_Figure"
                        
                    extracted_results.append({
                        "bytes": img_bytes,
                        "name": file_name_clean,
                        "page": page_idx + 1
                    })
                    
                    # --- 写入 PPT ---
                    # 使用空白版式
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    
                    ppt_w = prs.slide_width
                    ppt_h = prs.slide_height
                    
                    # 1. 放置图片
                    # 计算图片缩放比例 (Contain)
                    margin = Inches(0.5) # 边距
                    max_w = ppt_w - 2 * margin
                    max_h = ppt_h - 2 * Inches(1.0) # 底部留多一点给文字
                    
                    img_w_px = pix.width
                    img_h_px = pix.height
                    aspect = img_w_px / img_h_px
                    
                    target_w = max_w
                    target_h = target_w / aspect
                    
                    if target_h > max_h:
                        target_h = max_h
                        target_w = target_h * aspect
                        
                    left = (ppt_w - target_w) / 2
                    top = (ppt_h - target_h) / 2 - Inches(0.3) # 稍微往上提一点
                    
                    image_stream = io.BytesIO(img_bytes)
                    slide.shapes.add_picture(image_stream, left, top, width=target_w, height=target_h)
                    
                    # 2. 放置图注 (标题)
                    textbox_height = Inches(1.0)
                    txBox = slide.shapes.add_textbox(margin, top + target_h + Inches(0.1), max_w, textbox_height)
                    tf = txBox.text_frame
                    tf.word_wrap = True # 自动换行
                    p = tf.add_paragraph()
                    p.text = clean_caption
                    p.font.size = Pt(16) # 字号
                    p.font.bold = True
                    p.font.name = 'Microsoft YaHei' # 尝试设置微软雅黑
                    p.alignment = PP_ALIGN.CENTER
        
        status_text.text("✅ 处理完成！")
        
        if not extracted_results:
            st.error("未找到以'图'或'Figure'开头的图注。请检查PDF是否包含可搜索文本。")
        else:
            st.success(f"成功提取 {len(extracted_results)} 张高清图表！")
            
            # --- 下载区域 ---
            c1, c2 = st.columns(2)
            
            # 1. 下载 PPT
            out_ppt = io.BytesIO()
            prs.save(out_ppt)
            out_ppt.seek(0)
            c1.download_button(
                label=f"📥 下载 PPT ({ppt_ratio})",
                data=out_ppt,
                file_name="paper_figures.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary"
            )
            
            # 2. 下载图片包
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, item in enumerate(extracted_results):
                    # 文件名格式: P1_图1_xxx.png
                    fname = f"P{item['page']}_{item['name']}.png"
                    zf.writestr(fname, item['bytes'])
            zip_buffer.seek(0)
            
            c2.download_button(
                label="📦 下载高清图片包 (ZIP)",
                data=zip_buffer,
                file_name="figures_hd.zip",
                mime="application/zip"
            )
            
            st.divider()
            st.subheader("🖼️ 提取结果预览")
            for item in extracted_results:
                st.image(item['bytes'], caption=f"P{item['page']} | {item['name']}")
