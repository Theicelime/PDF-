import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import zipfile

# --- 页面配置 ---
st.set_page_config(page_title="论文图表智能重构工具", page_icon="🧩", layout="wide")

# --- 核心工具函数 ---

def sanitize_filename(text):
    """清洗文件名"""
    text = re.sub(r'\s+', ' ', text)  # 合并空格
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text.strip()[:60]

def is_caption(text):
    """精准识别图注，支持中文和英文"""
    # 匹配: "图1", "图 1", "Fig.1", "Figure 1", "Fig 1"
    # 忽略大小写
    pattern = r'^\s*(图|Fig(ure)?\.?)\s*\d+'
    return re.match(pattern, text, re.IGNORECASE) is not None

def get_smart_bbox(page, caption_rect, text_blocks):
    """
    【重构核心】不再盲目截图，而是基于对象（Object-Based）计算包围盒。
    
    逻辑：
    1. 找到图注 (Bottom Limit)。
    2. 找到图注正上方最近的一段文字 (Top Limit)。
    3. 获取该区域内所有的 图片(Images) 和 绘图(Drawings)。
    4. 计算这些对象的并集矩形 (Union Rect)。
    """
    
    # 1. 确定搜索区域的 左右边界 (处理双栏)
    page_w = page.rect.width
    mid_x = page_w / 2
    
    # 判断图注在左栏、右栏还是跨栏
    if caption_rect.x1 < mid_x + 20: # 左栏
        search_x0, search_x1 = 0, mid_x + 20
    elif caption_rect.x0 > mid_x - 20: # 右栏
        search_x0, search_x1 = mid_x - 20, page_w
    else: # 通栏
        search_x0, search_x1 = 0, page_w
        
    # 2. 确定搜索区域的 上下边界
    # 下界：图注的顶部
    y_bottom = caption_rect.y0 
    
    # 上界：寻找正上方最近的一个文本块
    # 默认上界为页眉位置 (假设50)
    y_top = 50 
    
    # 在所有文本块中，找到位于图注上方、且在同栏内的最近文本
    closest_gap = float('inf')
    
    for b in text_blocks:
        b_rect = fitz.Rect(b[:4])
        # 排除当前的图注本身
        if abs(b_rect.y0 - caption_rect.y0) < 5:
            continue
            
        # 必须在图注上方
        if b_rect.y1 < y_bottom:
            # 必须在同栏 (水平方向有交集)
            if not (b_rect.x1 < search_x0 or b_rect.x0 > search_x1):
                gap = y_bottom - b_rect.y1
                if gap < closest_gap:
                    closest_gap = gap
                    y_top = b_rect.y1 # 更新上界为这段文字的底部

    # 稍微放宽一点上界，防止紧贴
    y_top = max(50, y_top + 2) 

    # 定义“感兴趣区域” (ROI)
    roi_rect = fitz.Rect(search_x0, y_top, search_x1, y_bottom)

    # 3. 获取所有视觉对象 (Images & Drawings)
    # PyMuPDF 的 get_drawings 获取所有矢量路径
    drawings = page.get_drawings()
    # get_images 获取位图
    images = page.get_images(full=True)
    
    # 容器：存放所有属于该图的对象矩形
    target_rects = []
    
    # 筛选矢量绘图
    for draw in drawings:
        r = draw["rect"]
        # 如果这个矢量图在 ROI 内部，或者与 ROI 高度重叠
        intersect = r & roi_rect # 计算交集
        if intersect.get_area() > 0:
            # 排除巨大的背景色块 (比如整个页面的背景)
            if r.width > page_w * 0.9 and r.height > page.rect.height * 0.9:
                continue
            target_rects.append(r)
            
    # 筛选图片对象
    for img in images:
        try:
            img_rect = page.get_image_bbox(img)
            intersect = img_rect & roi_rect
            if intersect.get_area() > 0:
                target_rects.append(img_rect)
        except:
            pass

    # 4. 计算最终包围盒 (Merge)
    if not target_rects:
        # 如果真的啥也没抓到（极少见），回退到几何切割
        return roi_rect
    
    # 计算所有矩形的并集
    final_rect = target_rects[0]
    for r in target_rects:
        final_rect |= r # Union操作
        
    # 5. 最终修正
    # 确保宽度不会因为某个错误的线条变得无限宽，限制在栏宽内
    final_rect.x0 = max(search_x0, final_rect.x0)
    final_rect.x1 = min(search_x1, final_rect.x1)
    
    # 确保底部不覆盖图注
    final_rect.y1 = min(final_rect.y1, caption_rect.y0)
    
    # 增加一点点内边距，为了美观
    return final_rect


# --- UI 主程序 ---
st.title("🧩 论文图表智能重构 (Refactored)")
st.caption("使用 对象聚类算法 (Object Clustering) 替代传统的截图扫描，精准提取矢量图与混合图表。")

with st.sidebar:
    st.header("设置")
    ppt_orientation = st.radio("PPT版式", ["3:4 (竖版/阅读模式)", "16:9 (横版/演示模式)"])
    dpi_scale = 4.0 # 强制高清

uploaded_file = st.file_uploader("上传 PDF 论文", type="pdf")

if uploaded_file and st.button("🚀 开始重构与提取", type="primary"):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    
    # 初始化 PPT
    prs = Presentation()
    if ppt_orientation.startswith("3:4"):
        prs.slide_width = Inches(8.27)  # A4 宽度
        prs.slide_height = Inches(11.69) # A4 高度
    else:
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
    results = []
    status = st.empty()
    bar = st.progress(0)
    
    for page_idx, page in enumerate(doc):
        status.text(f"正在解析结构: 第 {page_idx + 1} 页...")
        bar.progress((page_idx + 1) / len(doc))
        
        # 1. 获取所有文本块 (用于定位上下文)
        text_blocks = page.get_text("blocks")
        # 2. 找出所有图注
        captions = []
        for b in text_blocks:
            text = b[4].replace('\n', ' ').strip()
            if is_caption(text):
                captions.append({
                    "rect": fitz.Rect(b[:4]),
                    "text": text
                })
        
        if not captions:
            continue
            
        # 3. 针对每个图注，智能计算其对应的图形区域
        for cap in captions:
            # 核心重构方法调用
            figure_rect = get_smart_bbox(page, cap["rect"], text_blocks)
            
            # 过滤无效区域
            if figure_rect.width < 10 or figure_rect.height < 10:
                continue
                
            # 4. 高清渲染该区域
            # fitz.Matrix(4, 4) = 300 DPI
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale), clip=figure_rect, alpha=False)
            img_bytes = pix.tobytes("png")
            
            # 结果存入列表
            results.append({
                "bytes": img_bytes,
                "name": sanitize_filename(cap["text"]),
                "page": page_idx + 1,
                "w": pix.width,
                "h": pix.height
            })
            
            # --- 5. 写入 PPT ---
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # 空白页
            
            ppt_w = prs.slide_width
            ppt_h = prs.slide_height
            margin = Inches(0.5)
            
            # 布局计算
            avail_w = ppt_w - 2 * margin
            avail_h = ppt_h - 2 * Inches(1.0) # 留出文本空间
            
            img_ratio = pix.width / pix.height
            
            # 适应逻辑 (Contain)
            final_w = avail_w
            final_h = final_w / img_ratio
            
            if final_h > avail_h:
                final_h = avail_h
                final_w = final_h * img_ratio
            
            left = (ppt_w - final_w) / 2
            top = (avail_h - final_h) / 2 + Inches(0.2)
            
            # 插入图片
            slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=final_w, height=final_h)
            
            # 插入图名 (底部居中)
            txbox = slide.shapes.add_textbox(margin, top + final_h + Inches(0.1), avail_w, Inches(1))
            tf = txbox.text_frame
            p = tf.add_paragraph()
            p.text = cap["text"]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(14)
            p.font.bold = True
            
    status.text("✅ 重构完成！")
    
    if results:
        col1, col2 = st.columns(2)
        
        # PPT 下载
        out_ppt = io.BytesIO()
        prs.save(out_ppt)
        out_ppt.seek(0)
        col1.download_button("📥 下载 PPT", out_ppt, "smart_layout.pptx")
        
        # ZIP 下载
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for item in results:
                fname = f"P{item['page']}_{item['name']}.png"
                zf.writestr(fname, item['bytes'])
        zip_buf.seek(0)
        col2.download_button("📦 下载图片包 (ZIP)", zip_buf, "smart_images.zip")
        
        # 预览
        st.divider()
        st.subheader(f"成功提取 {len(results)} 个图表结构")
        for res in results:
            st.image(res["bytes"], caption=f"Page {res['page']}: {res['name']}")
    else:
        st.warning("未检测到图表。请确认PDF包含 '图 X' 或 'Figure X' 格式的图注。")
