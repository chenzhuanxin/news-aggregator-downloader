# -*- coding: utf-8 -*-
"""
新闻资讯下载器 —— 通用工具模块
============================================================
提供 HTTP 请求、图片/视频资源处理、公文格式 docx（含图文+视频链接）
生成、日期/路径工具，被 sites.py 中 15 个站点抓取函数复用。
"""
import datetime
import os
import re
import urllib.parse
import urllib3

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36')

SESSION = requests.Session()
SESSION.headers.update({'User-Agent': UA})
VERIFY = False          # 部分站点证书链不完整，统一关闭证书校验
TIMEOUT = 30

IMG_EXTS = ('.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp')
VIDEO_MARKERS = ('video', '.mp4', '.m3u8', 'player/index', 'video/player',
                 'getvideo', '.flv', '.webm', 'video_url', 'playurl')


# ============================================================
# HTTP 工具
# ============================================================
def http_get(url, params=None, headers=None, timeout=TIMEOUT, encoding=None):
    """GET 请求，返回 requests.Response（自动处理编码）。"""
    h = dict(headers or {})
    r = SESSION.get(url, params=params, headers=h, timeout=timeout, verify=VERIFY)
    r.raise_for_status()
    if encoding:
        r.encoding = encoding
    return r


def http_download(url, dest_path, headers=None, timeout=120):
    """下载二进制文件到 dest_path，返回字节内容。"""
    h = dict(headers or {})
    r = SESSION.get(url, headers=h, timeout=timeout, verify=VERIFY)
    r.raise_for_status()
    with open(dest_path, 'wb') as f:
        f.write(r.content)
    return r.content


def soup_of(resp_or_html):
    """把响应文本或 HTML 字符串解析为 BeautifulSoup 对象。"""
    if isinstance(resp_or_html, (str, bytes)):
        return BeautifulSoup(resp_or_html, 'lxml')
    return BeautifulSoup(resp_or_html.text, 'lxml')


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


# ============================================================
# 日期 / 路径工具
# ============================================================
def today_date():
    return datetime.date.today()


def yesterday_date():
    return datetime.date.today() - datetime.timedelta(days=1)


def date_folder_name(dt):
    """子文件夹名：新闻资讯 20261008"""
    return f'新闻资讯 {dt.strftime("%Y%m%d")}'


def site_filename(site_name, dt, ext):
    """单站点文件：站点名 20261008.docx"""
    return f'{site_name} {dt.strftime("%Y%m%d")}.{ext}'


# ============================================================
# 公文格式 docx 工具（标题黑体三号居中 / 正文仿宋三号缩进）
# ============================================================
def _set_run_font(run, east_asia, size_pt=16, bold=False,
                  ascii_font=None, color=(0, 0, 0)):
    """同时设置中西文字体、字号、加粗、颜色。"""
    if ascii_font is None:
        ascii_font = east_asia
    run.font.name = ascii_font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    try:
        run.font.color.rgb = RGBColor(*color)
    except Exception:
        pass
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), ascii_font)
    rfonts.set(qn('w:hAnsi'), ascii_font)
    rfonts.set(qn('w:eastAsia'), east_asia)
    rfonts.set(qn('w:cs'), east_asia)


def _set_fixed_line_spacing(p, spacing_pt=30):
    """行距固定 30 磅，段前段后 0 磅。"""
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(spacing_pt)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def _add_first_line_indent(p, size_pt=16):
    """首行缩进 2 字符（firstLineChars=200），并以 firstLine 兜底。"""
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), '200')
    ind.set(qn('w:firstLine'), str(int(size_pt * 2 * 20)))


def _docx_default_style(doc, font_name='仿宋', size_pt=16):
    """设置文档默认样式（防 Word 回退默认字体）。"""
    try:
        normal = doc.styles['Normal']
        normal.font.name = font_name
        normal.font.size = Pt(size_pt)
        normal.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    except Exception:
        pass


def _add_news_title(doc, text, size=16):
    """新闻标题：黑体 三号 加粗 居中。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_fixed_line_spacing(p)
    run = p.add_run(text)
    _set_run_font(run, '黑体', size, bold=True, ascii_font='SimHei')


def _add_news_body(doc, text, justify=True, indent=True, ascii_font=None):
    """正文段落：仿宋 三号，首行缩进 2 字符，可选两端对齐。"""
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_fixed_line_spacing(p)
    if indent:
        _add_first_line_indent(p, 16)
    run = p.add_run(text)
    _set_run_font(run, '仿宋', 16, bold=False, ascii_font=ascii_font)


def _add_meta_line(doc, text, size=12):
    """附注行：来源/发布时间等，仿宋小字居中或左对齐。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_fixed_line_spacing(p, 24)
    run = p.add_run(text)
    _set_run_font(run, '仿宋', size, bold=False, ascii_font='Times New Roman')


def _add_video_line(doc, text, size=14):
    """视频链接行：加【视频】前缀。"""
    p = doc.add_paragraph()
    _set_fixed_line_spacing(p, 26)
    _add_first_line_indent(p, 14)
    run = p.add_run(text)
    _set_run_font(run, '仿宋', size, bold=False, ascii_font='Times New Roman')


# ============================================================
# 图片下载与去重
# ============================================================
def _safe_ext(url):
    path = urllib.parse.urlparse(url).path
    ext = os.path.splitext(path)[1].lower()
    return ext if ext in IMG_EXTS else '.jpg'


def download_image(url, work_dir, index=0, headers=None):
    """下载单张图片到 work_dir，返回本地路径；失败返回 None。"""
    try:
        full = urllib.parse.urljoin('', url) if url.startswith(('http://', 'https://')) else url
        if not url.startswith(('http://', 'https://')):
            return None
        local = os.path.join(work_dir, f'img_{index}{_safe_ext(url)}')
        http_download(url, local, headers=headers)
        return local if os.path.getsize(local) > 0 else None
    except Exception:
        return None


def embed_images(doc, image_urls, work_dir, max_width_cm=14.5, headers=None):
    """把多张新闻图片按顺序嵌入 docx 中。"""
    for i, u in enumerate(image_urls):
        if not u:
            continue
        local = download_image(u, work_dir, index=i, headers=headers)
        if not local:
            continue
        try:
            from PIL import Image as PILImage
            with PILImage.open(local) as im:
                w, h = im.size
            max_h = 19.0
            ratio = min(max_width_cm / w, max_h / h) if w and h else max_width_cm / 800
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_fixed_line_spacing(p, 20)
            run = p.add_run()
            run.add_picture(local, width=Cm(min(max_width_cm, w * ratio)))
        except Exception:
            pass


# ============================================================
# 构建站点 docx（图文 + 视频链接）
# ============================================================
def build_site_docx(site_name, date_obj, articles, out_path, work_dir):
    """
    把某个站点某一天的文章列表生成一个 docx。
    articles: list of dict，字段：
        title      : 标题
        meta       : 附注（发布时间 / 作者 / 栏目 / 来源），可为空
        paragraphs : 正文段落列表
        images     : 图片 URL 列表
        video_links: 视频链接列表（[{'label':..,'url':..}] 或 [str]）
    布局：每篇文章之间分页，标题黑体居中，图片随文居中嵌入，
          正文仿宋缩进，视频链接以【视频】列出。
    """
    doc = Document()
    _docx_default_style(doc)
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.8)

    # 文档主标题
    main = doc.add_paragraph()
    main.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_fixed_line_spacing(main, 36)
    run = main.add_run(f'{site_name}（{date_obj:%Y年%m月%d日}）')
    _set_run_font(run, '黑体', 22, bold=True, ascii_font='SimHei')
    _add_meta_line(doc, f'共 {len(articles)} 篇　·　本文件由「新闻资讯下载器」生成', 11)

    for i, art in enumerate(articles):
        if i > 0:
            doc.add_page_break()
        title = art.get('title') or '（无标题）'
        _add_news_title(doc, f'{i + 1}. {title}')
        meta = art.get('meta')
        if meta:
            _add_meta_line(doc, meta)
        # 图片
        images = art.get('images') or []
        if images:
            embed_images(doc, images, work_dir)
        # 正文
        for para in art.get('paragraphs') or []:
            para = re.sub(r'\s+', ' ', str(para)).strip()
            if para:
                _add_news_body(doc, para)
        # 视频链接
        videos = art.get('video_links') or []
        if videos:
            blank = doc.add_paragraph()
            _set_fixed_line_spacing(blank, 20)
            for v in videos:
                if isinstance(v, dict):
                    label = v.get('label') or '相关视频'
                    url = v.get('url', '')
                else:
                    label, url = '相关视频', str(v)
                _add_video_line(doc, f'【视频】{label}：{url}')

    doc.save(out_path)
    return out_path


# ============================================================
# 从 HTML 片段提取正文段落、图片、视频（站点通用）
# ============================================================
def clean_paragraphs(node, skip_images=True):
    """从容器节点提取 <p> 段落纯文本（可跳过含图的段落）。"""
    paras = []
    if node is None:
        return paras
    for el in node.find_all(['p', 'section', 'div']):
        if el.name == 'div' and (el.find('p') or el.find('section')):
            continue
        if skip_images and el.find('img'):
            continue
        txt = ' '.join(el.stripped_strings)
        txt = re.sub(r'\s+', ' ', txt).strip()
        if len(txt) >= 2:
            paras.append(txt)
    return paras


def extract_images(node, base_url=None, attr_candidates=('src', 'data-src', 'data-original', 'data-lazy-src')):
    """从容器节点提取所有图片 URL。"""
    urls, seen = [], set()
    if node is None:
        return urls
    for img in node.find_all('img'):
        u = None
        for attr in attr_candidates:
            val = img.get(attr)
            if val:
                u = val.strip()
                break
        if not u:
            continue
        if base_url:
            u = urllib.parse.urljoin(base_url, u)
        if u.startswith(('http://', 'https://')) and u not in seen:
            seen.add(u)
            urls.append(u)
    return urls


def extract_videos(node, base_url=None):
    """从容器节点提取视频链接（video/source/iframe）。"""
    links, seen = [], set()
    if node is None:
        return links
    for v in node.find_all(['video', 'source', 'audio', 'iframe']):
        u = v.get('src') or v.get('data-src') or ''
        if not u:
            continue
        if base_url:
            u = urllib.parse.urljoin(base_url, u)
        if u.startswith(('http://', 'https://')) and u not in seen:
            seen.add(u)
            links.append(u)
    # 兜底：a[href] 指向视频文件
    for a in node.find_all('a', href=True):
        href = a['href'].strip()
        if base_url:
            href = urllib.parse.urljoin(base_url, href)
        low = href.lower()
        if low.startswith(('http://', 'https://')) and \
                any(mk in low for mk in ('.mp4', '.m3u8', '.flv', '.webm')):
            if href not in seen:
                seen.add(href)
                links.append(href)
    return links
