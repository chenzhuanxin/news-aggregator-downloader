# -*- coding: utf-8 -*-
"""
新闻资讯下载器 —— 15 个新闻站点抓取
============================================================
每个抓取函数签名：
    fetch_xxx(selected_dates: set[datetime.date], work_dir: str)
        -> dict[str, list[dict]]   # { 'YYYY-MM-DD': [article, ...] }

article 结构：
    {'title', 'meta', 'paragraphs': [..], 'images': [url..], 'video_links': [..]}

在原始“仅文字”脚本基础上，为每个站点新增：
  - 图片（新闻配图，随文嵌入 docx）
  - 视频链接（video/iframe/播放链接，在 docx 中【视频】列出）
并按“今天 / 前一天”勾选分组返回。
"""
import datetime
import hashlib
import html as html_lib
import json
import os
import random
import re
import time
import urllib.parse

from .core import (
    http_get, http_download, soup_of, ensure_dir, extract_images,
    extract_videos, clean_paragraphs, IMG_EXTS,
)
from .core import today_date, yesterday_date, UA

# 可选：Selenium（腾讯两个爬虫需要）
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.common.by import By as SeleniumBy
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False


def _article(title='', meta='', paragraphs=None, images=None, video_links=None):
    return {
        'title': (title or '').strip(),
        'meta': meta or '',
        'paragraphs': paragraphs or [],
        'images': images or [],
        'video_links': video_links or [],
    }


def _push(result, date_obj, art):
    key = date_obj.isoformat()
    result.setdefault(key, []).append(art)


def _keep(selected_dates, dt):
    return dt in selected_dates


# =====================================================================
# 湖北日报 通用（热推 / 通城 / 要闻 共用）
# =====================================================================
HBRB_API = "https://hbrbapi.hubeidaily.net/amc/client/listContentByColumn"
HBRB_REFERER = "https://news.hubeidaily.net/"


def _hbrb_list(column_id, page_no, page_size=21, device_id=None, retries=2):
    if device_id is None:
        device_id = ''.join(random.choice('0123456789abcdef') for _ in range(32))
    params = {
        "column": column_id, "deviceId": device_id, "focusNo": 5,
        "pageNo": page_no, "pageSize": page_size,
    }
    last_err = None
    for _ in range(retries + 1):
        try:
            t = int(time.time() * 1000)
            s1 = hashlib.md5(f"hbrb-app-amc${t}".encode("utf-8")).hexdigest()
            token = hashlib.md5(f"h5Client-id${s1}${t}".encode("utf-8")).hexdigest()
            headers = {"token": token, "requestTime": str(t),
                       "User-Agent": UA, "Referer": HBRB_REFERER}
            r = http_get(HBRB_API, params=params, headers=headers, timeout=30)
            data = r.json()
            if data.get("suc") == 1:
                return data.get("data", {}) or {}
            last_err = f"api suc={data.get('suc')}"
        except Exception as e:
            last_err = str(e)
        time.sleep(1)
    return {}


def _hbrb_fetch_html(url, retries=2):
    last_err = None
    for _ in range(retries + 1):
        try:
            r = http_get(url, headers={"User-Agent": UA, "Referer": HBRB_REFERER}, timeout=30)
            r.encoding = r.apparent_encoding or "utf-8"
            return r.text
        except Exception as e:
            last_err = str(e)
        time.sleep(1)
    return None


def _hbrb_detail(url, fallback_title):
    """解析湖北日报详情页 -> article（含图片/视频）。"""
    html = _hbrb_fetch_html(url)
    if not html:
        return None
    soup = soup_of(html)
    title_el = soup.select_one("#news-title") or soup.select_one(".news-title")
    title = title_el.get_text(strip=True) if title_el else fallback_title
    pub_date = None
    src_item = soup.select_one(".source .source-item")
    if src_item:
        m = re.search(r"(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})", src_item.get_text())
        if m:
            y, mo, d = (int(x) for x in m.groups())
            pub_date = datetime.date(y, mo, d)
    box = soup.select_one(".detail-cont")
    if box is None:
        return None
    images = extract_images(box, base_url='https://news.hubeidaily.net/')
    videos = extract_videos(box, base_url='https://news.hubeidaily.net/')
    paras = clean_paragraphs(box, skip_images=True)
    if not paras and not images:
        return None
    return {
        'title': title, 'pub_date': pub_date, 'paragraphs': paras,
        'images': images, 'video_links': [{'label': '新闻视频', 'url': v} for v in videos],
    }


# ---------------------------------------------------------------------
# 1. 湖北日报热推新闻
# ---------------------------------------------------------------------
def fetch_hubei_hot(selected_dates, work_dir):
    result = {}
    seen_ids = set()
    hit_too_old = False
    for page_no in range(0, 8):
        if hit_too_old:
            break
        data = _hbrb_list(1476, page_no, 21, 'hubei-daily-crawler')
        content_list = data.get("contentList", []) or []
        if not content_list:
            break
        for item in content_list:
            cid = item.get("contentId")
            if item.get("contentType") in (7, 8):
                continue
            pc_url = item.get("pcUrl") or item.get("shareUrl") or item.get("mobileUrl")
            if not pc_url or cid in seen_ids:
                continue
            seen_ids.add(cid)
            art = _hbrb_detail(pc_url, item.get("title", ""))
            if not art:
                continue
            pd = art['pub_date']
            if pd is not None:
                if pd < min(selected_dates) if selected_dates else False:
                    hit_too_old = True
                    break
                if not _keep(selected_dates, pd):
                    continue
            else:
                continue
            _push(result, pd, _article(art['title'], paragraphs=art['paragraphs'],
                                       images=art['images'], video_links=art['video_links']))
            time.sleep(0.2)
    return result


# ---------------------------------------------------------------------
# 2. 湖北日报通城新闻
# ---------------------------------------------------------------------
def fetch_hubei_tongcheng(selected_dates, work_dir):
    want = {d.isoformat() for d in selected_dates}
    result = {}
    device_id = ''.join(random.choice('0123456789abcdef') for _ in range(32))
    seen = set()
    for page_no in range(10):
        data = _hbrb_list('1691', page_no, 21, device_id)
        if not data:
            break
        batch = list(data.get('focusList') or []) + list(data.get('contentList') or [])
        if not batch:
            break
        for item in batch:
            pub = (item.get('publishDetailedTime') or '').strip()
            day = pub[:10]
            if day in want and item.get('contentId') not in seen:
                seen.add(item['contentId'])
                url = (item.get('mobileUrl') or '').strip() or \
                    f"https://news.hubeidaily.net/mobile/c_{item.get('contentId')}.html"
                art = _hbrb_detail(url, item.get('contentTitle') or item.get('title') or '')
                if art and art['pub_date'] and _keep(selected_dates, art['pub_date']):
                    _push(result, art['pub_date'],
                          _article(art['title'], paragraphs=art['paragraphs'],
                                   images=art['images'], video_links=art['video_links']))
    return result


# ---------------------------------------------------------------------
# 3. 湖北日报要闻新闻
# ---------------------------------------------------------------------
def fetch_hubei_yaowen(selected_dates, work_dir):
    result = {}
    seen = set()
    for page_no in range(10):
        data = _hbrb_list('1478', page_no, 21, "pc")
        if not data:
            break
        items = data.get("contentList") or []
        if not items:
            break
        for item in items:
            pub_str = (item.get("publishDetailedTime") or "").strip()
            try:
                pub_date = datetime.datetime.strptime(pub_str, "%Y-%m-%d %H:%M:%S").date()
            except ValueError:
                continue
            if not _keep(selected_dates, pub_date):
                continue
            cid = item.get("contentId")
            url = item.get("pcUrl") or item.get("mobileUrl")
            if not url or cid in seen:
                continue
            seen.add(cid)
            art = _hbrb_detail(url, item.get("contentTitle") or '')
            if art and art['pub_date'] and _keep(selected_dates, art['pub_date']):
                _push(result, art['pub_date'],
                      _article(art['title'], paragraphs=art['paragraphs'],
                               images=art['images'], video_links=art['video_links']))
    return result


# =====================================================================
# 4. 观察者网
# =====================================================================
def fetch_guancha(selected_dates, work_dir):
    base_url = "https://www.guancha.cn"
    session = __import__('requests').Session()
    session.headers.update({'User-Agent': UA})
    today = today_date()
    keys = {d.strftime("%Y_%m_%d") for d in selected_dates}
    result = {}

    def get_html(url):
        try:
            resp = session.get(url, timeout=20)
            resp.encoding = resp.apparent_encoding or "utf-8"
            return resp.text
        except Exception:
            return None

    def normalize_url(href):
        if not href:
            return None
        href = href.strip()
        if href.startswith(("http://", "https://")):
            return href
        if href.startswith("//"):
            return "https:" + href
        if href.startswith("/"):
            return base_url + href
        return None

    def first_links():
        html = get_html(base_url)
        if not html:
            return []
        soup = soup_of(html)
        out, seen = [], set()
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            full = normalize_url(href)
            if not full:
                continue
            m1 = re.search(r"/([^/]+)/(\d{4}_\d{2}_\d{2})_\d+\.shtml", full)
            if m1:
                date_str = m1.group(2)
                clean = re.sub(r"(\d{4}_\d{2}_\d{2}_\d+\.shtml).*$", r"\1", full)
                if clean not in seen:
                    seen.add(clean)
                    out.append((a.get_text(strip=True), clean, date_str))
        return out

    def fetch_first(url):
        html = get_html(url)
        if not html:
            return [], []
        soup = soup_of(html)
        node = soup.select_one("div.main.content-main") or soup.select_one("div.article-content")
        if node is None:
            return [], []
        images = extract_images(node, base_url=base_url)
        videos = extract_videos(node, base_url=base_url)
        paras = clean_paragraphs(node, skip_images=True)
        if "余下全文" in "\n".join(paras):
            s_url = url[:-len(".shtml")] + "_s.shtml" if url.endswith(".shtml") else url + "_s"
            html2 = get_html(s_url)
            if html2:
                soup2 = soup_of(html2)
                node2 = soup2.select_one("div.main.content-main") or soup2.select_one("div.article-content")
                if node2 is not None:
                    images = extract_images(node2, base_url=base_url) or images
                    videos = extract_videos(node2, base_url=base_url) or videos
                    paras = clean_paragraphs(node2, skip_images=True) or paras
        return paras, images + videos

    for title, url, date_str in first_links():
        if date_str not in keys:
            continue
        paras, media = fetch_first(url)
        if not paras and not media:
            continue
        dt = datetime.date(*(int(x) for x in date_str.split("_")))
        _push(result, dt, _article(title, paragraphs=paras,
                                   images=[m for m in media if not m.lower().endswith(tuple(
                                       ('.mp4', '.m3u8', '.flv', '.webm')))],
                                   video_links=[{'label': '相关视频', 'url': m} for m in media
                                                if m.lower().endswith(('.mp4', '.m3u8', '.flv', '.webm'))]))
        time.sleep(0.2)
    return result


# =====================================================================
# 5. 金十数据市场快讯
# =====================================================================
def fetch_jin10(selected_dates, work_dir):
    API_URL = "https://flash-api.jin10.com/get_flash_list"
    headers = {
        "User-Agent": UA, "x-app-id": "bVBF4FyRTn5NJF5n", "x-version": "1.0.0",
        "Referer": "https://www.jin10.com/",
        "Accept": "application/json, text/plain, */*",
    }

    def fetch_page(max_time=None):
        params = {"channel": -8200, "vip": 1}
        if max_time:
            params["max_time"] = max_time
        last_err = None
        for attempt in range(5):
            try:
                resp = http_get(API_URL, params=params, headers=headers, timeout=20)
                data = resp.json()
                if data.get("status") != 200:
                    raise RuntimeError(f"API status={data.get('status')}")
                return data.get("data") or []
            except Exception as e:
                last_err = e
                time.sleep(1.5 * (attempt + 1))
        raise RuntimeError(f"请求接口失败: {last_err}")

    want = {d.isoformat() for d in selected_dates}
    earliest = min(selected_dates).strftime("%Y-%m-%d 00:00:00")
    collected, seen_ids = [], set()
    cursor, page_no = None, 0
    while page_no < 300:
        page = fetch_page(cursor)
        if not page:
            break
        for item in page:
            if item.get("id") in seen_ids:
                continue
            seen_ids.add(item.get("id"))
            collected.append(item)
        page_no += 1
        oldest = page[-1].get("time", "") if page else ""
        if not oldest or oldest < earliest:
            break
        cursor = oldest
        time.sleep(0.3)

    result = {}
    for item in collected:
        t = item.get("time", "")
        if not t or t[:10] not in want:
            continue
        d = item.get("data") or {}
        content = html_lib.unescape(re.sub(r"<[^>]+>", "", d.get("content") or "")).strip()
        title = html_lib.unescape((d.get("title") or "").strip())
        if not title:
            m = re.match(r"^【([^】]+)】", content)
            if m:
                title = m.group(1).strip()
                content = content[m.end():].strip()
            elif content:
                title = re.split(r"[，。；;]", content, maxsplit=1)[0].strip()[:40] or "市场快讯"
            else:
                continue
        dt = datetime.date.fromisoformat(t[:10])
        display = t[11:] if len(t) >= 19 else (t[5:] if len(t) >= 5 else '')
        _push(result, dt, _article(f'{display} {title}',
                                   meta='金十数据市场快讯', paragraphs=[content] if content else []))
    # 每条按时序
    for k in result:
        result[k] = sorted(result[k], key=lambda a: a['title'])
    return result


# =====================================================================
# 6. 澎湃新闻
# =====================================================================
def fetch_thepaper(selected_dates, work_dir):
    base_url = "https://www.thepaper.cn"
    want = {d.isoformat() for d in selected_dates}
    session = __import__('requests').Session()
    session.headers.update({
        "User-Agent": UA, "Referer": "https://www.thepaper.cn/",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    })
    crawled_ids = set()
    news_list = []
    channels = [(25949, "要闻"), (25489, "舆论场"), (25432, "自贸区连线"),
                (122903, "澎湃世界观"), (25464, "澎湃质量观"), (25433, "地产界")]

    def get_page(url, retry=3):
        for i in range(retry):
            try:
                resp = session.get(url, timeout=15)
                resp.raise_for_status()
                resp.encoding = "utf-8"
                return resp.text
            except Exception:
                if i < retry - 1:
                    time.sleep(1)
        return None

    def parse_next(html):
        soup = soup_of(html)
        tag = soup.find("script", id="__NEXT_DATA__")
        if not tag:
            return None
        try:
            return json.loads(tag.string)
        except json.JSONDecodeError:
            return None

    def add_news(item):
        cid = str(item.get("contId", ""))
        if not cid or cid in crawled_ids:
            return
        pub = item.get("publishTime", "") or item.get("pubTime", "")
        if not pub or pub[:10] not in want:
            return
        crawled_ids.add(cid)
        news_list.append({"contId": cid,
                          "title": (item.get("name", "") or "").strip(),
                          "pubTime": pub, "url": f"{base_url}/newsDetail_forward_{cid}"})

    html = get_page(base_url)
    if html:
        data = parse_next(html)
        if data:
            page_data = data.get("props", {}).get("pageProps", {}).get("data", {})
            for group in page_data.get("recommendTxt", []):
                for item in (group if isinstance(group, list) else [group]):
                    add_news(item)
            for item in page_data.get("recommendImg", []):
                add_news(item)
            for ch in page_data.get("recommendChannels", []):
                for item in ch.get("contentList", []):
                    add_news(item)
    for channel_id, _ in channels:
        html = get_page(f"{base_url}/list_{channel_id}")
        if not html:
            continue
        data = parse_next(html)
        if not data:
            continue
        page_data = data.get("props", {}).get("pageProps", {}).get("data", {})
        for key in ["contentList", "recommendList", "list"]:
            for item in page_data.get(key, []) or []:
                if isinstance(item, dict):
                    add_news(item)

    result = {}
    for news in news_list:
        html = get_page(news["url"])
        if not html:
            continue
        data = parse_next(html)
        if not data:
            continue
        detail = (data.get("props", {}).get("pageProps", {}).get("detailData", {})
                  or {}).get("contentDetail", {})
        if not detail:
            continue
        raw = detail.get("content", "") or ""
        soup = soup_of(raw)
        images = extract_images(soup, base_url=base_url)
        videos = extract_videos(soup, base_url=base_url)
        for tag in soup.find_all(["img", "video", "audio", "iframe", "embed", "source", "figure"]):
            tag.decompose()
        paras = [re.sub(r"\s+", " ", ln).strip() for ln in
                 soup.get_text(separator="\n", strip=True).split("\n") if ln.strip()]
        pub = detail.get("pubTime", news["pubTime"])
        try:
            dt = datetime.date.fromisoformat(pub[:10])
        except ValueError:
            continue
        if not _keep(selected_dates, dt):
            continue
        meta = f"发布时间：{pub}  |  来源：澎湃新闻"
        _push(result, dt, _article(detail.get("name", news["title"]), meta=meta,
                                   paragraphs=paras, images=images,
                                   video_links=[{'label': '新闻视频', 'url': v} for v in videos]))
        time.sleep(0.2)
    for k in result:
        result[k] = sorted(result[k], key=lambda a: a['meta'], reverse=True)
    return result


# =====================================================================
# 腾讯通用（Selenium 列表 + r.inews 详情）
# =====================================================================
def _tencent_fetch_list(base_url, scroll_times=3):
    if not HAS_SELENIUM:
        raise RuntimeError("未安装 selenium，无法抓取腾讯新闻")
    options = ChromeOptions()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument(f'user-agent={UA}')
    options.add_argument('--window-size=1920,1080')
    driver = webdriver.Chrome(options=options)
    news_list, crawled = [], set()
    try:
        driver.get(base_url)
        time.sleep(3)
        for _ in range(scroll_times):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
        for elem in driver.find_elements(SeleniumBy.CSS_SELECTOR, 'a[href*="/rain/a/"]'):
            try:
                href = elem.get_attribute("href") or ""
                m = re.search(r'/a/([A-Z0-9]+)', href)
                if not m:
                    continue
                nid = m.group(1)
                if nid in crawled:
                    continue
                title = elem.text.strip()
                title = re.sub(r'\s*\d+分钟前\s*$', '', title)
                title = re.sub(r'\s*\d+小时前\s*$', '', title)
                title = re.sub(r'^\s*(热点精选|独家)\s*', '', title)
                if len(title) < 5:
                    continue
                crawled.add(nid)
                news_list.append({"newsId": nid, "title": title,
                                  "url": href.split("?")[0]})
            except Exception:
                continue
    finally:
        driver.quit()
    return news_list


def _tencent_detail(nid, session):
    try:
        resp = session.get(f"https://r.inews.qq.com/getSimpleNews?id={nid}", timeout=15)
        resp.raise_for_status()
        return resp.json()
    except Exception:
        return None


def _tencent_media(detail, base_url):
    images, videos = [], []
    attr = detail.get("attribute", {}) or {}
    for key, val in attr.items():
        if key.startswith("VIDEO_") and isinstance(val, dict):
            u = val.get("playurl") or ""
            if u:
                videos.append({'label': val.get("title") or '相关视频', 'url': u})
        elif key.startswith("IMAGE_") and isinstance(val, dict):
            u = val.get("url") or val.get("imgurl") or ""
            if u and u.startswith(("http://", "https://")):
                images.append(u)
    content_text = (detail.get("content", {}) or {}).get("text", "") or ""
    if content_text:
        soup = soup_of(content_text)
        images += extract_images(soup, base_url=base_url)
        videos += [{'label': '相关视频', 'url': v} for v in extract_videos(soup, base_url=base_url)]
    # 去重
    vis, vus = [], set()
    for v in videos:
        if v['url'] not in vus:
            vus.add(v['url']); vis.append(v)
    iis, ius = [], set()
    for i in images:
        if i not in ius:
            ius.add(i); iis.append(i)
    return iis, vis


def _tencent_crawl(base_url, channel, selected_dates, scroll_times):
    want = {d.isoformat() for d in selected_dates}
    session = __import__('requests').Session()
    session.headers.update({"User-Agent": UA, "Referer": "https://news.qq.com/"})
    news_list = _tencent_fetch_list(base_url, scroll_times)
    if not news_list:
        return {}
    result = {}
    for news in news_list:
        detail = _tencent_detail(news["newsId"], session)
        if not detail or not detail.get("id"):
            continue
        pub = detail.get("time", "") or ""
        if pub[:10] not in want:
            continue
        try:
            dt = datetime.date.fromisoformat(pub[:10])
        except ValueError:
            continue
        if not _keep(selected_dates, dt):
            continue
        raw = (detail.get("content", {}) or {}).get("text", "") or ""
        soup = soup_of(raw)
        for tag in soup.find_all(["img", "video", "audio", "iframe", "embed", "source"]):
            tag.decompose()
        paras = [re.sub(r"\s+", " ", ln).strip() for ln in
                 soup.get_text(separator="\n", strip=True).split("\n") if ln.strip()]
        images, videos = _tencent_media(detail, "https://news.qq.com/")
        meta = f"发布时间：{pub}  |  栏目：{channel}  |  来源：{detail.get('source', '') or '腾讯新闻'}"
        _push(result, dt, _article(detail.get("title", news["title"]), meta=meta,
                                   paragraphs=paras, images=images, video_links=videos))
        time.sleep(0.2)
    return result


# ---------------------------------------------------------------------
# 7. 腾讯国际新闻
# ---------------------------------------------------------------------
def fetch_tencent_world(selected_dates, work_dir):
    if not HAS_SELENIUM:
        raise RuntimeError("未安装 selenium，无法抓取腾讯国际新闻")
    return _tencent_crawl("https://news.qq.com/ch/world/", "国际", selected_dates, 3)


# ---------------------------------------------------------------------
# 8. 腾讯要闻新闻
# ---------------------------------------------------------------------
def fetch_tencent_topnews(selected_dates, work_dir):
    if not HAS_SELENIUM:
        raise RuntimeError("未安装 selenium，无法抓取腾讯要闻新闻")
    return _tencent_crawl("https://news.qq.com/", "要闻", selected_dates, 5)


# =====================================================================
# 网易 通用
# =====================================================================
_NETEASE_HEADERS = {"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9"}


def _netease_parse_detail(url):
    r = http_get(url, headers=_NETEASE_HEADERS, timeout=15)
    r.encoding = 'utf-8'
    soup = soup_of(r.text)
    h1 = soup.find('h1')
    title = h1.get_text(strip=True) if h1 else ''
    body = soup.find('div', class_='post_body') or \
        (soup.find('div', class_='post_content') or soup).find('div', class_='post_body')
    paras, images, videos = [], [], []
    if body:
        images = extract_images(body, base_url=url)
        videos = extract_videos(body, base_url=url)
        for p in body.find_all('p'):
            if p.find('img'):
                continue
            t = p.get_text(strip=True)
            if t:
                paras.append(t)
    return title, paras, images, videos


def _netease_channel_crawl(list_url, selected_dates):
    r = http_get(list_url, headers=_NETEASE_HEADERS, timeout=15)
    r.encoding = 'gbk'
    m = re.search(r'data_callback\((\[.*\])\)', r.text, re.DOTALL)
    if not m:
        raise RuntimeError("接口格式变更，未获取到新闻列表")
    news_list = json.loads(m.group(1))
    result = {}
    for news in news_list:
        ts = news.get("time", "")
        try:
            pub = datetime.datetime.strptime(ts, "%m/%d/%Y %H:%M:%S")
        except (ValueError, TypeError):
            continue
        if not _keep(selected_dates, pub.date()):
            continue
        url = news.get("docurl", "")
        if not url:
            continue
        try:
            title, paras, images, videos = _netease_parse_detail(url)
        except Exception:
            continue
        if not title or not paras:
            continue
        _push(result, pub.date(),
              _article(title, meta=f"发布时间：{pub:%Y-%m-%d %H:%M}  |  来源：网易新闻",
                       paragraphs=paras, images=images,
                       video_links=[{'label': '相关视频', 'url': v} for v in videos]))
        time.sleep(0.2)
    return result


# ---------------------------------------------------------------------
# 9. 网易国际新闻（首页 + 详情）
# ---------------------------------------------------------------------
def fetch_netease_world(selected_dates, work_dir):
    INDEX_URL = "https://news.163.com/world/"
    PATTERNS = [re.compile(r"https?://www\.163\.com/(?:dy|news)/article/[A-Z0-9]+\.html")]
    STOP = ("相关新闻", "相关推荐", "责任编辑", "来源：", "延伸阅读", "版权声明", "扫码关注")

    def fetch(url, retries=3):
        for i in range(retries):
            try:
                r = requests_get(url)
                return r
            except Exception:
                time.sleep(1.5 * (i + 1))
        return None

    def requests_get(url):
        r = http_get(url, headers=_NETEASE_HEADERS, timeout=20)
        r.encoding = r.apparent_encoding or "utf-8"
        return r

    index = fetch(INDEX_URL)
    if not index:
        raise RuntimeError("网易国际首页抓取失败")
    urls = set()
    for pat in PATTERNS:
        for m in pat.finditer(index.text):
            urls.add(m.group(0))
    result = {}
    seen_titles = set()
    for url in sorted(urls):
        html = fetch(url)
        if not html:
            continue
        title, paras, images, videos = _netease_parse_detail(url)
        # 解析发布日期
        pub_date = _netease_extract_date(html.text)
        if pub_date is None or not _keep(selected_dates, pub_date):
            continue
        if not paras or title in seen_titles:
            continue
        seen_titles.add(title)
        _push(result, pub_date,
              _article(title, meta=f"发布时间：{pub_date}  |  来源：网易新闻",
                       paragraphs=paras, images=images,
                       video_links=[{'label': '相关视频', 'url': v} for v in videos]))
        time.sleep(0.3)
    return result


def _netease_extract_date(text):
    patterns = [
        r'(?:time|ptime|publishTime)["\']?\s*(?::|=)\s*["\'](\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2})?)',
        r'datePublished["\']?\s*(?::|=)\s*["\'](\d{4}-\d{2}-\d{2}T\d{2}:\d{2})',
        r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}(?::\d{2})?)',
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            s = m.group(1).replace("T", " ")
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
                try:
                    return datetime.datetime.strptime(s, fmt).date()
                except ValueError:
                    continue
    return None


# ---------------------------------------------------------------------
# 10. 网易国内新闻 / 11. 网易军事新闻
# ---------------------------------------------------------------------
def fetch_netease_domestic(selected_dates, work_dir):
    return _netease_channel_crawl("https://temp.163.com/special/00804KVA/cm_guonei.js",
                                  selected_dates)


def fetch_netease_military(selected_dates, work_dir):
    return _netease_channel_crawl("https://temp.163.com/special/00804KVA/cm_war.js",
                                  selected_dates)


# =====================================================================
# 12. 新华网国际新闻
# =====================================================================
def fetch_xinhua_world(selected_dates, work_dir):
    INDEX_URL = "http://www.news.cn/world/index.html"
    HEADERS = {"User-Agent": UA}
    segs = {d.strftime("%Y%m%d") for d in selected_dates}

    def fetch(url):
        r = http_get(url, headers=HEADERS, timeout=20)
        r.encoding = r.apparent_encoding or "utf-8"
        time.sleep(0.3)
        return r

    soup = soup_of(fetch(INDEX_URL).text)
    news_list = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        title = a.get_text(strip=True)
        if len(title) < 6:
            continue
        if not href.startswith("http"):
            href = urllib.parse.urljoin(INDEX_URL, href)
        if any(f"/{seg}/" in href for seg in segs) and href.rstrip("/").endswith("c.html"):
            news_list.append((title, href))
    result = {}
    for title, url in news_list:
        try:
            d_soup = soup_of(fetch(url).text)
            h1 = d_soup.find("h1")
            t = h1.get_text(strip=True) if h1 else ""
            node = d_soup.find(id="detail") or d_soup
            paras = [p.get_text(strip=True) for p in node.find_all("p")
                     if p.get_text(strip=True)]
            images = extract_images(d_soup, base_url="http://www.news.cn")
            videos = extract_videos(d_soup, base_url="http://www.news.cn")
        except Exception:
            continue
        if not t or not paras:
            continue
        # 用 URL 中的日期段判断
        dt = _date_from_url(url, segs)
        if dt is None or not _keep(selected_dates, dt):
            continue
        _push(result, dt, _article(t, meta=f"来源：新华网", paragraphs=paras,
                                   images=images,
                                   video_links=[{'label': '相关视频', 'url': v} for v in videos]))
    return result


def _date_from_url(url, segs):
    for seg in segs:
        m = re.search(rf"/{seg}/", url)
        if m:
            return datetime.datetime.strptime(seg, "%Y%m%d").date()
    return None


# =====================================================================
# 13. 新华网湖北新闻
# =====================================================================
def fetch_xinhua_hubei(selected_dates, work_dir):
    INDEX_URL = "http://www.hb.news.cn/index.htm"
    BASE_URL = "http://www.hb.news.cn/"
    HEADERS = {"User-Agent": UA}
    segs = {d.strftime("%Y%m%d") for d in selected_dates}
    PATTERN = re.compile(r"/(\d{8})/[0-9a-fA-F]{20,}/c\.html")

    r = http_get(INDEX_URL, headers=HEADERS, timeout=15)
    r.encoding = "utf-8"
    soup = soup_of(r.text)
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith(("javascript:", "#", "mailto:")):
            continue
        full = urllib.parse.urljoin(BASE_URL, href)
        m = PATTERN.search(full)
        if not m or m.group(1) not in segs:
            continue
        if not re.search(r"https?://(?:www\.)?(?:hb\.news\.cn|news\.cn)/", full):
            continue
        key = urllib.parse.urlparse(full).path
        if key not in [x[0] for x in links]:
            links.append((full, m.group(1)))
    result = {}
    for url, date_seg in links:
        try:
            r2 = http_get(url, headers=HEADERS, timeout=15)
            r2.encoding = "utf-8"
            d_soup = soup_of(r2.text)
            title_el = d_soup.select_one("h1 span.title") or d_soup.find("span", class_="title")
            title = title_el.get_text(strip=True) if title_el else ""
            if not title:
                continue
            node = d_soup.find("span", id="detailContent")
            paras = []
            if node:
                for el in node.find_all(["p", "figcaption"]):
                    txt = re.sub(r"\s+", " ", el.get_text().replace("\xa0", " ").replace("\u3000", " ")).strip()
                    if txt:
                        paras.append(txt)
            images = extract_images(d_soup, base_url=BASE_URL)
            videos = extract_videos(d_soup, base_url=BASE_URL)
        except Exception:
            continue
        if not paras:
            continue
        dt = datetime.datetime.strptime(date_seg, "%Y%m%d").date()
        if not _keep(selected_dates, dt):
            continue
        _push(result, dt, _article(title, meta=f"来源：新华网湖北", paragraphs=paras,
                                   images=images,
                                   video_links=[{'label': '相关视频', 'url': v} for v in videos]))
        time.sleep(0.3)
    return result


# =====================================================================
# 14. 新华网即时新闻
# =====================================================================
def fetch_xinhua_jsxw(selected_dates, work_dir):
    BASE = "http://www.news.cn"
    LIST_JSON = "http://www.news.cn/world/jsxw/ds_29089f6bdec84f03b12804d9fe4897be.json"
    HEADERS = {"User-Agent": UA, "Referer": "http://www.news.cn/world/jsxw/index.html"}
    want = {d.isoformat() for d in selected_dates}

    r = http_get(LIST_JSON, headers=HEADERS, timeout=20)
    r.encoding = "utf-8"
    try:
        data = json.loads(r.text)
    except json.JSONDecodeError:
        raise RuntimeError("新华网即时列表 JSON 解析失败")
    picked = []
    for it in data.get("datasource", []):
        pt = (it.get("publishTime") or "").strip()
        if pt[:10] in want:
            picked.append({"time": pt,
                           "title": (it.get("showTitle") or "").strip(),
                           "url": BASE + it.get("publishUrl", "")})
    result = {}
    for n in picked:
        try:
            r2 = http_get(n["url"], headers=HEADERS, timeout=20)
            r2.encoding = "utf-8"
            d_soup = soup_of(r2.text)
            h1 = d_soup.select_one("h1 span.title") or d_soup.find("span", class_="title")
            title = h1.get_text(strip=True) if h1 else ""
            if not title:
                continue
            box = d_soup.find(id="detailContent") or d_soup.find(id="detail")
            paras = [re.sub(r"\s+", " ", p.get_text()).strip()
                     for p in box.find_all("p") if p.get_text(strip=True)] if box else []
            images = extract_images(d_soup, base_url=BASE)
            videos = extract_videos(d_soup, base_url=BASE)
        except Exception:
            continue
        if not paras:
            continue
        dt = datetime.date.fromisoformat(n["time"][:10])
        if not _keep(selected_dates, dt):
            continue
        _push(result, dt, _article(title, meta=f"发布时间：{n['time']}  |  来源：新华网",
                                   paragraphs=paras, images=images,
                                   video_links=[{'label': '相关视频', 'url': v} for v in videos]))
        time.sleep(0.2)
    return result


# =====================================================================
# 15. 云上通城
# =====================================================================
def fetch_yunshang(selected_dates, work_dir):
    BASE_URL = "https://www.hbtctv.cn/"
    HEADERS = {"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9", "Referer": BASE_URL}
    want = {d.isoformat() for d in selected_dates}

    def fetch(url, session, retries=3):
        last_err = None
        for i in range(retries):
            try:
                resp = session.get(url, headers=HEADERS, timeout=20)
                resp.raise_for_status()
                resp.encoding = resp.apparent_encoding or "utf-8"
                return resp.text
            except Exception as e:
                last_err = e
                time.sleep(1.5 * (i + 1))
        raise RuntimeError(f"请求失败: {url} -> {last_err}")

    session = __import__('requests').Session()
    home = fetch(BASE_URL, session)
    # 解析首页文章列表
    articles, seen = [], set()
    for li in re.findall(r"<li>(.*?)</li>", home, re.S):
        m = re.search(r'href="(/p/(\d+).html)"[^>]*title="([^"]*)"', li)
        t = re.search(r"<time>(.*?)</time>", li, re.S)
        if m and t:
            url, aid, title = m.group(1), m.group(2), m.group(3).strip()
            dt = t.group(1).strip()
            if aid not in seen:
                seen.add(aid)
                articles.append({"url": url, "title": title, "datetime": dt})
    for m in re.finditer(
            r'<h2><a href="(/p/(\d+).html)" title="([^"]*)"[^>]*>.*?</a></h2>.*?<time>(.*?)</time>',
            home, re.S):
        url, aid, title, dt = m.group(1), m.group(2), m.group(3).strip(), m.group(4).strip()
        if aid not in seen:
            seen.add(aid)
            articles.append({"url": url, "title": title, "datetime": dt})

    picked = [a for a in articles if a["datetime"][:10] in want]
    picked.sort(key=lambda a: a["datetime"], reverse=True)
    result = {}
    for art in picked:
        full = BASE_URL + art["url"] if art["url"].startswith("/") else art["url"]
        try:
            html = fetch(full, session)
            soup = soup_of(html)
            h1 = soup.select_one("h1.article-title") or soup.select_one("h1.video-title")
            title = h1.get_text(strip=True) if h1 else art["title"]
            # 视频
            video_url = ''
            vc = soup.select_one("div.video-content")
            if vc and vc.find("iframe"):
                video_url = vc.find("iframe").get("src", "")
            ac = soup.select_one("div.article-content")
            if not video_url and ac and ac.find("iframe"):
                video_url = ac.find("iframe").get("src", "")
            images = extract_images(ac, base_url=BASE_URL) if ac else []
            paras = []
            if ac:
                for p in ac.find_all("p"):
                    txt = re.sub(r"\s+", " ", p.get_text(" ", strip=True)).strip()
                    if txt:
                        paras.append(txt)
        except Exception:
            continue
        dt = datetime.date.fromisoformat(art["datetime"][:10])
        if not _keep(selected_dates, dt):
            continue
        videos = [{'label': '新闻视频', 'url': video_url}] if video_url else []
        _push(result, dt, _article(title, meta=f"发布时间：{art['datetime']}  |  来源：云上通城",
                                   paragraphs=paras, images=images, video_links=videos))
    return result


# =====================================================================
# 站点清单
# =====================================================================
SITES = [
    {'id': 'hubei_hot',      'name': '湖北日报热推新闻', 'fn': fetch_hubei_hot,
     'desc': '湖北日报 · 热推栏目（图文+视频）'},
    {'id': 'hubei_tongcheng','name': '湖北日报通城新闻', 'fn': fetch_hubei_tongcheng,
     'desc': '湖北日报 · 通城本地（图文+视频）'},
    {'id': 'hubei_yaowen',   'name': '湖北日报要闻新闻', 'fn': fetch_hubei_yaowen,
     'desc': '湖北日报 · 要闻栏目（图文+视频）'},
    {'id': 'guancha',        'name': '观察者网',         'fn': fetch_guancha,
     'desc': '观察者网 · 首页文章（图文+视频）'},
    {'id': 'jin10',          'name': '金十数据市场快讯', 'fn': fetch_jin10,
     'desc': '金十数据 · 市场快讯（文字为主）'},
    {'id': 'thepaper',       'name': '澎湃新闻',         'fn': fetch_thepaper,
     'desc': '澎湃新闻 · 要闻/栏目（图文+视频）'},
    {'id': 'tencent_world',  'name': '腾讯国际新闻',     'fn': fetch_tencent_world,
     'desc': '腾讯新闻 · 国际（需 Selenium，图文+视频）'},
    {'id': 'tencent_topnews','name': '腾讯要闻新闻',     'fn': fetch_tencent_topnews,
     'desc': '腾讯新闻 · 要闻（需 Selenium，图文+视频）'},
    {'id': 'netease_world',  'name': '网易国际新闻',     'fn': fetch_netease_world,
     'desc': '网易 · 国际（图文+视频）'},
    {'id': 'netease_domestic','name': '网易国内新闻',    'fn': fetch_netease_domestic,
     'desc': '网易 · 国内（图文+视频）'},
    {'id': 'netease_military','name': '网易军事新闻',    'fn': fetch_netease_military,
     'desc': '网易 · 军事（图文+视频）'},
    {'id': 'xinhua_world',   'name': '新华网国际新闻',   'fn': fetch_xinhua_world,
     'desc': '新华网 · 国际（图文+视频）'},
    {'id': 'xinhua_hubei',   'name': '新华网湖北新闻',   'fn': fetch_xinhua_hubei,
     'desc': '新华网 · 湖北（图文+视频）'},
    {'id': 'xinhua_jsxw',    'name': '新华网即时新闻',   'fn': fetch_xinhua_jsxw,
     'desc': '新华网 · 即时（图文+视频）'},
    {'id': 'yunshang',       'name': '云上通城',         'fn': fetch_yunshang,
     'desc': '云上通城 · 通城本地（图文+视频）'},
]

SITES_BY_ID = {s['id']: s for s in SITES}
