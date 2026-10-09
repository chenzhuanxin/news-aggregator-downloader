# -*- coding: utf-8 -*-
"""
新闻资讯下载器 —— Flask 后端
============================================================
在本地启动 Web 服务：前端面板勾选新闻站点、选择「今天/前一天」、
选择保存目录与输出格式后，后台线程逐站抓取（文字+图片+视频链接）
并按日期自动建子文件夹保存 docx（可转 pdf），实时回传进度。

运行：python app.py  （默认 http://127.0.0.1:18763，端口被占自动避让）
"""
import datetime
import os
import shutil
import subprocess
import sys
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import Flask, jsonify, render_template, request

from crawler.core import (
    ensure_dir, date_folder_name, site_filename, build_site_docx,
    today_date, yesterday_date,
)
from crawler.sites import SITES, SITES_BY_ID


def _resource(rel):
    """兼容 PyInstaller 打包：exe 运行时资源在 _MEIPASS 解压目录。"""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# 打包为无控制台(windows)模式时 sys.stdout/stderr 可能为 None，避免崩溃
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w')


app = Flask(__name__,
            template_folder=_resource('templates'),
            static_folder=_resource('static'))


@app.after_request
def _no_cache(resp):
    """禁用浏览器缓存，确保前端总是拿到最新页面/JS（避免旧缓存导致按钮失效）。"""
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp


# 并发抓取线程数（1~8，默认 4）。并发越高越快，但更耗网络/内存、易被目标站限速。
MAX_WORKERS = 4

JOBS = {}
JOBS_LOCK = threading.Lock()
RUNNING_THREADS = {}
RUNNING_THREADS_LOCK = threading.Lock()


def _public_site(s):
    return {'id': s['id'], 'name': s['name'], 'desc': s['desc']}


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/sites')
def sites():
    return jsonify([_public_site(s) for s in SITES])


@app.route('/api/browse', methods=['POST'])
def browse():
    """弹出系统原生“选择文件夹”对话框，返回所选路径。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askdirectory(title='选择新闻资讯保存目录')
        root.destroy()
        if path:
            return jsonify({'path': path})
        return jsonify({'path': None})
    except Exception as e:
        return jsonify({'error': f'无法打开文件夹选择对话框：{e}'}), 500


@app.route('/api/open', methods=['POST'])
def open_dir():
    data = request.get_json(silent=True) or {}
    path = (data.get('path') or '').strip()
    if not path or not os.path.isdir(path):
        return jsonify({'ok': False, 'error': '目录不存在：' + path}), 400
    try:
        os.startfile(path)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': '无法打开目录：' + str(e)}), 500


def _find_soffice():
    """定位 LibreOffice（用于 docx -> pdf）。"""
    cand = shutil.which('soffice')
    if cand:
        return cand
    for base in (r'C:\Program Files\LibreOffice\program\soffice.exe',
                 r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
        if os.path.exists(base):
            return base
    return None


def _convert_pdf(docx_path, pdf_path):
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError('未安装 LibreOffice，无法生成 PDF（已改为保存 docx）')
    tmp_dir = os.path.dirname(pdf_path)
    # soffice 会把 docx 转成同目录同名的 pdf，再改名到目标
    subprocess.run(
        [soffice, '--headless', '--convert-to', 'pdf', '--outdir', tmp_dir, docx_path],
        check=True, capture_output=True, timeout=180,
    )
    auto_pdf = os.path.splitext(docx_path)[0] + '.pdf'
    if os.path.exists(auto_pdf):
        shutil.move(auto_pdf, pdf_path)
    if not os.path.exists(pdf_path):
        raise RuntimeError('PDF 转换失败')


@app.route('/api/run', methods=['POST'])
def run():
    data = request.get_json(silent=True) or {}
    site_ids = data.get('sites') or []
    dates = data.get('dates') or []          # ['today', 'yesterday']
    out_dir = (data.get('out_dir') or '').strip()
    fmt = (data.get('format') or 'docx').strip().lower()

    if not site_ids:
        return jsonify({'error': '请至少勾选一个新闻站点'}), 400
    unknown = [s for s in site_ids if s not in SITES_BY_ID]
    if unknown:
        return jsonify({'error': f'未知的站点：{unknown}'}), 400
    if not dates:
        return jsonify({'error': '请至少勾选一个日期（今天 / 前一天）'}), 400
    if fmt not in ('docx', 'pdf'):
        return jsonify({'error': '输出格式仅支持 docx 或 pdf'}), 400
    if not out_dir:
        return jsonify({'error': '请选择保存目录'}), 400
    if not os.path.isdir(out_dir):
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            return jsonify({'error': f'无法创建保存目录：{e}'}), 400

    date_map = {'today': today_date(), 'yesterday': yesterday_date()}
    selected_dates = [date_map[d] for d in dates if d in date_map]
    if not selected_dates:
        return jsonify({'error': '日期参数无效'}), 400

    job_id = uuid.uuid4().hex[:12]
    max_workers = int(data.get('max_workers') or 0) or MAX_WORKERS
    max_workers = max(1, min(8, max_workers))
    with JOBS_LOCK:
        JOBS[job_id] = {
            'id': job_id,
            'state': 'running',
            'out_dir': os.path.abspath(out_dir),
            'dates': [d.isoformat() for d in selected_dates],
            'format': fmt,
            'total': len(site_ids),
            'max_workers': max_workers,
            'current_index': 0,
            'current_site': '',
            'steps': [],
            'created': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }

    def _worker():
        job = JOBS[job_id]
        try:
            root_out = job['out_dir']
            work_dir = ensure_dir(os.path.join(root_out, '.crawler_tmp'))
            selected = [SITES_BY_ID[s] for s in site_ids]

            # 先预置所有站点步骤，保证前端进度顺序稳定
            with JOBS_LOCK:
                for site in selected:
                    job['steps'].append({'site': site['name'], 'site_id': site['id'],
                                         'status': 'running', 'message': '排队中…',
                                         'files': []})

            def _do_one(site):
                """单个站点：抓取 + 按日期生成文件。返回步骤结果 dict。"""
                if job['state'] != 'running':
                    return {'site_id': site['id'], 'status': 'cancel',
                            'message': '已停止', 'files': []}
                # 每个站点独立临时子目录，避免并发下图片互相覆盖
                site_work = ensure_dir(os.path.join(work_dir, site['id']))
                try:
                    data = site['fn'](selected_dates, site_work)   # {date_iso: [articles]}
                    files, total_articles = [], 0
                    pdf_note = ''
                    for dt in sorted(selected_dates):
                        iso = dt.isoformat()
                        arts = data.get(iso) or []
                        if not arts:
                            continue
                        total_articles += len(arts)
                        folder = ensure_dir(os.path.join(root_out, date_folder_name(dt)))
                        ext = 'pdf' if fmt == 'pdf' else 'docx'
                        fname = site_filename(site['name'], dt, ext)
                        final_path = os.path.join(folder, fname)
                        docx_tmp = os.path.join(site_work, f'tmp_{iso}.docx')
                        build_site_docx(site['name'], dt, arts, docx_tmp, site_work)
                        if fmt == 'pdf':
                            try:
                                _convert_pdf(docx_tmp, final_path)
                                if os.path.exists(docx_tmp):
                                    os.remove(docx_tmp)
                            except Exception as e:
                                final_path = os.path.join(
                                    folder, site_filename(site['name'], dt, 'docx'))
                                shutil.copy(docx_tmp, final_path)
                                if os.path.exists(docx_tmp):
                                    os.remove(docx_tmp)
                                pdf_note = str(e)[:120]
                        else:
                            shutil.move(docx_tmp, final_path)
                        files.append(final_path)
                    if not files:
                        return {'site_id': site['id'], 'status': 'empty',
                                'message': '所选日期无符合条件的新闻', 'files': []}
                    size = sum(os.path.getsize(f) for f in files if os.path.exists(f))
                    msg = (f'共 {total_articles} 篇 · {len(files)} 个文件 · '
                           f'{size / 1024 / 1024:.2f} MB')
                    if pdf_note:
                        msg += '　(转PDF不可用，已存docx)'
                    return {'site_id': site['id'], 'status': 'success',
                            'message': msg, 'files': files}
                except Exception as e:
                    return {'site_id': site['id'], 'status': 'fail',
                            'message': str(e)[:300], 'files': []}

            with ThreadPoolExecutor(max_workers=max_workers) as ex:
                futures = {ex.submit(_do_one, s): s for s in selected}
                for fut in as_completed(futures):
                    if job['state'] == 'stopping':
                        # 停止时取消尚未开始的
                        for f in futures:
                            f.cancel()
                    try:
                        res = fut.result()
                    except Exception as e:
                        res = {'site_id': futures[fut]['id'], 'status': 'fail',
                               'message': str(e)[:300], 'files': []}
                    if not res:
                        continue
                    with JOBS_LOCK:
                        for st in job['steps']:
                            if st['site_id'] == res['site_id']:
                                st['status'] = res['status']
                                st['message'] = res['message']
                                st['files'] = res['files']
                                break

            if job['state'] == 'running':
                job['state'] = 'done'
            elif job['state'] == 'stopping':
                job['state'] = 'stopped'
            shutil.rmtree(work_dir, ignore_errors=True)
        except Exception as e:
            job['state'] = 'error'
            job['steps'].append({'site': '', 'site_id': '', 'status': 'fail',
                                 'message': '全局异常：' + str(e)[:300], 'files': []})

    thread = threading.Thread(target=_worker, daemon=True)
    with RUNNING_THREADS_LOCK:
        RUNNING_THREADS[job_id] = thread
    thread.start()
    return jsonify({'job_id': job_id})


@app.route('/api/jobs/<job_id>')
def job_status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({'error': '任务不存在'}), 404
        snap = dict(job)
        snap['steps'] = list(job['steps'])
        return jsonify(snap)


@app.route('/api/stop', methods=['POST'])
def stop():
    data = request.get_json(silent=True) or {}
    job_id = data.get('job_id')
    with JOBS_LOCK:
        if job_id and job_id in JOBS:
            JOBS[job_id]['state'] = 'stopping'
            return jsonify({'ok': True})
    return jsonify({'ok': False, 'error': '任务不存在'}), 404


DEFAULT_PORT = 18763          # 不常用端口，降低与其他程序冲突的概率


def _already_running():
    """命名互斥量：返回 True 表示已有实例在运行（真正防止重复双击）。"""
    try:
        import ctypes
        handle = ctypes.windll.kernel32.CreateMutexW(None, False,
                                                     'Local\\NewsCrawler_202610')
        err = ctypes.windll.kernel32.GetLastError()
        if err == 183:          # ERROR_ALREADY_EXISTS
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    except Exception:
        return False


def _find_free_port(start):
    """从 start 起找第一个空闲端口，避免端口被占导致启动失败。"""
    import socket
    for p in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('127.0.0.1', p))
                return p
            except OSError:
                continue
    return start


def _port_file():
    import tempfile
    return os.path.join(tempfile.gettempdir(), 'news_crawler_port.txt')


def _save_port(port):
    try:
        with open(_port_file(), 'w', encoding='utf-8') as f:
            f.write(str(port))
    except Exception:
        pass


def _load_port():
    try:
        with open(_port_file(), 'r', encoding='utf-8') as f:
            return int(f.read().strip())
    except Exception:
        return None


def main():
    import webbrowser
    if _already_running():
        # 已有实例在运行：自动打开已有界面并提示，避免重复启动
        p = _load_port() or DEFAULT_PORT
        try:
            webbrowser.open(f'http://127.0.0.1:{p}')
        except Exception:
            pass
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            messagebox.showinfo(
                '新闻资讯下载器',
                '程序已在运行，请勿重复双击。\n已为你打开浏览器界面，直接使用即可。')
            root.destroy()
        except Exception:
            pass
        return

    start = int(os.environ.get('PORT', DEFAULT_PORT))
    port = _find_free_port(start)          # 被占用则自动避让到下一个空闲端口
    _save_port(port)
    url = f'http://127.0.0.1:{port}'

    print('=' * 56)
    print('  新闻资讯下载器 已启动')
    print(f'  浏览器界面:  {url}')
    print('  支持 15 类新闻站点 · 文字+图片+视频链接 · 今天/前一天')
    print('  端口被占用时会自动切换到空闲端口，关闭本窗口即可退出')
    print('=' * 56)
    # 等服务起来后自动打开浏览器
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host='127.0.0.1', port=port, debug=False, threaded=True)


if __name__ == '__main__':
    main()
