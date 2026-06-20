#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""智慧文档管理系统 - 桌面应用启动器 (Waitress + WebView2)
动态端口: 每次启动自动寻找空闲端口，彻底解决多应用端口冲突问题
"""

import sys
import os
import time
import socket
import threading
import traceback
import json
import shutil
import subprocess
from datetime import datetime

# ===== Windows COM 初始化（WebView2 需要） =====
try:
    import pythoncom
    pythoncom.CoInitialize()
except ImportError:
    pass  # 打包后 pywin32 可能缺失，不影响核心功能
except Exception:
    pass

# ===== 路径配置 =====
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    EXE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR = BASE_DIR

# 数据目录：使用安装目录（exe所在目录），所有数据都放在用户安装的文件夹里
DATA_DIR = EXE_DIR
LOG_FILE = os.path.join(DATA_DIR, 'launcher.log')
CONSOLE_LOG = os.path.join(DATA_DIR, 'console.log')
CRASH_LOG = os.path.join(DATA_DIR, 'crash.log')
FATAL_LOG = os.path.join(DATA_DIR, 'fatal_error.log')

os.makedirs(DATA_DIR, exist_ok=True)

# ===== 动态端口：每次启动自动寻找空闲端口，避免与其他应用冲突 =====
def find_free_port():
    """让操作系统自动分配一个空闲端口，用完即释放（waitress 会重新绑定）"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        sock.bind(('127.0.0.1', 0))
        return sock.getsockname()[1]

def show_message_box(msg, title='智慧文档管理系统', style=0x10):
    """显示 Windows 消息框"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, title, style)
    except Exception:
        pass

SERVER_PORT = find_free_port()

# ===== 日志 =====
def log(msg, level='INFO'):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{ts}] [{level}] {msg}'
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except OSError:
        pass
    # 也尝试输出到控制台（如果存在）
    try:
        print(line)
    except Exception:
        pass

# ===== 修复 --noconsole 下 stdout/stderr 为 None =====
def fix_stdout_stderr():
    """将 None 的 stdout/stderr 重定向到 console.log"""
    need_fix = sys.stdout is None or sys.stderr is None
    if not need_fix:
        return
    f = open(CONSOLE_LOG, 'a', encoding='utf-8')
    if sys.stdout is None:
        sys.stdout = f
    if sys.stderr is None:
        sys.stderr = f
    sys.stdout.write(f'[{datetime.now()}] stdout/stderr redirected to console.log\n')
    sys.stdout.flush()

fix_stdout_stderr()

# ===== 复制数据文件 =====
def copy_files_to_data_dir(file_list, overwrite=False):
    """复制文件到数据目录"""
    for fname in file_list:
        src = os.path.join(BASE_DIR, fname)
        dst = os.path.join(DATA_DIR, fname)
        if os.path.exists(src) and (overwrite or not os.path.exists(dst)):
            try:
                shutil.copy2(src, dst)
                log(f'复制文件: {fname}')
            except Exception as e:
                log(f'复制 {fname} 失败: {e}', 'WARN')

copy_files_to_data_dir(['config.json', 'login_config.json', 'dashboard_config.json', 'dismissed_reminders.json', 'cert_data.json', 'accounts.json'])

# ===== WebView2 检测与安装 =====
def check_webview2():
    """检测 WebView2 运行时"""
    import winreg

    def _check_webview2_registry(hive, hive_name):
        """检查注册表中 WebView2 是否已安装"""
        try:
            key = winreg.OpenKey(hive,
                r'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}' if hive == winreg.HKEY_LOCAL_MACHINE else
                r'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}')
            ver, _ = winreg.QueryValueEx(key, 'pv')
            winreg.CloseKey(key)
            log(f'WebView2 已安装 (注册表 {hive_name}), 版本: {ver}')
            return True
        except Exception:
            return False

    # 方法1: 注册表检查
    if _check_webview2_registry(winreg.HKEY_LOCAL_MACHINE, 'HKLM'):
        return True
    if _check_webview2_registry(winreg.HKEY_CURRENT_USER, 'HKCU'):
        return True

    # 方法2: 目录检查
    paths = [
        os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'),
                     'Microsoft', 'EdgeWebView', 'Application', 'msedgewebview2.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''),
                     'Microsoft', 'EdgeWebView', 'Application', 'msedgewebview2.exe'),
    ]
    for p in paths:
        if os.path.exists(p):
            log(f'WebView2 已安装 (文件检测): {p}')
            return True

    # 未安装，尝试自动安装
    log('WebView2 未安装，正在下载...')
    return install_webview2()

def install_webview2():
    """下载并安装 WebView2 Evergreen Bootstrapper"""
    import urllib.request

    # 下载前提示用户
    show_message_box('检测到系统未安装 WebView2 运行时组件。\n\n点击"确定"后将自动下载并安装（约 100-200 MB），请耐心等待。', '智慧文档管理系统 - 正在安装组件', 0x40)

    bootstrapper_url = 'https://go.microsoft.com/fwlink/p/?LinkId=2124701'
    bootstrapper_path = os.path.join(DATA_DIR, 'MicrosoftEdgeWebview2Setup.exe')
    
    try:
        urllib.request.urlretrieve(bootstrapper_url, bootstrapper_path)
        log('WebView2 下载完成，正在静默安装...')
        result = subprocess.run([bootstrapper_path, '/silent', '/install'],
                                capture_output=True, timeout=120)
        log(f'WebView2 安装结果: {result.returncode}')
        try:
            os.remove(bootstrapper_path)
        except Exception:
            pass
        return result.returncode == 0
    except Exception as e:
        log(f'WebView2 安装失败: {e}', 'ERROR')
        return False

# ===== Flask 启动 (使用 waitress) =====
def start_flask():
    """用 waitress 启动 Flask WSGI 应用"""
    import importlib.util
    
    try:
        # 关键: 设置数据目录，让 main.py 知道数据文件在哪里
        os.environ['_data_dir'] = DATA_DIR
        log(f'数据目录设置为: {DATA_DIR}')
        
        # 动态导入 main 模块
        log('导入 main 模块...')
        main_path = os.path.join(BASE_DIR, 'main.py')
        log(f'main.py 路径: {main_path}')
        spec = importlib.util.spec_from_file_location('main', main_path)
        main_module = importlib.util.module_from_spec(spec)
        
        # 关键: 在导入前修改 sys.path 确保 main.py 能找到它的数据文件
        sys.path.insert(0, DATA_DIR)
        
        spec.loader.exec_module(main_module)
        log('main 模块加载成功')
        
        app = main_module.app
        
        # 用 waitress 启动（动态端口，不与其他应用冲突）
        log(f'启动 waitress WSGI 服务器 (127.0.0.1:{SERVER_PORT})...')
        from waitress import serve
        try:
            serve(app, host='127.0.0.1', port=SERVER_PORT, threads=8, 
                  _quiet=True,  # 减少日志输出
                  clear_untrusted_proxy_headers=True)
        except Exception as e:
            log(f'waitress 服务器异常退出: {e}', 'ERROR')
            traceback.print_exc()
    except Exception as e:
        log(f'Flask 启动失败: {e}', 'ERROR')
        log(traceback.format_exc(), 'ERROR')
        # 写入 crash 日志让主线程能读取到错误
        crash_log = CRASH_LOG
        with open(crash_log, 'w', encoding='utf-8') as f:
            f.write(f'Flask 启动失败:\n{traceback.format_exc()}')

# ===== 主函数 =====
def main():
    log('=' * 50)
    log('智慧文档管理系统 启动 (Waitress版本)')
    log(f'Python: {sys.version.split()[0]}')
    log(f'Frozen: {getattr(sys, "frozen", False)}')
    log(f'BASE_DIR: {BASE_DIR}')
    log(f'DATA_DIR: {DATA_DIR}')

    # 检测 WebView2
    log('检测 WebView2 运行时...')
    if not check_webview2():
        log('ERROR: WebView2 安装失败，无法启动浏览器', 'ERROR')
        show_message_box('WebView2 运行时安装失败。\n\n请手动下载安装 Microsoft Edge WebView2:\nhttps://developer.microsoft.com/microsoft-edge/webview2/', '智慧文档管理系统 - 错误')
        return

    # 复制资源文件
    # index.html 是静态资源，每次都复制（确保前端代码最新）
    # 配置文件只在不存在时复制（保留用户已保存的数据）
    # 复制资源文件（每次覆盖确保最新）
    copy_files_to_data_dir(['index.html'], overwrite=True)
    # 复制配置文件（仅首次）
    copy_files_to_data_dir(['config.json', 'login_config.json', 'dashboard_config.json',
                            'dismissed_reminders.json', 'cert_data.json', 'accounts.json'])

    # 启动 Flask/Waitress 线程
    log('启动服务器线程...')
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()

    # 等待服务器就绪
    log('等待服务器就绪...')
    import urllib.request
    max_wait = 30
    started = False
    for i in range(max_wait):
        time.sleep(1)
        try:
            req = urllib.request.Request(f'http://127.0.0.1:{SERVER_PORT}/', method='HEAD')
            resp = urllib.request.urlopen(req, timeout=3)
            if resp.status < 500:
                log(f'服务器就绪 (状态码={resp.status}, 等待了{i+1}秒)')
                started = True
                break
        except urllib.error.HTTPError as e:
            # 404 等也是正常响应（可能首页路由问题）
            log(f'服务器就绪 (状态码={e.code}, 等待了{i+1}秒)')
            started = True
            break
        except Exception:
            # 连接被拒绝，继续等待
            continue

    if not started:
        log('ERROR: 服务器启动超时', 'ERROR')
        show_message_box('服务启动失败。\n\n可能原因:\n1. 系统端口资源不足\n2. 杀毒软件拦截\n3. 系统缺少必要组件\n\n日志文件:\n' + LOG_FILE)
        return

    # 读取窗口标题
    window_title = '智慧文档管理系统'
    try:
        login_config_path = os.path.join(DATA_DIR, 'login_config.json')
        if os.path.exists(login_config_path):
            with open(login_config_path, 'r', encoding='utf-8') as f:
                lc = json.load(f)
            window_title = lc.get('window_title', '智慧文档管理系统') or '智慧文档管理系统'
            log(f'窗口标题: {window_title}')
    except Exception as e:
        log(f'读取窗口标题失败，使用默认: {e}', 'WARN')

    # 启动内置浏览器
    log('启动内置浏览器...')
    try:
        import webview
        
        # JS API: 暴露原生文件夹选择器等能力给前端
        class VillageApi:
            def select_folder_dialog(self):
                """打开原生文件夹选择对话框，返回所选路径"""
                try:
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes('-topmost', True)
                    folder = filedialog.askdirectory(title='选择资料文件夹')
                    root.destroy()
                    return folder if folder else ''
                except Exception as e:
                    log(f'文件夹选择对话框失败: {e}', 'ERROR')
                    return ''

            def set_window_title(self, title):
                """动态修改窗口标题"""
                try:
                    if webview.windows:
                        webview.windows[0].set_title(title)
                    return True
                except Exception as e:
                    log(f'设置窗口标题失败: {e}', 'ERROR')
                    return False
        
        api = VillageApi()
        webview.create_window(window_title, f'http://127.0.0.1:{SERVER_PORT}/',
                              width=1280, height=800, min_size=(900, 600),
                              fullscreen=False, easy_drag=False,
                              js_api=api)
        webview.start(gui='edgechromium')
        log('浏览器窗口已关闭')
    except Exception as e:
        log(f'浏览器启动失败: {e}', 'ERROR')
        traceback.print_exc()
        # 浏览器失败但服务器还在运行，显示消息
        show_message_box(f'浏览器启动失败: {e}\n\n服务器仍在运行，可尝试重新启动程序。')

log('程序退出')

# ===== 入口点 =====
if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        # 终极崩溃捕获
        crash_msg = f'FATAL CRASH:\n{traceback.format_exc()}'
        try:
            with open(FATAL_LOG, 'a', encoding='utf-8') as f:
                f.write(f'[{datetime.now()}]\n{crash_msg}\n\n')
        except Exception:
            pass
        try:
            log(crash_msg, 'FATAL')
        except Exception:
            pass
        # 显示错误消息
        try:
            show_message_box(f'程序发生严重错误:\n\n{str(e)}\n\n详细信息已保存到:\n{FATAL_LOG}', '智慧文档管理系统 - 严重错误')
        except Exception:
            pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
