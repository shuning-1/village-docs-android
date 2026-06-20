# -*- mode: python ; coding: utf-8 -*-
"""智慧文档管理系统 - PyInstaller 打包配置
目标：最小化 exe 体积，单文件模式，无控制台窗口
关键修复：main.py 通过 importlib 动态加载，需要 collect_all 收集所有依赖
"""

import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# 收集动态加载模块的完整依赖（main.py 通过 importlib 加载，PyInstaller 无法自动检测）
flask_datas, flask_binaries, flask_hidden = collect_all('flask')
jinja2_datas, jinja2_binaries, jinja2_hidden = collect_all('jinja2')
werkzeug_datas, werkzeug_binaries, werkzeug_hidden = collect_all('werkzeug')
openpyxl_datas, openpyxl_binaries, openpyxl_hidden = collect_all('openpyxl')
docx_datas, docx_binaries, docx_hidden = collect_all('docx')
waitress_datas, waitress_binaries, waitress_hidden = collect_all('waitress')

all_datas = [
    ('index.html', '.'),
    ('main.py', '.'),
    ('config.json', '.'),
    ('login_config.json', '.'),
    ('dashboard_config.json', '.'),
    ('accounts.json', '.'),
    ('cert_data.json', '.'),
    ('dismissed_reminders.json', '.'),
    ('app_icon.ico', '.'),
] + flask_datas + jinja2_datas + werkzeug_datas + openpyxl_datas + docx_datas + waitress_datas

all_binaries = flask_binaries + jinja2_binaries + werkzeug_binaries + openpyxl_binaries + docx_binaries + waitress_binaries

all_hidden = flask_hidden + jinja2_hidden + werkzeug_hidden + openpyxl_hidden + docx_hidden + waitress_hidden + [
    'xlrd',
    'webview',
    'webview.platforms.edgechromium',
    'webview.platforms.winforms',
    'pythoncom',
    'win32com',
    'win32com.client',
    'pywintypes',
    'tkinter',
    'tkinter.filedialog',
    'calendar',
    'csv',
    'markupsafe',
    'itsdangerous',
    'click',
]

a = Analysis(
    ['launcher_webview.py'],
    pathex=[],
    binaries=all_binaries,
    datas=all_datas,
    hiddenimports=all_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'unittest',
        'pydoc',
        'doctest',
        'pdb',
        'profile',
        'pstats',
        'test',
        'tests',
        'lib2to3',
        'tkinter.test',
        'unittest.test',
        'numpy',
        'pandas',
        'matplotlib',
        'scipy',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'gtk',
        'gi',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='智慧文档管理系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
)
