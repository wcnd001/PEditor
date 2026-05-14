# -*- mode: python ; coding: utf-8 -*-

import os
import re
import shutil

# 从 PEditor.py 读取 __version__
version = "1.0.0"  # 默认值
try:
    with open('PEditor.py', 'r', encoding='utf-8') as f:
        content = f.read()
        match = re.search(r"__version__\s*=\s*['\"]([^'\"]+)['\"]", content)
        if match:
            version = match.group(1)
except:
    pass

APP_NAME = f"PEditor_v{version}"

a = Analysis(
    ['PEditor.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 这里不要放 PEditor_教程.txt
        # 否则 PyInstaller 6 默认会把它放进 _internal
    ],
    hiddenimports=['gui_helpers'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=APP_NAME,
)

# 只把教程文件单独放到 exe 同级目录，其它依赖仍保持在 _internal
try:
    spec_dir = SPECPATH
except NameError:
    spec_dir = os.path.dirname(os.path.abspath(__file__))

tutorial_src = os.path.join(spec_dir, 'PEditor_教程.txt')
tutorial_dst = os.path.join(DISTPATH, APP_NAME, 'PEditor_教程.txt')

if os.path.exists(tutorial_src):
    os.makedirs(os.path.dirname(tutorial_dst), exist_ok=True)
    shutil.copy2(tutorial_src, tutorial_dst)
else:
    print(f'警告：未找到教程文件：{tutorial_src}')