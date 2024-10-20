# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Gui_module.py'],
    pathex=[
        r'D:\Programming Study\0908-Fist-try\.venv\Lib\site-packages',
        r'D:\Programming Study\0908-Fist-try\LD_Scripts_Tools'
    ],
    binaries=[],
    datas=[('Game+.png', '.')],  # 添加数据文件
    hiddenimports=['numpy'],  # 添加隐藏导入模块
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    exclude_binaries=True,
    name='LD-helpper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 设置为 False 以关闭控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='Game+.ico'  # 指定 .ico 文件作为应用程序图标
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='LD-helpper',
)