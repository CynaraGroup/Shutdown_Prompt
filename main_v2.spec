# -*- mode: python ; coding: utf-8 -*-

import os
from pathlib import Path

block_cipher = None

# 获取当前文件所在目录，确保路径兼容不同操作系统
current_dir = Path(__file__).parent.absolute()

a = Analysis(['main_v2.py'],
             pathex=[str(current_dir)],  # 使用当前目录作为路径
             binaries=[],
             # 使用绝对路径引用资源文件，确保在CI环境中能正确找到
             datas=[
                 (str(current_dir / 'dingtalk.ttf'), '.'),
                 (str(current_dir / 'zhengqingke.ttf'), '.'),
                 (str(current_dir / 'logo.ico'), '.')
             ],
             hiddenimports=['win32com.client', 'pythoncom'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='ShutdownPrompt',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          # 使用绝对路径引用图标
          icon=str(current_dir / 'logo.ico'),
          entitlements_file=None )
