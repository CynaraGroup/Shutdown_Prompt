# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['main_v2.py'],
             pathex=['d:\\Users\\Desktop\\Project\\Shutdown_Prompt'],
             binaries=[],
             datas=[('dingtalk.ttf', '.'), ('zhengqingke.ttf', '.'),('logo.ico', '.')],
             hiddenimports=['win32com.client', 'pythoncom'],  # 添加Office相关隐藏导入
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
          console=False,  # 设置为False不显示控制台窗口
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          icon='./logo.ico'
          entitlements_file=None )
