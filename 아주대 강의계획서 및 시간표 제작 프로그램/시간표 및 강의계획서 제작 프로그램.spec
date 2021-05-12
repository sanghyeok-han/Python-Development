# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['시간표 및 강의계획서 제작 프로그램.py'],
             pathex=['C:\\Users\\User\\PycharmProjects\\MyProject2\\기타 스크래핑 필요한 프로그램'],
             binaries=[],
             datas=[('data_files/chromedriver78.exe', 'data_files'), ('data_files/chromedriver79.exe', 'data_files'), ('data_files/chromedriver80.exe', 'data_files'), ('data_files/wkhtmltopdf', 'data_files')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='시간표 및 강의계획서 제작 프로그램',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='시간표 및 강의계획서 제작 프로그램')
