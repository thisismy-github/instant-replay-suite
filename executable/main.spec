# -*- mode: python ; coding: utf-8 -*-


import os
import sys
block_cipher = None

CWD = os.path.dirname(os.path.realpath(sys.argv[0]))
SCRIPT_DIR = os.path.dirname(CWD)
VERSION_FILE = os.path.join(CWD, 'version_info_main.txt')
ICON = os.path.join(CWD, 'icon_main.ico')


a = Analysis([os.path.join(SCRIPT_DIR, 'irs.pyw')],
             pathex=[],
             binaries=[],
             datas=[(os.path.join(SCRIPT_DIR, 'resources'), 'resources'),],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
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
          name='InstantReplaySuite',
		  contents_directory='bin',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None,
          version=VERSION_FILE if os.path.exists(VERSION_FILE) else None,
          icon=ICON if os.path.exists(ICON) else None)

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas, 
               strip=False,
               upx=True,
               upx_exclude=[],
               name='release')
