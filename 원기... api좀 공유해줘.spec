# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

add_files = [ ("./file/g_flag.png", "."), 
      	    ("./file/g_mission.png", "."), 
	    ("./file/g_under.png", "."), 
	    ("./file/bg.png", ".") ]

a = Analysis(['원기... api좀 공유해줘.py'],
             pathex=['C:\\Users\\nerin\\OneDrive\\바탕 화면\\create_exe\\원기... api좀 공유해줘'],
             binaries=[],
             datas=add_files,
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='원기... api좀 공유해줘',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
