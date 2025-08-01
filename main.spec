# -*- mode: python ; coding: utf-8 -*-


block_cipher = None
a = Analysis(['ui.py'],
     pathex=['C:/Users/redmiG/anaconda3/envs/pdfOCR/Lib/site-packages/paddleocr', 'C:/Users/redmiG/anaconda3/envs/pdfOCR/Lib/site-packages//paddle/libs',],
     binaries=[('C:/Users/redmiG/anaconda3/envs/pdfOCR/Lib/site-packages/paddle/libs', '.')],
     datas=[('C:/Users/redmiG/anaconda3/envs/pdfOCR/Lib/site-packages/pypdfium2', 'pypdfium2'), 
     ('C:/Users/redmiG/anaconda3/envs/pdfOCR/Lib/site-packages/pypdfium2_raw', 'pypdfium2_raw'),
     ('C:/Users/redmiG/anaconda3/envs/pdfOCR/Lib/site-packages/paddleocr/ppocr', 'ppocr'),
     ('C:/Users/redmiG/anaconda3/envs/pdfOCR/Lib/site-packages/paddleocr/tools', 'tools')],
     hiddenimports=['pdfplumber', 'pypdfium2', 'pypdfium2_raw'],
     hookspath=['.'],
     runtime_hooks=[],
     excludes=['matplotlib'],
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
     name='CustomsExtractor',
     debug=False,
     bootloader_ignore_signals=False,
     strip=False,
     upx=True,
     console=False,
     disable_windowed_traceback=False,
     argv_emulation=False,
     target_arch=None,
     codesign_identity=None,
     entitlements_file=None
)
coll = COLLECT(exe,
     a.binaries,
     a.zipfiles,
     a.datas,
     strip=False,
     upx=True,
     upx_exclude=[],
     name='main'
)
