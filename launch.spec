# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launch.py'],
    pathex=[],
    binaries=[],
    datas=[('background.png','.')],
    hiddenimports=[
        'PIL',
        'openpyxl',
        'numpy',
        'matplotlib',
        'matplotlib.pyplot',
        'matplotlib.gridspec',
        'matplotlib.colors',
        'PyQt5',
        'PyQt5.QtWidgets',
        'PyQt5.QtGui',
        'PyQt5.QtCore',
        'io'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='双优计划图表生成器',
          debug=False,
          bootloader_ignore_signals=False,
          bootloader_ignore_pycs=False,
          noarchive=False,
          runtime_tmpdir=None,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_hooks=[],
          win_no_prefer_redirects=False,
          win_private_assemblies=False,
          console=True)

'''
app = BUNDLE(exe,
             name='双优计划图表生成器.app',
             icon=None,
             bundle_identifier=None)
'''