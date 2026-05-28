# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['openpyxl', 'openpyxl.utils', 'openpyxl.cell', 'pandas', 'numpy', 'plotly', 'plotly.express', 'plotly.graph_objects', 'dash', 'dash.dcc', 'dash.html', 'psutil']
hiddenimports += collect_submodules('plotly')
hiddenimports += collect_submodules('dash')


a = Analysis(
    ['_Python\\tkinter\\main.py'],
    pathex=['_Python\\tkinter'],
    binaries=[],
    datas=[('Figures', 'Figures'), ('template.inp', '.')],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['win32com', 'pythoncom', 'win32api', 'pywin32_system32', 'xlwings', 'matplotlib', 'matplotlib.tests', 'scipy', 'scipy.tests', 'sklearn', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'IPython', 'ipykernel', 'jupyter', 'notebook', 'pytest', 'numpy.tests', 'numpy.distutils', 'numpy.f2py', 'pandas.tests', 'pandas.io.tests', 'openpyxl.tests', 'test', 'tests', 'tornado', 'zmq', 'lxml', 'h5py', 'tables'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [('O', None, 'OPTION'), ('O', None, 'OPTION')],
    exclude_binaries=True,
    name='REMFluor-MD',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name='REMFluor-MD',
)
