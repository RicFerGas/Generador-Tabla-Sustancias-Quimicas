# build.spec
# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# Collect all submodules required by spaCy and other libraries
hiddenimports = collect_submodules('spacy') + collect_submodules('srsly') + collect_submodules('thinc') + collect_submodules('preshed') + collect_submodules('catalogue') + collect_submodules('blis') + collect_submodules('pyqt5')

# Collect data files from the spaCy model
datas = [
    (os.path.abspath('./models/xx_ent_wiki_sm'), 'models/xx_ent_wiki_sm'),
    (os.path.abspath('./data_sets'), 'data_sets')
]
a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Generador de Tabla HDS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    bundle_files=1,
    info_plist={
        'CFBundleName': 'Generador de Tabla HDS',
        'CFBundleDisplayName': 'Generador de Tabla HDS',
        'CFBundleIdentifier': 'com.SIGASH.generadorhds',
        'CFBundleVersion': '1.1.0',
        'CFBundleShortVersionString': '1.1.0',
    },
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Generador de Tabla HDS',
)