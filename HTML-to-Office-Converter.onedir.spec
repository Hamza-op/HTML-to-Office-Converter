# -*- mode: python ; coding: utf-8 -*-
#
# Fast-start build:
# - Produces a `dist/HTML-to-Office-Converter/` folder (onedir).
# - Startup is much faster than onefile because nothing is unpacked to a temp dir.
#
# Build:
#   pyinstaller HTML-to-Office-Converter.onedir.spec
#

hidden_imports = [
    # Playwright uses some dynamic imports internally.
    "playwright",
    "playwright.async_api",
    "playwright._impl._driver",
]

# Packages that get auto-pulled but are not needed — saves significant MB
excludes = [
    # Data Science & Math
    "matplotlib", "scipy", "pandas", "numpy.distutils",

    # GUI Frameworks we don't use
    "PyQt5", "PyQt6", "PySide2", "PySide6", "wx",

    # Image/Video Processing we don't use
    "skimage",

    # Misc heavy stuff
    "IPython", "jupyter", "notebook", "tornado", "zmq", "jinja2", "markupsafe",
    "boto3", "botocore", "awscli", "google", "azure",

    # Standard library modules that shouldn't be imported at runtime for this app
    "http.server", "unittest", "pydoc",
    "doctest", "difflib", "optparse", "shelve", "dbm", "sqlite3",
    "test", "_pytest", "pytest",
]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="HTML-to-Office-Converter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="HTML-to-Office-Converter",
)
