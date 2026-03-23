# -*- mode: python ; coding: utf-8 -*-

hidden_imports = [
    "playwright",
    "playwright.async_api",
    "playwright._impl._driver",
    "pdf2docx",
    "fitz",
    "pptx",
    "pptx.util",
    "pptx.dml.color",
    "pptx.enum.shapes",
    "customtkinter",
    "PIL",
    "PIL.Image",
    "PIL.ImageTk",
]

# Packages that get auto-pulled but are not needed — saves significant MB
excludes = [
    # Data Science & Math
    "matplotlib", "scipy", "pandas", "numpy.distutils",
    
    # GUI Frameworks we don't use
    "PyQt5", "PyQt6", "PySide2", "PySide6", "wx",
    
    # Image/Video Processing we don't use
    "skimage", 
    
    # Misc Heavy stuff
    "IPython", "jupyter", "notebook", "tornado", "zmq", "jinja2", "markupsafe",
    "boto3", "botocore", "awscli", "google", "azure",
    
    # Standard library stuff that isn't needed for this app
    "http.server", "unittest", "pydoc", 
    "doctest", "difflib", "optparse", "shelve", "dbm", "sqlite3",
    "cryptography", "setuptools", "pkg_resources", "distutils",
    "test", "_pytest", "pytest",
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    # Don't bundle Playwright browser binaries; use system Chrome/Edge (channel=...)
    # or an installed Playwright browser cache.
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=1,           # ← bytecode optimisation (removes docstrings etc.)
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='HTML-to-Office-Converter',
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
