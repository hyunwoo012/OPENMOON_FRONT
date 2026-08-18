# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

root = Path(SPECPATH)

datas = [
    (str(root / "frontend" / "dist"), "frontend/dist"),
    (str(root / "backend" / "data" / "templates"), "backend/data/templates"),
    (str(root / "backend" / "data" / "source"), "backend/data/source"),
    (str(root / "backend" / "assets" / "email_signatures"), "backend/assets/email_signatures"),
    (str(root / ".env.example"), "."),
]

a = Analysis(
    ["launcher.py"],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=["uvicorn.logging", "uvicorn.loops.auto", "uvicorn.protocols.http.auto", "uvicorn.protocols.websockets.auto"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="OPENMOON_AI",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="OPENMOON_AI",
)
