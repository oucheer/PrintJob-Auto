"""应用运行目录定位（EXE 单文件模式与源码模式通用）。"""
from __future__ import annotations

import sys
from pathlib import Path


def app_base_dir() -> Path:
    """EXE（PyInstaller onefile）返回解包目录；源码运行返回仓库根目录。"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parents[2]
