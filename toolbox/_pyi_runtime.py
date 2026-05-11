# -*- coding: utf-8 -*-
"""
PyInstaller 运行时钩子：在解压完成后尽早执行（先于业务 import）。
"""
import os
import shutil
import sys


def _app_root():
    return os.path.dirname(os.path.abspath(sys.executable))


def _meipass():
    return getattr(sys, "_MEIPASS", "") or ""


def _seed_dir(rel_toolbox_sub: str):
    """把内置资源复制到 exe 旁，便于用户读写。"""
    if not getattr(sys, "frozen", False):
        return
    me = _meipass()
    if not me:
        return
    src = os.path.join(me, "toolbox", rel_toolbox_sub)
    if not os.path.isdir(src):
        return
    dst = os.path.join(_app_root(), rel_toolbox_sub)
    if os.path.isdir(dst):
        return
    try:
        shutil.copytree(src, dst)
    except Exception:
        pass


def _pyi_main():
    if not getattr(sys, "frozen", False):
        return
    # Playwright 驱动仍走包内；若将来改用内置 Chromium 可指向此目录
    os.environ.setdefault(
        "PLAYWRIGHT_BROWSERS_PATH",
        os.path.join(_app_root(), "ms-playwright"),
    )
    # 尺码映射：首次在 exe 旁生成可编辑副本
    _seed_dir("尺码映射")


_pyi_main()
