#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import platform
import sys
import importlib


def check_module(name: str):
    try:
        mod = importlib.import_module(name)
        ver = getattr(mod, "__version__", "unknown")
        return True, str(ver), ""
    except Exception as e:
        return False, "", str(e)


def main():
    print("=== 唯品工具箱 环境检测 ===")
    print(f"Python: {sys.version.split()[0]}")
    print(f"Executable: {sys.executable}")
    print(f"Platform: {platform.platform()}")
    print(f"Arch: {platform.architecture()[0]}")
    print("")

    required = ["rembg", "onnxruntime", "PIL", "PySide6"]
    failed = []

    for name in required:
        ok, ver, err = check_module(name)
        if ok:
            print(f"[OK] {name:<12} version={ver}")
        else:
            print(f"[FAIL] {name:<12} error={err}")
            failed.append(name)

    root_dir = os.path.dirname(os.path.abspath(__file__))
    model_path = os.path.join(root_dir, "u2net.onnx")
    if os.path.exists(model_path):
        size_mb = os.path.getsize(model_path) / (1024 * 1024)
        print(f"[OK] u2net.onnx    found ({size_mb:.1f} MB)")
    else:
        print("[FAIL] u2net.onnx    not found (请放到项目根目录)")
        failed.append("u2net.onnx")

    print("")
    if not failed:
        print("结论：环境正常，可直接运行 main.py")
        return 0

    print("结论：存在缺失项。建议执行：")
    print(f"\"{sys.executable}\" -m pip install rembg onnxruntime Pillow PySide6")
    if "u2net.onnx" in failed:
        print("并将 u2net.onnx 放到项目根目录。")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

