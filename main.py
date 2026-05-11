#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import traceback
import logging
import threading
import subprocess
import importlib
from logging.handlers import RotatingFileHandler

# 配置日志（打包后写到 exe 同级，开发时写到 toolbox/ 内）
if getattr(sys, "frozen", False):
    LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "toolbox.log")
else:
    LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "toolbox", "toolbox.log")

def setup_logging():
    """日志优化：默认 INFO，单文件 5MB 自动轮转，最多保留 3 份。"""
    is_debug = os.environ.get("TOOLBOX_DEBUG", "").strip() in ("1", "true", "True")
    log_level = logging.DEBUG if is_debug else logging.INFO
    log_format = "%(asctime)s [%(levelname)s] %(message)s"
    formatter = logging.Formatter(log_format)

    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    root_logger.handlers.clear()

    try:
        file_handler = RotatingFileHandler(
            LOG_FILE,
            maxBytes=5 * 1024 * 1024,
            backupCount=3,
            encoding="utf-8",
        )
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)
    except OSError as e:
        # 无写权限或路径异常时仍保证控制台有日志
        file_handler = None
        print(f"toolbox: 无法创建日志文件 {LOG_FILE!r}: {e}", file=sys.stderr)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)

    if file_handler is not None:
        root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)


setup_logging()

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtGui import QFont
from toolbox.ui.main_window import MainWindow


def _module_ok(name: str) -> bool:
    try:
        importlib.import_module(name)
        return True
    except Exception:
        return False


def ensure_runtime_dependencies(parent=None):
    missing = []
    for m in ("rembg", "onnxruntime"):
        if not _module_ok(m):
            missing.append(m)
    if not missing:
        return True

    msg = (
        "检测到当前 Python 环境缺少抠图依赖：\n"
        f"{', '.join(missing)}\n\n"
        "是否现在自动安装到当前程序 Python？\n"
        f"Python: {sys.executable}"
    )
    ret = QMessageBox.question(
        parent,
        "依赖缺失",
        msg,
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.Yes,
    )
    if ret != QMessageBox.Yes:
        return False

    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "rembg", "onnxruntime"]
        )
    except Exception as e:
        QMessageBox.critical(
            parent,
            "安装失败",
            "自动安装失败，请手动执行：\n"
            f"\"{sys.executable}\" -m pip install rembg onnxruntime\n\n"
            f"错误：{e}",
        )
        return False

    QMessageBox.information(parent, "安装完成", "rembg / onnxruntime 已安装完成。")
    return True


def main():
    app = QApplication(sys.argv)
    
    app.setFont(QFont("Microsoft YaHei", 10))
    
    # 全局异常捕获
    def exception_hook(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, Exception):
            error_msg = "".join(
                traceback.format_exception(exc_type, exc_value, exc_traceback)
            )
            logging.error("程序异常: %s", error_msg)
        try:
            if issubclass(exc_type, Exception):
                QMessageBox.critical(
                    None, "程序异常", f"发生错误:\n{exc_value}"
                )
        except Exception:
            pass
    
    sys.excepthook = exception_hook

    def _thread_excepthook(args):
        """子线程未捕获异常也写入日志，避免只崩线程、界面按钮一直灰。"""
        if args.exc_type is None or args.exc_value is None:
            return
        if not issubclass(args.exc_type, Exception):
            return
        import traceback as _tb

        msg = "".join(
            _tb.format_exception(
                args.exc_type, args.exc_value, args.exc_traceback
            )
        )
        logging.error("后台线程异常: %s", msg)

    if hasattr(threading, "excepthook"):
        threading.excepthook = _thread_excepthook

    if not ensure_runtime_dependencies():
        logging.warning("用户取消依赖安装或安装失败，程序退出。")
        return 0

    window = MainWindow()
    window.show()

    logging.info("程序启动成功")
    return app.exec()


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        logging.exception("主程序未捕获异常")
        raise
    else:
        sys.exit(code)
