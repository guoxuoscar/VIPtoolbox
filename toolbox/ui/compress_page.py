# -*- coding: utf-8 -*-
import os
import threading

from PySide6.QtCore import Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from toolbox.core.utils import clean, compress_to_size, letterbox_square_white
from toolbox.ui.path_drop import enable_path_drop


class CompressPage(QWidget):
    # 定义信号用于线程安全地更新UI
    log_signal = Signal(str)
    done_signal = Signal(tuple)  # 压缩完成信号
    done_signal2 = Signal(tuple)  # 50图完成信号

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self.on_log)
        self.done_signal.connect(self.on_done)
        self.done_signal2.connect(self.on_done_50)
        self.init_ui()

    def on_log(self, msg):
        self.log_message(msg)

    def on_done(self, data):
        processed, total, output_dir = data
        self.compress_btn.setEnabled(True)
        QMessageBox.information(self, "✅ 压缩完成", f"图片压缩完成！\n总计:{total} 处理:{processed}\n保存位置: {output_dir}")
        try:
            os.startfile(output_dir)
        except Exception:
            pass

    def on_done_50(self, data):
        processed, total, output_dir = data
        self.gen50_btn.setEnabled(True)
        QMessageBox.information(self, "✅ 50图完成", f"50图生成完成！\n总计:{total} 处理:{processed}\n保存位置: {output_dir}")
        try:
            os.startfile(output_dir)
        except Exception:
            pass

    def select_dir(self, line_edit):
        """选择目录（供浏览按钮调用）"""
        path = QFileDialog.getExistingDirectory(self, "选择目录")
        if path:
            line_edit.setText(path)

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)
        self.setStyleSheet("""
            QGroupBox {
                border: 1px solid #D6E4F0;
                border-radius: 8px;
                margin-top: 10px;
                background: #FFFFFF;
                font-family: Microsoft YaHei;
                font-size: 11px;
                font-weight: bold;
                color: #1565C0;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
            }
            QLabel {
                color: #333333;
                font-family: Microsoft YaHei;
                font-size: 10px;
            }
            QLineEdit, QSpinBox, QComboBox {
                border: 1px solid #CFD8DC;
                border-radius: 6px;
                padding: 6px 8px;
                background: #FAFAFA;
                min-height: 26px;
                font-family: Microsoft YaHei;
                font-size: 10px;
            }
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
                border: 1px solid #1565C0;
                background: #FFFFFF;
            }
            QCheckBox {
                color: #37474F;
                font-family: Microsoft YaHei;
                font-size: 10px;
            }
            QTextEdit {
                border: 1px solid #CFD8DC;
                border-radius: 6px;
                background: #FAFAFA;
                font-family: Consolas, Microsoft YaHei;
                font-size: 9px;
            }
            QPushButton {
                background-color: #E3F2FD;
                color: #1565C0;
                border: 1px solid #90CAF9;
                border-radius: 6px;
                padding: 7px 12px;
                font-family: Microsoft YaHei;
                font-size: 10px;
            }
            QPushButton:hover {
                background-color: #BBDEFB;
            }
        """)
        title = QLabel("图片压缩工具")
        title.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        title.setStyleSheet("color: #1565C0;")
        layout.addWidget(title)
        desc = QLabel(
            "将图片批量压缩到指定宽度，保持比例不变；可选白底 1:1；"
            "默认在「输入文件夹」和「下面第 1 层子文件夹」里找图（层数可调）。"
            "体积上限内会尽量用满：逐档画质 + 更占体积的色度采样，仍远小于上限多半是图本身简单（JPEG 有上限）。"
        )
        desc.setWordWrap(True)
        desc.setFont(QFont("Microsoft YaHei", 10))
        desc.setStyleSheet("color: #666;")
        layout.addWidget(desc)
        compress_group = QGroupBox("🖼️ 图片批量压缩")
        compress_group.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        compress_layout = QVBoxLayout()
        compress_layout.setSpacing(10)
        params_row = QHBoxLayout()
        params_row.addWidget(QLabel("目标宽度:"))
        self.compress_width = QSpinBox()
        self.compress_width.setRange(100, 3000)
        self.compress_width.setValue(self.config.get("compress_size", 1200))
        self.compress_width.setSuffix(" px")
        self.compress_width.setFixedWidth(120)
        params_row.addWidget(self.compress_width)
        params_row.addWidget(QLabel("  最大体积:"))
        self.compress_maxkb = QSpinBox()
        self.compress_maxkb.setRange(10, 5000)
        self.compress_maxkb.setValue(self.config.get("compress_maxkb", 1024))
        self.compress_maxkb.setSuffix(" KB")
        self.compress_maxkb.setFixedWidth(120)
        params_row.addWidget(self.compress_maxkb)
        params_row.addStretch()
        compress_layout.addLayout(params_row)
        opt_row = QHBoxLayout()
        self.compress_square_cb = QCheckBox("非 1:1 图转 1:1（白底居中，边留白，边长=目标宽度）")
        self.compress_square_cb.setChecked(self.config.get("compress_square_11", False))
        opt_row.addWidget(self.compress_square_cb)
        opt_row.addWidget(QLabel("  扫描子文件夹层数:"))
        self.compress_depth = QSpinBox()
        self.compress_depth.setRange(1, 8)
        self.compress_depth.setValue(int(self.config.get("compress_scan_depth", 2)))
        self.compress_depth.setFixedWidth(60)
        opt_row.addWidget(self.compress_depth)
        opt_row.addStretch()
        compress_layout.addLayout(opt_row)
        _nm = str(self.config.get("compress_naming_mode", "original") or "original")
        if _nm not in ("original", "detail", "main", "custom"):
            _nm = "original"
        naming_row = QHBoxLayout()
        naming_row.addWidget(QLabel("压缩后命名:"))
        self.compress_naming_mode = QComboBox()
        self.compress_naming_mode.addItem("原文件名", "original")
        self.compress_naming_mode.addItem("详情图式（601、602、603…）", "detail")
        self.compress_naming_mode.addItem("主图式（1、2、3、4 后 15、16、17…）", "main")
        self.compress_naming_mode.addItem("自定义", "custom")
        for k in range(self.compress_naming_mode.count()):
            if self.compress_naming_mode.itemData(k) == _nm:
                self.compress_naming_mode.setCurrentIndex(k)
                break
        naming_row.addWidget(self.compress_naming_mode)
        self.compress_naming_custom = QLineEdit()
        self.compress_naming_custom.setText((self.config.get("compress_naming_custom") or "{i}.jpg").strip() or "{i}.jpg")
        self.compress_naming_mode.currentIndexChanged.connect(self._on_compress_naming_mode_changed)
        self._on_compress_naming_mode_changed()
        naming_row.addWidget(self.compress_naming_custom, 1)
        compress_layout.addLayout(naming_row)
        input_row = QHBoxLayout()
        input_row.addWidget(QLabel("输入来源:"))
        self.compress_input = QLineEdit()
        enable_path_drop(self.compress_input, mode="file_or_dir", extensions=(".jpg", ".jpeg", ".png", ".bmp", ".webp"), multi=True)
        self.compress_input.setText(self.config.get("compress_dir", ""))
        input_row.addWidget(self.compress_input)
        btn_input = QPushButton("浏览")
        btn_input.clicked.connect(lambda: self.select_input(self.compress_input))
        input_row.addWidget(btn_input)
        compress_layout.addLayout(input_row)
        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("输出目录:"))
        self.compress_output = QLineEdit()
        enable_path_drop(self.compress_output, mode="dir")
        self.compress_output.setText(self.config.get("compress_output", ""))
        output_row.addWidget(self.compress_output)
        btn_output = QPushButton("浏览")
        btn_output.clicked.connect(lambda: self.select_dir(self.compress_output))
        output_row.addWidget(btn_output)
        compress_layout.addLayout(output_row)
        self.compress_btn = QPushButton("▶️ 开始压缩")
        self.compress_btn.setStyleSheet("""
            QPushButton {
                background-color: #1565C0;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 9px 16px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #B0BEC5;
                color: #ECEFF1;
            }
        """)
        self.compress_btn.clicked.connect(self.start_compress)
        compress_layout.addWidget(self.compress_btn)
        compress_group.setLayout(compress_layout)
        layout.addWidget(compress_group)

        gen50_group = QGroupBox("📷 批量生成50图")
        gen50_layout = QVBoxLayout()
        gen50_params = QHBoxLayout()
        gen50_params.addWidget(QLabel("最大体积:"))
        self.gen50_maxkb = QSpinBox()
        self.gen50_maxkb.setRange(10, 2000)
        self.gen50_maxkb.setValue(self.config.get("gen50_maxkb", 1024))
        self.gen50_maxkb.setSuffix(" KB")
        gen50_params.addWidget(self.gen50_maxkb)
        gen50_params.addWidget(QLabel("  扫描子文件夹层数:"))
        _g50d = int(self.config.get("gen50_scan_depth", self.config.get("compress_scan_depth", 2)))
        _g50d = max(1, min(8, _g50d))
        self.gen50_scan_depth = QSpinBox()
        self.gen50_scan_depth.setRange(1, 8)
        self.gen50_scan_depth.setValue(_g50d)
        gen50_params.addWidget(self.gen50_scan_depth)
        gen50_params.addStretch()
        gen50_layout.addLayout(gen50_params)
        self.gen50_name_as_50 = QCheckBox("输出统一命名为 50.jpg（多图时自动 50_2、50_3… 防覆盖；不选则与源图同名）")
        self.gen50_name_as_50.setChecked(bool(self.config.get("gen50_name_as_50", False)))
        gen50_layout.addWidget(self.gen50_name_as_50)
        gen50_input_row = QHBoxLayout()
        gen50_input_row.addWidget(QLabel("输入目录:"))
        self.gen50_input = QLineEdit()
        enable_path_drop(self.gen50_input, mode="dir")
        self.gen50_input.setText(self.config.get("gen50_dir", ""))
        gen50_input_row.addWidget(self.gen50_input)
        btn_gen50_input = QPushButton("浏览")
        btn_gen50_input.clicked.connect(lambda: self.select_dir(self.gen50_input))
        gen50_input_row.addWidget(btn_gen50_input)
        gen50_layout.addLayout(gen50_input_row)
        gen50_output_row = QHBoxLayout()
        gen50_output_row.addWidget(QLabel("输出目录:"))
        self.gen50_output = QLineEdit()
        enable_path_drop(self.gen50_output, mode="dir")
        self.gen50_output.setText(self.config.get("gen50_output", ""))
        gen50_output_row.addWidget(self.gen50_output)
        btn_gen50_output = QPushButton("浏览")
        btn_gen50_output.clicked.connect(lambda: self.select_dir(self.gen50_output))
        gen50_output_row.addWidget(btn_gen50_output)
        gen50_layout.addLayout(gen50_output_row)
        self.gen50_btn = QPushButton("📷 开始生成50图")
        self.gen50_btn.setStyleSheet("""
            QPushButton {
                background-color: #00796B;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 9px 16px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #00897B;
            }
            QPushButton:disabled {
                background-color: #B0BEC5;
                color: #ECEFF1;
            }
        """)
        self.gen50_btn.clicked.connect(self.start_gen50)
        gen50_layout.addWidget(self.gen50_btn)
        gen50_group.setLayout(gen50_layout)
        layout.addWidget(gen50_group)

        log_group = QGroupBox("📝 操作日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(100)
        self.log_text.setMaximumHeight(240)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group, 1)
        inner = QWidget()
        inner.setLayout(layout)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)
        scroll.setFrameShape(QFrame.NoFrame)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(scroll)

    def select_input(self, line_edit):
        msg = QMessageBox(self)
        msg.setWindowTitle("选择输入方式")
        msg.setText("请选择输入类型：")
        msg.setIcon(QMessageBox.Question)
        folder_btn = msg.addButton("📁 文件夹（批量处理）", QMessageBox.ActionRole)
        file_btn = msg.addButton("📄 图片文件（多选）", QMessageBox.ActionRole)
        msg.addButton("取消", QMessageBox.RejectRole)
        msg.exec()
        if msg.clickedButton() == folder_btn:
            path = QFileDialog.getExistingDirectory(self, "选择图片文件夹")
            if path:
                line_edit.setText(path)
        elif msg.clickedButton() == file_btn:
            files, _ = QFileDialog.getOpenFileNames(self, "选择图片文件", "", "图片文件 (*.jpg *.jpeg *.png *.bmp *.webp)")
            if files:
                line_edit.setText(";".join(files))

    def log_message(self, msg):
        self.log_text.append(msg)

    def _on_compress_naming_mode_changed(self, _=None):
        if not hasattr(self, "compress_naming_mode") or not hasattr(self, "compress_naming_custom"):
            return
        self.compress_naming_custom.setVisible(self.compress_naming_mode.currentData() == "custom")

    @staticmethod
    def _next_compress_output_name(naming, custom_tmpl, out_sub, folder_index, src_path):
        out_key = os.path.normpath(os.path.abspath(out_sub))
        base_stem = clean(os.path.splitext(os.path.basename(src_path))[0]) or "image"
        if naming == "original":
            return f"{base_stem}.jpg"
        k = folder_index[out_key]
        folder_index[out_key] = k + 1
        if naming == "detail":
            return f"{601 + k}.jpg"
        if naming == "main":
            n = (k + 1) if k < 4 else 15 + (k - 4)
            return f"{n}.jpg"
        i = k + 1
        tmpl = (custom_tmpl or "{i}.jpg").strip() or "{i}.jpg"
        try:
            out = tmpl.format(i=i, idx=i, stem=base_stem)
        except (KeyError, ValueError, IndexError):
            out = f"{i}.jpg"
        b, e = os.path.splitext(out or f"{i}.jpg")
        if e.lower() not in (".jpg", ".jpeg"):
            e = ".jpg"
        return f"{clean(b) or str(i)}{e}"

    @staticmethod
    def _collect_compress_files(root, max_depth, supported):
        root = os.path.abspath(os.path.normpath(root))
        tasks = []
        if max_depth < 1:
            max_depth = 1
        try:
            for name in sorted(os.listdir(root)):
                p = os.path.join(root, name)
                if os.path.isfile(p) and os.path.splitext(name)[1].lower() in supported:
                    rd = os.path.dirname(os.path.relpath(p, root))
                    tasks.append((p, "." if rd in (".", "") else rd))
        except OSError:
            pass
        if max_depth <= 1:
            return tasks

        def scan_dir(dpath, level):
            if level > max_depth:
                return
            try:
                for name in sorted(os.listdir(dpath)):
                    p = os.path.join(dpath, name)
                    if os.path.isfile(p) and os.path.splitext(name)[1].lower() in supported:
                        rd = os.path.dirname(os.path.relpath(p, root))
                        tasks.append((p, "." if rd in (".", "") else rd))
                    elif os.path.isdir(p):
                        scan_dir(p, level + 1)
            except OSError:
                pass

        try:
            for name in sorted(os.listdir(root)):
                sub = os.path.join(root, name)
                if os.path.isdir(sub):
                    scan_dir(sub, 2)
        except OSError:
            pass
        return tasks

    def start_compress(self):
        input_dir = self.compress_input.text().strip()
        output_dir = self.compress_output.text().strip()
        if not input_dir or not output_dir:
            QMessageBox.warning(self, "提示", "请选择输入来源和输出目录")
            return
        naming = self.compress_naming_mode.currentData() or "original"
        custom_tmpl = (self.compress_naming_custom.text() or "{i}.jpg").strip() or "{i}.jpg"
        target_width = int(self.compress_width.value())
        max_kb = int(self.compress_maxkb.value())
        max_depth = int(self.compress_depth.value())
        square = bool(self.compress_square_cb.isChecked())
        self.compress_btn.setEnabled(False)

        def work():
            import logging
            from collections import defaultdict
            from PIL import Image
            processed, total = 0, 0
            try:
                os.makedirs(output_dir, exist_ok=True)
                supported = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
                raw = input_dir.strip()
                tasks = []
                if ";" in raw:
                    for p in raw.split(";"):
                        p = p.strip()
                        if os.path.isfile(p) and os.path.splitext(p)[1].lower() in supported:
                            tasks.append((p, "."))
                elif os.path.isfile(raw):
                    if os.path.splitext(raw)[1].lower() in supported:
                        tasks.append((raw, "."))
                elif os.path.isdir(raw):
                    tasks = CompressPage._collect_compress_files(raw, max_depth, supported)
                if not tasks:
                    self.log_signal.emit("没有找到符合条件的图片（请检查文件夹层数或扩展名）")
                    return
                total = len(tasks)
                folder_index = defaultdict(int)
                for src, rel in tasks:
                    try:
                        rel_dir = "." if rel in (".", "") else rel
                        out_sub = output_dir if rel_dir == "." else os.path.join(output_dir, rel_dir)
                        os.makedirs(out_sub, exist_ok=True)
                        out_name = self._next_compress_output_name(str(naming), custom_tmpl, out_sub, folder_index, src)
                        out_path = os.path.join(out_sub, out_name)
                        with Image.open(src) as img:
                            img.load()
                            ratio = target_width / float(img.width)
                            new_height = max(1, int(img.height * ratio))
                            img = img.resize((target_width, new_height), Image.Resampling.LANCZOS)
                            if square:
                                img = letterbox_square_white(img, target_width)
                        compress_to_size(img, out_path, max_kb=max_kb)
                        processed += 1
                    except Exception as e:
                        self.log_signal.emit(f"  错误: {src} - {e}")
            except Exception as e:
                logging.getLogger("toolbox.compress").exception("批量压缩任务")
                self.log_signal.emit(f"压缩任务异常（已记日志）: {e}")
            finally:
                self.done_signal.emit((processed, total, output_dir))

        threading.Thread(target=work, daemon=True).start()

    def start_gen50(self):
        input_dir = self.gen50_input.text().strip()
        output_dir = self.gen50_output.text().strip()
        if not input_dir or not output_dir:
            QMessageBox.warning(self, "提示", "请选择输入和输出目录")
            return
        if not os.path.isdir(input_dir):
            QMessageBox.warning(self, "提示", "输入不是有效的图片文件夹")
            return
        gen50_as_50 = self.gen50_name_as_50.isChecked()
        gen50_maxkb_v = int(self.gen50_maxkb.value())
        gen50_depth_v = int(self.gen50_scan_depth.value())
        self.gen50_btn.setEnabled(False)

        def work():
            import logging
            input_folder_name = os.path.basename(input_dir.rstrip("/\\"))
            output_dir_final = os.path.join(output_dir, f"{input_folder_name}_1200")
            processed, total = 0, 0
            try:
                os.makedirs(output_dir_final, exist_ok=True)
                target_w, target_h = 950, 1200
                as_50 = bool(gen50_as_50)
                supported = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
                tasks = CompressPage._collect_compress_files(input_dir, max(1, min(8, int(gen50_depth_v))), supported)
                all_paths = [p for p, _ in tasks]
                n_total = len(all_paths)
                total = n_total
                for f_path in all_paths:
                    fname = os.path.basename(f_path)
                    try:
                        from PIL import Image
                        with Image.open(f_path) as img:
                            img.load()
                            orig_w, orig_h = img.size
                            scale_h = target_h / orig_h
                            scaled_w = int(orig_w * scale_h)
                            if scaled_w >= target_w:
                                img_scaled = img.resize((scaled_w, target_h), Image.Resampling.LANCZOS)
                                left = (scaled_w - target_w) // 2
                                img_cropped = img_scaled.crop((left, 0, left + target_w, target_h))
                            else:
                                scale_w = target_w / orig_w
                                new_h = int(orig_h * scale_w)
                                img_scaled = img.resize((target_w, new_h), Image.Resampling.LANCZOS)
                                top = (new_h - target_h) // 2
                                img_cropped = img_scaled.crop((0, top, target_w, top + target_h))
                            if as_50:
                                out_basename = "50.jpg" if n_total == 1 or processed == 0 else f"50_{processed + 1}.jpg"
                            else:
                                out_basename = f"{clean(os.path.splitext(fname)[0]) or 'out'}.jpg"
                            out_path = os.path.join(output_dir_final, out_basename)
                            compress_to_size(img_cropped, out_path, max_kb=gen50_maxkb_v)
                            processed += 1
                    except Exception as e:
                        self.log_signal.emit(f"  错误: {fname} - {e}")
            except Exception as e:
                logging.getLogger("toolbox.gen50").exception("50图任务")
                self.log_signal.emit(f"50图任务异常（已记日志）: {e}")
            finally:
                self.done_signal2.emit((processed, total, output_dir_final))

        threading.Thread(target=work, daemon=True).start()

    def save_settings(self):
        self.config["compress_size"] = self.compress_width.value()
        self.config["compress_maxkb"] = self.compress_maxkb.value()
        self.config["compress_dir"] = self.compress_input.text()
        self.config["compress_output"] = self.compress_output.text()
        self.config["compress_square_11"] = self.compress_square_cb.isChecked()
        self.config["compress_scan_depth"] = int(self.compress_depth.value())
        self.config["compress_naming_mode"] = self.compress_naming_mode.currentData() or "original"
        self.config["compress_naming_custom"] = self.compress_naming_custom.text().strip()
        self.config["gen50_maxkb"] = self.gen50_maxkb.value()
        self.config["gen50_scan_depth"] = int(self.gen50_scan_depth.value())
        self.config["gen50_dir"] = self.gen50_input.text()
        self.config["gen50_output"] = self.gen50_output.text()
        self.config["gen50_name_as_50"] = self.gen50_name_as_50.isChecked()
