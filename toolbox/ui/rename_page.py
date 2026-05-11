# -*- coding: utf-8 -*-
import os
import threading

from PySide6.QtCore import Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from toolbox.ui.path_drop import enable_path_drop


class RenamePage(QWidget):
    # 定义信号用于线程安全地更新UI
    done_signal = Signal(tuple)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.done_signal.connect(self.on_done)
        self.init_ui()

    def on_done(self, data):
        count, output_dir = data
        self.rename_btn.setEnabled(True)
        QMessageBox.information(self, "✅ 重命名完成", f"重命名完成！\n共处理 {count} 个文件\n保存位置: {output_dir}")
        try:
            os.startfile(output_dir)
        except Exception:
            pass

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        self.setLayout(layout)

        title = QLabel("批量重命名")
        title.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        title.setStyleSheet("color: #1565C0;")
        layout.addWidget(title)

        desc = QLabel("按模板批量重命名文件，支持多种预设规则")
        desc.setFont(QFont("Microsoft YaHei", 10))
        desc.setStyleSheet("color: #666;")
        layout.addWidget(desc)

        preset_group = QGroupBox("📋 预设规则")
        preset_group.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        preset_layout = QVBoxLayout()

        preset_info = QLabel("💡 选择一个预设或自定义规则：")
        preset_info.setFont(QFont("Microsoft YaHei", 9))
        preset_info.setStyleSheet("color: #666;")
        preset_layout.addWidget(preset_info)

        self.preset_combo = QComboBox()
        self.preset_combo.addItems(
            [
                "自选编号 (001, 002, 003...)",
                "日期前缀 (20240101_001...)",
                "流水号 (0001, 0002...)",
                "原名保持",
                "自定义模板",
            ]
        )
        self.preset_combo.setFont(QFont("Microsoft YaHei", 10))
        preset_layout.addWidget(self.preset_combo)

        custom_row = QHBoxLayout()
        custom_row.addWidget(QLabel("自定义模板:"))
        self.custom_template = QLineEdit()
        self.custom_template.setText(self.config.get("rename_template", "{n:03d}"))
        self.custom_template.setPlaceholderText("如: {n:03d} 或 prefix_{n:04d}")
        self.custom_template.setStyleSheet(
            """
            QLineEdit {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                padding: 8px 12px;
                background-color: #FAFAFA;
            }
            """
        )
        custom_row.addWidget(self.custom_template)
        preset_layout.addLayout(custom_row)

        template_help = QLabel(
            "📝 可用变量: {n} = 序号, {n:02d} = 2位序号, {n:03d} = 3位序号\n              {name} = 原文件名, {ext} = 扩展名"
        )
        template_help.setFont(QFont("Microsoft YaHei", 8))
        template_help.setStyleSheet("color: #888;")
        preset_layout.addWidget(template_help)

        preset_group.setLayout(preset_layout)
        layout.addWidget(preset_group)

        io_group = QGroupBox("📁 输入输出")
        io_group.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        io_layout = QVBoxLayout()
        io_layout.setSpacing(10)

        input_row = QHBoxLayout()
        input_row.addWidget(QLabel("输入目录:"))
        self.rename_input = QLineEdit()
        enable_path_drop(self.rename_input, mode="dir")
        self.rename_input.setText(self.config.get("rename_dir", ""))
        self.rename_input.setPlaceholderText("选择要重命名的文件所在目录...")
        self.rename_input.setStyleSheet(
            """
            QLineEdit {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                padding: 8px 12px;
                background-color: #FAFAFA;
            }
            """
        )
        input_row.addWidget(self.rename_input)

        btn_input = QPushButton("浏览")
        btn_input.setFixedWidth(80)
        btn_input.setStyleSheet(
            """
            QPushButton {
                background-color: #E3F2FD;
                color: #1565C0;
                border: 1px solid #1565C0;
                padding: 8px 12px;
                border-radius: 4px;
            }
            """
        )
        btn_input.clicked.connect(lambda: self.select_dir(self.rename_input))
        input_row.addWidget(btn_input)
        io_layout.addLayout(input_row)

        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("输出目录:"))
        self.rename_output = QLineEdit()
        enable_path_drop(self.rename_output, mode="dir")
        self.rename_output.setText(self.config.get("rename_output", ""))
        self.rename_output.setPlaceholderText("选择重命名后的保存位置（可同原目录）...")
        self.rename_output.setStyleSheet(
            """
            QLineEdit {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                padding: 8px 12px;
                background-color: #FAFAFA;
            }
            """
        )
        output_row.addWidget(self.rename_output)

        btn_output = QPushButton("浏览")
        btn_output.setFixedWidth(80)
        btn_output.setStyleSheet(
            """
            QPushButton {
                background-color: #E3F2FD;
                color: #1565C0;
                border: 1px solid #1565C0;
                padding: 8px 12px;
                border-radius: 4px;
            }
            """
        )
        btn_output.clicked.connect(lambda: self.select_dir(self.rename_output))
        output_row.addWidget(btn_output)
        io_layout.addLayout(output_row)

        io_group.setLayout(io_layout)
        layout.addWidget(io_group)

        btn_row = QHBoxLayout()
        self.preview_btn = QPushButton("👁️ 预览效果")
        self.preview_btn.setFont(QFont("Microsoft YaHei", 10))
        self.preview_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #9C27B0;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
            }
            """
        )
        btn_row.addWidget(self.preview_btn)

        self.rename_btn = QPushButton("▶️ 开始重命名")
        self.rename_btn.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        self.rename_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #43A047;
                color: white;
                border: none;
                padding: 10px 25px;
                border-radius: 4px;
            }
            """
        )
        self.rename_btn.clicked.connect(self.start_rename)
        btn_row.addWidget(self.rename_btn)

        self.preview_btn.clicked.connect(self.preview_rename)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        preview_group = QGroupBox("👁️ 预览")
        preview_group.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        preview_layout = QVBoxLayout()

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont("Consolas", 9))
        self.preview_text.setMaximumHeight(150)
        preview_layout.addWidget(self.preview_text)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        layout.addStretch()

    def select_dir(self, line_edit):
        path = QFileDialog.getExistingDirectory(self, "选择目录")
        if path:
            line_edit.setText(path)

    def _get_template(self):
        preset = self.preset_combo.currentIndex()
        template = self.custom_template.text().strip()
        if preset == 0:
            template = "{n:03d}"
        elif preset == 1:
            from datetime import datetime

            template = f"{datetime.now().strftime('%Y%m%d')}_{{n:03d}}"
        elif preset == 2:
            template = "{n:04d}"
        elif preset == 3:
            template = "{name}"
        return template

    def _build_name(self, template, fname, counter):
        ext = os.path.splitext(fname)[1]
        new_name = (
            template.replace("{n}", str(counter))
            .replace("{n:02d}", f"{counter:02d}")
            .replace("{n:03d}", f"{counter:03d}")
            .replace("{n:04d}", f"{counter:04d}")
            .replace("{name}", os.path.splitext(fname)[0])
            .replace("{ext}", ext)
        )
        if "." not in new_name:
            new_name += ext
        return new_name

    def start_rename(self):
        import logging
        import shutil

        input_dir = self.rename_input.text().strip()
        output_dir = self.rename_output.text().strip()
        if not input_dir:
            QMessageBox.warning(self, "提示", "请选择输入目录")
            return
        if not output_dir:
            output_dir = input_dir
        try:
            os.makedirs(output_dir, exist_ok=True)
        except OSError as e:
            QMessageBox.warning(self, "错误", f"无法创建输出目录: {e}")
            return
        supported = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
        try:
            files = sorted([f for f in os.listdir(input_dir) if os.path.splitext(f)[1].lower() in supported])
        except OSError as e:
            QMessageBox.warning(self, "错误", f"无法读取输入目录: {e}")
            return
        if not files:
            QMessageBox.warning(self, "提示", "未找到图片文件")
            return

        template = self._get_template()
        self.rename_btn.setEnabled(False)
        work_input = input_dir
        work_output = output_dir
        work_tmpl = template
        work_files = list(files)

        def work():
            log = logging.getLogger("toolbox.rename")
            n_ok = 0
            try:
                counter = 1
                for fname in work_files:
                    f_path = os.path.join(work_input, fname)
                    new_name = self._build_name(work_tmpl, fname, counter)
                    new_path = os.path.join(work_output, new_name)
                    shutil.copy2(f_path, new_path)
                    n_ok = counter
                    counter += 1
            except Exception:
                log.exception("批量重命名")
            finally:
                self.done_signal.emit((n_ok, work_output))

        threading.Thread(target=work, daemon=True).start()

    def preview_rename(self):
        input_dir = self.rename_input.text().strip()
        if not input_dir:
            QMessageBox.warning(self, "提示", "请先选择输入目录")
            return

        supported = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
        files = sorted([f for f in os.listdir(input_dir) if os.path.splitext(f)[1].lower() in supported])
        if not files:
            QMessageBox.warning(self, "提示", "未找到图片文件")
            return

        template = self._get_template()
        preview_text = "📋 预览效果（前5个文件）:\n\n"
        for i, fname in enumerate(files[:5]):
            new_name = self._build_name(template, fname, i + 1)
            preview_text += f"  {fname}\n  → {new_name}\n\n"

        if len(files) > 5:
            preview_text += f"  ... 还有 {len(files) - 5} 个文件"

        QMessageBox.information(self, "👁️ 预览", preview_text)

    def save_settings(self):
        self.config["rename_template"] = self.custom_template.text()
        self.config["rename_dir"] = self.rename_input.text()
        self.config["rename_output"] = self.rename_output.text()
