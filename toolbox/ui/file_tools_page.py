# -*- coding: utf-8 -*-
import os
import re
import shutil
import threading

import openpyxl

from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from toolbox.ui.path_drop import ExcelDropLineEdit, DirDropLineEdit, enable_path_drop


# ── Tab 1: 批量新建文件夹 ──────────────────────────────────────────
class FolderCreatorTab(QWidget):
    log_signal = Signal(str)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self._append_log)
        self.init_ui()

    def _append_log(self, msg):
        self.log_area.append(msg)
        bar = self.log_area.verticalScrollBar()
        if bar:
            bar.setValue(bar.maximum())

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        desc = QLabel(
            "根据Excel表格中的列批量创建文件夹。选一列 = 单层文件夹，选多列 = 多层嵌套文件夹（按选择顺序作为层级）。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc)

        # Excel 选择
        excel_row = QHBoxLayout()
        excel_row.addWidget(QLabel("Excel表格:"))
        self.excel_input = ExcelDropLineEdit("拖拽或浏览选择Excel文件...")
        self.excel_input.setText(self.config.get("folder_creator_excel", ""))
        excel_row.addWidget(self.excel_input, 1)
        excel_btn = QPushButton("浏览")
        excel_btn.clicked.connect(self._browse_excel)
        excel_row.addWidget(excel_btn)
        layout.addLayout(excel_row)

        # 列选择区域
        col_group = QGroupBox("📋 选择列（多选=多层嵌套，顺序=层级顺序）")
        col_group.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        col_layout = QVBoxLayout()
        col_layout.setSpacing(6)

        col_top = QHBoxLayout()
        self.cb_vip_style = QCheckBox("唯品款号（默认）")
        self.cb_vip_style.setChecked(True)
        self.cb_vip_style.toggled.connect(self._on_col_check_changed)
        self.cb_vip_goods = QCheckBox("唯品货号")
        self.cb_vip_goods.setChecked(False)
        self.cb_vip_goods.toggled.connect(self._on_col_check_changed)
        self.cb_custom = QCheckBox("自定义列名:")
        self.cb_custom.toggled.connect(self._on_col_check_changed)
        self.custom_col = QLineEdit()
        self.custom_col.setPlaceholderText("输入列名（如: 款号、分类名称）")
        self.custom_col.setEnabled(False)
        col_top.addWidget(self.cb_vip_style)
        col_top.addWidget(self.cb_vip_goods)
        col_top.addWidget(self.cb_custom)
        col_top.addWidget(self.custom_col, 1)
        col_layout.addLayout(col_top)

        col_layout.addWidget(QLabel("💡 提示: 将按勾选顺序从上到下作为文件夹层级。例如: 唯品货号 → 唯品款号 = 货号文件夹/款号子文件夹"))
        col_group.setLayout(col_layout)
        layout.addWidget(col_group)

        # 输出目录
        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("输出目录:"))
        self.output_dir = DirDropLineEdit("默认为表格所在目录，可自选")
        self.output_dir.setText(self.config.get("folder_creator_output", ""))
        out_row.addWidget(self.output_dir, 1)
        out_btn = QPushButton("浏览")
        out_btn.clicked.connect(self._browse_output)
        out_row.addWidget(out_btn)
        layout.addLayout(out_row)

        # 运行按钮
        self.run_btn = QPushButton("▶️ 开始批量建文件夹")
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #1565C0; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #1976D2; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.run_btn.clicked.connect(self._start)
        layout.addWidget(self.run_btn)

        # 日志
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        self.log_area.setStyleSheet("font-family: Consolas; font-size: 9px; background: #FAFAFA;")
        layout.addWidget(self.log_area)

        layout.addStretch()
        self.setLayout(layout)

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel", "", "Excel (*.xlsx *.xls)")
        if path:
            self.excel_input.setText(path)

    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.output_dir.setText(path)

    def _on_col_check_changed(self):
        self.custom_col.setEnabled(self.cb_custom.isChecked())

    def _get_column_names(self):
        """按勾选顺序返回列名列表"""
        cols = []
        if self.cb_vip_style.isChecked():
            cols.append("唯品款号")
        if self.cb_vip_goods.isChecked():
            cols.append("唯品货号")
        if self.cb_custom.isChecked():
            txt = self.custom_col.text().strip()
            if txt:
                cols.append(txt)
        return cols

    def _start(self):
        excel_path = self.excel_input.text().strip()
        if not excel_path or not os.path.isfile(excel_path):
            QMessageBox.warning(self, "提示", "请先选择Excel文件")
            return

        cols = self._get_column_names()
        if not cols:
            QMessageBox.warning(self, "提示", "请至少选择一列")
            return

        output_dir = self.output_dir.text().strip()
        if not output_dir:
            output_dir = os.path.dirname(excel_path)
            self.output_dir.setText(output_dir)

        # 保存配置
        self.config["folder_creator_excel"] = excel_path
        self.config["folder_creator_output"] = output_dir

        self.run_btn.setEnabled(False)
        self.log_area.clear()
        threading.Thread(target=self._do_create, args=(excel_path, cols, output_dir), daemon=True).start()

    def _do_create(self, excel_path, cols, output_dir):
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True)
            ws = wb.active
            headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
            header_map = {}
            for idx, h in enumerate(headers):
                if h:
                    header_map[str(h).strip()] = idx

            # 找到各列对应的列索引
            col_indices = []
            for col_name in cols:
                found = None
                for h, idx in header_map.items():
                    if col_name == h or col_name in h:
                        found = idx
                        break
                if found is None:
                    self.log_signal.emit(f"[错误] 未找到列: {col_name}")
                    return
                col_indices.append((col_name, found))

            self.log_signal.emit(f"找到列: {[c[0] for c in col_indices]}")
            self.log_signal.emit(f"输出目录: {output_dir}")
            self.log_signal.emit("")

            created = 0
            skipped = 0
            seen = set()

            for row_data in ws.iter_rows(min_row=2, values_only=True):
                parts = []
                all_valid = True
                for col_name, col_idx in col_indices:
                    val = row_data[col_idx] if col_idx < len(row_data) else None
                    if val is None or str(val).strip() == "":
                        all_valid = False
                        break
                    # 清理文件夹名非法字符
                    folder_name = str(val).strip()
                    folder_name = re.sub(r'[<>:"/\\|?*]', '_', folder_name)
                    parts.append(folder_name)

                if not all_valid or not parts:
                    skipped += 1
                    continue

                target = os.path.join(output_dir, *parts)
                if target in seen:
                    skipped += 1
                    continue
                seen.add(target)

                if not os.path.exists(target):
                    os.makedirs(target, exist_ok=True)
                    self.log_signal.emit(f"[创建] {os.path.relpath(target, output_dir)}")
                    created += 1
                else:
                    skipped += 1

            wb.close()
            self.log_signal.emit("")
            self.log_signal.emit(f"创建 {created} 个文件夹，跳过 {skipped} 个（已存在或无数据）")
            self.log_signal.emit("=" * 50)
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
        finally:
            self.run_btn.setEnabled(True)


# ── Tab 2: 提取文件汇总（按文件夹名匹配） ──────────────────────
class FileCollectExtractTab(QWidget):
    """按文件夹名匹配目标文件夹，提取其中文件，支持多种输出结构和命名规则。"""
    log_signal = Signal(str)

    FOLDER_PRESETS = ["SKU", "详情页", "详情图片", "主图", "白底图", "透明图", "视频", "素材图"]
    OUTPUT_MODES = [
        ("按父文件夹归类", "parent", "如: GZ001/front.jpg"),
        ("按匹配文件夹归类", "matched", "如: SKU/GZ001_front.jpg"),
        ("父文件夹/匹配文件夹 双层", "both", "如: GZ001/SKU/front.jpg"),
        ("扁平汇总(全部放一起)", "flat", "如: GZ001_SKU_front.jpg"),
    ]
    NAMING_MODES = [
        ("保持原文件名", "original", "{name}{ext}"),
        ("父文件夹_原文件名", "parent_name", "{parent}_{name}{ext}"),
        ("父文件夹_序号", "parent_num", "{parent}_{num:04d}{ext}"),
        ("匹配文件夹_父文件夹_原文件名", "matched_parent_name", "{matched}_{parent}_{name}{ext}"),
        ("自定义模板", "custom", ""),
    ]

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self._append_log)
        self.init_ui()

    def _append_log(self, msg):
        self.log_area.append(msg)
        bar = self.log_area.verticalScrollBar()
        if bar:
            bar.setValue(bar.maximum())

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        desc = QLabel(
            "按文件夹名匹配目标文件夹（如SKU、详情图片），提取其中文件，支持多种输出结构和命名规则。\n"
            "例如: 勾选 SKU、详情页 → 自动找到所有名为SKU和详情页的文件夹 → 提取内部文件 → 按规则命名汇总。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc)

        # 源文件夹
        src_row = QHBoxLayout()
        src_row.addWidget(QLabel("源文件夹:"))
        self.src_dir = DirDropLineEdit("拖拽或浏览选择根目录...")
        self.src_dir.setText(self.config.get("file_collect_src", ""))
        src_row.addWidget(self.src_dir, 1)
        src_btn = QPushButton("浏览")
        src_btn.clicked.connect(self._browse_src)
        src_row.addWidget(src_btn)
        layout.addLayout(src_row)

        # 匹配文件夹名
        folder_group = QGroupBox("📁 匹配文件夹名 (勾选需要提取的文件夹)")
        folder_group.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        folder_layout = QVBoxLayout()
        folder_layout.setSpacing(4)

        self.folder_cbs = {}
        row = QHBoxLayout()
        row.setSpacing(12)
        for i, name in enumerate(self.FOLDER_PRESETS):
            cb = QCheckBox(name)
            cb.setChecked(name in str(self.config.get("file_collect_folders", "SKU")).split(","))
            cb.toggled.connect(self._on_folder_toggled)
            self.folder_cbs[name] = cb
            row.addWidget(cb)
            if (i + 1) % 4 == 0 and i < len(self.FOLDER_PRESETS) - 1:
                folder_layout.addLayout(row)
                row = QHBoxLayout()
                row.setSpacing(12)
        folder_layout.addLayout(row)

        custom_row = QHBoxLayout()
        custom_row.addWidget(QLabel("自定义:"))
        self.custom_folder = QLineEdit()
        self.custom_folder.setPlaceholderText("空格分隔，如: 缩略图 角度图")
        self.custom_folder.setText(self.config.get("file_collect_custom_folders", ""))
        custom_row.addWidget(self.custom_folder, 1)
        folder_layout.addLayout(custom_row)

        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)

        # 文件筛选
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("文件类型:"))
        self.ext_input = QLineEdit()
        self.ext_input.setPlaceholderText("空格分隔，如: .jpg .png（留空=所有文件）")
        self.ext_input.setText(self.config.get("file_collect_ext", ""))
        filter_row.addWidget(self.ext_input, 1)
        filter_row.addWidget(QLabel("  文件名含:"))
        self.kw_input = QLineEdit()
        self.kw_input.setPlaceholderText("如: front（留空=不筛选）")
        self.kw_input.setText(self.config.get("file_collect_kw", ""))
        filter_row.addWidget(self.kw_input, 1)
        layout.addLayout(filter_row)

        # 输出结构
        out_group = QGroupBox("📂 输出结构 & ✏️ 命名规则")
        out_group.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        out_layout = QVBoxLayout()
        out_layout.setSpacing(8)

        # 输出结构
        struct_row = QHBoxLayout()
        struct_row.addWidget(QLabel("输出结构:"))
        self.struct_combo = QComboBox()
        for label, val, hint in self.OUTPUT_MODES:
            self.struct_combo.addItem(f"{label} ({hint})", val)
        prev_struct = self.config.get("file_collect_struct", "parent")
        for k in range(self.struct_combo.count()):
            if self.struct_combo.itemData(k) == prev_struct:
                self.struct_combo.setCurrentIndex(k)
                break
        struct_row.addWidget(self.struct_combo, 1)
        out_layout.addLayout(struct_row)

        # 命名规则
        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("命名规则:"))
        self.name_combo = QComboBox()
        for label, val, hint in self.NAMING_MODES:
            text = f"{label} ({hint})" if hint else label
            self.name_combo.addItem(text, val)
        prev_name = self.config.get("file_collect_naming", "original")
        for k in range(self.name_combo.count()):
            if self.name_combo.itemData(k) == prev_name:
                self.name_combo.setCurrentIndex(k)
                break
        self.name_combo.currentIndexChanged.connect(self._on_naming_changed)
        name_row.addWidget(self.name_combo, 1)
        out_layout.addLayout(name_row)

        # 自定义模板
        templ_row = QHBoxLayout()
        templ_row.addWidget(QLabel("自定义模板:"))
        self.custom_template = QLineEdit()
        self.custom_template.setPlaceholderText("变量: {parent} {matched} {name} {ext} {num}")
        self.custom_template.setText(self.config.get("file_collect_template", "{parent}_{matched}_{name}{ext}"))
        templ_row.addWidget(self.custom_template, 1)
        out_layout.addLayout(templ_row)

        hint = QLabel("💡 {parent}=父文件夹名  {matched}=匹配文件夹名  {name}=原文件名(无扩展名)  {ext}=扩展名  {num}=自动序号")
        hint.setStyleSheet("color: #999; font-size: 9px;")
        out_layout.addWidget(hint)
        out_group.setLayout(out_layout)

        self._on_naming_changed()
        layout.addWidget(out_group)

        # 输出目录
        outdir_row = QHBoxLayout()
        outdir_row.addWidget(QLabel("输出目录:"))
        self.output_dir = DirDropLineEdit("汇总输出位置...")
        self.output_dir.setText(self.config.get("file_collect_output", ""))
        outdir_row.addWidget(self.output_dir, 1)
        out_btn = QPushButton("浏览")
        out_btn.clicked.connect(self._browse_output)
        outdir_row.addWidget(out_btn)
        layout.addLayout(outdir_row)

        self.run_btn = QPushButton("▶️ 开始提取汇总")
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #00796B; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #00897B; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.run_btn.clicked.connect(self._start)
        layout.addWidget(self.run_btn)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        self.log_area.setStyleSheet("font-family: Consolas; font-size: 9px; background: #FAFAFA;")
        layout.addWidget(self.log_area)
        layout.addStretch()
        self.setLayout(layout)

    def _browse_src(self):
        path = QFileDialog.getExistingDirectory(self, "选择源文件夹")
        if path:
            self.src_dir.setText(path)

    def _browse_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.output_dir.setText(path)

    def _on_folder_toggled(self):
        pass

    def _on_naming_changed(self):
        is_custom = self.name_combo.currentData() == "custom"
        self.custom_template.setEnabled(is_custom)

    def _get_folder_names(self):
        names = []
        for name, cb in self.folder_cbs.items():
            if cb.isChecked():
                names.append(name)
        custom = self.custom_folder.text().strip()
        if custom:
            for n in custom.split():
                n = n.strip()
                if n and n not in names:
                    names.append(n)
        return names

    def _start(self):
        src = self.src_dir.text().strip()
        if not src or not os.path.isdir(src):
            QMessageBox.warning(self, "提示", "请先选择源文件夹")
            return

        folder_names = self._get_folder_names()
        if not folder_names:
            QMessageBox.warning(self, "提示", "请至少选择一个目标文件夹名")
            return

        output = self.output_dir.text().strip()
        if not output:
            QMessageBox.warning(self, "提示", "请选择输出目录")
            return

        exts = [e.strip().lower() for e in self.ext_input.text().strip().split() if e.strip()]
        kw = self.kw_input.text().strip()
        struct = self.struct_combo.currentData()
        naming = self.name_combo.currentData()
        template = self.custom_template.text().strip() if naming == "custom" else ""

        # 保存配置
        self.config["file_collect_src"] = src
        self.config["file_collect_folders"] = ",".join(folder_names)
        self.config["file_collect_custom_folders"] = self.custom_folder.text().strip()
        self.config["file_collect_ext"] = self.ext_input.text().strip()
        self.config["file_collect_kw"] = kw
        self.config["file_collect_struct"] = struct
        self.config["file_collect_naming"] = naming
        self.config["file_collect_template"] = template
        self.config["file_collect_output"] = output

        self.run_btn.setEnabled(False)
        self.log_area.clear()
        threading.Thread(
            target=self._do_collect,
            args=(src, folder_names, exts, kw, struct, naming, template, output),
            daemon=True,
        ).start()

    def _do_collect(self, src, folder_names, exts, kw, struct, naming, template, output):
        try:
            src = os.path.abspath(src)
            output = os.path.abspath(output)
            folder_set = set(folder_names)
            found = 0
            copied = 0
            counters = {}  # per-dest-dir counters for {num}

            for root, dirs, files in os.walk(src):
                current_name = os.path.basename(root)
                if current_name not in folder_set:
                    continue

                # 找到了匹配的文件夹
                rel = os.path.relpath(root, src)
                parts = rel.split(os.sep) if rel != "." else []
                parent = parts[-2] if len(parts) >= 2 else os.path.basename(src)
                matched = current_name

                for f in files:
                    # 文件筛选
                    ext = os.path.splitext(f)[1].lower()
                    if exts and ext not in exts:
                        continue
                    if kw and kw.lower() not in f.lower():
                        continue

                    found += 1
                    src_path = os.path.join(root, f)
                    name_no_ext = os.path.splitext(f)[0]

                    # 确定目标子目录
                    if struct == "parent":
                        dest_sub = parent
                    elif struct == "matched":
                        dest_sub = matched
                    elif struct == "both":
                        dest_sub = os.path.join(parent, matched)
                    else:  # flat
                        dest_sub = ""

                    # 确定新文件名
                    if naming == "original":
                        new_fname = f
                    elif naming == "parent_name":
                        new_fname = f"{parent}_{f}"
                    elif naming == "parent_num":
                        dest_key = dest_sub or "_flat"
                        idx = counters.get(dest_key, 0) + 1
                        counters[dest_key] = idx
                        new_fname = f"{parent}_{idx:04d}{ext}"
                    elif naming == "matched_parent_name":
                        new_fname = f"{matched}_{parent}_{f}"
                    elif naming == "custom":
                        dest_key = dest_sub or "_flat"
                        idx = counters.get(dest_key, 0) + 1
                        counters[dest_key] = idx
                        new_fname = template.replace("{parent}", parent).replace("{matched}", matched).replace("{name}", name_no_ext).replace("{ext}", ext).replace("{num}", str(idx)).replace("{num:04d}", f"{idx:04d}")
                    else:
                        new_fname = f

                    # 构建目标路径
                    if dest_sub:
                        dest_dir = os.path.join(output, dest_sub)
                    else:
                        dest_dir = output
                    os.makedirs(dest_dir, exist_ok=True)
                    dest_path = os.path.join(dest_dir, new_fname)

                    # 防重名
                    base, ext_f = os.path.splitext(new_fname)
                    counter = 2
                    while os.path.exists(dest_path):
                        dest_path = os.path.join(dest_dir, f"{base}_{counter}{ext_f}")
                        counter += 1

                    shutil.copy2(src_path, dest_path)
                    copied += 1
                    self.log_signal.emit(f"[提取] {os.path.relpath(src_path, src)} → {os.path.relpath(dest_path, output)}")

            self.log_signal.emit("")
            self.log_signal.emit(f"扫描到 {found} 个匹配文件，成功复制 {copied} 个")
            self.log_signal.emit(f"匹配文件夹: {', '.join(folder_names)}")
            self.log_signal.emit("=" * 50)
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
            import traceback
            self.log_signal.emit(traceback.format_exc())
        finally:
            self.run_btn.setEnabled(True)


# ── Tab 3: 图片归入同名文件夹 ────────────────────────────────────
class ImageToFolderTab(QWidget):
    log_signal = Signal(str)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self._append_log)
        self.init_ui()

    def _append_log(self, msg):
        self.log_area.append(msg)
        bar = self.log_area.verticalScrollBar()
        if bar:
            bar.setValue(bar.maximum())

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        desc = QLabel(
            "将目标目录下的图片文件，自动移入与之同名的文件夹中。\n"
            "例如: A001.jpg 移入 A001 文件夹；支持多种图片格式。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc)

        dir_row = QHBoxLayout()
        dir_row.addWidget(QLabel("目标目录:"))
        self.target_dir = DirDropLineEdit("拖拽或浏览选择包含图片和文件夹的目录...")
        self.target_dir.setText(self.config.get("img2folder_dir", ""))
        dir_row.addWidget(self.target_dir, 1)
        dir_btn = QPushButton("浏览")
        dir_btn.clicked.connect(self._browse_dir)
        dir_row.addWidget(dir_btn)
        layout.addLayout(dir_row)

        ext_row = QHBoxLayout()
        ext_row.addWidget(QLabel("图片格式:"))
        self.ext_input = QLineEdit()
        self.ext_input.setText(self.config.get("img2folder_ext", ".jpg .jpeg .png .bmp .webp .gif"))
        self.ext_input.setPlaceholderText("空格分隔，如: .jpg .png .bmp")
        ext_row.addWidget(self.ext_input, 1)
        layout.addLayout(ext_row)

        self.run_btn = QPushButton("▶️ 开始归入")
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #E65100; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #F57C00; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.run_btn.clicked.connect(self._start)
        layout.addWidget(self.run_btn)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        self.log_area.setStyleSheet("font-family: Consolas; font-size: 9px; background: #FAFAFA;")
        layout.addWidget(self.log_area)
        layout.addStretch()
        self.setLayout(layout)

    def _browse_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择目录")
        if path:
            self.target_dir.setText(path)

    def _start(self):
        target = self.target_dir.text().strip()
        if not target or not os.path.isdir(target):
            QMessageBox.warning(self, "提示", "请选择目标目录")
            return
        exts = [e.strip().lower() for e in self.ext_input.text().strip().split() if e.strip()]
        if not exts:
            QMessageBox.warning(self, "提示", "请至少输入一个图片格式")
            return

        self.config["img2folder_dir"] = target
        self.config["img2folder_ext"] = " ".join(exts)

        self.run_btn.setEnabled(False)
        self.log_area.clear()
        threading.Thread(target=self._do_move, args=(target, exts), daemon=True).start()

    def _do_move(self, target, exts):
        try:
            moved = 0
            skipped = 0
            for f in os.listdir(target):
                fp = os.path.join(target, f)
                if not os.path.isfile(fp):
                    continue
                ext = os.path.splitext(f)[1].lower()
                if ext not in exts:
                    continue
                name = os.path.splitext(f)[0]
                dest_folder = os.path.join(target, name)
                os.makedirs(dest_folder, exist_ok=True)
                dest = os.path.join(dest_folder, f)
                if os.path.exists(dest):
                    self.log_signal.emit(f"[跳过] {f} (目标已存在)")
                    skipped += 1
                    continue
                shutil.move(fp, dest)
                self.log_signal.emit(f"[移入] {f} → {name}/")
                moved += 1
            self.log_signal.emit("")
            self.log_signal.emit(f"移入 {moved} 个文件，跳过 {skipped} 个")
            self.log_signal.emit("=" * 50)
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
        finally:
            self.run_btn.setEnabled(True)


# ── Tab 4: 导出文件清单 ──────────────────────────────────────────
class FileListExportTab(QWidget):
    log_signal = Signal(str)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self._append_log)
        self.init_ui()

    def _append_log(self, msg):
        self.log_area.append(msg)
        bar = self.log_area.verticalScrollBar()
        if bar:
            bar.setValue(bar.maximum())

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        desc = QLabel("将目录下的所有文件/文件夹名称导出为文本文件，支持递归扫描。")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc)

        dir_row = QHBoxLayout()
        dir_row.addWidget(QLabel("扫描目录:"))
        self.target_dir = DirDropLineEdit("拖拽或浏览选择目录...")
        self.target_dir.setText(self.config.get("filelist_dir", ""))
        dir_row.addWidget(self.target_dir, 1)
        dir_btn = QPushButton("浏览")
        dir_btn.clicked.connect(self._browse_dir)
        dir_row.addWidget(dir_btn)
        layout.addLayout(dir_row)

        opt_row = QHBoxLayout()
        self.recursive_cb = QCheckBox("递归子文件夹")
        self.recursive_cb.setChecked(True)
        opt_row.addWidget(self.recursive_cb)
        self.only_files_cb = QCheckBox("仅文件")
        opt_row.addWidget(self.only_files_cb)
        self.only_dirs_cb = QCheckBox("仅文件夹")
        opt_row.addWidget(self.only_dirs_cb)
        opt_row.addStretch()
        layout.addLayout(opt_row)

        self.export_btn = QPushButton("📋 导出文件清单")
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #5E35B1; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #7E57C2; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.export_btn.clicked.connect(self._start)
        layout.addWidget(self.export_btn)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        self.log_area.setStyleSheet("font-family: Consolas; font-size: 9px; background: #FAFAFA;")
        layout.addWidget(self.log_area)
        layout.addStretch()
        self.setLayout(layout)

    def _browse_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择目录")
        if path:
            self.target_dir.setText(path)

    def _start(self):
        target = self.target_dir.text().strip()
        if not target or not os.path.isdir(target):
            QMessageBox.warning(self, "提示", "请选择目录")
            return
        self.config["filelist_dir"] = target
        self.export_btn.setEnabled(False)
        self.log_area.clear()
        threading.Thread(target=self._do_export, args=(target,), daemon=True).start()

    def _do_export(self, target):
        try:
            recursive = self.recursive_cb.isChecked()
            only_files = self.only_files_cb.isChecked()
            only_dirs = self.only_dirs_cb.isChecked()
            out_file = os.path.join(target, "文件清单.txt")
            lines = []
            lines.append(f"目录: {target}")
            lines.append(f"导出时间: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            lines.append("=" * 60)
            lines.append("")

            if recursive:
                for root, dirs, files in os.walk(target):
                    rel = os.path.relpath(root, target)
                    if not only_files:
                        for d in dirs:
                            lines.append(f"[D] {os.path.join(rel, d) if rel != '.' else d}")
                    if not only_dirs:
                        for f in files:
                            lines.append(f"[F] {os.path.join(rel, f) if rel != '.' else f}")
            else:
                for item in sorted(os.listdir(target)):
                    full = os.path.join(target, item)
                    is_dir = os.path.isdir(full)
                    if only_files and is_dir:
                        continue
                    if only_dirs and not is_dir:
                        continue
                    prefix = "[D]" if is_dir else "[F]"
                    lines.append(f"{prefix} {item}")

            with open(out_file, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))

            self.log_signal.emit(f"清单已导出: {out_file}")
            self.log_signal.emit(f"共 {len(lines) - 4} 条记录")
            self.log_signal.emit("=" * 50)
            try:
                os.startfile(out_file)
            except Exception:
                pass
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
        finally:
            self.export_btn.setEnabled(True)


# ── Tab 5: 批量分发文件 ──────────────────────────────────────────
class FileDistributeTab(QWidget):
    log_signal = Signal(str)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self._append_log)
        self.init_ui()

    def _append_log(self, msg):
        self.log_area.append(msg)
        bar = self.log_area.verticalScrollBar()
        if bar:
            bar.setValue(bar.maximum())

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        desc = QLabel(
            "将指定的文件批量复制到目标目录下的每一个子文件夹中。\n"
            "常用于: 将封面图、默认图等素材分发到各商品文件夹。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc)

        # 源文件
        src_row = QHBoxLayout()
        src_row.addWidget(QLabel("源文件:"))
        self.src_files = QLineEdit()
        self.src_files.setPlaceholderText("支持拖拽多个文件到此处")
        self.src_files.setText(self.config.get("distribute_files", ""))
        enable_path_drop(self.src_files, mode="file", multi=True)
        src_row.addWidget(self.src_files, 1)
        src_btn = QPushButton("浏览")
        src_btn.clicked.connect(self._browse_files)
        src_row.addWidget(src_btn)
        layout.addLayout(src_row)

        # 目标父目录
        dest_row = QHBoxLayout()
        dest_row.addWidget(QLabel("目标父目录:"))
        self.dest_dir = DirDropLineEdit("拖拽或浏览选择含子文件夹的父目录...")
        self.dest_dir.setText(self.config.get("distribute_dest", ""))
        dest_row.addWidget(self.dest_dir, 1)
        dest_btn = QPushButton("浏览")
        dest_btn.clicked.connect(self._browse_dest)
        dest_row.addWidget(dest_btn)
        layout.addLayout(dest_row)

        self.run_btn = QPushButton("▶️ 开始批量分发")
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #2E7D32; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #388E3C; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.run_btn.clicked.connect(self._start)
        layout.addWidget(self.run_btn)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        self.log_area.setStyleSheet("font-family: Consolas; font-size: 9px; background: #FAFAFA;")
        layout.addWidget(self.log_area)
        layout.addStretch()
        self.setLayout(layout)

    def _browse_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择要分发的文件", "", "所有文件 (*.*)")
        if paths:
            self.src_files.setText(";".join(paths))

    def _browse_dest(self):
        path = QFileDialog.getExistingDirectory(self, "选择目标父目录")
        if path:
            self.dest_dir.setText(path)

    def _start(self):
        src_text = self.src_files.text().strip()
        if not src_text:
            QMessageBox.warning(self, "提示", "请选择要分发的源文件")
            return
        dest = self.dest_dir.text().strip()
        if not dest or not os.path.isdir(dest):
            QMessageBox.warning(self, "提示", "请选择目标父目录")
            return

        self.config["distribute_files"] = src_text
        self.config["distribute_dest"] = dest

        self.run_btn.setEnabled(False)
        self.log_area.clear()

        src_paths = [p.strip() for p in src_text.split(";") if p.strip()]
        threading.Thread(target=self._do_distribute, args=(src_paths, dest), daemon=True).start()

    def _do_distribute(self, src_paths, dest):
        try:
            subdirs = [os.path.join(dest, d) for d in os.listdir(dest)
                       if os.path.isdir(os.path.join(dest, d))]
            if not subdirs:
                self.log_signal.emit("[错误] 目标目录下没有子文件夹")
                return

            total = 0
            for sub in subdirs:
                for src in src_paths:
                    if not os.path.isfile(src):
                        continue
                    fname = os.path.basename(src)
                    dest_path = os.path.join(sub, fname)
                    shutil.copy2(src, dest_path)
                    total += 1
                    self.log_signal.emit(f"[复制] {fname} → {os.path.basename(sub)}/")

            self.log_signal.emit("")
            self.log_signal.emit(f"分发完成: {len(src_paths)} 个文件 × {len(subdirs)} 个文件夹 = {total} 次复制")
            self.log_signal.emit("=" * 50)
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
        finally:
            self.run_btn.setEnabled(True)


# ── (已合并到 Tab 2: FileCollectExtractTab) ─────────────────
# ── Tab 7: 供应商款号提取 ────────────────────────────────────────
class SupplierCodeTab(QWidget):
    log_signal = Signal(str)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self._append_log)
        self._pywin32_ok = False
        self._check_pywin32()
        self.init_ui()

    def _check_pywin32(self):
        try:
            import pythoncom
            import win32com.client
            self._pywin32_ok = True
        except ImportError:
            self._pywin32_ok = False

    def _append_log(self, msg):
        self.log_area.append(msg)
        bar = self.log_area.verticalScrollBar()
        if bar:
            bar.setValue(bar.maximum())

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        desc = QLabel(
            "从Excel的商家编码/商品编码列自动提取供应商款号，在原文件旁生成带前缀的新文件。\n"
            "编码规则: 去尾部后缀(DD/DK/DZ/单数字/EKxx等)取最后剩余段为款号。\n"
            "使用 win32com 操作以保留 DISPIMG 嵌入式图片。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc)

        # 输入模式
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("输入方式:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItem("选择单个Excel文件", "file")
        self.mode_combo.addItem("选择文件夹(批量处理)", "folder")
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        mode_row.addWidget(self.mode_combo)
        mode_row.addStretch()
        layout.addLayout(mode_row)

        # 输入路径
        src_row = QHBoxLayout()
        src_row.addWidget(QLabel("输入路径:"))
        self.src_input = QLineEdit()
        self.src_input.setPlaceholderText("拖拽Excel或浏览选择...")
        self.src_input.setText(self.config.get("supplier_code_input", ""))
        enable_path_drop(self.src_input, mode="file_or_dir")
        src_row.addWidget(self.src_input, 1)
        self.src_btn = QPushButton("浏览")
        self.src_btn.clicked.connect(self._browse_src)
        src_row.addWidget(self.src_btn)
        layout.addLayout(src_row)

        # 目标列选择
        col_row = QHBoxLayout()
        col_row.addWidget(QLabel("目标列名:"))
        self.col_input = QLineEdit()
        self.col_input.setPlaceholderText("自动识别，也可手动输入(如: 商家编码)")
        self.col_input.setText(self.config.get("supplier_code_col", ""))
        col_row.addWidget(self.col_input, 1)
        layout.addLayout(col_row)
        hint = QLabel("留空则自动匹配: 商家编码 > 商品编码 > 货号")
        hint.setStyleSheet("color: #999; font-size: 9px;")
        layout.addWidget(hint)

        # 输出前缀
        prefix_row = QHBoxLayout()
        prefix_row.addWidget(QLabel("输出前缀:"))
        self.prefix_input = QLineEdit()
        self.prefix_input.setText(self.config.get("supplier_code_prefix", "供应商款号-"))
        prefix_row.addWidget(self.prefix_input)
        prefix_row.addWidget(QLabel("  新文件名 = 前缀 + 原文件名"))
        prefix_row.addStretch()
        layout.addLayout(prefix_row)

        self.run_btn = QPushButton("▶️ 开始提取供应商款号")
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #6A1B9A; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #8E24AA; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.run_btn.clicked.connect(self._start)
        layout.addWidget(self.run_btn)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        self.log_area.setStyleSheet("font-family: Consolas; font-size: 9px; background: #FAFAFA;")
        layout.addWidget(self.log_area)
        layout.addStretch()
        self.setLayout(layout)

    def _on_mode_changed(self):
        mode = self.mode_combo.currentData()
        if mode == "file":
            self.src_input.setPlaceholderText("拖拽Excel文件或浏览选择...")
        else:
            self.src_input.setPlaceholderText("拖拽文件夹或浏览选择...")

    def _browse_src(self):
        mode = self.mode_combo.currentData()
        if mode == "file":
            path, _ = QFileDialog.getOpenFileName(self, "选择Excel", "", "Excel (*.xlsx *.xls)")
            if path:
                self.src_input.setText(path)
        else:
            path = QFileDialog.getExistingDirectory(self, "选择文件夹")
            if path:
                self.src_input.setText(path)

    def _start(self):
        if not self._pywin32_ok:
            ret = QMessageBox.question(
                self, "缺少依赖",
                "供应商款号提取需要 pywin32 库，是否现在安装？\n\n"
                "安装后需要重启程序。",
                QMessageBox.Yes | QMessageBox.No
            )
            if ret == QMessageBox.Yes:
                import subprocess, sys
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32", "-q"])
                    QMessageBox.information(self, "完成", "pywin32 安装完成，请重启程序。")
                    self._pywin32_ok = True
                except Exception as e:
                    QMessageBox.critical(self, "失败", f"安装失败: {e}")
            return

        src = self.src_input.text().strip()
        if not src or not os.path.exists(src):
            QMessageBox.warning(self, "提示", "请选择有效的输入路径")
            return

        self.config["supplier_code_input"] = src
        self.config["supplier_code_col"] = self.col_input.text().strip()
        self.config["supplier_code_prefix"] = self.prefix_input.text().strip() or "供应商款号-"

        self.run_btn.setEnabled(False)
        self.log_area.clear()
        threading.Thread(target=self._do_extract, daemon=True).start()

    def _find_files(self, src):
        mode = self.mode_combo.currentData()
        if mode == "file":
            if src.endswith((".xlsx", ".xls")) and not os.path.basename(src).startswith("~$"):
                return [src]
            return []
        else:
            import glob
            files = glob.glob(os.path.join(src, "*.xlsx"))
            files = [f for f in files
                     if not os.path.basename(f).startswith("供应商款号-")
                     and not os.path.basename(f).startswith("~$")]
            return sorted(files)

    def _find_column(self, ws, max_col, target_col_name):
        for col in range(1, max_col + 1):
            val = ws.Cells(1, col).Value
            if val and target_col_name in str(val):
                return col
        return None

    def _extract_code(self, code):
        if not code:
            return ""
        parts = str(code).split("-")
        if len(parts) <= 1:
            return parts[0]

        def is_suffix(seg):
            if re.match(r'^[0-9]$', seg):
                return True
            if seg in ("01", "02"):
                return True
            if seg in ("DD", "DK", "DZ"):
                return True
            if re.match(r'^EK.{2}$', seg):
                return True
            return False

        suffix_count = 0
        for i in range(len(parts) - 1, -1, -1):
            if is_suffix(parts[i]):
                suffix_count += 1
            else:
                break

        if suffix_count > 0:
            start = len(parts) - suffix_count - 1
            return "-".join(parts[start:])

        if len(parts[-1]) < 4 and len(parts) >= 2:
            return "-".join(parts[-2:])
        return parts[-1]

    def _do_extract(self):
        import pythoncom
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = __import__("win32com.client", fromlist=["win32com"]).Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            src = self.src_input.text().strip()
            files = self._find_files(src)
            if not files:
                self.log_signal.emit("[错误] 未找到可处理的Excel文件")
                return

            col_name = self.col_input.text().strip() or ""
            prefix = self.prefix_input.text().strip() or "供应商款号-"
            XL_FORMAT = 51

            self.log_signal.emit(f"找到 {len(files)} 个文件待处理")
            self.log_signal.emit(f"输出前缀: {prefix}")
            self.log_signal.emit("")

            success = 0
            for i, fpath in enumerate(files, 1):
                fname = os.path.basename(fpath)
                self.log_signal.emit(f"--- 处理 ({i}/{len(files)}): {fname} ---")
                try:
                    wb = excel.Workbooks.Open(fpath)
                    ws = wb.ActiveSheet
                    used = ws.UsedRange
                    max_row = used.Rows.Count
                    max_col = used.Columns.Count

                    # 找列
                    col_idx = None
                    if col_name:
                        col_idx = self._find_column(ws, max_col, col_name)
                    if col_idx is None:
                        for cn in ("商家编码", "商品编码", "货号"):
                            col_idx = self._find_column(ws, max_col, cn)
                            if col_idx is not None:
                                break

                    if col_idx is None:
                        self.log_signal.emit(f"  [跳过] 未找到目标列，表头: {[ws.Cells(1,c).Value for c in range(1,min(6,max_col+1))]}")
                        wb.Close(SaveChanges=False)
                        continue

                    header = ws.Cells(1, col_idx).Value
                    self.log_signal.emit(f"  定位列: \"{header}\" (第{col_idx}列)")

                    new_col = col_idx + 1
                    try:
                        ws.Columns(new_col).Insert()
                    except Exception:
                        pass
                    ws.Cells(1, new_col).Value = "供应商款号"

                    changed = 0
                    for row in range(2, max_row + 1):
                        val = ws.Cells(row, col_idx).Value
                        if val is None:
                            continue
                        new_val = self._extract_code(str(val))
                        ws.Cells(row, new_col).Value = new_val
                        changed += 1
                        if row <= 5:
                            self.log_signal.emit(f"  Row {row}: {str(val)[:40]} → {new_val}")

                    folder = os.path.dirname(fpath) or "."
                    name = os.path.basename(fpath)
                    out_path = os.path.join(folder, f"{prefix}{name}")
                    if os.path.exists(out_path):
                        os.remove(out_path)
                    wb.SaveAs(out_path, XL_FORMAT)
                    self.log_signal.emit(f"  [完成] 处理 {changed} 行 → {prefix}{name}")
                    success += 1
                    wb.Close(SaveChanges=False)
                except Exception as e:
                    self.log_signal.emit(f"  [错误] {e}")
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass

            self.log_signal.emit("")
            self.log_signal.emit(f"全部完成: 成功 {success}/{len(files)}")
            self.log_signal.emit("=" * 50)
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
            import traceback
            self.log_signal.emit(traceback.format_exc())
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self.run_btn.setEnabled(True)


# ── Tab 8: 清理重复文件 ──────────────────────────────────────────
class DedupCleanerTab(QWidget):
    log_signal = Signal(str)

    PATTERNS = [
        re.compile(r'^(.*)\((\d+)\)(\.[^.]+)$'),
        re.compile(r'^(.*) - 副本(?: \((\d+)\))?(\.[^.]+)?$'),
        re.compile(r'^副本 (.*)(\.[^.]+)$'),
        re.compile(r'^(.*) - Copy(?: \((\d+)\))?(\.[^.]+)?$'),
        re.compile(r'^Copy of (.*)(\.[^.]+)$'),
        re.compile(r'^(.*) \((\d+)\)(\.[^.]+)$'),
        re.compile(r"^(.*) \([\w-]+'s conflicted copy \d{4}-\d{2}-\d{2}\)(\.[^.]+)$"),
        re.compile(r'^(.*)_conflicted copy_\d{4}-\d{2}-\d{2}(\.[^.]+)$'),
        re.compile(r'^(.*)-([0-9a-f]{8,40})(\.[^.]+)$'),
        re.compile(r'^(.*)_([0-9a-f]{8,40})(\.[^.]+)$'),
        re.compile(r'^(.*)(\.backup_\d{8})(\.[^.]+)$'),
    ]
    SKIP_DIRS = {".opencode", "__pycache__", "node_modules", ".git", ".svn"}

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.log_signal.connect(self._append_log)
        self._duplicates = []
        self.init_ui()

    def _append_log(self, msg):
        self.log_area.append(msg)
        bar = self.log_area.verticalScrollBar()
        if bar:
            bar.setValue(bar.maximum())

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        desc = QLabel(
            "扫描目录中的重复文件(Windows自动重命名、中文/英文副本、OneDrive冲突/哈希备份、手动备份等11种模式)。\n"
            "删除前会列出全部匹配项并二次确认，原始文件不受影响。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(desc)

        dir_row = QHBoxLayout()
        dir_row.addWidget(QLabel("目标目录:"))
        self.target_dir = DirDropLineEdit("拖拽或浏览选择要清理的目录...")
        self.target_dir.setText(self.config.get("dedup_dir", ""))
        dir_row.addWidget(self.target_dir, 1)
        dir_btn = QPushButton("浏览")
        dir_btn.clicked.connect(self._browse_dir)
        dir_row.addWidget(dir_btn)
        layout.addLayout(dir_row)

        btn_row = QHBoxLayout()
        self.scan_btn = QPushButton("🔍 扫描重复文件")
        self.scan_btn.setStyleSheet("""
            QPushButton {
                background-color: #1565C0; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #1976D2; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.scan_btn.clicked.connect(self._scan)
        btn_row.addWidget(self.scan_btn)

        self.delete_btn = QPushButton("🗑️ 确认删除重复文件")
        self.delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #C62828; color: white; border: none;
                border-radius: 6px; padding: 9px 16px; font-size: 11px; font-weight: bold;
            }
            QPushButton:hover { background-color: #D32F2F; }
            QPushButton:disabled { background-color: #B0BEC5; color: #ECEFF1; }
        """)
        self.delete_btn.setEnabled(False)
        self.delete_btn.clicked.connect(self._confirm_delete)
        btn_row.addWidget(self.delete_btn)
        layout.addLayout(btn_row)

        self.stat_label = QLabel("")
        self.stat_label.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(self.stat_label)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(300)
        self.log_area.setStyleSheet("font-family: Consolas; font-size: 9px; background: #FAFAFA;")
        layout.addWidget(self.log_area)
        layout.addStretch()
        self.setLayout(layout)

    def _browse_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择目录")
        if path:
            self.target_dir.setText(path)

    def _format_size(self, b):
        if b < 1024:
            return f"{b}B"
        elif b < 1024 ** 2:
            return f"{b / 1024:.1f}KB"
        elif b < 1024 ** 3:
            return f"{b / 1024 ** 2:.1f}MB"
        else:
            return f"{b / 1024 ** 3:.2f}GB"

    def _match(self, name, parent_path, is_dir=False):
        for pat in self.PATTERNS:
            m = pat.match(name)
            if m:
                stem = m.group(1)
                if not is_dir and m.lastindex >= 3:
                    suffix = m.group(3) or ""
                elif not is_dir:
                    suffix = m.group(2) or ""
                else:
                    suffix = ""
                orig = os.path.join(parent_path, stem + suffix)
                backup = os.path.join(parent_path, name)
                if os.path.exists(orig):
                    return (orig, backup)
        return None

    def _scan(self):
        target = self.target_dir.text().strip()
        if not target or not os.path.isdir(target):
            QMessageBox.warning(self, "提示", "请选择有效的目录")
            return
        self.config["dedup_dir"] = target

        self.scan_btn.setEnabled(False)
        self.delete_btn.setEnabled(False)
        self._duplicates = []
        self.log_area.clear()
        self.stat_label.setText("正在扫描...")
        threading.Thread(target=self._do_scan, args=(target,), daemon=True).start()

    def _do_scan(self, target):
        try:
            scanned = 0
            for dirpath, dirnames, filenames in os.walk(target):
                rel = os.path.relpath(dirpath, target)
                parts = set(p.lower() for p in (rel.split(os.sep) if rel != "." else []))
                if parts & self.SKIP_DIRS:
                    continue
                for fname in filenames:
                    scanned += 1
                    r = self._match(fname, dirpath)
                    if r:
                        self._duplicates.append(r)
                for dname in dirnames[:]:
                    r = self._match(dname, dirpath, is_dir=True)
                    if r and os.path.isdir(r[0]):
                        self._duplicates.append(r)

            self._duplicates.sort(key=lambda x: x[1])
            total_bytes = sum(
                os.path.getsize(f) for _, f in self._duplicates if os.path.isfile(f)
            )

            target_abs = os.path.abspath(target)
            self.log_signal.emit(f"扫描目录: {target}")
            self.log_signal.emit(f"扫描文件数: {scanned}")
            self.log_signal.emit(f"发现重复: {len(self._duplicates)} 个")
            self.log_signal.emit(f"可释放空间: {self._format_size(total_bytes)}")
            self.log_signal.emit("-" * 60)

            MAX_SHOW = 100
            for orig, backup in self._duplicates[:MAX_SHOW]:
                sz = os.path.getsize(backup) if os.path.isfile(backup) else 0
                rel_bak = os.path.relpath(backup, target_abs)
                rel_org = os.path.relpath(orig, target_abs)
                self.log_signal.emit(f"[删] {self._format_size(sz):>8}  {rel_bak}")
                self.log_signal.emit(f"[留]           {rel_org}")

            if len(self._duplicates) > MAX_SHOW:
                self.log_signal.emit(f"  ... 还有 {len(self._duplicates) - MAX_SHOW} 个未显示")

            self.log_signal.emit("=" * 60)
            self.stat_label.setText(
                f"扫描完成: 发现 {len(self._duplicates)} 个重复文件, "
                f"可释放 {self._format_size(total_bytes)}"
            )
            if self._duplicates:
                self.delete_btn.setEnabled(True)
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
        finally:
            self.scan_btn.setEnabled(True)

    def _confirm_delete(self):
        if not self._duplicates:
            return
        total_bytes = sum(
            os.path.getsize(f) for _, f in self._duplicates if os.path.isfile(f)
        )
        ret = QMessageBox.warning(
            self, "⚠️ 确认删除",
            f"即将删除 {len(self._duplicates)} 个重复文件\n"
            f"可释放空间: {self._format_size(total_bytes)}\n\n"
            f"原始文件不受影响，仅删除重复/备份文件。\n\n"
            f"确定要继续吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if ret != QMessageBox.Yes:
            return
        self.delete_btn.setEnabled(False)
        self.scan_btn.setEnabled(False)
        self.log_signal.emit("")
        self.log_signal.emit("正在删除...")
        threading.Thread(target=self._do_delete, daemon=True).start()

    def _do_delete(self):
        try:
            ok = fail = total = 0
            target = self.target_dir.text().strip()
            log_dir = os.path.join(target, ".opencode")
            os.makedirs(log_dir, exist_ok=True)
            from datetime import datetime
            log_path = os.path.join(log_dir, f"清理重复文件_{datetime.now():%Y%m%d_%H%M%S}.log")
            log_lines = []

            for orig, backup in self._duplicates:
                try:
                    sz = os.path.getsize(backup) if os.path.isfile(backup) else 0
                    total += sz
                    if os.path.isdir(backup):
                        shutil.rmtree(backup)
                    else:
                        os.remove(backup)
                    ok += 1
                    msg = f"[已删] ({self._format_size(sz)}) {os.path.relpath(backup, target)}"
                    self.log_signal.emit(msg)
                    log_lines.append(msg)
                except Exception as e:
                    fail += 1
                    msg = f"[失败] {os.path.relpath(backup, target)} -> {e}"
                    self.log_signal.emit(msg)
                    log_lines.append(msg)

            with open(log_path, "w", encoding="utf-8") as f:
                f.write("\n".join(log_lines))

            self.log_signal.emit("")
            self.log_signal.emit(f"清理完成! 成功: {ok} | 失败: {fail} | 释放: {self._format_size(total)}")
            self.log_signal.emit(f"日志已保存: {log_path}")
            self.log_signal.emit("=" * 50)
            self.stat_label.setText(f"清理完成: 删除 {ok} 个, 失败 {fail} 个")
        except Exception as e:
            self.log_signal.emit(f"[异常] {e}")
        finally:
            self.scan_btn.setEnabled(True)
            self.delete_btn.setEnabled(False)
            self._duplicates = []


# ── 主页面：文件工具 ──────────────────────────────────────────────
class FileToolsPage(QWidget):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(6)
        self.setLayout(layout)

        # 全局样式
        self.setStyleSheet("""
            QGroupBox {
                border: 1px solid #D6E4F0;
                border-radius: 8px;
                margin-top: 10px;
                background: #FFFFFF;
                font-family: Microsoft YaHei;
                font-size: 10px;
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
            QLineEdit, QComboBox {
                border: 1px solid #CFD8DC;
                border-radius: 6px;
                padding: 6px 8px;
                background: #FAFAFA;
                min-height: 26px;
                font-family: Microsoft YaHei;
                font-size: 10px;
            }
            QLineEdit:focus, QComboBox:focus {
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
            QTabWidget::pane {
                border: 1px solid #D6E4F0;
                border-radius: 8px;
                background: #FFFFFF;
            }
            QTabBar::tab {
                background: #E3F2FD;
                color: #1565C0;
                border: 1px solid #BBDEFB;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 8px 16px;
                font-family: Microsoft YaHei;
                font-size: 10px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                color: #0D47A1;
                font-weight: bold;
                border-bottom: 2px solid #1565C0;
            }
            QTabBar::tab:hover {
                background: #BBDEFB;
            }
        """)

        title = QLabel("📁 文件工具")
        title.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        title.setStyleSheet("color: #1565C0;")
        layout.addWidget(title)

        desc = QLabel("文件夹批量操作 & 表格驱动工具集，涵盖建文件夹、提取汇总、分发、清单导出等日常办公高频操作")
        desc.setWordWrap(True)
        desc.setFont(QFont("Microsoft YaHei", 10))
        desc.setStyleSheet("color: #666;")
        layout.addWidget(desc)

        # 滚动区域包裹标签页
        self.tabs = QTabWidget()

        self.tab1 = FolderCreatorTab(self.config)
        self.tab2 = FileCollectExtractTab(self.config)
        self.tab3 = ImageToFolderTab(self.config)
        self.tab4 = FileListExportTab(self.config)
        self.tab5 = FileDistributeTab(self.config)
        self.tab6 = SupplierCodeTab(self.config)
        self.tab7 = DedupCleanerTab(self.config)

        # 包裹到 QScrollArea
        for tab, name in [
            (self.tab1, "📁 批量建文件夹"),
            (self.tab2, "📦 提取文件汇总"),
            (self.tab3, "🖼️ 图片归入同名文件夹"),
            (self.tab4, "📋 导出文件清单"),
            (self.tab5, "📨 批量分发文件"),
            (self.tab6, "🔖 供应商款号提取"),
            (self.tab7, "🧹 清理重复文件"),
        ]:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setWidget(tab)
            scroll.setFrameShape(QScrollArea.NoFrame)
            self.tabs.addTab(scroll, name)

        layout.addWidget(self.tabs)
