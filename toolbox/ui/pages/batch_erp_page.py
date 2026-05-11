# -*- coding: utf-8 -*-
"""批量模板 & ERP商品资料 合并页面"""
import os
import logging
import subprocess

from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
    QLineEdit, QPushButton, QCheckBox, QProgressBar,
    QTextEdit, QFileDialog, QMessageBox, QApplication,
    QSpinBox, QFrame,
)

from toolbox.core.batch_template import TemplateGenerator
from toolbox.core.erp_product import ErpProductGenerator
from toolbox.core.utils import save_config
from toolbox.ui.path_drop import DirDropLineEdit, ExcelDropLineEdit

LOGGER = logging.getLogger("batch_erp_page")

# ============ 后台线程 ============
class BatchErpThread(QThread):
    progress_signal = Signal(str, int, int)
    done_signal = Signal(object)

    def __init__(self, sel_path, info_path, inv_path, output_dir, min_stock,
                 run_batch_qa, run_batch_acc, run_batch_tryon, run_batch_attr,
                 run_erp_barcode, run_erp_add, run_erp_price, parent=None):
        super().__init__(parent)
        self.sel_path = sel_path
        self.info_path = info_path
        self.inv_path = inv_path
        self.output_dir = output_dir
        self.min_stock = min_stock
        self.run_batch_qa = run_batch_qa
        self.run_batch_acc = run_batch_acc
        self.run_batch_tryon = run_batch_tryon
        self.run_batch_attr = run_batch_attr
        self.run_erp_barcode = run_erp_barcode
        self.run_erp_add = run_erp_add
        self.run_erp_price = run_erp_price

    def run(self):
        try:
            from toolbox.core.batch_template import BatchResult
            from toolbox.core.erp_product import ErpBatchResult, ErpGenerationResult

            # 计算总步数
            total_steps = 0
            if self.run_batch_qa: total_steps += 1
            if self.run_batch_acc: total_steps += 1
            if self.run_batch_tryon: total_steps += 1
            if self.run_batch_attr: total_steps += 1
            if self.run_erp_barcode: total_steps += 1
            if self.run_erp_add: total_steps += 1
            if self.run_erp_price: total_steps += 1

            def cancelled():
                return self.isInterruptionRequested()

            current = 0
            batch_result = None
            erp_result = None

            # === 批量模板生成 ===
            needs_batch = any([self.run_batch_qa, self.run_batch_acc,
                               self.run_batch_tryon, self.run_batch_attr])
            if needs_batch and not cancelled():
                bt_out = os.path.join(self.output_dir, 'QA&属性等')
                os.makedirs(bt_out, exist_ok=True)
                gen = TemplateGenerator(
                    self.sel_path, bt_out,
                    progress_callback=lambda s, c, t: self.progress_signal.emit(s, current + c, total_steps),
                )
                gen._load()

                batch_result = BatchResult(
                    input_path=self.sel_path, output_dir=bt_out,
                    product_count=len(gen.products),
                )

                if self.run_batch_qa and not cancelled():
                    current += 1
                    self.progress_signal.emit('QA', current, total_steps)
                    batch_result.qa = gen.run_qa()

                if self.run_batch_acc and not cancelled():
                    current += 1
                    self.progress_signal.emit('配件明细', current, total_steps)
                    batch_result.accessories = gen.run_accessories()

                if self.run_batch_tryon and not cancelled():
                    current += 1
                    self.progress_signal.emit('试穿报告', current, total_steps)
                    batch_result.tryon = gen.run_tryon()

                if self.run_batch_attr and not cancelled():
                    current += 1
                    self.progress_signal.emit('属性表格', current, total_steps)
                    batch_result.attributes = gen.run_attributes()

            # === ERP商品资料生成 ===
            needs_erp = any([self.run_erp_barcode, self.run_erp_add, self.run_erp_price])
            if needs_erp and not cancelled():
                erp_out = os.path.join(self.output_dir, 'erp')
                os.makedirs(erp_out, exist_ok=True)
                gen2 = ErpProductGenerator(
                    self.info_path, self.inv_path, self.sel_path, erp_out,
                    self.min_stock,
                    progress_callback=lambda s, c, t: self.progress_signal.emit(s, current + c, total_steps),
                    error_dir=self.output_dir,
                )

                erp_result = ErpBatchResult()

                if self.run_erp_barcode and not cancelled():
                    current += 1
                    self.progress_signal.emit('ERP条码对照表', current, total_steps)
                    erp_result.erp_barcode = gen2.run_erp_barcode()

                if self.run_erp_add and not cancelled():
                    current += 1
                    self.progress_signal.emit('批量新增商品资料', current, total_steps)
                    erp_result.batch_add = gen2.run_batch_add()

                if self.run_erp_price and not cancelled():
                    current += 1
                    self.progress_signal.emit('定价导入', current, total_steps)
                    erp_result.pricing = gen2.run_pricing()

                # 收集警告
                all_w = []
                for r in [erp_result.erp_barcode, erp_result.batch_add, erp_result.pricing]:
                    if r and r.warnings:
                        all_w.extend(r.warnings)
                erp_result.warnings = all_w

            self.progress_signal.emit('完成', total_steps, total_steps)
            self.done_signal.emit((batch_result, erp_result))
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.done_signal.emit(e)


# ============ 主页面 ============
class BatchErpPage(QWidget):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self._thread = None
        self._last_output_dir = ""
        self.init_ui()
        self._restore_paths()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 18, 24, 18)
        main_layout.setSpacing(12)
        font = QFont("Microsoft YaHei", 10)

        # === 标题 ===
        title = QLabel("批量模板 & ERP商品资料")
        title.setFont(QFont("Microsoft YaHei", 18, QFont.Bold))
        title.setStyleSheet("color: #333;")
        main_layout.addWidget(title)

        desc = QLabel("选款资料表必选；商品信息表+商品库存表选填（ERP功能需要3表齐全）")
        desc.setFont(font)
        desc.setStyleSheet("color: #888;")
        desc.setWordWrap(True)
        main_layout.addWidget(desc)

        # === 输入文件 ===
        input_group = QGroupBox("输入文件")
        input_group.setFont(font)
        input_layout = QVBoxLayout(input_group)
        input_layout.setSpacing(8)

        # Row 1: 选款资料表
        r1 = QHBoxLayout()
        lbl = QLabel("选款资料表:")
        lbl.setFont(font); lbl.setFixedWidth(100)
        lbl.setStyleSheet("color: #D32F2F; font-weight: bold;")  # 必填标红
        r1.addWidget(lbl)
        self.sel_edit = ExcelDropLineEdit("拖拽或选择选款表（必填，含 商家编码/唯品款号/品牌/类目/价格）")
        self.sel_edit.excel_dropped.connect(self._on_input_changed)
        r1.addWidget(self.sel_edit)
        btn = QPushButton("浏览..."); btn.setFont(font); btn.setFixedWidth(80)
        btn.clicked.connect(lambda: self._browse_file(self.sel_edit, "选择选款资料表"))
        r1.addWidget(btn)
        input_layout.addLayout(r1)

        # Row 2: 商品信息表
        r2 = QHBoxLayout()
        lbl2 = QLabel("商品信息表:")
        lbl2.setFont(font); lbl2.setFixedWidth(100)
        lbl2.setStyleSheet("color: #888;")  # 选填
        r2.addWidget(lbl2)
        self.info_edit = ExcelDropLineEdit("拖拽或选择商品信息表（ERP功能需要）")
        self.info_edit.excel_dropped.connect(self._on_input_changed)
        r2.addWidget(self.info_edit)
        btn2 = QPushButton("浏览..."); btn2.setFont(font); btn2.setFixedWidth(80)
        btn2.clicked.connect(lambda: self._browse_file(self.info_edit, "选择商品信息表"))
        r2.addWidget(btn2)
        input_layout.addLayout(r2)

        # Row 3: 商品库存表
        r3 = QHBoxLayout()
        lbl3 = QLabel("商品库存表:")
        lbl3.setFont(font); lbl3.setFixedWidth(100)
        lbl3.setStyleSheet("color: #888;")
        r3.addWidget(lbl3)
        self.inv_edit = ExcelDropLineEdit("拖拽或选择商品库存表（ERP功能需要）")
        self.inv_edit.excel_dropped.connect(self._on_input_changed)
        r3.addWidget(self.inv_edit)
        btn3 = QPushButton("浏览..."); btn3.setFont(font); btn3.setFixedWidth(80)
        btn3.clicked.connect(lambda: self._browse_file(self.inv_edit, "选择商品库存表"))
        r3.addWidget(btn3)
        input_layout.addLayout(r3)

        main_layout.addWidget(input_group)

        # === 库存阈值 + 输出目录 ===
        settings_row = QHBoxLayout()
        settings_row.setSpacing(16)

        stock_group = QGroupBox("库存阈值 (ERP用)")
        stock_group.setFont(font)
        stock_layout = QHBoxLayout(stock_group)
        self.stock_spin = QSpinBox()
        self.stock_spin.setRange(1, 9999)
        self.stock_spin.setValue(10)
        self.stock_spin.setFont(font); self.stock_spin.setFixedWidth(80)
        self.stock_spin.setSuffix(" 件")
        self.stock_spin.setEnabled(False)
        self.stock_spin.setToolTip("需要3个输入表全部选择后才可调整")
        stock_layout.addWidget(self.stock_spin)
        stock_layout.addWidget(QLabel("色库存低于此值删除该颜色"))
        settings_row.addWidget(stock_group)

        out_group = QGroupBox("输出目录")
        out_group.setFont(font)
        out_layout = QHBoxLayout(out_group)
        self.out_edit = DirDropLineEdit("拖拽或选择输出目录，默认为选款表同目录下的 output")
        out_layout.addWidget(self.out_edit)
        btn_out = QPushButton("浏览..."); btn_out.setFont(font); btn_out.setFixedWidth(80)
        btn_out.clicked.connect(lambda: self._browse_dir())
        out_layout.addWidget(btn_out)
        settings_row.addWidget(out_group, 1)

        main_layout.addLayout(settings_row)

        # === 导出内容 ===
        opt_group = QGroupBox("导出内容")
        opt_group.setFont(font)
        opt_layout = QVBoxLayout(opt_group)
        opt_layout.setSpacing(8)

        # 批量模板组
        batch_label = QLabel("▸ 批量模板生成（仅需选款资料表）")
        batch_label.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        batch_label.setStyleSheet("color: #333;")
        opt_layout.addWidget(batch_label)

        batch_row = QHBoxLayout()
        self.chk_qa = QCheckBox("QA表格")
        self.chk_acc = QCheckBox("配件明细")
        self.chk_tryon = QCheckBox("试穿报告")
        self.chk_attr = QCheckBox("属性表格")
        for chk in [self.chk_qa, self.chk_acc, self.chk_tryon, self.chk_attr]:
            chk.setFont(font); chk.setChecked(True)
            batch_row.addWidget(chk)
        batch_row.addStretch()
        opt_layout.addLayout(batch_row)

        # 分隔线
        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("background-color: #E0E0E0;"); sep.setFixedHeight(1)
        opt_layout.addWidget(sep)

        # ERP组
        erp_label = QLabel("▸ ERP商品资料（需3表齐全）")
        erp_label.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        erp_label.setStyleSheet("color: #333;")
        opt_layout.addWidget(erp_label)

        erp_row = QHBoxLayout()
        self.chk_erp_barcode = QCheckBox("ERP条码对照表")
        self.chk_erp_add = QCheckBox("批量新增商品资料")
        self.chk_erp_price = QCheckBox("定价导入-扣点模式")
        for chk in [self.chk_erp_barcode, self.chk_erp_add, self.chk_erp_price]:
            chk.setFont(font); chk.setChecked(True)
            erp_row.addWidget(chk)
        erp_row.addStretch()
        opt_layout.addLayout(erp_row)

        main_layout.addWidget(opt_group)

        # === 运行按钮 + 进度条 ===
        run_layout = QHBoxLayout()
        self.run_btn = QPushButton("开始生成")
        self.run_btn.setFont(QFont("Microsoft YaHei", 13, QFont.Bold))
        self.run_btn.setFixedHeight(44); self.run_btn.setFixedWidth(140)
        self.run_btn.setStyleSheet("""
            QPushButton { background-color: #1E88E5; color: white;
                          border: none; border-radius: 6px; padding: 8px 20px; }
            QPushButton:hover { background-color: #1976D2; }
            QPushButton:pressed { background-color: #1565C0; }
            QPushButton:disabled { background-color: #BDBDBD; }
        """)
        self.run_btn.clicked.connect(self._start_generation)
        run_layout.addWidget(self.run_btn)

        self.stop_btn = QPushButton("⏹ 停止")
        self.stop_btn.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        self.stop_btn.setFixedHeight(44)
        self.stop_btn.setStyleSheet("""
            QPushButton { background-color: #D32F2F; color: white;
                          border: none; border-radius: 6px; padding: 8px 16px; }
            QPushButton:hover { background-color: #C62828; }
            QPushButton:disabled { background-color: #BDBDBD; }
        """)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop_generation)
        run_layout.addWidget(self.stop_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFont(font); self.progress_bar.setFixedHeight(32)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("就绪 - 请选择输入文件")
        self.progress_bar.setStyleSheet("""
            QProgressBar { border: 1px solid #E0E0E0; border-radius: 4px;
                           text-align: center; background: #F5F5F5; }
            QProgressBar::chunk { background-color: #1E88E5; border-radius: 3px; }
        """)
        run_layout.addWidget(self.progress_bar, 1)
        main_layout.addLayout(run_layout)

        # === 日志 ===
        log_group = QGroupBox("运行日志")
        log_group.setFont(font)
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(130)
        self.log_text.setStyleSheet("QTextEdit{border:1px solid #E0E0E0;background:white;padding:6px;}")
        log_layout.addWidget(self.log_text)
        main_layout.addWidget(log_group)

        # 初始状态刷新
        self._refresh_erp_state()

    # ========== 动态控制 ==========
    def _has_sel(self): return bool(self.sel_edit.text().strip()) and os.path.exists(self.sel_edit.text().strip())
    def _has_info(self): return bool(self.info_edit.text().strip()) and os.path.exists(self.info_edit.text().strip())
    def _has_inv(self): return bool(self.inv_edit.text().strip()) and os.path.exists(self.inv_edit.text().strip())
    def _erp_ready(self): return self._has_sel() and self._has_info() and self._has_inv()

    def _refresh_erp_state(self):
        """根据输入状态启用/禁用ERP相关控件"""
        erp_ok = self._erp_ready()
        for chk in [self.chk_erp_barcode, self.chk_erp_add, self.chk_erp_price]:
            chk.setEnabled(erp_ok)
            if not erp_ok:
                chk.setToolTip("需要选款资料表+商品信息表+商品库存表全部选择后才可用")
        self.stock_spin.setEnabled(erp_ok)
        if not erp_ok:
            self.stock_spin.setToolTip("需要3个输入表全部选择后才可调整")

    def _on_input_changed(self, path):
        """任一输入变化时刷新状态"""
        self._refresh_erp_state()
        if not self.out_edit.text() and self._has_sel():
            self.out_edit.setText(os.path.join(os.path.dirname(self.sel_edit.text().strip()), 'output'))

    # ========== 文件浏览 ==========
    def _browse_file(self, edit, title):
        path, _ = QFileDialog.getOpenFileName(self, title, "", "Excel 文件 (*.xlsx *.xls);;所有文件 (*)")
        if path:
            edit.setText(path)
            self._on_input_changed(path)
            if not self.out_edit.text() and edit == self.sel_edit:
                self.out_edit.setText(os.path.join(os.path.dirname(path), 'output'))

    def _browse_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if path:
            self.out_edit.setText(path)

    # ========== 生成 ==========
    def _start_generation(self):
        sel_path = self.sel_edit.text().strip()
        info_path = self.info_edit.text().strip()
        inv_path = self.inv_edit.text().strip()
        output_dir = self.out_edit.text().strip()
        min_stock = self.stock_spin.value()

        # 至少需要选款资料表
        if not sel_path or not os.path.exists(sel_path):
            QMessageBox.warning(self, "提示", "请选择选款资料表（必填）")
            return

        if not output_dir:
            output_dir = os.path.join(os.path.dirname(sel_path), 'output')
            self.out_edit.setText(output_dir)

        # 检查批量模板勾选
        run_batch_qa = self.chk_qa.isChecked()
        run_batch_acc = self.chk_acc.isChecked()
        run_batch_tryon = self.chk_tryon.isChecked()
        run_batch_attr = self.chk_attr.isChecked()
        has_batch = any([run_batch_qa, run_batch_acc, run_batch_tryon, run_batch_attr])

        # 检查ERP勾选 (且必须3表齐全)
        erp_ok = self._erp_ready()
        run_erp_barcode = self.chk_erp_barcode.isChecked() and erp_ok
        run_erp_add = self.chk_erp_add.isChecked() and erp_ok
        run_erp_price = self.chk_erp_price.isChecked() and erp_ok
        has_erp = any([run_erp_barcode, run_erp_add, run_erp_price])

        if not has_batch and not has_erp:
            QMessageBox.warning(self, "提示", "请至少选择一种导出内容")
            return

        # ERP功能需要3表但用户勾选了 → 提示
        if not erp_ok and any([self.chk_erp_barcode.isChecked(),
                               self.chk_erp_add.isChecked(),
                               self.chk_erp_price.isChecked()]):
            QMessageBox.warning(
                self, "提示",
                "ERP功能需要同时选择 选款资料表 + 商品信息表 + 商品库存表。\n\n"
                "当前缺失商品信息表/商品库存表，已勾选的ERP选项将跳过。"
            )

        self._last_output_dir = output_dir

        # 保存配置
        self.config['batch_erp_sel'] = sel_path
        self.config['batch_erp_info'] = info_path
        self.config['batch_erp_inv'] = inv_path
        self.config['batch_erp_output'] = output_dir
        self.config['batch_erp_min_stock'] = min_stock
        save_config(self.config)

        self.run_btn.setEnabled(False); self.run_btn.setText("生成中...")
        self.stop_btn.setEnabled(True)
        self.log_text.clear()
        self.log_text.append(f"选款资料表: {sel_path}")
        if info_path: self.log_text.append(f"商品信息表: {info_path}")
        if inv_path: self.log_text.append(f"商品库存表: {inv_path}")
        self.log_text.append(f"输出目录: {output_dir}")
        self.log_text.append(f"库存阈值: {min_stock} 件\n")
        QApplication.setOverrideCursor(Qt.WaitCursor)

        self._thread = BatchErpThread(
            sel_path, info_path, inv_path, output_dir, min_stock,
            run_batch_qa, run_batch_acc, run_batch_tryon, run_batch_attr,
            run_erp_barcode, run_erp_add, run_erp_price,
        )
        self._thread.progress_signal.connect(self._on_progress)
        self._thread.done_signal.connect(self._on_done)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

        # 计算总步数设置进度条范围
        total = 0
        if run_batch_qa: total += 1
        if run_batch_acc: total += 1
        if run_batch_tryon: total += 1
        if run_batch_attr: total += 1
        if run_erp_barcode: total += 1
        if run_erp_add: total += 1
        if run_erp_price: total += 1
        self.progress_bar.setRange(0, max(1, total))
        self.progress_bar.setFormat("启动中...")

    def _on_progress(self, step, current, total):
        self.progress_bar.setValue(min(current, self.progress_bar.maximum()))
        self.progress_bar.setFormat(f"{step}  ({current}/{self.progress_bar.maximum()})")

    def _stop_generation(self):
        if self._thread and self._thread.isRunning():
            self._thread.requestInterruption()
            self.stop_btn.setEnabled(False)
            self.log_text.append("\n正在停止...（等待当前步骤完成）")

    def _on_done(self, result):
        QApplication.restoreOverrideCursor()
        self.run_btn.setEnabled(True); self.run_btn.setText("开始生成")
        self.stop_btn.setEnabled(False)

        if isinstance(result, Exception):
            self.progress_bar.setFormat("生成失败")
            self.log_text.append(f"\n[错误] {result}")
            QMessageBox.critical(self, "生成失败", f"发生错误:\n{result}")
            return

        self.progress_bar.setFormat("完成")
        batch_result, erp_result = result

        # 汇总日志
        self.log_text.append(f"\n{'=' * 50}")
        all_warnings = []
        ok_count = 0
        error_list = []

        if batch_result:
            self.log_text.append(batch_result.summary())
            for r in [batch_result.qa, batch_result.accessories]:
                if r and r.success: ok_count += 1
                elif r: error_list.append(f"批量: {r.warnings}")
            for r in (batch_result.tryon or {}).values():
                if r.success: ok_count += 1
            for r in (batch_result.attributes or {}).values():
                if r.success: ok_count += 1
            if batch_result.warnings:
                all_warnings.extend(batch_result.warnings)

        if erp_result:
            self.log_text.append(erp_result.summary())
            for r in [erp_result.erp_barcode, erp_result.batch_add, erp_result.pricing]:
                if r and r.success: ok_count += 1
                elif r: error_list.append(f"ERP: {r.warnings}")
            if erp_result.warnings:
                all_warnings.extend(erp_result.warnings)
            if erp_result.errors:
                error_list.extend(erp_result.errors)

        # ERP报错文件
        error_file = ''
        if erp_result and erp_result.error_file:
            error_file = erp_result.error_file

        # === 严重警告优先弹窗 ===
        has_critical = any('冲突' in w or '异常' in w or '超过' in w or '已删除' in w for w in all_warnings)
        if has_critical:
            critical_msgs = [w for w in all_warnings if '冲突' in w or '异常' in w or '超过' in w or '已删除' in w]
            critical_text = "\n".join(f"  ⚠ {w}" for w in critical_msgs[:12])
            if len(critical_msgs) > 12:
                critical_text += f"\n  ... 等共 {len(critical_msgs)} 条"
            if error_file:
                critical_text += f"\n\n报错详情已输出至:\n{error_file}"

            crit_box = QMessageBox(self)
            crit_box.setWindowTitle("⚠ 数据异常警告")
            crit_box.setIcon(QMessageBox.Warning)
            crit_box.setText("生成完成，但发现以下数据异常，请务必检查！")
            crit_box.setInformativeText(critical_text)
            crit_box.setStyleSheet("QLabel{min-width:500px;}")
            crit_box.exec()

        # === 完成弹窗 ===
        msg_lines = [f"生成完成！共输出 {ok_count} 个文件。"]
        if error_file:
            msg_lines.append(f"报错文件: {os.path.basename(error_file)}")

        has_warnings = bool(all_warnings)
        has_errors = bool(error_list)

        box = QMessageBox(self)
        box.setWindowTitle("生成完成")
        icon = QMessageBox.Critical if has_errors else (QMessageBox.Warning if has_warnings else QMessageBox.Information)
        box.setIcon(icon)
        box.setText("\n".join(msg_lines))

        if has_warnings:
            detail = "\n".join(f"  - {w}" for w in all_warnings[:15])
            if len(all_warnings) > 15:
                detail += f"\n  ... 等共 {len(all_warnings)} 条"
            box.setDetailedText("警告详情:\n" + detail)

        open_btn = box.addButton("打开输出目录", QMessageBox.AcceptRole)
        err_btn = None
        if error_file and os.path.exists(error_file):
            err_btn = box.addButton("打开报错文件", QMessageBox.ActionRole)
        box.addButton("关闭", QMessageBox.RejectRole)
        box.exec()

        clicked = box.clickedButton()
        if clicked == open_btn:
            try:
                os.startfile(self._last_output_dir)
            except Exception:
                subprocess.Popen(['explorer', self._last_output_dir])
        elif err_btn and clicked == err_btn:
            try:
                os.startfile(error_file)
            except Exception:
                subprocess.Popen(['explorer', '/select,', error_file])

        if has_errors:
            QMessageBox.critical(
                self, "错误",
                "生成过程中发生错误:\n\n" + "\n".join(f"  - {e}" for e in error_list)
            )

    def _restore_paths(self):
        if self.config.get('batch_erp_sel'):
            self.sel_edit.setText(self.config['batch_erp_sel'])
        if self.config.get('batch_erp_info'):
            self.info_edit.setText(self.config['batch_erp_info'])
        if self.config.get('batch_erp_inv'):
            self.inv_edit.setText(self.config['batch_erp_inv'])
        if self.config.get('batch_erp_output'):
            self.out_edit.setText(self.config['batch_erp_output'])
        if self.config.get('batch_erp_min_stock'):
            self.stock_spin.setValue(int(self.config['batch_erp_min_stock']))
        self._refresh_erp_state()
