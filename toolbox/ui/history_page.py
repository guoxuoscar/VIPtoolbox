# -*- coding: utf-8 -*-

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from toolbox.core.utils import save_config


class HistoryPage(QWidget):
    """操作历史页面（从 main_window 拆分，便于维护）"""

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        self.setLayout(layout)

        title = QLabel("操作历史")
        title.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        title.setStyleSheet("color: #1565C0;")
        layout.addWidget(title)

        toolbar = QHBoxLayout()
        self.clear_btn = QPushButton("🗑 清空历史")
        self.clear_btn.setFont(QFont("Microsoft YaHei", 10))
        self.clear_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #F44336;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            """
        )
        self.clear_btn.clicked.connect(self._clear_history)
        toolbar.addWidget(self.clear_btn)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        self.history_group = QGroupBox("📋 操作记录")
        self.history_group.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        layout.addWidget(self.history_group, 1)
        self._render_history()

    def _render_history(self):
        # 先清空旧内容，再渲染新内容
        old_layout = self.history_group.layout()
        if old_layout is not None:
            self._clear_layout(old_layout)
            old_layout.deleteLater()

        history_layout = QVBoxLayout()
        history = self.config.get("operation_history", [])

        if not history:
            no_data = QLabel("暂无操作记录")
            no_data.setStyleSheet("color: #999; font-size: 14px; padding: 20px;")
            no_data.setAlignment(Qt.AlignCenter)
            history_layout.addWidget(no_data)
        else:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)

            container = QWidget()
            container_layout = QVBoxLayout()

            for item in history[:50]:
                item_widget = QFrame()
                item_widget.setStyleSheet(
                    "background-color: #F5F5F5; padding: 10px; border-radius: 4px; margin-bottom: 5px;"
                )
                item_layout = QVBoxLayout()
                item_widget.setLayout(item_layout)

                time_label = QLabel(item.get("time", ""))
                time_label.setStyleSheet("color: #1565C0; font-weight: bold;")
                item_layout.addWidget(time_label)

                op_label = QLabel(item.get("operation", ""))
                op_label.setStyleSheet("color: #333; font-weight: bold;")
                item_layout.addWidget(op_label)

                detail_label = QLabel(item.get("details", ""))
                detail_label.setStyleSheet("color: #666;")
                detail_label.setWordWrap(True)
                item_layout.addWidget(detail_label)

                container_layout.addWidget(item_widget)

            container.setLayout(container_layout)
            scroll.setWidget(container)
            history_layout.addWidget(scroll)

        self.history_group.setLayout(history_layout)

    def _clear_history(self):
        if not self.config.get("operation_history"):
            QMessageBox.information(self, "提示", "当前没有可清空的历史")
            return
        reply = QMessageBox.question(self, "确认", "确定要清空全部操作历史吗？")
        if reply != QMessageBox.Yes:
            return
        self.config["operation_history"] = []
        save_config(self.config)
        QMessageBox.information(self, "完成", "操作历史已清空")
        self._render_history()

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            child_layout = item.layout()
            if widget is not None:
                widget.deleteLater()
            elif child_layout is not None:
                self._clear_layout(child_layout)
