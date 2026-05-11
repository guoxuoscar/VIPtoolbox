# -*- coding: utf-8 -*-
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
)

from toolbox.core.utils import (
    load_anti_ban_config,
    load_proxies,
    save_anti_ban_config,
)


class AntiBanSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🛡️ 防封设置")
        self.setMinimumWidth(500)
        self.setStyleSheet("QDialog { background-color: white; }")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        self.setLayout(layout)

        title = QLabel("防封策略设置")
        title.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        title.setStyleSheet("color: #1565C0;")
        layout.addWidget(title)

        info = QLabel("💡 建议开启所有防封选项，下载速度会稍慢但更稳定")
        info.setFont(QFont("Microsoft YaHei", 9))
        info.setStyleSheet("color: #666; background-color: #E3F2FD; padding: 8px; border-radius: 4px;")
        layout.addWidget(info)

        config = load_anti_ban_config()

        self.enable_check = QCheckBox("启用防封策略")
        self.enable_check.setChecked(config.get("enabled", True))
        self.enable_check.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        layout.addWidget(self.enable_check)

        delay_group = QGroupBox("延迟设置")
        delay_group.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        delay_layout = QGridLayout()
        delay_layout.addWidget(QLabel("最小延迟:"), 0, 0)
        self.min_delay = QSpinBox()
        self.min_delay.setRange(1, 30)
        self.min_delay.setValue(config.get("min_delay", 3))
        self.min_delay.setSuffix(" 秒")
        delay_layout.addWidget(self.min_delay, 0, 1)
        delay_layout.addWidget(QLabel("最大延迟:"), 0, 2)
        self.max_delay = QSpinBox()
        self.max_delay.setRange(2, 60)
        self.max_delay.setValue(config.get("max_delay", 8))
        self.max_delay.setSuffix(" 秒")
        delay_layout.addWidget(self.max_delay, 0, 3)
        delay_layout.addWidget(QLabel("批次大小:"), 1, 0)
        self.batch_size = QSpinBox()
        self.batch_size.setRange(1, 20)
        self.batch_size.setValue(config.get("batch_size", 5))
        self.batch_size.setSuffix(" 个")
        delay_layout.addWidget(self.batch_size, 1, 1)
        delay_layout.addWidget(QLabel("休息时间:"), 1, 2)
        rest_layout = QHBoxLayout()
        self.batch_rest_min = QSpinBox()
        self.batch_rest_min.setRange(10, 300)
        self.batch_rest_min.setValue(config.get("batch_rest_min", 30))
        self.batch_rest_min.setSuffix("秒")
        rest_layout.addWidget(self.batch_rest_min)
        rest_layout.addWidget(QLabel("-"))
        self.batch_rest_max = QSpinBox()
        self.batch_rest_max.setRange(20, 600)
        self.batch_rest_max.setValue(config.get("batch_rest_max", 60))
        self.batch_rest_max.setSuffix("秒")
        rest_layout.addWidget(self.batch_rest_max)
        delay_layout.addLayout(rest_layout, 1, 3)
        delay_group.setLayout(delay_layout)
        layout.addWidget(delay_group)

        proxy_group = QGroupBox("代理设置")
        proxy_group.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        proxy_layout = QVBoxLayout()
        self.proxy_check = QCheckBox("启用代理轮换")
        self.proxy_check.setChecked(config.get("use_proxy", False))
        proxy_layout.addWidget(self.proxy_check)
        proxies = load_proxies()
        proxy_label = QLabel(f"已加载代理: {len(proxies)} 个")
        proxy_label.setFont(QFont("Microsoft YaHei", 9))
        proxy_label.setStyleSheet("color: #666;")
        proxy_layout.addWidget(proxy_label)
        proxy_note = QLabel("代理格式: ip:port 或 http://ip:port\n在项目目录的 proxies.txt 中配置")
        proxy_note.setFont(QFont("Microsoft YaHei", 8))
        proxy_note.setStyleSheet("color: #999;")
        proxy_layout.addWidget(proxy_note)
        proxy_group.setLayout(proxy_layout)
        layout.addWidget(proxy_group)

        other_group = QGroupBox("其他选项")
        other_group.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        other_layout = QVBoxLayout()
        self.random_ua = QCheckBox("随机User-Agent (轮换浏览器标识)")
        self.random_ua.setChecked(config.get("random_ua", True))
        other_layout.addWidget(self.random_ua)
        self.human_scroll = QCheckBox("人类滚动行为 (模拟真人操作)")
        self.human_scroll.setChecked(config.get("human_scroll", True))
        other_layout.addWidget(self.human_scroll)
        self.auto_retry = QCheckBox("自动重试失败项 (最多3次)")
        self.auto_retry.setChecked(config.get("auto_retry", True))
        other_layout.addWidget(self.auto_retry)
        other_group.setLayout(other_layout)
        layout.addWidget(other_group)
        layout.addStretch()

        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("✅ 确定保存")
        ok_btn.setFont(QFont("Microsoft YaHei", 11, QFont.Bold))
        ok_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; border: none; padding: 10px 25px; border-radius: 4px; }")
        ok_btn.clicked.connect(self.save_settings)
        btn_layout.addWidget(ok_btn)
        cancel_btn = QPushButton("取消")
        cancel_btn.setFont(QFont("Microsoft YaHei", 10))
        cancel_btn.setStyleSheet("QPushButton { background-color: #9E9E9E; color: white; border: none; padding: 10px 20px; border-radius: 4px; }")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def save_settings(self):
        new_config = {
            "enabled": self.enable_check.isChecked(),
            "min_delay": self.min_delay.value(),
            "max_delay": self.max_delay.value(),
            "batch_size": self.batch_size.value(),
            "batch_rest_min": self.batch_rest_min.value(),
            "batch_rest_max": self.batch_rest_max.value(),
            "use_proxy": self.proxy_check.isChecked(),
            "random_ua": self.random_ua.isChecked(),
            "human_scroll": self.human_scroll.isChecked(),
            "auto_retry": self.auto_retry.isChecked(),
        }
        save_anti_ban_config(new_config)
        QMessageBox.information(self, "保存成功", "防封设置已保存。\n代理与 User-Agent 将在下次点击「打开登录页面」时生效。")
        self.accept()
