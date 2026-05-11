import os

from PySide6.QtCore import QEvent, QObject, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QLineEdit


class PathDropFilter(QObject):
    """给路径输入框提供拖拽支持（文件/目录/混合）。"""

    def __init__(self, mode="file", extensions=None, multi=False, on_accept=None, parent=None):
        super().__init__(parent)
        self.mode = mode
        self.extensions = {e.lower() for e in (extensions or [])}
        self.multi = multi
        self.on_accept = on_accept

    def _valid_file_ext(self, path):
        if not self.extensions:
            return True
        ext = os.path.splitext(path)[1].lower()
        return ext in self.extensions

    def _is_valid_path(self, path):
        mode = self.mode
        if mode in ("dir", "folder"):
            return os.path.isdir(path)
        if mode in ("both", "file_or_dir"):
            if os.path.isdir(path):
                return True
            return os.path.isfile(path) and self._valid_file_ext(path)
        return os.path.isfile(path) and self._valid_file_ext(path)

    def _collect_valid(self, event):
        md = event.mimeData()
        if not md or not md.hasUrls():
            return []
        valid = []
        for url in md.urls():
            if not url.isLocalFile():
                continue
            p = url.toLocalFile()
            if self._is_valid_path(p):
                valid.append(p)
        return valid

    def eventFilter(self, obj, event):
        et = event.type()
        if et == QEvent.DragEnter:
            if self._collect_valid(event):
                event.acceptProposedAction()
            else:
                event.ignore()
            return True
        if et == QEvent.Drop:
            valid = self._collect_valid(event)
            if not valid:
                event.ignore()
                return True
            value = ";".join(valid) if self.multi else valid[0]
            obj.setText(value)
            if self.on_accept:
                try:
                    self.on_accept(value)
                except Exception:
                    pass
            event.acceptProposedAction()
            return True
        return False


def enable_path_drop(line_edit, mode="file", extensions=None, multi=False, on_accept=None):
    """启用路径输入框拖拽。"""
    line_edit.setAcceptDrops(True)
    drop_filter = PathDropFilter(mode=mode, extensions=extensions, multi=multi, on_accept=on_accept, parent=line_edit)
    line_edit.installEventFilter(drop_filter)
    line_edit._path_drop_filter = drop_filter


class ExcelDropLineEdit(QLineEdit):
    """仅接受 Excel 文件拖拽的输入框。"""
    excel_dropped = Signal(str)

    def __init__(self, placeholder="", parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFont(QFont("Microsoft YaHei", 10))
        self.setPlaceholderText(placeholder)
        self._btn_style = """
            QLineEdit {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                padding: 8px 10px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #1E88E5;
            }
        """
        self.setStyleSheet(self._btn_style)

    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md.hasUrls():
            for url in md.urls():
                if not url.isLocalFile():
                    continue
                path = url.toLocalFile()
                if path.lower().endswith((".xlsx", ".xls")):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        md = event.mimeData()
        if not md.hasUrls():
            event.ignore()
            return
        for url in md.urls():
            if not url.isLocalFile():
                continue
            path = url.toLocalFile()
            if path.lower().endswith((".xlsx", ".xls")):
                self.setText(path)
                self.excel_dropped.emit(path)
                event.acceptProposedAction()
                return
        event.ignore()


class DirDropLineEdit(QLineEdit):
    """仅接受目录拖拽的输入框。"""
    dir_dropped = Signal(str)

    def __init__(self, placeholder="", parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFont(QFont("Microsoft YaHei", 10))
        self.setPlaceholderText(placeholder)
        self.setStyleSheet("""
            QLineEdit {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                padding: 8px 10px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #1E88E5;
            }
        """)

    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md.hasUrls():
            for url in md.urls():
                if url.isLocalFile() and os.path.isdir(url.toLocalFile()):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        md = event.mimeData()
        if not md.hasUrls():
            event.ignore()
            return
        for url in md.urls():
            if not url.isLocalFile():
                continue
            path = url.toLocalFile()
            if os.path.isdir(path):
                self.setText(path)
                self.dir_dropped.emit(path)
                event.acceptProposedAction()
                return
        event.ignore()
