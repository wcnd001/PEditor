import json
import os
from datetime import datetime

from PyQt5.QtWidgets import QDialog, QDialogButtonBox, QHBoxLayout, QMessageBox, QPushButton, QPlainTextEdit, QVBoxLayout

from utils import resource_path


LOG_FILE = resource_path('modify_log.txt')


class ChangeLogger:
    @staticmethod
    def get_log_path():
        return LOG_FILE

    @staticmethod
    def _ensure_parent():
        folder = os.path.dirname(LOG_FILE)
        if folder and not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)

    @staticmethod
    def _normalize(value):
        if isinstance(value, (dict, list, tuple)):
            try:
                return json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True)
            except Exception:
                return str(value)
        if value is None:
            return 'None'
        return str(value)

    @classmethod
    def write(cls, title, before=None, after=None, extra=None):
        cls._ensure_parent()
        stamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        lines = [
            '=' * 80,
            '[%s] %s' % (stamp, title),
        ]
        if extra not in (None, ''):
            lines.append('附加信息:')
            lines.append(cls._normalize(extra))
        if before is not None:
            lines.append('修改前:')
            lines.append(cls._normalize(before))
        if after is not None:
            lines.append('修改后:')
            lines.append(cls._normalize(after))
        lines.append('')
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write('\n'.join(lines))

    @classmethod
    def read_text(cls):
        cls._ensure_parent()
        if not os.path.exists(LOG_FILE):
            return '暂无日志。'
        try:
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
                text = f.read()
            return text if text.strip() else '暂无日志。'
        except Exception as e:
            return '读取日志失败：%s' % e

    @classmethod
    def clear(cls):
        cls._ensure_parent()
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write('')


def log_change(title, before=None, after=None, extra=None):
    try:
        ChangeLogger.write(title, before=before, after=after, extra=extra)
    except Exception:
        pass


class LogViewerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('修改日志')
        self.resize(920, 680)
        self.init_ui()
        self.refresh_text()

    def init_ui(self):
        layout = QVBoxLayout(self)
        self.text_edit = QPlainTextEdit()
        self.text_edit.setReadOnly(True)
        layout.addWidget(self.text_edit)

        btn_row = QHBoxLayout()
        self.refresh_btn = QPushButton('刷新')
        self.refresh_btn.clicked.connect(self.refresh_text)
        btn_row.addWidget(self.refresh_btn)
        self.clear_btn = QPushButton('清空日志')
        self.clear_btn.clicked.connect(self.clear_log)
        btn_row.addWidget(self.clear_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def refresh_text(self):
        self.text_edit.setPlainText(ChangeLogger.read_text())
        cursor = self.text_edit.textCursor()
        cursor.movePosition(cursor.End)
        self.text_edit.setTextCursor(cursor)

    def clear_log(self):
        reply = QMessageBox.question(self, '确认', '确定清空修改日志吗？', QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        ChangeLogger.clear()
        self.refresh_text()
