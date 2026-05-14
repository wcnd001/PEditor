from contextlib import contextmanager

from PyQt5.QtCore import Qt, QSignalBlocker


@contextmanager
def signal_blocked(obj):
    """异常安全地临时阻断 QObject 信号。"""
    blocker = QSignalBlocker(obj)
    try:
        yield obj
    finally:
        del blocker


def wrap_text_flags():
    """文本框自动换行用绘制标志。下拉菜单已恢复原生样式，不再使用自定义下拉换行代理。"""
    flags = Qt.TextWordWrap
    try:
        flags = flags | Qt.TextWrapAnywhere
    except Exception:
        pass
    return flags
