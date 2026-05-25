import sys
import json
import os
import traceback
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QComboBox, QPushButton, QTextEdit,
    QMessageBox, QFileDialog, QLabel, QInputDialog, QGridLayout,
    QCheckBox, QDialog, QListWidget, QListWidgetItem, QDialogButtonBox,
    QGroupBox, QAbstractItemView, QScrollArea, QButtonGroup, QMenu,
    QAction, QSplitter, QFormLayout, QSpinBox, QFrame, QToolBar,
    QToolButton, QSizePolicy,
    QPlainTextEdit, QTableView, QTableWidget, QDoubleSpinBox, QAbstractSpinBox,
    QStyleOptionComboBox, QStyle, QColorDialog
)
from PyQt5.QtCore import Qt, QTimer, QEvent, QSize
from PyQt5.QtGui import QFont, QFontMetrics, QTextOption, QPalette, QColor
from dbutils import Database
from datamanager import DataManagerWindow
from template_editor import TemplateEditorWindow
from template_db import TemplateDB
from datamatch import DataMatcher, RuleManagerDialog
import export
from webcontrol import BrowserFlowWindow
from utils import resource_path
from log import LogViewerDialog, log_change
from gui_helpers import signal_blocked, wrap_text_flags as _wrap_text_flags

__version__ = '3.4'
# 主界面输入区的最小布局宽度。左侧分割区域小于此值时，行内控件按该宽度稳定布局，避免长文本反复重排导致卡顿。
input_option_min_layout_width = 360
# 打包命令：pyinstaller --clean PEditor.spec --distpath "D:\Microsoft Visual Studio\code"



# 界面主题预设和颜色字段。颜色统一使用 #RRGGBB，便于写入 settings.json。
# QColor 是 PyQt5 提供的颜色类，用于校验颜色字符串并生成调色板。
theme_color_fields = [
    ('color_window', '窗口背景'),
    ('color_panel', '面板/卡片背景'),
    ('color_text', '普通文字'),
    ('color_input_bg', '输入框背景'),
    ('color_button_bg', '按钮背景'),
    ('color_button_text', '按钮文字'),
    ('color_border', '边框颜色'),
    ('color_primary', '选中主色'),
    ('color_selected_bg', '选中背景'),
    ('color_handle_bg', '拖动柄背景'),
    ('color_handle_selected_bg', '选中拖动柄背景'),
    ('color_preview_bg', '预览框背景'),
]

theme_presets = {
    '浅色': {
        'color_window': '#f3f4f6',
        'color_panel': '#ffffff',
        'color_text': '#111827',
        'color_input_bg': '#ffffff',
        'color_button_bg': '#f3f4f6',
        'color_button_text': '#111827',
        'color_border': '#d1d5db',
        'color_primary': '#2563eb',
        'color_selected_bg': '#dbeafe',
        'color_handle_bg': '#f3f4f6',
        'color_handle_selected_bg': '#bfdbfe',
        'color_preview_bg': '#ffffff',
    },
    '深色': {
        'color_window': '#111827',
        'color_panel': '#1f2937',
        'color_text': '#f9fafb',
        'color_input_bg': '#374151',
        'color_button_bg': '#374151',
        'color_button_text': '#f9fafb',
        'color_border': '#4b5563',
        'color_primary': '#60a5fa',
        'color_selected_bg': '#1e3a8a',
        'color_handle_bg': '#374151',
        'color_handle_selected_bg': '#1d4ed8',
        'color_preview_bg': '#111827',
    },
    '护眼': {
        'color_window': '#eef6e8',
        'color_panel': '#f8fff2',
        'color_text': '#1f2933',
        'color_input_bg': '#ffffff',
        'color_button_bg': '#e8f3dc',
        'color_button_text': '#1f2933',
        'color_border': '#b7c8a6',
        'color_primary': '#4d7c0f',
        'color_selected_bg': '#d9f99d',
        'color_handle_bg': '#e8f3dc',
        'color_handle_selected_bg': '#bbf7d0',
        'color_preview_bg': '#fbfff5',
    },
}


def normalize_color_value(value, default='#ffffff'):
    """校验并规范化颜色字符串，非法颜色回退到 default。"""
    color = QColor(str(value or '').strip())
    if color.isValid():
        return color.name()
    fallback = QColor(str(default or '').strip())
    return fallback.name() if fallback.isValid() else '#ffffff'


def resolve_theme_colors(settings: dict):
    """根据设置生成实际使用的主题颜色。"""
    settings = settings if isinstance(settings, dict) else {}
    theme_name = str(settings.get('theme_name') or '浅色')
    preset = theme_presets.get(theme_name, theme_presets['浅色'])
    colors = {}
    for key, _label in theme_color_fields:
        default = preset.get(key, theme_presets['浅色'].get(key, '#ffffff'))
        colors[key] = normalize_color_value(settings.get(key, default), default)
    return colors


class ToolbarMenuButton(QToolButton):
    """工具栏菜单按钮：使用原生菜单三角，并缓存稳定的 sizeHint。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self._stable_size_sig = None
        self._stable_size_hint = QSize()

    def stable_size_hint(self):
        sig = (
            self.text(),
            self.font().family(),
            round(float(self.font().pointSizeF()), 2),
            type(self.style()).__name__,
            int(self.popupMode()),
            bool(self.menu()),
        )
        if sig != self._stable_size_sig or not self._stable_size_hint.isValid():
            try:
                self.ensurePolished()
                hint = super().sizeHint()
                min_hint = super().minimumSizeHint()
                self._stable_size_hint = QSize(max(hint.width(), min_hint.width()), max(hint.height(), min_hint.height()))
            except Exception:
                self._stable_size_hint = QSize(max(56, len(self.text()) * 14), 28)
            self._stable_size_sig = sig
        return QSize(self._stable_size_hint)

    def invalidate_stable_size(self):
        self._stable_size_sig = None

class OptionEditDialog(QDialog):
    def __init__(self, options_config: list, main_db: Database, parent=None):
        super().__init__(parent)
        self.main_db = main_db
        self.original_config = json.loads(json.dumps(options_config, ensure_ascii=False))
        self.options_config = json.loads(json.dumps(options_config, ensure_ascii=False))
        self.setWindowTitle('编辑输入选项')
        self.resize(650, 550)
        self.init_ui()
        self.load_options()

    def init_ui(self):
        layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        self.list_widget.currentItemChanged.connect(self.on_current_changed)
        layout.addWidget(QLabel('选项列表（可拖拽排序）：'))
        layout.addWidget(self.list_widget)

        edit_group = QGroupBox('编辑选中选项')
        edit_layout = QFormLayout(edit_group)
        self.label_edit = QLineEdit()
        edit_layout.addRow('标签:', self.label_edit)

        self.type_combo = QComboBox()
        self.type_combo.addItems(['文本框', '下拉菜单', '可输入下拉菜单', '复选框'])
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        edit_layout.addRow('类型:', self.type_combo)

        self.source_group = QGroupBox('数据源配置')
        source_layout = QFormLayout(self.source_group)
        self.source_table_combo = QComboBox()
        self.source_table_combo.addItem('（固定选项）', None)
        for t in self.main_db.get_tables():
            self.source_table_combo.addItem(t, t)
        self.source_table_combo.currentIndexChanged.connect(self.on_table_changed)
        source_layout.addRow('数据表:', self.source_table_combo)

        self.source_column_combo = QComboBox()
        source_layout.addRow('显示列:', self.source_column_combo)

        self.fixed_values_edit = QTextEdit()
        self.fixed_values_edit.setPlaceholderText('每行一个选项值；支持 {字段名} 与 #公式#，公式可返回多行并自动拆成多个下拉项')
        source_layout.addRow('固定选项:', self.fixed_values_edit)
        self.fixed_values_hint = QLabel('固定选项支持函数：例如 #dbjoin(\'工序表\', \'产品型号\', {产品型号}, \'{工序内容}\', nl())#\n若公式返回多行，程序会自动拆成多个下拉选项；dbrows(...) 也可继续使用。')
        self.fixed_values_hint.setWordWrap(True)
        self.fixed_values_hint.setStyleSheet('color: #555;')
        source_layout.addRow(self.fixed_values_hint)
        edit_layout.addRow(self.source_group)

        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton('新增')
        self.add_btn.clicked.connect(self.add_option)
        btn_layout.addWidget(self.add_btn)
        self.delete_btn = QPushButton('删除')
        self.delete_btn.clicked.connect(self.delete_option)
        btn_layout.addWidget(self.delete_btn)
        self.apply_btn = QPushButton('应用修改')
        self.apply_btn.clicked.connect(self.apply_changes)
        btn_layout.addWidget(self.apply_btn)
        edit_layout.addRow(btn_layout)
        layout.addWidget(edit_group)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.on_type_changed(self.type_combo.currentText())

    def load_options(self):
        self.list_widget.clear()
        for opt in sorted(self.options_config, key=lambda x: x.get('order', 0)):
            item = QListWidgetItem(opt['label'])
            item.setData(Qt.UserRole, opt)
            self.list_widget.addItem(item)
        if self.list_widget.count():
            self.list_widget.setCurrentRow(0)
        else:
            self.clear_editor()

    def clear_editor(self):
        self.label_edit.clear()
        self.type_combo.setCurrentIndex(0)
        self.source_table_combo.setCurrentIndex(0)
        self.source_column_combo.clear()
        self.fixed_values_edit.clear()

    def on_current_changed(self, current, previous):
        if not current:
            self.clear_editor()
            return

        opt = current.data(Qt.UserRole)
        self.label_edit.setText(opt.get('label', ''))
        type_map = {'文本框': 'text', '下拉菜单': 'combo', '可输入下拉菜单': 'editable_combo', '复选框': 'checkbox'}
        rev = {v: k for k, v in type_map.items()}
        self.type_combo.setCurrentText(rev.get(opt.get('type', 'text'), '文本框'))

        source = opt.get('source', {}) or {}
        if source.get('type') == 'table':
            idx = self.source_table_combo.findData(source.get('table'))
            if idx >= 0:
                self.source_table_combo.setCurrentIndex(idx)
            self.on_table_changed()
            idx = self.source_column_combo.findText(source.get('column', ''))
            if idx >= 0:
                self.source_column_combo.setCurrentIndex(idx)
            self.fixed_values_edit.clear()
        else:
            self.source_table_combo.setCurrentIndex(0)
            self.on_table_changed()
            self.fixed_values_edit.setPlainText('\n'.join(source.get('values', [])))

    def on_type_changed(self, text):
        self.source_group.setVisible(text in ('下拉菜单', '可输入下拉菜单'))

    def on_table_changed(self):
        self.source_column_combo.clear()
        table = self.source_table_combo.currentData()
        if table:
            for col in self.main_db.get_table_info(table):
                self.source_column_combo.addItem(col['name'])
        else:
            self.source_column_combo.addItem('')

    def add_option(self):
        opt = {'label': '新选项', 'type': 'text', 'order': len(self.options_config), 'source': {}}
        self.options_config.append(opt)
        item = QListWidgetItem(opt['label'])
        item.setData(Qt.UserRole, opt)
        self.list_widget.addItem(item)
        self.list_widget.setCurrentItem(item)

    def delete_option(self):
        item = self.list_widget.currentItem()
        if item:
            opt = item.data(Qt.UserRole)
            if opt in self.options_config:
                self.options_config.remove(opt)
            self.list_widget.takeItem(self.list_widget.row(item))

    def _label_exists_elsewhere(self, label: str, current_item=None) -> bool:
        target = str(label or '').strip()
        if not target:
            return False
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item is None or item is current_item:
                continue
            if str(item.text() or '').strip() == target:
                return True
        return False

    def apply_changes(self):
        item = self.list_widget.currentItem()
        if not item:
            return True
        opt = item.data(Qt.UserRole)
        label = self.label_edit.text().strip()
        original_label = str(opt.get('label', '') or '').strip()
        if not label:
            QMessageBox.warning(self, '错误', '标签不能为空')
            return False
        # 当前项未改名时，直接视为合法，避免误把自身计入重复项。
        if label != original_label and self._label_exists_elsewhere(label, current_item=item):
            QMessageBox.warning(self, '错误', f'标签“{label}”已存在，请选择其他名称')
            return False
        opt['label'] = label
        type_map = {'文本框': 'text', '下拉菜单': 'combo', '可输入下拉菜单': 'editable_combo', '复选框': 'checkbox'}
        opt['type'] = type_map[self.type_combo.currentText()]

        source = {}
        if self.type_combo.currentText() in ('下拉菜单', '可输入下拉菜单'):
            table = self.source_table_combo.currentData()
            if table:
                source = {'type': 'table', 'table': table, 'column': self.source_column_combo.currentText()}
            else:
                vals = [v.strip() for v in self.fixed_values_edit.toPlainText().split('\n') if v.strip()]
                source = {'type': 'fixed', 'values': vals}
        opt['source'] = source
        item.setText(opt['label'])
        item.setData(Qt.UserRole, opt)
        return True

    def accept(self):
        if self.list_widget.currentItem() and not self.apply_changes():
            return

        final = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            opt = item.data(Qt.UserRole)
            opt['order'] = i
            final.append(opt)
        self.options_config = final
        super().accept()

    def reject(self):
        current = self._get_current_config()
        if current != self.original_config:
            reply = QMessageBox.question(self, '未保存', '是否保存更改？', QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Yes:
                self.accept()
            elif reply == QMessageBox.No:
                super().reject()
        else:
            super().reject()

    def _get_current_config(self):
        temp = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            opt = item.data(Qt.UserRole).copy()
            opt['order'] = i
            temp.append(opt)
        return temp


def scaled_point_size(settings: dict, key: str, default: int) -> int:
    """根据总字体缩放比例计算实际 QFont 点数。

    字体大小统一只通过 QFont.setPointSize 控制，样式表不再写 font-size，
    避免 pt 与 px 混用导致“调大反而变小”。
    """
    try:
        base_size = int((settings or {}).get(key, default))
    except Exception:
        base_size = int(default)
    try:
        scale = int((settings or {}).get('font_scale', 100))
    except Exception:
        scale = 100
    scale = max(30, min(300, scale))
    return max(1, int(round(base_size * scale / 100.0)))


def set_widget_point_size(widget, point_size: int):
    """安全设置控件字号。"""
    if widget is None:
        return
    try:
        font = widget.font() if hasattr(widget, 'font') else QFont()
        font.setPointSize(int(point_size))
        widget.setFont(font)
        if isinstance(widget, QTextEdit):
            widget.document().setDefaultFont(font)
            viewport = widget.viewport()
            if viewport is not None:
                viewport.setFont(font)
        elif isinstance(widget, QComboBox):
            try:
                view = widget.view()
                if view is not None:
                    view.setFont(font)
            except Exception:
                pass
            try:
                line_edit = widget.lineEdit()
                if line_edit is not None:
                    line_edit.setFont(font)
            except Exception:
                pass
    except RuntimeError:
        return
    except Exception:
        return


def set_text_edit_font(text_edit: QTextEdit, font: QFont):
    """同步 QTextEdit 控件、viewport 和 document 的字体，避免输入区文字比其他控件小。"""
    try:
        text_edit.setFont(font)
        text_edit.document().setDefaultFont(font)
        viewport = text_edit.viewport()
        if viewport is not None:
            viewport.setFont(font)
    except Exception:
        pass


def font_metrics_for(widget_or_font):
    """返回 QFontMetrics，用于根据字体计算控件自适应尺寸。"""
    try:
        if isinstance(widget_or_font, QFont):
            return QFontMetrics(widget_or_font)
        return QFontMetrics(widget_or_font.font())
    except Exception:
        return QFontMetrics(QFont())


def text_width_for(widget_or_font, text: str) -> int:
    """兼容不同 PyQt5 版本的文字宽度计算。"""
    metrics = font_metrics_for(widget_or_font)
    text = '' if text is None else str(text)
    try:
        return metrics.horizontalAdvance(text)
    except AttributeError:
        return metrics.width(text)


def adaptive_control_height(widget_or_font, compactness: int, min_height: int = 22, lines: int = 1) -> int:
    """根据字体行高和紧凑程度计算控件高度。"""
    metrics = font_metrics_for(widget_or_font)
    try:
        line_height = metrics.lineSpacing()
    except Exception:
        line_height = 16
    try:
        compactness = int(compactness)
    except Exception:
        compactness = 4
    return max(int(min_height), int(line_height * max(1, lines) + compactness * 2 + 8))


def adaptive_text_width(widget_or_font, text: str, compactness: int, min_width: int = 40, max_width: int = 300) -> int:
    """根据文字内容和字体计算建议宽度。"""
    try:
        compactness = int(compactness)
    except Exception:
        compactness = 4
    width = text_width_for(widget_or_font, text) + compactness * 2 + 20
    return max(int(min_width), min(int(max_width), int(width)))




def qt_safe_single_shot(delay_ms, callback):
    """安全延后执行 PyQt 回调。

    QTimer.singleShot 使用 Python callable 时，如果控件在定时器触发前
    已经被删除，回调里访问 Qt C++ 对象会抛出 RuntimeError。
    这里统一吞掉这类删除后回调，避免窗口关闭/重建输入项时闪退。
    """
    def runner():
        try:
            callback()
        except RuntimeError:
            return
        except Exception:
            return
    try:
        QTimer.singleShot(int(delay_ms), runner)
    except RuntimeError:
        return
    except Exception:
        return


def wrap_text_by_width(text: str, widget_or_font, max_width: int, max_lines: int = 3) -> str:
    """把固定宽度按钮文本按当前字体插入换行，避免字号变大后文字被裁切。"""
    text = '' if text is None else str(text)
    if not text:
        return ''
    max_width = max(20, int(max_width or 20))
    metrics = font_metrics_for(widget_or_font)
    lines = []
    current = ''
    for ch in text:
        if ch == '\n':
            lines.append(current)
            current = ''
            continue
        candidate = current + ch
        try:
            width = metrics.horizontalAdvance(candidate)
        except AttributeError:
            width = metrics.width(candidate)
        if current and width > max_width:
            lines.append(current)
            current = ch
            if len(lines) >= max_lines:
                break
        else:
            current = candidate
    if current and len(lines) < max_lines:
        lines.append(current)
    if len(lines) >= max_lines:
        consumed = ''.join(lines)
        if len(consumed) < len(text.replace('\n', '')):
            lines[-1] = lines[-1].rstrip('…') + '…'
    return '\n'.join(lines)




class AutoWrapTextEdit(QTextEdit):
    """用于主界面输入项的自动换行文本框。

    初始高度保持为单行控件高度。只有文本按当前真实输入框宽度排版后
    所需高度超过单行基准高度时，才逐步增高；达到设置的最大高度后
    再启用内部垂直滚动条。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._single_line_height = 24
        self._max_auto_height = 96
        self._height_updating = False
        self._available_text_width = 1
        self.setAcceptRichText(False)
        self.setLineWrapMode(QTextEdit.WidgetWidth)
        self.setWordWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        self.setTabChangesFocus(True)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        try:
            self.document().setDocumentMargin(2)
        except Exception:
            pass
        self.document().contentsChanged.connect(self.update_auto_height)
        self.set_auto_limits(24, 96)

    def set_auto_limits(self, single_line_height, max_height):
        try:
            self._single_line_height = max(20, int(single_line_height))
        except Exception:
            self._single_line_height = 24
        try:
            self._max_auto_height = max(self._single_line_height, int(max_height))
        except Exception:
            self._max_auto_height = max(self._single_line_height, 96)
        self.setMinimumHeight(self._single_line_height)
        self.setMaximumHeight(self._max_auto_height)
        if self.height() <= 0 or self.toPlainText().strip() == '':
            self.setFixedHeight(self._single_line_height)
        self._sync_document_text_width()
        self.update_auto_height()

    def set_available_text_width(self, width):
        """由外部行布局传入本轮真实文本可用宽度。"""
        try:
            width = int(width)
        except Exception:
            width = 1
        self._available_text_width = max(1, width)
        self._sync_document_text_width()
        self.update_auto_height()

    def _fallback_text_width(self):
        try:
            frame = self.frameWidth() * 2
        except Exception:
            frame = 0
        return max(1, int(self.width()) - frame - 6)

    def _current_text_width(self):
        width = max(1, int(getattr(self, '_available_text_width', 1) or 1))
        if width <= 1:
            width = self._fallback_text_width()
        return max(1, int(width))

    def _sync_document_text_width(self):
        """把 QTextDocument 的排版宽度同步为本轮真实输入框宽度。

        这里不能固定为 -1。-1 会让 QTextDocument 按理想自然宽度排版，
        在输入框较宽或全屏时连续长文本可能不按控件宽度换行。
        每次重算都使用当前传入宽度，避免旧宽度残留。
        """
        try:
            self.document().setTextWidth(self._current_text_width())
        except Exception:
            pass

    def setFont(self, font):
        super().setFont(font)
        try:
            self.document().setDefaultFont(font)
            viewport = self.viewport()
            if viewport is not None:
                viewport.setFont(font)
        except Exception:
            pass
        self._sync_document_text_width()
        self.update_auto_height()

    def setPlainTextPreserveSignal(self, text):
        self.setPlainText('' if text is None else str(text))
        self._sync_document_text_width()
        self.update_auto_height()

    def _height_components(self, text_height):
        try:
            frame = self.frameWidth() * 2
        except Exception:
            frame = 0
        margins = self.contentsMargins()
        try:
            doc_margin = int(self.document().documentMargin())
        except Exception:
            doc_margin = 2
        return int(text_height) + margins.top() + margins.bottom() + frame + doc_margin * 2 + 4

    def _document_required_height(self):
        text = self.toPlainText()
        if text in (None, ''):
            return self._single_line_height
        width = self._current_text_width()
        metrics = QFontMetrics(self.font())
        flags = _wrap_text_flags()
        try:
            rect = metrics.boundingRect(0, 0, int(width), 100000, flags, text)
            text_height = rect.height()
        except Exception:
            explicit_lines = str(text).count('\n') + 1
            text_height = metrics.lineSpacing() * max(1, explicit_lines)
        return max(self._single_line_height, self._height_components(text_height))

    def _single_line_required_height(self):
        metrics = QFontMetrics(self.font())
        try:
            text_height = metrics.boundingRect(0, 0, max(80, self._current_text_width()), 1000, _wrap_text_flags(), 'Hg').height()
        except Exception:
            text_height = metrics.lineSpacing()
        return max(self._single_line_height, self._height_components(text_height))

    def update_auto_height(self):
        if getattr(self, '_height_updating', False):
            return
        self._height_updating = True
        try:
            self._sync_document_text_width()
            wanted = self._document_required_height()
            baseline = self._single_line_required_height()
            if wanted <= baseline + 2:
                height = self._single_line_height
            else:
                height = min(self._max_auto_height, wanted)
            height = max(self._single_line_height, int(height))
            self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded if wanted > self._max_auto_height else Qt.ScrollBarAlwaysOff)
            if self.height() != height:
                self.setFixedHeight(height)
            self.updateGeometry()
        except Exception:
            pass
        finally:
            self._height_updating = False

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._sync_document_text_width()
        qt_safe_single_shot(0, self.update_auto_height)




class SafeEditableComboBox(QComboBox):
    """可输入下拉框：避免内部输入框覆盖右侧下拉按钮，并保证按钮热区可点击。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self._line_edit_syncing = False
        self._popup_hot_pressed = False
        try:
            line_edit = self.lineEdit()
            if line_edit is not None:
                # QLineEdit 是 QComboBox 内部输入框，去掉边框并交给下拉框外壳统一绘制。
                line_edit.setFrame(False)
                line_edit.setTextMargins(0, 0, 4, 0)
                line_edit.setStyleSheet('QLineEdit { border: none; background: transparent; padding: 0px; min-height: 0px; }')
                line_edit.installEventFilter(self)
        except Exception:
            pass
        qt_safe_single_shot(0, self.sync_line_edit_area)

    def _style_arrow_width(self):
        """读取当前 Qt 样式中的下拉箭头区域宽度。"""
        try:
            option = QStyleOptionComboBox()
            self.initStyleOption(option)
            rect = self.style().subControlRect(QStyle.CC_ComboBox, option, QStyle.SC_ComboBoxArrow, self)
            return max(0, int(rect.width()))
        except Exception:
            return 0

    def _reserved_button_width(self):
        """右侧为下拉按钮预留宽度，避免可输入区域盖住箭头。"""
        try:
            height = int(self.height())
        except Exception:
            height = 24
        # 这里比实际箭头按钮略宽一点，解决箭头左侧边缘被内部输入框吃掉的问题。
        return max(30, min(58, max(height + 10, self._style_arrow_width() + 12)))

    def _point_in_popup_zone(self, point):
        """判断鼠标点是否落在右侧下拉按钮热区。"""
        try:
            return int(point.x()) >= max(0, int(self.width()) - self._reserved_button_width())
        except Exception:
            return False

    def _show_popup_from_hot_zone(self):
        """从右侧热区点击时打开下拉菜单。"""
        try:
            self.sync_line_edit_area()
            self.setFocus(Qt.MouseFocusReason)
            self.showPopup()
        except RuntimeError:
            return
        except Exception:
            return

    def sync_line_edit_area(self):
        """把内部 QLineEdit 限制在下拉箭头左侧。"""
        if getattr(self, '_line_edit_syncing', False):
            return
        self._line_edit_syncing = True
        try:
            if not self.isEditable():
                return
            line_edit = self.lineEdit()
            if line_edit is None:
                return
            try:
                line_edit.setFrame(False)
                line_edit.setTextMargins(0, 0, 4, 0)
                line_edit.setStyleSheet('QLineEdit { border: none; background: transparent; padding: 0px; min-height: 0px; }')
                line_edit.installEventFilter(self)
            except Exception:
                pass
            width = max(1, int(self.width()))
            height = max(1, int(self.height()))
            left = 4
            top = 2
            bottom = 2
            reserved = self._reserved_button_width()
            edit_width = max(20, width - reserved - left - 4)
            edit_height = max(1, height - top - bottom)
            line_edit.setGeometry(left, top, edit_width, edit_height)
        except RuntimeError:
            return
        except Exception:
            return
        finally:
            self._line_edit_syncing = False

    def eventFilter(self, watched, event):
        """拦截内部输入框右侧热区点击，避免事件被 QLineEdit 吃掉。"""
        try:
            line_edit = self.lineEdit()
            is_line_edit = watched is not None and watched == line_edit
        except Exception:
            is_line_edit = False
        if is_line_edit and event.type() in (QEvent.MouseButtonPress, QEvent.MouseButtonRelease, QEvent.MouseButtonDblClick):
            try:
                combo_point = watched.mapTo(self, event.pos())
                if event.button() == Qt.LeftButton and self._point_in_popup_zone(combo_point):
                    if event.type() in (QEvent.MouseButtonPress, QEvent.MouseButtonDblClick):
                        self._popup_hot_pressed = True
                        self._show_popup_from_hot_zone()
                    elif self._popup_hot_pressed:
                        self._popup_hot_pressed = False
                    event.accept()
                    return True
            except Exception:
                pass
        return super().eventFilter(watched, event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._point_in_popup_zone(event.pos()):
            self._popup_hot_pressed = True
            self._show_popup_from_hot_zone()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and getattr(self, '_popup_hot_pressed', False):
            self._popup_hot_pressed = False
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def setFont(self, font):
        super().setFont(font)
        try:
            line_edit = self.lineEdit()
            if line_edit is not None:
                line_edit.setFont(font)
        except Exception:
            pass
        qt_safe_single_shot(0, self.sync_line_edit_area)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.sync_line_edit_area()
        qt_safe_single_shot(0, self.sync_line_edit_area)

    def showEvent(self, event):
        super().showEvent(event)
        qt_safe_single_shot(0, self.sync_line_edit_area)



class UiSettingsDialog(QDialog):
    """界面设置窗口：左侧字号/尺寸，右侧主题颜色，中间可拖动调整宽度。"""

    def __init__(self, current_settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle('界面设置')
        self.resize(980, 560)
        self._settings = dict(current_settings or {})
        self.color_edits = {}
        self._changing_theme = False
        self.init_ui()

    @staticmethod
    def _int_value(settings, key, default):
        try:
            return int(settings.get(key, default))
        except Exception:
            return int(default)

    def _make_spin(self, key, default, minimum, maximum):
        spin = QSpinBox()
        spin.setRange(minimum, maximum)
        spin.setValue(self._int_value(self._settings, key, default))
        return spin

    def init_ui(self):
        layout = QVBoxLayout(self)
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        layout.addWidget(splitter, 1)

        left_scroll = QScrollArea()
        self.left_settings_scroll = left_scroll
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 8, 8, 8)
        left_layout.setSpacing(8)
        left_title = QLabel('字号与尺寸设置')
        left_title.setWordWrap(True)
        left_layout.addWidget(left_title)

        form = QFormLayout()
        form.setRowWrapPolicy(QFormLayout.DontWrapRows)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)

        self.font_scale_spin = self._make_spin('font_scale', 100, 30, 300)
        form.addRow('总字体缩放(%):', self.font_scale_spin)

        self.option_font_spin = self._make_spin('option_font_size', 10, 8, 30)
        form.addRow('选项基础字号:', self.option_font_spin)
        self.option_compact_spin = self._make_spin('option_compactness', 4, 0, 30)
        form.addRow('选项紧凑程度:', self.option_compact_spin)
        self.option_input_height_spin = self._make_spin('option_input_height', 96, 24, 600)
        form.addRow('选项输入框高度:', self.option_input_height_spin)

        self.toolbar_font_spin = self._make_spin('toolbar_font_size', 10, 8, 30)
        form.addRow('工具栏基础字号:', self.toolbar_font_spin)
        self.toolbar_compact_spin = self._make_spin('toolbar_compactness', 4, 0, 30)
        form.addRow('工具栏紧凑程度:', self.toolbar_compact_spin)

        self.copy_font_spin = self._make_spin('copy_font_size', 10, 8, 30)
        form.addRow('复制按钮基础字号:', self.copy_font_spin)
        self.copy_compact_spin = self._make_spin('copy_compactness', 4, 0, 30)
        form.addRow('复制按钮紧凑程度:', self.copy_compact_spin)

        self.preview_font_spin = self._make_spin('preview_font_size', 11, 8, 36)
        form.addRow('预览框基础字号:', self.preview_font_spin)
        self.preview_compact_spin = self._make_spin('preview_compactness', 4, 0, 30)
        form.addRow('预览框紧凑程度:', self.preview_compact_spin)

        self.size_form = form
        self._polish_settings_form(form)
        left_layout.addLayout(form)
        hint = QLabel('总字体缩放会按比例放大/缩小全部文字；基础字号使用 QFont 点数；紧凑程度为数字，数值越小越紧凑；选项输入框高度为文本输入框自动增高后的最大高度。')
        hint.setWordWrap(True)
        left_layout.addWidget(hint)
        left_layout.addStretch()
        left_scroll.setWidget(left_panel)
        splitter.addWidget(left_scroll)

        right_scroll = QScrollArea()
        self.right_settings_scroll = right_scroll
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QFrame.NoFrame)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 8, 8, 8)
        right_layout.setSpacing(8)
        right_title = QLabel('主题与控件颜色')
        right_title.setWordWrap(True)
        right_layout.addWidget(right_title)

        theme_form = QFormLayout()
        theme_form.setRowWrapPolicy(QFormLayout.DontWrapRows)
        theme_form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        theme_form.setHorizontalSpacing(12)
        theme_form.setVerticalSpacing(8)

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(list(theme_presets.keys()) + ['自定义'])
        theme_name = str(self._settings.get('theme_name') or '浅色')
        self.theme_combo.setCurrentText(theme_name if theme_name in theme_presets or theme_name == '自定义' else '浅色')
        self.theme_combo.currentTextChanged.connect(self.on_theme_changed)
        theme_form.addRow('主题预设:', self.theme_combo)

        current_theme = self.theme_combo.currentText()
        preset = theme_presets.get(current_theme, theme_presets['浅色'])
        for key, label in theme_color_fields:
            widget = self._make_color_widget(key, preset.get(key, '#ffffff'))
            theme_form.addRow(label + ':', widget)

        self.theme_form = theme_form
        self._polish_settings_form(theme_form)
        right_layout.addLayout(theme_form)
        theme_hint = QLabel('颜色支持 #RRGGBB 格式。选择预设会填入对应颜色；手动改颜色或点击“选择”后会自动切换为“自定义”。')
        theme_hint.setWordWrap(True)
        right_layout.addWidget(theme_hint)
        right_layout.addStretch()
        right_scroll.setWidget(right_panel)
        splitter.addWidget(right_scroll)
        splitter.setSizes([430, 550])
        try:
            left_scroll.viewport().installEventFilter(self)
            right_scroll.viewport().installEventFilter(self)
            left_scroll.installEventFilter(self)
            right_scroll.installEventFilter(self)
        except Exception:
            pass
        qt_safe_single_shot(0, self.update_settings_form_widths)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _make_color_widget(self, key, default_color):
        widget = QWidget()
        row = QHBoxLayout(widget)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)
        edit = QLineEdit()
        edit.setText(normalize_color_value(self._settings.get(key, default_color), default_color))
        edit.setPlaceholderText('#RRGGBB')
        edit.setFixedWidth(92)
        edit.editingFinished.connect(self.mark_custom_theme)
        button = QPushButton('选择')
        button.setFixedWidth(64)
        # QColorDialog 是 PyQt5 内置颜色选择对话框，用于从调色盘中选择颜色。
        button.clicked.connect(lambda _checked=False, name=key: self.choose_color(name))
        row.addWidget(edit, 0)
        row.addWidget(button, 0)
        row.addStretch(1)
        widget.setMinimumWidth(0)
        self.color_edits[key] = edit
        return widget

    def mark_custom_theme(self):
        if self._changing_theme:
            return
        if hasattr(self, 'theme_combo') and self.theme_combo.currentText() != '自定义':
            self.theme_combo.setCurrentText('自定义')

    def choose_color(self, key):
        edit = self.color_edits.get(key)
        if edit is None:
            return
        current = QColor(edit.text().strip())
        if not current.isValid():
            current = QColor('#ffffff')
        selected = QColorDialog.getColor(current, self, '选择颜色')
        if selected.isValid():
            edit.setText(selected.name())
            self.mark_custom_theme()

    def on_theme_changed(self, theme_name):
        if theme_name not in theme_presets:
            return
        self._changing_theme = True
        try:
            for key, value in theme_presets[theme_name].items():
                edit = self.color_edits.get(key)
                if edit is not None:
                    edit.setText(value)
        finally:
            self._changing_theme = False

    def _spin_width_for(self, spin):
        """按界面设置窗口数字输入框的内容长度计算宽度，避免无意义拉长。"""
        try:
            texts = [str(spin.value()), str(spin.minimum()), str(spin.maximum())]
            width = max(text_width_for(spin, text) for text in texts)
        except Exception:
            width = 42
        return max(58, min(90, int(width) + 34))

    def _fit_settings_field_widget(self, widget):
        """仅用于界面设置窗口：字段控件按内容收缩，右侧按钮/下拉框不被输入框顶出。"""
        if widget is None:
            return 0
        try:
            widget.setMinimumWidth(0)
            widget.setMaximumWidth(16777215)
            if isinstance(widget, QAbstractSpinBox):
                width = self._spin_width_for(widget)
                widget.setFixedWidth(width)
                return int(width)
            elif isinstance(widget, QLineEdit):
                text = widget.text() or widget.placeholderText() or ''
                width = max(70, min(150, int(text_width_for(widget, text) + 26)))
                widget.setFixedWidth(width)
                return int(width)
            elif isinstance(widget, QPushButton):
                width = adaptive_text_width(widget, widget.text(), 4, min_width=54, max_width=90)
                widget.setFixedWidth(width)
                return int(width)
            elif isinstance(widget, QComboBox):
                texts = [widget.itemText(i) for i in range(widget.count())]
                max_text = max(texts or [''], key=lambda item: text_width_for(widget, item))
                width = adaptive_text_width(widget, max_text, 4, min_width=86, max_width=150)
                widget.setFixedWidth(width)
                return int(width)
            else:
                total = 0
                children = [child for child in widget.findChildren(QWidget) if isinstance(child, (QAbstractSpinBox, QLineEdit, QPushButton, QComboBox))]
                for child in children:
                    total += self._fit_settings_field_widget(child)
                try:
                    layout = widget.layout()
                    if layout is not None:
                        left, _, right, _ = layout.getContentsMargins()
                        spacing = layout.spacing()
                        if children:
                            total += left + right + spacing * max(0, len(children) - 1)
                except Exception:
                    pass
                if total > 0:
                    widget.setMinimumWidth(int(total))
                return int(total)
        except Exception:
            return 0

    def _polish_settings_form(self, form, panel_width=None):
        """仅用于界面设置窗口：标签按内容自适应，空间不足时自动换行，并优先保留右侧控件。"""
        try:
            if panel_width is None or int(panel_width) <= 0:
                panel_width = self.width() // 2
            panel_width = max(220, int(panel_width))
            row_infos = []
            max_field_width = 0
            max_label_natural = 0
            for row in range(form.rowCount()):
                label_item = form.itemAt(row, QFormLayout.LabelRole)
                field_item = form.itemAt(row, QFormLayout.FieldRole)
                label_widget = label_item.widget() if label_item is not None else None
                field_widget = field_item.widget() if field_item is not None else None
                field_width = self._fit_settings_field_widget(field_widget)
                label_natural = 80
                if label_widget is not None and hasattr(label_widget, 'text'):
                    label_natural = text_width_for(label_widget, label_widget.text()) + 18
                max_field_width = max(max_field_width, int(field_width or 0))
                max_label_natural = max(max_label_natural, int(label_natural))
                row_infos.append((label_widget, field_widget, int(label_natural), int(field_width or 0)))

            try:
                h_spacing = int(form.horizontalSpacing())
                if h_spacing < 0:
                    h_spacing = 12
            except Exception:
                h_spacing = 12
            # 标签不再限制为分割区四分之一；只保留一个右侧控件所需的安全宽度。
            safety_margin = 36
            field_reserve = max(70, max_field_width)
            label_max_by_space = max(70, panel_width - field_reserve - h_spacing - safety_margin)
            shared_label_width = max(70, min(max_label_natural, label_max_by_space))

            for label_widget, field_widget, _label_natural, _field_width in row_infos:
                if label_widget is not None:
                    label_widget.setFixedWidth(int(shared_label_width))
                    if isinstance(label_widget, QLabel):
                        label_widget.setWordWrap(True)
                        label_widget.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self._fit_settings_field_widget(field_widget)
        except Exception:
            pass

    def update_settings_form_widths(self):
        """左右分割条拖动或窗口缩放后，只刷新界面设置窗口内部表单宽度。"""
        try:
            left_width = self.left_settings_scroll.viewport().width() if hasattr(self, 'left_settings_scroll') else self.width() // 2
            right_width = self.right_settings_scroll.viewport().width() if hasattr(self, 'right_settings_scroll') else self.width() // 2
            if hasattr(self, 'size_form'):
                self._polish_settings_form(self.size_form, left_width)
            if hasattr(self, 'theme_form'):
                self._polish_settings_form(self.theme_form, right_width)
        except Exception:
            pass

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Resize:
            try:
                if watched in (self.left_settings_scroll, self.left_settings_scroll.viewport(), self.right_settings_scroll, self.right_settings_scroll.viewport()):
                    qt_safe_single_shot(0, self.update_settings_form_widths)
            except Exception:
                pass
        return super().eventFilter(watched, event)

    def get_settings(self):
        payload = {
            'font_scale': int(self.font_scale_spin.value()),
            'option_font_size': int(self.option_font_spin.value()),
            'option_compactness': int(self.option_compact_spin.value()),
            'option_input_height': int(self.option_input_height_spin.value()),
            'toolbar_font_size': int(self.toolbar_font_spin.value()),
            'toolbar_compactness': int(self.toolbar_compact_spin.value()),
            'copy_font_size': int(self.copy_font_spin.value()),
            'copy_compactness': int(self.copy_compact_spin.value()),
            'preview_font_size': int(self.preview_font_spin.value()),
            'preview_compactness': int(self.preview_compact_spin.value()),
            'theme_name': self.theme_combo.currentText(),
        }
        theme_name = payload['theme_name']
        preset = theme_presets.get(theme_name, theme_presets['浅色'])
        for key, _label in theme_color_fields:
            edit = self.color_edits.get(key)
            default = preset.get(key, theme_presets['浅色'].get(key, '#ffffff'))
            payload[key] = normalize_color_value(edit.text() if edit is not None else default, default)
        return payload


class BrowserSettingsDialog(QDialog):
    """主界面工具栏中的浏览器参数设置窗口。"""

    def __init__(self, browser_settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle('浏览器参数设置')
        self.resize(680, 220)
        self.browser_settings = dict(browser_settings or {})
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.driver_edit = QLineEdit()
        self.driver_edit.setPlaceholderText('chromedriver.exe 路径')
        self.driver_edit.setText(str(self.browser_settings.get('chromedriver_path', '') or ''))
        form.addRow('Driver路径:', self.driver_edit)

        self.binary_edit = QLineEdit()
        self.binary_edit.setPlaceholderText('chrome.exe 路径，可留空')
        self.binary_edit.setText(str(self.browser_settings.get('chrome_binary', '') or ''))
        form.addRow('浏览器路径:', self.binary_edit)

        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText('点击打开浏览器时使用的 URL')
        self.url_edit.setText(str(self.browser_settings.get('start_url', '') or ''))
        form.addRow('启动URL:', self.url_edit)

        layout.addLayout(form)
        hint = QLabel('这些参数仍然保存到当前模板的浏览器流程配置中。')
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_browser_settings(self):
        return {
            'chromedriver_path': self.driver_edit.text().strip(),
            'chrome_binary': self.binary_edit.text().strip(),
            'start_url': self.url_edit.text().strip(),
        }


class InputOptionRow(QWidget):
    """GTE 风格的输入选项行：左侧拖动柄，中间标签，右侧输入控件。"""

    def __init__(self, option_config: dict, editor_widget, main_window, parent=None):
        super().__init__(parent)
        self.setObjectName('inputOptionRow')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.option_config = option_config
        self.editor = editor_widget
        self.main_window = main_window
        self.drag_start_y = None
        self.mouse_dragging = False
        self._selected = False
        self._adaptive_updating = False

        self.row_layout = QHBoxLayout(self)
        self.row_layout.setContentsMargins(3, 3, 3, 3)
        self.row_layout.setSpacing(4)

        self.handle = QLabel('☰')
        self.handle.setObjectName('inputOptionHandle')
        self.handle.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.handle.setCursor(Qt.SizeAllCursor)
        self.handle.mousePressEvent = self.handle_mouse_press
        self.handle.mouseMoveEvent = self.handle_mouse_move
        self.handle.mouseReleaseEvent = self.handle_mouse_release

        self.label = QLabel(self.get_wrap_text(self.get_name()))
        self.label.setObjectName('inputOptionLabel')
        self.label.setWordWrap(True)
        self.label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.editor.installEventFilter(self)
        try:
            viewport = self.editor.viewport()
            if viewport is not None:
                viewport.installEventFilter(self)
        except Exception:
            pass
        try:
            line_edit = self.editor.lineEdit() if isinstance(self.editor, QComboBox) and self.editor.isEditable() else None
            if line_edit is not None:
                line_edit.installEventFilter(self)
        except Exception:
            pass
        try:
            if isinstance(self.editor, QTextEdit):
                self.editor.document().contentsChanged.connect(lambda: qt_safe_single_shot(0, self.update_adaptive_size))
        except Exception:
            pass

        self.row_layout.addWidget(self.handle, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.row_layout.addWidget(self.label, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.row_layout.addWidget(self.editor, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.set_selected(False)

    def get_name(self):
        return str(self.option_config.get('label', '') or '')

    def set_name(self, name):
        self.option_config['label'] = str(name or '').strip()
        self.label.setText(self.get_wrap_text(self.get_name()))
        self.update_adaptive_size()

    def get_wrap_text(self, text):
        return '\u200b'.join(str(text))

    def _current_settings(self):
        if hasattr(self.main_window, '_ui_settings'):
            return self.main_window._normalize_ui_settings(self.main_window._ui_settings)
        return {'font_scale': 100, 'option_font_size': 10, 'option_compactness': 4, 'option_input_height': 96}

    def _available_row_width(self):
        """返回当前滚动视口可用宽度，不使用当前行 fixedWidth 作为计算基准。"""
        width = 0
        try:
            width = self.main_window.input_scroll.viewport().width()
        except Exception:
            width = 0
        if width <= 0:
            width = 420
        try:
            left, _, right, _ = self.main_window.input_layout.getContentsMargins()
            width -= left + right
        except Exception:
            width -= 10
        return max(80, int(width))

    def _minimum_layout_width(self):
        try:
            return int(self.main_window.option_row_min_layout_width())
        except Exception:
            return int(input_option_min_layout_width)

    def _effective_layout_width(self):
        """计算本轮布局宽度。

        分割栏缩得过窄时，行内标签、输入框不再继续缩小到极端宽度，
        而是固定使用最小布局宽度；这样长文本不会在窄宽度下反复重排，
        避免卡死和抖动。
        """
        return max(self._minimum_layout_width(), self._available_row_width())

    def _layout_horizontal_cost(self):
        try:
            left, _, right, _ = self.row_layout.getContentsMargins()
            spacing = self.row_layout.spacing()
        except Exception:
            left = right = spacing = 4
        return int(left + right + spacing * 2)

    def reset_adaptive_constraints(self):
        """每次重新计算前先清空上一轮固定尺寸，防止拖动分隔栏时尺寸叠加。"""
        for widget in (self, self.handle, self.label, self.editor):
            try:
                widget.setMinimumWidth(0)
                widget.setMaximumWidth(16777215)
            except Exception:
                pass
        for widget in (self, self.handle, self.label):
            try:
                widget.setMinimumHeight(0)
                widget.setMaximumHeight(16777215)
            except Exception:
                pass

    def _label_size(self, available_width, font_size, compact):
        font = self.label.font()
        font.setPointSize(font_size)
        metrics = QFontMetrics(font)
        natural_width = text_width_for(font, self.get_name()) + compact * 2 + 16
        label_max_width = max(48, int(available_width) // 4)
        label_min_width = min(label_max_width, max(48, font_size * 4))
        label_width = max(label_min_width, min(label_max_width, natural_width))
        self.label.setFixedWidth(int(label_width))
        try:
            text_height = self.label.heightForWidth(int(label_width))
        except Exception:
            text_height = metrics.lineSpacing()
        label_height = max(metrics.lineSpacing() + compact * 2, int(text_height) + compact)
        return int(label_width), int(label_height)

    def _editor_single_line_height(self, compact):
        return adaptive_control_height(self.editor, compact, min_height=24)

    def update_editor_adaptive_size(self, compact: int, editor_width: int):
        single_line_height = self._editor_single_line_height(compact)
        max_input_height = int(self._current_settings().get('option_input_height', 96))
        max_input_height = max(single_line_height, max_input_height)
        try:
            self.editor.setMinimumWidth(0)
            self.editor.setMaximumWidth(16777215)
            self.editor.setFixedWidth(max(60, int(editor_width)))
        except Exception:
            pass
        if isinstance(self.editor, QComboBox):
            # 下拉菜单和可输入下拉菜单恢复 PyQt 原生单行样式，不再做自定义换行/弹窗绘制。
            self.editor.setMinimumHeight(single_line_height)
            self.editor.setMaximumHeight(single_line_height)
            self.editor.setFixedHeight(single_line_height)
            try:
                self.editor.view().setMinimumWidth(max(100, int(editor_width)))
                self.editor.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            except Exception:
                pass
            try:
                if hasattr(self.editor, 'sync_line_edit_area'):
                    self.editor.sync_line_edit_area()
                    qt_safe_single_shot(0, self.editor.sync_line_edit_area)
            except Exception:
                pass
        elif isinstance(self.editor, QTextEdit):
            self.editor.setLineWrapMode(QTextEdit.WidgetWidth)
            try:
                self.editor.setWordWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
            except Exception:
                pass
            if isinstance(self.editor, AutoWrapTextEdit):
                # 使用本轮计算出的 editor_width 作为文档换行基准，不能用上一轮 viewport 宽度。
                text_width = max(1, int(editor_width) - 12)
                self.editor.set_available_text_width(text_width)
                self.editor.set_auto_limits(single_line_height, max_input_height)
                self.editor.update_auto_height()
            else:
                self.editor.setMinimumHeight(single_line_height)
                self.editor.setMaximumHeight(max_input_height)
                self.editor.setFixedHeight(single_line_height)
        elif isinstance(self.editor, (QLineEdit, QCheckBox)):
            self.editor.setFixedHeight(single_line_height)

    def update_adaptive_size(self):
        if getattr(self, '_adaptive_updating', False):
            return
        self._adaptive_updating = True
        try:
            self.reset_adaptive_constraints()
            settings = self._current_settings()
            option_font_size = scaled_point_size(settings, 'option_font_size', 10)
            compact = int(settings.get('option_compactness', 4))
            available_width = self._effective_layout_width()
            handle_width = max(24, text_width_for(self.handle, '☰') + compact * 2 + 14)
            self.handle.setFixedWidth(int(handle_width))
            label_width, label_height = self._label_size(available_width, option_font_size, compact)
            editor_width = max(60, available_width - handle_width - label_width - self._layout_horizontal_cost())
            self.update_editor_adaptive_size(compact, editor_width)
            editor_height = self.editor.height() if self.editor.height() > 0 else self.editor.minimumHeight()
            metrics = font_metrics_for(self.label)
            handle_height = max(editor_height, metrics.lineSpacing() + compact * 2 + 8)
            self.handle.setFixedHeight(int(handle_height))
            try:
                self.handle.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            except Exception:
                pass
            _, margin_top, _, margin_bottom = self.row_layout.getContentsMargins()
            row_content_height = max(label_height, editor_height, handle_height)
            target_height = max(30, int(row_content_height + margin_top + margin_bottom))
            self.setFixedWidth(int(available_width))
            self.setFixedHeight(target_height)
            self.updateGeometry()
        finally:
            self._adaptive_updating = False

    def set_selected(self, selected):
        self._selected = bool(selected)
        colors = self.main_window._theme_colors() if self.main_window is not None and hasattr(self.main_window, '_theme_colors') else resolve_theme_colors({})
        border = colors['color_primary'] if selected else colors['color_border']
        background = colors['color_selected_bg'] if selected else colors['color_panel']
        handle_background = colors['color_handle_selected_bg'] if selected else colors['color_handle_bg']
        self.setStyleSheet(
            f"QWidget#inputOptionRow {{ background-color: {background}; border: 1px solid {border}; border-radius: 3px; }}"
            f"QLabel#inputOptionHandle {{ background-color: {handle_background}; border-right: 1px solid {colors['color_border']}; }}"
            "QLabel#inputOptionLabel { background-color: transparent; border: none; }"
        )

    def apply_ui_settings(self, settings):
        option_font_size = scaled_point_size(settings, 'option_font_size', 10)
        option_compact = int(settings.get('option_compactness', 4))
        font = QFont()
        font.setPointSize(option_font_size)
        for widget in (self, self.handle, self.label, self.editor):
            widget.setFont(font)
        if isinstance(self.editor, AutoWrapTextEdit):
            self.editor.update_auto_height()
        margin = max(1, option_compact // 2)
        spacing = max(1, option_compact // 2)
        self.row_layout.setContentsMargins(margin, margin, margin, margin)
        self.row_layout.setSpacing(spacing)
        self.set_selected(getattr(self, '_selected', False))
        self.update_adaptive_size()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 尺寸由父级滚动区统一重算。这里不再递归更新，避免拖动分隔栏时反复叠加。

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.main_window.select_option_row(self)
        super().mousePressEvent(event)

    def label_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            self.main_window.select_option_row(self)

    def handle_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_y = event.globalY()
            self.mouse_dragging = False
            self.main_window.select_option_row(self, toggle=False)

    def handle_mouse_move(self, event):
        if self.drag_start_y is None:
            return
        try:
            if not (event.buttons() & Qt.LeftButton):
                self.drag_start_y = None
                self.mouse_dragging = False
                return
        except Exception:
            pass
        if abs(event.globalY() - self.drag_start_y) >= 3:
            self.mouse_dragging = True
            self.main_window.move_option_row_by_mouse(self, event.globalPos())

    def handle_mouse_release(self, event):
        self.drag_start_y = None
        if self.mouse_dragging:
            self.main_window.sync_option_order_from_rows(save=True)
        self.mouse_dragging = False

    def eventFilter(self, watched, event):
        if event.type() in (QEvent.MouseButtonPress, QEvent.FocusIn):
            try:
                is_editor_part = watched == self.editor or watched == self.editor.viewport()
            except Exception:
                is_editor_part = watched == self.editor
            try:
                line_edit = self.editor.lineEdit() if isinstance(self.editor, QComboBox) and self.editor.isEditable() else None
                is_editor_part = is_editor_part or watched == line_edit
            except Exception:
                pass
            if is_editor_part:
                if event.type() == QEvent.FocusIn or getattr(event, 'button', lambda: None)() == Qt.LeftButton:
                    self.main_window.select_option_row(self, toggle=False)
        return super().eventFilter(watched, event)


class TemplateSelectRow(QWidget):
    """固定在选项面板顶部的模板选择行，不参与输入值采集和拖拽排序。"""

    def __init__(self, template_combo, main_window, parent=None):
        super().__init__(parent)
        self.setObjectName('templateSelectRow')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.template_combo = template_combo
        self.main_window = main_window
        self._adaptive_updating = False

        self.row_layout = QHBoxLayout(self)
        self.row_layout.setContentsMargins(3, 3, 3, 3)
        self.row_layout.setSpacing(4)

        self.fixed_mark = QLabel('')
        self.fixed_mark.setObjectName('templateFixedMark')
        self.fixed_mark.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.label = QLabel('模板')
        self.label.setObjectName('templateSelectLabel')
        self.label.setWordWrap(True)
        self.label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        self.template_combo.setMinimumWidth(0)
        self.template_combo.view().setMinimumWidth(260)
        self.row_layout.addWidget(self.fixed_mark, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.row_layout.addWidget(self.label, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.row_layout.addWidget(self.template_combo, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.setStyleSheet(
            'QWidget#templateSelectRow { background-color: #f9fafb; border: 1px solid #9ca3af; border-radius: 3px; }'
            'QLabel#templateFixedMark { background-color: transparent; border: none; }'
            'QLabel#templateSelectLabel { background-color: transparent; border: none; font-weight: bold; }'
        )

    def _current_settings(self):
        if hasattr(self.main_window, '_ui_settings'):
            return self.main_window._normalize_ui_settings(self.main_window._ui_settings)
        return {'font_scale': 100, 'option_font_size': 10, 'option_compactness': 4}

    def _minimum_layout_width(self):
        try:
            return int(self.main_window.option_row_min_layout_width())
        except Exception:
            return int(input_option_min_layout_width)

    def reset_adaptive_constraints(self):
        for widget in (self, self.fixed_mark, self.label, self.template_combo):
            try:
                widget.setMinimumWidth(0)
                widget.setMaximumWidth(16777215)
            except Exception:
                pass
        for widget in (self, self.fixed_mark, self.label, self.template_combo):
            try:
                widget.setMinimumHeight(0)
                widget.setMaximumHeight(16777215)
            except Exception:
                pass

    def apply_ui_settings(self, settings):
        option_font_size = scaled_point_size(settings, 'option_font_size', 10)
        option_compact = int(settings.get('option_compactness', 4))
        font = QFont()
        font.setPointSize(option_font_size)
        self.setFont(font)
        for child in self.findChildren(QWidget):
            child.setFont(font)
        colors = self.main_window._theme_colors(settings) if self.main_window is not None and hasattr(self.main_window, '_theme_colors') else resolve_theme_colors(settings)
        self.setStyleSheet(
            f"QWidget#templateSelectRow {{ background-color: {colors['color_panel']}; border: 1px solid {colors['color_border']}; border-radius: 3px; }}"
            "QLabel#templateFixedMark { background-color: transparent; border: none; }"
            "QLabel#templateSelectLabel { background-color: transparent; border: none; font-weight: bold; }"
        )
        margin = max(1, option_compact // 2)
        spacing = max(1, option_compact // 2)
        self.row_layout.setContentsMargins(margin, margin, margin, margin)
        self.row_layout.setSpacing(spacing)
        self.update_adaptive_size()

    def update_adaptive_size(self):
        if getattr(self, '_adaptive_updating', False):
            return
        self._adaptive_updating = True
        try:
            self.reset_adaptive_constraints()
            settings = self._current_settings()
            option_font_size = scaled_point_size(settings, 'option_font_size', 10)
            option_compact = int(settings.get('option_compactness', 4))
            metrics = font_metrics_for(self.label)
            try:
                available_width = self.main_window.input_scroll.viewport().width()
                left, _, right, _ = self.main_window.input_layout.getContentsMargins()
                available_width -= left + right
            except Exception:
                available_width = 420
            available_width = max(self._minimum_layout_width(), int(available_width))
            handle_width = max(24, text_width_for(self.fixed_mark, '☰') + option_compact * 2 + 14)
            self.fixed_mark.setFixedWidth(int(handle_width))
            max_label_width = max(48, available_width // 4)
            label_width = adaptive_text_width(self.label, '模板', option_compact, min_width=48, max_width=max_label_width)
            self.label.setFixedWidth(int(label_width))
            try:
                left, _, right, _ = self.row_layout.getContentsMargins()
                spacing = self.row_layout.spacing()
                horizontal_cost = left + right + spacing * 2
            except Exception:
                horizontal_cost = 12
            combo_width = max(80, available_width - handle_width - label_width - horizontal_cost)
            combo_height = adaptive_control_height(self.template_combo, option_compact, min_height=24)
            self.template_combo.setFixedWidth(int(combo_width))
            self.template_combo.setFixedHeight(int(combo_height))
            try:
                self.template_combo.view().setMinimumWidth(max(int(combo_width), option_font_size * 20))
            except Exception:
                pass
            _, margin_top, _, margin_bottom = self.row_layout.getContentsMargins()
            label_height = self.label.heightForWidth(self.label.width()) if self.label.hasHeightForWidth() else metrics.lineSpacing()
            label_height = max(metrics.lineSpacing() + option_compact * 2, int(label_height) + option_compact)
            row_height = max(label_height, combo_height, metrics.lineSpacing() + option_compact * 2 + 8) + margin_top + margin_bottom
            target_height = max(30, int(row_height))
            self.setFixedWidth(int(available_width))
            self.setFixedHeight(target_height)
            self.updateGeometry()
        finally:
            self._adaptive_updating = False

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 由父级滚动区统一刷新，避免行自身 resizeEvent 和 setFixedWidth 互相触发。

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.main_window.clear_selected_option_row()
        super().mousePressEvent(event)


class CopyButtonItem(QWidget):
    """下方复制按钮区的 GTE 风格项目：上方小拖动柄，下方按钮。"""

    def __init__(self, field_name: str, main_window, parent=None):
        super().__init__(parent)
        self.setObjectName('copyButtonItem')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Minimum)
        self.field_name = str(field_name or '').strip()
        self.main_window = main_window
        self.drag_start_x = None
        self.mouse_dragging = False
        self._adaptive_updating = False
        self._hidden_by_filter = False

        self.item_layout = QVBoxLayout(self)
        self.item_layout.setContentsMargins(4, 4, 4, 4)
        self.item_layout.setSpacing(3)

        self.handle = QLabel('•••')
        self.handle.setObjectName('copyButtonHandle')
        self.handle.setAlignment(Qt.AlignCenter)
        self.handle.setCursor(Qt.SizeAllCursor)
        self.handle.mousePressEvent = self.handle_mouse_press
        self.handle.mouseMoveEvent = self.handle_mouse_move
        self.handle.mouseReleaseEvent = self.handle_mouse_release

        self.button = QPushButton(self.field_name)
        self.button.setCheckable(True)
        self.button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Minimum)
        self.button.clicked.connect(lambda checked, f=self.field_name: self.main_window.on_copy_button_clicked(self, f, checked))

        self.item_layout.addWidget(self.handle)
        self.item_layout.addWidget(self.button)
        self.set_selected(False)
        self.update_adaptive_size()

    def _current_settings(self):
        if hasattr(self.main_window, '_ui_settings'):
            return self.main_window._normalize_ui_settings(self.main_window._ui_settings)
        return {'font_scale': 100, 'copy_font_size': 10, 'copy_compactness': 4}

    def update_adaptive_size(self):
        if getattr(self, '_adaptive_updating', False):
            return
        self._adaptive_updating = True
        try:
            settings = self._current_settings()
            copy_font_size = scaled_point_size(settings, 'copy_font_size', 10)
            copy_compact = int(settings.get('copy_compactness', 4))
            font = self.button.font()
            font.setPointSize(copy_font_size)
            self.button.setFont(font)
            self.handle.setFont(font)
            margin_left, margin_top, margin_right, margin_bottom = self.item_layout.getContentsMargins()
            # 复制按钮文字不再插入换行，按钮宽度直接按完整文字宽度自适应。
            if self.button.text() != self.field_name:
                self.button.setText(self.field_name)
            button_width = text_width_for(font, self.field_name) + copy_compact * 2 + 28
            button_width = max(24, int(button_width))
            # 拖动柄只作为排序标记使用，保持很小高度，避免挤占复制按钮区纵向空间。
            handle_height = max(8, min(12, int(copy_compact) + 8))
            button_height = adaptive_control_height(font, copy_compact, min_height=24, lines=1)
            handle_font = QFont(font)
            handle_font.setPointSize(max(6, min(copy_font_size, 8)))
            self.handle.setFont(handle_font)
            self.handle.setFixedHeight(handle_height)
            self.handle.setFixedWidth(button_width)
            self.button.setFixedWidth(button_width)
            self.button.setMinimumHeight(button_height)
            self.button.setMaximumHeight(button_height)
            total_width = button_width + margin_left + margin_right
            total_height = handle_height + button_height + self.item_layout.spacing() + margin_top + margin_bottom
            self.setMinimumSize(total_width, total_height)
            self.setMaximumWidth(total_width)
            self.updateGeometry()
        finally:
            self._adaptive_updating = False

    def apply_ui_settings(self, settings):
        copy_font_size = scaled_point_size(settings, 'copy_font_size', 10)
        copy_compact = int(settings.get('copy_compactness', 4))
        font = QFont()
        font.setPointSize(copy_font_size)
        for widget in (self, self.handle, self.button):
            widget.setFont(font)
        margin = max(1, copy_compact // 2)
        spacing = max(1, copy_compact // 2)
        if self.item_layout is not None:
            self.item_layout.setContentsMargins(margin, margin, margin, margin)
            self.item_layout.setSpacing(spacing)
        self.set_selected(self.isChecked())
        self.update_adaptive_size()
        self.adjustSize()

    def isChecked(self):
        return self.button.isChecked()

    def setChecked(self, checked):
        self.button.setChecked(bool(checked))
        self.set_selected(bool(checked))

    def blockSignals(self, block):
        return self.button.blockSignals(block)

    def set_selected(self, selected):
        self._selected = bool(selected)
        colors = self.main_window._theme_colors() if self.main_window is not None and hasattr(self.main_window, '_theme_colors') else resolve_theme_colors({})
        border = colors['color_primary'] if selected else colors['color_border']
        background = colors['color_selected_bg'] if selected else colors['color_panel']
        handle_background = colors['color_handle_selected_bg'] if selected else colors['color_handle_bg']
        self.setStyleSheet(
            f"QWidget#copyButtonItem {{ background-color: {background}; border: 1px solid {border}; border-radius: 3px; }}"
            f"QLabel#copyButtonHandle {{ background-color: {handle_background}; border-bottom: 1px solid {colors['color_border']}; }}"
        )

    def resizeEvent(self, event):
        super().resizeEvent(event)
        try:
            viewport_width = self.main_window.copy_scroll.viewport().width() if self.main_window is not None and hasattr(self.main_window, 'copy_scroll') else 0
            sig = (int(event.size().width()), int(event.size().height()), int(viewport_width), round(float(self.font().pointSizeF()), 2))
            if sig == getattr(self, '_last_resize_sig', None):
                return
            self._last_resize_sig = sig
            qt_safe_single_shot(0, self.update_adaptive_size)
        except Exception:
            qt_safe_single_shot(0, self.update_adaptive_size)

    def handle_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_x = event.globalX()
            self.mouse_dragging = False
            self.main_window.select_copy_item(self, toggle=False)

    def handle_mouse_move(self, event):
        if self.drag_start_x is None:
            return
        try:
            if not (event.buttons() & Qt.LeftButton):
                self.drag_start_x = None
                self.mouse_dragging = False
                return
        except Exception:
            pass
        if abs(event.globalX() - self.drag_start_x) >= 3:
            self.mouse_dragging = True
            self.main_window.move_copy_item_by_mouse(self, event.globalPos())

    def handle_mouse_release(self, event):
        self.drag_start_x = None
        if self.mouse_dragging:
            self.main_window.sync_copy_order_from_items(save=True)
        self.mouse_dragging = False


class PEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.template_db = TemplateDB()
        self.data_matcher = DataMatcher(self.db, self.template_db)
        self.current_template_name = None
        self.current_options_config = []
        self.current_rules_config = []
        self.option_rows = []
        self.selected_option_row = None
        self.input_widgets = []
        self.copy_buttons = []
        self.copy_button_group = QButtonGroup(self)
        self.copy_button_group.setExclusive(False)
        self._copy_selected_fields = []
        self._copy_last_selected_field = ''
        self.data_manager_window = None
        self.template_editor_window = None
        self.browser_flow_window = None
        self.settings_file = self._get_settings_path()
        self._live_process_content = None
        self._last_export_signature = None
        self._last_render_result_text = ''
        self._last_final_fields = {}
        self._last_input_values = {}
        self._input_value_cache = {}
        self._last_data_pool = {}
        self._updating_option_sources = False
        self._live_refresh_timer = QTimer(self)
        self._live_refresh_timer.setInterval(250)
        self._live_refresh_timer.timeout.connect(self._poll_live_updates)

        self.init_ui()
        self.load_template_list()
        self.refresh_input_area()
        self.update_copy_buttons_from_config()
        self.load_settings()
        self._live_refresh_timer.start()


    def _init_window_geometry(self):
        screen = QApplication.primaryScreen()
        if screen is not None:
            geo = screen.availableGeometry()
            width = min(980, max(780, geo.width() - 60))
            height = min(880, max(640, geo.height() - 80))
            self.resize(width, height)
        else:
            self.resize(980, 880)
        self.setMinimumSize(780, 640)

    def option_row_min_layout_width(self):
        """首页输入选项行的最小稳定布局宽度。

        字号变大时最小宽度随字体适当增加；左侧区域小于该宽度时，
        行内容继续按该宽度排版，并通过水平滚动查看，避免长文本在
        极窄宽度下反复换行重算。
        """
        try:
            settings = self._normalize_ui_settings(getattr(self, '_ui_settings', self._default_ui_settings()))
            option_font_size = scaled_point_size(settings, 'option_font_size', 10)
            return max(int(input_option_min_layout_width), option_font_size * 28)
        except Exception:
            return int(input_option_min_layout_width)

    @staticmethod
    def _default_ui_settings():
        settings = {
            'font_scale': 100,
            'option_font_size': 10,
            'option_compactness': 4,
            'option_input_height': 96,
            'toolbar_font_size': 10,
            'toolbar_compactness': 4,
            'copy_font_size': 10,
            'copy_compactness': 4,
            'preview_font_size': 11,
            'preview_compactness': 4,
            'theme_name': '浅色',
        }
        settings.update(theme_presets['浅色'])
        return settings

    def _load_settings_payload(self):
        settings = self._default_ui_settings()
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f) or {}
                    if isinstance(loaded, dict):
                        settings.update(loaded)
                        # 兼容旧版 settings.json：旧版只有 font_size/button_compactness 时，自动迁移到四类设置。
                        if 'font_size' in loaded:
                            old_font = self._safe_int(loaded.get('font_size'), settings['option_font_size'], 8, 36)
                            for key in ('option_font_size', 'toolbar_font_size', 'copy_font_size', 'preview_font_size'):
                                if key not in loaded:
                                    settings[key] = old_font
        except Exception as e:
            print(f'加载设置失败：{e}')
        return self._normalize_ui_settings(settings)

    @staticmethod
    def _safe_int(value, default, minimum=0, maximum=99):
        try:
            number = int(value)
        except Exception:
            number = int(default)
        return max(int(minimum), min(int(maximum), number))

    def _normalize_ui_settings(self, settings: dict):
        defaults = self._default_ui_settings()
        payload = dict(defaults)
        if isinstance(settings, dict):
            payload.update(settings)
        payload['font_scale'] = self._safe_int(payload.get('font_scale'), defaults['font_scale'], 30, 300)
        for key in ('option_font_size', 'toolbar_font_size', 'copy_font_size'):
            payload[key] = self._safe_int(payload.get(key), defaults[key], 8, 30)
        payload['preview_font_size'] = self._safe_int(payload.get('preview_font_size'), defaults['preview_font_size'], 8, 36)
        for key in ('option_compactness', 'toolbar_compactness', 'copy_compactness', 'preview_compactness'):
            payload[key] = self._safe_int(payload.get(key), defaults[key], 0, 30)
        payload['option_input_height'] = self._safe_int(payload.get('option_input_height'), defaults['option_input_height'], 24, 600)
        theme_name = str(payload.get('theme_name') or '浅色')
        if theme_name not in theme_presets and theme_name != '自定义':
            theme_name = '浅色'
        payload['theme_name'] = theme_name
        preset = theme_presets.get(theme_name, theme_presets['浅色'])
        for key, _label in theme_color_fields:
            default_color = preset.get(key, defaults.get(key, '#ffffff'))
            payload[key] = normalize_color_value(payload.get(key, default_color), default_color)
        return payload

    def _theme_colors(self, settings=None):
        return resolve_theme_colors(self._normalize_ui_settings(settings or getattr(self, '_ui_settings', self._default_ui_settings())))

    def _build_theme_palette(self, settings: dict):
        colors = self._theme_colors(settings)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(colors['color_window']))
        palette.setColor(QPalette.WindowText, QColor(colors['color_text']))
        palette.setColor(QPalette.Base, QColor(colors['color_input_bg']))
        palette.setColor(QPalette.AlternateBase, QColor(colors['color_panel']))
        palette.setColor(QPalette.Text, QColor(colors['color_text']))
        palette.setColor(QPalette.Button, QColor(colors['color_button_bg']))
        palette.setColor(QPalette.ButtonText, QColor(colors['color_button_text']))
        palette.setColor(QPalette.Highlight, QColor(colors['color_primary']))
        palette.setColor(QPalette.HighlightedText, QColor(colors['color_preview_bg']))
        return palette

    def _build_app_stylesheet(self, settings: dict):
        """生成主窗口样式表。字号仍由 QFont 控制，QSS 负责颜色、边框、padding 和高度。"""
        settings = self._normalize_ui_settings(settings)
        colors = self._theme_colors(settings)
        option_font = scaled_point_size(settings, 'option_font_size', 10)
        toolbar_font = scaled_point_size(settings, 'toolbar_font_size', 10)
        copy_font = scaled_point_size(settings, 'copy_font_size', 10)
        preview_font = scaled_point_size(settings, 'preview_font_size', 11)
        option_compact = settings['option_compactness']
        toolbar_compact = settings['toolbar_compactness']
        copy_compact = settings['copy_compactness']
        preview_compact = settings['preview_compactness']

        option_min_height = max(22, option_font + option_compact * 2 + 10)
        toolbar_min_height = max(22, toolbar_font + toolbar_compact * 2 + 8)
        copy_min_height = max(22, copy_font + copy_compact * 2 + 8)
        preview_padding = max(0, preview_compact)
        option_padding_v = max(0, option_compact // 2)
        option_padding_h = max(2, option_compact + 2)
        toolbar_padding_v = max(0, toolbar_compact // 2)
        toolbar_padding_h = max(2, toolbar_compact + 2)
        copy_padding_v = max(0, copy_compact // 2)
        copy_padding_h = max(2, copy_compact + 2)

        return (
            f"QMainWindow {{ background-color: {colors['color_window']}; color: {colors['color_text']}; }}\n"
            f"QWidget {{ color: {colors['color_text']}; }}\n"
            f"QFrame#gteLeftPanel, QFrame#gteRightPanel, QFrame#gteDownPanel {{ "
            f"background-color: {colors['color_panel']}; border: 1px solid {colors['color_border']}; }}\n"
            f"QScrollArea, QScrollArea > QWidget > QWidget {{ background-color: {colors['color_panel']}; }}\n"
            f"QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QListWidget, QTableWidget, QTableView, QSpinBox, QDoubleSpinBox {{ "
            f"background-color: {colors['color_input_bg']}; color: {colors['color_text']}; border: 1px solid {colors['color_border']}; }}\n"
            f"QPushButton, QToolButton {{ background-color: {colors['color_button_bg']}; color: {colors['color_button_text']}; "
            f"border: 1px solid {colors['color_border']}; border-radius: 3px; }}\n"
            f"QPushButton:hover, QToolButton:hover {{ border: 1px solid {colors['color_primary']}; }}\n"
            f"QMenu {{ background-color: {colors['color_panel']}; color: {colors['color_text']}; border: 1px solid {colors['color_border']}; }}\n"
            f"QMenu::item:selected {{ background-color: {colors['color_selected_bg']}; color: {colors['color_text']}; }}\n"
            f"QToolBar#mainToolbar QToolButton {{ padding: {toolbar_padding_v}px {toolbar_padding_h}px; min-height: {toolbar_min_height}px; }}\n"
            f"QWidget#inputOptionRow QTextEdit, QWidget#inputOptionRow QComboBox, QWidget#templateSelectRow QComboBox {{ "
            f"padding: {option_padding_v}px {option_padding_h}px; min-height: {option_min_height}px; }}\n"
            f"QWidget#copyButtonItem QPushButton {{ padding: {copy_padding_v}px {copy_padding_h}px; min-height: {copy_min_height}px; }}\n"
            f"QTextEdit#previewTextEdit {{ padding: {preview_padding}px; border: none; background: {colors['color_preview_bg']}; color: {colors['color_text']}; }}\n"
            "QTextEdit#previewTextEdit:focus { border: none; }\n"
            "QTextEdit#previewTextEdit QFrame { border: none; }\n"
            "QPlainTextEdit, QListWidget, QTableWidget { padding: 2px; }"
        )


    def apply_ui_settings(self, settings: dict):
        """应用界面设置。

        旧版在这里直接调用 QApplication.setStyleSheet，并在同一轮事件里
        逐个刷新工具栏、选项行、复制按钮、预览框。部分控件已经被
        deleteLater() 标记但还残留在列表中时，点击“确定”会触发
        RuntimeError，严重时表现为闪退。
        
        现在改为：
        1. 先规范化设置；
        2. 只给主窗口设置样式，避免影响正在关闭的设置对话框；
        3. 逐类控件容错刷新，任何单个控件异常都不会导致主程序退出。
        """
        merged = self._normalize_ui_settings(settings)
        self._ui_settings = merged
        self._force_next_adaptive_refresh = True
        self._toolbar_metrics_sig = None
        try:
            app = QApplication.instance()
            if app is not None:
                app_font = app.font()
                app_font.setPointSize(scaled_point_size(merged, 'option_font_size', 10))
                app.setFont(app_font)
                app.setPalette(self._build_theme_palette(merged))
            self.setPalette(self._build_theme_palette(merged))
        except Exception:
            pass
        stylesheet = self._build_app_stylesheet(merged)
        try:
            self.setStyleSheet(stylesheet)
        except Exception:
            print('应用主窗口样式失败：')
            traceback.print_exc()
        self._apply_runtime_ui_settings()
        return merged

    @staticmethod
    def _safe_set_font(widget, point_size):
        if widget is None:
            return
        try:
            font = widget.font() if hasattr(widget, 'font') else QFont()
            font.setPointSize(int(point_size))
            widget.setFont(font)
        except RuntimeError:
            # 控件可能已经被 Qt 删除，仅跳过，避免设置窗口确认后闪退。
            return
        except Exception:
            return

    @staticmethod
    def _safe_set_min_height(widget, height):
        if widget is None:
            return
        try:
            widget.setMinimumHeight(int(height))
        except RuntimeError:
            return
        except Exception:
            return

    @staticmethod
    def _safe_set_min_width(widget, width):
        if widget is None:
            return
        try:
            widget.setMinimumWidth(int(width))
        except RuntimeError:
            return
        except Exception:
            return

    @staticmethod
    def _safe_set_fixed_height(widget, height):
        if widget is None:
            return
        try:
            widget.setFixedHeight(int(height))
        except RuntimeError:
            return
        except Exception:
            return

    def _apply_font_to_text_like_widget(self, widget, font):
        """同步文本类控件内部 document/view 的字体，避免子控件字号不一致。"""
        try:
            widget.setFont(font)
            if isinstance(widget, QTextEdit):
                widget.document().setDefaultFont(font)
                viewport = widget.viewport()
                if viewport is not None:
                    viewport.setFont(font)
            elif isinstance(widget, QPlainTextEdit):
                viewport = widget.viewport()
                if viewport is not None:
                    viewport.setFont(font)
            elif isinstance(widget, QComboBox):
                view = widget.view()
                if view is not None:
                    view.setFont(font)
                line_edit = widget.lineEdit() if widget.isEditable() else None
                if line_edit is not None:
                    line_edit.setFont(font)
                if hasattr(widget, 'sync_line_edit_area'):
                    widget.sync_line_edit_area()
        except RuntimeError:
            return
        except Exception:
            return

    def _apply_generic_window_ui_settings(self, window):
        """把统一字体缩放和基础自适应尺寸应用到子窗口。

        子窗口中的表单行按“窗口可用宽度 - 标签最大宽度 = 输入控件宽度”
        的思路处理，避免字号放大后标签和输入框互相挤压。
        """
        if window is None:
            return
        try:
            settings = self._normalize_ui_settings(getattr(self, '_ui_settings', self._default_ui_settings()))
            base_size = scaled_point_size(settings, 'option_font_size', 10)
            compact = int(settings.get('option_compactness', 4))
            font = QFont()
            font.setPointSize(base_size)
            colors = self._theme_colors(settings)
            try:
                window.setPalette(self._build_theme_palette(settings))
                window.setStyleSheet(
                    f"QMainWindow, QDialog {{ background-color: {colors['color_window']}; }}"
                    f"QWidget {{ color: {colors['color_text']}; }}"
                    f"QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QListWidget, QTableWidget, QTableView, QSpinBox, QDoubleSpinBox {{ background-color: {colors['color_input_bg']}; color: {colors['color_text']}; border: 1px solid {colors['color_border']}; }}"
                    f"QPushButton, QToolButton {{ background-color: {colors['color_button_bg']}; color: {colors['color_button_text']}; border: 1px solid {colors['color_border']}; border-radius: 3px; }}"
                    f"QPushButton:hover, QToolButton:hover {{ border: 1px solid {colors['color_primary']}; }}"
                )
            except Exception:
                pass
            widgets = [window] + list(window.findChildren(QWidget))
            min_height = adaptive_control_height(font, compact, min_height=22)
            text_min_height = adaptive_control_height(font, compact, min_height=60, lines=3)
            for widget in widgets:
                try:
                    self._apply_font_to_text_like_widget(widget, font)
                    if isinstance(widget, QLabel):
                        widget.setWordWrap(True)
                        widget.setMinimumHeight(max(widget.minimumHeight(), adaptive_control_height(font, compact, min_height=18)))
                    elif isinstance(widget, (QPushButton, QToolButton)):
                        widget.setMinimumHeight(min_height)
                        width = adaptive_text_width(font, widget.text(), compact, min_width=max(48, base_size * 4), max_width=max(260, base_size * 22))
                        widget.setMinimumWidth(width)
                    elif isinstance(widget, (QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QAbstractSpinBox, QCheckBox)):
                        widget.setMinimumHeight(min_height)
                        if isinstance(widget, QComboBox):
                            widget.setMinimumWidth(max(widget.minimumWidth(), base_size * 10))
                            try:
                                wrap_width = max(80, int(widget.width() or widget.minimumWidth()) - 34)
                            except Exception:
                                pass
                    elif isinstance(widget, (QTextEdit, QPlainTextEdit)):
                        widget.setMinimumHeight(max(widget.minimumHeight(), text_min_height))
                        try:
                            widget.setLineWrapMode(QTextEdit.WidgetWidth)
                        except Exception:
                            pass
                    elif isinstance(widget, (QTableView, QTableWidget)):
                        try:
                            widget.verticalHeader().setDefaultSectionSize(max(min_height, base_size * 2 + compact * 2))
                            widget.horizontalHeader().setMinimumHeight(min_height)
                        except Exception:
                            pass
                except RuntimeError:
                    continue
                except Exception:
                    continue

            # QFormLayout 子窗口表单宽度策略：标签最大为可用宽度 1/4，控件拿剩余宽度。
            # 对短标签设置稳定的最小/最大宽度，避免字号放大后被挤成两行；
            # 只有标签文本本身超过 1/4 宽度时才允许在标签内部换行。
            for form in list(window.findChildren(QFormLayout)):
                try:
                    form.setRowWrapPolicy(QFormLayout.DontWrapRows)
                    form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    form.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
                    form.setHorizontalSpacing(max(6, compact))
                    form.setVerticalSpacing(max(3, compact // 2 + 2))
                    parent_widget = form.parentWidget()
                    available_width = parent_widget.width() if parent_widget is not None and parent_widget.width() > 0 else window.width()
                    available_width = max(360, int(available_width) - 24)
                    label_max_width = max(base_size * 10, available_width // 3)
                    label_natural_widths = []
                    for row in range(form.rowCount()):
                        label_item = form.itemAt(row, QFormLayout.LabelRole)
                        label_widget = label_item.widget() if label_item is not None else None
                        if label_widget is not None and hasattr(label_widget, 'text'):
                            label_natural_widths.append(text_width_for(font, label_widget.text()) + compact * 2 + 16)
                    shared_label_width = min(label_max_width, max([base_size * 8] + label_natural_widths))
                    field_width = max(180, available_width - shared_label_width - max(12, compact * 3))
                    for row in range(form.rowCount()):
                        label_item = form.itemAt(row, QFormLayout.LabelRole)
                        field_item = form.itemAt(row, QFormLayout.FieldRole)
                        label_widget = label_item.widget() if label_item is not None else None
                        field_widget = field_item.widget() if field_item is not None else None
                        if label_widget is not None:
                            label_widget.setFont(font)
                            label_text = label_widget.text() if hasattr(label_widget, 'text') else ''
                            natural = text_width_for(font, label_text) + compact * 2 + 16
                            if isinstance(label_widget, QLabel):
                                label_widget.setWordWrap(natural > label_max_width)
                                label_widget.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                            label_widget.setMinimumWidth(int(shared_label_width))
                            label_widget.setMaximumWidth(int(shared_label_width))
                        if field_widget is not None:
                            field_widget.setMinimumWidth(field_width)
                            if isinstance(field_widget, (QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QAbstractSpinBox)):
                                field_widget.setMinimumHeight(min_height)
                            elif isinstance(field_widget, (QPlainTextEdit, QTextEdit)):
                                field_widget.setMinimumHeight(text_min_height)
                            elif isinstance(field_widget, QWidget):
                                for child in field_widget.findChildren(QWidget):
                                    child.setFont(font)
                                    if isinstance(child, (QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QAbstractSpinBox, QPushButton)):
                                        child.setMinimumHeight(min_height)
                except RuntimeError:
                    continue
                except Exception:
                    continue
        except RuntimeError:
            return
        except Exception:
            print('应用子窗口字体设置失败：')
            traceback.print_exc()

    def _refresh_open_child_windows_ui_settings(self):
        for window in (
            getattr(self, 'data_manager_window', None),
            getattr(self, 'template_editor_window', None),
            getattr(self, 'browser_flow_window', None),
        ):
            self._apply_generic_window_ui_settings(window)

    def _apply_runtime_ui_settings(self):
        settings = self._normalize_ui_settings(getattr(self, '_ui_settings', self._default_ui_settings()))

        toolbar_font_size = scaled_point_size(settings, 'toolbar_font_size', 10)
        toolbar_compact = int(settings.get('toolbar_compactness', 4))
        toolbar_font = QFont()
        toolbar_font.setPointSize(toolbar_font_size)
        toolbar_height = adaptive_control_height(toolbar_font, toolbar_compact, min_height=22)
        toolbar_sig = (
            toolbar_font_size,
            toolbar_compact,
            type(self.style()).__name__,
            tuple(button.text() for button in getattr(self, 'toolbar_buttons', []) or []),
        )
        refresh_toolbar_width = toolbar_sig != getattr(self, '_toolbar_metrics_sig', None)
        self._toolbar_metrics_sig = toolbar_sig
        if hasattr(self, 'main_toolbar'):
            self._safe_set_font(self.main_toolbar, toolbar_font_size)
            self._safe_set_fixed_height(self.main_toolbar, toolbar_height + max(2, toolbar_compact // 2))
            for button in list(getattr(self, 'toolbar_buttons', []) or []):
                try:
                    self._safe_set_font(button, toolbar_font_size)
                    button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                    button.setToolButtonStyle(Qt.ToolButtonTextOnly)
                    if refresh_toolbar_width:
                        if hasattr(button, 'invalidate_stable_size'):
                            button.invalidate_stable_size()
                        if hasattr(button, 'stable_size_hint'):
                            hint = button.stable_size_hint()
                        else:
                            button.ensurePolished()
                            hint = button.sizeHint()
                        button_width = max(int(hint.width()), int(button.minimumSizeHint().width()), max(56, toolbar_font_size * 4))
                        button.setFixedWidth(button_width)
                    self._safe_set_min_height(button, toolbar_height)
                    menu = button.menu() if hasattr(button, 'menu') else None
                    if menu is not None:
                        self._safe_set_font(menu, toolbar_font_size)
                        for action in menu.actions():
                            action.setFont(menu.font())
                except RuntimeError:
                    continue
                except Exception:
                    print('应用工具栏样式失败：')
                    traceback.print_exc()

        try:
            if hasattr(self, 'template_row') and self.template_row is not None:
                self.template_row.apply_ui_settings(settings)
        except RuntimeError:
            pass
        except Exception:
            print('应用模板行样式失败：')
            traceback.print_exc()

        for row in list(getattr(self, 'option_rows', []) or []):
            try:
                if row is not None and hasattr(row, 'apply_ui_settings'):
                    row.apply_ui_settings(settings)
            except RuntimeError:
                continue
            except Exception:
                print('应用选项行样式失败：')
                traceback.print_exc()

        # copy_buttons 正常结构是 [(CopyButtonItem, field_name), ...]。这里兼容
        # 旧中间版本可能遗留的 [CopyButtonItem, ...]，避免解包异常导致闪退。
        for entry in list(getattr(self, 'copy_buttons', []) or []):
            try:
                item_widget = entry[0] if isinstance(entry, (list, tuple)) else entry
                if item_widget is not None and hasattr(item_widget, 'apply_ui_settings'):
                    item_widget.apply_ui_settings(settings)
            except RuntimeError:
                continue
            except Exception:
                print('应用复制按钮样式失败：')
                traceback.print_exc()

        try:
            self._relayout_copy_buttons()
        except Exception:
            pass

        if hasattr(self, 'result_text'):
            try:
                preview_font_size = scaled_point_size(settings, 'preview_font_size', 11)
                preview_padding = settings.get('preview_compactness', 4)
                colors = self._theme_colors(settings)
                self._safe_set_font(self.result_text, preview_font_size)
                self.result_text.setFrameShape(QFrame.NoFrame)
                self.result_text.setLineWidth(0)
                self.result_text.setMidLineWidth(0)
                self.result_text.setStyleSheet(
                    'QTextEdit#previewTextEdit {'
                    f'padding: {preview_padding}px; '
                    f'border: 0px; background: {colors["color_preview_bg"]}; color: {colors["color_text"]};'
                    '} QTextEdit#previewTextEdit:focus { border: 0px; }'
                )
                viewport = self.result_text.viewport()
                if viewport is not None:
                    viewport.setStyleSheet(f'border: 0px; background: {colors["color_preview_bg"]};')
                    viewport.setAutoFillBackground(False)
            except RuntimeError:
                pass
            except Exception:
                print('应用预览框样式失败：')
                traceback.print_exc()

        self._refresh_open_child_windows_ui_settings()

    def _schedule_refresh_main_adaptive_rows(self):
        """合并滚动区 resize 信号，避免一次拖动分隔栏排队触发多轮布局计算。"""
        if getattr(self, '_refresh_adaptive_rows_pending', False):
            return
        self._refresh_adaptive_rows_pending = True

        def run_refresh():
            self._refresh_adaptive_rows_pending = False
            self._refresh_main_adaptive_rows()

        qt_safe_single_shot(60, run_refresh)

    def _refresh_main_adaptive_rows(self):
        """刷新首页输入行和复制项尺寸。

        每轮刷新都以 QScrollArea.viewport().width() 为唯一基准，并先清空上一轮
        setFixedWidth/setFixedHeight 产生的限制，避免拖动分隔栏时尺寸逐轮叠加。
        """
        try:
            raw_input_width = int(self.input_scroll.viewport().width()) if hasattr(self, 'input_scroll') else 0
            effective_input_width = max(raw_input_width, int(self.option_row_min_layout_width()))
            settings = self._normalize_ui_settings(getattr(self, '_ui_settings', self._default_ui_settings()))
            current_widths = (
                effective_input_width,
                int(self.copy_scroll.viewport().width()) if hasattr(self, 'copy_scroll') else 0,
                scaled_point_size(settings, 'option_font_size', 10),
                int(settings.get('option_compactness', 4)),
                int(settings.get('option_input_height', 96)),
                scaled_point_size(settings, 'copy_font_size', 10),
                int(settings.get('copy_compactness', 4)),
            )
            if getattr(self, '_last_adaptive_refresh_widths', None) == current_widths and not getattr(self, '_force_next_adaptive_refresh', False):
                return
            self._last_adaptive_refresh_widths = current_widths
            self._force_next_adaptive_refresh = False
        except Exception:
            pass
        widgets = [getattr(self, 'template_row', None)] + list(getattr(self, 'option_rows', []) or [])
        for widget in widgets:
            try:
                if widget is not None and hasattr(widget, 'reset_adaptive_constraints'):
                    widget.reset_adaptive_constraints()
            except RuntimeError:
                continue
            except Exception:
                continue
        try:
            if hasattr(self, 'input_scroll_panel') and self.input_scroll_panel is not None:
                viewport_width = self.input_scroll.viewport().width() if hasattr(self, 'input_scroll') else self.input_scroll_panel.width()
                effective_width = max(int(viewport_width or 0), int(self.option_row_min_layout_width()))
                self.input_scroll_panel.setMinimumWidth(0)
                self.input_scroll_panel.setMaximumWidth(16777215)
                if effective_width > 0:
                    self.input_scroll_panel.setFixedWidth(int(effective_width))
        except Exception:
            pass
        for widget in widgets:
            try:
                if widget is not None and hasattr(widget, 'update_adaptive_size'):
                    widget.update_adaptive_size()
            except RuntimeError:
                continue
            except Exception:
                continue
        for entry in list(getattr(self, 'copy_buttons', []) or []):
            try:
                item_widget = entry[0] if isinstance(entry, (list, tuple)) else entry
                if item_widget is not None and hasattr(item_widget, 'update_adaptive_size'):
                    item_widget.update_adaptive_size()
            except RuntimeError:
                continue
            except Exception:
                continue
        try:
            self._relayout_copy_buttons()
        except Exception:
            pass

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Resize:
            try:
                watched_targets = (self.input_scroll, self.input_scroll.viewport(), self.copy_scroll, self.copy_scroll.viewport())
                if watched in watched_targets:
                    self._schedule_refresh_main_adaptive_rows()
            except Exception:
                pass
        return super().eventFilter(watched, event)

    def normalize_main_browser_url_input(self):
        normalized = export.get_engine().normalize_url(self.browser_url_edit.text())
        if normalized != self.browser_url_edit.text().strip():
            self.browser_url_edit.setText(normalized)

    @staticmethod
    def _get_settings_path():
        """
        返回应用设置文件的路径。

        默认情况下将设置文件放在可执行文件或当前脚本所在目录下，名称为
        ``settings.json``。在打包模式下（PyInstaller)，会将文件保存在
        可执行文件同级目录中；在开发模式下，则存放在当前模块文件所在
        的目录。这样可以确保 JSON 设置文件与 exe/脚本文件位于同一目录，
        方便一并打包和分发。

        返回:
            str: 设置文件的绝对路径。
        """
        if getattr(sys, 'frozen', False):
            # 打包后的执行环境
            base_dir = os.path.dirname(sys.executable)
        else:
            # 开发环境，以当前文件所在目录为基准
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, 'settings.json')

    @staticmethod
    def _normalize_text(value):
        return '' if value is None else str(value)

    def _build_render_signature(self, template_name, input_vals, result_text, final_fields):
        payload = {'template': template_name, 'input': input_vals, 'result': result_text, 'final_fields': final_fields}
        return json.dumps(payload, ensure_ascii=False, sort_keys=True)

    def _get_active_process_content_override(self):
        if (
            self.template_editor_window is not None
            and self.current_template_name
            and self.template_editor_window.template_name == self.current_template_name
            and self._live_process_content is not None
        ):
            return self._live_process_content
        return None

    def _poll_live_updates(self):
        try:
            self.update_result_text()
        except Exception:
            pass

    def _set_widget_value(self, widget, value):
        text_value = self._normalize_text(value)
        if isinstance(widget, AutoWrapTextEdit):
            widget.setPlainTextPreserveSignal(text_value)
        elif isinstance(widget, QTextEdit):
            widget.setPlainText(text_value)
        elif isinstance(widget, QLineEdit):
            widget.setText(text_value)
        elif isinstance(widget, QComboBox):
            index = widget.findText(text_value)
            if index >= 0:
                widget.setCurrentIndex(index)
            elif widget.isEditable():
                widget.setCurrentText(text_value)
            else:
                if text_value and widget.findText(text_value) < 0:
                    widget.addItem(text_value)
                    widget.setCurrentText(text_value)
                else:
                    widget.setCurrentIndex(0)
        elif isinstance(widget, QCheckBox):
            widget.setChecked(str(value).lower() in ('true', '1', 'yes', 'checked'))

    def _set_widget_value_silent(self, widget, value):
        """设置控件值但不触发输入刷新，用于配置变动后恢复用户已输入内容。"""
        if widget is None:
            return
        try:
            with signal_blocked(widget):
                self._set_widget_value(widget, value)
        except Exception:
            self._set_widget_value(widget, value)

    def _remember_input_values(self):
        """保存当前输入区内容。除 clear_inputs 外，任何配置改动前都应调用。"""
        try:
            values = self.collect_input_values()
        except Exception:
            values = {}
        if values:
            self._input_value_cache.update(values)
        return dict(self._input_value_cache)

    def _restore_input_values(self, values=None, refresh_result=False):
        """按字段名恢复输入内容，避免拖拽、规则、模板、浏览器配置变动后清空输入。"""
        values = dict(values or self._input_value_cache or {})
        if not values:
            return
        for row in list(getattr(self, 'option_rows', []) or []):
            try:
                label = row.get_name()
                if label in values:
                    self._set_widget_value_silent(row.editor, values[label])
            except RuntimeError:
                continue
            except Exception:
                continue
        self._input_value_cache.update(values)
        if refresh_result:
            self.update_result_text(force=True)

    def on_live_process_template_changed(self, template_name, content):
        if template_name != self.current_template_name:
            return
        self._remember_input_values()
        self._live_process_content = content
        if self.browser_flow_window is not None and self.browser_flow_window.template_name == self.current_template_name:
            try:
                self.browser_flow_window.refresh_field_combo()
            except Exception:
                pass
        self.update_result_text(force=True)

    def on_external_data_changed(self):
        preserved = self._remember_input_values()
        self.refresh_input_area(preserved)
        self._restore_input_values(preserved)
        self.update_result_text(force=True)

    def on_data_manager_closed(self):
        self.data_manager_window = None

    def init_ui(self):
        self.setWindowTitle(f'PEditor_v{__version__}')
        self._init_window_geometry()
        self.create_toolbar()
        self.create_main_area()

    def _new_toolbar_action(self, menu, text, handler=None, attr_name=None, checkable=False):
        action = QAction(text, self)
        action.setCheckable(checkable)
        if handler is not None:
            if checkable:
                action.toggled.connect(handler)
            else:
                action.triggered.connect(lambda checked=False, h=handler: h())
        menu.addAction(action)
        if attr_name:
            setattr(self, attr_name, action)
        return action

    def _add_toolbar_menu(self, toolbar, title, action_items):
        tool_btn = ToolbarMenuButton(self)
        tool_btn.setObjectName('toolbarMenuButton')
        tool_btn.setText(title)
        tool_btn.setPopupMode(QToolButton.InstantPopup)
        tool_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        menu = QMenu(tool_btn)
        menu.setObjectName('toolbarMenu')
        for item in action_items:
            if item is None:
                menu.addSeparator()
                continue
            text, handler, attr_name = item
            self._new_toolbar_action(menu, text, handler, attr_name)
        tool_btn.setMenu(menu)
        toolbar.addWidget(tool_btn)
        if hasattr(self, 'toolbar_buttons'):
            self.toolbar_buttons.append(tool_btn)
        return tool_btn

    def create_toolbar(self):
        toolbar = QToolBar('工具栏')
        toolbar.setObjectName('mainToolbar')
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self.main_toolbar = toolbar
        self.toolbar_buttons = []
        self.addToolBar(toolbar)

        # 这些 QLineEdit 只作为浏览器参数的状态载体，必须隐藏；否则会在主界面左上角形成无意义输入框。
        self.browser_driver_edit = QLineEdit()
        self.browser_binary_edit = QLineEdit()
        self.browser_url_edit = QLineEdit()
        for hidden_edit in (self.browser_driver_edit, self.browser_binary_edit, self.browser_url_edit):
            hidden_edit.hide()
        self.browser_driver_edit.editingFinished.connect(self.on_main_browser_settings_changed)
        self.browser_binary_edit.editingFinished.connect(self.on_main_browser_settings_changed)
        self.browser_url_edit.editingFinished.connect(self.normalize_main_browser_url_input)
        self.browser_url_edit.editingFinished.connect(self.on_main_browser_settings_changed)

        self._add_toolbar_menu(toolbar, '文件', [
            ('新建模板', self.new_template, 'new_btn'),
            ('重命名模板', self.rename_template, 'rename_btn'),
            ('删除模板', self.delete_template, 'del_btn'),
            None,
            ('导入模板', self.import_template, 'import_btn'),
            ('导出模板', self.export_template, 'export_btn'),
            ('导出工艺TXT', self.save_to_summary_file, 'save_btn'),
            None,
            ('退出', self.close, 'exit_btn'),
        ])

        self._add_toolbar_menu(toolbar, '编辑', [
            ('添加选项行', self.add_option_row, 'add_option_row_action'),
            ('删除选中选项行', self.delete_option_row, 'delete_option_row_action'),
            ('编辑选中选项名称', self.edit_option_row, 'edit_option_row_action'),
            None,
            ('输入项配置', self.edit_options, 'edit_opt_btn'),
            ('工序模板', self.open_template_editor, 'edit_field_btn'),
            ('编辑规则', self.edit_rules, 'edit_rule_btn'),
            ('数据库', self.open_data_manager, 'db_btn'),
            None,
            ('清空输入', self.clear_inputs, 'clear_btn'),
        ])

        self._add_toolbar_menu(toolbar, '浏览器', [
            ('浏览器参数', self.open_browser_settings_dialog, 'browser_settings_btn'),
            ('打开浏览器', self.open_browser_from_main, 'open_browser_btn'),
            ('浏览器配置', self.open_browser_flow_editor, 'browser_cfg_btn'),
            ('导出至浏览器', self.export_current_to_browser, 'browser_export_btn'),
        ])

        self.copy_tool_btn = ToolbarMenuButton(self)
        self.copy_tool_btn.setObjectName('toolbarMenuButton')
        self.copy_tool_btn.setText('复制选项编辑')
        self.copy_tool_btn.setPopupMode(QToolButton.InstantPopup)
        self.copy_tool_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        copy_menu = QMenu(self.copy_tool_btn)
        copy_menu.setObjectName('toolbarMenu')
        self._new_toolbar_action(copy_menu, '添加工序', self.show_add_copy_button_menu, 'add_copy_btn')
        self._new_toolbar_action(copy_menu, '删除选中', self.delete_selected_copy_button, 'del_copy_btn')
        copy_menu.addSeparator()
        self.copy_multi_check = self._new_toolbar_action(copy_menu, '多选复制', self.on_copy_multi_mode_changed, checkable=True)
        self.copy_tool_btn.setMenu(copy_menu)
        toolbar.addWidget(self.copy_tool_btn)
        self.toolbar_buttons.append(self.copy_tool_btn)

        self._add_toolbar_menu(toolbar, '帮助', [
            ('界面设置', self.open_settings, 'settings_btn'),
            ('教程', self.open_tutorial, 'tutorial_btn'),
            ('日志', self.open_log_viewer, 'log_btn'),
        ])

    def create_main_area(self):
        splitter1 = QSplitter(Qt.Vertical)
        splitter2 = QSplitter(Qt.Horizontal)
        # QSplitter 是 GTE 主界面的拖动分割边框控件，用户可直接拖动边框调整区域大小。
        splitter1.setHandleWidth(8)
        splitter2.setHandleWidth(8)
        splitter1.setChildrenCollapsible(False)
        splitter2.setChildrenCollapsible(False)

        left_panel = self.create_left_panel()
        right_panel = self.create_right_panel()
        down_panel = self.create_down_panel()

        splitter2.addWidget(left_panel)
        splitter2.addWidget(right_panel)
        splitter1.addWidget(splitter2)
        splitter1.addWidget(down_panel)
        splitter1.setSizes([700, 260])
        splitter2.setSizes([500, 500])
        self.setCentralWidget(splitter1)
        self.splitter = splitter1
        self.top_splitter = splitter2

    def _make_panel(self, object_name='gtePanel', white=False):
        panel = QFrame()
        panel.setObjectName(object_name)
        # 面板颜色由 _build_app_stylesheet() 统一控制，避免主题设置被局部样式覆盖。
        return panel

    def create_left_panel(self):
        panel = self._make_panel('gteLeftPanel')
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.input_scroll = QScrollArea()
        self.input_scroll.setWidgetResizable(True)
        self.input_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.input_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.input_scroll.setFrameShape(QFrame.NoFrame)

        self.input_scroll_panel = QFrame()
        self.input_scroll_panel.setFrameShape(QFrame.NoFrame)
        self.input_scroll_panel.setMinimumWidth(0)
        self.input_scroll_panel.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Minimum)
        self.input_layout = QVBoxLayout(self.input_scroll_panel)
        self.input_layout.setContentsMargins(5, 5, 5, 5)
        self.input_layout.setSpacing(5)

        self.template_combo = QComboBox()
        self.template_combo.currentTextChanged.connect(self.on_template_changed)
        self.template_row = TemplateSelectRow(self.template_combo, self)
        self.input_layout.addWidget(self.template_row)
        self.input_layout.addStretch()

        self.input_scroll.setWidget(self.input_scroll_panel)
        try:
            self.input_scroll.viewport().installEventFilter(self)
            self.input_scroll.installEventFilter(self)
        except Exception:
            pass
        layout.addWidget(self.input_scroll)
        return panel

    def create_right_panel(self):
        panel = self._make_panel('gteRightPanel', white=True)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        self.result_text = QTextEdit()
        self.result_text.setObjectName('previewTextEdit')
        self.result_text.setFrameShape(QFrame.NoFrame)
        self.result_text.setStyleSheet('border: none;')
        self.result_text.setReadOnly(True)
        self.result_text.setPlaceholderText('工艺段落将实时显示在此...')
        self.result_text.setLineWrapMode(QTextEdit.WidgetWidth)
        layout.addWidget(self.result_text)
        return panel

    def create_down_panel(self):
        panel = self._make_panel('gteDownPanel')
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        self.copy_scroll = QScrollArea()
        # 复制按钮区按视口宽度自动换行，横向滚动关闭，超出高度时使用纵向滚动查看。
        self.copy_scroll.setWidgetResizable(True)
        self.copy_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.copy_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.copy_scroll.setFrameShape(QFrame.NoFrame)
        self.copy_scroll_widget = QFrame()
        self.copy_scroll_widget.setFrameShape(QFrame.NoFrame)
        self.copy_scroll_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        # 外层用纵向行容器，每一行内部单独横向排列，避免 QGridLayout 造成上下行列宽对齐。
        self.copy_buttons_layout = QVBoxLayout(self.copy_scroll_widget)
        self.copy_buttons_layout.setContentsMargins(5, 5, 5, 5)
        self.copy_buttons_layout.setSpacing(6)
        self.copy_buttons_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self._copy_row_widgets = []
        self.copy_scroll.setWidget(self.copy_scroll_widget)
        try:
            self.copy_scroll.viewport().installEventFilter(self)
        except Exception:
            pass
        layout.addWidget(self.copy_scroll)
        return panel

    def open_browser_settings_dialog(self):
        preserved = self.collect_input_values()
        browser = self._get_browser_settings_for_current_template() if self.current_template_name else BrowserFlowWindow._default_browser()
        dlg = BrowserSettingsDialog(browser, self)
        self._apply_generic_window_ui_settings(dlg)
        if dlg.exec_() != QDialog.Accepted:
            return
        self._input_value_cache.update(preserved)
        settings = dlg.get_browser_settings()
        self.browser_driver_edit.setText(settings.get('chromedriver_path', ''))
        self.browser_binary_edit.setText(settings.get('chrome_binary', ''))
        self.browser_url_edit.setText(settings.get('start_url', ''))
        self.on_main_browser_settings_changed()


    def _get_browser_flow_for_current_template(self):
        if not self.current_template_name:
            return None
        flow = self.template_db.get_browser_flow(self.current_template_name)
        if not flow:
            flow = {'browser': BrowserFlowWindow._default_browser(), 'steps': []}
            self.template_db.update_browser_flow(self.current_template_name, flow)
        flow.setdefault('browser', BrowserFlowWindow._default_browser())
        flow.setdefault('steps', [])
        return flow

    def _get_browser_settings_for_current_template(self):
        flow = self._get_browser_flow_for_current_template()
        if not flow:
            return BrowserFlowWindow._default_browser()
        browser = BrowserFlowWindow._default_browser()
        browser.update(flow.get('browser', {}) or {})
        return browser

    def _load_browser_settings_to_main(self):
        browser = self._get_browser_settings_for_current_template() if self.current_template_name else BrowserFlowWindow._default_browser()
        for widget, value in (
            (self.browser_driver_edit, browser.get('chromedriver_path', '')),
            (self.browser_binary_edit, browser.get('chrome_binary', '')),
            (self.browser_url_edit, browser.get('start_url', '')),
        ):
            with signal_blocked(widget):
                widget.setText(self._normalize_text(value))

    def _save_browser_settings_for_current_template(self, silent=True):
        if not self.current_template_name:
            return
        flow = self._get_browser_flow_for_current_template() or {'browser': BrowserFlowWindow._default_browser(), 'steps': []}
        browser = BrowserFlowWindow._default_browser()
        browser.update(flow.get('browser', {}) or {})
        browser.update({
            'connect_mode': 'launch',
            'chromedriver_path': self.browser_driver_edit.text().strip(),
            'chrome_binary': self.browser_binary_edit.text().strip(),
            'start_url': export.get_engine().normalize_url(self.browser_url_edit.text()),
        })
        flow['browser'] = browser
        self.template_db.update_browser_flow(self.current_template_name, flow)
        if self.browser_flow_window is not None and self.browser_flow_window.template_name == self.current_template_name:
            try:
                self.browser_flow_window.apply_external_browser_settings(browser)
            except Exception:
                pass
        if not silent:
            QMessageBox.information(self, '成功', '浏览器设置已保存到当前模板。')

    def on_main_browser_settings_changed(self):
        self.normalize_main_browser_url_input()
        self._save_browser_settings_for_current_template(silent=True)

    def open_browser_from_main(self):
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择模板。')
            return
        self._save_browser_settings_for_current_template(silent=True)
        try:
            browser = self._get_browser_settings_for_current_template()
            export.get_engine().launch_browser(
                chromedriver_path=browser.get('chromedriver_path', ''),
                chrome_binary=browser.get('chrome_binary', ''),
                start_url=browser.get('start_url', ''),
                debug_port=int(browser.get('debug_port', 9222) or 9222),
            )
            QMessageBox.information(self, '成功', '浏览器已打开。')
            if self.browser_flow_window is not None and self.browser_flow_window.template_name == self.current_template_name:
                self.browser_flow_window.refresh_windows()
        except Exception as e:
            QMessageBox.critical(self, '错误', f'打开浏览器失败：{e}')

    def load_settings(self):
        settings = self._load_settings_payload()
        self.apply_ui_settings(settings)

    def save_settings(self, settings):
        payload = self._normalize_ui_settings(settings)
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f'保存设置失败：{e}')

    def open_settings(self):
        current = dict(getattr(self, '_ui_settings', self._default_ui_settings()))
        dlg = UiSettingsDialog(current, self)
        self._apply_generic_window_ui_settings(dlg)
        if dlg.exec_() != QDialog.Accepted:
            return
        try:
            settings = self._normalize_ui_settings(dlg.get_settings())
            self.apply_ui_settings(settings)
            self.save_settings(settings)
        except Exception as e:
            # 避免设置窗口确认后异常冒泡导致主程序退出。
            print('应用界面设置失败：')
            traceback.print_exc()
            QMessageBox.critical(self, '界面设置错误', f'应用界面设置失败：{e}')

    def open_tutorial(self):
        tutorial_text = ''
        tutorial_path = resource_path('PEditor_教程.txt')
        try:
            if os.path.exists(tutorial_path):
                with open(tutorial_path, 'r', encoding='utf-8') as f:
                    tutorial_text = f.read()
        except Exception:
            tutorial_text = ''
        if not tutorial_text.strip():
            tutorial_text = '未找到教程文件 PEditor_教程.txt，请确认该文件与程序一起提供。'

        dlg = QDialog(self)
        dlg.setWindowTitle('PEditor 使用教程')
        dlg.resize(900, 680)
        layout = QVBoxLayout(dlg)
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(tutorial_text)
        layout.addWidget(text_edit)
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok)
        btn_box.accepted.connect(dlg.accept)
        layout.addWidget(btn_box)
        self._apply_generic_window_ui_settings(dlg)
        dlg.exec_()


    def open_log_viewer(self):
        dlg = LogViewerDialog(self)
        self._apply_generic_window_ui_settings(dlg)
        dlg.exec_()

    def _show_browser_alert(self, message, level='info'):
        text = str(message or '').strip()
        if not text:
            return
        if str(level).lower() in ('warning', 'timeout', 'error'):
            QMessageBox.warning(self, '浏览器提示', text)
        else:
            QMessageBox.information(self, '浏览器提示', text)

    def load_template_list(self):
        previous_values = self.collect_input_values() if getattr(self, 'option_rows', None) else dict(getattr(self, '_input_value_cache', {}) or {})
        previous = self.current_template_name
        temps = self.template_db.get_main_templates()
        with signal_blocked(self.template_combo):
            self.template_combo.clear()
            for t in temps:
                self.template_combo.addItem(t['name'], t)

        self._live_process_content = None
        if temps:
            target_name = previous if previous and any(t['name'] == previous for t in temps) else temps[0]['name']
            with signal_blocked(self.template_combo):
                self.template_combo.setCurrentText(target_name)
            self._load_template_by_name(target_name)
        else:
            self.current_template_name = None
            self.current_options_config = []
            self.current_rules_config = []
            self.result_text.clear()

        self.refresh_input_area(previous_values)
        self._load_browser_settings_to_main()
        self.update_copy_buttons_from_config()
        self.update_result_text(force=True)

    def _load_template_by_name(self, name):
        self.current_template_name = name
        index = self.template_combo.findText(name)
        data = self.template_combo.itemData(index) if index >= 0 else None
        if data:
            config = data.get('config', {}) or {}
            self.current_options_config = config.get('options', [])
            self.current_rules_config = config.get('rules', [])
        else:
            self.current_options_config = []
            self.current_rules_config = []
        if self.browser_flow_window is not None:
            self.browser_flow_window.set_template_name(name)

    def on_template_changed(self, name):
        if not name:
            self.current_template_name = None
            self.current_options_config = []
            self.current_rules_config = []
            self.refresh_input_area()
            self._load_browser_settings_to_main()
            self.update_copy_buttons_from_config()
            self.result_text.clear()
            return

        self._live_process_content = None
        preserved = self.collect_input_values() if getattr(self, 'option_rows', None) else dict(getattr(self, '_input_value_cache', {}) or {})
        self._load_template_by_name(name)
        self.refresh_input_area(preserved)
        self._load_browser_settings_to_main()
        self.update_copy_buttons_from_config()
        self.update_result_text(force=True)

    def new_template(self):
        name, ok = QInputDialog.getText(self, '新建模板', '名称:')
        if ok and name.strip():
            name = name.strip()
            if any(t['name'] == name for t in self.template_db.get_main_templates()):
                QMessageBox.warning(self, '错误', '名称已存在')
                return
            config = {'options': [], 'rules': [], 'copy_buttons': []}
            self.template_db.add_main_template(name, config)
            self.template_db.add_process_template(name, {'available_fields': {}, 'available_field_names': [], 'selected_fields': [], 'preview_format': '', 'field_conditions': {}})
            self.template_db.update_browser_flow(name, {'browser': BrowserFlowWindow._default_browser(), 'steps': []})
            self.load_template_list()
            self.template_combo.setCurrentText(name)

    def rename_template(self):
        if not self.current_template_name:
            return
        new, ok = QInputDialog.getText(self, '重命名', '新名称:', text=self.current_template_name)
        if ok and new.strip() and new != self.current_template_name:
            new = new.strip()
            if any(t['name'] == new for t in self.template_db.get_main_templates()):
                QMessageBox.warning(self, '错误', '名称已存在')
                return
            old_name = self.current_template_name
            config = {'options': self.current_options_config, 'rules': self.current_rules_config, 'copy_buttons': self.get_current_copy_button_fields()}
            self.template_db.update_main_template(old_name, new, config)
            proc = self.template_db.get_process_template(old_name)
            if proc:
                self.template_db.add_process_template(new, proc['content'])
                self.template_db.delete_process_template(old_name)
            # The browser flow record is automatically renamed by
            # TemplateDB.update_main_template via rename_browser_flow.  The
            # previous implementation manually copied and deleted the row,
            # which could leave duplicate entries when run multiple times.
            # Because rename_browser_flow already updates the flow's
            # template_name, there is no need to duplicate this logic here.
            self.current_template_name = new
            self.load_template_list()
            self.template_combo.setCurrentText(new)

    def delete_template(self):
        if not self.current_template_name:
            return
        if QMessageBox.question(self, '确认', f"删除模板 '{self.current_template_name}'？") == QMessageBox.Yes:
            self.template_db.delete_main_template(self.current_template_name)
            self.template_db.delete_process_template(self.current_template_name)
            self.template_db.delete_browser_flow(self.current_template_name)
            self.current_template_name = None
            self.load_template_list()

    def import_template(self):
        path, _ = QFileDialog.getOpenFileName(self, '导入', '', 'JSON (*.json)')
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                imported = json.load(f)
            if 'main_template' in imported or 'process_template' in imported or 'browser_flow' in imported:
                main_config = imported.get('main_template', {}) or {}
                process_content = imported.get('process_template', {}) or {}
                browser_flow = imported.get('browser_flow', {}) or {'browser': BrowserFlowWindow._default_browser(), 'steps': []}
            else:
                main_config = imported
                process_content = {'available_fields': {}, 'available_field_names': [], 'selected_fields': [], 'preview_format': '', 'field_conditions': {}}
                browser_flow = {'browser': BrowserFlowWindow._default_browser(), 'steps': []}

            name = os.path.splitext(os.path.basename(path))[0]
            name, ok = QInputDialog.getText(self, '名称', '模板名称:', text=name)
            if ok and name.strip():
                name = name.strip()
                if any(t['name'] == name for t in self.template_db.get_main_templates()):
                    QMessageBox.warning(self, '错误', '名称已存在')
                    return
                main_config.setdefault('options', [])
                main_config.setdefault('rules', [])
                main_config.setdefault('copy_buttons', [])
                process_content.setdefault('available_fields', {})
                process_content.setdefault('available_field_names', [])
                process_content.setdefault('selected_fields', [])
                process_content.setdefault('preview_format', '')
                process_content.setdefault('field_conditions', {})
                browser_flow.setdefault('browser', BrowserFlowWindow._default_browser())
                browser_flow.setdefault('steps', [])
                self.template_db.add_main_template(name, main_config)
                self.template_db.add_process_template(name, process_content)
                self.template_db.update_browser_flow(name, browser_flow)
                self.load_template_list()
                self.template_combo.setCurrentText(name)
        except Exception as e:
            QMessageBox.critical(self, '错误', str(e))

    def export_template(self):
        if not self.current_template_name:
            return
        main_config = {'options': self.current_options_config, 'rules': self.current_rules_config, 'copy_buttons': self.get_current_copy_button_fields()}
        process_tpl = self.template_db.get_process_template(self.current_template_name)
        process_content = process_tpl.get('content', {}) if process_tpl else {}
        browser_flow = self.template_db.get_browser_flow(self.current_template_name) or {}
        payload = {'main_template': main_config, 'process_template': process_content, 'browser_flow': browser_flow}
        path, _ = QFileDialog.getSaveFileName(self, '导出', f'{self.current_template_name}.json', 'JSON (*.json)')
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

    def edit_options(self):
        preserved = self._remember_input_values()
        dlg = OptionEditDialog(self.current_options_config, self.db, self)
        self._apply_generic_window_ui_settings(dlg)
        if dlg.exec_() == QDialog.Accepted:
            self.current_options_config = dlg.options_config
            self.refresh_input_area(preserved)
            self._save_current_template()
            self._restore_input_values(preserved)
            if self.browser_flow_window is not None and self.browser_flow_window.template_name == self.current_template_name:
                try:
                    self.browser_flow_window.refresh_field_combo()
                except Exception:
                    pass
            self.update_result_text(force=True)

    def edit_rules(self):
        field_pool = self._build_rule_field_pool()
        dlg = RuleManagerDialog(self.current_rules_config, self.db, field_pool, self)
        self._apply_generic_window_ui_settings(dlg)
        preserved = self._remember_input_values()
        if dlg.exec_() == QDialog.Accepted:
            self._input_value_cache.update(preserved)
            self.current_rules_config = dlg.get_rules()
            self._save_current_template()
            self._restore_input_values(preserved)
            if self.browser_flow_window is not None and self.browser_flow_window.template_name == self.current_template_name:
                try:
                    self.browser_flow_window.refresh_field_combo()
                except Exception:
                    pass
            self.update_result_text(force=True)

    def _build_rule_field_pool(self):
        process_override = self._get_active_process_content_override()
        input_labels = [opt.get('label', '') for opt in self.current_options_config]
        return self.data_matcher.get_rule_field_pool(
            self.current_template_name,
            input_option_labels=input_labels,
            main_config_override={'options': self.current_options_config, 'rules': self.current_rules_config},
            process_content_override=process_override,
        )

    def _sync_current_template_combo_item(self, config=None):
        if not self.current_template_name:
            return
        index = self.template_combo.findText(self.current_template_name)
        if index < 0:
            return
        payload = self.template_combo.itemData(index) or {'name': self.current_template_name, 'config': {}}
        payload = dict(payload)
        payload['name'] = self.current_template_name
        payload['config'] = config if config is not None else payload.get('config', {})
        self.template_combo.setItemData(index, payload)

    def _save_current_template(self):
        if not self.current_template_name:
            return
        before = self.template_db.get_main_template(self.current_template_name) or {}
        before_config = (before.get('config', {}) if before else {}) or {}
        config = {
            'options': self.current_options_config,
            'rules': self.current_rules_config,
            'copy_buttons': self._sanitize_copy_button_fields(self.get_current_copy_button_fields()),
        }
        self.template_db.update_main_template(self.current_template_name, self.current_template_name, config)
        self._sync_current_template_combo_item(config)
        if before_config != config:
            log_change(f'主模板修改 - {self.current_template_name}', before=before_config, after=config)

    def _clear_input_option_rows(self):
        self.selected_option_row = None
        for row in list(self.option_rows):
            self.input_layout.removeWidget(row)
            row.deleteLater()
        self.option_rows.clear()
        self.input_widgets.clear()

    def refresh_input_area(self, preserved_values=None):
        # 除非用户明确点击“清空输入”，否则重建输入区时优先保留当前已输入内容。
        if preserved_values is None:
            try:
                preserved_values = self.collect_input_values()
            except Exception:
                preserved_values = dict(getattr(self, '_input_value_cache', {}) or {})
        preserved_values = dict(preserved_values or {})
        if not preserved_values and getattr(self, '_input_value_cache', None):
            preserved_values = dict(self._input_value_cache)
        if preserved_values:
            self._input_value_cache.update(preserved_values)
        self._clear_input_option_rows()
        opts = sorted(self.current_options_config, key=lambda x: x.get('order', 0))
        initial_values = dict(preserved_values or {})
        for opt in opts:
            widget = self._create_widget(opt, initial_values)
            if preserved_values and opt.get('label') in preserved_values:
                self._set_widget_value(widget, preserved_values.get(opt.get('label')))
            row = InputOptionRow(opt, widget, self)
            insert_index = max(1, self.input_layout.count() - 1)
            self.input_layout.insertWidget(insert_index, row)
            self.option_rows.append(row)
            self.input_widgets.append(widget)
        self.setup_input_change_tracking()
        self._apply_runtime_ui_settings()
        self.refresh_dynamic_combo_options()
        try:
            qt_safe_single_shot(0, self._refresh_main_adaptive_rows)
            qt_safe_single_shot(80, self._refresh_main_adaptive_rows)
        except Exception:
            pass

    def _create_widget(self, opt, initial_values=None):
        widget_type = opt.get('type', 'text')
        if widget_type == 'text':
            editor = AutoWrapTextEdit()
            editor.setPlaceholderText('')
            return editor
        if widget_type in ('combo', 'editable_combo'):
            combo = SafeEditableComboBox() if widget_type == 'editable_combo' else QComboBox()
            combo.addItem('')
            options = self.data_matcher.get_field_options(opt.get('source', {}), input_values=initial_values or {})
            combo.addItems(options if options else ['（无选项）'])
            combo.setCurrentIndex(0)
            try:
                combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            except Exception:
                pass
            return combo
        if widget_type == 'checkbox':
            return QCheckBox()
        return AutoWrapTextEdit()

    def setup_input_change_tracking(self):
        for widget in self.input_widgets:
            if getattr(widget, '_pe_input_tracking_bound', False):
                continue
            if isinstance(widget, QTextEdit):
                widget.textChanged.connect(self.on_input_widget_changed)
            elif isinstance(widget, QLineEdit):
                widget.textChanged.connect(self.on_input_widget_changed)
            elif isinstance(widget, QComboBox):
                widget.currentTextChanged.connect(self.on_input_widget_changed)
            elif isinstance(widget, QCheckBox):
                widget.stateChanged.connect(self.on_input_widget_changed)
            try:
                widget._pe_input_tracking_bound = True
            except Exception:
                pass

    def clear_selected_option_row(self):
        if self.selected_option_row is not None:
            self.selected_option_row.set_selected(False)
        self.selected_option_row = None

    def select_option_row(self, row, toggle=True):
        if row is None:
            return
        if self.selected_option_row is row:
            if toggle:
                row.set_selected(False)
                self.selected_option_row = None
            else:
                row.set_selected(True)
            return
        if self.selected_option_row is not None:
            self.selected_option_row.set_selected(False)
        self.selected_option_row = row
        self.selected_option_row.set_selected(True)

    def add_option_row(self):
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择或新建模板。')
            return
        base = '新选项'
        existing = {str(opt.get('label', '')).strip() for opt in self.current_options_config}
        index = 1
        name = base
        while name in existing:
            index += 1
            name = f'{base}{index}'
        preserved = self.collect_input_values()
        self.current_options_config.append({'label': name, 'type': 'text', 'order': len(self.current_options_config), 'source': {}})
        self.refresh_input_area(preserved)
        if self.option_rows:
            self.select_option_row(self.option_rows[-1], toggle=False)
        self._save_current_template()
        self.update_result_text(force=True)

    def delete_option_row(self):
        if self.selected_option_row is None:
            QMessageBox.warning(self, '提示', '请先选中一个选项行。')
            return
        row = self.selected_option_row
        label = row.get_name()
        reply = QMessageBox.question(self, '确认删除', f'确定删除选项“{label}”吗？')
        if reply != QMessageBox.Yes:
            return
        preserved = self.collect_input_values()
        self.current_options_config = [opt for opt in self.current_options_config if opt is not row.option_config]
        preserved.pop(label, None)
        self.refresh_input_area(preserved)
        self.sync_option_order_from_rows(save=True)
        self.update_result_text(force=True)

    def edit_option_row(self):
        if self.selected_option_row is None:
            QMessageBox.warning(self, '提示', '请先选中一个选项行。')
            return
        row = self.selected_option_row
        old_name = row.get_name()
        new_name, ok = QInputDialog.getText(self, '编辑选项', '请输入新的选项名称：', text=old_name)
        if not ok:
            return
        new_name = new_name.strip()
        if not new_name:
            QMessageBox.warning(self, '提示', '选项名称不能为空。')
            return
        if new_name != old_name and any(opt.get('label') == new_name for opt in self.current_options_config):
            QMessageBox.warning(self, '提示', f'选项“{new_name}”已存在。')
            return
        preserved = self.collect_input_values()
        if old_name in preserved:
            preserved[new_name] = preserved.pop(old_name)
        row.set_name(new_name)
        self.refresh_input_area(preserved)
        for option_row in self.option_rows:
            if option_row.get_name() == new_name:
                self.select_option_row(option_row, toggle=False)
                break
        self._save_current_template()
        self.update_result_text(force=True)

    def find_option_row_at_pos(self, global_pos):
        widget = QApplication.widgetAt(global_pos)
        while widget is not None:
            if isinstance(widget, InputOptionRow):
                return widget
            widget = widget.parentWidget()
        return None

    def move_option_row_by_mouse(self, row, global_pos):
        target_row = self.find_option_row_at_pos(global_pos)
        if target_row is None or target_row is row:
            return
        row_index = self.input_layout.indexOf(row)
        target_index = self.input_layout.indexOf(target_row)
        if row_index < 0 or target_index < 0:
            return
        local_pos = target_row.mapFromGlobal(global_pos)
        if local_pos.y() < target_row.height() / 2:
            new_index = target_index
        else:
            new_index = target_index + 1
        if row_index < new_index:
            new_index -= 1
        max_index = self.input_layout.count() - 2
        if new_index < 0:
            new_index = 0
        if new_index > max_index:
            new_index = max_index
        if new_index == row_index:
            return
        self.input_layout.removeWidget(row)
        self.input_layout.insertWidget(new_index, row)
        self.select_option_row(row, toggle=False)

    def sync_option_order_from_rows(self, save=False):
        rows = []
        for index in range(self.input_layout.count()):
            widget = self.input_layout.itemAt(index).widget()
            if isinstance(widget, InputOptionRow):
                rows.append(widget)
        self.option_rows = rows
        self.input_widgets = [row.editor for row in rows]
        self.current_options_config = [row.option_config for row in rows]
        for order, opt in enumerate(self.current_options_config):
            opt['order'] = order
        if save and self.current_template_name:
            preserved = self._remember_input_values()
            self._save_current_template()
            self._restore_input_values(preserved)
            self.update_result_text(force=True)


    def on_input_widget_changed(self, *args):
        source_widget = self.sender()
        if self._updating_option_sources:
            try:
                self._remember_input_values()
            except Exception:
                pass
            self.update_result_text()
            return
        try:
            self._remember_input_values()
        except Exception:
            pass
        self.refresh_dynamic_combo_options(source_widget=source_widget)
        try:
            self._remember_input_values()
        except Exception:
            pass
        self.update_result_text()

    @staticmethod
    def _editable_combo_is_user_typing(widget, source_widget=None):
        if not isinstance(widget, QComboBox) or not widget.isEditable():
            return False
        line_edit = widget.lineEdit()
        try:
            return widget is source_widget or line_edit is source_widget or bool(line_edit and line_edit.hasFocus())
        except RuntimeError:
            return False

    def _update_combo_items_preserve_text(self, widget, normalized_options, current_text):
        """刷新下拉项，同时保留可输入下拉框当前文本和光标位置。"""
        line_edit = widget.lineEdit() if isinstance(widget, QComboBox) and widget.isEditable() else None
        cursor_pos = None
        selection_start = -1
        selection_length = 0
        if line_edit is not None:
            try:
                cursor_pos = line_edit.cursorPosition()
                if line_edit.hasSelectedText():
                    selection_start = line_edit.selectionStart()
                    selection_length = len(line_edit.selectedText())
            except RuntimeError:
                line_edit = None

        blockers = []
        try:
            blockers.append(signal_blocked(widget))
            blockers[-1].__enter__()
            if line_edit is not None:
                blockers.append(signal_blocked(line_edit))
                blockers[-1].__enter__()
            widget.clear()
            widget.addItems(normalized_options)
            if widget.isEditable():
                widget.setEditText(current_text)
                if line_edit is not None:
                    line_edit.setText(current_text)
                    if selection_start >= 0 and selection_length > 0:
                        line_edit.setSelection(selection_start, selection_length)
                    elif cursor_pos is not None:
                        line_edit.setCursorPosition(max(0, min(int(cursor_pos), len(current_text))))
            else:
                if current_text in normalized_options:
                    widget.setCurrentText(current_text)
                else:
                    widget.setCurrentIndex(0)
        finally:
            for blocker in reversed(blockers):
                try:
                    blocker.__exit__(None, None, None)
                except Exception:
                    pass

    def refresh_dynamic_combo_options(self, source_widget=None):
        if self._updating_option_sources:
            return
        self._updating_option_sources = True
        try:
            input_vals = self.collect_input_values()
            if getattr(self, 'option_rows', None):
                row_iter = [(row.option_config, row.editor) for row in self.option_rows]
            else:
                row_iter = list(zip(sorted(self.current_options_config, key=lambda x: x.get('order', 0)), self.input_widgets))
            for opt, widget in row_iter:
                if opt.get('type') not in ('combo', 'editable_combo') or not isinstance(widget, QComboBox):
                    continue
                options = self.data_matcher.get_field_options(opt.get('source', {}), input_values=input_vals)
                normalized_options = [''] + (options if options else ['（无选项）'])
                current_items = [widget.itemText(i) for i in range(widget.count())]
                current_text = widget.currentText()

                if widget.isEditable():
                    # 用户正在当前可输入下拉框里打字时，不刷新这个控件自己的候选项。
                    # 否则 clear/addItems/setEditText 会让 QLineEdit 光标跳到末尾，造成“a1bc11111”这类错位输入。
                    if self._editable_combo_is_user_typing(widget, source_widget):
                        continue
                    # 可输入下拉框允许任意文本，不需要把用户当前文本临时追加到候选项中。
                    # 这样候选列表不会随每个输入字符变化，也不会反复触发重排和光标复位。
                    if current_items == normalized_options:
                        continue
                    self._update_combo_items_preserve_text(widget, normalized_options, current_text)
                    continue

                # 普通不可输入下拉框无法显示列表外文本；为避免动态候选项变化时清空已有选择，临时保留当前值。
                if current_text and current_text not in normalized_options:
                    normalized_options.append(current_text)
                if current_items == normalized_options:
                    continue
                self._update_combo_items_preserve_text(widget, normalized_options, current_text)
        finally:
            self._updating_option_sources = False

    def _set_result_text_preserve_view(self, text: str):
        vbar = self.result_text.verticalScrollBar()
        hbar = self.result_text.horizontalScrollBar()
        old_v_max = max(1, vbar.maximum())
        old_h_max = max(1, hbar.maximum())
        old_v_value = vbar.value()
        old_h_value = hbar.value()
        v_ratio = old_v_value / old_v_max if old_v_max else 0
        h_ratio = old_h_value / old_h_max if old_h_max else 0
        self.result_text.setPlainText(text)
        new_v_max = vbar.maximum()
        new_h_max = hbar.maximum()
        vbar.setValue(int(round(v_ratio * new_v_max)) if new_v_max else 0)
        hbar.setValue(int(round(h_ratio * new_h_max)) if new_h_max else 0)

    def _get_all_process_field_names(self):
        if not self.current_template_name:
            return []
        content = self._get_current_process_content() or {}
        names = []
        names.extend(content.get('available_field_names', []) or [])
        names.extend((content.get('available_fields', {}) or {}).keys())
        names.extend(content.get('selected_fields', []) or [])
        result = []
        seen = set()
        for name in names:
            name = str(name).strip()
            if name and name not in seen:
                seen.add(name)
                result.append(name)
        return result

    def _get_current_process_content(self):
        content = self._get_active_process_content_override()
        if isinstance(content, dict):
            return content
        if not self.current_template_name:
            return {}
        proc_tpl = self.template_db.get_process_template(self.current_template_name) or {}
        return (proc_tpl.get('content', {}) or {})

    def _get_visible_copy_button_fields(self):
        content = self._get_current_process_content() or {}
        field_conditions = content.get('field_conditions', {}) or {}
        field_configs = content.get('available_fields', {}) or {}
        visible = set()
        input_values = self.collect_input_values() if self.current_template_name else {}
        data_pool = dict(self._last_data_pool or {})
        for field in self._get_all_process_field_names():
            expr = str(field_conditions.get(field, '') or '').strip()
            if not expr:
                visible.add(field)
                continue
            try:
                if self.data_matcher._evaluate_condition(expr, data_pool, input_values, field_configs):
                    visible.add(field)
            except Exception:
                visible.add(field)
        return visible

    def _sanitize_copy_button_fields(self, fields):
        available = set(self._get_all_process_field_names())
        result = []
        seen = set()
        for field in fields or []:
            name = str(field or '').strip()
            if not name or name in seen:
                continue
            if available and name not in available:
                continue
            seen.add(name)
            result.append(name)
        return result

    def _detach_copy_items_for_relayout(self):
        """重排行容器前先把复制项安全临时移回滚动面板，避免变成顶层小窗口。"""
        parent = getattr(self, 'copy_scroll_widget', None)
        for item_widget, _ in list(getattr(self, 'copy_buttons', []) or []):
            try:
                item_widget.hide()
                if parent is not None:
                    item_widget.setParent(parent)
            except RuntimeError:
                continue
            except Exception:
                continue

    def _clear_copy_row_widgets(self, detach_items=True):
        """清理自动换行产生的行容器。detach_items=True 时保留 CopyButtonItem 本体。"""
        if not hasattr(self, 'copy_buttons_layout'):
            return
        if detach_items:
            self._detach_copy_items_for_relayout()
        while self.copy_buttons_layout.count():
            layout_item = self.copy_buttons_layout.takeAt(0)
            row_widget = layout_item.widget() if layout_item is not None else None
            if row_widget is not None:
                try:
                    row_widget.hide()
                    row_widget.deleteLater()
                except Exception:
                    pass
        self._copy_row_widgets = []

    def _clear_copy_button_widgets(self):
        old_widgets = [item_widget for item_widget, _ in self.copy_buttons]
        self._clear_copy_row_widgets(detach_items=True)
        for item_widget in old_widgets:
            internal_button = getattr(item_widget, 'button', item_widget)
            try:
                self.copy_button_group.removeButton(internal_button)
            except Exception:
                pass
            try:
                item_widget.hide()
                item_widget.deleteLater()
            except Exception:
                pass
        self.copy_buttons.clear()

    def _copy_item_width(self, item_widget):
        try:
            width = item_widget.sizeHint().width()
        except Exception:
            width = 0
        try:
            width = max(width, item_widget.minimumWidth())
        except Exception:
            pass
        try:
            width = max(width, item_widget.width())
        except Exception:
            pass
        return max(1, int(width or 1))

    def _make_copy_row_widget(self):
        row_widget = QWidget(self.copy_scroll_widget)
        row_widget.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(0, 0, 0, 0)
        try:
            row_layout.setSpacing(max(0, int(self.copy_buttons_layout.spacing())))
        except Exception:
            row_layout.setSpacing(6)
        row_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.copy_buttons_layout.addWidget(row_widget, 0, Qt.AlignLeft | Qt.AlignTop)
        self._copy_row_widgets.append(row_widget)
        return row_widget, row_layout

    def _relayout_copy_buttons(self):
        """按复制区当前宽度把按钮重排为多行；每行单独横向排列，不做列对齐。"""
        if not hasattr(self, 'copy_buttons_layout') or getattr(self, '_copy_relayouting', False):
            return
        self._copy_relayouting = True
        try:
            visible_entries = [(item_widget, field) for item_widget, field in self.copy_buttons if not getattr(item_widget, '_hidden_by_filter', False)]
            self._clear_copy_row_widgets(detach_items=True)
            try:
                viewport_width = int(self.copy_scroll.viewport().width())
            except Exception:
                viewport_width = 0
            left, top, right, bottom = self.copy_buttons_layout.getContentsMargins()
            spacing = max(0, int(self.copy_buttons_layout.spacing()))
            available_width = max(80, viewport_width - left - right - 4) if viewport_width > 0 else 600
            row_layout = None
            used_width = 0
            for item_widget, _ in visible_entries:
                item_width = self._copy_item_width(item_widget)
                next_width = item_width if row_layout is None or row_layout.count() == 0 else used_width + spacing + item_width
                if row_layout is None or (row_layout.count() > 0 and next_width > available_width):
                    _, row_layout = self._make_copy_row_widget()
                    used_width = 0
                row_layout.addWidget(item_widget, 0, Qt.AlignLeft | Qt.AlignTop)
                item_widget.show()
                used_width = item_width if used_width == 0 else used_width + spacing + item_width
            try:
                self.copy_scroll_widget.setMinimumWidth(0)
                self.copy_scroll_widget.setMaximumWidth(16777215)
                self.copy_buttons_layout.activate()
                content_height = self.copy_buttons_layout.sizeHint().height() + top + bottom
                self.copy_scroll_widget.setMinimumHeight(max(1, int(content_height)))
                self.copy_scroll_widget.updateGeometry()
            except Exception:
                pass
        finally:
            self._copy_relayouting = False

    def _normalize_copy_selected_fields(self, fields=None):
        ordered = []
        selected = set(self._sanitize_copy_button_fields(fields if fields is not None else self._copy_selected_fields))
        for _, field in self.copy_buttons:
            if field in selected and field not in ordered:
                ordered.append(field)
        self._copy_selected_fields = ordered
        if self._copy_last_selected_field and self._copy_last_selected_field not in {field for _, field in self.copy_buttons}:
            self._copy_last_selected_field = ''
        return ordered

    def _checked_copy_field_names(self):
        result = []
        for btn, field in self.copy_buttons:
            if btn.isChecked() and field not in result:
                result.append(field)
        return result

    def _selected_copy_field_names_for_action(self):
        selected = list(self._normalize_copy_selected_fields())
        if not selected:
            selected = self._sanitize_copy_button_fields(self._checked_copy_field_names())
            if selected:
                self._copy_selected_fields = list(selected)
        if not selected and (not getattr(self, 'copy_multi_check', None) or not self.copy_multi_check.isChecked()):
            fallback = self._sanitize_copy_button_fields([self._copy_last_selected_field])
            if fallback:
                selected = fallback
                self._copy_selected_fields = list(fallback)
        return list(selected)

    def _apply_copy_button_checked_state(self, fields=None):
        selected = set(self._normalize_copy_selected_fields(fields))
        for item_widget, field in self.copy_buttons:
            with signal_blocked(item_widget):
                item_widget.setChecked(field in selected)

    def select_copy_item(self, item_widget, toggle=True):
        if item_widget is None:
            return
        field_name = getattr(item_widget, 'field_name', '')
        if not field_name:
            return
        currently_selected = field_name in self._selected_copy_field_names_for_action()
        multi_enabled = bool(getattr(self, 'copy_multi_check', None) and self.copy_multi_check.isChecked())
        if multi_enabled:
            selected = self._selected_copy_field_names_for_action()
            if currently_selected and toggle:
                selected = [field for field in selected if field != field_name]
            elif field_name not in selected:
                selected.append(field_name)
            self._copy_selected_fields = selected
        else:
            self._copy_selected_fields = [] if currently_selected and toggle else [field_name]
        self._copy_last_selected_field = self._copy_selected_fields[-1] if self._copy_selected_fields else ''
        self._apply_copy_button_checked_state(self._copy_selected_fields)

    def find_copy_item_at_pos(self, global_pos):
        widget = QApplication.widgetAt(global_pos)
        while widget is not None:
            if isinstance(widget, CopyButtonItem):
                return widget
            widget = widget.parentWidget()
        return None

    def move_copy_item_by_mouse(self, item_widget, global_pos):
        target_item = self.find_copy_item_at_pos(global_pos)
        if target_item is None or target_item is item_widget:
            return
        item_index = next((i for i, (widget, _) in enumerate(self.copy_buttons) if widget is item_widget), -1)
        target_index = next((i for i, (widget, _) in enumerate(self.copy_buttons) if widget is target_item), -1)
        if item_index < 0 or target_index < 0:
            return
        local_pos = target_item.mapFromGlobal(global_pos)
        new_index = target_index if local_pos.x() < target_item.width() / 2 else target_index + 1
        moving = self.copy_buttons.pop(item_index)
        if item_index < new_index:
            new_index -= 1
        new_index = max(0, min(new_index, len(self.copy_buttons)))
        if new_index == item_index:
            self.copy_buttons.insert(item_index, moving)
            return
        self.copy_buttons.insert(new_index, moving)
        self.select_copy_item(item_widget, toggle=False)
        self.sync_copy_order_from_items(save=False)

    def sync_copy_order_from_items(self, save=False):
        self._relayout_copy_buttons()
        if save and self.current_template_name:
            preserved = self._remember_input_values()
            self._save_current_template()
            self._restore_input_values(preserved)

    def _rebuild_copy_button_widgets(self, fields, checked_fields=None):
        checked_fields = self._sanitize_copy_button_fields(checked_fields or self._copy_selected_fields)
        self._clear_copy_button_widgets()
        for field_name in self._sanitize_copy_button_fields(fields):
            item_widget = CopyButtonItem(field_name, self)
            self.copy_button_group.addButton(item_widget.button)
            self.copy_buttons.append((item_widget, field_name))
        self._relayout_copy_buttons()
        self._apply_runtime_ui_settings()
        self._apply_copy_button_checked_state(checked_fields)

    def get_current_copy_button_fields(self):
        return self._sanitize_copy_button_fields([field for _, field in self.copy_buttons])

    def update_copy_buttons_from_config(self):
        previous_selected = list(self._sanitize_copy_button_fields(self._copy_selected_fields))
        previous_last = self._sanitize_copy_button_fields([self._copy_last_selected_field])
        self._clear_copy_button_widgets()
        self._copy_selected_fields = []
        self._copy_last_selected_field = previous_last[0] if previous_last else ''
        if not self.current_template_name:
            return
        main_tpl = self.template_db.get_main_template(self.current_template_name) or {}
        main_cfg = main_tpl.get('config', {}) or {}
        raw_fields = main_cfg.get('copy_buttons', []) or []
        clean_fields = self._sanitize_copy_button_fields(raw_fields)
        # 复制按钮与字段池隔离：这里只显示主模板中保存的 copy_buttons，
        # 不再因为字段仍存在于“可用字段”中而自动补回，避免删除后刷新又恢复。
        display_fields = clean_fields
        restored_selected = self._sanitize_copy_button_fields(previous_selected)
        self._rebuild_copy_button_widgets(display_fields, checked_fields=restored_selected)
        saved_fields = self._sanitize_copy_button_fields(display_fields)
        if self._copy_last_selected_field and self._copy_last_selected_field not in saved_fields:
            self._copy_last_selected_field = ''
        if raw_fields != saved_fields and self.current_template_name:
            clean_config = {
                'options': self.current_options_config,
                'rules': self.current_rules_config,
                'copy_buttons': saved_fields,
            }
            self.template_db.update_main_template(self.current_template_name, self.current_template_name, clean_config)
            self._sync_current_template_combo_item(clean_config)
        self.refresh_copy_button_visibility(self._last_final_fields)

    def on_copy_multi_mode_changed(self, checked):
        if not checked:
            fallback = self._sanitize_copy_button_fields([self._copy_last_selected_field])
            self._copy_selected_fields = fallback
            self._apply_copy_button_checked_state(fallback)

    def on_copy_button_clicked(self, button, field_name, checked=False):
        field_name = str(field_name or '').strip()
        if not field_name:
            return
        multi_enabled = bool(getattr(self, 'copy_multi_check', None) and self.copy_multi_check.isChecked())
        if not multi_enabled:
            if checked:
                self._copy_selected_fields = [field_name]
                self._copy_last_selected_field = field_name
            else:
                self._copy_selected_fields = []
                if self._copy_last_selected_field == field_name:
                    self._copy_last_selected_field = ''
            self._apply_copy_button_checked_state(self._copy_selected_fields)
            # 选中和取消选中都执行复制；取消选中只是改变高亮状态，不影响复制当前按钮内容。
            qt_safe_single_shot(0, lambda f=field_name: self.copy_field_content(f))
            return

        self._copy_last_selected_field = field_name
        selected = list(self._selected_copy_field_names_for_action())
        if checked:
            if field_name not in selected:
                selected.append(field_name)
        else:
            selected = [field for field in selected if field != field_name]
        self._copy_selected_fields = selected
        if selected:
            self._copy_last_selected_field = selected[-1]
        else:
            self._copy_last_selected_field = ''
        self._apply_copy_button_checked_state(self._copy_selected_fields)
        # 多选模式下取消某个按钮选中时，也复制该按钮对应字段内容。
        qt_safe_single_shot(0, lambda f=field_name: self.copy_field_content(f))

    def _selected_copy_field_names(self):
        return list(self._selected_copy_field_names_for_action())

    def move_selected_copy_buttons(self, delta):
        selected_fields = self._selected_copy_field_names_for_action()
        selected = set(selected_fields)
        if not selected:
            QMessageBox.information(self, '提示', '请先选中要移动的按钮。')
            return

        items = list(self.copy_buttons)
        indexes = [i for i, (_, field) in enumerate(items) if field in selected]
        if delta < 0:
            for index in indexes:
                if index > 0 and items[index - 1][1] not in selected:
                    items[index - 1], items[index] = items[index], items[index - 1]
        elif delta > 0:
            for index in reversed(indexes):
                if index < len(items) - 1 and items[index + 1][1] not in selected:
                    items[index + 1], items[index] = items[index], items[index + 1]
        else:
            return

        self.copy_buttons = items
        self._relayout_copy_buttons()

        self._copy_selected_fields = list(selected_fields)
        if selected_fields:
            self._copy_last_selected_field = selected_fields[-1]
        self._apply_copy_button_checked_state(selected_fields)
        preserved = self._remember_input_values()
        self.refresh_copy_button_visibility(self._last_final_fields)
        self._save_current_template()
        self._restore_input_values(preserved)

    def refresh_copy_button_visibility(self, final_fields=None):
        visible_fields = self._get_visible_copy_button_fields() if self.current_template_name else set()
        changed = False
        for item_widget, field in self.copy_buttons:
            new_hidden = True if not visible_fields else field not in visible_fields
            if getattr(item_widget, '_hidden_by_filter', False) != new_hidden:
                changed = True
            item_widget._hidden_by_filter = new_hidden
            if new_hidden:
                try:
                    item_widget.hide()
                except Exception:
                    pass
        if changed:
            self._relayout_copy_buttons()

    def show_add_copy_button_menu(self):
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择模板')
            return
        all_fields = self._get_all_process_field_names()
        if not all_fields:
            QMessageBox.warning(self, '提示', '请先编辑字段模板，定义字段')
            return
        menu = QMenu(self)
        for field in all_fields:
            action = QAction(field, self)
            action.triggered.connect(lambda checked, f=field: self.add_copy_button_by_field(f))
            menu.addAction(action)
        anchor = getattr(self, 'copy_tool_btn', None)
        if anchor is not None:
            pos = anchor.mapToGlobal(anchor.rect().bottomLeft())
        else:
            pos = self.mapToGlobal(self.rect().center())
        menu.exec_(pos)

    def add_copy_button_by_field(self, field_name):
        field_name = str(field_name or '').strip()
        if not field_name:
            return
        all_fields = set(self._get_all_process_field_names())
        if all_fields and field_name not in all_fields:
            QMessageBox.warning(self, '提示', f'字段“{field_name}”不存在，无法添加。')
            return
        for _, existing in self.copy_buttons:
            if existing == field_name:
                QMessageBox.information(self, '提示', f'按钮“{field_name}”已存在')
                return
        selected = self._selected_copy_field_names()
        self._rebuild_copy_button_widgets(self.get_current_copy_button_fields() + [field_name], checked_fields=selected)
        preserved = self._remember_input_values()
        self.refresh_copy_button_visibility(self._last_final_fields)
        self._save_current_template()
        self._restore_input_values(preserved)

    def delete_selected_copy_button(self):
        selected_fields = self._selected_copy_field_names_for_action()
        selected = set(selected_fields)
        if not selected:
            QMessageBox.information(self, '提示', '请先选中要删除的按钮（点击按钮使其高亮）')
            return
        remain = [field for field in self.get_current_copy_button_fields() if field not in selected]
        self._copy_selected_fields = [field for field in selected_fields if field not in selected]
        if self._copy_last_selected_field in selected:
            self._copy_last_selected_field = self._copy_selected_fields[-1] if self._copy_selected_fields else ''
        self._rebuild_copy_button_widgets(remain, checked_fields=self._copy_selected_fields)
        if getattr(self, 'copy_multi_check', None) and not self.copy_multi_check.isChecked():
            self.on_copy_multi_mode_changed(False)
        preserved = self._remember_input_values()
        self.refresh_copy_button_visibility(self._last_final_fields)
        self._save_current_template()
        self._restore_input_values(preserved)

    def copy_field_content(self, field_name):
        if not self.current_template_name:
            return
        try:
            input_values = self.collect_input_values()
            copy_text = ''
            # 普通点击复制按钮时不再强制刷新预览和重排复制区；输入未变化时直接用缓存。
            if field_name in (self._last_final_fields or {}) and input_values == (self._last_input_values or {}):
                copy_text = self._last_final_fields.get(field_name, '')
            else:
                content = dict(self._get_current_process_content() or {})
                content['selected_fields'] = [field_name]
                _, _, final_fields = self.data_matcher.render(
                    self.current_template_name,
                    input_values,
                    process_content_override=content,
                )
                copy_text = final_fields.get(field_name, '')
            QApplication.clipboard().setText(str(copy_text or ''))
        except Exception as e:
            QMessageBox.critical(self, '错误', f'复制失败：{e}')

    def collect_input_values(self):
        input_vals = {}
        if getattr(self, 'option_rows', None):
            row_iter = [(row.option_config, row.editor) for row in self.option_rows]
        else:
            row_iter = list(zip(sorted(self.current_options_config, key=lambda x: x.get('order', 0)), self.input_widgets))
        for opt, widget in row_iter:
            label = opt.get('label', '')
            if not label:
                continue
            if isinstance(widget, QTextEdit):
                input_vals[label] = widget.toPlainText()
            elif isinstance(widget, QLineEdit):
                input_vals[label] = widget.text()
            elif isinstance(widget, QComboBox):
                input_vals[label] = widget.currentText()
            elif isinstance(widget, QCheckBox):
                input_vals[label] = str(widget.isChecked())
        return input_vals

    def update_result_text(self, force=False):
        if not self.current_template_name:
            self.result_text.clear()
            self._last_export_signature = None
            self._last_render_result_text = ''
            self._last_final_fields = {}
            self._last_input_values = {}
            self._last_data_pool = {}
            self.refresh_copy_button_visibility({})
            return
        input_vals = self.collect_input_values()
        if input_vals:
            self._input_value_cache.update(input_vals)
        try:
            process_override = self._get_active_process_content_override()
            result, data_pool, final_fields = self.data_matcher.render(self.current_template_name, input_vals, process_content_override=process_override)
            signature = self._build_render_signature(self.current_template_name, input_vals, result, final_fields)
            if force or signature != self._last_export_signature:
                self._set_result_text_preserve_view(result)
                export.export_data(final_fields)
                self._last_export_signature = signature
            self._last_render_result_text = result
            self._last_final_fields = final_fields
            self._last_input_values = input_vals
            self._last_data_pool = data_pool
            self.refresh_copy_button_visibility(final_fields)
        except Exception as e:
            error_text = f'生成失败：{e}'
            if force or self.result_text.toPlainText() != error_text:
                self._set_result_text_preserve_view(error_text)
                self._last_export_signature = None
            self._last_render_result_text = ''
            self._last_final_fields = {}
            self._last_input_values = input_vals
            self._last_data_pool = {}
            self.refresh_copy_button_visibility({})

    def save_to_summary_file(self):
        text = self.result_text.toPlainText()
        if not text.strip():
            return
        path, _ = QFileDialog.getSaveFileName(self, '保存', '汇总.txt', '文本 (*.txt)')
        if path:
            with open(path, 'a', encoding='utf-8') as f:
                f.write(text + '\n' + '=' * 50 + '\n\n')

    def export_current_to_browser(self):
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择模板。')
            return
        self.update_result_text(force=True)
        flow_override = None
        if self.browser_flow_window is not None:
            try:
                flow_override = self.browser_flow_window.collect_flow()
            except Exception as e:
                QMessageBox.warning(self, '提示', f'读取当前浏览器流程配置失败：{e}')
                return
        success, message = export.export_to_browser(
            self.template_db,
            self.current_template_name,
            self._last_render_result_text,
            self._last_final_fields,
            self._last_input_values,
            data_pool=self._last_data_pool,
            flow_override=flow_override,
            alert_handler=self._show_browser_alert,
        )
        if success:
            QMessageBox.information(self, '成功', '已执行浏览器导出。\n\n' + (message[-800:] if message else ''))
        else:
            QMessageBox.warning(self, '提示', message or '浏览器导出失败。')

    def clear_inputs(self):
        for widget in self.input_widgets:
            if isinstance(widget, QTextEdit):
                widget.clear()
            elif isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
            elif isinstance(widget, QCheckBox):
                widget.setChecked(False)
        self._input_value_cache = self.collect_input_values()
        self.update_result_text(force=True)

    def open_data_manager(self):
        if self.data_manager_window is None:
            self.data_manager_window = DataManagerWindow()
            self.data_manager_window.setAttribute(Qt.WA_DeleteOnClose, True)
            self.data_manager_window.data_changed.connect(self.on_external_data_changed)
            self.data_manager_window.destroyed.connect(self.on_data_manager_closed)
        self._apply_generic_window_ui_settings(self.data_manager_window)
        self.data_manager_window.show()
        self.data_manager_window.raise_()
        self.data_manager_window.activateWindow()

    def open_template_editor(self):
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择或新建模板')
            return
        if self.template_editor_window is None:
            self.template_editor_window = TemplateEditorWindow(self.db, self.current_template_name, self.template_db)
            self.template_editor_window.setAttribute(Qt.WA_DeleteOnClose, True)
            self.template_editor_window.content_changed.connect(self.on_live_process_template_changed)
            self.template_editor_window.destroyed.connect(self.on_template_editor_closed)
        else:
            self.template_editor_window.template_name = self.current_template_name
            self.template_editor_window.setWindowTitle(f'模板编辑 - {self.current_template_name}')
            self.template_editor_window.load_template_content()
        self._apply_generic_window_ui_settings(self.template_editor_window)
        self.template_editor_window.show()
        self.template_editor_window.raise_()
        self.template_editor_window.activateWindow()

    def open_browser_flow_editor(self):
        self._remember_input_values()
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择或新建模板')
            return
        if self.browser_flow_window is None:
            self.browser_flow_window = BrowserFlowWindow(self.template_db, self.current_template_name, export.get_engine(), self)
            self.browser_flow_window.setAttribute(Qt.WA_DeleteOnClose, True)
            self.browser_flow_window.destroyed.connect(self.on_browser_flow_window_closed)
        else:
            self.browser_flow_window.set_template_name(self.current_template_name)
        self._apply_generic_window_ui_settings(self.browser_flow_window)
        self.browser_flow_window.show()
        self.browser_flow_window.raise_()
        self.browser_flow_window.activateWindow()

    def on_browser_flow_window_closed(self):
        self._remember_input_values()
        self.browser_flow_window = None
        self._load_browser_settings_to_main()

    def on_template_editor_closed(self):
        self._remember_input_values()
        self.template_editor_window = None
        self._live_process_content = None
        try:
            self.template_db.ensure_connection()
            self.update_copy_buttons_from_config()
            self.update_result_text(force=True)
        except Exception as e:
            print(f'模板编辑器关闭后刷新失败：{e}')


    def _try_close_child_window(self, window):
        if window is None:
            return True
        try:
            closed = window.close()
        except Exception:
            return False
        if closed is False:
            return False
        try:
            return not window.isVisible()
        except Exception:
            return True

    def closeEvent(self, event):
        children_ok = True
        for window in (self.browser_flow_window, self.template_editor_window, self.data_manager_window):
            if not self._try_close_child_window(window):
                children_ok = False
                break
        if not children_ok:
            event.ignore()
            return
        self._live_refresh_timer.stop()
        self.db.close()
        self.template_db.close()
        event.accept()


def write_startup_error(error_text):
    try:
        base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base_dir, 'startup_error.log')
        with open(path, 'w', encoding='utf-8') as f:
            f.write(error_text)
    except Exception:
        pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    try:
        window = PEditor()
        window.show()
        sys.exit(app.exec_())
    except Exception:
        error_text = traceback.format_exc()
        write_startup_error(error_text)
        try:
            QMessageBox.critical(None, 'PEditor 启动失败', '程序启动失败，错误已写入 startup_error.log。\n\n' + error_text[-1200:])
        except Exception:
            print(error_text)
        sys.exit(1)
