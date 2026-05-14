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
    QToolButton, QSizePolicy
)
from PyQt5.QtCore import Qt, QTimer, QEvent
from PyQt5.QtGui import QFont
from dbutils import Database
from datamanager import DataManagerWindow
from template_editor import TemplateEditorWindow
from template_db import TemplateDB
from datamatch import DataMatcher, RuleManagerDialog
import export
from webcontrol import BrowserFlowWindow
from utils import resource_path
from log import LogViewerDialog, log_change

__version__ = '3.0'
# 打包命令：pyinstaller --clean PEditor.spec --distpath "D:\Microsoft Visual Studio\code"


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
        self.fixed_values_hint.setStyleSheet('color: #555; font-size: 13px;')
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


class UiSettingsDialog(QDialog):
    """界面设置窗口：分别设置选项、工具栏、复制按钮、预览框的字体和紧凑程度。"""

    def __init__(self, current_settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle('界面设置')
        self.resize(520, 330)
        self._settings = dict(current_settings or {})
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
        form = QFormLayout()

        self.option_font_spin = self._make_spin('option_font_size', 10, 8, 30)
        form.addRow('选项字体大小:', self.option_font_spin)
        self.option_compact_spin = self._make_spin('option_compactness', 4, 0, 30)
        form.addRow('选项紧凑程度:', self.option_compact_spin)

        self.toolbar_font_spin = self._make_spin('toolbar_font_size', 10, 8, 30)
        form.addRow('工具栏字体大小:', self.toolbar_font_spin)
        self.toolbar_compact_spin = self._make_spin('toolbar_compactness', 4, 0, 30)
        form.addRow('工具栏紧凑程度:', self.toolbar_compact_spin)

        self.copy_font_spin = self._make_spin('copy_font_size', 10, 8, 30)
        form.addRow('复制按钮字体大小:', self.copy_font_spin)
        self.copy_compact_spin = self._make_spin('copy_compactness', 4, 0, 30)
        form.addRow('复制按钮紧凑程度:', self.copy_compact_spin)

        self.preview_font_spin = self._make_spin('preview_font_size', 11, 8, 36)
        form.addRow('预览框字体大小:', self.preview_font_spin)
        self.preview_compact_spin = self._make_spin('preview_compactness', 4, 0, 30)
        form.addRow('预览框紧凑程度:', self.preview_compact_spin)

        layout.addLayout(form)
        hint = QLabel('紧凑程度为数字：数值越小越紧凑，数值越大控件内边距和行高越大。')
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_settings(self):
        return {
            'option_font_size': int(self.option_font_spin.value()),
            'option_compactness': int(self.option_compact_spin.value()),
            'toolbar_font_size': int(self.toolbar_font_spin.value()),
            'toolbar_compactness': int(self.toolbar_compact_spin.value()),
            'copy_font_size': int(self.copy_font_spin.value()),
            'copy_compactness': int(self.copy_compact_spin.value()),
            'preview_font_size': int(self.preview_font_spin.value()),
            'preview_compactness': int(self.preview_compact_spin.value()),
        }


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
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setMinimumHeight(34)
        self.option_config = option_config
        self.editor = editor_widget
        self.main_window = main_window
        self.drag_start_y = None
        self.mouse_dragging = False

        self.row_layout = QHBoxLayout(self)
        self.row_layout.setContentsMargins(3, 3, 3, 3)
        self.row_layout.setSpacing(4)

        self.handle = QLabel('☰')
        self.handle.setObjectName('inputOptionHandle')
        self.handle.setFixedWidth(24)
        self.handle.setAlignment(Qt.AlignCenter)
        self.handle.setCursor(Qt.SizeAllCursor)
        self.handle.mousePressEvent = self.handle_mouse_press
        self.handle.mouseMoveEvent = self.handle_mouse_move
        self.handle.mouseReleaseEvent = self.handle_mouse_release

        self.label = QLabel(self.get_wrap_text(self.get_name()))
        self.label.setObjectName('inputOptionLabel')
        self.label.setWordWrap(True)
        # ===============================================================
        # 选项标签长度/宽度相关修改位置
        # 后续如果要改标签宽度、最大宽度、换行效果，可优先修改这里和 update_label_max_width()
        # ===============================================================
        self.update_label_max_width()
        self.label.mousePressEvent = self.label_mouse_press

        self.editor.installEventFilter(self)
        self.row_layout.addWidget(self.handle)
        self.row_layout.addWidget(self.label)
        self.row_layout.addWidget(self.editor, 1)
        self.set_selected(False)

    def get_name(self):
        return str(self.option_config.get('label', '') or '')

    def set_name(self, name):
        self.option_config['label'] = str(name or '').strip()
        self.label.setText(self.get_wrap_text(self.get_name()))
        self.update_label_max_width()

    def get_wrap_text(self, text):
        return '\u200b'.join(str(text))

    # ===============================================================
    # 选项标签长度/宽度相关修改位置
    # 主要修改点：label_width = max(72, min(170, left_panel_width // 4))
    # 72 表示最小固定宽度，170 表示最大固定宽度，left_panel_width // 4 表示约占左侧面板四分之一
    # 使用 setFixedWidth 固定标签宽度，可避免点击选中变色时输入框宽度抖动
    # ===============================================================
    def update_label_max_width(self):
        """稳定输入行标签宽度，避免点击选中后因为样式刷新导致整行宽度抖动。"""
        left_panel_width = 0
        try:
            if hasattr(self.main_window, 'input_scroll') and self.main_window.input_scroll is not None:
                left_panel_width = self.main_window.input_scroll.viewport().width()
        except Exception:
            left_panel_width = 0
        if left_panel_width <= 0:
            left_panel_width = 420
        # 标签宽度跟随左侧面板整体宽度变化，但使用固定宽度而不是最大宽度；
        # 这样选中/取消选中只改变颜色，不会触发布局重新分配导致输入框宽度变化。
        label_width = max(72, min(170, left_panel_width // 4))
        self.label.setFixedWidth(label_width)

    def set_selected(self, selected):
        border = '#2563eb' if selected else '#d1d5db'
        background = '#dbeafe' if selected else '#ffffff'
        handle_background = '#bfdbfe' if selected else '#f3f4f6'
        self.setStyleSheet(
            f"QWidget#inputOptionRow {{ background-color: {background}; border: 1px solid {border}; border-radius: 3px; }}"
            f"QLabel#inputOptionHandle {{ background-color: {handle_background}; border-right: 1px solid #d1d5db; }}"
            "QLabel#inputOptionLabel { background-color: transparent; border: none; }"
        )

    def apply_ui_settings(self, settings):
        option_font_size = int(settings.get('option_font_size', 10))
        option_compact = int(settings.get('option_compactness', 4))
        font = QFont()
        font.setPointSize(option_font_size)
        for widget in (self, self.handle, self.label, self.editor):
            widget.setFont(font)
        margin = max(1, option_compact // 2)
        spacing = max(1, option_compact // 2)
        self.row_layout.setContentsMargins(margin, margin, margin, margin)
        self.row_layout.setSpacing(spacing)
        row_height = max(30, option_font_size + option_compact * 2 + 12)
        editor_height = max(22, option_font_size + option_compact * 2 + 10)
        self.setMinimumHeight(row_height)
        self.editor.setMinimumHeight(editor_height)
        self.update_label_max_width()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_label_max_width()

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
        if abs(event.globalY() - self.drag_start_y) >= 3:
            self.mouse_dragging = True
            self.main_window.move_option_row_by_mouse(self, event.globalPos())

    def handle_mouse_release(self, event):
        self.drag_start_y = None
        if self.mouse_dragging:
            self.main_window.sync_option_order_from_rows(save=True)
        self.mouse_dragging = False

    def eventFilter(self, watched, event):
        if watched == self.editor and event.type() == QEvent.MouseButtonPress:
            if event.button() == Qt.LeftButton:
                self.main_window.select_option_row(self, toggle=False)
        return super().eventFilter(watched, event)


class TemplateSelectRow(QWidget):
    """固定在选项面板顶部的模板选择行，不参与输入值采集和拖拽排序。"""

    def __init__(self, template_combo, main_window, parent=None):
        super().__init__(parent)
        self.setObjectName('templateSelectRow')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setMinimumHeight(34)
        self.template_combo = template_combo
        self.main_window = main_window

        layout = QHBoxLayout(self)
        layout.setContentsMargins(3, 3, 3, 3)
        layout.setSpacing(4)

        # 固定首行不再显示“固定”字样，仅保留一个空白占位宽度，使模板行和普通选项行左侧对齐。
        fixed_mark = QLabel('')
        fixed_mark.setObjectName('templateFixedMark')
        fixed_mark.setFixedWidth(24)
        fixed_mark.setAlignment(Qt.AlignCenter)

        label = QLabel('模板')
        label.setObjectName('templateSelectLabel')
        label.setFixedWidth(60)
        label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)

        self.template_combo.setMinimumWidth(120)
        self.template_combo.view().setMinimumWidth(260)
        layout.addWidget(fixed_mark)
        layout.addWidget(label)
        layout.addWidget(self.template_combo, 1)
        self.setStyleSheet(
            'QWidget#templateSelectRow { background-color: #f9fafb; border: 1px solid #9ca3af; border-radius: 3px; }'
            'QLabel#templateFixedMark { background-color: transparent; border: none; }'
            'QLabel#templateSelectLabel { background-color: transparent; border: none; font-weight: bold; }'
        )

    def apply_ui_settings(self, settings):
        option_font_size = int(settings.get('option_font_size', 10))
        option_compact = int(settings.get('option_compactness', 4))
        font = QFont()
        font.setPointSize(option_font_size)
        self.setFont(font)
        for child in self.findChildren(QWidget):
            child.setFont(font)
        margin = max(1, option_compact // 2)
        spacing = max(1, option_compact // 2)
        layout = self.layout()
        if layout is not None:
            layout.setContentsMargins(margin, margin, margin, margin)
            layout.setSpacing(spacing)
        row_height = max(30, option_font_size + option_compact * 2 + 12)
        combo_height = max(22, option_font_size + option_compact * 2 + 10)
        self.setMinimumHeight(row_height)
        self.template_combo.setMinimumHeight(combo_height)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.main_window.clear_selected_option_row()
        super().mousePressEvent(event)


class CopyButtonItem(QWidget):
    """下方复制按钮区的 GTE 风格横向滚动项目：上方拖动柄，下方按钮。"""

    def __init__(self, field_name: str, main_window, parent=None):
        super().__init__(parent)
        self.setObjectName('copyButtonItem')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.field_name = str(field_name or '').strip()
        self.main_window = main_window
        self.drag_start_x = None
        self.mouse_dragging = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(3)

        self.handle = QLabel('☰')
        self.handle.setObjectName('copyButtonHandle')
        self.handle.setAlignment(Qt.AlignCenter)
        self.handle.setFixedHeight(22)
        self.handle.setCursor(Qt.SizeAllCursor)
        self.handle.mousePressEvent = self.handle_mouse_press
        self.handle.mouseMoveEvent = self.handle_mouse_move
        self.handle.mouseReleaseEvent = self.handle_mouse_release

        self.button = QPushButton(self.field_name)
        self.button.setCheckable(True)
        self.button.setMinimumWidth(92)
        self.button.setMaximumWidth(180)
        self.button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.button.clicked.connect(lambda checked, f=self.field_name: self.main_window.on_copy_button_clicked(self, f, checked))

        layout.addWidget(self.handle)
        layout.addWidget(self.button)
        self.set_selected(False)

    def apply_ui_settings(self, settings):
        copy_font_size = int(settings.get('copy_font_size', 10))
        copy_compact = int(settings.get('copy_compactness', 4))
        font = QFont()
        font.setPointSize(copy_font_size)
        for widget in (self, self.handle, self.button):
            widget.setFont(font)
        margin = max(1, copy_compact // 2)
        spacing = max(1, copy_compact // 2)
        layout = self.layout()
        if layout is not None:
            layout.setContentsMargins(margin, margin, margin, margin)
            layout.setSpacing(spacing)
        self.handle.setFixedHeight(max(18, copy_font_size + copy_compact + 8))
        self.button.setMinimumHeight(max(22, copy_font_size + copy_compact * 2 + 8))
        self.adjustSize()

    def isChecked(self):
        return self.button.isChecked()

    def setChecked(self, checked):
        self.button.setChecked(bool(checked))
        self.set_selected(bool(checked))

    def blockSignals(self, block):
        return self.button.blockSignals(block)

    def set_selected(self, selected):
        border = '#2563eb' if selected else '#d1d5db'
        background = '#dbeafe' if selected else '#ffffff'
        handle_background = '#bfdbfe' if selected else '#f3f4f6'
        self.setStyleSheet(
            f"QWidget#copyButtonItem {{ background-color: {background}; border: 1px solid {border}; border-radius: 3px; }}"
            f"QLabel#copyButtonHandle {{ background-color: {handle_background}; border-bottom: 1px solid #d1d5db; }}"
        )

    def handle_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_x = event.globalX()
            self.mouse_dragging = False
            self.main_window.select_copy_item(self, toggle=False)

    def handle_mouse_move(self, event):
        if self.drag_start_x is None:
            return
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

    @staticmethod
    def _default_ui_settings():
        return {
            'option_font_size': 10,
            'option_compactness': 4,
            'toolbar_font_size': 10,
            'toolbar_compactness': 4,
            'copy_font_size': 10,
            'copy_compactness': 4,
            'preview_font_size': 11,
            'preview_compactness': 4,
        }

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
        for key in ('option_font_size', 'toolbar_font_size', 'copy_font_size'):
            payload[key] = self._safe_int(payload.get(key), defaults[key], 8, 30)
        payload['preview_font_size'] = self._safe_int(payload.get('preview_font_size'), defaults['preview_font_size'], 8, 36)
        for key in ('option_compactness', 'toolbar_compactness', 'copy_compactness', 'preview_compactness'):
            payload[key] = self._safe_int(payload.get(key), defaults[key], 0, 30)
        return payload

    def _build_app_stylesheet(self, settings: dict):
        settings = self._normalize_ui_settings(settings)
        option_font = settings['option_font_size']
        option_compact = settings['option_compactness']
        toolbar_font = settings['toolbar_font_size']
        toolbar_compact = settings['toolbar_compactness']
        copy_font = settings['copy_font_size']
        copy_compact = settings['copy_compactness']
        preview_font = settings['preview_font_size']
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
            f"QToolBar#mainToolbar, QToolBar#mainToolbar QToolButton, QMenu#toolbarMenu {{ font-size: {toolbar_font}px; }}\n"
            f"QToolBar#mainToolbar QToolButton {{ padding: {toolbar_padding_v}px {toolbar_padding_h}px; min-height: {toolbar_min_height}px; }}\n"
            f"QWidget#inputOptionRow QLabel, QWidget#templateSelectRow QLabel, "
            f"QWidget#inputOptionRow QLineEdit, QWidget#inputOptionRow QComboBox, QWidget#inputOptionRow QCheckBox, "
            f"QWidget#templateSelectRow QComboBox {{ font-size: {option_font}px; }}\n"
            f"QWidget#inputOptionRow QLineEdit, QWidget#inputOptionRow QComboBox, QWidget#templateSelectRow QComboBox {{ "
            f"padding: {option_padding_v}px {option_padding_h}px; min-height: {option_min_height}px; }}\n"
            f"QWidget#copyButtonItem QPushButton, QWidget#copyButtonItem QLabel {{ font-size: {copy_font}px; }}\n"
            f"QWidget#copyButtonItem QPushButton {{ padding: {copy_padding_v}px {copy_padding_h}px; min-height: {copy_min_height}px; }}\n"
            f"QTextEdit#previewTextEdit {{ font-size: {preview_font}px; padding: {preview_padding}px; border: none; background: white; }}\n"
            f"QTextEdit#previewTextEdit:focus {{ border: none; }}\n"
            f"QTextEdit#previewTextEdit QFrame {{ border: none; }}\n"
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

    def _apply_runtime_ui_settings(self):
        settings = self._normalize_ui_settings(getattr(self, '_ui_settings', self._default_ui_settings()))

        toolbar_font_size = settings.get('toolbar_font_size', 10)
        toolbar_height = max(22, toolbar_font_size + settings.get('toolbar_compactness', 4) * 2 + 8)
        if hasattr(self, 'main_toolbar'):
            self._safe_set_font(self.main_toolbar, toolbar_font_size)
            for button in list(getattr(self, 'toolbar_buttons', []) or []):
                try:
                    self._safe_set_font(button, toolbar_font_size)
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

        if hasattr(self, 'result_text'):
            try:
                preview_font_size = settings.get('preview_font_size', 11)
                preview_padding = settings.get('preview_compactness', 4)
                self._safe_set_font(self.result_text, preview_font_size)
                self.result_text.setFrameShape(QFrame.NoFrame)
                self.result_text.setLineWidth(0)
                self.result_text.setMidLineWidth(0)
                self.result_text.setStyleSheet(
                    'QTextEdit#previewTextEdit {'
                    f'font-size: {preview_font_size}px; '
                    f'padding: {preview_padding}px; '
                    'border: 0px; background: white;'
                    '} QTextEdit#previewTextEdit:focus { border: 0px; }'
                )
                viewport = self.result_text.viewport()
                if viewport is not None:
                    viewport.setStyleSheet('border: 0px; background: white;')
                    viewport.setAutoFillBackground(False)
            except RuntimeError:
                pass
            except Exception:
                print('应用预览框样式失败：')
                traceback.print_exc()

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
        if isinstance(widget, QLineEdit):
            widget.setText(text_value)
        elif isinstance(widget, QComboBox):
            index = widget.findText(text_value)
            if index >= 0:
                widget.setCurrentIndex(index)
            elif widget.isEditable():
                widget.setCurrentText(text_value)
            else:
                widget.setCurrentIndex(0)
        elif isinstance(widget, QCheckBox):
            widget.setChecked(str(value).lower() in ('true', '1', 'yes', 'checked'))

    def on_live_process_template_changed(self, template_name, content):
        if template_name != self.current_template_name:
            return
        self._live_process_content = content
        if self.browser_flow_window is not None and self.browser_flow_window.template_name == self.current_template_name:
            try:
                self.browser_flow_window.refresh_field_combo()
            except Exception:
                pass
        self.update_result_text(force=True)

    def on_external_data_changed(self):
        preserved = self.collect_input_values()
        self.refresh_input_area(preserved)
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
        tool_btn = QToolButton(self)
        tool_btn.setObjectName('toolbarMenuButton')
        tool_btn.setText(title)
        tool_btn.setPopupMode(QToolButton.InstantPopup)
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

        self.copy_tool_btn = QToolButton(self)
        self.copy_tool_btn.setObjectName('toolbarMenuButton')
        self.copy_tool_btn.setText('复制')
        self.copy_tool_btn.setPopupMode(QToolButton.InstantPopup)
        copy_menu = QMenu(self.copy_tool_btn)
        copy_menu.setObjectName('toolbarMenu')
        self._new_toolbar_action(copy_menu, '添加工序按钮', self.show_add_copy_button_menu, 'add_copy_btn')
        self._new_toolbar_action(copy_menu, '删除选中按钮', self.delete_selected_copy_button, 'del_copy_btn')
        self._new_toolbar_action(copy_menu, '前移选中按钮', lambda: self.move_selected_copy_buttons(-1), 'move_copy_left_btn')
        self._new_toolbar_action(copy_menu, '后移选中按钮', lambda: self.move_selected_copy_buttons(1), 'move_copy_right_btn')
        copy_menu.addSeparator()
        self.copy_multi_check = self._new_toolbar_action(copy_menu, '多选复制按钮', self.on_copy_multi_mode_changed, checkable=True)
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
        color = 'white' if white else 'transparent'
        panel.setStyleSheet(f'QFrame#{object_name} {{ background-color: {color}; border: 1px solid #999; }}')
        return panel

    def create_left_panel(self):
        panel = self._make_panel('gteLeftPanel')
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.input_scroll = QScrollArea()
        self.input_scroll.setWidgetResizable(True)
        self.input_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.input_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.input_scroll.setFrameShape(QFrame.NoFrame)

        self.input_scroll_panel = QFrame()
        self.input_scroll_panel.setFrameShape(QFrame.NoFrame)
        self.input_layout = QVBoxLayout(self.input_scroll_panel)
        self.input_layout.setContentsMargins(5, 5, 5, 5)
        self.input_layout.setSpacing(5)

        self.template_combo = QComboBox()
        self.template_combo.currentTextChanged.connect(self.on_template_changed)
        self.template_row = TemplateSelectRow(self.template_combo, self)
        self.input_layout.addWidget(self.template_row)
        self.input_layout.addStretch()

        self.input_scroll.setWidget(self.input_scroll_panel)
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
        self.result_text.setStyleSheet('border: none; background: white;')
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
        # 复制按钮区需要横向滚动；不让内容容器被强行压缩到视口宽度。
        self.copy_scroll.setWidgetResizable(False)
        self.copy_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.copy_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.copy_scroll.setFrameShape(QFrame.NoFrame)
        self.copy_scroll_widget = QFrame()
        self.copy_scroll_widget.setFrameShape(QFrame.NoFrame)
        self.copy_scroll_widget.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        self.copy_buttons_layout = QHBoxLayout(self.copy_scroll_widget)
        self.copy_buttons_layout.setContentsMargins(5, 5, 5, 5)
        self.copy_buttons_layout.setSpacing(6)
        self.copy_buttons_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.copy_scroll.setWidget(self.copy_scroll_widget)
        layout.addWidget(self.copy_scroll)
        return panel

    def open_browser_settings_dialog(self):
        browser = self._get_browser_settings_for_current_template() if self.current_template_name else BrowserFlowWindow._default_browser()
        dlg = BrowserSettingsDialog(browser, self)
        if dlg.exec_() != QDialog.Accepted:
            return
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
            old = widget.blockSignals(True)
            widget.setText(self._normalize_text(value))
            widget.blockSignals(old)

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
        dlg.exec_()


    def open_log_viewer(self):
        dlg = LogViewerDialog(self)
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
        previous = self.current_template_name
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        temps = self.template_db.get_main_templates()
        for t in temps:
            self.template_combo.addItem(t['name'], t)
        self.template_combo.blockSignals(False)

        self._live_process_content = None
        if temps:
            target_name = previous if previous and any(t['name'] == previous for t in temps) else temps[0]['name']
            self.template_combo.setCurrentText(target_name)
            self._load_template_by_name(target_name)
        else:
            self.current_template_name = None
            self.current_options_config = []
            self.current_rules_config = []
            self.result_text.clear()

        self.refresh_input_area()
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
        self._load_template_by_name(name)
        self.refresh_input_area()
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
        preserved = self.collect_input_values()
        dlg = OptionEditDialog(self.current_options_config, self.db, self)
        if dlg.exec_() == QDialog.Accepted:
            self.current_options_config = dlg.options_config
            self.refresh_input_area(preserved)
            self._save_current_template()
            if self.browser_flow_window is not None and self.browser_flow_window.template_name == self.current_template_name:
                try:
                    self.browser_flow_window.refresh_field_combo()
                except Exception:
                    pass
            self.update_result_text(force=True)

    def edit_rules(self):
        field_pool = self._build_rule_field_pool()
        dlg = RuleManagerDialog(self.current_rules_config, self.db, field_pool, self)
        if dlg.exec_() == QDialog.Accepted:
            self.current_rules_config = dlg.get_rules()
            self._save_current_template()
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

    def _create_widget(self, opt, initial_values=None):
        widget_type = opt.get('type', 'text')
        if widget_type == 'text':
            return QLineEdit()
        if widget_type in ('combo', 'editable_combo'):
            combo = QComboBox()
            if widget_type == 'editable_combo':
                combo.setEditable(True)
            combo.addItem('')
            options = self.data_matcher.get_field_options(opt.get('source', {}), input_values=initial_values or {})
            combo.addItems(options if options else ['（无选项）'])
            combo.setCurrentIndex(0)
            return combo
        if widget_type == 'checkbox':
            return QCheckBox()
        return QLineEdit()

    def setup_input_change_tracking(self):
        for widget in self.input_widgets:
            try:
                widget.disconnect()
            except Exception:
                pass
            if isinstance(widget, QLineEdit):
                widget.textChanged.connect(self.on_input_widget_changed)
            elif isinstance(widget, QComboBox):
                widget.currentTextChanged.connect(self.on_input_widget_changed)
            elif isinstance(widget, QCheckBox):
                widget.stateChanged.connect(self.on_input_widget_changed)

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
            self._save_current_template()
            self.update_result_text(force=True)


    def on_input_widget_changed(self, *args):
        if self._updating_option_sources:
            self.update_result_text()
            return
        self.refresh_dynamic_combo_options()
        self.update_result_text()

    def refresh_dynamic_combo_options(self):
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
                if current_items == normalized_options:
                    continue
                old = widget.blockSignals(True)
                widget.clear()
                widget.addItems(normalized_options)
                if current_text in normalized_options:
                    widget.setCurrentText(current_text)
                elif widget.isEditable() and current_text:
                    widget.setEditText(current_text)
                else:
                    widget.setCurrentIndex(0)
                widget.blockSignals(old)
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

    def _clear_copy_button_widgets(self):
        for item_widget, _ in self.copy_buttons:
            internal_button = getattr(item_widget, 'button', item_widget)
            try:
                self.copy_button_group.removeButton(internal_button)
            except Exception:
                pass
            item_widget.deleteLater()
        self.copy_buttons.clear()
        while self.copy_buttons_layout.count():
            item = self.copy_buttons_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.copy_scroll_widget.adjustSize()

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
            old = item_widget.blockSignals(True)
            item_widget.setChecked(field in selected)
            item_widget.blockSignals(old)

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
        item_index = self.copy_buttons_layout.indexOf(item_widget)
        target_index = self.copy_buttons_layout.indexOf(target_item)
        if item_index < 0 or target_index < 0:
            return
        local_pos = target_item.mapFromGlobal(global_pos)
        if local_pos.x() < target_item.width() / 2:
            new_index = target_index
        else:
            new_index = target_index + 1
        if item_index < new_index:
            new_index -= 1
        max_index = self.copy_buttons_layout.count() - 1
        new_index = max(0, min(new_index, max_index))
        if new_index == item_index:
            return
        self.copy_buttons_layout.removeWidget(item_widget)
        self.copy_buttons_layout.insertWidget(new_index, item_widget)
        self.select_copy_item(item_widget, toggle=False)
        self.sync_copy_order_from_items(save=False)

    def sync_copy_order_from_items(self, save=False):
        ordered = []
        known = {item_widget: field for item_widget, field in self.copy_buttons}
        for index in range(self.copy_buttons_layout.count()):
            widget = self.copy_buttons_layout.itemAt(index).widget()
            if isinstance(widget, CopyButtonItem) and widget in known:
                ordered.append((widget, known[widget]))
        self.copy_buttons = ordered
        self.copy_scroll_widget.adjustSize()
        if save and self.current_template_name:
            self._save_current_template()

    def _rebuild_copy_button_widgets(self, fields, checked_fields=None):
        checked_fields = self._sanitize_copy_button_fields(checked_fields or self._copy_selected_fields)
        self._clear_copy_button_widgets()
        for field_name in self._sanitize_copy_button_fields(fields):
            item_widget = CopyButtonItem(field_name, self)
            self.copy_buttons_layout.addWidget(item_widget)
            self.copy_button_group.addButton(item_widget.button)
            self.copy_buttons.append((item_widget, field_name))
        self.copy_scroll_widget.adjustSize()
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
            if checked:
                self.copy_field_content(field_name)
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
        if checked:
            self.copy_field_content(field_name)

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
        while self.copy_buttons_layout.count():
            item = self.copy_buttons_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
        for item_widget, _ in self.copy_buttons:
            self.copy_buttons_layout.addWidget(item_widget)
        self.copy_scroll_widget.adjustSize()

        self._copy_selected_fields = list(selected_fields)
        if selected_fields:
            self._copy_last_selected_field = selected_fields[-1]
        self._apply_copy_button_checked_state(selected_fields)
        self.refresh_copy_button_visibility(self._last_final_fields)
        self._save_current_template()

    def refresh_copy_button_visibility(self, final_fields=None):
        visible_fields = self._get_visible_copy_button_fields() if self.current_template_name else set()
        if not visible_fields:
            for item_widget, _ in self.copy_buttons:
                item_widget.setVisible(False)
            self.copy_scroll_widget.adjustSize()
            return
        for item_widget, field in self.copy_buttons:
            item_widget.setVisible(field in visible_fields)
        self.copy_scroll_widget.adjustSize()

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
        self.refresh_copy_button_visibility(self._last_final_fields)
        self._save_current_template()

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
        self.refresh_copy_button_visibility(self._last_final_fields)
        self._save_current_template()

    def copy_field_content(self, field_name):
        if not self.current_template_name:
            return
        try:
            self.update_result_text(force=True)
            copy_text = self._last_final_fields.get(field_name, '')
            if field_name not in self._last_final_fields:
                content = dict(self._get_current_process_content() or {})
                content['selected_fields'] = [field_name]
                _, _, final_fields = self.data_matcher.render(
                    self.current_template_name,
                    self.collect_input_values(),
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
            if isinstance(widget, QLineEdit):
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
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
            elif isinstance(widget, QCheckBox):
                widget.setChecked(False)
        self.update_result_text(force=True)

    def open_data_manager(self):
        if self.data_manager_window is None:
            self.data_manager_window = DataManagerWindow()
            self.data_manager_window.setAttribute(Qt.WA_DeleteOnClose, True)
            self.data_manager_window.data_changed.connect(self.on_external_data_changed)
            self.data_manager_window.destroyed.connect(self.on_data_manager_closed)
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
        self.template_editor_window.show()
        self.template_editor_window.raise_()
        self.template_editor_window.activateWindow()

    def open_browser_flow_editor(self):
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择或新建模板')
            return
        if self.browser_flow_window is None:
            self.browser_flow_window = BrowserFlowWindow(self.template_db, self.current_template_name, export.get_engine(), self)
            self.browser_flow_window.setAttribute(Qt.WA_DeleteOnClose, True)
            self.browser_flow_window.destroyed.connect(self.on_browser_flow_window_closed)
        else:
            self.browser_flow_window.set_template_name(self.current_template_name)
        self.browser_flow_window.show()
        self.browser_flow_window.raise_()
        self.browser_flow_window.activateWindow()

    def on_browser_flow_window_closed(self):
        self.browser_flow_window = None
        self._load_browser_settings_to_main()

    def on_template_editor_closed(self):
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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PEditor()
    window.show()
    sys.exit(app.exec_())
