import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QComboBox, QPushButton, QTextEdit,
    QMessageBox, QFileDialog, QLabel, QInputDialog, QGridLayout,
    QCheckBox, QDialog, QListWidget, QListWidgetItem, QDialogButtonBox,
    QGroupBox, QAbstractItemView, QScrollArea, QButtonGroup, QMenu,
    QAction, QSplitter, QFormLayout, QSpinBox
)
from PyQt5.QtCore import Qt, QTimer
from dbutils import Database
from datamanager import DataManagerWindow
from template_editor import TemplateEditorWindow
from template_db import TemplateDB
from datamatch import DataMatcher, RuleManagerDialog
import export
from webcontrol import BrowserFlowWindow
from utils import resource_path

__version__ = '2.6'
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
    def __init__(self, current_settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle('界面设置')
        self.resize(360, 180)
        self._settings = dict(current_settings or {})
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 24)
        self.font_size_spin.setValue(int(self._settings.get('font_size', 10) or 10))
        form.addRow('字体大小:', self.font_size_spin)

        self.compact_combo = QComboBox()
        self.compact_combo.addItems(['很紧凑', '紧凑', '标准', '宽松'])
        compactness = str(self._settings.get('button_compactness', '紧凑') or '紧凑')
        idx = self.compact_combo.findText(compactness)
        self.compact_combo.setCurrentIndex(idx if idx >= 0 else 1)
        form.addRow('按钮紧凑度:', self.compact_combo)

        layout.addLayout(form)
        hint = QLabel('按钮内边距会跟随字体大小自动联动，字体越大按钮会同步变高。')
        hint.setWordWrap(True)
        layout.addWidget(hint)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_settings(self):
        return {
            'font_size': int(self.font_size_spin.value()),
            'button_compactness': self.compact_combo.currentText(),
        }


class PEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.template_db = TemplateDB()
        self.data_matcher = DataMatcher(self.db, self.template_db)
        self.current_template_name = None
        self.current_options_config = []
        self.current_rules_config = []
        self.input_widgets = []
        self.copy_buttons = []
        self.copy_button_group = QButtonGroup(self)
        self.copy_button_group.setExclusive(False)
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
            'font_size': 10,
            'button_compactness': '紧凑',
        }

    def _load_settings_payload(self):
        settings = self._default_ui_settings()
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f) or {}
                    if isinstance(loaded, dict):
                        settings.update(loaded)
        except Exception as e:
            print(f'加载设置失败：{e}')
        return settings

    def _build_app_stylesheet(self, font_size: int, button_compactness: str):
        try:
            font_size = max(8, min(24, int(font_size)))
        except Exception:
            font_size = 10
        compact_map = {
            '很紧凑': (4, 1, 6),
            '紧凑': (6, 2, 8),
            '标准': (8, 4, 10),
            '宽松': (12, 6, 12),
        }
        pad_h, pad_v, extra_h = compact_map.get(button_compactness, compact_map['紧凑'])
        min_height = max(font_size + extra_h + pad_v * 2, 22)
        combo_height = max(min_height, font_size + extra_h + 2)
        return (
            f"* {{ font-size: {font_size}px; }}\n"
            f"QPushButton {{\n    padding: {pad_v}px {pad_h}px;\n    min-height: {min_height}px;\n}}\n"
            f"QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{\n    min-height: {combo_height}px;\n}}\n"
            "QAbstractButton {\n    spacing: 2px;\n}\n"
            "QPlainTextEdit, QTextEdit, QListWidget, QTableWidget {\n    padding: 2px;\n}"
        )

    def apply_ui_settings(self, settings: dict):
        merged = self._default_ui_settings()
        if isinstance(settings, dict):
            merged.update(settings)
        self._ui_settings = merged
        stylesheet = self._build_app_stylesheet(merged.get('font_size', 10), merged.get('button_compactness', '紧凑'))
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(stylesheet)
        return merged

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
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        t_layout = QHBoxLayout()
        t_layout.addWidget(QLabel('模板:'))
        self.template_combo = QComboBox()
        self.template_combo.setMinimumWidth(88)
        self.template_combo.setMaximumWidth(96)
        self.template_combo.view().setMinimumWidth(220)
        self.template_combo.currentTextChanged.connect(self.on_template_changed)
        t_layout.addWidget(self.template_combo)

        # 在模板操作按钮组中增加“教程”按钮，以便用户查阅使用说明
        for text, handler in [
            ('新增', self.new_template),
            ('重命名', self.rename_template),
            ('删除', self.delete_template),
            ('导入', self.import_template),
            ('导出', self.export_template),
            ('输入项', self.edit_options),
            ('工序', self.open_template_editor),
            ('编辑规则', self.edit_rules),
            ('浏览器配置', self.open_browser_flow_editor),
            ('数据库', self.open_data_manager),
            ('设置', self.open_settings),
            ('教程', self.open_tutorial),
        ]:
            btn = QPushButton(text)
            btn.clicked.connect(handler)
            t_layout.addWidget(btn)
            if text == '新增':
                self.new_btn = btn
            elif text == '重命名':
                self.rename_btn = btn
            elif text == '删除':
                self.del_btn = btn
            elif text == '导入':
                self.import_btn = btn
            elif text == '导出':
                self.export_btn = btn
            elif text == '输入项':
                self.edit_opt_btn = btn
            elif text == '工序':
                self.edit_field_btn = btn
            elif text == '编辑规则':
                self.edit_rule_btn = btn
            elif text == '浏览器配置':
                self.browser_cfg_btn = btn
            elif text == '数据库':
                self.db_btn = btn
            elif text == '设置':
                self.settings_btn = btn
            elif text == '教程':
                self.tutorial_btn = btn
        t_layout.addStretch()
        main_layout.addLayout(t_layout)


        browser_bar = QHBoxLayout()
        browser_bar.addWidget(QLabel('Driver路径:'))
        self.browser_driver_edit = QLineEdit()
        self.browser_driver_edit.setPlaceholderText('chromedriver.exe 路径')
        self.browser_driver_edit.editingFinished.connect(self.on_main_browser_settings_changed)
        browser_bar.addWidget(self.browser_driver_edit, 2)
        browser_bar.addWidget(QLabel('浏览器路径:'))
        self.browser_binary_edit = QLineEdit()
        self.browser_binary_edit.setPlaceholderText('chrome.exe 路径，可留空')
        self.browser_binary_edit.editingFinished.connect(self.on_main_browser_settings_changed)
        browser_bar.addWidget(self.browser_binary_edit, 2)
        browser_bar.addWidget(QLabel('启动URL:'))
        self.browser_url_edit = QLineEdit()
        self.browser_url_edit.setPlaceholderText('点击打开浏览器时使用的 URL')
        self.browser_url_edit.editingFinished.connect(self.normalize_main_browser_url_input)
        self.browser_url_edit.editingFinished.connect(self.on_main_browser_settings_changed)
        browser_bar.addWidget(self.browser_url_edit, 3)
        self.open_browser_btn = QPushButton('打开浏览器')
        self.open_browser_btn.clicked.connect(self.open_browser_from_main)
        browser_bar.addWidget(self.open_browser_btn)
        main_layout.addLayout(browser_bar)

        self.input_group = QGroupBox('工艺参数输入')
        self.input_layout = QGridLayout(self.input_group)
        self.input_layout.setColumnStretch(1, 1)
        self.input_layout.setColumnStretch(3, 1)
        self.input_layout.setColumnStretch(5, 1)
        main_layout.addWidget(self.input_group)

        self.splitter = QSplitter(Qt.Vertical)
        main_layout.addWidget(self.splitter, 1)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setPlaceholderText('工艺段落将实时显示在此...')
        self.splitter.addWidget(self.result_text)

        copy_group = QGroupBox('复制工序内容')
        copy_main_layout = QVBoxLayout(copy_group)
        copy_toolbar = QHBoxLayout()
        self.add_copy_btn = QPushButton('添加工序')
        self.add_copy_btn.clicked.connect(self.show_add_copy_button_menu)
        copy_toolbar.addWidget(self.add_copy_btn)
        self.del_copy_btn = QPushButton('删除选中')
        self.del_copy_btn.clicked.connect(self.delete_selected_copy_button)
        copy_toolbar.addWidget(self.del_copy_btn)
        self.copy_multi_check = QCheckBox('多选')
        self.copy_multi_check.toggled.connect(self.on_copy_multi_mode_changed)
        copy_toolbar.addWidget(self.copy_multi_check)
        copy_toolbar.addStretch()
        copy_main_layout.addLayout(copy_toolbar)

        self.copy_scroll = QScrollArea()
        self.copy_scroll.setWidgetResizable(True)
        self.copy_scroll_widget = QWidget()
        self.copy_buttons_layout = QHBoxLayout(self.copy_scroll_widget)
        self.copy_buttons_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.copy_scroll.setWidget(self.copy_scroll_widget)
        copy_main_layout.addWidget(self.copy_scroll)

        self.input_group.setMinimumHeight(110)
        self.result_text.setMinimumHeight(120)
        copy_group.setMinimumHeight(120)
        self.splitter.addWidget(copy_group)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setSizes([400, 150])

        bottom = QHBoxLayout()
        self.save_btn = QPushButton('导出工艺文件内容为TXT')
        self.save_btn.clicked.connect(self.save_to_summary_file)
        bottom.addWidget(self.save_btn)
        self.browser_export_btn = QPushButton('导出至浏览器')
        self.browser_export_btn.clicked.connect(self.export_current_to_browser)
        bottom.addWidget(self.browser_export_btn)
        self.clear_btn = QPushButton('清空输入')
        self.clear_btn.clicked.connect(self.clear_inputs)
        bottom.addWidget(self.clear_btn)
        self.exit_btn = QPushButton('退出')
        self.exit_btn.clicked.connect(self.close)
        bottom.addWidget(self.exit_btn)
        main_layout.addLayout(bottom)


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
        payload = self._default_ui_settings()
        if isinstance(settings, dict):
            payload.update(settings)
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f'保存设置失败：{e}')

    def open_settings(self):
        current = dict(getattr(self, '_ui_settings', self._default_ui_settings()))
        dlg = UiSettingsDialog(current, self)
        if dlg.exec_() == QDialog.Accepted:
            settings = dlg.get_settings()
            self.apply_ui_settings(settings)
            self.save_settings(settings)

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
        config = {'options': self.current_options_config, 'rules': self.current_rules_config, 'copy_buttons': self.get_current_copy_button_fields()}
        self.template_db.update_main_template(self.current_template_name, self.current_template_name, config)
        self._sync_current_template_combo_item(config)

    def refresh_input_area(self, preserved_values=None):
        for w in self.input_widgets:
            w.deleteLater()
        self.input_widgets.clear()
        while self.input_layout.count():
            item = self.input_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        opts = sorted(self.current_options_config, key=lambda x: x.get('order', 0))
        row = col = 0
        initial_values = dict(preserved_values or {})
        for opt in opts:
            label = QLabel(opt['label'])
            self.input_layout.addWidget(label, row, col * 2)
            widget = self._create_widget(opt, initial_values)
            self.input_layout.addWidget(widget, row, col * 2 + 1)
            self.input_widgets.append(widget)
            if preserved_values and opt['label'] in preserved_values:
                self._set_widget_value(widget, preserved_values.get(opt['label']))
            col += 1
            if col >= 3:
                col = 0
                row += 1
        self.setup_input_change_tracking()
        self.refresh_dynamic_combo_options()

    def _create_widget(self, opt, initial_values=None):
        widget_type = opt['type']
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
            if isinstance(widget, QLineEdit):
                widget.textChanged.connect(self.on_input_widget_changed)
            elif isinstance(widget, QComboBox):
                widget.currentTextChanged.connect(self.on_input_widget_changed)
            elif isinstance(widget, QCheckBox):
                widget.stateChanged.connect(self.on_input_widget_changed)

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
            for opt, widget in zip(sorted(self.current_options_config, key=lambda x: x.get('order', 0)), self.input_widgets):
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

    def get_current_copy_button_fields(self):
        return [field for _, field in self.copy_buttons]

    def update_copy_buttons_from_config(self):
        for btn, _ in self.copy_buttons:
            self.copy_button_group.removeButton(btn)
            btn.deleteLater()
        self.copy_buttons.clear()
        while self.copy_buttons_layout.count():
            item = self.copy_buttons_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.current_template_name:
            return
        main_tpl = self.template_db.get_main_template(self.current_template_name)
        if main_tpl:
            saved_fields = main_tpl.get('config', {}).get('copy_buttons', [])
            seen = set()
            for field in saved_fields:
                if field not in seen:
                    seen.add(field)
                    self._add_copy_button_internal(field)
        else:
            proc_tpl = self.template_db.get_process_template(self.current_template_name)
            if proc_tpl:
                seen = set()
                for field in proc_tpl.get('content', {}).get('selected_fields', []):
                    if field not in seen:
                        seen.add(field)
                        self._add_copy_button_internal(field)
        self.refresh_copy_button_visibility(self._last_final_fields)

    def on_copy_multi_mode_changed(self, checked):
        if not checked:
            for btn, _ in self.copy_buttons:
                old = btn.blockSignals(True)
                btn.setChecked(False)
                btn.blockSignals(old)

    def on_copy_button_clicked(self, button, field_name, checked=False):
        if not getattr(self, 'copy_multi_check', None) or not self.copy_multi_check.isChecked():
            for other, _ in self.copy_buttons:
                if other is not button and other.isChecked():
                    old = other.blockSignals(True)
                    other.setChecked(False)
                    other.blockSignals(old)
        self.copy_field_content(field_name)
        if not getattr(self, 'copy_multi_check', None) or not self.copy_multi_check.isChecked():
            old = button.blockSignals(True)
            button.setChecked(False)
            button.blockSignals(old)

    def _add_copy_button_internal(self, field_name):
        btn = QPushButton(field_name)
        btn.setCheckable(True)
        btn.setMinimumWidth(80)
        btn.clicked.connect(lambda checked, b=btn, f=field_name: self.on_copy_button_clicked(b, f, checked))
        self.copy_buttons_layout.addWidget(btn)
        self.copy_button_group.addButton(btn)
        self.copy_buttons.append((btn, field_name))

    def refresh_copy_button_visibility(self, final_fields=None):
        visible_fields = set((final_fields or {}).keys()) if isinstance(final_fields, dict) else set()
        if not visible_fields and self.current_template_name:
            for btn, _ in self.copy_buttons:
                btn.setVisible(True)
            return
        for btn, field in self.copy_buttons:
            btn.setVisible(field in visible_fields)

    def show_add_copy_button_menu(self):
        if not self.current_template_name:
            QMessageBox.warning(self, '提示', '请先选择模板')
            return
        proc_tpl = self.template_db.get_process_template(self.current_template_name)
        if not proc_tpl:
            QMessageBox.warning(self, '提示', '请先编辑字段模板，定义已选字段')
            return
        selected_fields = proc_tpl.get('content', {}).get('selected_fields', [])
        if not selected_fields:
            QMessageBox.warning(self, '提示', '已选字段列表为空')
            return
        menu = QMenu(self)
        for field in list(dict.fromkeys(selected_fields)):
            action = QAction(field, self)
            action.triggered.connect(lambda checked, f=field: self.add_copy_button_by_field(f))
            menu.addAction(action)
        menu.exec_(self.add_copy_btn.mapToGlobal(self.add_copy_btn.rect().bottomLeft()))

    def add_copy_button_by_field(self, field_name):
        for _, existing in self.copy_buttons:
            if existing == field_name:
                QMessageBox.information(self, '提示', f'按钮“{field_name}”已存在')
                return
        self._add_copy_button_internal(field_name)
        self._save_current_template()

    def delete_selected_copy_button(self):
        to_remove = []
        for btn, field in self.copy_buttons:
            if btn.isChecked():
                self.copy_button_group.removeButton(btn)
                btn.deleteLater()
                to_remove.append((btn, field))
        if not to_remove:
            QMessageBox.information(self, '提示', '请先选中要删除的按钮（点击按钮使其高亮）')
            return
        for item in to_remove:
            self.copy_buttons.remove(item)
        if getattr(self, 'copy_multi_check', None) and not self.copy_multi_check.isChecked():
            self.on_copy_multi_mode_changed(False)
        self._save_current_template()

    def copy_field_content(self, field_name):
        if not self.current_template_name:
            return
        try:
            self.update_result_text(force=True)
            QApplication.clipboard().setText(str(self._last_final_fields.get(field_name, '')))
        except Exception as e:
            QMessageBox.critical(self, '错误', f'复制失败：{e}')

    def collect_input_values(self):
        input_vals = {}
        sorted_options = sorted(self.current_options_config, key=lambda x: x.get('order', 0))
        for opt, widget in zip(sorted_options, self.input_widgets):
            label = opt['label']
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
