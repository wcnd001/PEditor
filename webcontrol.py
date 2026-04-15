import json
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

ACTION_CLICK = '点击元素'
ACTION_INPUT = '输入文本'
ACTION_WAIT_ELEMENT = '等待元素'
ACTION_WAIT_NEW_WINDOW = '等待新窗口'
ACTION_SWITCH_WINDOW = '切换窗口'
ACTION_SWITCH_MAIN_WINDOW = '切回主窗口'
ACTION_SWITCH_IFRAME = '切换iframe'
ACTION_SWITCH_DEFAULT = '切回默认文档'
ACTION_SLEEP = '延时'
ACTION_DRAG = '拖拽元素'

STEP_ACTIONS = [
    ACTION_CLICK,
    ACTION_INPUT,
    ACTION_WAIT_ELEMENT,
    ACTION_WAIT_NEW_WINDOW,
    ACTION_SWITCH_WINDOW,
    ACTION_SWITCH_MAIN_WINDOW,
    ACTION_SWITCH_IFRAME,
    ACTION_SWITCH_DEFAULT,
    ACTION_SLEEP,
    ACTION_DRAG,
]

LOCATOR_TYPES = ['id', 'name', 'xpath', 'css selector', 'class name', 'tag name', 'link text', 'partial link text']
WINDOW_MATCH_TYPES = ['标题包含', 'URL包含', '序号']
NEWLINE_MODES = ['直接输入', '删除', '转为空格', '转为\\n']
TAB_MODES = ['直接输入', '删除', '转为4空格', '转为\\t']
SPACE_MODES = ['直接输入', '压缩为1个', '删除全部']


class BrowserFlowWindow(QMainWindow):
    def _init_window_geometry(self):
        screen = None
        try:
            from PyQt5.QtWidgets import QApplication
            app = QApplication.instance()
            screen = app.primaryScreen() if app else None
        except Exception:
            screen = None
        if screen:
            geo = screen.availableGeometry()
            width = min(1360, max(980, geo.width() - 60))
            height = min(900, max(700, geo.height() - 80))
            self.resize(width, height)
        else:
            self.resize(1280, 860)
        self.setMinimumSize(900, 620)

    def __init__(self, template_db, template_name, engine, parent=None):
        super().__init__(parent)
        self.template_db = template_db
        self.template_name = template_name
        self.engine = engine
        self.flow_config = {}
        self._loading_step = False
        self._loaded_signature = ''
        self.recorded_element = None
        self.record_timer = QTimer(self)
        self.record_timer.setInterval(250)
        self.record_timer.timeout.connect(self.poll_recorded_element)

        self.setWindowTitle(f'浏览器流程配置 - {self.template_name}')
        self._init_window_geometry()
        self.init_ui()
        self.load_flow()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        browser_group = QGroupBox('浏览器连接')
        browser_layout = QGridLayout(browser_group)

        self.chromedriver_edit = QLineEdit()
        self.chrome_binary_edit = QLineEdit()
        self.debug_port_edit = QLineEdit('9222')
        self.start_url_edit = QLineEdit()
        self.start_url_edit.editingFinished.connect(self.normalize_start_url_input)
        self.implicit_wait_edit = QLineEdit('2')

        browser_layout.addWidget(QLabel('chromedriver:'), 0, 0)
        browser_layout.addWidget(self.chromedriver_edit, 0, 1, 1, 3)
        browser_layout.addWidget(QLabel('Chrome路径:'), 1, 0)
        browser_layout.addWidget(self.chrome_binary_edit, 1, 1, 1, 3)
        browser_layout.addWidget(QLabel('启动网址:'), 2, 0)
        browser_layout.addWidget(self.start_url_edit, 2, 1, 1, 3)
        browser_layout.addWidget(QLabel('调试端口:'), 3, 0)
        browser_layout.addWidget(self.debug_port_edit, 3, 1)
        browser_layout.addWidget(QLabel('隐式等待(秒):'), 3, 2)
        browser_layout.addWidget(self.implicit_wait_edit, 3, 3)

        btn_row = QHBoxLayout()
        self.launch_btn = QPushButton('启动浏览器')
        self.launch_btn.clicked.connect(self.launch_browser)
        btn_row.addWidget(self.launch_btn)
        self.refresh_windows_btn = QPushButton('刷新窗口列表')
        self.refresh_windows_btn.clicked.connect(self.refresh_windows)
        btn_row.addWidget(self.refresh_windows_btn)
        self.inspect_btn = QPushButton('自动录制')
        self.inspect_btn.clicked.connect(self.start_element_recording)
        btn_row.addWidget(self.inspect_btn)
        self.test_btn = QPushButton('测试导入')
        self.test_btn.clicked.connect(self.test_import)
        btn_row.addWidget(self.test_btn)
        btn_row.addStretch()
        browser_layout.addLayout(btn_row, 4, 0, 1, 4)

        window_row = QHBoxLayout()
        self.window_combo = QComboBox()
        self.window_combo.setMinimumWidth(300)
        self.window_combo.setMaximumWidth(420)
        self.window_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.window_combo.currentIndexChanged.connect(self.on_window_selection_changed)
        window_row.addWidget(QLabel('已打开窗口:'))
        window_row.addWidget(self.window_combo)
        self.switch_window_btn = QPushButton('使用选中窗口')
        self.switch_window_btn.clicked.connect(self.switch_selected_window)
        window_row.addWidget(self.switch_window_btn)
        self.add_window_step_btn = QPushButton('将选中窗口生成步骤')
        self.add_window_step_btn.clicked.connect(self.add_switch_window_step_from_selection)
        window_row.addWidget(self.add_window_step_btn)
        window_row.addStretch()
        browser_layout.addLayout(window_row, 5, 0, 1, 4)

        main_layout.addWidget(browser_group)
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        main_layout.addWidget(splitter, 1)

        step_panel = QWidget()
        step_panel.setMinimumWidth(260)
        step_layout = QVBoxLayout(step_panel)
        step_layout.addWidget(QLabel('流程步骤'))
        self.step_list = QListWidget()
        self.step_list.currentRowChanged.connect(self.load_selected_step)
        step_layout.addWidget(self.step_list, 1)

        step_btns = QHBoxLayout()
        self.add_step_btn = QPushButton('新增步骤')
        self.add_step_btn.clicked.connect(self.add_step)
        step_btns.addWidget(self.add_step_btn)
        self.del_step_btn = QPushButton('删除步骤')
        self.del_step_btn.clicked.connect(self.delete_step)
        step_btns.addWidget(self.del_step_btn)
        self.up_step_btn = QPushButton('上移')
        self.up_step_btn.clicked.connect(self.move_step_up)
        step_btns.addWidget(self.up_step_btn)
        self.down_step_btn = QPushButton('下移')
        self.down_step_btn.clicked.connect(self.move_step_down)
        step_btns.addWidget(self.down_step_btn)
        step_layout.addLayout(step_btns)

        save_btns = QHBoxLayout()
        self.save_step_btn = QPushButton('应用当前步骤修改')
        self.save_step_btn.clicked.connect(self.apply_current_step_changes)
        save_btns.addWidget(self.save_step_btn)
        self.save_flow_btn = QPushButton('保存流程配置')
        self.save_flow_btn.clicked.connect(self.save_flow)
        save_btns.addWidget(self.save_flow_btn)
        step_layout.addLayout(save_btns)

        editor_panel = QWidget()
        editor_panel.setMinimumWidth(420)
        editor_layout = QVBoxLayout(editor_panel)

        editor_group = QGroupBox('步骤编辑')
        self.step_form = QFormLayout(editor_group)
        form = self.step_form
        self.step_name_edit = QLineEdit()
        self.step_condition_edit = QLineEdit()
        self.step_condition_edit.setPlaceholderText("留空=始终执行；例如 {是否二段硫化} == '是' 或 {二段硫化段落} != ''")
        self.action_combo = QComboBox()
        self.action_combo.addItems(STEP_ACTIONS)
        self.action_combo.currentTextChanged.connect(self.update_action_visibility)
        self.locator_type_combo = QComboBox()
        self.locator_type_combo.addItems(LOCATOR_TYPES)
        self.locator_value_edit = QLineEdit()
        self.target_locator_type_combo = QComboBox()
        self.target_locator_type_combo.addItems(LOCATOR_TYPES)
        self.target_locator_value_edit = QLineEdit()
        self.drop_position_combo = QComboBox()
        self.drop_position_combo.addItems(['中间', '上方', '下方', '自定义偏移'])
        self.drag_offset_x_spin = QSpinBox()
        self.drag_offset_x_spin.setRange(-9999, 9999)
        self.drag_offset_y_spin = QSpinBox()
        self.drag_offset_y_spin.setRange(-9999, 9999)
        self.drag_offset_widget = QWidget()
        drag_offset_layout = QHBoxLayout(self.drag_offset_widget)
        drag_offset_layout.setContentsMargins(0, 0, 0, 0)
        drag_offset_layout.addWidget(QLabel('X:'))
        drag_offset_layout.addWidget(self.drag_offset_x_spin)
        drag_offset_layout.addWidget(QLabel('Y:'))
        drag_offset_layout.addWidget(self.drag_offset_y_spin)
        drag_offset_layout.addStretch()
        self.value_template_edit = QPlainTextEdit()
        self.value_template_edit.setPlaceholderText('可使用 {字段名}、{__RESULT__}、{__NL__} 或 #if(...)#')
        self.wait_timeout_spin = QDoubleSpinBox()
        self.wait_timeout_spin.setRange(0.1, 9999)
        self.wait_timeout_spin.setValue(10)
        self.wait_timeout_spin.setDecimals(1)
        self.window_match_type_combo = QComboBox()
        self.window_match_type_combo.addItems(WINDOW_MATCH_TYPES)
        self.window_match_value_edit = QLineEdit()
        self.sleep_seconds_spin = QDoubleSpinBox()
        self.sleep_seconds_spin.setRange(0.1, 3600)
        self.sleep_seconds_spin.setValue(1)
        self.sleep_seconds_spin.setDecimals(1)
        self.use_js_click_check = QCheckBox('点击时使用 JS click')
        self.clear_before_input_check = QCheckBox('输入前清空原值')
        self.clear_before_input_check.setChecked(True)
        self.wait_clickable_check = QCheckBox('等待元素时要求可点击')
        self.note_edit = QPlainTextEdit()
        self.newline_mode_combo = QComboBox()
        self.newline_mode_combo.addItems(NEWLINE_MODES)
        self.tab_mode_combo = QComboBox()
        self.tab_mode_combo.addItems(TAB_MODES)
        self.space_mode_combo = QComboBox()
        self.space_mode_combo.addItems(SPACE_MODES)
        self.field_combo = QComboBox()
        self.insert_field_btn = QPushButton('插入字段')
        self.insert_field_btn.clicked.connect(self.insert_selected_field)
        self.insert_result_btn = QPushButton('插入{__RESULT__}')
        self.insert_result_btn.clicked.connect(lambda: self.insert_template_text('{__RESULT__}'))

        self.field_insert_widget = QWidget()
        field_row = QHBoxLayout(self.field_insert_widget)
        field_row.setContentsMargins(0, 0, 0, 0)
        field_row.addWidget(QLabel('已选字段去重:'))
        field_row.addWidget(self.field_combo, 1)
        field_row.addWidget(self.insert_field_btn)
        field_row.addWidget(self.insert_result_btn)

        form.addRow('步骤名称:', self.step_name_edit)
        form.addRow('执行条件:', self.step_condition_edit)
        form.addRow('动作类型:', self.action_combo)
        form.addRow('定位方式:', self.locator_type_combo)
        form.addRow('定位值:', self.locator_value_edit)
        form.addRow('目标定位方式:', self.target_locator_type_combo)
        form.addRow('目标定位值:', self.target_locator_value_edit)
        form.addRow('释放位置:', self.drop_position_combo)
        form.addRow('拖拽偏移:', self.drag_offset_widget)
        form.addRow('输入模板:', self.value_template_edit)
        form.addRow('字段插入:', self.field_insert_widget)
        form.addRow('回车/换行处理:', self.newline_mode_combo)
        form.addRow('Tab/缩进处理:', self.tab_mode_combo)
        form.addRow('空格处理:', self.space_mode_combo)
        form.addRow('等待超时(秒):', self.wait_timeout_spin)
        form.addRow('窗口匹配方式:', self.window_match_type_combo)
        form.addRow('窗口匹配值:', self.window_match_value_edit)
        form.addRow('延时秒数:', self.sleep_seconds_spin)
        form.addRow(self.use_js_click_check)
        form.addRow(self.clear_before_input_check)
        form.addRow(self.wait_clickable_check)
        form.addRow('备注:', self.note_edit)
        editor_layout.addWidget(editor_group)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText('浏览器连接、窗口切换、元素抓取日志会显示在这里。')
        editor_layout.addWidget(self.log_text, 1)

        elements_panel = QWidget()
        elements_panel.setMinimumWidth(360)
        elements_layout = QVBoxLayout(elements_panel)
        elements_group = QGroupBox('自动录制结果（只读）')
        eg_layout = QVBoxLayout(elements_group)
        self.record_text = QPlainTextEdit()
        self.record_text.setReadOnly(True)
        self.record_text.setPlaceholderText('点击“自动录制”后，切换到浏览器点击目标元素，XPath / 推荐定位信息会显示在这里。')
        eg_layout.addWidget(self.record_text, 1)
        element_btns = QHBoxLayout()
        self.apply_element_btn = QPushButton('将录制结果填入当前步骤')
        self.apply_element_btn.clicked.connect(self.apply_selected_element_to_step)
        element_btns.addWidget(self.apply_element_btn)
        self.add_click_step_btn = QPushButton('用录制结果新增点击步骤')
        self.add_click_step_btn.clicked.connect(lambda: self.add_step_from_element(default_action=ACTION_CLICK))
        element_btns.addWidget(self.add_click_step_btn)
        self.add_input_step_btn = QPushButton('用录制结果新增输入步骤')
        self.add_input_step_btn.clicked.connect(lambda: self.add_step_from_element(default_action=ACTION_INPUT))
        element_btns.addWidget(self.add_input_step_btn)
        eg_layout.addLayout(element_btns)
        elements_layout.addWidget(elements_group, 1)

        splitter.addWidget(step_panel)
        splitter.addWidget(editor_panel)
        splitter.addWidget(elements_panel)
        splitter.setSizes([300, 520, 420])
        self.update_action_visibility(self.action_combo.currentText())


    def _set_form_row_visible(self, field_widget, visible):
        label = self.step_form.labelForField(field_widget) if hasattr(self, 'step_form') else None
        if label is not None:
            label.setVisible(visible)
        field_widget.setVisible(visible)


    def log(self, message):
        self.log_text.appendPlainText(str(message))

    def set_template_name(self, template_name):
        self.template_name = template_name
        self.setWindowTitle(f'浏览器流程配置 - {self.template_name}')
        self.load_flow()

    @staticmethod
    def _default_browser():
        return {
            'connect_mode': 'launch',
            'chromedriver_path': '',
            'chrome_binary': '',
            'debug_address': '127.0.0.1:9222',
            'debug_port': 9222,
            'start_url': '',
            'implicit_wait': 2,
        }

    def _default_flow(self):
        return {'browser': self._default_browser(), 'steps': []}

    def _signature(self, flow):
        return json.dumps(flow or {}, ensure_ascii=False, sort_keys=True)

    @staticmethod
    def _truncate_window_text(title, url, max_len=78):
        title = title or '(无标题)'
        url = url or ''
        text = f'{title} | {url}' if url else title
        return text if len(text) <= max_len else text[: max_len - 1] + '…'

    def _dedup_selected_fields(self):
        proc = self.template_db.get_process_template(self.template_name) or {}
        content = proc.get('content', {}) or {}
        selected = content.get('selected_fields', []) or []
        unique = []
        seen = set()
        for field in selected:
            name = str(field).strip()
            if name and name not in seen:
                seen.add(name)
                unique.append(name)
        return unique

    def refresh_field_combo(self):
        names = []
        parent = self.parent()
        if parent is not None and hasattr(parent, '_build_rule_field_pool'):
            try:
                names.extend(parent._build_rule_field_pool())
            except Exception:
                pass
        names.extend(self._dedup_selected_fields())
        seen = set()
        ordered = []
        for name in names:
            name = str(name).strip()
            if name and name not in seen:
                seen.add(name)
                ordered.append(name)
        self.field_combo.blockSignals(True)
        self.field_combo.clear()
        for name in ordered:
            self.field_combo.addItem(name)
        self.field_combo.blockSignals(False)
    def normalize_start_url_input(self):
        normalized = self.engine.normalize_url(self.start_url_edit.text())
        if normalized != self.start_url_edit.text().strip():
            self.start_url_edit.setText(normalized)


    def _parse_int_value(self, text_value, default, field_name, strict=False):
        text_value = (text_value or '').strip()
        if not text_value:
            return default
        try:
            return int(float(text_value))
        except Exception:
            if strict:
                raise ValueError(f'{field_name}必须是数字')
            return default

    def _parse_float_value(self, text_value, default, field_name, strict=False):
        text_value = (text_value or '').strip()
        if not text_value:
            return default
        try:
            return float(text_value)
        except Exception:
            if strict:
                raise ValueError(f'{field_name}必须是数字')
            return default

    def apply_external_browser_settings(self, browser_settings: dict):
        browser_settings = browser_settings or {}
        self.apply_current_step_changes(silent=True)

        def _set_text(widget, value):
            old = widget.blockSignals(True)
            widget.setText('' if value is None else str(value))
            widget.blockSignals(old)

        browser = self.flow_config.setdefault('browser', self._default_browser())
        browser.update({
            'connect_mode': 'launch',
            'chromedriver_path': browser_settings.get('chromedriver_path', browser.get('chromedriver_path', '')),
            'chrome_binary': browser_settings.get('chrome_binary', browser.get('chrome_binary', '')),
            'debug_port': self._parse_int_value(str(browser_settings.get('debug_port', browser.get('debug_port', 9222))), 9222, '调试端口'),
            'start_url': self.engine.normalize_url(browser_settings.get('start_url', browser.get('start_url', ''))),
            'implicit_wait': self._parse_float_value(str(browser_settings.get('implicit_wait', browser.get('implicit_wait', 2))), 2.0, '隐式等待'),
        })

        _set_text(self.chromedriver_edit, browser.get('chromedriver_path', ''))
        _set_text(self.chrome_binary_edit, browser.get('chrome_binary', ''))
        _set_text(self.debug_port_edit, browser.get('debug_port', 9222))
        _set_text(self.start_url_edit, browser.get('start_url', ''))
        _set_text(self.implicit_wait_edit, browser.get('implicit_wait', 2))

        try:
            loaded = json.loads(self._loaded_signature) if self._loaded_signature else self._default_flow()
        except Exception:
            loaded = self._default_flow()
        loaded.setdefault('browser', self._default_browser())
        loaded['browser'].update(browser)
        self._loaded_signature = self._signature(loaded)

    def _collect_browser_settings(self, strict=False):
        return {
            'connect_mode': 'launch',
            'chromedriver_path': self.chromedriver_edit.text().strip(),
            'chrome_binary': self.chrome_binary_edit.text().strip(),
            'debug_port': self._parse_int_value(self.debug_port_edit.text(), 9222, '调试端口', strict=strict),
            'start_url': self.engine.normalize_url(self.start_url_edit.text()),
            'implicit_wait': self._parse_float_value(self.implicit_wait_edit.text(), 2.0, '隐式等待', strict=strict),
        }

    def on_window_selection_changed(self, index):
        if index < 0:
            return
        self.set_selected_window_as_target(auto=True)

    def set_selected_window_as_target(self, auto=False):
        item = self.window_combo.currentData()
        if not item:
            return
        try:
            self.engine.set_preferred_window(item['handle'])
            title = item.get('title') or item.get('url') or f"窗口 {item.get('index', '')}"
            prefix = '已自动选择目标窗口' if auto else '已选择目标窗口'
            self.log(f'{prefix}：{title}（不会拉起浏览器前台）')
        except Exception as e:
            self.log(f'选择目标窗口失败：{e}')
            if not auto:
                QMessageBox.critical(self, '错误', f'选择目标窗口失败：{e}')

    def load_flow(self):
        flow = self.template_db.get_browser_flow(self.template_name) or self._default_flow()
        flow.setdefault('browser', self._default_browser())
        flow.setdefault('steps', [])
        self.flow_config = flow
        browser = flow.get('browser', {}) or {}
        self.chromedriver_edit.setText(browser.get('chromedriver_path', ''))
        self.chrome_binary_edit.setText(browser.get('chrome_binary', ''))
        self.debug_port_edit.setText(str(browser.get('debug_port', 9222)))
        self.start_url_edit.setText(browser.get('start_url', ''))
        self.implicit_wait_edit.setText(str(browser.get('implicit_wait', 2)))
        self.refresh_field_combo()
        self.refresh_step_list()
        self.refresh_windows()
        self._loaded_signature = self._signature(self.collect_flow())

    def collect_flow(self, strict=False):
        self.apply_current_step_changes(silent=True)
        browser = self._default_browser()
        browser.update(self._collect_browser_settings(strict=strict))
        steps = [dict(step) for step in (self.flow_config.get('steps', []) or [])]
        return {'browser': browser, 'steps': steps}

    def has_unsaved_changes(self):
        self.apply_current_step_changes(silent=True)
        return self._signature(self.collect_flow()) != self._loaded_signature

    def save_flow(self):
        try:
            self.apply_current_step_changes(silent=True)
            self.flow_config = self.collect_flow(strict=True)
            self.template_db.update_browser_flow(self.template_name, self.flow_config)
            self._loaded_signature = self._signature(self.flow_config)
            self.log('浏览器流程配置已保存。')
            QMessageBox.information(self, '成功', '浏览器流程配置已保存。')
            return True
        except Exception as e:
            self.log(f'保存浏览器流程配置失败：{e}')
            QMessageBox.critical(self, '错误', f'保存浏览器流程配置失败：{e}')
            return False

    def refresh_step_list(self):
        self.step_list.clear()
        steps = self.flow_config.get('steps', []) or []
        for idx, step in enumerate(steps, start=1):
            title = f"{idx}. {step.get('name') or step.get('action') or '未命名步骤'}"
            item = QListWidgetItem(title)
            item.setData(Qt.UserRole, step)
            self.step_list.addItem(item)
        if self.step_list.count():
            self.step_list.setCurrentRow(0)
        else:
            self.clear_step_editor()

    def clear_step_editor(self):
        self._loading_step = True
        self.step_name_edit.clear()
        self.step_condition_edit.clear()
        self.action_combo.setCurrentIndex(0)
        self.locator_type_combo.setCurrentIndex(0)
        self.locator_value_edit.clear()
        self.target_locator_type_combo.setCurrentIndex(0)
        self.target_locator_value_edit.clear()
        self.drop_position_combo.setCurrentText('中间')
        self.drag_offset_x_spin.setValue(0)
        self.drag_offset_y_spin.setValue(0)
        self.value_template_edit.clear()
        self.newline_mode_combo.setCurrentText('直接输入')
        self.tab_mode_combo.setCurrentText('直接输入')
        self.space_mode_combo.setCurrentText('直接输入')
        self.wait_timeout_spin.setValue(10)
        self.window_match_type_combo.setCurrentIndex(0)
        self.window_match_value_edit.clear()
        self.sleep_seconds_spin.setValue(1)
        self.use_js_click_check.setChecked(False)
        self.clear_before_input_check.setChecked(True)
        self.wait_clickable_check.setChecked(False)
        self.note_edit.clear()
        self._loading_step = False
        self.update_action_visibility(self.action_combo.currentText())

    def add_step(self):
        step = {
            'name': f'步骤{len(self.flow_config.get("steps", [])) + 1}',
            'condition_expr': '',
            'action': ACTION_CLICK,
            'locator_type': 'xpath',
            'locator_value': '',
            'target_locator_type': 'xpath',
            'target_locator_value': '',
            'drop_position': '中间',
            'drag_offset_x': 0,
            'drag_offset_y': 0,
            'value_template': '',
            'newline_mode': '直接输入',
            'tab_mode': '直接输入',
            'space_mode': '直接输入',
            'wait_timeout': 10,
            'window_match_type': '标题包含',
            'window_match_value': '',
            'sleep_seconds': 1,
            'use_js_click': False,
            'clear_before_input': True,
            'wait_clickable': False,
            'note': '',
        }
        self.flow_config.setdefault('steps', []).append(step)
        self.refresh_step_list()
        self.step_list.setCurrentRow(self.step_list.count() - 1)

    def delete_step(self):
        row = self.step_list.currentRow()
        if row < 0:
            return
        del self.flow_config['steps'][row]
        self.refresh_step_list()

    def move_step_up(self):
        row = self.step_list.currentRow()
        if row > 0:
            steps = self.flow_config['steps']
            steps[row - 1], steps[row] = steps[row], steps[row - 1]
            self.refresh_step_list()
            self.step_list.setCurrentRow(row - 1)

    def move_step_down(self):
        row = self.step_list.currentRow()
        steps = self.flow_config.get('steps', [])
        if 0 <= row < len(steps) - 1:
            steps[row + 1], steps[row] = steps[row], steps[row + 1]
            self.refresh_step_list()
            self.step_list.setCurrentRow(row + 1)

    def load_selected_step(self, row):
        steps = self.flow_config.get('steps', []) or []
        if not (0 <= row < len(steps)):
            self.clear_step_editor()
            return
        step = steps[row]
        self._loading_step = True
        self.step_name_edit.setText(step.get('name', ''))
        self.step_condition_edit.setText(step.get('condition_expr', ''))
        self.action_combo.setCurrentText(step.get('action', ACTION_CLICK))
        self.locator_type_combo.setCurrentText(step.get('locator_type', 'xpath'))
        self.locator_value_edit.setText(step.get('locator_value', ''))
        self.target_locator_type_combo.setCurrentText(step.get('target_locator_type', 'xpath'))
        self.target_locator_value_edit.setText(step.get('target_locator_value', ''))
        self.drop_position_combo.setCurrentText(step.get('drop_position', '中间'))
        self.drag_offset_x_spin.setValue(int(step.get('drag_offset_x', 0) or 0))
        self.drag_offset_y_spin.setValue(int(step.get('drag_offset_y', 0) or 0))
        self.value_template_edit.setPlainText(step.get('value_template', ''))
        self.newline_mode_combo.setCurrentText(step.get('newline_mode', '直接输入'))
        self.tab_mode_combo.setCurrentText(step.get('tab_mode', '直接输入'))
        self.space_mode_combo.setCurrentText(step.get('space_mode', '直接输入'))
        self.wait_timeout_spin.setValue(float(step.get('wait_timeout', 10) or 10))
        self.window_match_type_combo.setCurrentText(step.get('window_match_type', '标题包含'))
        self.window_match_value_edit.setText(step.get('window_match_value', ''))
        self.sleep_seconds_spin.setValue(float(step.get('sleep_seconds', 1) or 1))
        self.use_js_click_check.setChecked(bool(step.get('use_js_click', False)))
        self.clear_before_input_check.setChecked(bool(step.get('clear_before_input', True)))
        self.wait_clickable_check.setChecked(bool(step.get('wait_clickable', False)))
        self.note_edit.setPlainText(step.get('note', ''))
        self._loading_step = False
        self.update_action_visibility(self.action_combo.currentText())

    def current_step_dict(self):
        return {
            'name': self.step_name_edit.text().strip(),
            'condition_expr': self.step_condition_edit.text().strip(),
            'action': self.action_combo.currentText(),
            'locator_type': self.locator_type_combo.currentText(),
            'locator_value': self.locator_value_edit.text().strip(),
            'target_locator_type': self.target_locator_type_combo.currentText(),
            'target_locator_value': self.target_locator_value_edit.text().strip(),
            'drop_position': self.drop_position_combo.currentText(),
            'drag_offset_x': self.drag_offset_x_spin.value(),
            'drag_offset_y': self.drag_offset_y_spin.value(),
            'value_template': self.value_template_edit.toPlainText(),
            'newline_mode': self.newline_mode_combo.currentText(),
            'tab_mode': self.tab_mode_combo.currentText(),
            'space_mode': self.space_mode_combo.currentText(),
            'wait_timeout': self.wait_timeout_spin.value(),
            'window_match_type': self.window_match_type_combo.currentText(),
            'window_match_value': self.window_match_value_edit.text().strip(),
            'sleep_seconds': self.sleep_seconds_spin.value(),
            'use_js_click': self.use_js_click_check.isChecked(),
            'clear_before_input': self.clear_before_input_check.isChecked(),
            'wait_clickable': self.wait_clickable_check.isChecked(),
            'note': self.note_edit.toPlainText(),
        }

    def apply_current_step_changes(self, silent=False):
        if self._loading_step:
            return
        row = self.step_list.currentRow()
        if row < 0:
            return
        self.flow_config['steps'][row] = self.current_step_dict()
        self.refresh_step_list()
        self.step_list.setCurrentRow(row)
        if not silent:
            self.log('已应用当前步骤修改。')

    def update_action_visibility(self, action):
        locator_needed = action in (ACTION_CLICK, ACTION_INPUT, ACTION_WAIT_ELEMENT, ACTION_SWITCH_IFRAME, ACTION_DRAG)
        value_needed = action == ACTION_INPUT
        window_needed = action == ACTION_SWITCH_WINDOW
        sleep_needed = action == ACTION_SLEEP
        drag_needed = action == ACTION_DRAG

        self._set_form_row_visible(self.locator_type_combo, locator_needed)
        self._set_form_row_visible(self.locator_value_edit, locator_needed)
        self._set_form_row_visible(self.target_locator_type_combo, drag_needed)
        self._set_form_row_visible(self.target_locator_value_edit, drag_needed)
        self._set_form_row_visible(self.drop_position_combo, drag_needed)
        self._set_form_row_visible(self.drag_offset_widget, drag_needed)
        self._set_form_row_visible(self.value_template_edit, value_needed)
        self._set_form_row_visible(self.field_insert_widget, value_needed)
        self._set_form_row_visible(self.newline_mode_combo, value_needed)
        self._set_form_row_visible(self.tab_mode_combo, value_needed)
        self._set_form_row_visible(self.space_mode_combo, value_needed)
        self._set_form_row_visible(self.window_match_type_combo, window_needed)
        self._set_form_row_visible(self.window_match_value_edit, window_needed)
        self._set_form_row_visible(self.sleep_seconds_spin, sleep_needed)
        self.use_js_click_check.setVisible(action == ACTION_CLICK)
        self.clear_before_input_check.setVisible(action == ACTION_INPUT)
        self.wait_clickable_check.setVisible(action == ACTION_WAIT_ELEMENT)

    def insert_template_text(self, text):
        cursor = self.value_template_edit.textCursor()
        cursor.insertText(text)
        self.value_template_edit.setTextCursor(cursor)

    def insert_selected_field(self):
        field = self.field_combo.currentText().strip()
        if field:
            self.insert_template_text(f'{{{field}}}')

    def launch_browser(self):
        try:
            self.normalize_start_url_input()
            browser = self._collect_browser_settings(strict=True)
            self.engine.launch_browser(
                chromedriver_path=browser.get('chromedriver_path', ''),
                chrome_binary=browser.get('chrome_binary', ''),
                start_url=browser.get('start_url', ''),
                debug_port=browser.get('debug_port', 9222),
                logger=self.log,
            )
            self.refresh_windows()
        except Exception as e:
            QMessageBox.critical(self, '错误', f'启动浏览器失败：{e}')
            self.log(f'启动浏览器失败：{e}')

    def refresh_windows(self):
        """Refresh the list of open browser windows.

        This method will not attempt to launch a new browser session.  If the
        engine is not currently connected to a running browser (i.e. there is
        no existing WebDriver instance), it simply logs a message and exits.
        This avoids inadvertently starting a new browser window when the user
        only wants to view the current list of windows.
        """
        # If there is no active browser connection, do not start one implicitly.
        if not getattr(self.engine, 'is_connected', lambda: False)():
            # Clear the combo box but leave any previous selection untouched.
            self.window_combo.blockSignals(True)
            self.window_combo.clear()
            self.window_combo.blockSignals(False)
            self.log('浏览器未连接，无法刷新窗口列表。请先启动或连接浏览器。')
            return

        selected = self.window_combo.currentData() or {}
        selected_handle = selected.get('handle') or getattr(self.engine, 'preferred_window_handle', None)
        self.window_combo.blockSignals(True)
        self.window_combo.clear()
        try:
            windows = self.engine.list_windows()
            target_index = -1
            for idx, item in enumerate(windows):
                text = f"[{item['index']}] {self._truncate_window_text(item.get('title'), item.get('url'))}"
                self.window_combo.addItem(text, item)
                if item.get('handle') == selected_handle and target_index < 0:
                    target_index = idx
            if target_index < 0 and windows:
                target_index = 0
            if target_index >= 0:
                self.window_combo.setCurrentIndex(target_index)
            self.log(f'已获取窗口数量：{len(windows)}')
        except Exception as e:
            self.log(f'刷新窗口列表失败：{e}')
        finally:
            self.window_combo.blockSignals(False)
        # Automatically set the selected window as the target if any exist.
        if self.window_combo.count() > 0:
            self.set_selected_window_as_target(auto=True)

    def switch_selected_window(self):
        if self.window_combo.count() <= 0:
            QMessageBox.information(self, '提示', '当前没有可选择的浏览器窗口。')
            return
        self.set_selected_window_as_target(auto=False)

    def add_switch_window_step_from_selection(self):
        item = self.window_combo.currentData()
        if not item:
            QMessageBox.information(self, '提示', '请先选择一个浏览器窗口。')
            return
        title = (item.get('title') or '').strip()
        url = (item.get('url') or '').strip()
        match_type = '标题包含' if title else 'URL包含'
        match_value = title if title else url
        step = {
            'name': '切换到目标窗口',
            'condition_expr': '',
            'action': ACTION_SWITCH_WINDOW,
            'locator_type': 'xpath',
            'locator_value': '',
            'target_locator_type': 'xpath',
            'target_locator_value': '',
            'drop_position': '中间',
            'drag_offset_x': 0,
            'drag_offset_y': 0,
            'value_template': '',
            'newline_mode': '直接输入',
            'tab_mode': '直接输入',
            'space_mode': '直接输入',
            'wait_timeout': 10,
            'window_match_type': match_type,
            'window_match_value': match_value,
            'sleep_seconds': 1,
            'use_js_click': False,
            'clear_before_input': True,
            'wait_clickable': False,
            'note': '由当前选中窗口自动生成',
        }
        self.flow_config.setdefault('steps', []).append(step)
        self.refresh_step_list()
        self.step_list.setCurrentRow(self.step_list.count() - 1)

    def start_element_recording(self):
        if not getattr(self.engine, 'is_connected', lambda: False)():
            QMessageBox.information(self, '提示', '浏览器未连接，请先启动或连接浏览器后再开始自动录制。')
            self.log('浏览器未连接，无法启动自动录制。')
            return
        try:
            self.record_timer.stop()
            self.recorded_element = None
            self.record_text.setPlainText('自动录制已启动，请切换到浏览器点击目标元素。\n点击将只用于录制定位信息，不建议连续快速点击。')
            self.engine.start_element_recording(logger=self.log)
            self.record_timer.start()
        except Exception as e:
            QMessageBox.critical(self, '错误', f'启动自动录制失败：{e}')
            self.log(f'启动自动录制失败：{e}')

    def poll_recorded_element(self):
        try:
            info = self.engine.poll_recorded_element(consume=True)
        except Exception as e:
            self.record_timer.stop()
            self.log(f'自动录制失败：{e}')
            self.record_text.setPlainText(f'自动录制失败：{e}')
            return
        if not info:
            return
        self.record_timer.stop()
        self.recorded_element = info
        self.record_text.setPlainText(self.format_recorded_element_info(info))
        locator_type, locator_value = self.choose_best_locator(info)
        self.log(f'已录制元素：{locator_type} = {locator_value}')

    @staticmethod
    def format_recorded_element_info(element):
        if not element:
            return '暂无录制结果。'
        if element.get('error'):
            return f"录制失败：{element.get('error')}"
        lines = [
            f"标签: {element.get('tag', '')}",
            f"文本: {element.get('text', '')}",
            f"id: {element.get('id', '')}",
            f"name: {element.get('name', '')}",
            f"placeholder: {element.get('placeholder', '')}",
            f"title: {element.get('title', '')}",
            f"XPath: {element.get('xpath', '')}",
            f"可点击XPath: {element.get('clickable_xpath', '')}",
            f"CSS: {element.get('css', '')}",
            f"推荐定位方式: {element.get('recommended_locator_type', '')}",
            f"推荐定位值: {element.get('recommended_locator_value', '')}",
        ]
        frame_chain = element.get('frame_chain') or []
        if frame_chain:
            lines.append('所在iframe链:')
            for idx, item in enumerate(frame_chain, start=1):
                lines.append(f'  {idx}. {item}')
        return '\n'.join(lines)

    def selected_element_info(self):
        return self.recorded_element

    @staticmethod
    def choose_best_locator(element):
        if not element:
            return 'xpath', ''
        if element.get('recommended_locator_type') and element.get('recommended_locator_value'):
            return element.get('recommended_locator_type'), element.get('recommended_locator_value')
        if element.get('id'):
            return 'id', element['id']
        if element.get('name'):
            return 'name', element['name']
        if element.get('clickable_xpath'):
            return 'xpath', element['clickable_xpath']
        if element.get('xpath'):
            return 'xpath', element['xpath']
        if element.get('clickable_css'):
            return 'css selector', element['clickable_css']
        if element.get('css'):
            return 'css selector', element['css']
        return 'xpath', ''

    def apply_selected_element_to_step(self, *args, **kwargs):
        element = self.selected_element_info()
        if not element:
            QMessageBox.information(self, '提示', '请先完成一次自动录制。')
            return
        locator_type, locator_value = self.choose_best_locator(element)
        action = self.action_combo.currentText()
        target_mode = action == ACTION_DRAG and bool(self.locator_value_edit.text().strip()) and not bool(self.target_locator_value_edit.text().strip())
        if action not in (ACTION_CLICK, ACTION_INPUT, ACTION_WAIT_ELEMENT, ACTION_SWITCH_IFRAME, ACTION_DRAG):
            action = ACTION_INPUT if element.get('tag') in ('input', 'textarea', 'select') else ACTION_CLICK
            self.action_combo.setCurrentText(action)
        if target_mode:
            self.target_locator_type_combo.setCurrentText(locator_type)
            self.target_locator_value_edit.setText(locator_value)
            self.log(f'已将录制结果填入拖拽目标：{locator_type} = {locator_value}')
            return
        self.locator_type_combo.setCurrentText(locator_type)
        self.locator_value_edit.setText(locator_value)
        if not self.step_name_edit.text().strip():
            desc = element.get('text') or element.get('placeholder') or element.get('id') or element.get('name') or element.get('tag')
            self.step_name_edit.setText(f'{action}-{desc}')
        self.log(f'已将录制结果填入当前步骤：{locator_type} = {locator_value}')

    def add_step_from_element(self, default_action=ACTION_CLICK):
        element = self.selected_element_info()
        if not element:
            QMessageBox.information(self, '提示', '请先完成一次自动录制。')
            return
        locator_type, locator_value = self.choose_best_locator(element)
        desc = element.get('text') or element.get('placeholder') or element.get('id') or element.get('name') or element.get('tag')
        step = {
            'name': f'{default_action}-{desc}',
            'condition_expr': '',
            'action': default_action,
            'locator_type': locator_type,
            'locator_value': locator_value,
            'target_locator_type': 'xpath',
            'target_locator_value': '',
            'drop_position': '中间',
            'drag_offset_x': 0,
            'drag_offset_y': 0,
            'value_template': '{__RESULT__}' if default_action == ACTION_INPUT else '',
            'newline_mode': '直接输入',
            'tab_mode': '直接输入',
            'space_mode': '直接输入',
            'wait_timeout': 10,
            'window_match_type': '标题包含',
            'window_match_value': '',
            'sleep_seconds': 1,
            'use_js_click': False,
            'clear_before_input': True,
            'wait_clickable': False,
            'note': '由自动录制生成',
        }
        self.flow_config.setdefault('steps', []).append(step)
        self.refresh_step_list()
        self.step_list.setCurrentRow(self.step_list.count() - 1)

    def test_import(self):
        parent = self.parent()
        if parent is None or not hasattr(parent, '_last_render_result_text'):
            QMessageBox.information(self, '提示', '未找到主窗口上下文。')
            return
        try:
            parent.update_result_text(force=True)
            payload = {
                'template_name': parent.current_template_name,
                'result_text': parent._last_render_result_text,
                'final_fields': parent._last_final_fields,
                'input_values': parent._last_input_values,
                'data_pool': getattr(parent, '_last_data_pool', {}) or {},
            }
            flow = self.collect_flow(strict=True)
            self.engine.execute_flow(flow, payload, logger=self.log)
            QMessageBox.information(self, '成功', '测试导入执行完成。')
        except Exception as e:
            QMessageBox.critical(self, '错误', f'测试导入失败：{e}')
            self.log(f'测试导入失败：{e}')

    def closeEvent(self, event):
        try:
            self.record_timer.stop()
            self.engine.stop_element_recording(logger=self.log)
        except Exception:
            pass
        if self.has_unsaved_changes():
            reply = QMessageBox.question(self, '未保存', '浏览器流程配置已更改，是否保存？', QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Yes:
                if self.save_flow():
                    event.accept()
                else:
                    event.ignore()
                    return
            elif reply == QMessageBox.No:
                event.accept()
            else:
                event.ignore()
                return
        else:
            event.accept()
