import copy
import json
from log import log_change
from PyQt5.QtCore import Qt, QTimer, QEvent, QSize
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
    QScrollArea,
    QFrame,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QStyle,
    QApplication,
    QSizePolicy,
)

from PyQt5.QtGui import QFontMetrics, QPalette
from gui_helpers import signal_blocked




def qt_safe_single_shot(delay_ms, callback):
    """安全延后执行 PyQt 回调，避免窗口关闭后定时回调访问已删除控件。"""
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

# 浏览器流程“步骤编辑”表单的最小稳定布局宽度。窗口缩得过窄时不再继续压缩标签/输入框，避免长内容反复重排。
step_editor_min_layout_width = 520




ACTION_CLICK = '点击元素'
ACTION_INPUT = '输入文本'
ACTION_MULTI_SELECT = '多选元素/多行选择'
ACTION_RIGHT_CLICK = '右键点击'
ACTION_KEY_COMBO = '键盘组合键'  # 兼容旧流程，界面新建步骤请使用“键盘按键”或“全选操作”
ACTION_KEY_PRESS = '键盘按键'
ACTION_SELECT_ALL = '全选操作'
ACTION_RIGHT_CLICK_MENU = '右键菜单项点击'
ACTION_DROPDOWN_TWO_STAGE = '下拉菜单两段式操作'
ACTION_WAIT_ELEMENT = '等待元素'
ACTION_WAIT_ELEMENT_GONE = '等待元素消失'
ACTION_WAIT_NEW_WINDOW = '等待新窗口'
ACTION_WAIT_WINDOW_BACK_MAIN = '等待窗口关闭/等待回到主窗口'
ACTION_SWITCH_WINDOW = '切换窗口'
ACTION_SWITCH_MAIN_WINDOW = '切回主窗口'
ACTION_SWITCH_IFRAME = '切换iframe'
ACTION_SWITCH_DEFAULT = '切回默认文档'
ACTION_SLEEP = '延时'
ACTION_DRAG = '拖拽元素'
ACTION_PAGE_CONDITION = '页面条件判断'
ACTION_ADD_TABLE = '添加表格'
ACTION_FILL_TABLE = '自动填单元格'
ACTION_CUSTOM_JS = '执行JavaScript命令'
ACTION_MOUSE_MULTI_CLICK = '鼠标连击'
ACTION_LOOP_START = '循环开始'
ACTION_LOOP_END = '循环结束'
ACTION_JUMP_STEP = '跳转至步骤'

STEP_ACTIONS = [
    ACTION_CLICK,
    ACTION_INPUT,
    ACTION_MULTI_SELECT,
    ACTION_RIGHT_CLICK,
    ACTION_SELECT_ALL,
    ACTION_KEY_PRESS,
    ACTION_RIGHT_CLICK_MENU,
    ACTION_DROPDOWN_TWO_STAGE,
    ACTION_WAIT_ELEMENT,
    ACTION_WAIT_ELEMENT_GONE,
    ACTION_WAIT_NEW_WINDOW,
    ACTION_WAIT_WINDOW_BACK_MAIN,
    ACTION_SWITCH_WINDOW,
    ACTION_SWITCH_MAIN_WINDOW,
    ACTION_SWITCH_IFRAME,
    ACTION_SWITCH_DEFAULT,
    ACTION_SLEEP,
    ACTION_DRAG,
    ACTION_PAGE_CONDITION,
    ACTION_ADD_TABLE,
    ACTION_FILL_TABLE,
    ACTION_CUSTOM_JS,
    ACTION_MOUSE_MULTI_CLICK,
    ACTION_LOOP_START,
    ACTION_LOOP_END,
    ACTION_JUMP_STEP,
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
        self._updating_step_form_widths = False
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
        self.browser_connection_group = browser_group
        browser_layout = QVBoxLayout(browser_group)
        browser_layout.setContentsMargins(8, 8, 8, 8)
        browser_layout.setSpacing(4)
        self._browser_connection_labels = []

        self.chromedriver_edit = QLineEdit()
        self.chrome_binary_edit = QLineEdit()
        self.debug_port_edit = QLineEdit('9222')
        self.start_url_edit = QLineEdit()
        self.start_url_edit.editingFinished.connect(self.normalize_start_url_input)
        self.implicit_wait_edit = QLineEdit('0.5')
        for edit in (self.chromedriver_edit, self.chrome_binary_edit, self.start_url_edit, self.debug_port_edit, self.implicit_wait_edit):
            # 让浏览器连接区输入框按当前可用宽度动态填充，右侧贴近连接框右边界。
            edit.setMinimumWidth(80)
            edit.setMaximumWidth(16777215)
            edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        def add_browser_label(row_layout, text):
            label = QLabel(text)
            label.setWordWrap(True)
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Preferred)
            self._browser_connection_labels.append(label)
            row_layout.addWidget(label, 0, Qt.AlignLeft | Qt.AlignVCenter)
            return label

        def add_browser_setting_row(label_text, edit_widget):
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(6)
            add_browser_label(row_layout, label_text)
            row_layout.addWidget(edit_widget, 1, Qt.AlignVCenter)
            browser_layout.addWidget(row)
            return row

        def add_browser_setting_pair_row(first_label, first_edit, second_label, second_edit):
            """同一行放置两组浏览器参数。

            每个标签仍按自身文字和浏览器连接区宽度计算宽度；
            两个输入框分别吃掉所在半行的剩余空间，第二个输入框右侧贴近连接框右边界。
            """
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(6)
            add_browser_label(row_layout, first_label)
            row_layout.addWidget(first_edit, 1, Qt.AlignVCenter)
            row_layout.addSpacing(8)
            add_browser_label(row_layout, second_label)
            row_layout.addWidget(second_edit, 1, Qt.AlignVCenter)
            browser_layout.addWidget(row)
            return row

        add_browser_setting_row('chromedriver:', self.chromedriver_edit)
        add_browser_setting_row('Chrome路径:', self.chrome_binary_edit)
        add_browser_setting_row('启动网址:', self.start_url_edit)
        add_browser_setting_pair_row('调试端口:', self.debug_port_edit, '隐式等待(秒):', self.implicit_wait_edit)

        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(0, 2, 0, 0)
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
        btn_row.addStretch(1)
        browser_layout.addLayout(btn_row)

        window_row = QHBoxLayout()
        window_row.setContentsMargins(0, 0, 0, 0)
        self.window_combo = QComboBox()
        self.window_combo.setMinimumWidth(260)
        self.window_combo.setMaximumWidth(420)
        self.window_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.window_combo.currentIndexChanged.connect(self.on_window_selection_changed)
        add_browser_label(window_row, '已打开窗口:')
        window_row.addWidget(self.window_combo, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.switch_window_btn = QPushButton('使用选中窗口')
        self.switch_window_btn.clicked.connect(self.switch_selected_window)
        window_row.addWidget(self.switch_window_btn)
        self.add_window_step_btn = QPushButton('将选中窗口生成步骤')
        self.add_window_step_btn.clicked.connect(self.add_switch_window_step_from_selection)
        window_row.addWidget(self.add_window_step_btn)
        window_row.addStretch(1)
        browser_layout.addLayout(window_row)

        elements_group = QGroupBox('自动录制结果（只读）')
        eg_layout = QVBoxLayout(elements_group)
        self.record_text = QPlainTextEdit()
        self.record_text.setReadOnly(True)
        self.record_text.setPlaceholderText('点击“自动录制”后，切换到浏览器点击目标元素，XPath / 推荐定位信息会显示在这里。\n当前录制结果只显示，不再自动填充步骤。')
        eg_layout.addWidget(self.record_text, 1)

        top_splitter = QSplitter(Qt.Horizontal)
        top_splitter.setChildrenCollapsible(False)
        top_splitter.addWidget(browser_group)
        top_splitter.addWidget(elements_group)
        top_splitter.setSizes([620, 520])
        main_layout.addWidget(top_splitter)
        self._update_browser_connection_label_widths()

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
        self.copy_step_btn = QPushButton('复制步骤')
        self.copy_step_btn.clicked.connect(self.copy_step)
        step_btns.addWidget(self.copy_step_btn)
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
        self.save_flow_btn = QPushButton('手动保存流程')
        self.save_flow_btn.clicked.connect(self.save_flow)
        save_btns.addWidget(self.save_flow_btn)
        step_layout.addLayout(save_btns)

        editor_panel = QWidget()
        editor_panel.setMinimumWidth(420)
        editor_layout = QVBoxLayout(editor_panel)
        editor_layout.setContentsMargins(0, 0, 0, 0)

        self.step_scroll = QScrollArea()
        self.step_scroll.setWidgetResizable(True)
        self.step_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.step_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.step_scroll.setFrameShape(QFrame.NoFrame)
        self.step_scroll_panel = QWidget()
        step_scroll_layout = QVBoxLayout(self.step_scroll_panel)
        step_scroll_layout.setContentsMargins(5, 5, 5, 5)
        step_scroll_layout.setSpacing(5)

        editor_group = QGroupBox('步骤编辑')
        self.step_form = QFormLayout(editor_group)
        self.step_form.setRowWrapPolicy(QFormLayout.DontWrapRows)
        self.step_form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.step_form.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        form = self.step_form
        self.step_name_edit = QLineEdit()
        self.step_condition_edit = QLineEdit()
        self.step_condition_edit.setPlaceholderText("留空=始终执行；例如 {是否二段硫化} == '是' 或 {二段硫化段落} != ''")
        self.action_combo = QComboBox()
        self.action_combo.addItems(STEP_ACTIONS)
        self.action_combo.currentTextChanged.connect(self.update_action_visibility)
        self.locator_type_combo = QComboBox()
        self.locator_type_combo.addItems(LOCATOR_TYPES)
        self.locator_value_edit = QPlainTextEdit()
        self.locator_value_edit.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.locator_value_edit.setMinimumHeight(46)
        self.locator_value_edit.setMaximumHeight(96)
        self.locator_value_edit.setPlaceholderText('路径较长时可自动换行，支持循环标记 [[i]]，例如 //a[[i]]')
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
        self.value_template_edit.setPlaceholderText('普通动作可使用 {字段名}、{__RESULT__}、#if(...)#；执行JavaScript命令请使用 [[字段名]]、[[__RESULT__]]')
        self.page_condition_expr_edit = QLineEdit()
        self.page_condition_expr_edit.setPlaceholderText('页面条件判断时可填表达式；留空则按元素定位检测')
        self.detect_mode_combo = QComboBox()
        self.detect_mode_combo.addItems(['立即判断', '等待判断'])
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
        self.mouse_click_widget = QWidget()
        mouse_click_layout = QHBoxLayout(self.mouse_click_widget)
        mouse_click_layout.setContentsMargins(0, 0, 0, 0)
        mouse_click_layout.addWidget(QLabel('按键:'))
        self.mouse_click_button_combo = QComboBox()
        self.mouse_click_button_combo.addItems(['左键连击', '右键连击'])
        mouse_click_layout.addWidget(self.mouse_click_button_combo)
        mouse_click_layout.addWidget(QLabel('次数:'))
        self.mouse_click_count_spin = QSpinBox()
        self.mouse_click_count_spin.setRange(1, 20)
        self.mouse_click_count_spin.setValue(2)
        mouse_click_layout.addWidget(self.mouse_click_count_spin)
        mouse_click_layout.addStretch()

        self.keyboard_action_widget = QWidget()
        keyboard_action_layout = QHBoxLayout(self.keyboard_action_widget)
        keyboard_action_layout.setContentsMargins(0, 0, 0, 0)
        keyboard_action_layout.addWidget(QLabel('方式:'))
        self.keyboard_press_mode_combo = QComboBox()
        self.keyboard_press_mode_combo.addItems(['点按', '长按'])
        keyboard_action_layout.addWidget(self.keyboard_press_mode_combo)
        keyboard_action_layout.addWidget(QLabel('长按秒数:'))
        self.keyboard_hold_seconds_spin = QDoubleSpinBox()
        self.keyboard_hold_seconds_spin.setRange(0.1, 30)
        self.keyboard_hold_seconds_spin.setDecimals(1)
        self.keyboard_hold_seconds_spin.setValue(1.0)
        keyboard_action_layout.addWidget(self.keyboard_hold_seconds_spin)
        keyboard_action_layout.addStretch()

        self.jump_step_widget = QWidget()
        jump_step_layout = QHBoxLayout(self.jump_step_widget)
        jump_step_layout.setContentsMargins(0, 0, 0, 0)
        jump_step_layout.addWidget(QLabel('跳转到第'))
        self.jump_step_spin = QSpinBox()
        self.jump_step_spin.setRange(1, 9999)
        self.jump_step_spin.setValue(1)
        jump_step_layout.addWidget(self.jump_step_spin)
        jump_step_layout.addWidget(QLabel('步'))
        jump_step_layout.addStretch()

        self.loop_end_widget = QWidget()
        loop_end_layout = QHBoxLayout(self.loop_end_widget)
        loop_end_layout.setContentsMargins(0, 0, 0, 0)
        loop_end_layout.addWidget(QLabel('循环标记:'))
        self.loop_marker_edit = QLineEdit('i')
        self.loop_marker_edit.setMaximumWidth(90)
        loop_end_layout.addWidget(self.loop_marker_edit)
        loop_end_layout.addWidget(QLabel('起始:'))
        self.loop_start_spin = QSpinBox()
        self.loop_start_spin.setRange(1, 9999)
        self.loop_start_spin.setValue(1)
        loop_end_layout.addWidget(self.loop_start_spin)
        loop_end_layout.addWidget(QLabel('结束:'))
        self.loop_stop_spin = QSpinBox()
        self.loop_stop_spin.setRange(1, 9999)
        self.loop_stop_spin.setValue(1)
        loop_end_layout.addWidget(self.loop_stop_spin)
        loop_end_layout.addWidget(QLabel('循环块内定位值可写 [[i]]、[[j]]'))
        loop_end_layout.addStretch()

        self.loop_widget = QWidget()
        loop_layout = QHBoxLayout(self.loop_widget)
        loop_layout.setContentsMargins(0, 0, 0, 0)
        self.loop_enabled_check = QCheckBox('启用循环内容表')
        loop_layout.addWidget(self.loop_enabled_check)
        loop_layout.addWidget(QLabel('行数:'))
        self.loop_content_rows_spin = QSpinBox()
        self.loop_content_rows_spin.setRange(1, 9999)
        self.loop_content_rows_spin.setValue(1)
        loop_layout.addWidget(self.loop_content_rows_spin)
        loop_layout.addWidget(QLabel('列数:'))
        self.loop_content_cols_spin = QSpinBox()
        self.loop_content_cols_spin.setRange(1, 30)
        self.loop_content_cols_spin.setValue(1)
        loop_layout.addWidget(self.loop_content_cols_spin)
        loop_layout.addWidget(QLabel('列标记:'))
        self.loop_columns_edit = QLineEdit('内容')
        self.loop_columns_edit.setPlaceholderText('逗号分隔，例如：内容,j,备注；列名可在模板/定位中用 [[列名]]')
        loop_layout.addWidget(self.loop_columns_edit, 1)
        self.loop_sync_table_btn = QPushButton('同步表格')
        self.loop_sync_table_btn.clicked.connect(self.sync_loop_value_table_rows)
        loop_layout.addWidget(self.loop_sync_table_btn)

        self.loop_value_table = QTableWidget(0, 2)
        self.loop_value_table.setHorizontalHeaderLabels(['序号', '内容'])
        self.loop_value_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.loop_value_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.loop_value_table.setMinimumHeight(90)
        self.loop_value_table.setMaximumHeight(220)

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
        self.insert_result_btn.clicked.connect(self.insert_result_placeholder)

        self.field_insert_widget = QWidget()
        field_row = QHBoxLayout(self.field_insert_widget)
        field_row.setContentsMargins(0, 0, 0, 0)
        field_row.addWidget(self.field_combo, 1)
        field_row.addWidget(self.insert_field_btn)
        field_row.addWidget(self.insert_result_btn)

        self.on_found_step_spin = QSpinBox()
        self.on_found_step_spin.setRange(0, 9999)
        self.on_not_found_step_spin = QSpinBox()
        self.on_not_found_step_spin.setRange(0, 9999)
        self.on_timeout_step_spin = QSpinBox()
        self.on_timeout_step_spin.setRange(0, 9999)
        self.on_found_message_edit = QLineEdit()
        self.on_not_found_message_edit = QLineEdit()
        self.on_timeout_message_edit = QLineEdit()

        self.branch_found_widget = QWidget()
        found_row = QHBoxLayout(self.branch_found_widget)
        found_row.setContentsMargins(0, 0, 0, 0)
        found_row.addWidget(QLabel('跳转步骤:'))
        found_row.addWidget(self.on_found_step_spin)
        found_row.addWidget(QLabel('提示:'))
        found_row.addWidget(self.on_found_message_edit, 1)

        self.branch_not_found_widget = QWidget()
        not_found_row = QHBoxLayout(self.branch_not_found_widget)
        not_found_row.setContentsMargins(0, 0, 0, 0)
        not_found_row.addWidget(QLabel('跳转步骤:'))
        not_found_row.addWidget(self.on_not_found_step_spin)
        not_found_row.addWidget(QLabel('提示:'))
        not_found_row.addWidget(self.on_not_found_message_edit, 1)

        self.branch_timeout_widget = QWidget()
        timeout_row = QHBoxLayout(self.branch_timeout_widget)
        timeout_row.setContentsMargins(0, 0, 0, 0)
        timeout_row.addWidget(QLabel('跳转步骤:'))
        timeout_row.addWidget(self.on_timeout_step_spin)
        timeout_row.addWidget(QLabel('提示:'))
        timeout_row.addWidget(self.on_timeout_message_edit, 1)

        form.addRow('步骤名称:', self.step_name_edit)
        form.addRow('执行条件:', self.step_condition_edit)
        form.addRow('动作类型:', self.action_combo)
        form.addRow('定位方式:', self.locator_type_combo)
        form.addRow('定位值:', self.locator_value_edit)
        form.addRow('目标定位方式:', self.target_locator_type_combo)
        form.addRow('目标定位值:', self.target_locator_value_edit)
        form.addRow('释放位置:', self.drop_position_combo)
        form.addRow('拖拽偏移:', self.drag_offset_widget)
        form.addRow('输入/参数模板:', self.value_template_edit)
        form.addRow('字段插入:', self.field_insert_widget)
        form.addRow('页面判断表达式:', self.page_condition_expr_edit)
        form.addRow('检测模式:', self.detect_mode_combo)
        form.addRow('找到分支:', self.branch_found_widget)
        form.addRow('找不到分支:', self.branch_not_found_widget)
        form.addRow('超时分支:', self.branch_timeout_widget)
        form.addRow('回车/换行处理:', self.newline_mode_combo)
        form.addRow('Tab/缩进处理:', self.tab_mode_combo)
        form.addRow('空格处理:', self.space_mode_combo)
        form.addRow('等待超时(秒):', self.wait_timeout_spin)
        form.addRow('窗口匹配方式:', self.window_match_type_combo)
        form.addRow('窗口匹配值:', self.window_match_value_edit)
        form.addRow('延时秒数:', self.sleep_seconds_spin)
        form.addRow('连击设置:', self.mouse_click_widget)
        form.addRow('键盘设置:', self.keyboard_action_widget)
        form.addRow('跳转目标:', self.jump_step_widget)
        form.addRow('循环参数:', self.loop_end_widget)
        form.addRow('输入循环:', self.loop_widget)
        form.addRow('循环内容表:', self.loop_value_table)
        form.addRow(self.use_js_click_check)
        form.addRow(self.clear_before_input_check)
        form.addRow(self.wait_clickable_check)
        form.addRow('备注:', self.note_edit)
        step_scroll_layout.addWidget(editor_group)
        step_scroll_layout.addStretch()
        self.step_scroll.setWidget(self.step_scroll_panel)
        try:
            self.step_scroll.viewport().installEventFilter(self)
            self.step_scroll.installEventFilter(self)
        except Exception:
            pass
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText('浏览器连接、窗口切换、元素抓取日志会显示在这里。')
        editor_splitter = QSplitter(Qt.Vertical)
        editor_splitter.setChildrenCollapsible(False)
        editor_splitter.addWidget(self.step_scroll)
        editor_splitter.addWidget(self.log_text)
        editor_splitter.setSizes([620, 220])
        editor_layout.addWidget(editor_splitter, 1)

        splitter.addWidget(step_panel)
        splitter.addWidget(editor_panel)
        splitter.setSizes([300, 900])
        self.update_action_visibility(self.action_combo.currentText())
        self._update_step_form_widths()
        self._connect_auto_save_signals()


    def _set_form_row_visible(self, field_widget, visible):
        label = self.step_form.labelForField(field_widget) if hasattr(self, 'step_form') else None
        if label is not None:
            label.setVisible(visible)
        field_widget.setVisible(visible)



    def _update_browser_connection_label_widths(self):
        """浏览器连接区标签按文字自适应，过长时最多占连接区宽度四分之一。"""
        labels = getattr(self, '_browser_connection_labels', []) or []
        if not labels:
            return
        try:
            group_width = self.browser_connection_group.width() if hasattr(self, 'browser_connection_group') else self.width()
        except Exception:
            group_width = self.width()
        max_width = max(60, int((group_width or 240) / 4))
        for label in labels:
            try:
                label.setMaximumWidth(max_width)
                metrics = label.fontMetrics()
                try:
                    natural = metrics.horizontalAdvance(label.text()) + 4
                except AttributeError:
                    natural = metrics.width(label.text()) + 4
                label.setMinimumWidth(0)
                label.setFixedWidth(min(max_width, max(1, int(natural))))
                label.setWordWrap(natural > max_width)
            except RuntimeError:
                continue
            except Exception:
                continue

    def _schedule_step_form_width_update(self):
        if getattr(self, '_step_form_width_update_pending', False):
            return
        self._step_form_width_update_pending = True

        def run_update():
            self._step_form_width_update_pending = False
            self._update_step_form_widths()

        qt_safe_single_shot(20, run_update)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        try:
            self._schedule_step_form_width_update()
        except Exception:
            pass
        try:
            self._update_browser_connection_label_widths()
        except Exception:
            pass

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Resize:
            try:
                if watched in (self.step_scroll, self.step_scroll.viewport()):
                    self._schedule_step_form_width_update()
            except Exception:
                pass
        return super().eventFilter(watched, event)

    def _update_step_form_widths(self):
        if not hasattr(self, 'step_form') or getattr(self, '_updating_step_form_widths', False):
            return
        self._updating_step_form_widths = True
        try:
            try:
                viewport_width = self.step_scroll.viewport().width() if hasattr(self, 'step_scroll') else self.width()
            except Exception:
                viewport_width = self.width()
            effective_width = max(int(step_editor_min_layout_width), int(viewport_width or 0))
            try:
                if hasattr(self, 'step_scroll_panel') and effective_width > 0:
                    self.step_scroll_panel.setMinimumWidth(0)
                    self.step_scroll_panel.setMaximumWidth(16777215)
                    self.step_scroll_panel.setFixedWidth(int(effective_width))
            except Exception:
                pass
            available_width = max(220, int(effective_width) - 30)
            label_max_width = max(120, available_width // 3)
            try:
                self.step_form.setRowWrapPolicy(QFormLayout.DontWrapRows)
                self.step_form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.step_form.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
            except Exception:
                pass
            natural_widths = []
            for row in range(self.step_form.rowCount()):
                label_item = self.step_form.itemAt(row, QFormLayout.LabelRole)
                label_widget = label_item.widget() if label_item is not None else None
                if label_widget is not None:
                    try:
                        text = label_widget.text() if hasattr(label_widget, 'text') else ''
                        metrics = label_widget.fontMetrics()
                        try:
                            natural_widths.append(metrics.horizontalAdvance(text) + 20)
                        except AttributeError:
                            natural_widths.append(metrics.width(text) + 20)
                    except Exception:
                        pass
            shared_label_width = min(label_max_width, max([96] + natural_widths))
            for row in range(self.step_form.rowCount()):
                label_item = self.step_form.itemAt(row, QFormLayout.LabelRole)
                field_item = self.step_form.itemAt(row, QFormLayout.FieldRole)
                label_widget = label_item.widget() if label_item is not None else None
                field_widget = field_item.widget() if field_item is not None else None
                label_width = shared_label_width
                if label_widget is not None:
                    try:
                        label_widget.setMinimumWidth(0)
                        label_widget.setMaximumWidth(16777215)
                        text = label_widget.text() if hasattr(label_widget, 'text') else ''
                        metrics = label_widget.fontMetrics()
                        try:
                            natural = metrics.horizontalAdvance(text) + 20
                        except AttributeError:
                            natural = metrics.width(text) + 20
                        label_widget.setFixedWidth(int(label_width))
                        if hasattr(label_widget, 'setWordWrap'):
                            label_widget.setWordWrap(natural > label_max_width)
                        if hasattr(label_widget, 'setAlignment'):
                            label_widget.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    except Exception:
                        pass
                if field_widget is not None:
                    try:
                        field_widget.setMinimumWidth(0)
                        field_widget.setMaximumWidth(16777215)
                        field_width = max(140, available_width - int(label_width) - 28)
                        field_widget.setFixedWidth(field_width)
                        if field_widget in (getattr(self, 'value_template_edit', None), getattr(self, 'note_edit', None)):
                            field_widget.setMinimumHeight(90)
                    except Exception:
                        pass
        finally:
            self._updating_step_form_widths = False


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
            'implicit_wait': 0.5,
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
        with signal_blocked(self.field_combo):
            self.field_combo.clear()
            for name in ordered:
                self.field_combo.addItem(name)
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
            with signal_blocked(widget):
                widget.setText('' if value is None else str(value))

        browser = self.flow_config.setdefault('browser', self._default_browser())
        browser.update({
            'connect_mode': 'launch',
            'chromedriver_path': browser_settings.get('chromedriver_path', browser.get('chromedriver_path', '')),
            'chrome_binary': browser_settings.get('chrome_binary', browser.get('chrome_binary', '')),
            'debug_port': self._parse_int_value(str(browser_settings.get('debug_port', browser.get('debug_port', 9222))), 9222, '调试端口'),
            'start_url': self.engine.normalize_url(browser_settings.get('start_url', browser.get('start_url', ''))),
            'implicit_wait': self._parse_float_value(str(browser_settings.get('implicit_wait', browser.get('implicit_wait', 0.5))), 0.5, '隐式等待'),
        })

        _set_text(self.chromedriver_edit, browser.get('chromedriver_path', ''))
        _set_text(self.chrome_binary_edit, browser.get('chrome_binary', ''))
        _set_text(self.debug_port_edit, browser.get('debug_port', 9222))
        _set_text(self.start_url_edit, browser.get('start_url', ''))
        _set_text(self.implicit_wait_edit, browser.get('implicit_wait', 0.5))

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
            'implicit_wait': self._parse_float_value(self.implicit_wait_edit.text(), 0.5, '隐式等待', strict=strict),
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
        self.implicit_wait_edit.setText(str(browser.get('implicit_wait', 0.5)))
        self.refresh_field_combo()
        self.refresh_step_list()
        self.refresh_windows()
        self._loaded_signature = self._signature(self.collect_flow())

    def collect_flow(self, strict=False):
        self._save_current_editor_to_flow(auto=False, persist=False)
        browser = self._default_browser()
        browser.update(self._collect_browser_settings(strict=strict))
        steps = [dict(step) for step in (self.flow_config.get('steps', []) or [])]
        return {'browser': browser, 'steps': steps}

    def has_unsaved_changes(self):
        self._save_current_editor_to_flow(auto=False, persist=False)
        return self._signature(self.collect_flow()) != self._loaded_signature

    def save_flow(self, silent=False):
        try:
            self._save_current_editor_to_flow(auto=False, persist=False)
            self.flow_config = self.collect_flow(strict=True)
            old_flow = None
            try:
                old_flow = json.loads(self._loaded_signature) if self._loaded_signature else None
            except Exception:
                old_flow = None
            self.template_db.update_browser_flow(self.template_name, self.flow_config)
            self._loaded_signature = self._signature(self.flow_config)
            if old_flow != self.flow_config:
                log_change(f'浏览器流程修改 - {self.template_name}', before=old_flow, after=self.flow_config)
            if not silent:
                self.log('浏览器流程配置已保存。')
                QMessageBox.information(self, '成功', '浏览器流程配置已保存。')
            return True
        except Exception as e:
            self.log(f'保存浏览器流程配置失败：{e}')
            if not silent:
                QMessageBox.critical(self, '错误', f'保存浏览器流程配置失败：{e}')
            return False

    def refresh_step_list(self, keep_row=None):
        current = self.step_list.currentRow() if keep_row is None else keep_row
        with signal_blocked(self.step_list):
            self.step_list.clear()
            steps = self.flow_config.get('steps', []) or []
            for idx, step in enumerate(steps, start=1):
                title = f"{idx}. {step.get('name') or step.get('action') or '未命名步骤'}"
                item = QListWidgetItem(title)
                item.setData(Qt.UserRole, step)
                self.step_list.addItem(item)
        if self.step_list.count():
            if current is None or current < 0:
                current = 0
            self.step_list.setCurrentRow(max(0, min(current, self.step_list.count() - 1)))
        else:
            self.clear_step_editor()

    def _new_step(self, **overrides):
        step = {
            'name': '',
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
            'page_condition_expr': '',
            'detect_mode': '等待判断',
            'on_found_step': 0,
            'on_not_found_step': 0,
            'on_timeout_step': 0,
            'on_found_message': '',
            'on_not_found_message': '',
            'on_timeout_message': '',
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
            'mouse_click_button': '左键连击',
            'mouse_click_count': 2,
            'keyboard_press_mode': '点按',
            'keyboard_hold_seconds': 1.0,
            'jump_step': 1,
            'loop_marker': 'i',
            'loop_start': 1,
            'loop_stop': 1,
            'loop_enabled': False,
            'loop_content_rows': 1,
            'loop_content_cols': 1,
            'loop_columns': ['内容'],
            'loop_values': [],
            'note': '',
        }
        step.update(overrides)
        return step

    @staticmethod
    def _normalize_action_for_ui(action):
        # 兼容旧版流程中的“键盘组合键”，新界面统一显示为“键盘按键”。
        if action == ACTION_KEY_COMBO:
            return ACTION_KEY_PRESS
        return action or ACTION_CLICK

    def clear_step_editor(self):
        self._loading_step = True
        step = self._new_step()
        self.step_name_edit.setText(step.get('name', ''))
        self.step_condition_edit.setText(step.get('condition_expr', ''))
        self.action_combo.setCurrentText(self._normalize_action_for_ui(step.get('action', ACTION_CLICK)))
        self.locator_type_combo.setCurrentText(step.get('locator_type', 'xpath'))
        self.locator_value_edit.setPlainText(step.get('locator_value', ''))
        self.target_locator_type_combo.setCurrentText(step.get('target_locator_type', 'xpath'))
        self.target_locator_value_edit.setText(step.get('target_locator_value', ''))
        self.drop_position_combo.setCurrentText(step.get('drop_position', '中间'))
        self.drag_offset_x_spin.setValue(int(step.get('drag_offset_x', 0) or 0))
        self.drag_offset_y_spin.setValue(int(step.get('drag_offset_y', 0) or 0))
        self.value_template_edit.setPlainText(step.get('value_template', ''))
        self.page_condition_expr_edit.setText(step.get('page_condition_expr', ''))
        self.detect_mode_combo.setCurrentText(step.get('detect_mode', '等待判断'))
        self.on_found_step_spin.setValue(int(step.get('on_found_step', 0) or 0))
        self.on_not_found_step_spin.setValue(int(step.get('on_not_found_step', 0) or 0))
        self.on_timeout_step_spin.setValue(int(step.get('on_timeout_step', 0) or 0))
        self.on_found_message_edit.setText(step.get('on_found_message', ''))
        self.on_not_found_message_edit.setText(step.get('on_not_found_message', ''))
        self.on_timeout_message_edit.setText(step.get('on_timeout_message', ''))
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
        self.mouse_click_button_combo.setCurrentText(step.get('mouse_click_button', '左键连击'))
        self.mouse_click_count_spin.setValue(max(1, min(20, int(step.get('mouse_click_count', 2) or 2))))
        self.keyboard_press_mode_combo.setCurrentText(step.get('keyboard_press_mode', '点按') or '点按')
        self.keyboard_hold_seconds_spin.setValue(float(step.get('keyboard_hold_seconds', 1.0) or 1.0))
        self.jump_step_spin.setValue(max(1, int(step.get('jump_step', 1) or 1)))
        self.loop_marker_edit.setText(str(step.get('loop_marker', 'i') or 'i'))
        self.loop_start_spin.setValue(max(1, min(9999, int(step.get('loop_start', 1) or 1))))
        self.loop_stop_spin.setValue(max(1, min(9999, int(step.get('loop_stop', 1) or 1))))
        self.loop_enabled_check.setChecked(bool(step.get('loop_enabled', False)))
        self.loop_content_rows_spin.setValue(max(1, min(9999, int(step.get('loop_content_rows', len(step.get('loop_values', []) or []) or 1) or 1))))
        self.loop_content_cols_spin.setValue(max(1, min(30, int(step.get('loop_content_cols', len(step.get('loop_columns', []) or []) or 1) or 1))))
        self.loop_columns_edit.setText(','.join(step.get('loop_columns', ['内容']) or ['内容']))
        self.set_loop_table_values(step.get('loop_values', []) or [])
        self.note_edit.setPlainText(step.get('note', ''))
        self._loading_step = False
        self.update_action_visibility(self.action_combo.currentText())

    def add_step(self):
        step = self._new_step(name=f'步骤{len(self.flow_config.get("steps", [])) + 1}')
        self.flow_config.setdefault('steps', []).append(step)
        self.refresh_step_list(keep_row=len(self.flow_config.get('steps', [])) - 1)
        self._save_flow_config_silent()

    def delete_step(self):
        row = self.step_list.currentRow()
        if row < 0:
            return
        del self.flow_config['steps'][row]
        next_row = row - 1 if row > 0 else 0
        self.refresh_step_list(keep_row=next_row)
        self._save_flow_config_silent()

    def copy_step(self):
        row = self.step_list.currentRow()
        steps = self.flow_config.setdefault('steps', [])
        if not (0 <= row < len(steps)):
            QMessageBox.information(self, '提示', '请先选择要复制的步骤。')
            return
        new_step = copy.deepcopy(steps[row])
        base_name = str(new_step.get('name') or new_step.get('action') or f'步骤{row + 1}').strip()
        new_step['name'] = base_name + '_复制'
        insert_row = row + 1
        steps.insert(insert_row, new_step)
        self.refresh_step_list(keep_row=insert_row)
        self._save_flow_config_silent()

    def move_step_up(self):
        row = self.step_list.currentRow()
        if row > 0:
            steps = self.flow_config['steps']
            steps[row - 1], steps[row] = steps[row], steps[row - 1]
            self.refresh_step_list(keep_row=row - 1)
            self._save_flow_config_silent()

    def move_step_down(self):
        row = self.step_list.currentRow()
        steps = self.flow_config.get('steps', [])
        if 0 <= row < len(steps) - 1:
            steps[row + 1], steps[row] = steps[row], steps[row + 1]
            self.refresh_step_list(keep_row=row + 1)
            self._save_flow_config_silent()

    def load_selected_step(self, row):
        steps = self.flow_config.get('steps', []) or []
        if not (0 <= row < len(steps)):
            self.clear_step_editor()
            return
        step = self._new_step(**steps[row])
        self._loading_step = True
        self.step_name_edit.setText(step.get('name', ''))
        self.step_condition_edit.setText(step.get('condition_expr', ''))
        self.action_combo.setCurrentText(self._normalize_action_for_ui(step.get('action', ACTION_CLICK)))
        self.locator_type_combo.setCurrentText(step.get('locator_type', 'xpath'))
        self.locator_value_edit.setPlainText(step.get('locator_value', ''))
        self.target_locator_type_combo.setCurrentText(step.get('target_locator_type', 'xpath'))
        self.target_locator_value_edit.setText(step.get('target_locator_value', ''))
        self.drop_position_combo.setCurrentText(step.get('drop_position', '中间'))
        self.drag_offset_x_spin.setValue(int(step.get('drag_offset_x', 0) or 0))
        self.drag_offset_y_spin.setValue(int(step.get('drag_offset_y', 0) or 0))
        self.value_template_edit.setPlainText(step.get('value_template', ''))
        self.page_condition_expr_edit.setText(step.get('page_condition_expr', ''))
        self.detect_mode_combo.setCurrentText(step.get('detect_mode', '等待判断'))
        self.on_found_step_spin.setValue(int(step.get('on_found_step', 0) or 0))
        self.on_not_found_step_spin.setValue(int(step.get('on_not_found_step', 0) or 0))
        self.on_timeout_step_spin.setValue(int(step.get('on_timeout_step', 0) or 0))
        self.on_found_message_edit.setText(step.get('on_found_message', ''))
        self.on_not_found_message_edit.setText(step.get('on_not_found_message', ''))
        self.on_timeout_message_edit.setText(step.get('on_timeout_message', ''))
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
        self.mouse_click_button_combo.setCurrentText(step.get('mouse_click_button', '左键连击'))
        self.mouse_click_count_spin.setValue(max(1, min(20, int(step.get('mouse_click_count', 2) or 2))))
        self.keyboard_press_mode_combo.setCurrentText(step.get('keyboard_press_mode', '点按') or '点按')
        self.keyboard_hold_seconds_spin.setValue(float(step.get('keyboard_hold_seconds', 1.0) or 1.0))
        self.jump_step_spin.setValue(max(1, int(step.get('jump_step', 1) or 1)))
        self.loop_marker_edit.setText(str(step.get('loop_marker', 'i') or 'i'))
        self.loop_start_spin.setValue(max(1, min(9999, int(step.get('loop_start', 1) or 1))))
        self.loop_stop_spin.setValue(max(1, min(9999, int(step.get('loop_stop', 1) or 1))))
        self.loop_enabled_check.setChecked(bool(step.get('loop_enabled', False)))
        self.loop_content_rows_spin.setValue(max(1, min(9999, int(step.get('loop_content_rows', len(step.get('loop_values', []) or []) or 1) or 1))))
        self.loop_content_cols_spin.setValue(max(1, min(30, int(step.get('loop_content_cols', len(step.get('loop_columns', []) or []) or 1) or 1))))
        self.loop_columns_edit.setText(','.join(step.get('loop_columns', ['内容']) or ['内容']))
        self.set_loop_table_values(step.get('loop_values', []) or [])
        self.note_edit.setPlainText(step.get('note', ''))
        self._loading_step = False
        self.update_action_visibility(self.action_combo.currentText())

    def current_step_dict(self):
        return self._new_step(
            name=self.step_name_edit.text().strip(),
            condition_expr=self.step_condition_edit.text().strip(),
            action=self.action_combo.currentText(),
            locator_type=self.locator_type_combo.currentText(),
            locator_value=self.locator_value_edit.toPlainText().strip(),
            target_locator_type=self.target_locator_type_combo.currentText(),
            target_locator_value=self.target_locator_value_edit.text().strip(),
            drop_position=self.drop_position_combo.currentText(),
            drag_offset_x=self.drag_offset_x_spin.value(),
            drag_offset_y=self.drag_offset_y_spin.value(),
            value_template=self.value_template_edit.toPlainText(),
            page_condition_expr=self.page_condition_expr_edit.text().strip(),
            detect_mode=self.detect_mode_combo.currentText(),
            on_found_step=self.on_found_step_spin.value(),
            on_not_found_step=self.on_not_found_step_spin.value(),
            on_timeout_step=self.on_timeout_step_spin.value(),
            on_found_message=self.on_found_message_edit.text().strip(),
            on_not_found_message=self.on_not_found_message_edit.text().strip(),
            on_timeout_message=self.on_timeout_message_edit.text().strip(),
            newline_mode=self.newline_mode_combo.currentText(),
            tab_mode=self.tab_mode_combo.currentText(),
            space_mode=self.space_mode_combo.currentText(),
            wait_timeout=self.wait_timeout_spin.value(),
            window_match_type=self.window_match_type_combo.currentText(),
            window_match_value=self.window_match_value_edit.text().strip(),
            sleep_seconds=self.sleep_seconds_spin.value(),
            use_js_click=self.use_js_click_check.isChecked(),
            clear_before_input=self.clear_before_input_check.isChecked(),
            wait_clickable=self.wait_clickable_check.isChecked(),
            mouse_click_button=self.mouse_click_button_combo.currentText(),
            mouse_click_count=self.mouse_click_count_spin.value(),
            keyboard_press_mode=self.keyboard_press_mode_combo.currentText(),
            keyboard_hold_seconds=self.keyboard_hold_seconds_spin.value(),
            jump_step=self.jump_step_spin.value(),
            loop_marker=self.loop_marker_edit.text().strip() or 'i',
            loop_start=self.loop_start_spin.value(),
            loop_stop=self.loop_stop_spin.value(),
            loop_enabled=self.loop_enabled_check.isChecked(),
            loop_content_rows=self.loop_content_rows_spin.value(),
            loop_content_cols=self.loop_content_cols_spin.value(),
            loop_columns=self.get_loop_columns(),
            loop_values=self.get_loop_table_values(),
            note=self.note_edit.toPlainText(),
        )

    def apply_current_step_changes(self, silent=False):
        # 兼容旧入口。当前版本已改为编辑后自动保存，不再需要手动应用。
        return self._save_current_editor_to_flow(auto=False, persist=True)

    def _save_current_editor_to_flow(self, auto=True, persist=True):
        if getattr(self, '_loading_step', False):
            return False
        row = self.step_list.currentRow()
        if row < 0:
            return False
        steps = self.flow_config.setdefault('steps', [])
        if not (0 <= row < len(steps)):
            return False
        step = self.current_step_dict()
        steps[row] = step
        item = self.step_list.item(row)
        if item is not None:
            item.setText(f"{row + 1}. {step.get('name') or step.get('action') or '未命名步骤'}")
            item.setData(Qt.UserRole, step)
        if persist:
            self._save_flow_config_silent()
        return True

    def _save_flow_config_silent(self):
        try:
            browser = self._default_browser()
            browser.update(self._collect_browser_settings(strict=False))
            self.flow_config = {'browser': browser, 'steps': [dict(step) for step in (self.flow_config.get('steps', []) or [])]}
            self.template_db.update_browser_flow(self.template_name, self.flow_config)
            self._loaded_signature = self._signature(self.flow_config)
        except Exception as e:
            self.log(f'自动保存浏览器流程失败：{e}')

    def update_action_visibility(self, action):
        locator_needed = action in (
            ACTION_CLICK, ACTION_INPUT, ACTION_MULTI_SELECT, ACTION_RIGHT_CLICK,
            ACTION_RIGHT_CLICK_MENU, ACTION_DROPDOWN_TWO_STAGE, ACTION_WAIT_ELEMENT,
            ACTION_WAIT_ELEMENT_GONE, ACTION_SWITCH_IFRAME, ACTION_DRAG,
            ACTION_PAGE_CONDITION, ACTION_ADD_TABLE, ACTION_FILL_TABLE, ACTION_KEY_COMBO,
            ACTION_KEY_PRESS, ACTION_SELECT_ALL, ACTION_CUSTOM_JS, ACTION_MOUSE_MULTI_CLICK,
        )
        value_needed = action in (ACTION_INPUT, ACTION_MULTI_SELECT, ACTION_KEY_COMBO, ACTION_KEY_PRESS, ACTION_FILL_TABLE, ACTION_CUSTOM_JS)
        target_needed = action in (ACTION_DRAG, ACTION_RIGHT_CLICK_MENU, ACTION_DROPDOWN_TWO_STAGE)
        window_needed = action == ACTION_SWITCH_WINDOW
        sleep_needed = action == ACTION_SLEEP
        drag_needed = action == ACTION_DRAG
        detect_needed = action == ACTION_PAGE_CONDITION
        mouse_click_needed = action == ACTION_MOUSE_MULTI_CLICK
        keyboard_needed = action in (ACTION_KEY_PRESS, ACTION_KEY_COMBO)
        jump_needed = action == ACTION_JUMP_STEP
        loop_end_needed = action in (ACTION_LOOP_START, ACTION_LOOP_END)
        value_loop_supported = action in (ACTION_INPUT, ACTION_MULTI_SELECT, ACTION_KEY_COMBO, ACTION_KEY_PRESS, ACTION_FILL_TABLE, ACTION_CUSTOM_JS)
        use_loop_table = value_loop_supported and self.loop_enabled_check.isChecked()

        self._set_form_row_visible(self.locator_type_combo, locator_needed)
        self._set_form_row_visible(self.locator_value_edit, locator_needed)
        self._set_form_row_visible(self.target_locator_type_combo, target_needed)
        self._set_form_row_visible(self.target_locator_value_edit, target_needed)
        self._set_form_row_visible(self.drop_position_combo, drag_needed)
        self._set_form_row_visible(self.drag_offset_widget, drag_needed)
        self._set_form_row_visible(self.value_template_edit, value_needed and not use_loop_table)
        self._set_form_row_visible(self.field_insert_widget, value_needed and not use_loop_table)
        self._set_form_row_visible(self.page_condition_expr_edit, detect_needed)
        self._set_form_row_visible(self.detect_mode_combo, detect_needed)
        self._set_form_row_visible(self.branch_found_widget, detect_needed)
        self._set_form_row_visible(self.branch_not_found_widget, detect_needed)
        self._set_form_row_visible(self.branch_timeout_widget, detect_needed)
        self._set_form_row_visible(self.newline_mode_combo, action == ACTION_INPUT)
        self._set_form_row_visible(self.tab_mode_combo, action == ACTION_INPUT)
        self._set_form_row_visible(self.space_mode_combo, action == ACTION_INPUT)
        self._set_form_row_visible(self.window_match_type_combo, window_needed)
        self._set_form_row_visible(self.window_match_value_edit, window_needed)
        self._set_form_row_visible(self.sleep_seconds_spin, sleep_needed)
        self._set_form_row_visible(self.mouse_click_widget, mouse_click_needed)
        self._set_form_row_visible(self.keyboard_action_widget, keyboard_needed)
        self._set_form_row_visible(self.jump_step_widget, jump_needed)
        self._set_form_row_visible(self.loop_end_widget, loop_end_needed)
        self._set_form_row_visible(self.loop_widget, value_loop_supported)
        self._set_form_row_visible(self.loop_value_table, use_loop_table)
        self.use_js_click_check.setVisible(action == ACTION_CLICK)
        self.clear_before_input_check.setVisible(action in (ACTION_INPUT, ACTION_FILL_TABLE))
        self.wait_clickable_check.setVisible(action in (ACTION_WAIT_ELEMENT, ACTION_WAIT_ELEMENT_GONE))
        if action == ACTION_CUSTOM_JS:
            self.value_template_edit.setPlaceholderText('JS 命令不会替换普通 {字段名}；需要字段值请使用 [[字段名]] 或 [[__RESULT__]]')
            self.insert_result_btn.setText('插入[[__RESULT__]]')
        elif action in (ACTION_KEY_PRESS, ACTION_KEY_COMBO):
            self.value_template_edit.setPlaceholderText('填写按键名，例如 ENTER、TAB、ESC、DELETE、BACKSPACE、UP、DOWN、LEFT、RIGHT、F2；也兼容 CTRL+A')
            self.insert_result_btn.setText('插入{__RESULT__}')
        else:
            self.value_template_edit.setPlaceholderText('可使用 {字段名}、{__RESULT__}、{__NL__} 或 #if(...)#')
            self.insert_result_btn.setText('插入{__RESULT__}')

    def insert_template_text(self, text):
        cursor = self.value_template_edit.textCursor()
        cursor.insertText(text)
        self.value_template_edit.setTextCursor(cursor)

    def _template_placeholder_for_current_action(self, field):
        field = str(field or '').strip()
        if self.action_combo.currentText() == ACTION_CUSTOM_JS:
            return f'[[{field}]]'
        return f'{{{field}}}'

    def insert_selected_field(self):
        field = self.field_combo.currentText().strip()
        if field:
            self.insert_template_text(self._template_placeholder_for_current_action(field))

    def insert_result_placeholder(self):
        if self.action_combo.currentText() == ACTION_CUSTOM_JS:
            self.insert_template_text('[[__RESULT__]]')
        else:
            self.insert_template_text('{__RESULT__}')


    def get_loop_columns(self):
        text = str(self.loop_columns_edit.text() or '').strip()
        columns = [item.strip() for item in text.replace('，', ',').split(',') if item.strip()]
        count = max(1, min(30, int(self.loop_content_cols_spin.value() or 1)))
        if not columns:
            columns = ['内容']
        while len(columns) < count:
            columns.append(f'列{len(columns) + 1}')
        return columns[:count]

    def get_loop_table_values(self):
        columns = self.get_loop_columns()
        values = []
        for row in range(self.loop_value_table.rowCount()):
            row_data = {}
            for col_index, name in enumerate(columns, start=1):
                item = self.loop_value_table.item(row, col_index)
                row_data[name] = item.text() if item is not None else ''
            values.append(row_data)
        return values

    def set_loop_table_values(self, values):
        values = list(values or [])
        self.loop_value_table.blockSignals(True)
        try:
            columns = self.get_loop_columns()
            count = max(1, min(9999, int(self.loop_content_rows_spin.value() or 1))) if self.loop_enabled_check.isChecked() else max(len(values), 0)
            self.loop_value_table.setColumnCount(len(columns) + 1)
            self.loop_value_table.setHorizontalHeaderLabels(['序号'] + columns)
            self.loop_value_table.setRowCount(count)
            self.loop_value_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            for col in range(1, len(columns) + 1):
                self.loop_value_table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Stretch)
            for row in range(count):
                idx_item = QTableWidgetItem(str(row + 1))
                idx_item.setFlags(idx_item.flags() & ~Qt.ItemIsEditable)
                self.loop_value_table.setItem(row, 0, idx_item)
                src = values[row] if row < len(values) else {}
                if not isinstance(src, dict):
                    src = {columns[0]: str(src)}
                for col_index, name in enumerate(columns, start=1):
                    self.loop_value_table.setItem(row, col_index, QTableWidgetItem(str(src.get(name, ''))))
        finally:
            self.loop_value_table.blockSignals(False)

    def sync_loop_value_table_rows(self):
        old_values = self.get_loop_table_values()
        self.set_loop_table_values(old_values)
        self._save_current_editor_to_flow(auto=True, persist=True)

    def _connect_auto_save_signals(self):
        widgets_text = [
            self.step_name_edit, self.step_condition_edit, self.target_locator_value_edit,
            self.page_condition_expr_edit, self.window_match_value_edit,
        ]
        widgets_text.append(self.loop_marker_edit)
        for widget in widgets_text:
            widget.textChanged.connect(lambda *args: self._save_current_editor_to_flow(auto=True, persist=True))
        self.locator_value_edit.textChanged.connect(lambda: self._save_current_editor_to_flow(auto=True, persist=True))
        self.value_template_edit.textChanged.connect(lambda: self._save_current_editor_to_flow(auto=True, persist=True))
        self.note_edit.textChanged.connect(lambda: self._save_current_editor_to_flow(auto=True, persist=True))
        for combo in (
            self.action_combo, self.locator_type_combo, self.target_locator_type_combo,
            self.drop_position_combo, self.detect_mode_combo, self.newline_mode_combo,
            self.tab_mode_combo, self.space_mode_combo, self.window_match_type_combo,
            self.mouse_click_button_combo, self.keyboard_press_mode_combo,
        ):
            combo.currentTextChanged.connect(lambda *args: self._save_current_editor_to_flow(auto=True, persist=True))
        for spin in (
            self.drag_offset_x_spin, self.drag_offset_y_spin, self.wait_timeout_spin,
            self.sleep_seconds_spin, self.on_found_step_spin, self.on_not_found_step_spin,
            self.on_timeout_step_spin, self.mouse_click_count_spin,
            self.keyboard_hold_seconds_spin, self.jump_step_spin, self.loop_start_spin, self.loop_stop_spin, self.loop_content_rows_spin, self.loop_content_cols_spin,
        ):
            spin.valueChanged.connect(lambda *args: self._save_current_editor_to_flow(auto=True, persist=True))
        for line in (self.on_found_message_edit, self.on_not_found_message_edit, self.on_timeout_message_edit):
            line.textChanged.connect(lambda *args: self._save_current_editor_to_flow(auto=True, persist=True))
        for check in (self.use_js_click_check, self.clear_before_input_check, self.wait_clickable_check):
            check.stateChanged.connect(lambda *args: self._save_current_editor_to_flow(auto=True, persist=True))
        self.loop_enabled_check.stateChanged.connect(lambda *args: (self.update_action_visibility(self.action_combo.currentText()), self.sync_loop_value_table_rows()))
        self.loop_content_rows_spin.valueChanged.connect(lambda *args: self.sync_loop_value_table_rows())
        self.loop_content_cols_spin.valueChanged.connect(lambda *args: self.sync_loop_value_table_rows())
        self.loop_columns_edit.textChanged.connect(lambda *args: self.sync_loop_value_table_rows())
        self.loop_value_table.itemChanged.connect(lambda *args: self._save_current_editor_to_flow(auto=True, persist=True))
        for edit in (self.chromedriver_edit, self.chrome_binary_edit, self.debug_port_edit, self.start_url_edit, self.implicit_wait_edit):
            edit.textChanged.connect(lambda *args: self._save_flow_config_silent())

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
            with signal_blocked(self.window_combo):
                self.window_combo.clear()
            self.log('浏览器未连接，无法刷新窗口列表。请先启动或连接浏览器。')
            return

        selected = self.window_combo.currentData() or {}
        selected_handle = selected.get('handle') or getattr(self.engine, 'preferred_window_handle', None)
        try:
            windows = self.engine.list_windows()
            target_index = -1
            with signal_blocked(self.window_combo):
                self.window_combo.clear()
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
        step = self._new_step(
            name='切换到目标窗口',
            action=ACTION_SWITCH_WINDOW,
            window_match_type=match_type,
            window_match_value=match_value,
            note='由当前选中窗口自动生成',
        )
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
    def _add_locator_candidate(candidates, seen, locator_type, locator_value, label=''):
        locator_type = str(locator_type or '').strip()
        locator_value = str(locator_value or '').strip()
        if not locator_type or not locator_value:
            return
        key = (locator_type.lower(), locator_value)
        if key in seen:
            return
        seen.add(key)
        candidates.append({'type': locator_type, 'value': locator_value, 'label': str(label or '').strip()})

    @staticmethod
    def get_locator_candidates(element):
        candidates = []
        seen = set()
        if not element:
            return candidates

        raw_candidates = element.get('locator_candidates') or []
        if isinstance(raw_candidates, list):
            for item in raw_candidates:
                if not isinstance(item, dict):
                    continue
                BrowserFlowWindow._add_locator_candidate(
                    candidates,
                    seen,
                    item.get('type'),
                    item.get('value'),
                    item.get('label'),
                )

        recommended_type = element.get('recommended_locator_type')
        recommended_value = element.get('recommended_locator_value')
        if str(recommended_type or '').strip().lower() != 'id':
            BrowserFlowWindow._add_locator_candidate(candidates, seen, recommended_type, recommended_value, '历史推荐定位')

        BrowserFlowWindow._add_locator_candidate(candidates, seen, 'name', element.get('name'), 'name 属性，通常比动态 id 稳定')
        placeholder = str(element.get('placeholder') or '').strip()
        if placeholder and "'" not in placeholder:
            BrowserFlowWindow._add_locator_candidate(candidates, seen, 'xpath', f"//*[@placeholder='{placeholder}']", 'placeholder XPath')
        title = str(element.get('title') or '').strip()
        if title and "'" not in title:
            BrowserFlowWindow._add_locator_candidate(candidates, seen, 'xpath', f"//*[@title='{title}']", 'title XPath')
            BrowserFlowWindow._add_locator_candidate(candidates, seen, 'xpath', f"//*[@aria-label='{title}']", 'aria-label XPath')
        BrowserFlowWindow._add_locator_candidate(candidates, seen, 'xpath', element.get('clickable_xpath'), '可点击元素绝对 XPath')
        BrowserFlowWindow._add_locator_candidate(candidates, seen, 'xpath', element.get('xpath'), '当前元素绝对 XPath')
        BrowserFlowWindow._add_locator_candidate(candidates, seen, 'css selector', element.get('clickable_css'), '可点击元素 CSS')
        BrowserFlowWindow._add_locator_candidate(candidates, seen, 'css selector', element.get('css'), '当前元素 CSS')
        BrowserFlowWindow._add_locator_candidate(candidates, seen, 'id', element.get('id'), 'id 属性，可能随刷新变化，建议确认后再用')
        return candidates

    @staticmethod
    def format_recorded_element_info(element):
        if not element:
            return '暂无录制结果。'
        if element.get('error'):
            return f"录制失败：{element.get('error')}"
        candidates = BrowserFlowWindow.get_locator_candidates(element)
        locator_type, locator_value = BrowserFlowWindow.choose_best_locator(element)
        lines = [
            f"标签: {element.get('tag', '')}",
            f"文本: {element.get('text', '')}",
            f"id: {element.get('id', '')}",
            f"name: {element.get('name', '')}",
            f"placeholder: {element.get('placeholder', '')}",
            f"title/aria-label: {element.get('title', '')}",
            f"自动填入定位: {locator_type} = {locator_value}",
            '可用定位方式（按推荐顺序；id 已靠后，避免动态 id 优先）：',
        ]
        if candidates:
            for idx, item in enumerate(candidates, start=1):
                label = item.get('label') or ''
                suffix = f'  # {label}' if label else ''
                lines.append(f"  {idx}. {item.get('type', '')} = {item.get('value', '')}{suffix}")
        else:
            lines.append('  （未生成可用定位方式）')
        lines.extend([
            f"XPath: {element.get('xpath', '')}",
            f"可点击XPath: {element.get('clickable_xpath', '')}",
            f"CSS: {element.get('css', '')}",
        ])
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
        candidates = BrowserFlowWindow.get_locator_candidates(element)
        for item in candidates:
            if str(item.get('type', '')).strip().lower() != 'id':
                return item.get('type', 'xpath'), item.get('value', '')
        if candidates:
            item = candidates[0]
            return item.get('type', 'xpath'), item.get('value', '')
        return 'xpath', ''
    def apply_selected_element_to_step(self, *args, **kwargs):
        element = self.selected_element_info()
        if not element:
            QMessageBox.information(self, '提示', '请先完成一次自动录制。')
            return
        locator_type, locator_value = self.choose_best_locator(element)
        action = self.action_combo.currentText()
        target_mode = action == ACTION_DRAG and bool(self.locator_value_edit.toPlainText().strip()) and not bool(self.target_locator_value_edit.text().strip())
        if action not in (ACTION_CLICK, ACTION_INPUT, ACTION_WAIT_ELEMENT, ACTION_SWITCH_IFRAME, ACTION_DRAG):
            action = ACTION_INPUT if element.get('tag') in ('input', 'textarea', 'select') else ACTION_CLICK
            self.action_combo.setCurrentText(action)
        if target_mode:
            self.target_locator_type_combo.setCurrentText(locator_type)
            self.target_locator_value_edit.setText(locator_value)
            self.log(f'已将录制结果填入拖拽目标：{locator_type} = {locator_value}')
            return
        self.locator_type_combo.setCurrentText(locator_type)
        self.locator_value_edit.setPlainText(locator_value)
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
        step = self._new_step(
            name=f'{default_action}-{desc}',
            action=default_action,
            locator_type=locator_type,
            locator_value=locator_value,
            value_template='{__RESULT__}' if default_action == ACTION_INPUT else '',
            note='由自动录制生成',
        )
        self.flow_config.setdefault('steps', []).append(step)
        self.refresh_step_list()
        self.step_list.setCurrentRow(self.step_list.count() - 1)

    def _show_flow_alert(self, message, level='info'):
        text = str(message or '').strip()
        if not text:
            return
        if str(level).lower() in ('warning', 'timeout', 'error'):
            QMessageBox.warning(self, '流程提示', text)
        else:
            QMessageBox.information(self, '流程提示', text)

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
            self.engine.execute_flow(flow, payload, logger=self.log, alert_handler=self._show_flow_alert)
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
