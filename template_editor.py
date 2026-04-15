import copy
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QListWidget, QTextEdit, QMessageBox, QInputDialog,
    QLabel, QGroupBox, QAbstractItemView, QDialog,
    QDialogButtonBox
)
from template_db import TemplateDB
from dbutils import Database


class FieldEditDialog(QDialog):
    def __init__(self, field_name, field_config, parent=None):
        super().__init__(parent)
        self.field_name = field_name
        self.original_config = field_config if isinstance(field_config, str) else ""
        self.setWindowTitle(f"编辑字段 - {field_name}")
        self.resize(560, 320)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"字段名: {self.field_name}"))
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(self.original_config)
        layout.addWidget(self.text_edit)
        layout.addWidget(QLabel("输入配置内容（空格/换行均保留；支持 {字段名}、{__NL__}、#if(...)#、#dbjoin(...)#、#dbrows(...)# 等公式标签；其中 dbjoin 采用行模板写法）"))
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_config(self):
        return self.text_edit.toPlainText()

    def closeEvent(self, event):
        if self.get_config() != self.original_config:
            reply = QMessageBox.question(
                self, "未保存", "是否保存更改？",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            if reply == QMessageBox.Yes:
                self.accept()
            elif reply == QMessageBox.No:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


class TemplateEditorWindow(QMainWindow):
    content_changed = pyqtSignal(str, dict)

    def __init__(self, main_db: Database, template_name: str, template_db: TemplateDB):
        super().__init__()
        self.main_db = main_db
        self.template_db = template_db
        self.template_name = template_name
        self.original_content = None
        self.available_fields = {}
        self.available_field_names = []
        self.field_conditions = {}
        self._suspend_live_emit = False

        self.init_ui()
        self.load_template_content()
        self.refresh_available_list()

    def init_ui(self):
        self.setWindowTitle(f"模板编辑 - {self.template_name}")
        self.resize(1120, 620)
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        mid_layout = QHBoxLayout()
        left_group = QGroupBox("可用字段（双击编辑内容）")
        left_inner = QVBoxLayout(left_group)
        self.available_list = QListWidget()
        self.available_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.available_list.itemDoubleClicked.connect(self.edit_available_field)
        left_inner.addWidget(self.available_list)

        mid_btn = QVBoxLayout()
        mid_btn.addStretch()
        field_group = QGroupBox("字段编辑")
        f_layout = QVBoxLayout(field_group)
        self.add_btn = QPushButton("添加")
        self.add_btn.clicked.connect(self.add_available_field)
        f_layout.addWidget(self.add_btn)
        self.del_btn = QPushButton("删除")
        self.del_btn.clicked.connect(self.delete_available_field)
        f_layout.addWidget(self.del_btn)
        self.rename_btn = QPushButton("重命名")
        self.rename_btn.clicked.connect(self.rename_available_field)
        f_layout.addWidget(self.rename_btn)
        mid_btn.addWidget(field_group)
        mid_btn.addSpacing(10)
        self.add_sel_btn = QPushButton("→")
        self.add_sel_btn.clicked.connect(self.add_fields_to_selected)
        mid_btn.addWidget(self.add_sel_btn)
        self.rem_sel_btn = QPushButton("←")
        self.rem_sel_btn.clicked.connect(self.remove_fields_from_selected)
        mid_btn.addWidget(self.rem_sel_btn)
        mid_btn.addStretch()

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_group = QGroupBox("已选字段")
        r_inner = QVBoxLayout(right_group)
        self.selected_list = QListWidget()
        self.selected_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.selected_list.currentItemChanged.connect(self.update_condition_summary)
        r_inner.addWidget(self.selected_list)
        cond_btn_layout = QHBoxLayout()
        self.set_cond_btn = QPushButton("显示条件")
        self.set_cond_btn.clicked.connect(self.edit_selected_field_condition)
        cond_btn_layout.addWidget(self.set_cond_btn)
        self.clear_cond_btn = QPushButton("清除条件")
        self.clear_cond_btn.clicked.connect(self.clear_selected_field_condition)
        cond_btn_layout.addWidget(self.clear_cond_btn)
        r_inner.addLayout(cond_btn_layout)
        self.condition_hint_label = QLabel("当前字段条件：无")
        self.condition_hint_label.setWordWrap(True)
        r_inner.addWidget(self.condition_hint_label)
        right_layout.addWidget(right_group)
        move_layout = QHBoxLayout()
        move_layout.addStretch()
        self.up_btn = QPushButton("↑")
        self.up_btn.clicked.connect(self.move_up)
        move_layout.addWidget(self.up_btn)
        self.down_btn = QPushButton("↓")
        self.down_btn.clicked.connect(self.move_down)
        move_layout.addWidget(self.down_btn)
        move_layout.addStretch()
        right_layout.addLayout(move_layout)

        preview_group = QGroupBox("模板预览（只读）")
        p_layout = QVBoxLayout(preview_group)
        self.preview_text = QTextEdit()
        self.preview_text.setLineWrapMode(QTextEdit.WidgetWidth)
        self.preview_text.setReadOnly(True)
        self.preview_text.setPlaceholderText("这里实时显示已选字段按当前顺序拼接后的原始模板内容。")
        p_layout.addWidget(self.preview_text)
        p_layout.addWidget(QLabel("说明：这里显示的是未替换占位符/公式标签的模板原始内容；真正替换后的结果请看主界面大文本框。"))

        mid_layout.addWidget(left_group, 3)
        mid_layout.addLayout(mid_btn)
        mid_layout.addWidget(right_widget, 3)
        mid_layout.addWidget(preview_group, 4)
        main_layout.addLayout(mid_layout)

        bottom = QHBoxLayout()
        bottom.addStretch()
        self.save_btn = QPushButton("保存模板")
        self.save_btn.clicked.connect(self.save_template)
        bottom.addWidget(self.save_btn)
        main_layout.addLayout(bottom)

    def _selected_field_name(self):
        item = self.selected_list.currentItem()
        return item.text().strip() if item else ''

    def _computed_preview_text(self):
        fields = [self.selected_list.item(i).text() for i in range(self.selected_list.count())]
        return ''.join(str(self.available_fields.get(name, '')) for name in fields)

    def update_preview(self):
        old = self.preview_text.blockSignals(True)
        self.preview_text.setPlainText(self._computed_preview_text())
        self.preview_text.blockSignals(old)

    def update_condition_summary(self, *args):
        name = self._selected_field_name()
        if not name:
            self.condition_hint_label.setText("当前字段条件：无")
            return
        expr = (self.field_conditions.get(name) or '').strip()
        self.condition_hint_label.setText(f"当前字段条件：{expr or '无'}")

    def edit_selected_field_condition(self):
        name = self._selected_field_name()
        if not name:
            QMessageBox.information(self, "提示", "请先在已选字段中选择一个字段。")
            return
        current = self.field_conditions.get(name, '')
        prompt = (
            f"为字段“{name}”设置显示条件。\n"
            "留空表示始终显示。\n"
            "示例：\n"
            "{是否二段硫化} == '是'\n"
            "{牌号} == '5470'\n"
            "if({牌号}=='1145', False, True)"
        )
        text, ok = QInputDialog.getMultiLineText(self, "字段显示条件", prompt, current)
        if ok:
            expr = text.strip()
            if expr:
                self.field_conditions[name] = expr
            else:
                self.field_conditions.pop(name, None)
            self.update_condition_summary()
            self._emit_live_content_changed()

    def clear_selected_field_condition(self):
        name = self._selected_field_name()
        if not name:
            return
        self.field_conditions.pop(name, None)
        self.update_condition_summary()
        self._emit_live_content_changed()

    def load_template_content(self):
        self._suspend_live_emit = True
        try:
            data = self.template_db.get_process_template(self.template_name)
            if data:
                content = data.get('content', {}) or {}
            else:
                content = {
                    "available_fields": {},
                    "available_field_names": [],
                    "selected_fields": [],
                    "preview_format": "",
                    "field_conditions": {},
                }
            self.available_fields = copy.deepcopy(content.get('available_fields', {}))
            self.available_field_names = list(content.get('available_field_names', []))
            self.field_conditions = copy.deepcopy(content.get('field_conditions', {}))
            self.selected_list.clear()
            self.selected_list.addItems(content.get('selected_fields', []))
            self.refresh_available_list()
            self.update_preview()
            self.update_condition_summary()
            self.original_content = self.get_current_content()
        finally:
            self._suspend_live_emit = False
        self._emit_live_content_changed()

    def refresh_available_list(self):
        self.available_list.clear()
        self.available_list.addItems(self.available_field_names)

    def add_available_field(self):
        name, ok = QInputDialog.getText(self, "添加字段", "字段名:")
        if ok and name.strip():
            name = name.strip()
            if name in self.available_fields:
                QMessageBox.warning(self, "错误", "字段已存在")
                return
            self.available_fields[name] = ""
            self.available_field_names.append(name)
            self.refresh_available_list()
            self.update_preview()
            self._emit_live_content_changed()

    def delete_available_field(self):
        selected_names = [item.text() for item in self.available_list.selectedItems()]
        if not selected_names:
            return
        for name in selected_names:
            self.available_fields.pop(name, None)
            if name in self.available_field_names:
                self.available_field_names.remove(name)
            self.field_conditions.pop(name, None)
        for i in range(self.selected_list.count() - 1, -1, -1):
            if self.selected_list.item(i).text() in selected_names:
                self.selected_list.takeItem(i)
        self.refresh_available_list()
        self.update_preview()
        self.update_condition_summary()
        self._emit_live_content_changed()

    def rename_available_field(self):
        item = self.available_list.currentItem()
        if not item:
            return
        old = item.text()
        new, ok = QInputDialog.getText(self, "重命名", "新名称:", text=old)
        if ok and new.strip() and new != old:
            new = new.strip()
            if new in self.available_fields:
                QMessageBox.warning(self, "错误", "字段已存在")
                return
            self.available_fields[new] = self.available_fields.pop(old)
            idx = self.available_field_names.index(old)
            self.available_field_names[idx] = new
            if old in self.field_conditions:
                self.field_conditions[new] = self.field_conditions.pop(old)
            self.refresh_available_list()
            for i in range(self.selected_list.count()):
                if self.selected_list.item(i).text() == old:
                    self.selected_list.item(i).setText(new)

            old_token = '{' + old + '}'
            new_token = '{' + new + '}'
            for key, value in list(self.available_fields.items()):
                if isinstance(value, str) and old_token in value:
                    self.available_fields[key] = value.replace(old_token, new_token)
            for key, value in list(self.field_conditions.items()):
                if isinstance(value, str) and old_token in value:
                    self.field_conditions[key] = value.replace(old_token, new_token)
            self.update_preview()
            self.update_condition_summary()
            self._emit_live_content_changed()

    def edit_available_field(self, item):
        name = item.text()
        dlg = FieldEditDialog(name, self.available_fields.get(name, ""), self)
        if dlg.exec_() == QDialog.Accepted:
            self.available_fields[name] = dlg.get_config()
            self.update_preview()
            self._emit_live_content_changed()

    def add_fields_to_selected(self):
        for item in self.available_list.selectedItems():
            self.selected_list.addItem(item.text())
        self.update_preview()
        self._emit_live_content_changed()

    def remove_fields_from_selected(self):
        selected_rows = sorted({self.selected_list.row(item) for item in self.selected_list.selectedItems()}, reverse=True)
        for row in selected_rows:
            self.selected_list.takeItem(row)
        self.update_preview()
        self.update_condition_summary()
        self._emit_live_content_changed()

    def move_up(self):
        row = self.selected_list.currentRow()
        if row > 0:
            item = self.selected_list.takeItem(row)
            self.selected_list.insertItem(row - 1, item)
            self.selected_list.setCurrentRow(row - 1)
            self.update_preview()
            self._emit_live_content_changed()

    def move_down(self):
        row = self.selected_list.currentRow()
        if 0 <= row < self.selected_list.count() - 1:
            item = self.selected_list.takeItem(row)
            self.selected_list.insertItem(row + 1, item)
            self.selected_list.setCurrentRow(row + 1)
            self.update_preview()
            self._emit_live_content_changed()

    def get_current_content(self):
        return {
            "available_fields": copy.deepcopy(self.available_fields),
            "available_field_names": list(self.available_field_names),
            "selected_fields": [self.selected_list.item(i).text() for i in range(self.selected_list.count())],
            "preview_format": self._computed_preview_text(),
            "field_conditions": copy.deepcopy(self.field_conditions),
        }

    def _emit_live_content_changed(self):
        if self._suspend_live_emit:
            return
        try:
            self.update_preview()
            self.content_changed.emit(self.template_name, self.get_current_content())
        except Exception:
            pass

    def save_template(self):
        content = self.get_current_content()
        self.template_db.update_process_template(self.template_name, content)
        self.original_content = content
        self._emit_live_content_changed()
        QMessageBox.information(self, "成功", "模板已保存")

    def closeEvent(self, event):
        if self.original_content and self.get_current_content() != self.original_content:
            reply = QMessageBox.question(
                self, "未保存", "保存更改？",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            if reply == QMessageBox.Yes:
                self.save_template()
                event.accept()
            elif reply == QMessageBox.No:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
