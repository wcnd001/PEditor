import csv
import io
import random
import time
from typing import List, Tuple

try:
    from charset_normalizer import from_bytes as charset_from_bytes
except Exception:
    charset_from_bytes = None

try:
    import chardet
except Exception:
    chardet = None
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QPushButton,
    QTableView, QMessageBox, QFileDialog, QHeaderView, QAbstractItemView,
    QStyledItemDelegate, QPlainTextEdit, QLineEdit, QLabel, QMenu, QAction,
    QInputDialog, QDialog, QFormLayout, QCheckBox, QDialogButtonBox, QGroupBox
)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QEvent, pyqtSignal
from PyQt5.QtGui import QFont
from dbutils import Database


class SqlTableDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QPlainTextEdit(parent)
        editor.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        editor.installEventFilter(self)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        if value is not None:
            editor.setPlainText(str(value))
        else:
            editor.setPlainText("")

    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText())

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

    def eventFilter(self, obj, event):
        if isinstance(obj, QPlainTextEdit) and event.type() == QEvent.KeyPress:
            key_event = event
            if key_event.key() == Qt.Key_Return and key_event.modifiers() == Qt.AltModifier:
                obj.insertPlainText('\n')
                return True
            elif key_event.key() == Qt.Key_Return and key_event.modifiers() == Qt.NoModifier:
                self.commitData.emit(obj)
                self.closeEditor.emit(obj)
                return True
        return super().eventFilter(obj, event)


class PandasModel(QAbstractTableModel):
    def __init__(self, data, headers, table_name, db, unique_columns=None, change_callback=None):
        super().__init__()
        self._data = data
        self._headers = headers
        self.table_name = table_name
        self.db = db
        self.unique_columns = unique_columns or set()
        self.change_callback = change_callback

    def rowCount(self, parent=QModelIndex()):
        return len(self._data)

    def columnCount(self, parent=QModelIndex()):
        return len(self._headers)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.DisplayRole or role == Qt.EditRole:
            row = self._data[index.row()]
            col = self._headers[index.column()]
            value = row.get(col, "")
            return "" if value is None else str(value)
        if role == Qt.FontRole:
            col = self._headers[index.column()]
            if col in self.unique_columns:
                font = QFont()
                font.setBold(True)
                return font
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._headers[section]
        if orientation == Qt.Horizontal and role == Qt.EditRole:
            return self._headers[section]
        return None

    def setHeaderData(self, section, orientation, value, role=Qt.EditRole):
        if orientation == Qt.Horizontal and role == Qt.EditRole:
            old_name = self._headers[section]
            new_name = value.strip()
            if not new_name or new_name == old_name:
                return False
            if new_name in self._headers:
                QMessageBox.warning(None, "错误", "列名已存在")
                return False
            try:
                self.db.execute(f'ALTER TABLE "{self.table_name}" RENAME COLUMN "{old_name}" TO "{new_name}"')
                self._headers[section] = new_name
                for row in self._data:
                    if old_name in row:
                        row[new_name] = row.pop(old_name)
                self.headerDataChanged.emit(orientation, section, section)
                if callable(self.change_callback):
                    self.change_callback()
                return True
            except Exception as e:
                QMessageBox.critical(None, "错误", f"重命名列失败：{str(e)}")
                return False
        return False

    def flags(self, index):
        if self._headers[index.column()] == 'id':
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def setData(self, index, value, role=Qt.EditRole):
        if role == Qt.EditRole:
            row = self._data[index.row()]
            col = self._headers[index.column()]
            new_value = value if value != "" else None
            row[col] = new_value
            try:
                record_id = row.get('id')
                if record_id is not None:
                    sql = f'UPDATE "{self.table_name}" SET "{col}" = ? WHERE id = ?'
                    self.db.execute(sql, (new_value, record_id))
            except Exception as e:
                QMessageBox.critical(None, "错误", f"更新数据库失败：{str(e)}")
                return False
            self.dataChanged.emit(index, index)
            if callable(self.change_callback):
                self.change_callback()
            return True
        return False

    def sort(self, column, order=Qt.AscendingOrder):
        if not self._data:
            return
        col_name = self._headers[column]
        reverse = (order == Qt.DescendingOrder)

        def sort_key(item):
            val = item.get(col_name)
            if val is None:
                return (0, "") if reverse else (1, "")
            return (1, str(val)) if reverse else (0, str(val))

        self._data.sort(key=sort_key, reverse=reverse)
        self.layoutChanged.emit()


class DataManagerWindow(QMainWindow):
    data_changed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.db = Database()
        self.current_table = None
        self.model = None
        self._header_editor = None
        self.init_ui()
        self.refresh_table_list()
        self.load_table_data()

    def _notify_data_changed(self):
        self.data_changed.emit()

    def _read_csv_rows_with_fallback(self, path):
        encodings = ['utf-8-sig', 'utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'gb18030', 'gbk', 'cp936', 'big5', 'cp1252', 'latin1']
        with open(path, 'rb') as f:
            raw = f.read()

        candidates: List[str] = []

        def add_candidate(encoding_name: str):
            encoding_name = str(encoding_name or '').strip()
            if encoding_name and encoding_name not in candidates:
                candidates.append(encoding_name)

        if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
            add_candidate('utf-16')
        if b'\x00' in raw[:128]:
            add_candidate('utf-16')

        if charset_from_bytes is not None:
            try:
                best = charset_from_bytes(raw).best()
                if best and getattr(best, 'encoding', None):
                    add_candidate(best.encoding)
            except Exception:
                pass

        if chardet is not None:
            try:
                detected = chardet.detect(raw) or {}
                add_candidate(detected.get('encoding'))
            except Exception:
                pass

        for item in encodings:
            add_candidate(item)

        last_error = None
        for encoding in candidates:
            try:
                text = raw.decode(encoding)
                sample = text[:8192]
                delimiter = ','
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
                    delimiter = dialect.delimiter
                except Exception:
                    delimiter = ','
                reader = csv.reader(io.StringIO(text), delimiter=delimiter)
                rows = list(reader)
                if rows:
                    return rows, encoding
            except UnicodeDecodeError as e:
                last_error = e
                continue
            except Exception as e:
                last_error = e
                continue
        raise UnicodeDecodeError(
            last_error.encoding if isinstance(last_error, UnicodeDecodeError) else 'unknown',
            last_error.object if isinstance(last_error, UnicodeDecodeError) else b'',
            last_error.start if isinstance(last_error, UnicodeDecodeError) else 0,
            last_error.end if isinstance(last_error, UnicodeDecodeError) else 1,
            '该 CSV 可能不是 UTF-8 编码，请先转为 UTF-8 后再导入'
        )

    def init_ui(self):
        self.setWindowTitle("数据库管理")
        self.resize(1000, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("数据表:"))

        self.table_combo = QComboBox()
        self.table_combo.currentTextChanged.connect(self.on_table_changed)
        top_layout.addWidget(self.table_combo)

        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.load_table_data)
        top_layout.addWidget(self.refresh_btn)

        self.add_row_btn = QPushButton("新增行")
        self.add_row_btn.clicked.connect(self.add_record)
        top_layout.addWidget(self.add_row_btn)

        self.delete_row_btn = QPushButton("删除行")
        self.delete_row_btn.clicked.connect(self.delete_selected)
        top_layout.addWidget(self.delete_row_btn)

        self.add_column_btn = QPushButton("新增列")
        self.add_column_btn.clicked.connect(self.add_column)
        top_layout.addWidget(self.add_column_btn)

        self.delete_column_btn = QPushButton("删除列")
        self.delete_column_btn.clicked.connect(self.delete_column)
        top_layout.addWidget(self.delete_column_btn)

        # 列移动按钮
        self.move_left_btn = QPushButton("← 左移列")
        self.move_left_btn.clicked.connect(self.move_column_left)
        top_layout.addWidget(self.move_left_btn)

        self.move_right_btn = QPushButton("右移列 →")
        self.move_right_btn.clicked.connect(self.move_column_right)
        top_layout.addWidget(self.move_right_btn)

        self.table_ops_btn = QPushButton("表操作")
        self.table_ops_btn.clicked.connect(self.show_table_ops_menu)
        top_layout.addWidget(self.table_ops_btn)

        self.import_btn = QPushButton("导入CSV")
        self.import_btn.clicked.connect(self.import_csv)
        top_layout.addWidget(self.import_btn)

        self.export_btn = QPushButton("导出CSV")
        self.export_btn.clicked.connect(self.export_csv)
        top_layout.addWidget(self.export_btn)

        layout.addLayout(top_layout)

        self.table_view = QTableView()
        self.table_view.setSortingEnabled(True)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.setWordWrap(True)
        self.table_view.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        self.table_view.setItemDelegate(SqlTableDelegate(self))
        # 双击表头编辑列名
        self.table_view.horizontalHeader().sectionDoubleClicked.connect(self.start_header_edit)
        layout.addWidget(self.table_view)

    def start_header_edit(self, logicalIndex):
        """双击表头时开始编辑列名"""
        header = self.table_view.horizontalHeader()
        # 获取表头区域矩形，需要将逻辑索引转换为视觉索引，或者直接使用 sectionViewportPosition + sectionSize
        x = header.sectionViewportPosition(logicalIndex)
        width = header.sectionSize(logicalIndex)
        height = header.height()
        rect = self.table_view.viewport().mapToGlobal(header.geometry().topLeft())
        editor = QLineEdit(header.viewport())
        editor.setText(self.model._headers[logicalIndex])
        editor.setGeometry(x, 0, width, height)
        editor.setFocus()
        editor.selectAll()
        editor.show()

        self._header_editor = editor
        self._editing_index = logicalIndex

        editor.editingFinished.connect(lambda: self.finish_header_edit(logicalIndex))
        editor.installEventFilter(self)

    def finish_header_edit(self, logicalIndex):
        editor = self._header_editor
        if editor is None:
            return
        new_name = editor.text().strip()
        editor.deleteLater()
        self._header_editor = None

        if new_name and new_name != self.model._headers[logicalIndex]:
            if self.model.setHeaderData(logicalIndex, Qt.Horizontal, new_name, Qt.EditRole):
                # 成功重命名，更新表头显示
                self.table_view.horizontalHeader().headerDataChanged(Qt.Horizontal, logicalIndex, logicalIndex)
            else:
                QMessageBox.warning(self, "错误", "列名无效或已存在")

    def eventFilter(self, obj, event):
        if hasattr(self, '_header_editor') and obj == self._header_editor:
            if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Escape:
                self._header_editor.deleteLater()
                self._header_editor = None
                return True
        return super().eventFilter(obj, event)

    def refresh_table_list(self):
        self.table_combo.blockSignals(True)
        self.table_combo.clear()
        tables = self.db.get_tables()
        self.table_combo.addItems(tables)
        self.table_combo.blockSignals(False)

        if not self.current_table and tables:
            self.table_combo.setCurrentIndex(0)
            self.current_table = tables[0]
            self.load_table_data()

    def on_table_changed(self, table_name):
        self.current_table = table_name
        self.load_table_data()

    def load_table_data(self):
        if not self.current_table:
            return
        try:
            rows = self.db.fetch_all(f'SELECT * FROM "{self.current_table}"')
            if not rows:
                info = self.db.get_table_info(self.current_table)
                headers = [col['name'] for col in info]
                rows = []
            else:
                headers = list(rows[0].keys())
            all_unique = self._get_unique_columns(self.current_table)
            business_unique = {col for col in all_unique if col != 'id'}
            self.model = PandasModel(rows, headers, self.current_table, self.db, business_unique, self._notify_data_changed)
            self.table_view.setModel(self.model)
            self.table_view.resizeRowsToContents()
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"加载表数据失败：{str(e)}")

    def _get_unique_columns(self, table_name):
        unique_cols = set()
        indexes = self.db.fetch_all(f"PRAGMA index_list({table_name})")
        for idx in indexes:
            if idx['unique']:
                info = self.db.fetch_all(f"PRAGMA index_info({idx['name']})")
                for col in info:
                    unique_cols.add(col['name'])
        table_info = self.db.get_table_info(table_name)
        for col in table_info:
            if col['pk']:
                unique_cols.add(col['name'])
        return unique_cols

    def add_record(self):
        if not self.current_table:
            QMessageBox.warning(self, "提示", "请先选择数据表")
            return
        try:
            info = self.db.get_table_info(self.current_table)
            unique_cols = self._get_unique_columns(self.current_table)
            columns = []
            values = []
            timestamp = str(int(time.time() * 1000))[-8:]
            rand_suffix = str(random.randint(1000, 9999))

            for col in info:
                if col['name'] == 'id':
                    continue
                columns.append(col['name'])
                if col['notnull'] == 1 or col['name'] in unique_cols:
                    if col['name'] in unique_cols:
                        temp_val = f"_NEW_{timestamp}_{rand_suffix}"
                        values.append(temp_val)
                    else:
                        col_type = col['type'].upper()
                        if 'INT' in col_type or 'REAL' in col_type or 'NUM' in col_type:
                            values.append(0)
                        else:
                            values.append('')
                else:
                    values.append(None)

            columns_quoted = [f'"{c}"' for c in columns]
            sql = f'INSERT INTO "{self.current_table}" ({", ".join(columns_quoted)}) VALUES ({", ".join(["?"]*len(columns))})'
            self.db.execute(sql, tuple(values))
            self.load_table_data()
            self._notify_data_changed()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"新增失败：{str(e)}")

    def delete_selected(self):
        if not self.current_table:
            return
        indexes = self.table_view.selectionModel().selectedIndexes()
        if not indexes:
            QMessageBox.information(self, "提示", "请先选择要删除的行中的任意单元格")
            return

        rows = set(idx.row() for idx in indexes)
        reply = QMessageBox.question(self, "确认删除", f"确定删除选中的 {len(rows)} 行记录吗？\n删除后剩余行的 ID 将自动重新排列。",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        try:
            model = self.table_view.model()
            has_id = 'id' in model._headers
            if not has_id:
                QMessageBox.critical(self, "错误", "该表没有 id 列，无法执行删除操作。")
                return

            ids_to_delete = []
            for row in rows:
                id_val = model.data(model.index(row, 0))
                if id_val:
                    ids_to_delete.append(int(id_val))

            for id_val in ids_to_delete:
                self.db.execute(f'DELETE FROM "{self.current_table}" WHERE id = ?', (id_val,))

            remaining = self.db.fetch_all(f'SELECT * FROM "{self.current_table}" ORDER BY id')
            for new_id, record in enumerate(remaining, start=1):
                old_id = record['id']
                if new_id != old_id:
                    self.db.execute(f'UPDATE "{self.current_table}" SET id = ? WHERE id = ?', (new_id, old_id))

            if remaining:
                max_id = len(remaining)
                self.db.execute(
                    "UPDATE sqlite_sequence SET seq = ? WHERE name = ?",
                    (max_id, self.current_table)
                )
            else:
                self.db.execute("DELETE FROM sqlite_sequence WHERE name = ?", (self.current_table,))

            self.load_table_data()
            self._notify_data_changed()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"删除失败：{str(e)}")

    # ---------- 列操作 ----------
    def add_column(self):
        if not self.current_table:
            QMessageBox.warning(self, "提示", "请先选择数据表")
            return
        existing = [self.model._headers[i] for i in range(self.model.columnCount())]
        base = "new_column"
        i = 1
        while f"{base}_{i}" in existing:
            i += 1
        col_name = f"{base}_{i}"
        try:
            self.db.execute(f'ALTER TABLE "{self.current_table}" ADD COLUMN "{col_name}" TEXT')
            self.load_table_data()
            self._notify_data_changed()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"新增列失败：{str(e)}")

    def delete_column(self):
        if not self.current_table:
            return
        indexes = self.table_view.selectionModel().selectedIndexes()
        if not indexes:
            QMessageBox.information(self, "提示", "请先选中要删除列中的任意单元格")
            return
        col = indexes[0].column()
        col_name = self.model._headers[col]
        if col_name == 'id':
            QMessageBox.warning(self, "错误", "不能删除 id 列")
            return
        reply = QMessageBox.question(self, "确认删除列", f"确定删除列 '{col_name}' 吗？\n该列的所有数据将永久丢失！",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        try:
            info = self.db.get_table_info(self.current_table)
            unique_cols = self._get_unique_columns(self.current_table)
            keep_cols = [c['name'] for c in info if c['name'] != col_name]
            temp_table = f'"{self.current_table}_temp_{int(time.time())}"'
            cols_def = []
            for col in info:
                if col['name'] == col_name:
                    continue
                definition = f'"{col["name"]}" {col["type"]}'
                if col['notnull']:
                    definition += " NOT NULL"
                if col['dflt_value'] is not None:
                    definition += f" DEFAULT {col['dflt_value']}"
                if col['pk']:
                    definition += " PRIMARY KEY"
                if col['name'] in unique_cols:
                    definition += " UNIQUE"
                cols_def.append(definition)
            create_sql = f'CREATE TABLE {temp_table} ({", ".join(cols_def)})'
            self.db.execute(create_sql)

            keep_cols_quoted = [f'"{c}"' for c in keep_cols]
            self.db.execute(f'INSERT INTO {temp_table} ({", ".join(keep_cols_quoted)}) SELECT {", ".join(keep_cols_quoted)} FROM "{self.current_table}"')

            self.db.execute(f'DROP TABLE "{self.current_table}"')
            self.db.execute(f'ALTER TABLE {temp_table} RENAME TO "{self.current_table}"')

            self.load_table_data()
            self._notify_data_changed()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"删除列失败：{str(e)}")

    def edit_header(self, index):
        self.table_view.horizontalHeader().editHeaderItem(index)

    def move_column_left(self):
        indexes = self.table_view.selectionModel().selectedIndexes()
        if not indexes:
            QMessageBox.information(self, "提示", "请先选中要移动的列中的任意单元格")
            return
        col = indexes[0].column()
        if col == 0:
            QMessageBox.information(self, "提示", "该列已在最左侧，无法左移")
            return
        self._swap_columns(col, col - 1)

    def move_column_right(self):
        indexes = self.table_view.selectionModel().selectedIndexes()
        if not indexes:
            QMessageBox.information(self, "提示", "请先选中要移动的列中的任意单元格")
            return
        col = indexes[0].column()
        if col == self.model.columnCount() - 1:
            QMessageBox.information(self, "提示", "该列已在最右侧，无法右移")
            return
        self._swap_columns(col, col + 1)

    def _swap_columns(self, col1, col2):
        """交换两列位置并更新数据库表结构"""
        headers = self.model._headers[:]
        headers[col1], headers[col2] = headers[col2], headers[col1]

        # 重建表以应用新顺序
        try:
            info = self.db.get_table_info(self.current_table)
            unique_cols = self._get_unique_columns(self.current_table)
            temp_table = f'"{self.current_table}_temp_{int(time.time())}"'
            cols_def = []
            for col_name in headers:
                col_def = next((c for c in info if c['name'] == col_name), None)
                if not col_def:
                    continue
                definition = f'"{col_name}" {col_def["type"]}'
                if col_def['notnull']:
                    definition += " NOT NULL"
                if col_def['dflt_value'] is not None:
                    definition += f" DEFAULT {col_def['dflt_value']}"
                if col_def['pk']:
                    definition += " PRIMARY KEY"
                if col_name in unique_cols:
                    definition += " UNIQUE"
                cols_def.append(definition)
            create_sql = f'CREATE TABLE {temp_table} ({", ".join(cols_def)})'
            self.db.execute(create_sql)

            new_cols_quoted = [f'"{c}"' for c in headers]
            self.db.execute(f'INSERT INTO {temp_table} ({", ".join(new_cols_quoted)}) SELECT {", ".join(new_cols_quoted)} FROM "{self.current_table}"')

            self.db.execute(f'DROP TABLE "{self.current_table}"')
            self.db.execute(f'ALTER TABLE {temp_table} RENAME TO "{self.current_table}"')

            self.load_table_data()
            # 移动后选中交换后的列
            self.table_view.selectColumn(col2)
            self._notify_data_changed()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"移动列失败：{str(e)}")

    # ---------- 表操作 ----------
    def show_table_ops_menu(self):
        menu = QMenu(self)
        new_action = QAction("新建表", self)
        new_action.triggered.connect(self.create_new_table)
        rename_action = QAction("重命名当前表", self)
        rename_action.triggered.connect(self.rename_current_table)
        delete_action = QAction("删除当前表", self)
        delete_action.triggered.connect(self.drop_current_table)
        menu.addAction(new_action)
        menu.addAction(rename_action)
        menu.addAction(delete_action)
        menu.exec_(self.table_ops_btn.mapToGlobal(self.table_ops_btn.rect().bottomLeft()))

    def create_new_table(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("新建表")
        layout = QVBoxLayout(dialog)

        form_layout = QFormLayout()
        name_edit = QLineEdit()
        form_layout.addRow("表名:", name_edit)
        layout.addLayout(form_layout)

        group_box = QGroupBox("业务唯一标识列")
        group_layout = QFormLayout(group_box)

        enable_check = QCheckBox("添加业务唯一标识列")
        enable_check.setChecked(True)
        group_layout.addRow(enable_check)

        code_name_edit = QLineEdit("code")
        code_name_edit.setEnabled(True)
        group_layout.addRow("列名:", code_name_edit)

        code_type_combo = QComboBox()
        code_type_combo.addItems(["TEXT", "INTEGER", "REAL"])
        code_type_combo.setEnabled(True)
        group_layout.addRow("数据类型:", code_type_combo)

        code_notnull_check = QCheckBox("非空 (NOT NULL)")
        code_notnull_check.setChecked(True)
        code_notnull_check.setEnabled(True)
        group_layout.addRow(code_notnull_check)

        def toggle_code_fields(checked):
            code_name_edit.setEnabled(checked)
            code_type_combo.setEnabled(checked)
            code_notnull_check.setEnabled(checked)

        enable_check.toggled.connect(toggle_code_fields)
        layout.addWidget(group_box)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec_() == QDialog.Accepted:
            table_name = name_edit.text().strip()
            if not table_name:
                QMessageBox.warning(self, "错误", "表名不能为空")
                return

            try:
                columns_def = ["id INTEGER PRIMARY KEY AUTOINCREMENT"]
                if enable_check.isChecked():
                    code_col = code_name_edit.text().strip()
                    if not code_col:
                        QMessageBox.warning(self, "错误", "唯一标识列名不能为空")
                        return
                    col_type = code_type_combo.currentText()
                    notnull = "NOT NULL" if code_notnull_check.isChecked() else ""
                    unique = "UNIQUE"
                    columns_def.append(f'"{code_col}" {col_type} {notnull} {unique}')

                create_sql = f'CREATE TABLE "{table_name}" ({", ".join(columns_def)})'
                self.db.execute(create_sql)
                self.db.conn.commit()

                self.refresh_table_list()
                self._notify_data_changed()
                QMessageBox.information(self, "成功", f"表 '{table_name}' 创建成功")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"创建表失败：{str(e)}")

    def rename_current_table(self):
        if not self.current_table:
            QMessageBox.warning(self, "提示", "请先选择数据表")
            return
        new_name, ok = QInputDialog.getText(self, "重命名表", "请输入新表名:", text=self.current_table)
        if ok and new_name.strip() and new_name != self.current_table:
            new_name = new_name.strip()
            try:
                self.db.execute(f'ALTER TABLE "{self.current_table}" RENAME TO "{new_name}"')
                self.current_table = new_name
                self.refresh_table_list()
                self.table_combo.setCurrentText(new_name)
                self._notify_data_changed()
                QMessageBox.information(self, "成功", f"表已重命名为 '{new_name}'")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"重命名失败：{str(e)}")

    def drop_current_table(self):
        if not self.current_table:
            QMessageBox.warning(self, "提示", "请先选择数据表")
            return
        reply = QMessageBox.question(self, "确认删除表",
                                     f"确定要永久删除表 '{self.current_table}' 吗？\n此操作不可恢复！",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                self.db.execute(f'DROP TABLE "{self.current_table}"')
                self.current_table = None
                self.refresh_table_list()
                self.load_table_data()
                self._notify_data_changed()
                QMessageBox.information(self, "成功", "表已删除")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"删除表失败：{str(e)}")

    # ---------- 导入导出 ----------
    def import_csv(self):
        if not self.current_table:
            QMessageBox.warning(self, "提示", "请先选择数据表")
            return
        path, _ = QFileDialog.getOpenFileName(self, "选择CSV文件", "", "CSV文件 (*.csv)")
        if not path:
            return

        try:
            rows, used_encoding = self._read_csv_rows_with_fallback(path)

            if not rows:
                QMessageBox.warning(self, "提示", "CSV文件为空")
                return

            headers = rows[0]
            data_rows = rows[1:]

            has_id = 'id' in headers
            ignore_id = False
            if has_id:
                reply = QMessageBox.question(self, "ID列处理",
                    "CSV文件中包含 'id' 列。\n"
                    "选择 'Yes'：忽略 id 列，让数据库自动生成。\n"
                    "选择 'No' ：使用 CSV 中的 id 值（可能与已有记录冲突）。",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                ignore_id = (reply == QMessageBox.Yes)

            table_info = self.db.get_table_info(self.current_table)
            not_null_columns = [col['name'] for col in table_info if col['notnull'] == 1 and col['dflt_value'] is None]

            if has_id and ignore_id:
                target_headers = [h for h in headers if h != 'id']
            else:
                target_headers = headers

            not_null_indices = [i for i, h in enumerate(target_headers) if h in not_null_columns]

            insert_data = []
            skipped_empty = 0
            skipped_not_null = 0
            for row in data_rows:
                if len(row) < len(headers):
                    row.extend([''] * (len(headers) - len(row)))
                elif len(row) > len(headers):
                    row = row[:len(headers)]

                if has_id and ignore_id:
                    id_index = headers.index('id')
                    row_values = [v for i, v in enumerate(row) if i != id_index]
                else:
                    row_values = row

                converted = [None if (v is None or str(v).strip() == '') else str(v).strip() for v in row_values]

                if all(v is None for v in converted):
                    skipped_empty += 1
                    continue

                invalid = False
                for idx in not_null_indices:
                    if converted[idx] is None:
                        invalid = True
                        break
                if invalid:
                    skipped_not_null += 1
                    continue

                insert_data.append(tuple(converted))

            if not insert_data:
                QMessageBox.information(self, "提示",
                    f"没有可导入的有效数据行。\n全空行：{skipped_empty}，违反非空约束：{skipped_not_null}")
                return

            confirm = QMessageBox.warning(self, "确认覆盖操作",
                f"⚠️ 此操作将清空表 “{self.current_table}” 中的所有现有数据！\n"
                f"然后导入 {len(insert_data)} 行有效数据。\n"
                f"（跳过全空行 {skipped_empty} 行，违反非空约束 {skipped_not_null} 行）\n\n"
                "确定要继续吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if confirm != QMessageBox.Yes:
                return

            try:
                self.db.begin_transaction()
                self.db.execute(f'DELETE FROM "{self.current_table}"')
                self.db.execute("DELETE FROM sqlite_sequence WHERE name = ?", (self.current_table,))

                col_quoted = [f'"{c}"' for c in target_headers]
                sql = f'INSERT INTO "{self.current_table}" ({", ".join(col_quoted)}) VALUES ({", ".join(["?"]*len(target_headers))})'
                self.db.executemany(sql, insert_data)

                self.db.commit_transaction()
                self.load_table_data()
                self._notify_data_changed()
                QMessageBox.information(self, "完成",
                    f"导入成功！\n已清空原有数据，写入 {len(insert_data)} 行新数据。\n"
                    f"（跳过全空行 {skipped_empty} 行，违反非空约束 {skipped_not_null} 行）")
            except Exception as e:
                self.db.rollback_transaction()
                raise e

        except UnicodeDecodeError:
            QMessageBox.critical(self, "错误", "导入失败：该 CSV 可能不是 UTF-8 编码，请先转为 UTF-8 后再导入。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导入失败：{str(e)}")

    def export_csv(self):
        if not self.current_table:
            return
        path, _ = QFileDialog.getSaveFileName(self, "保存CSV文件", "", "CSV文件 (*.csv)")
        if not path:
            return
        try:
            rows = self.db.fetch_all(f'SELECT * FROM "{self.current_table}"')
            if not rows:
                QMessageBox.warning(self, "提示", "表中无数据")
                return
            headers = list(rows[0].keys())
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                writer.writerows(rows)
            QMessageBox.information(self, "成功", "导出成功")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

    def closeEvent(self, event):
        event.accept()