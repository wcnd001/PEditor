import json
import math
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, ROUND_UP
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QMessageBox, QHeaderView, QAbstractItemView, QFormLayout,
    QLineEdit, QComboBox, QDialogButtonBox, QLabel
)
from dbutils import Database
from template_db import TemplateDB


class RuleEditDialog(QDialog):
    """编辑单条规则"""

    def __init__(self, rule_item=None, main_db=None, field_pool=None, parent=None):
        super().__init__(parent)
        self.main_db = main_db
        self.field_pool = [str(item).strip() for item in (field_pool or []) if str(item).strip()]
        self.rule_item = rule_item or {
            "input_field": "",
            "operation": "直接传递",
            "output_field": "",
            "params": {}
        }
        self.setWindowTitle("编辑规则")
        self.resize(560, 460)
        self.init_ui()
        self.load_rule_item()
        self.original_rule_json = json.dumps(self.get_rule_item(), ensure_ascii=False, sort_keys=True)

    def init_ui(self):
        layout = QFormLayout(self)

        self.input_combo = QComboBox()
        self.input_combo.setEditable(True)
        self.input_combo.addItems(self.field_pool)
        layout.addRow("输入字段名:", self.input_combo)

        self.op_combo = QComboBox()
        self.op_combo.addItems(["直接传递", "数学公式", "数据库查询"])
        self.op_combo.currentTextChanged.connect(self.on_op_changed)
        layout.addRow("操作类型:", self.op_combo)

        self.formula_edit = QLineEdit()
        self.formula_edit.setPlaceholderText("例如: if({牌号}=='1145','无需二段硫化','需要二段硫化')")
        layout.addRow("数学公式:", self.formula_edit)

        hint_label = QLabel(
            "语法说明：\n"
            "• 字段名必须用花括号包裹，如 {厚度}、{加料（g）}\n"
            "• 支持运算符：+ - * / // % ** 以及比较运算 == != > >= < <=\n"
            "• 支持函数：int(x)、float(x)、round(x, n)、roundup(x, n)、abs(x)、isblank(x)、nl()\n"
            "• 支持模板查库：dbjoin(表, 查询列, 查询值, 行模板, 分隔符)，dbrows(...) 为兼容别名\n"
            "• 条件写法同时支持：if(条件, 真值, 假值)、and(a,b,...)、or(a,b,...) 以及 Python 条件表达式\n"
            "• 输出字段可与现有字段同名，后面的规则输出会覆盖前面的原始值。"
        )
        hint_label.setWordWrap(True)
        hint_label.setStyleSheet("color: #555; font-size: 13px;")
        layout.addRow(hint_label)

        self.db_table_combo = QComboBox()
        self.db_key_column_combo = QComboBox()
        self.db_value_column_combo = QComboBox()
        layout.addRow("查询表:", self.db_table_combo)
        layout.addRow("查询列 (条件):", self.db_key_column_combo)
        layout.addRow("返回列:", self.db_value_column_combo)

        self.output_combo = QComboBox()
        self.output_combo.setEditable(True)
        self.output_combo.addItems(self.field_pool)
        layout.addRow("输出字段名:", self.output_combo)

        self.on_op_changed(self.op_combo.currentText())

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

        if self.main_db:
            tables = self.main_db.get_tables()
            self.db_table_combo.addItems(tables)
            self.db_table_combo.currentTextChanged.connect(self.update_columns)
            self.update_columns()

    def update_columns(self):
        table = self.db_table_combo.currentText()
        self.db_key_column_combo.clear()
        self.db_value_column_combo.clear()
        if table and self.main_db:
            info = self.main_db.get_table_info(table)
            for col in info:
                self.db_key_column_combo.addItem(col['name'])
                self.db_value_column_combo.addItem(col['name'])

    def on_op_changed(self, text):
        is_formula = text == "数学公式"
        is_db = text == "数据库查询"
        self.formula_edit.setVisible(is_formula)
        self.db_table_combo.setVisible(is_db)
        self.db_key_column_combo.setVisible(is_db)
        self.db_value_column_combo.setVisible(is_db)

    def load_rule_item(self):
        self.input_combo.setCurrentText(self.rule_item.get("input_field", ""))
        self.output_combo.setCurrentText(self.rule_item.get("output_field", ""))
        op = self.rule_item.get("operation", "直接传递")
        self.op_combo.setCurrentText(op)
        params = self.rule_item.get("params", {}) or {}

        if op == "数学公式":
            self.formula_edit.setText(params.get("formula", ""))
        elif op == "数据库查询":
            table = params.get("table", "")
            key_col = params.get("key_column", "")
            value_col = params.get("value_column", "")

            idx = self.db_table_combo.findText(table)
            if idx >= 0:
                self.db_table_combo.setCurrentIndex(idx)
            self.update_columns()

            idx = self.db_key_column_combo.findText(key_col)
            if idx >= 0:
                self.db_key_column_combo.setCurrentIndex(idx)
            idx = self.db_value_column_combo.findText(value_col)
            if idx >= 0:
                self.db_value_column_combo.setCurrentIndex(idx)

    def get_rule_item(self):
        op = self.op_combo.currentText()
        params = {}
        if op == "数学公式":
            params["formula"] = self.formula_edit.text().strip()
        elif op == "数据库查询":
            params["table"] = self.db_table_combo.currentText()
            params["key_column"] = self.db_key_column_combo.currentText()
            params["value_column"] = self.db_value_column_combo.currentText()
        return {
            "input_field": self.input_combo.currentText().strip(),
            "operation": op,
            "output_field": self.output_combo.currentText().strip(),
            "params": params
        }

    def reject(self):
        current_rule_json = json.dumps(self.get_rule_item(), ensure_ascii=False, sort_keys=True)
        if current_rule_json != self.original_rule_json:
            reply = QMessageBox.question(
                self,
                "未保存",
                "是否保存更改？",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            if reply == QMessageBox.Yes:
                self.accept()
            elif reply == QMessageBox.No:
                super().reject()
        else:
            super().reject()


class RuleManagerDialog(QDialog):
    """规则管理对话框"""

    def __init__(self, rules: list, main_db: Database, field_pool: list, parent=None):
        super().__init__(parent)
        self.main_db = main_db
        self.field_pool = [str(item).strip() for item in (field_pool or []) if str(item).strip()]
        self.rules = [json.loads(json.dumps(r, ensure_ascii=False)) for r in rules]
        self.original_rules = json.dumps(rules, ensure_ascii=False, sort_keys=True)
        self.setWindowTitle("编辑规则")
        self.resize(980, 520)
        self.init_ui()
        self.refresh_table()

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["输入字段", "操作", "输出字段", "参数"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("添加规则")
        self.add_btn.clicked.connect(self.add_rule)
        btn_layout.addWidget(self.add_btn)
        self.edit_btn = QPushButton("编辑规则")
        self.edit_btn.clicked.connect(self.edit_rule)
        btn_layout.addWidget(self.edit_btn)
        self.del_btn = QPushButton("删除规则")
        self.del_btn.clicked.connect(self.delete_rule)
        btn_layout.addWidget(self.del_btn)
        self.up_btn = QPushButton("上移")
        self.up_btn.clicked.connect(self.move_up)
        btn_layout.addWidget(self.up_btn)
        self.down_btn = QPushButton("下移")
        self.down_btn.clicked.connect(self.move_down)
        btn_layout.addWidget(self.down_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def refresh_table(self):
        self.table.setRowCount(len(self.rules))
        for i, rule in enumerate(self.rules):
            self.table.setItem(i, 0, QTableWidgetItem(rule.get('input_field', '')))
            self.table.setItem(i, 1, QTableWidgetItem(rule.get('operation', '')))
            self.table.setItem(i, 2, QTableWidgetItem(rule.get('output_field', '')))
            self.table.setItem(i, 3, QTableWidgetItem(str(rule.get('params', {}))))

    def get_dynamic_field_pool(self):
        labels = list(self.field_pool)
        for rule in self.rules:
            out = rule.get('output_field', '').strip()
            if out and out not in labels:
                labels.append(out)
        return labels

    def add_rule(self):
        dlg = RuleEditDialog(main_db=self.main_db, field_pool=self.get_dynamic_field_pool(), parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.rules.append(dlg.get_rule_item())
            self.refresh_table()

    def edit_rule(self):
        row = self.table.currentRow()
        if row < 0:
            return
        dlg = RuleEditDialog(self.rules[row], self.main_db, self.get_dynamic_field_pool(), self)
        if dlg.exec_() == QDialog.Accepted:
            self.rules[row] = dlg.get_rule_item()
            self.refresh_table()
            self.table.setCurrentCell(row, 0)

    def delete_rule(self):
        row = self.table.currentRow()
        if row >= 0:
            del self.rules[row]
            self.refresh_table()

    def move_up(self):
        row = self.table.currentRow()
        if row > 0:
            self.rules[row], self.rules[row - 1] = self.rules[row - 1], self.rules[row]
            self.refresh_table()
            self.table.setCurrentCell(row - 1, 0)

    def move_down(self):
        row = self.table.currentRow()
        if 0 <= row < self.table.rowCount() - 1:
            self.rules[row], self.rules[row + 1] = self.rules[row + 1], self.rules[row]
            self.refresh_table()
            self.table.setCurrentCell(row + 1, 0)

    def get_rules(self):
        return self.rules

    def reject(self):
        current = json.dumps(self.rules, ensure_ascii=False, sort_keys=True)
        if current != self.original_rules:
            reply = QMessageBox.question(
                self,
                "未保存",
                "是否保存更改？",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            if reply == QMessageBox.Yes:
                self.accept()
            elif reply == QMessageBox.No:
                super().reject()
        else:
            super().reject()


class DataMatcher:
    """规则引擎，供主程序调用"""

    _FULLWIDTH_TRANS = str.maketrans({
        '（': '(',
        '）': ')',
        '，': ',',
        '；': ';',
    })
    _PLACEHOLDER_PATTERN = re.compile(r'\{([^{}]+)\}')
    _FORMULA_TAG_PATTERN = re.compile(r'#([^#]+?)#', re.DOTALL)
    _SPECIAL_VALUES = {
        '__NL__': '\n',
        'NL': '\n',
        '__BR__': '\n',
        '换行': '\n',
    }

    def __init__(self, main_db: Database, template_db: TemplateDB):
        self.main_db = main_db
        self.template_db = template_db

    @staticmethod
    def _unique_preserve(items):
        result = []
        seen = set()
        for item in items:
            name = str(item).strip()
            if name and name not in seen:
                seen.add(name)
                result.append(name)
        return result

    def get_rule_field_pool(self, template_name: str = None, input_option_labels=None, main_config_override=None, process_content_override=None):
        labels = []
        labels.extend(input_option_labels or [])

        if template_name:
            main_tpl = self.template_db.get_main_template(template_name) if main_config_override is None else {'config': main_config_override}
            main_cfg = (main_tpl.get('config', {}) if main_tpl else {}) or {}
            for opt in main_cfg.get('options', []) or []:
                label = str(opt.get('label', '')).strip()
                if label:
                    labels.append(label)
            for rule in main_cfg.get('rules', []) or []:
                in_field = str(rule.get('input_field', '')).strip()
                out = str(rule.get('output_field', '')).strip()
                if in_field:
                    labels.append(in_field)
                if out:
                    labels.append(out)

            proc_tpl = self.template_db.get_process_template(template_name) if process_content_override is None else {'content': process_content_override}
            proc_cfg = (proc_tpl.get('content', {}) if proc_tpl else {}) or {}
            labels.extend(proc_cfg.get('available_field_names', []) or [])
            labels.extend((proc_cfg.get('available_fields', {}) or {}).keys())
            labels.extend(proc_cfg.get('selected_fields', []) or [])

        return self._unique_preserve(labels)

    @classmethod
    def _normalize_field_key(cls, key: str) -> str:
        return cls._translate_fullwidth_outside_strings(str(key or '').strip())

    @classmethod
    def _lookup_mapping_value(cls, mapping, key, default=""):
        if not isinstance(mapping, dict):
            return default
        if key in mapping:
            value = mapping[key]
            return default if value is None else value
        normalized_key = cls._normalize_field_key(key)
        for existing_key, value in mapping.items():
            if cls._normalize_field_key(existing_key) == normalized_key:
                return default if value is None else value
        return default

    @classmethod
    def _first_existing_value(cls, *sources, default=""):
        for mapping, key in sources:
            value = cls._lookup_mapping_value(mapping, key, default=None)
            if value is not None:
                return default if value is None else value
        return default

    @classmethod
    def _special_value(cls, field_name):
        return cls._SPECIAL_VALUES.get(str(field_name).strip(), None)

    @staticmethod
    def _convert_expression_value(value):
        if value is None:
            return ''
        if isinstance(value, str):
            stripped = value.strip()
            if stripped == '':
                return ''
            if stripped.lower() == 'true':
                return True
            if stripped.lower() == 'false':
                return False
            try:
                number = float(stripped)
                if number.is_integer():
                    return int(number)
                return number
            except ValueError:
                return stripped
        return value

    @classmethod
    def _find_matching_paren(cls, text: str, open_index: int) -> int:
        depth = 0
        quote = None
        escape = False
        for index in range(open_index, len(text)):
            ch = text[index]
            if quote:
                if escape:
                    escape = False
                elif ch == '\\':
                    escape = True
                elif ch == quote:
                    quote = None
                continue
            if ch in ('"', "'"):
                quote = ch
                continue
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
                if depth == 0:
                    return index
        raise ValueError('函数括号未正确闭合')

    @classmethod
    def _split_top_level_args(cls, text: str):
        args = []
        start = 0
        depth = 0
        quote = None
        escape = False
        for index, ch in enumerate(text):
            if quote:
                if escape:
                    escape = False
                elif ch == '\\':
                    escape = True
                elif ch == quote:
                    quote = None
                continue
            if ch in ('"', "'"):
                quote = ch
                continue
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            elif ch == ',' and depth == 0:
                args.append(text[start:index].strip())
                start = index + 1
        args.append(text[start:].strip())
        return args

    @staticmethod
    def _validate_logic_function_args(func_name: str, parts):
        if not parts:
            raise ValueError(f'{func_name} 函数至少需要 1 个参数')
        blank_indexes = [str(i + 1) for i, part in enumerate(parts) if not str(part).strip()]
        if blank_indexes:
            joined = '、'.join(blank_indexes)
            raise ValueError(f'{func_name} 函数第 {joined} 个参数不能为空')
        return parts

    @classmethod
    def _rewrite_if_functions(cls, text: str) -> str:
        result = []
        index = 0
        while index < len(text):
            match = re.match(r'(?i)if\s*\(', text[index:])
            prev_ok = index == 0 or not (text[index - 1].isalnum() or text[index - 1] == '_')
            if match and prev_ok:
                token = match.group(0)
                open_index = index + token.rfind('(')
                close_index = cls._find_matching_paren(text, open_index)
                inner = text[open_index + 1:close_index]
                parts = cls._split_top_level_args(inner)
                if len(parts) != 3:
                    raise ValueError('if 函数必须包含 3 个参数：if(条件, 真值, 假值)')
                cond = cls._rewrite_expression_functions(parts[0])
                true_expr = cls._rewrite_expression_functions(parts[1])
                false_expr = cls._rewrite_expression_functions(parts[2])
                result.append(f'(({true_expr}) if ({cond}) else ({false_expr}))')
                index = close_index + 1
                continue
            result.append(text[index])
            index += 1
        return ''.join(result)

    @classmethod
    def _rewrite_logic_function(cls, text: str, func_name: str, operator: str) -> str:
        result = []
        index = 0
        pattern = re.compile(rf'(?i){func_name}\(')
        while index < len(text):
            match = pattern.match(text[index:])
            prev_ok = index == 0 or not (text[index - 1].isalnum() or text[index - 1] == '_')
            if match and prev_ok:
                token = match.group(0)
                open_index = index + token.rfind('(')
                close_index = cls._find_matching_paren(text, open_index)
                inner = text[open_index + 1:close_index]
                raw_parts = cls._validate_logic_function_args(func_name, cls._split_top_level_args(inner))
                parts = [cls._rewrite_expression_functions(part) for part in raw_parts]
                joined = f' {operator} '.join(f'({part})' for part in parts)
                result.append(f'({joined})')
                index = close_index + 1
                continue
            result.append(text[index])
            index += 1
        return ''.join(result)

    @classmethod
    def _rewrite_expression_functions(cls, text: str) -> str:
        text = str(text or '')
        text = cls._rewrite_if_functions(text)
        text = cls._rewrite_logic_function(text, 'and', 'and')
        text = cls._rewrite_logic_function(text, 'or', 'or')
        return text

    @classmethod
    def _translate_fullwidth_outside_strings(cls, text: str) -> str:
        result = []
        quote = None
        escape = False
        for ch in str(text or ''):
            if quote:
                result.append(ch)
                if escape:
                    escape = False
                elif ch == '\\':
                    escape = True
                elif ch == quote:
                    quote = None
                continue
            if ch in ('"', "'"):
                quote = ch
                result.append(ch)
                continue
            result.append(ch.translate(cls._FULLWIDTH_TRANS))
        return ''.join(result)

    @classmethod
    def _normalize_expression_text(cls, expression: str) -> str:
        text = cls._translate_fullwidth_outside_strings(expression)
        return cls._rewrite_expression_functions(text)

    def _build_expression(self, expression: str, resolver):
        context = {}
        counter = 0
        raw_expression = str(expression or '')

        result = []
        index = 0
        quote = None
        escape = False
        while index < len(raw_expression):
            ch = raw_expression[index]
            if quote:
                result.append(ch)
                if escape:
                    escape = False
                elif ch == '\\':
                    escape = True
                elif ch == quote:
                    quote = None
                index += 1
                continue
            if ch in ('"', "'"):
                quote = ch
                result.append(ch)
                index += 1
                continue
            if ch == '{':
                close_index = raw_expression.find('}', index + 1)
                if close_index > index:
                    field = raw_expression[index + 1:close_index].strip()
                    var_name = f'__v{counter}'
                    counter += 1
                    context[var_name] = self._convert_expression_value(resolver(field))
                    result.append(var_name)
                    index = close_index + 1
                    continue
            result.append(ch)
            index += 1

        expr = self._normalize_expression_text(''.join(result))
        return expr, context

    @staticmethod
    def _to_decimal(value):
        if isinstance(value, Decimal):
            return value
        if value is None:
            return Decimal('0')
        if isinstance(value, bool):
            return Decimal(int(value))
        if isinstance(value, (int, float)):
            return Decimal(str(value))
        text = str(value).strip()
        if text == '':
            return Decimal('0')
        return Decimal(text)

    @classmethod
    def _safe_int(cls, value=0):
        return int(float(cls._convert_expression_value(value) or 0))

    @classmethod
    def _safe_float(cls, value=0):
        return float(cls._convert_expression_value(value) or 0)

    @classmethod
    def _safe_round(cls, value, ndigits=0):
        number = cls._to_decimal(value)
        ndigits = int(cls._safe_int(ndigits))
        quant = Decimal('1').scaleb(-ndigits)
        result = number.quantize(quant, rounding=ROUND_HALF_UP)
        if ndigits <= 0 and result == result.to_integral_value():
            return int(result)
        return float(result)

    @classmethod
    def _safe_roundup(cls, value, ndigits=0):
        number = cls._to_decimal(value)
        ndigits = int(cls._safe_int(ndigits))
        quant = Decimal('1').scaleb(-ndigits)
        result = number.quantize(quant, rounding=ROUND_UP)
        if ndigits <= 0 and result == result.to_integral_value():
            return int(result)
        return float(result)

    @staticmethod
    def _safe_iif(condition, true_value, false_value):
        return true_value if condition else false_value

    @staticmethod
    def _safe_isblank(value):
        if value is None:
            return True
        return str(value).strip() == ''

    @staticmethod
    def _safe_nl():
        return '\n'

    @staticmethod
    def _escape_identifier(name: str) -> str:
        return str(name or '').replace('"', '""')

    def _table_columns(self, table_name: str):
        return [col.get('name') for col in self.main_db.get_table_info(table_name)]

    def _build_db_order_sql(self, table_name: str) -> str:
        try:
            columns = self._table_columns(table_name)
        except Exception:
            columns = []
        if 'id' in columns:
            return ' ORDER BY "id"'
        return ''

    def _fetch_db_rows(self, table_name: str, key_column: str, key_value, select_columns=None):
        table_name = str(table_name or '').strip()
        key_column = str(key_column or '').strip()
        if not table_name or not key_column:
            raise ValueError('dbjoin 参数不完整：表名和查询列不能为空')

        all_columns = self._table_columns(table_name)
        if key_column not in all_columns:
            raise ValueError(f'数据表“{table_name}”中不存在查询列“{key_column}”')

        if select_columns:
            for col in select_columns:
                if col not in all_columns:
                    raise ValueError(f'数据表“{table_name}”中不存在返回列“{col}”')
            select_sql = ', '.join(f'"{self._escape_identifier(col)}"' for col in select_columns)
        else:
            select_sql = '*'

        table_sql = self._escape_identifier(table_name)
        key_sql = self._escape_identifier(key_column)
        sql = f'SELECT {select_sql} FROM "{table_sql}" WHERE "{key_sql}" = ?{self._build_db_order_sql(table_name)}'
        return self.main_db.fetch_all(sql, (key_value,))

    def _auto_eval_nested_expression(self, source_text: str, rendered_text: str, resolver):
        source_text = '' if source_text is None else str(source_text)
        rendered_text = '' if rendered_text is None else str(rendered_text)
        if not source_text or not rendered_text:
            return rendered_text
        # 仅在源文本本身包含占位符，且看起来像纯表达式时尝试二次求值，
        # 避免把普通文本（例如 147±2 或 100-120）误当成数学表达式。
        if '{' not in source_text or '}' not in source_text:
            return rendered_text
        if '#' in source_text:
            return rendered_text
        candidate = rendered_text.strip()
        if not candidate or '{' in candidate or '}' in candidate:
            return rendered_text
        if not re.search(r'[+\-*/%()]', candidate):
            return rendered_text
        try:
            value = self._evaluate_expression(candidate, resolver)
            return '' if value is None else str(value)
        except Exception:
            return rendered_text

    def _render_db_cell_value(self, raw_value, resolver):
        if raw_value is None:
            return ''
        if not isinstance(raw_value, str):
            return raw_value
        rendered = self._render_recursive_text(raw_value, resolver)
        return self._auto_eval_nested_expression(raw_value, rendered, resolver)

    def _render_db_row_template(self, row_template: str, row_data: dict, row_index: int, outer_resolver):
        row_payload = dict(row_data or {})
        row_payload.setdefault('序号', row_index)
        row_payload.setdefault('__INDEX__', row_index)

        def resolver(field_name):
            special = self._special_value(field_name)
            if special is not None:
                return special
            value = self._lookup_mapping_value(row_payload, field_name, default=None)
            if value is not None:
                return self._render_db_cell_value(value, outer_resolver)
            return outer_resolver(field_name)

        rendered = self._render_recursive_text(row_template, resolver)
        return self._auto_eval_nested_expression(row_template, rendered, resolver)

    def _dbjoin(self, resolver, table_name, key_column, key_value, row_template, separator=''):
        table_name = '' if table_name is None else str(table_name)
        key_column = '' if key_column is None else str(key_column)
        row_template = '' if row_template is None else str(row_template)
        sep = '' if separator is None else str(separator)
        rows = self._fetch_db_rows(table_name, key_column, key_value)
        rendered_rows = []
        for idx, row in enumerate(rows, start=1):
            rendered_rows.append(self._render_db_row_template(row_template, row, idx, resolver))
        return sep.join(rendered_rows)

    def _dbrows(self, resolver, table_name, key_column, key_value, row_template, separator='\n'):
        return self._dbjoin(resolver, table_name, key_column, key_value, row_template, separator)

    def _safe_globals(self, resolver=None):
        resolver = resolver or (lambda field_name: '')
        return {
            '__builtins__': {},
            'int': self._safe_int,
            'INT': self._safe_int,
            'float': self._safe_float,
            'FLOAT': self._safe_float,
            'round': self._safe_round,
            'ROUND': self._safe_round,
            'roundup': self._safe_roundup,
            'ROUNDUP': self._safe_roundup,
            'iif': self._safe_iif,
            'IIF': self._safe_iif,
            'abs': abs,
            'ABS': abs,
            'max': max,
            'MAX': max,
            'min': min,
            'MIN': min,
            'ceil': math.ceil,
            'CEIL': math.ceil,
            'floor': math.floor,
            'FLOOR': math.floor,
            'str': str,
            'STR': str,
            'len': len,
            'LEN': len,
            'isblank': self._safe_isblank,
            'ISBLANK': self._safe_isblank,
            'nl': self._safe_nl,
            'NL': self._safe_nl,
            'newline': self._safe_nl,
            'NEWLINE': self._safe_nl,
            'dbjoin': lambda table_name, key_column, key_value, row_template, separator='': self._dbjoin(resolver, table_name, key_column, key_value, row_template, separator),
            'DBJOIN': lambda table_name, key_column, key_value, row_template, separator='': self._dbjoin(resolver, table_name, key_column, key_value, row_template, separator),
            'dbrows': lambda table_name, key_column, key_value, row_template, separator='\n': self._dbrows(resolver, table_name, key_column, key_value, row_template, separator),
            'DBROWS': lambda table_name, key_column, key_value, row_template, separator='\n': self._dbrows(resolver, table_name, key_column, key_value, row_template, separator),
            'True': True,
            'False': False,
            'true': True,
            'false': False,
            'TRUE': True,
            'FALSE': False,
        }

    def _evaluate_expression(self, expression: str, resolver):
        expr, context = self._build_expression(expression, resolver)
        try:
            return eval(expr, self._safe_globals(resolver), context)
        except NameError as e:
            raise ValueError(f'存在未加花括号的字段名或非法名称: {e}')
        except SyntaxError as e:
            raise ValueError(f'表达式语法错误: {e}')
        except InvalidOperation as e:
            raise ValueError(f'数值处理错误: {e}')
        except Exception as e:
            raise ValueError(str(e))

    @staticmethod
    def _coerce_condition_result(value):
        if isinstance(value, bool):
            return value
        if value is None:
            return False
        if isinstance(value, (int, float, Decimal)):
            return value != 0
        text = str(value).strip().lower()
        return text not in ('', '0', 'false', 'none', 'null', 'no', 'n', '否', '不', '无需', '不需要')

    def _evaluate_condition(self, expression: str, data_pool: dict, input_values: dict, field_configs: dict):
        expression = (expression or '').strip()
        if not expression:
            return True

        def resolver(field_name):
            special = self._special_value(field_name)
            if special is not None:
                return special
            return self._first_existing_value(
                (data_pool, field_name),
                (input_values, field_name),
                (field_configs, field_name),
                default=''
            )

        result = self._evaluate_expression(expression, resolver)
        return self._coerce_condition_result(result)

    @staticmethod
    def _build_initial_pool(input_values: dict, field_configs: dict):
        pool = {}
        for key, value in (field_configs or {}).items():
            pool[str(key)] = '' if value is None else value
        for key, value in (input_values or {}).items():
            pool[str(key)] = '' if value is None else value
        return pool

    def get_field_options(self, config: dict, input_values: dict = None) -> list:
        if not config:
            return ['示例选项1', '示例选项2']

        if config.get('type') == 'fixed':
            values = config.get('values', [])
            if not values:
                return ['（请配置固定选项）']

            input_values = dict(input_values or {})

            def resolver(field_name):
                special = self._special_value(field_name)
                if special is not None:
                    return special
                return self._lookup_mapping_value(input_values, field_name, default='')

            result = []
            for raw_value in values:
                rendered = self._render_recursive_text(raw_value, resolver)
                parts = str(rendered).splitlines() or [str(rendered)]
                for part in parts:
                    part = str(part).strip()
                    if part:
                        result.append(part)
            return result or ['（请配置固定选项）']

        if config.get('type') == 'table':
            table = config.get('table')
            column = config.get('column')
            if table and column:
                try:
                    rows = self.main_db.fetch_all(
                        f'SELECT DISTINCT "{column}" FROM "{table}" ORDER BY "{column}"'
                    )
                    result = []
                    for row in rows:
                        value = row.get(column)
                        if value is not None:
                            result.append(str(value))
                    return result or ['（表中无数据）']
                except Exception:
                    return ['（查询失败）']
            return ['（未配置表/列）']

        return ['未知配置类型']

    def apply_rules(self, rules: list, initial_pool: dict) -> dict:
        data_pool = dict(initial_pool or {})

        for rule in rules:
            in_field = rule.get('input_field', '')
            in_value = data_pool.get(in_field, '')
            op = rule.get('operation')
            out_field = rule.get('output_field', '')
            params = rule.get('params', {}) or {}

            if not out_field:
                continue

            if op == '直接传递':
                data_pool[out_field] = in_value

            elif op == '数学公式':
                formula = params.get('formula', '')
                if not formula:
                    data_pool[out_field] = '[公式为空]'
                    continue

                try:
                    def resolver(field_name):
                        special = self._special_value(field_name)
                        if special is not None:
                            return special
                        return data_pool.get(field_name)
                    result = self._evaluate_expression(formula, resolver)
                    data_pool[out_field] = '' if result is None else str(result)
                except Exception as e:
                    data_pool[out_field] = f'[公式错误: {e}]'

            elif op == '数据库查询':
                table = params.get('table')
                key_col = params.get('key_column')
                value_col = params.get('value_column')

                if table and key_col and value_col:
                    try:
                        rows = self.main_db.fetch_all(
                            f'SELECT "{value_col}" FROM "{table}" WHERE "{key_col}" = ?',
                            (in_value,)
                        )
                        data_pool[out_field] = '' if not rows else ('' if rows[0][value_col] is None else str(rows[0][value_col]))
                    except Exception as e:
                        data_pool[out_field] = f'[查询失败: {e}]'
                else:
                    data_pool[out_field] = '[配置不完整]'

            else:
                data_pool[out_field] = f'[未知操作: {op}]'

        return data_pool

    def _render_formula_tags(self, text: str, resolver, max_passes: int = 12) -> str:
        current = '' if text is None else str(text)
        for _ in range(max_passes):
            changed = False

            def replace_formula(match):
                nonlocal changed
                expr = match.group(1).strip()
                try:
                    value = self._evaluate_expression(expr, resolver)
                    changed = True
                    return '' if value is None else str(value)
                except Exception as e:
                    changed = True
                    return f'[公式错误: {e}]'

            new_text = self._FORMULA_TAG_PATTERN.sub(replace_formula, current)
            current = new_text
            if not changed:
                break
        return current

    def _render_placeholders(self, text: str, resolver, selected_fields=None, visible_fields=None) -> str:
        selected_set = set(selected_fields or [])
        visible_set = set(visible_fields or [])

        def replace(match):
            key = match.group(1).strip()
            special = self._special_value(key)
            if special is not None:
                return special
            if key in selected_set and key not in visible_set:
                return ''
            value = resolver(key)
            return '' if value is None else str(value)

        return self._PLACEHOLDER_PATTERN.sub(replace, '' if text is None else str(text))

    def _render_recursive_text(self, text: str, resolver, selected_fields=None, visible_fields=None, max_passes: int = 12) -> str:
        current = '' if text is None else str(text)
        for _ in range(max_passes):
            new_text = self._render_formula_tags(current, resolver, max_passes=1)
            new_text = self._render_placeholders(new_text, resolver, selected_fields=selected_fields, visible_fields=visible_fields)
            if new_text == current:
                break
            current = new_text
        return current

    def render(self, main_template_name: str, input_values: dict,
               main_config_override: dict = None,
               process_content_override: dict = None) -> tuple:
        main_tpl = self.template_db.get_main_template(main_template_name)
        if not main_tpl and main_config_override is None:
            return '', {}, {}

        proc_tpl = self.template_db.get_process_template(main_template_name)
        if not proc_tpl and process_content_override is None:
            return '', {}, {}

        config = (main_config_override if main_config_override is not None
                  else (main_tpl.get('config', {}) if main_tpl else {})) or {}
        content = (process_content_override if process_content_override is not None
                   else (proc_tpl.get('content', {}) if proc_tpl else {})) or {}
        rules = config.get('rules', []) or []
        field_configs = content.get('available_fields', {}) or {}
        initial_pool = self._build_initial_pool(input_values or {}, field_configs)
        data_pool = self.apply_rules(rules, initial_pool)

        stored_preview = content.get('preview_format', '') or ''
        selected_fields = content.get('selected_fields', []) or []
        field_conditions = content.get('field_conditions', {}) or {}

        def resolve_field_value(field_name):
            special = self._special_value(field_name)
            if special is not None:
                return special
            return self._first_existing_value(
                (data_pool, field_name),
                (input_values, field_name),
                (field_configs, field_name),
                default=''
            )

        visible_fields = []
        for field in selected_fields:
            expr = field_conditions.get(field, '')
            try:
                visible = self._evaluate_condition(expr, data_pool, input_values or {}, field_configs)
            except Exception:
                visible = True
            if visible:
                visible_fields.append(field)

        template_text = ''.join(str(resolve_field_value(f)) for f in visible_fields)
        if not template_text and not visible_fields and stored_preview:
            template_text = stored_preview

        result = self._render_recursive_text(
            template_text,
            resolve_field_value,
            selected_fields=selected_fields,
            visible_fields=visible_fields,
        )

        final_fields = {}
        for field in visible_fields:
            raw_val = resolve_field_value(field)
            final_val = self._render_recursive_text(
                raw_val,
                resolve_field_value,
                selected_fields=selected_fields,
                visible_fields=visible_fields,
            )
            final_fields[field] = final_val

        return result, data_pool, final_fields
