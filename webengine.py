import re
import time
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP
from typing import Callable, List, Optional

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class BrowserEngineError(Exception):
    pass


class BrowserEngine:
    def __init__(self):
        self.driver = None
        self.main_window_handle = None
        self.preferred_window_handle = None
        self.current_config = {}

    def is_connected(self) -> bool:
        try:
            return self.driver is not None and len(self.driver.window_handles) >= 1
        except Exception:
            return False

    def _log(self, logger: Optional[Callable[[str], None]], message: str):
        if logger:
            logger(message)

    @staticmethod
    def _normalize_path(value: str) -> str:
        return (value or '').strip()

    @staticmethod
    def normalize_url(url: str) -> str:
        url = (url or '').strip()
        if not url:
            return ''
        lowered = url.lower()
        special_prefixes = ('about:', 'chrome://', 'edge://', 'file://', 'data:', 'javascript:')
        if lowered.startswith(special_prefixes):
            return url
        if re.match(r'^[a-zA-Z][a-zA-Z0-9+\-.]*://', url):
            return url
        if url.startswith('//'):
            return 'https:' + url
        return 'https://' + url

    def _build_options(self, chrome_binary: str = '', extra_args: Optional[List[str]] = None) -> Options:
        options = Options()
        chrome_binary = self._normalize_path(chrome_binary)
        if chrome_binary:
            options.binary_location = chrome_binary
        for arg in (extra_args or []):
            if arg:
                options.add_argument(arg)
        return options

    def _create_driver(self, chromedriver_path: str, options: Options):
        chromedriver_path = self._normalize_path(chromedriver_path)
        service = Service(executable_path=chromedriver_path) if chromedriver_path else Service()
        return webdriver.Chrome(service=service, options=options)

    def launch_browser(self, chromedriver_path: str, chrome_binary: str = '', start_url: str = '', debug_port: int = 9222, logger=None):
        start_url = self.normalize_url(start_url)
        if self.is_connected():
            self._log(logger, '检测到已有受控浏览器，会继续复用当前浏览器。')
            self.repair_session_window(logger=logger, activate_preferred=True)
            if start_url:
                self.open_url(start_url, logger=logger)
            return self.driver

        self._log(logger, '正在启动受控 Chrome ...')
        options = self._build_options(
            chrome_binary=chrome_binary,
            extra_args=[
                f'--remote-debugging-port={int(debug_port)}',
                '--disable-popup-blocking',
                '--start-maximized',
            ],
        )
        self.driver = self._create_driver(chromedriver_path, options)
        self.main_window_handle = self.driver.current_window_handle
        self.preferred_window_handle = self.main_window_handle
        self.current_config = {
            'mode': 'launch',
            'chromedriver_path': chromedriver_path,
            'chrome_binary': chrome_binary,
            'start_url': start_url,
            'debug_port': debug_port,
        }
        if start_url:
            self.driver.get(start_url)
        self._log(logger, '浏览器已启动。')
        return self.driver

    def ensure_connected(self, browser_config: dict, logger=None):
        if self.is_connected():
            self.repair_session_window(logger=logger, activate_preferred=True)
            return self.driver

        browser_config = browser_config or {}
        chromedriver_path = browser_config.get('chromedriver_path', '')
        chrome_binary = browser_config.get('chrome_binary', '')
        start_url = browser_config.get('start_url', '')
        debug_port = int(browser_config.get('debug_port', 9222) or 9222)
        return self.launch_browser(chromedriver_path, chrome_binary=chrome_binary, start_url=start_url, debug_port=debug_port, logger=logger)

    def disconnect(self):
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
        self.driver = None
        self.main_window_handle = None
        self.preferred_window_handle = None
        self.current_config = {}

    def _safe_window_handles(self) -> List[str]:
        if not self.driver:
            return []
        try:
            return list(self.driver.window_handles)
        except Exception as e:
            raise BrowserEngineError(f'无法获取浏览器窗口列表：{e}')

    def set_preferred_window(self, handle: str, logger=None):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        handles = self._safe_window_handles()
        if handle not in handles:
            raise BrowserEngineError('目标窗口已经关闭。')
        self.preferred_window_handle = handle
        self._log(logger, '已设置当前目标窗口，后续操作将针对该窗口。')
        return handle

    def repair_session_window(self, logger=None, activate_preferred: bool = True) -> Optional[str]:
        if not self.driver:
            return None
        handles = self._safe_window_handles()
        if not handles:
            self.main_window_handle = None
            self.preferred_window_handle = None
            raise BrowserEngineError('浏览器中没有可用窗口。')

        current_handle = None
        try:
            current_handle = self.driver.current_window_handle
        except Exception:
            current_handle = None

        if self.main_window_handle not in handles:
            self.main_window_handle = handles[0]
            self._log(logger, '主窗口句柄已更新为当前可用窗口。')
        if self.preferred_window_handle not in handles:
            self.preferred_window_handle = self.main_window_handle

        if current_handle not in handles:
            fallback = self.preferred_window_handle if self.preferred_window_handle in handles else self.main_window_handle
            fallback = fallback if fallback in handles else handles[0]
            try:
                self.driver.switch_to.window(fallback)
                self._log(logger, '检测到当前窗口句柄已失效，已自动切换到仍可用的窗口。')
            except Exception as e:
                raise BrowserEngineError(f'当前窗口已关闭，且自动切换失败：{e}')
            current_handle = fallback

        if activate_preferred and self.preferred_window_handle in handles and current_handle != self.preferred_window_handle:
            try:
                self.driver.switch_to.window(self.preferred_window_handle)
            except Exception as e:
                raise BrowserEngineError(f'切换到当前目标窗口失败：{e}')
            current_handle = self.preferred_window_handle

        return current_handle

    def open_url(self, url: str, logger=None):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        self.repair_session_window(logger=logger, activate_preferred=True)
        url = self.normalize_url(url)
        if url:
            self.driver.get(url)
            self._log(logger, f'已打开网址：{url}')

    def _read_current_window_meta(self):
        title = ''
        url = ''
        try:
            title = self.driver.title or ''
        except Exception:
            title = ''
        try:
            url = self.driver.current_url or ''
        except Exception:
            url = ''

        if not title or not url:
            try:
                meta = self.driver.execute_script(
                    'return {title: (document && document.title) || "", url: (window && window.location && window.location.href) || ""};'
                ) or {}
                if not title:
                    title = meta.get('title', '') or title
                if not url:
                    url = meta.get('url', '') or url
            except Exception:
                pass
        return title, url

    def list_windows(self) -> List[dict]:
        if not self.is_connected():
            return []

        handles = self._safe_window_handles()
        if not handles:
            return []
        self.repair_session_window(activate_preferred=False)

        result = []
        restore_handle = None
        try:
            restore_handle = self.driver.current_window_handle
        except Exception:
            restore_handle = None

        for idx, handle in enumerate(handles):
            title = ''
            url = ''
            try:
                self.driver.switch_to.window(handle)
                time.sleep(0.05)
                title, url = self._read_current_window_meta()
                if not title and not url:
                    time.sleep(0.1)
                    title, url = self._read_current_window_meta()
            except Exception:
                pass
            result.append({
                'index': idx,
                'handle': handle,
                'title': title,
                'url': url,
            })

        if restore_handle in handles:
            try:
                self.driver.switch_to.window(restore_handle)
            except Exception:
                pass
        return result

    def switch_to_window(self, handle: str):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        handles = self._safe_window_handles()
        if handle not in handles:
            raise BrowserEngineError('目标窗口已经关闭。')
        self.preferred_window_handle = handle
        self.driver.switch_to.window(handle)

    def switch_to_main_window(self):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        handles = self._safe_window_handles()
        if self.main_window_handle not in handles:
            self.main_window_handle = handles[0] if handles else None
        if self.main_window_handle:
            self.preferred_window_handle = self.main_window_handle
            self.driver.switch_to.window(self.main_window_handle)

    def switch_window_by_rule(self, match_type: str, match_value: str):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')

        windows = self.list_windows()
        value = (match_value or '').strip()
        if match_type == '序号':
            index = int(value or 0)
            for item in windows:
                if item['index'] == index:
                    self.switch_to_window(item['handle'])
                    return item
        elif match_type == 'URL包含':
            for item in windows:
                if value in (item['url'] or ''):
                    self.switch_to_window(item['handle'])
                    return item
        else:
            for item in windows:
                if value in (item['title'] or ''):
                    self.switch_to_window(item['handle'])
                    return item
        raise BrowserEngineError(f'未找到匹配窗口：{match_type} = {value}')

    @staticmethod
    def _locator_to_by(locator_type: str):
        locator_type = (locator_type or '').strip().lower()
        mapping = {
            'id': By.ID,
            'name': By.NAME,
            'xpath': By.XPATH,
            'css': By.CSS_SELECTOR,
            'css selector': By.CSS_SELECTOR,
            'class': By.CLASS_NAME,
            'class name': By.CLASS_NAME,
            'tag': By.TAG_NAME,
            'tag name': By.TAG_NAME,
            'link text': By.LINK_TEXT,
            'partial link text': By.PARTIAL_LINK_TEXT,
        }
        if locator_type not in mapping:
            raise BrowserEngineError(f'不支持的定位方式：{locator_type}')
        return mapping[locator_type]

    def wait_for_element(self, locator_type: str, locator_value: str, timeout: float = 10, clickable: bool = False):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        self.repair_session_window(activate_preferred=True)
        by = self._locator_to_by(locator_type)
        wait = WebDriverWait(self.driver, timeout)
        condition = EC.element_to_be_clickable((by, locator_value)) if clickable else EC.presence_of_element_located((by, locator_value))
        return wait.until(condition)

    def inspect_elements(self, limit: int = 3000) -> List[dict]:
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        self.repair_session_window(activate_preferred=True)
        script = r'''
        function cssPath(el) {
            if (!(el instanceof Element)) return '';
            var path = [];
            while (el && el.nodeType === Node.ELEMENT_NODE) {
                var selector = el.nodeName.toLowerCase();
                if (el.id) {
                    selector += '#' + el.id;
                    path.unshift(selector);
                    break;
                } else {
                    var sib = el, nth = 1;
                    while (sib = sib.previousElementSibling) {
                        if (sib.nodeName.toLowerCase() === selector) nth++;
                    }
                    selector += ':nth-of-type(' + nth + ')';
                }
                path.unshift(selector);
                el = el.parentNode;
            }
            return path.join(' > ');
        }
        function xpath(el) {
            if (el.id) return '//*[@id="' + el.id + '"]';
            if (el === document.body) return '/html/body';
            var ix = 0;
            var siblings = el.parentNode ? el.parentNode.childNodes : [];
            for (var i = 0; i < siblings.length; i++) {
                var sibling = siblings[i];
                if (sibling === el) {
                    return xpath(el.parentNode) + '/' + el.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                }
                if (sibling.nodeType === 1 && sibling.tagName === el.tagName) {
                    ix++;
                }
            }
            return '';
        }
        function textOf(el) {
            var text = el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.getAttribute('title') || '';
            text = (text || '').replace(/\s+/g, ' ').trim();
            return text.slice(0, 120);
        }
        function isVisible(el) {
            var rect = el.getBoundingClientRect();
            var style = window.getComputedStyle(el);
            if (style.display === 'none' || style.visibility === 'hidden') return false;
            if (style.opacity === '0') return false;
            if (rect.width === 0 && rect.height === 0) return false;
            return true;
        }
        var selector = [
            'input','textarea','select','button','a','label','option','iframe',
            '[id]','[name]','[placeholder]','[title]','[role]','[onclick]','[aria-label]','[contenteditable="true"]'
        ].join(',');
        var nodes = Array.from(document.querySelectorAll(selector));
        var seen = new Set();
        var filtered = [];
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            if (!(el instanceof Element)) continue;
            if (!isVisible(el)) continue;
            var xp = xpath(el);
            var key = xp || cssPath(el);
            if (!key || seen.has(key)) continue;
            seen.add(key);
            filtered.push({
                tag: (el.tagName || '').toLowerCase(),
                type: el.getAttribute('type') || '',
                text: textOf(el),
                id: el.id || '',
                name: el.getAttribute('name') || '',
                placeholder: el.getAttribute('placeholder') || '',
                title: el.getAttribute('title') || el.getAttribute('aria-label') || '',
                value: el.value || '',
                css: cssPath(el),
                xpath: xp
            });
            if (filtered.length >= LIMIT_PLACEHOLDER) break;
        }
        return filtered;
        '''.replace('LIMIT_PLACEHOLDER', str(int(limit)))
        return self.driver.execute_script(script)

    @staticmethod
    def _safe_round_half_up(value, ndigits=0):
        number = Decimal(str(value or 0))
        ndigits = int(float(ndigits or 0))
        quant = Decimal('1').scaleb(-ndigits)
        result = number.quantize(quant, rounding=ROUND_HALF_UP)
        if ndigits <= 0 and result == result.to_integral_value():
            return int(result)
        return float(result)

    @staticmethod
    def _safe_roundup(value, ndigits=0):
        number = Decimal(str(value or 0))
        ndigits = int(float(ndigits or 0))
        quant = Decimal('1').scaleb(-ndigits)
        result = number.quantize(quant, rounding=ROUND_UP)
        if ndigits <= 0 and result == result.to_integral_value():
            return int(result)
        return float(result)

    @staticmethod
    def _safe_isblank(value):
        if value is None:
            return True
        return str(value).strip() == ''

    @staticmethod
    def _safe_nl():
        return '\n'

    @staticmethod
    def _normalize_expr_text(expression: str) -> str:
        return str(expression or '').translate(str.maketrans({'（': '(', '）': ')', '，': ','}))

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
        raise BrowserEngineError('条件表达式括号未正确闭合')

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
                    raise BrowserEngineError('if 条件函数必须写成 if(条件, 真值, 假值)')
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
        pattern = re.compile(rf'(?i){func_name}\s*\(')
        while index < len(text):
            match = pattern.match(text[index:])
            prev_ok = index == 0 or not (text[index - 1].isalnum() or text[index - 1] == '_')
            if match and prev_ok:
                token = match.group(0)
                open_index = index + token.rfind('(')
                close_index = cls._find_matching_paren(text, open_index)
                inner = text[open_index + 1:close_index]
                parts = [cls._rewrite_expression_functions(part) for part in cls._split_top_level_args(inner) if part.strip()]
                if not parts:
                    raise BrowserEngineError(f'{func_name} 函数至少需要 1 个参数')
                joined = f' {operator} '.join(f'({part})' for part in parts)
                result.append(f'({joined})')
                index = close_index + 1
                continue
            result.append(text[index])
            index += 1
        return ''.join(result)

    @classmethod
    def _rewrite_expression_functions(cls, text: str) -> str:
        text = cls._rewrite_if_functions(str(text or ''))
        text = cls._rewrite_logic_function(text, 'and', 'and')
        text = cls._rewrite_logic_function(text, 'or', 'or')
        return text

    @classmethod
    def _template_safe_globals(cls):
        return {
            '__builtins__': {},
            'int': lambda value=0: int(float(str(value or 0).strip() or 0)),
            'INT': lambda value=0: int(float(str(value or 0).strip() or 0)),
            'float': lambda value=0: float(str(value or 0).strip() or 0),
            'FLOAT': lambda value=0: float(str(value or 0).strip() or 0),
            'round': cls._safe_round_half_up,
            'ROUND': cls._safe_round_half_up,
            'roundup': cls._safe_roundup,
            'ROUNDUP': cls._safe_roundup,
            'abs': abs,
            'ABS': abs,
            'max': max,
            'MAX': max,
            'min': min,
            'MIN': min,
            'isblank': cls._safe_isblank,
            'ISBLANK': cls._safe_isblank,
            'nl': cls._safe_nl,
            'NL': cls._safe_nl,
            'newline': cls._safe_nl,
            'NEWLINE': cls._safe_nl,
            'True': True,
            'False': False,
        }

    @staticmethod
    def _payload_lookup_static(payload: dict, key: str):
        special_map = {
            '__RESULT__': payload.get('result_text', ''),
            'RESULT_TEXT': payload.get('result_text', ''),
            '__TEMPLATE__': payload.get('template_name', ''),
            'TEMPLATE_NAME': payload.get('template_name', ''),
            '__NL__': '\n',
            'NL': '\n',
            '__BR__': '\n',
            '换行': '\n',
        }
        if key in special_map:
            return special_map.get(key, '')
        for bucket_name in ('data_pool', 'final_fields', 'input_values'):
            bucket = payload.get(bucket_name, {}) or {}
            if key in bucket:
                return bucket.get(key)
        return special_map.get(key, '')

    def _payload_lookup(self, payload: dict, key: str):
        return self._payload_lookup_static(payload, key)

    @classmethod
    def _evaluate_template_expression(cls, expression: str, payload: dict):
        expr_text = cls._rewrite_expression_functions(cls._normalize_expr_text(expression))
        context = {}
        counter = 0

        def replace_braced(match):
            nonlocal counter
            field = match.group(1).strip()
            var_name = f'__p{counter}'
            counter += 1
            value = cls._payload_lookup_static(payload, field)
            if isinstance(value, str):
                stripped = value.strip()
                if stripped.lower() == 'true':
                    value = True
                elif stripped.lower() == 'false':
                    value = False
                else:
                    try:
                        number = float(stripped)
                        value = int(number) if number.is_integer() else number
                    except Exception:
                        value = stripped
            context[var_name] = value
            return var_name

        expr = re.sub(r'\{([^{}]+)\}', replace_braced, expr_text)
        return eval(expr, cls._template_safe_globals(), context)

    @classmethod
    def _expand_formula_tags(cls, text: str, payload: dict) -> str:
        def replace_formula(match):
            expr = match.group(1).strip()
            try:
                value = cls._evaluate_template_expression(expr, payload)
                return '' if value is None else str(value)
            except Exception as e:
                return f'[公式错误: {e}]'
        return re.sub(r'#([^#]+?)#', replace_formula, str(text), flags=re.DOTALL)

    @classmethod
    def _replace_payload_placeholders(cls, text: str, payload: dict) -> str:
        def replace(match):
            key = match.group(1).strip()
            value = cls._payload_lookup_static(payload, key)
            return '' if value is None else str(value)
        return re.sub(r'\{([^{}]+)\}', replace, str(text))

    @classmethod
    def _expand_text_template(cls, template: str, payload: dict, max_passes: int = 12) -> str:
        if template is None:
            return ''
        current = str(template)
        for _ in range(max_passes):
            previous = current
            current = cls._expand_formula_tags(current, payload)
            current = cls._replace_payload_placeholders(current, payload)
            if current == previous:
                break
        return current

    @classmethod
    def render_value_template(cls, template: str, payload: dict) -> str:
        return cls._expand_text_template(template, payload)

    @staticmethod
    def _condition_truthy(value):
        if isinstance(value, bool):
            return value
        if value is None:
            return False
        if isinstance(value, (int, float)):
            return value != 0
        text = str(value).strip().lower()
        return text not in ('', '0', 'false', 'none', 'null', 'no', 'n', '否', '不', '无需', '不需要')

    def evaluate_condition_expression(self, expression: str, payload: dict) -> bool:
        expression = (expression or '').strip()
        if not expression:
            return True
        try:
            result = self._evaluate_template_expression(expression, payload)
        except Exception as e:
            raise BrowserEngineError(f'步骤条件表达式错误：{e}')
        return self._condition_truthy(result)

    @staticmethod
    def _apply_char_policy(value: str, mode: str, kind: str) -> str:
        mode = (mode or '直接输入').strip()
        if kind == 'newline':
            if mode == '删除':
                return value.replace('\r\n', '').replace('\n', '').replace('\r', '')
            if mode == '转为空格':
                return value.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
            if mode == '转为\\n':
                return value.replace('\r\n', '\\n').replace('\n', '\\n').replace('\r', '\\n')
            return value
        if kind == 'tab':
            if mode == '删除':
                return value.replace('\t', '')
            if mode == '转为4空格':
                return value.replace('\t', '    ')
            if mode == '转为\\t':
                return value.replace('\t', '\\t')
            return value
        if kind == 'space':
            if mode == '压缩为1个':
                return re.sub(r' {2,}', ' ', value)
            if mode == '删除全部':
                return value.replace(' ', '')
            return value
        return value

    def transform_input_value(self, value: str, step: dict) -> str:
        text = '' if value is None else str(value)
        text = self._apply_char_policy(text, step.get('newline_mode', '直接输入'), 'newline')
        text = self._apply_char_policy(text, step.get('tab_mode', '直接输入'), 'tab')
        text = self._apply_char_policy(text, step.get('space_mode', '直接输入'), 'space')
        return text

    @staticmethod
    def _send_text(element, value: str):
        if value == '':
            return
        element.send_keys(value)

    def execute_flow(self, flow_config: dict, payload: dict, logger=None):
        browser_cfg = (flow_config or {}).get('browser', {}) or {}
        self.ensure_connected(browser_cfg, logger=logger)

        implicit_wait = float(browser_cfg.get('implicit_wait', 0) or 0)
        if implicit_wait > 0:
            self.driver.implicitly_wait(implicit_wait)

        if not self.main_window_handle:
            try:
                self.main_window_handle = self.driver.current_window_handle
            except Exception:
                self.main_window_handle = None
        if not self.preferred_window_handle:
            self.preferred_window_handle = self.main_window_handle

        steps = (flow_config or {}).get('steps', []) or []
        if not steps:
            raise BrowserEngineError('当前模板未配置浏览器导出步骤。')

        context = {
            'main_handle': self.main_window_handle,
            'last_window_count': len(self.driver.window_handles),
        }

        for idx, step in enumerate(steps, start=1):
            action = (step.get('action') or '').strip()
            name = (step.get('name') or f'步骤{idx}').strip()
            timeout = float(step.get('wait_timeout', 10) or 10)
            self._log(logger, f'[{idx}] {name} - {action}')

            if not self.evaluate_condition_expression(step.get('condition_expr', ''), payload):
                self._log(logger, f'[{idx}] 已跳过：未满足执行条件')
                continue

            if action == '点击元素':
                element = self.wait_for_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout, clickable=True)
                if step.get('use_js_click'):
                    self.driver.execute_script('arguments[0].click();', element)
                else:
                    element.click()
                context['last_window_count'] = len(self.driver.window_handles)

            elif action == '输入文本':
                element = self.wait_for_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout, clickable=False)
                value = self.render_value_template(step.get('value_template', ''), payload)
                value = self.transform_input_value(value, step)
                if step.get('clear_before_input', True):
                    try:
                        element.clear()
                    except Exception:
                        pass
                self._send_text(element, value)

            elif action == '等待元素':
                self.wait_for_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout, clickable=bool(step.get('wait_clickable', False)))

            elif action == '等待新窗口':
                old_count = context.get('last_window_count', len(self.driver.window_handles))
                WebDriverWait(self.driver, timeout).until(lambda d: len(d.window_handles) > old_count)
                context['last_window_count'] = len(self.driver.window_handles)

            elif action == '切换窗口':
                item = self.switch_window_by_rule(step.get('window_match_type', '标题包含'), step.get('window_match_value', ''))
                self._log(logger, f"已切换到窗口：{item.get('title') or item.get('url')}")

            elif action == '切回主窗口':
                self.switch_to_main_window()

            elif action == '切换iframe':
                frame = self.wait_for_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout)
                self.driver.switch_to.frame(frame)

            elif action == '切回默认文档':
                self.driver.switch_to.default_content()

            elif action == '延时':
                time.sleep(float(step.get('sleep_seconds', 1) or 1))

            else:
                raise BrowserEngineError(f'不支持的动作：{action}')

        return True
