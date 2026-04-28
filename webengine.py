import re
import time
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP
from typing import Callable, List, Optional

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


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

    def wait_for_element_gone(self, locator_type: str, locator_value: str, timeout: float = 10):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        self.repair_session_window(activate_preferred=True)
        by = self._locator_to_by(locator_type)
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.invisibility_of_element_located((by, locator_value)))

    def _find_elements(self, locator_type: str, locator_value: str):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        self.repair_session_window(activate_preferred=True)
        by = self._locator_to_by(locator_type)
        return self.driver.find_elements(by, locator_value)

    def _scroll_into_view(self, element):
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", element)
        except Exception:
            try:
                self.driver.execute_script('arguments[0].scrollIntoView(true);', element)
            except Exception:
                pass

    def _retry_find_element(self, locator_type: str, locator_value: str, timeout: float = 10, clickable: bool = False, attempts: int = 3):
        last_error = None
        for _ in range(max(1, attempts)):
            try:
                element = self.wait_for_element(locator_type, locator_value, timeout=timeout, clickable=clickable)
                self._scroll_into_view(element)
                return element
            except Exception as e:
                last_error = e
                time.sleep(0.2)
        if last_error:
            raise last_error
        raise BrowserEngineError('未找到元素')

    @staticmethod
    def _is_retryable_click_error(error: Exception) -> bool:
        text = str(error or '').lower()
        retry_keywords = ('stale', 'not attached', 'other element would receive', 'intercept', 'not clickable', 'detached')
        return isinstance(error, (StaleElementReferenceException, WebDriverException)) or any(keyword in text for keyword in retry_keywords)

    def _click_element(self, locator_type: str, locator_value: str, timeout: float = 10, use_js_click: bool = False):
        last_error = None
        for attempt in range(3):
            try:
                element = self._retry_find_element(locator_type, locator_value, timeout=timeout, clickable=not use_js_click, attempts=1)
                if use_js_click:
                    self.driver.execute_script('arguments[0].click();', element)
                else:
                    element.click()
                return True
            except Exception as e:
                last_error = e
                if not self._is_retryable_click_error(e) or attempt >= 2:
                    break
                time.sleep(0.3)
        if last_error:
            raise last_error
        raise BrowserEngineError('点击元素失败')

    def _right_click_element(self, locator_type: str, locator_value: str, timeout: float = 10):
        element = self._retry_find_element(locator_type, locator_value, timeout=timeout, clickable=True)
        ActionChains(self.driver).move_to_element(element).context_click(element).perform()
        return True

    def _right_click_menu_item(self, step: dict, timeout: float = 10):
        self._right_click_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout)
        self._click_element(step.get('target_locator_type', ''), step.get('target_locator_value', ''), timeout=timeout)
        return True

    def _dropdown_two_stage(self, step: dict, timeout: float = 10):
        self._click_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout, use_js_click=bool(step.get('use_js_click', False)))
        self._click_element(step.get('target_locator_type', ''), step.get('target_locator_value', ''), timeout=timeout)
        return True

    def _input_text(self, locator_type: str, locator_value: str, value: str, timeout: float = 10, clear_before_input: bool = True):
        element = self._retry_find_element(locator_type, locator_value, timeout=timeout, clickable=False)
        if clear_before_input:
            try:
                element.clear()
            except Exception:
                pass
        self._send_text(element, value)
        return True

    def _set_element_value(self, element, value: str, clear_before_input: bool = True):
        tag = (getattr(element, 'tag_name', '') or '').lower()
        if tag == 'select':
            self._select_option_like(element, value)
            return
        if clear_before_input:
            try:
                element.clear()
            except Exception:
                try:
                    self.driver.execute_script("arguments[0].value='';", element)
                except Exception:
                    pass
        try:
            element.send_keys(value)
        except Exception:
            self.driver.execute_script(
                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input', {bubbles:true})); arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                element,
                value,
            )

    def _select_option_like(self, element, value: str):
        text = '' if value is None else str(value)
        try:
            select = Select(element)
            try:
                select.select_by_visible_text(text)
                return True
            except Exception:
                pass
            for option in select.options:
                if (option.get_attribute('value') or '') == text:
                    select.select_by_value(text)
                    return True
        except Exception:
            pass
        try:
            self.driver.execute_script(
                "var el=arguments[0], target=arguments[1];"
                "for (var i=0;i<el.options.length;i++){var o=el.options[i]; if(String(o.text).trim()===target || String(o.value)===target){el.selectedIndex=i; el.dispatchEvent(new Event('change',{bubbles:true})); return true;}} return false;",
                element,
                text,
            )
            return True
        except Exception:
            return False

    def _multi_select(self, locator_type: str, locator_values, timeout: float = 10):
        values = [str(v).strip() for v in (locator_values or []) if str(v).strip()]
        if not values:
            raise BrowserEngineError('多选元素/多行选择至少需要 1 个定位值')
        actions = ActionChains(self.driver)
        actions.key_down(Keys.CONTROL)
        performed = False
        for value in values:
            element = self._retry_find_element(locator_type, value, timeout=timeout, clickable=True)
            self._scroll_into_view(element)
            actions.click(element)
            performed = True
        actions.key_up(Keys.CONTROL)
        if performed:
            actions.perform()
        return performed

    @classmethod
    def _parse_key_combo(cls, text: str):
        mapping = {
            'CTRL': Keys.CONTROL,
            'CONTROL': Keys.CONTROL,
            'SHIFT': Keys.SHIFT,
            'ALT': Keys.ALT,
            'ENTER': Keys.ENTER,
            'RETURN': Keys.RETURN,
            'TAB': Keys.TAB,
            'ESC': Keys.ESCAPE,
            'ESCAPE': Keys.ESCAPE,
            'DELETE': Keys.DELETE,
            'BACKSPACE': Keys.BACKSPACE,
            'SPACE': Keys.SPACE,
            'UP': Keys.ARROW_UP,
            'DOWN': Keys.ARROW_DOWN,
            'LEFT': Keys.ARROW_LEFT,
            'RIGHT': Keys.ARROW_RIGHT,
            'HOME': Keys.HOME,
            'END': Keys.END,
            'PAGEUP': Keys.PAGE_UP,
            'PAGEDOWN': Keys.PAGE_DOWN,
        }
        result = []
        for part in re.split(r'\s*\+\s*', str(text or '').strip()):
            part = str(part).strip()
            if not part:
                continue
            upper = part.upper()
            if upper in mapping:
                result.append(mapping[upper])
            elif re.match(r'^F\d{1,2}$', upper):
                result.append(getattr(Keys, upper, upper))
            elif len(part) == 1:
                result.append(part)
            else:
                result.append(part)
        return result

    def _send_key_combo(self, locator_type: str, locator_value: str, combo_text: str, timeout: float = 10):
        keys = self._parse_key_combo(combo_text)
        if not keys:
            raise BrowserEngineError('键盘组合键不能为空')
        element = None
        if str(locator_value or '').strip():
            element = self._retry_find_element(locator_type, locator_value, timeout=timeout, clickable=False)
            try:
                element.click()
            except Exception:
                pass
        else:
            try:
                element = self.driver.switch_to.active_element
            except Exception:
                element = self.driver.find_element(By.TAG_NAME, 'body')
        ActionChains(self.driver).send_keys_to_element(element, *keys).perform()
        return True

    def _drag_drop_element(self, step: dict, timeout: float = 10):
        source = self._retry_find_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout, clickable=True)
        target = self._retry_find_element(step.get('target_locator_type', ''), step.get('target_locator_value', ''), timeout=timeout, clickable=True)
        target_rect = getattr(target, 'rect', {}) or {}
        width = max(1, int(float(target_rect.get('width', 1) or 1)))
        height = max(1, int(float(target_rect.get('height', 1) or 1)))
        mode = (step.get('drop_position') or '中间').strip()
        if mode == '上方':
            offset_x, offset_y = width // 2, 2
        elif mode == '下方':
            offset_x, offset_y = width // 2, max(1, height - 2)
        elif mode == '自定义偏移':
            offset_x = int(float(step.get('drag_offset_x', 0) or 0))
            offset_y = int(float(step.get('drag_offset_y', 0) or 0))
        else:
            offset_x, offset_y = width // 2, height // 2
        actions = ActionChains(self.driver)
        actions.move_to_element(source).click_and_hold(source).pause(0.15)
        actions.move_to_element_with_offset(target, offset_x, offset_y).pause(0.15).release().perform()
        return True

    def _add_table(self, locator_type: str, locator_value: str, timeout: float = 10):
        return self._click_element(locator_type, locator_value, timeout=timeout)

    def _fill_table_cells(self, locator_type: str, locator_value: str, table_text: str, timeout: float = 10, clear_before_input: bool = True):
        table = self._retry_find_element(locator_type, locator_value, timeout=timeout, clickable=False)
        rows = [line for line in str(table_text or '').splitlines() if line.strip() != '']
        if not rows:
            raise BrowserEngineError('自动填单元格内容不能为空')
        matrix = [line.split('\t') for line in rows]
        row_nodes = table.find_elements(By.CSS_SELECTOR, 'tr')
        if not row_nodes:
            row_nodes = [table]
        filled = 0
        for row_index, values in enumerate(matrix):
            if row_index >= len(row_nodes):
                break
            row_node = row_nodes[row_index]
            cell_inputs = row_node.find_elements(By.CSS_SELECTOR, 'input, textarea, select, [contenteditable="true"]')
            if not cell_inputs:
                cell_inputs = row_node.find_elements(By.CSS_SELECTOR, 'td, th')
            for col_index, value in enumerate(values):
                if col_index >= len(cell_inputs):
                    break
                cell = cell_inputs[col_index]
                tag = (getattr(cell, 'tag_name', '') or '').lower()
                if tag in ('td', 'th'):
                    nested = cell.find_elements(By.CSS_SELECTOR, 'input, textarea, select, [contenteditable="true"]')
                    if nested:
                        cell = nested[0]
                self._scroll_into_view(cell)
                self._set_element_value(cell, value, clear_before_input=clear_before_input)
                filled += 1
        if filled <= 0:
            raise BrowserEngineError('未找到可填写的表格单元格')
        return True

    def _wait_until_back_to_main(self, timeout: float = 10):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        main_handle = self.main_window_handle
        def condition(driver):
            handles = list(driver.window_handles)
            if main_handle and main_handle in handles and len(handles) == 1:
                return True
            try:
                current = driver.current_window_handle
            except Exception:
                current = None
            return bool(main_handle and current == main_handle and main_handle in handles)
        WebDriverWait(self.driver, timeout).until(condition)
        self.switch_to_main_window()
        return True

    @staticmethod
    def _emit_alert(alert_handler, message, level='info'):
        if alert_handler:
            alert_handler(message, level)

    @staticmethod
    def _coerce_branch_step(value):
        try:
            num = int(float(value or 0))
        except Exception:
            num = 0
        return num if num > 0 else 0

    def _resolve_step_index(self, step_no: int, step_count: int):
        if step_no <= 0:
            return None
        if step_no > step_count:
            raise BrowserEngineError(f'跳转目标步骤不存在：{step_no}')
        return step_no - 1

    def _handle_branch_result(self, step: dict, result_key: str, step_count: int, logger=None, alert_handler=None):
        message_key = {
            'found': 'on_found_message',
            'not_found': 'on_not_found_message',
            'timeout': 'on_timeout_message',
        }[result_key]
        step_key = {
            'found': 'on_found_step',
            'not_found': 'on_not_found_step',
            'timeout': 'on_timeout_step',
        }[result_key]
        message = str(step.get(message_key, '') or '').strip()
        if message:
            self._log(logger, f'页面条件判断提示：{message}')
            self._emit_alert(alert_handler, message, 'timeout' if result_key == 'timeout' else 'info')
        target = self._coerce_branch_step(step.get(step_key, 0))
        return self._resolve_step_index(target, step_count)

    def _evaluate_page_condition(self, step: dict, payload: dict, timeout: float = 10):
        expr = str(step.get('page_condition_expr', '') or '').strip()
        if expr:
            result = self.evaluate_condition_expression(expr, payload)
            return 'found' if result else 'not_found'
        locator_type = step.get('locator_type', '')
        locator_value = step.get('locator_value', '')
        if not str(locator_value or '').strip():
            raise BrowserEngineError('页面条件判断未配置定位值，也未配置判断表达式')
        mode = str(step.get('detect_mode', '等待判断') or '等待判断').strip()
        try:
            elements = self._find_elements(locator_type, locator_value)
            if elements:
                return 'found'
        except Exception:
            pass
        if mode == '立即判断':
            return 'not_found'
        try:
            self.wait_for_element(locator_type, locator_value, timeout=timeout, clickable=bool(step.get('wait_clickable', False)))
            return 'found'
        except TimeoutException:
            return 'timeout'
        except Exception:
            return 'timeout'

    def start_element_recording(self, logger=None):
        if not self.is_connected():
            raise BrowserEngineError('浏览器未连接')
        self.repair_session_window(logger=logger, activate_preferred=True)
        script = r'''
(function () {
    function safeText(text) {
        return String(text || '').replace(/\s+/g, ' ').trim();
    }
    function textLiteral(text) {
        text = String(text || '');
        if (text.indexOf("'") < 0) return "'" + text + "'";
        if (text.indexOf('\"') < 0) return '\"' + text + '\"';
        return null;
    }
    function cssPath(el) {
        if (!el || el.nodeType !== 1) return '';
        var path = [];
        while (el && el.nodeType === 1) {
            var selector = (el.nodeName || '').toLowerCase();
            if (!selector) break;
            var nth = 1;
            var sib = el;
            while ((sib = sib.previousElementSibling)) {
                if ((sib.nodeName || '').toLowerCase() === selector) nth++;
            }
            selector += ':nth-of-type(' + nth + ')';
            path.unshift(selector);
            el = el.parentElement;
        }
        return path.join(' > ');
    }
    function xpath(el) {
        if (!el || el.nodeType !== 1) return '';
        if (el === el.ownerDocument.body) return '/html/body';
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
    function closestClickable(el) {
        var cur = el;
        while (cur && cur.nodeType === 1) {
            var tag = (cur.tagName || '').toLowerCase();
            var cls = String(cur.className || '');
            var role = cur.getAttribute ? (cur.getAttribute('role') || '') : '';
            if (tag === 'a' || tag === 'button' || tag === 'option' || tag === 'label' || tag === 'li' || tag === 'tr' || role === 'button' || cur.getAttribute('onclick') || /(^|\s)(x-btn|btn|trigger|arrow|picker)(\s|$)/i.test(cls)) {
                return cur;
            }
            cur = cur.parentElement;
        }
        return el;
    }
    function makeTextLocator(el, text) {
        var literal = textLiteral(text);
        if (!literal) return '';
        var tag = (el.tagName || '').toLowerCase();
        if (tag === 'a' || tag === 'button' || tag === 'li' || tag === 'tr' || tag === 'span' || tag === 'div') {
            return '//' + tag + '[contains(.,' + literal + ')]';
        }
        return '//*[contains(text(),' + literal + ')]';
    }
    function makeRowLocator(el, text) {
        var literal = textLiteral(text);
        if (!literal) return '';
        var cur = el;
        while (cur && cur.nodeType === 1) {
            var tag = (cur.tagName || '').toLowerCase();
            var cls = String(cur.className || '');
            if (tag === 'tr' && cls.indexOf('x-grid-row') >= 0) {
                return '//*[contains(text(),' + literal + ')]/ancestor::tr[contains(@class,"x-grid-row")][1]';
            }
            if (tag === 'li') {
                return '//*[contains(text(),' + literal + ')]/ancestor::li[1]';
            }
            cur = cur.parentElement;
        }
        return '';
    }
    function buildInfo(el, frameChain) {
        var clickable = closestClickable(el);
        var text = safeText(el.innerText || el.textContent || el.value || el.getAttribute('aria-label') || el.getAttribute('title') || '');
        var clickableText = safeText(clickable.innerText || clickable.textContent || clickable.value || clickable.getAttribute('aria-label') || clickable.getAttribute('title') || '');
        var info = {
            tag: ((el.tagName || '').toLowerCase()),
            text: text.slice(0, 120),
            id: el.id || '',
            name: (el.getAttribute && el.getAttribute('name')) || '',
            placeholder: (el.getAttribute && el.getAttribute('placeholder')) || '',
            title: (el.getAttribute && (el.getAttribute('title') || el.getAttribute('aria-label'))) || '',
            css: cssPath(el),
            xpath: xpath(el),
            clickable_tag: ((clickable.tagName || '').toLowerCase()),
            clickable_text: clickableText.slice(0, 120),
            clickable_css: cssPath(clickable),
            clickable_xpath: xpath(clickable),
            frame_chain: frameChain || []
        };
        var candidates = [];
        function addCandidate(type, value, label) {
            type = String(type || '').trim();
            value = String(value || '').trim();
            if (!type || !value) return;
            for (var i = 0; i < candidates.length; i++) {
                if (candidates[i].type === type && candidates[i].value === value) return;
            }
            candidates.push({type: type, value: value, label: label || ''});
        }
        if (info.name) {
            addCandidate('name', info.name, 'name 属性，通常比动态 id 稳定');
        }
        if (info.placeholder) {
            var placeholderLiteral = textLiteral(info.placeholder);
            if (placeholderLiteral) {
                addCandidate('xpath', '//*[' + '@placeholder=' + placeholderLiteral + ']', 'placeholder XPath');
            }
        }
        if (info.title) {
            var titleLiteral = textLiteral(info.title);
            if (titleLiteral) {
                addCandidate('xpath', '//*[@title=' + titleLiteral + ']', 'title XPath');
                addCandidate('xpath', '//*[@aria-label=' + titleLiteral + ']', 'aria-label XPath');
            }
        }
        if (clickableText) {
            addCandidate('xpath', makeRowLocator(clickable, clickableText), '可点击行/列表 XPath');
            addCandidate('xpath', makeTextLocator(clickable, clickableText), '可点击文本 XPath');
            if ((clickable.tagName || '').toLowerCase() === 'a') {
                addCandidate('link text', clickableText, '链接完整文本');
                addCandidate('partial link text', clickableText, '链接部分文本');
            }
        }
        if (text && text !== clickableText) {
            addCandidate('xpath', makeRowLocator(el, text), '文本所在行/列表 XPath');
            addCandidate('xpath', makeTextLocator(el, text), '元素文本 XPath');
        }
        addCandidate('xpath', info.clickable_xpath, '可点击元素绝对 XPath');
        addCandidate('xpath', info.xpath, '当前元素绝对 XPath');
        addCandidate('css selector', info.clickable_css, '可点击元素 CSS');
        addCandidate('css selector', info.css, '当前元素 CSS');
        if (info.id) {
            addCandidate('id', info.id, 'id 属性，可能随刷新变化，建议确认后再用');
            addCandidate('xpath', '//*[@id=' + '"' + info.id + '"' + ']', 'id XPath，可能随刷新变化');
        }
        info.locator_candidates = candidates;
        info.recommended_locator_type = candidates.length ? candidates[0].type : '';
        info.recommended_locator_value = candidates.length ? candidates[0].value : '';
        return info;
    }
    function attach(doc, frameChain, handlers) {
        if (!doc || doc.__peRecorderBound) return;
        doc.__peRecorderBound = true;
        var handler = function (e) {
            e = e || window.event;
            var target = e.target || e.srcElement;
            try {
                window.top.__peRecorderResult = buildInfo(target, frameChain || []);
            } catch (err) {
                window.top.__peRecorderResult = {error: String((err && err.message) || err || 'unknown')};
            }
            if (e.preventDefault) e.preventDefault();
            if (e.stopPropagation) e.stopPropagation();
            e.cancelBubble = true;
            return false;
        };
        if (doc.addEventListener) doc.addEventListener('click', handler, true);
        else if (doc.attachEvent) doc.attachEvent('onclick', handler);
        handlers.push({doc: doc, handler: handler});
        var iframes = doc.getElementsByTagName('iframe');
        for (var i = 0; i < iframes.length; i++) {
            try {
                attach(iframes[i].contentWindow.document, (frameChain || []).concat([xpath(iframes[i]) || cssPath(iframes[i]) || ('iframe[' + i + ']')]), handlers);
            } catch (err) {}
        }
    }
    if (window.__peRecorderCleanup) {
        try { window.__peRecorderCleanup(); } catch (err) {}
    }
    window.__peRecorderResult = null;
    window.__peRecorderActive = true;
    var handlers = [];
    attach(document, [], handlers);
    window.__peRecorderCleanup = function () {
        for (var i = 0; i < handlers.length; i++) {
            try {
                if (handlers[i].doc.removeEventListener) handlers[i].doc.removeEventListener('click', handlers[i].handler, true);
            } catch (err) {}
            try { handlers[i].doc.__peRecorderBound = false; } catch (err) {}
        }
        window.__peRecorderActive = false;
    };
    return true;
})();
'''
        self.driver.execute_script(script)
        self._log(logger, '自动录制已启动，请切换到浏览器点击目标元素。')
        return True

    def poll_recorded_element(self, consume: bool = True):
        if not self.is_connected():
            return None
        self.repair_session_window(activate_preferred=False)
        info = self.driver.execute_script('return window.top.__peRecorderResult || null;')
        if info and consume:
            try:
                self.driver.execute_script('if (window.top.__peRecorderCleanup) { try { window.top.__peRecorderCleanup(); } catch (e) {} } window.top.__peRecorderResult = null;')
            except Exception:
                pass
        return info

    def stop_element_recording(self, logger=None):
        if not self.is_connected():
            return False
        try:
            self.driver.execute_script('if (window.top.__peRecorderCleanup) { try { window.top.__peRecorderCleanup(); } catch (e) {} } window.top.__peRecorderResult = null;')
            self._log(logger, '已停止自动录制。')
            return True
        except Exception:
            return False

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

    @staticmethod
    def _validate_logic_function_args(func_name: str, parts):
        if not parts:
            raise BrowserEngineError(f'{func_name} 函数至少需要 1 个参数')
        blank_indexes = [str(i + 1) for i, part in enumerate(parts) if not str(part).strip()]
        if blank_indexes:
            joined = '、'.join(blank_indexes)
            raise BrowserEngineError(f'{func_name} 函数第 {joined} 个参数不能为空')
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
            'true': True,
            'false': False,
            'TRUE': True,
            'FALSE': False,
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

    def execute_flow(self, flow_config: dict, payload: dict, logger=None, alert_handler=None):
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

        idx = 0
        while idx < len(steps):
            step = steps[idx] or {}
            action = (step.get('action') or '').strip()
            name = (step.get('name') or f'步骤{idx + 1}').strip()
            timeout = float(step.get('wait_timeout', 10) or 10)
            self._log(logger, f'[{idx + 1}] {name} - {action}')

            if not self.evaluate_condition_expression(step.get('condition_expr', ''), payload):
                self._log(logger, f'[{idx + 1}] 已跳过：未满足执行条件')
                idx += 1
                continue

            next_index = idx + 1

            if action == '点击元素':
                self._click_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout, use_js_click=bool(step.get('use_js_click')))
                context['last_window_count'] = len(self.driver.window_handles)

            elif action == '输入文本':
                value = self.render_value_template(step.get('value_template', ''), payload)
                value = self.transform_input_value(value, step)
                self._input_text(step.get('locator_type', ''), step.get('locator_value', ''), value, timeout=timeout, clear_before_input=bool(step.get('clear_before_input', True)))

            elif action == '多选元素/多行选择':
                rendered = self.render_value_template(step.get('value_template', ''), payload)
                locator_values = [line.strip() for line in str(rendered or '').splitlines() if line.strip()]
                if not locator_values and str(step.get('locator_value', '') or '').strip():
                    locator_values = [str(step.get('locator_value', '')).strip()]
                self._multi_select(step.get('locator_type', ''), locator_values, timeout=timeout)

            elif action == '右键点击':
                self._right_click_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout)

            elif action == '键盘组合键':
                combo_text = self.render_value_template(step.get('value_template', ''), payload)
                self._send_key_combo(step.get('locator_type', ''), step.get('locator_value', ''), combo_text, timeout=timeout)

            elif action == '右键菜单项点击':
                self._right_click_menu_item(step, timeout=timeout)

            elif action == '下拉菜单两段式操作':
                self._dropdown_two_stage(step, timeout=timeout)

            elif action == '拖拽元素':
                self._drag_drop_element(step, timeout=timeout)
                context['last_window_count'] = len(self.driver.window_handles)

            elif action == '等待元素':
                self.wait_for_element(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout, clickable=bool(step.get('wait_clickable', False)))

            elif action == '等待元素消失':
                self.wait_for_element_gone(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout)

            elif action == '等待新窗口':
                old_count = context.get('last_window_count', len(self.driver.window_handles))
                WebDriverWait(self.driver, timeout).until(lambda d: len(d.window_handles) > old_count)
                context['last_window_count'] = len(self.driver.window_handles)

            elif action == '等待窗口关闭/等待回到主窗口':
                self._wait_until_back_to_main(timeout=timeout)
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

            elif action == '页面条件判断':
                result_key = self._evaluate_page_condition(step, payload, timeout=timeout)
                self._log(logger, f'页面条件判断结果：{result_key}')
                target_index = self._handle_branch_result(step, result_key, len(steps), logger=logger, alert_handler=alert_handler)
                if target_index is not None:
                    next_index = target_index

            elif action == '添加表格':
                self._add_table(step.get('locator_type', ''), step.get('locator_value', ''), timeout=timeout)

            elif action == '自动填单元格':
                table_text = self.render_value_template(step.get('value_template', ''), payload)
                self._fill_table_cells(step.get('locator_type', ''), step.get('locator_value', ''), table_text, timeout=timeout, clear_before_input=bool(step.get('clear_before_input', True)))

            else:
                raise BrowserEngineError(f'不支持的动作：{action}')

            idx = next_index

        return True
