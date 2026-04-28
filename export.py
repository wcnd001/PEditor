from webengine import BrowserEngine

_engine = BrowserEngine()


def get_engine():
    return _engine


def export_data(data_pool: dict):
    print('=== 导出数据 ===', flush=True)
    for field, value in (data_pool or {}).items():
        print(f'{field}: {value}', flush=True)
    print('=== 导出完成 ===', flush=True)


def export_to_browser(template_db, template_name: str, result_text: str, final_fields: dict, input_values: dict = None, data_pool: dict = None, flow_override: dict = None, alert_handler=None):
    flow = flow_override or template_db.get_browser_flow(template_name)
    if not flow:
        return False, '当前模板未配置浏览器流程。'

    logs = []

    def logger(message):
        text = str(message)
        logs.append(text)
        print(text, flush=True)

    payload = {
        'template_name': template_name,
        'result_text': result_text,
        'final_fields': final_fields or {},
        'input_values': input_values or {},
        'data_pool': data_pool or {},
    }

    try:
        _engine.execute_flow(flow, payload, logger=logger, alert_handler=alert_handler)
        logger('浏览器导出执行完成。')
        return True, '\n'.join(logs)
    except Exception as e:
        logger(f'浏览器导出失败：{e}')
        return False, '\n'.join(logs)
