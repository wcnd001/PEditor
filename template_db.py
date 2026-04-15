import sqlite3
import json
import os
from typing import List, Dict, Any, Optional

try:
    from utils import resource_path
except ImportError:
    import sys

    def resource_path(relative_path):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath('.'), relative_path)


DB_PATH = resource_path('template.db')


class TemplateDB:
    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self.connect()
        self.init_db()

    def connect(self):
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.cursor = self.conn.cursor()

    def ensure_connection(self):
        try:
            self.conn.cursor()
        except (sqlite3.ProgrammingError, AttributeError):
            self.connect()

    def close(self):
        if self.conn:
            self.conn.close()
            self.conn = None
            self.cursor = None

    def init_db(self):
        self.ensure_connection()
        self.cursor.execute(
            '''
            CREATE TABLE IF NOT EXISTS main_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                config TEXT NOT NULL
            )
            '''
        )
        self.cursor.execute(
            '''
            CREATE TABLE IF NOT EXISTS process_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                content TEXT NOT NULL
            )
            '''
        )
        self.cursor.execute(
            '''
            CREATE TABLE IF NOT EXISTS browser_flows (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_name TEXT UNIQUE NOT NULL,
                flow_config TEXT NOT NULL,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            '''
        )
        self.conn.commit()

    @staticmethod
    def _loads_json(text: str, default):
        try:
            return json.loads(text)
        except Exception:
            return default

    def get_main_templates(self) -> List[Dict[str, Any]]:
        self.ensure_connection()
        self.cursor.execute('SELECT id, name, config FROM main_templates ORDER BY name')
        rows = self.cursor.fetchall()
        return [{'id': r[0], 'name': r[1], 'config': self._loads_json(r[2], {})} for r in rows]

    def get_main_template(self, name: str) -> Optional[Dict[str, Any]]:
        self.ensure_connection()
        self.cursor.execute('SELECT id, name, config FROM main_templates WHERE name = ?', (name,))
        row = self.cursor.fetchone()
        if row:
            return {'id': row[0], 'name': row[1], 'config': self._loads_json(row[2], {})}
        return None

    def add_main_template(self, name: str, config: dict):
        self.ensure_connection()
        self.cursor.execute('INSERT INTO main_templates (name, config) VALUES (?, ?)', (name, json.dumps(config, ensure_ascii=False)))
        self.conn.commit()

    def update_main_template(self, old_name: str, new_name: str, config: dict):
        self.ensure_connection()
        self.cursor.execute('UPDATE main_templates SET name = ?, config = ? WHERE name = ?', (new_name, json.dumps(config, ensure_ascii=False), old_name))
        self.conn.commit()
        if old_name != new_name:
            self.rename_browser_flow(old_name, new_name)

    def delete_main_template(self, name: str):
        self.ensure_connection()
        self.cursor.execute('DELETE FROM main_templates WHERE name = ?', (name,))
        self.conn.commit()
        self.delete_browser_flow(name)

    def get_process_templates(self) -> List[Dict[str, Any]]:
        self.ensure_connection()
        self.cursor.execute('SELECT id, name, content FROM process_templates ORDER BY name')
        rows = self.cursor.fetchall()
        return [{'id': r[0], 'name': r[1], 'content': self._loads_json(r[2], {})} for r in rows]

    def get_process_template(self, name: str) -> Optional[Dict[str, Any]]:
        self.ensure_connection()
        self.cursor.execute('SELECT id, name, content FROM process_templates WHERE name = ?', (name,))
        row = self.cursor.fetchone()
        if row:
            return {'id': row[0], 'name': row[1], 'content': self._loads_json(row[2], {})}
        return None

    def add_process_template(self, name: str, content: dict):
        self.ensure_connection()
        self.cursor.execute('INSERT INTO process_templates (name, content) VALUES (?, ?)', (name, json.dumps(content, ensure_ascii=False)))
        self.conn.commit()

    def update_process_template(self, name: str, content: dict):
        self.ensure_connection()
        exists = self.get_process_template(name)
        content_text = json.dumps(content, ensure_ascii=False)
        if exists:
            self.cursor.execute('UPDATE process_templates SET content = ? WHERE name = ?', (content_text, name))
        else:
            self.cursor.execute('INSERT INTO process_templates (name, content) VALUES (?, ?)', (name, content_text))
        self.conn.commit()

    def delete_process_template(self, name: str):
        self.ensure_connection()
        self.cursor.execute('DELETE FROM process_templates WHERE name = ?', (name,))
        self.conn.commit()

    def get_browser_flow(self, template_name: str) -> Optional[dict]:
        self.ensure_connection()
        self.cursor.execute('SELECT flow_config FROM browser_flows WHERE template_name = ?', (template_name,))
        row = self.cursor.fetchone()
        if row:
            return self._loads_json(row[0], {})
        return None

    def update_browser_flow(self, template_name: str, flow_config: dict):
        self.ensure_connection()
        text = json.dumps(flow_config or {}, ensure_ascii=False)
        exists = self.get_browser_flow(template_name)
        if exists is None:
            self.cursor.execute('INSERT INTO browser_flows (template_name, flow_config, updated_at) VALUES (?, ?, CURRENT_TIMESTAMP)', (template_name, text))
        else:
            self.cursor.execute('UPDATE browser_flows SET flow_config = ?, updated_at = CURRENT_TIMESTAMP WHERE template_name = ?', (text, template_name))
        self.conn.commit()

    def delete_browser_flow(self, template_name: str):
        self.ensure_connection()
        self.cursor.execute('DELETE FROM browser_flows WHERE template_name = ?', (template_name,))
        self.conn.commit()

    def rename_browser_flow(self, old_name: str, new_name: str):
        self.ensure_connection()
        self.cursor.execute('UPDATE browser_flows SET template_name = ?, updated_at = CURRENT_TIMESTAMP WHERE template_name = ?', (new_name, old_name))
        self.conn.commit()
