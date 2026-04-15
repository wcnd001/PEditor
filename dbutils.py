import sqlite3
from typing import List, Optional
from utils import resource_path

DB_PATH = resource_path("data.db")


class Database:
    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self._manual_transaction = False
        self.connect()

    def connect(self):
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.cursor = self.conn.cursor()
        self._manual_transaction = False

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
            self._manual_transaction = False

    def execute(self, sql: str, params: tuple = ()):
        self.ensure_connection()
        self.cursor.execute(sql, params)
        if not self._manual_transaction:
            self.conn.commit()
        return self.cursor

    def executemany(self, sql: str, params_list: List[tuple]):
        self.ensure_connection()
        self.cursor.executemany(sql, params_list)
        if not self._manual_transaction:
            self.conn.commit()
        return self.cursor

    def fetch_all(self, sql: str, params: tuple = ()) -> List[dict]:
        self.ensure_connection()
        self.cursor.execute(sql, params)
        return [dict(row) for row in self.cursor.fetchall()]

    def fetch_one(self, sql: str, params: tuple = ()) -> Optional[dict]:
        self.ensure_connection()
        self.cursor.execute(sql, params)
        row = self.cursor.fetchone()
        return dict(row) if row else None

    def get_tables(self) -> List[str]:
        self.ensure_connection()
        self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = [row[0] for row in self.cursor.fetchall()]
        return [t for t in tables if not t.startswith('sqlite_')]

    def get_table_info(self, table_name: str) -> List[dict]:
        self.ensure_connection()
        safe_name = table_name.replace('"', '""')
        self.cursor.execute(f'PRAGMA table_info("{safe_name}")')
        return [dict(row) for row in self.cursor.fetchall()]

    def begin_transaction(self):
        self.ensure_connection()
        if self._manual_transaction:
            return
        self.conn.execute("BEGIN TRANSACTION")
        self._manual_transaction = True

    def commit_transaction(self):
        self.ensure_connection()
        if self._manual_transaction:
            self.conn.commit()
            self._manual_transaction = False

    def rollback_transaction(self):
        self.ensure_connection()
        if self._manual_transaction:
            self.conn.rollback()
            self._manual_transaction = False
