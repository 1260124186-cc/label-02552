# -*- coding: utf-8 -*-
"""
历史批次与版本管理模块
功能：
  1. 每次运行生成唯一批次号
  2. 双存储架构：文件系统按日期归档 + SQLite 元数据索引
  3. 归档内容：总表、日志、检验报告、元数据
  4. 支持按时间、批次号、记录数等条件查询回溯
"""

import os
import sys
import json
import sqlite3
import shutil
import random
import string
import logging
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import List, Optional, Dict, Any, Tuple
from pathlib import Path


BATCH_DB_FILENAME = 'batch_history.db'
HISTORY_DIR_NAME = 'history'
BATCH_ID_PREFIX = 'BATCH'


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    return logging.getLogger('bankcheck')


def generate_batch_id() -> str:
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    random_suffix = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    return f"{BATCH_ID_PREFIX}{timestamp}{random_suffix}"


def get_date_str(dt: Optional[datetime] = None) -> str:
    if dt is None:
        dt = datetime.now()
    return dt.strftime('%Y-%m-%d')


@dataclass
class BatchInfo:
    batch_id: str
    start_time: str
    end_time: Optional[str] = None
    status: str = 'running'
    input_folder: str = ''
    output_folder: str = ''
    total_records: int = 0
    new_records: int = 0
    duplicate_records: int = 0
    processed_files: int = 0
    unprocessed_files: int = 0
    error_files: int = 0
    incremental_mode: bool = False
    operator: str = ''
    summary_table_path: str = ''
    log_file_path: str = ''
    report_path: str = ''
    batch_dir: str = ''
    remark: str = ''
    error_message: str = ''

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'BatchInfo':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


class BatchManager:
    def __init__(self, script_dir: Optional[str] = None):
        self.script_dir = script_dir or get_script_dir()
        self.db_path = os.path.join(self.script_dir, BATCH_DB_FILENAME)
        self.history_root = os.path.join(self.script_dir, HISTORY_DIR_NAME)
        self.current_batch: Optional[BatchInfo] = None
        self._init_db()
        self._ensure_history_dir()

    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS batches (
                batch_id TEXT PRIMARY KEY,
                start_time TEXT NOT NULL,
                end_time TEXT,
                status TEXT NOT NULL,
                input_folder TEXT,
                output_folder TEXT,
                total_records INTEGER DEFAULT 0,
                new_records INTEGER DEFAULT 0,
                duplicate_records INTEGER DEFAULT 0,
                processed_files INTEGER DEFAULT 0,
                unprocessed_files INTEGER DEFAULT 0,
                error_files INTEGER DEFAULT 0,
                incremental_mode INTEGER DEFAULT 0,
                operator TEXT,
                summary_table_path TEXT,
                log_file_path TEXT,
                report_path TEXT,
                batch_dir TEXT,
                remark TEXT,
                error_message TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_batches_start_time ON batches(start_time)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_batches_status ON batches(status)')
        conn.commit()
        conn.close()

    def _ensure_history_dir(self):
        os.makedirs(self.history_root, exist_ok=True)

    def _get_batch_dir(self, batch_id: str, date_str: str) -> str:
        return os.path.join(self.history_root, date_str, batch_id)

    def start_batch(self, input_folder: str = '', operator: str = '') -> BatchInfo:
        batch_id = generate_batch_id()
        date_str = get_date_str()
        batch_dir = self._get_batch_dir(batch_id, date_str)
        os.makedirs(batch_dir, exist_ok=True)

        batch_info = BatchInfo(
            batch_id=batch_id,
            start_time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            input_folder=input_folder,
            operator=operator,
            batch_dir=batch_dir,
            status='running',
        )

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO batches (
                batch_id, start_time, status, input_folder, operator, batch_dir
            ) VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            batch_info.batch_id,
            batch_info.start_time,
            batch_info.status,
            batch_info.input_folder,
            batch_info.operator,
            batch_info.batch_dir,
        ))
        conn.commit()
        conn.close()

        self.current_batch = batch_info
        logger = get_logger()
        logger.info('批次已创建: %s', batch_id)
        logger.info('归档目录: %s', batch_dir)

        return batch_info

    def _update_batch(self, batch_info: BatchInfo):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE batches SET
                end_time = ?,
                status = ?,
                input_folder = ?,
                output_folder = ?,
                total_records = ?,
                new_records = ?,
                duplicate_records = ?,
                processed_files = ?,
                unprocessed_files = ?,
                error_files = ?,
                incremental_mode = ?,
                operator = ?,
                summary_table_path = ?,
                log_file_path = ?,
                report_path = ?,
                batch_dir = ?,
                remark = ?,
                error_message = ?
            WHERE batch_id = ?
        ''', (
            batch_info.end_time,
            batch_info.status,
            batch_info.input_folder,
            batch_info.output_folder,
            batch_info.total_records,
            batch_info.new_records,
            batch_info.duplicate_records,
            batch_info.processed_files,
            batch_info.unprocessed_files,
            batch_info.error_files,
            1 if batch_info.incremental_mode else 0,
            batch_info.operator,
            batch_info.summary_table_path,
            batch_info.log_file_path,
            batch_info.report_path,
            batch_info.batch_dir,
            batch_info.remark,
            batch_info.error_message,
            batch_info.batch_id,
        ))
        conn.commit()
        conn.close()

    def archive_summary_table(self, batch_id: str, source_path: str) -> str:
        batch_info = self._get_batch_info(batch_id)
        if not batch_info or not os.path.exists(source_path):
            return ''

        filename = os.path.basename(source_path)
        dest_path = os.path.join(batch_info.batch_dir, f"{batch_id}_{filename}")
        shutil.copy2(source_path, dest_path)

        batch_info.summary_table_path = dest_path
        self._update_batch(batch_info)

        logger = get_logger()
        logger.info('总表已归档: %s', dest_path)
        return dest_path

    def archive_log_file(self, batch_id: str, source_path: str) -> str:
        batch_info = self._get_batch_info(batch_id)
        if not batch_info or not os.path.exists(source_path):
            return ''

        filename = os.path.basename(source_path)
        dest_path = os.path.join(batch_info.batch_dir, f"{batch_id}_{filename}")
        shutil.copy2(source_path, dest_path)

        batch_info.log_file_path = dest_path
        self._update_batch(batch_info)

        logger = get_logger()
        logger.info('日志已归档: %s', dest_path)
        return dest_path

    def generate_inspection_report(self, batch_id: str, result_data: Dict[str, Any]) -> str:
        batch_info = self._get_batch_info(batch_id)
        if not batch_info:
            return ''

        report_content = self._build_report_content(batch_info, result_data)
        report_path = os.path.join(batch_info.batch_dir, f"{batch_id}_检验报告.md")

        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_content)

        batch_info.report_path = report_path
        self._update_batch(batch_info)

        logger = get_logger()
        logger.info('检验报告已生成: %s', report_path)
        return report_path

    def _build_report_content(self, batch_info: BatchInfo, result_data: Dict[str, Any]) -> str:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        status_icon = '✅' if batch_info.status == 'success' else '⚠️' if batch_info.status == 'warning' else '❌'

        content = f"""# 银行流水检验报告

**批次号**: {batch_info.batch_id}
**状态**: {status_icon} {batch_info.status}
**开始时间**: {batch_info.start_time}
**结束时间**: {batch_info.end_time or now}
**操作员**: {batch_info.operator or '未指定'}
**输入文件夹**: {batch_info.input_folder}

---

## 一、处理统计

| 指标 | 数值 |
|------|------|
| 处理文件数 | {batch_info.processed_files} |
| 未识别文件数 | {batch_info.unprocessed_files} |
| 处理出错文件数 | {batch_info.error_files} |
| 运行模式 | {'增量合并' if batch_info.incremental_mode else '全量覆盖'} |
| 总记录数 | {batch_info.total_records} |
| 新增记录数 | {batch_info.new_records} |
| 重复记录数 | {batch_info.duplicate_records} |

---

## 二、文件清单

### 已处理文件
"""

        processed_files = result_data.get('processed_files', [])
        if processed_files:
            for f in processed_files:
                content += f"- {os.path.basename(f)}\n"
        else:
            content += "无\n"

        unprocessed_files = result_data.get('unprocessed_files', [])
        if unprocessed_files:
            content += "\n### 未识别文件\n"
            for f in unprocessed_files:
                content += f"- {os.path.basename(f)}\n"

        error_files = result_data.get('error_files', [])
        if error_files:
            content += "\n### 处理出错文件\n"
            for f, err in error_files:
                content += f"- {os.path.basename(f)}: {err}\n"

        content += f"""
---

## 三、归档文件

- **总表**: [{os.path.basename(batch_info.summary_table_path)}]({batch_info.summary_table_path})
- **日志**: [{os.path.basename(batch_info.log_file_path)}]({batch_info.log_file_path})
- **归档目录**: {batch_info.batch_dir}

"""

        if batch_info.error_message:
            content += f"""
---

## 四、错误信息

```
{batch_info.error_message}
```
"""

        if batch_info.remark:
            content += f"""
---

## 五、备注

{batch_info.remark}
"""

        content += f"\n---\n*报告生成时间: {now}*\n"
        return content

    def finish_batch(self, batch_id: str, result_data: Dict[str, Any],
                     status: str = 'success', error_message: str = '') -> BatchInfo:
        batch_info = self._get_batch_info(batch_id)
        if not batch_info:
            raise ValueError(f"批次不存在: {batch_id}")

        batch_info.end_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        batch_info.status = status
        batch_info.total_records = result_data.get('total_records', 0)
        batch_info.new_records = result_data.get('new_records', 0)
        batch_info.duplicate_records = result_data.get('duplicate_records', 0)
        batch_info.processed_files = len(result_data.get('processed_files', []))
        batch_info.unprocessed_files = len(result_data.get('unprocessed_files', []))
        batch_info.error_files = len(result_data.get('error_files', []))
        batch_info.incremental_mode = result_data.get('incremental_mode', False)
        batch_info.output_folder = result_data.get('output_folder', '')
        batch_info.error_message = error_message

        if result_data.get('summary_table_path'):
            self.archive_summary_table(batch_id, result_data['summary_table_path'])

        if result_data.get('log_file_path'):
            self.archive_log_file(batch_id, result_data['log_file_path'])

        self.generate_inspection_report(batch_id, result_data)

        self._save_metadata(batch_info, result_data)
        self._update_batch(batch_info)

        logger = get_logger()
        logger.info('批次完成: %s, 状态: %s', batch_id, status)

        return batch_info

    def _save_metadata(self, batch_info: BatchInfo, result_data: Dict[str, Any]):
        metadata = {
            'batch_info': batch_info.to_dict(),
            'result_data': result_data,
            'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
        metadata_path = os.path.join(batch_info.batch_dir, f"{batch_info.batch_id}_metadata.json")
        with open(metadata_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)

    def _get_batch_info(self, batch_id: str) -> Optional[BatchInfo]:
        if self.current_batch and self.current_batch.batch_id == batch_id:
            return self.current_batch

        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM batches WHERE batch_id = ?', (batch_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            data = dict(row)
            data['incremental_mode'] = bool(data.get('incremental_mode', 0))
            return BatchInfo.from_dict(data)
        return None

    def query_batches(self,
                      start_date: Optional[str] = None,
                      end_date: Optional[str] = None,
                      status: Optional[str] = None,
                      operator: Optional[str] = None,
                      min_records: Optional[int] = None,
                      limit: int = 100,
                      offset: int = 0) -> List[BatchInfo]:
        query = 'SELECT * FROM batches WHERE 1=1'
        params = []

        if start_date:
            query += ' AND date(start_time) >= date(?)'
            params.append(start_date)
        if end_date:
            query += ' AND date(start_time) <= date(?)'
            params.append(end_date)
        if status:
            query += ' AND status = ?'
            params.append(status)
        if operator:
            query += ' AND operator LIKE ?'
            params.append(f'%{operator}%')
        if min_records:
            query += ' AND total_records >= ?'
            params.append(min_records)

        query += ' ORDER BY start_time DESC LIMIT ? OFFSET ?'
        params.extend([limit, offset])

        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()

        batches = []
        for row in rows:
            data = dict(row)
            data['incremental_mode'] = bool(data.get('incremental_mode', 0))
            batches.append(BatchInfo.from_dict(data))
        return batches

    def get_batch_detail(self, batch_id: str) -> Optional[Dict[str, Any]]:
        batch_info = self._get_batch_info(batch_id)
        if not batch_info:
            return None

        metadata_path = os.path.join(batch_info.batch_dir, f"{batch_id}_metadata.json")
        metadata = {}
        if os.path.exists(metadata_path):
            with open(metadata_path, 'r', encoding='utf-8') as f:
                metadata = json.load(f)

        return {
            'batch_info': batch_info.to_dict(),
            'metadata': metadata.get('result_data', {}),
            'files': self._list_batch_files(batch_info.batch_dir),
        }

    def _list_batch_files(self, batch_dir: str) -> Dict[str, str]:
        files = {}
        if not os.path.exists(batch_dir):
            return files
        for f in os.listdir(batch_dir):
            fpath = os.path.join(batch_dir, f)
            if os.path.isfile(fpath):
                files[f] = fpath
        return files

    def restore_batch(self, batch_id: str, target_dir: Optional[str] = None) -> Dict[str, str]:
        batch_info = self._get_batch_info(batch_id)
        if not batch_info:
            raise ValueError(f"批次不存在: {batch_id}")

        if target_dir is None:
            target_dir = os.path.join(self.script_dir, f"restore_{batch_id}")
        os.makedirs(target_dir, exist_ok=True)

        restored = {}
        for name, src in self._list_batch_files(batch_info.batch_dir).items():
            dest = os.path.join(target_dir, name)
            shutil.copy2(src, dest)
            restored[name] = dest

        logger = get_logger()
        logger.info('批次 %s 已恢复到: %s', batch_id, target_dir)
        return restored

    def delete_batch(self, batch_id: str) -> bool:
        batch_info = self._get_batch_info(batch_id)
        if not batch_info:
            return False

        if os.path.exists(batch_info.batch_dir):
            shutil.rmtree(batch_info.batch_dir)

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM batches WHERE batch_id = ?', (batch_id,))
        conn.commit()
        conn.close()

        if self.current_batch and self.current_batch.batch_id == batch_id:
            self.current_batch = None

        logger = get_logger()
        logger.info('批次已删除: %s', batch_id)
        return True

    def get_statistics(self) -> Dict[str, Any]:
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT COUNT(*) FROM batches')
        total_batches = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*) FROM batches WHERE status = "success"')
        success_batches = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*) FROM batches WHERE status = "running"')
        running_batches = cursor.fetchone()[0]

        cursor.execute('SELECT IFNULL(SUM(total_records), 0) FROM batches')
        total_records = cursor.fetchone()[0]

        cursor.execute('SELECT IFNULL(SUM(new_records), 0) FROM batches')
        total_new_records = cursor.fetchone()[0]

        cursor.execute('''
            SELECT date(start_time) as dt, COUNT(*) as cnt
            FROM batches
            WHERE start_time >= date('now', '-30 days')
            GROUP BY date(start_time)
            ORDER BY dt DESC
            LIMIT 30
        ''')
        daily_stats = [{'date': r[0], 'count': r[1]} for r in cursor.fetchall()]

        conn.close()

        return {
            'total_batches': total_batches,
            'success_batches': success_batches,
            'failed_batches': total_batches - success_batches - running_batches,
            'running_batches': running_batches,
            'total_records': total_records,
            'total_new_records': total_new_records,
            'daily_stats': daily_stats,
        }

    def get_batch_ids_by_date(self, date_str: str) -> List[str]:
        date_dir = os.path.join(self.history_root, date_str)
        if not os.path.exists(date_dir):
            return []
        return [d for d in os.listdir(date_dir) if d.startswith(BATCH_ID_PREFIX)]


_global_batch_manager: Optional[BatchManager] = None


def get_batch_manager(script_dir: Optional[str] = None) -> BatchManager:
    global _global_batch_manager
    if _global_batch_manager is None:
        _global_batch_manager = BatchManager(script_dir)
    return _global_batch_manager
