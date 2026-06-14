# -*- coding: utf-8 -*-
"""
异步任务数据模型
定义任务状态、类型、优先级以及各阶段任务的负载数据结构
"""

import uuid
import json
import enum
import logging
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional, Tuple, Union
from datetime import datetime


logger = logging.getLogger('bankcheck.task_queue')


class TaskStatus(str, enum.Enum):
    """任务状态枚举"""
    PENDING = 'pending'
    QUEUED = 'queued'
    RUNNING = 'running'
    COMPLETED = 'completed'
    FAILED = 'failed'
    CANCELLED = 'cancelled'
    RETRYING = 'retrying'
    TIMEOUT = 'timeout'


class TaskType(str, enum.Enum):
    """任务类型枚举"""
    SCAN = 'scan'
    PARSE = 'parse'
    MERGE = 'merge'
    REPORT = 'report'
    PERSIST = 'persist'
    CLEANUP = 'cleanup'


class TaskPriority(int, enum.Enum):
    """任务优先级枚举"""
    LOW = 1
    NORMAL = 5
    HIGH = 8
    CRITICAL = 10


def generate_task_id() -> str:
    """生成唯一任务ID"""
    return f"task_{uuid.uuid4().hex[:24]}"


def generate_job_id() -> str:
    """生成唯一作业ID"""
    return f"job_{uuid.uuid4().hex[:24]}"


@dataclass
class TaskPayload:
    """基础任务负载"""
    task_type: str
    job_id: str = ''
    task_id: str = field(default_factory=generate_task_id)
    priority: int = TaskPriority.NORMAL.value
    retry_count: int = 0
    max_retries: int = 3
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'TaskPayload':
        if cls is TaskPayload:
            task_type = data.get('task_type')
            subclass_map = {
                TaskType.SCAN.value: ScanTaskPayload,
                TaskType.PARSE.value: ParseTaskPayload,
                TaskType.MERGE.value: MergeTaskPayload,
                TaskType.REPORT.value: ReportTaskPayload,
                TaskType.PERSIST.value: PersistTaskPayload,
                TaskType.CLEANUP.value: CleanupTaskPayload,
            }
            subclass = subclass_map.get(task_type, cls)
            return subclass(**{k: v for k, v in data.items()
                               if k in subclass.__dataclass_fields__})
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False)

    @classmethod
    def from_json(cls, json_str: str) -> 'TaskPayload':
        return cls.from_dict(json.loads(json_str))


@dataclass
class JobContext:
    """作业上下文，用于跟踪整个处理流程"""
    job_id: str
    user_id: str = ''
    operator: str = ''
    input_folder: str = ''
    script_dir: str = ''
    incremental: bool = True
    keep_strategy: str = 'keep_unprocessed'
    folder_strategy: str = 'copy_sibling'
    folder_output_dir: Optional[str] = None
    folder_suffix: str = '＋检验版'
    status: str = TaskStatus.PENDING.value
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    started_at: Optional[str] = None
    finished_at: Optional[str] = None
    total_files: int = 0
    processed_files: int = 0
    error_files: int = 0
    unprocessed_files: int = 0
    total_records: int = 0
    new_records: int = 0
    duplicate_records: int = 0
    error_message: Optional[str] = None
    batch_id: Optional[str] = None
    output_path: Optional[str] = None
    subject_summary_path: Optional[str] = None
    balance_check_path: Optional[str] = None
    duplicate_check_path: Optional[str] = None
    accounting_period_path: Optional[str] = None
    progress_stage: str = 'pending'
    progress_percent: int = 0
    progress_message: str = ''
    task_ids: List[str] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'JobContext':
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})

    def update_progress(self, stage: str, percent: int, message: str = ''):
        """更新进度信息"""
        self.progress_stage = stage
        self.progress_percent = min(max(percent, 0), 100)
        self.progress_message = message
        logger.info('作业 %s 进度: %s - %d%% - %s',
                    self.job_id, stage, percent, message)


@dataclass
class ScanTaskPayload(TaskPayload):
    """扫描任务负载"""
    task_type: str = TaskType.SCAN.value
    source_folder: str = ''
    recursive: bool = True
    file_extensions: List[str] = field(default_factory=lambda: ['.xlsx', '.xls'])
    folder_strategy: str = 'copy_sibling'
    folder_output_dir: Optional[str] = None
    folder_suffix: str = '＋检验版'

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ScanTaskPayload':
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class ParseTaskPayload(TaskPayload):
    """解析任务负载"""
    task_type: str = TaskType.PARSE.value
    file_path: str = ''
    lookup_file: Optional[str] = None
    lookup_data: Dict[str, Any] = field(default_factory=dict)
    bank_type: str = ''

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ParseTaskPayload':
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class MergeTaskPayload(TaskPayload):
    """合并任务负载"""
    task_type: str = TaskType.MERGE.value
    script_dir: str = ''
    incremental: bool = True
    parse_task_ids: List[str] = field(default_factory=list)
    lookup_source: Dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'MergeTaskPayload':
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class ReportTaskPayload(TaskPayload):
    """报告任务负载"""
    task_type: str = TaskType.REPORT.value
    script_dir: str = ''
    output_path: Optional[str] = None
    records: List[Dict[str, Any]] = field(default_factory=list)
    source_info: Dict[str, Any] = field(default_factory=dict)
    report_types: List[str] = field(
        default_factory=lambda: ['subject_summary', 'balance_check',
                                 'duplicate_check', 'accounting_period']
    )

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ReportTaskPayload':
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class PersistTaskPayload(TaskPayload):
    """持久化任务负载"""
    task_type: str = TaskType.PERSIST.value
    script_dir: str = ''
    batch_id: str = ''
    records: List[Dict[str, Any]] = field(default_factory=list)
    deduplicate: bool = True

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'PersistTaskPayload':
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})


@dataclass
class CleanupTaskPayload(TaskPayload):
    """清理任务负载"""
    task_type: str = TaskType.CLEANUP.value
    working_folder: str = ''
    excel_files: List[str] = field(default_factory=list)
    processed_files: List[str] = field(default_factory=list)
    error_files: List[Tuple[str, str]] = field(default_factory=list)
    unprocessed_files: List[str] = field(default_factory=list)
    strategy: str = 'keep_unprocessed'
    archive_dir_name: str = '已处理归档'

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'CleanupTaskPayload':
        payload = cls(**{k: v for k, v in data.items()
                         if k in cls.__dataclass_fields__})
        if 'error_files' in data and isinstance(data['error_files'], list):
            payload.error_files = [tuple(f) for f in data['error_files']]
        return payload


@dataclass
class TaskResult:
    """任务执行结果"""
    task_id: str
    job_id: str
    task_type: str
    status: str
    result: Dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None
    started_at: Optional[str] = None
    finished_at: Optional[str] = None
    duration_ms: float = 0
    worker_id: str = ''

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'TaskResult':
        return cls(**{k: v for k, v in data.items()
                      if k in cls.__dataclass_fields__})

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False)

    @classmethod
    def from_json(cls, json_str: str) -> 'TaskResult':
        return cls.from_dict(json.loads(json_str))


def create_payload_from_dict(data: Dict[str, Any]) -> TaskPayload:
    """根据任务类型创建对应的负载对象"""
    task_type = data.get('task_type', '')

    payload_classes = {
        TaskType.SCAN.value: ScanTaskPayload,
        TaskType.PARSE.value: ParseTaskPayload,
        TaskType.MERGE.value: MergeTaskPayload,
        TaskType.REPORT.value: ReportTaskPayload,
        TaskType.PERSIST.value: PersistTaskPayload,
        TaskType.CLEANUP.value: CleanupTaskPayload,
    }

    cls = payload_classes.get(task_type, TaskPayload)
    return cls.from_dict(data)


def create_payload_from_json(json_str: str) -> TaskPayload:
    """从JSON创建任务负载"""
    return create_payload_from_dict(json.loads(json_str))
