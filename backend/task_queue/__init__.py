# -*- coding: utf-8 -*-
"""
消息队列异步任务模块
将单次处理拆分为扫描、解析、合并、报告等异步任务
通过 Redis/RabbitMQ 调度，支撑多用户并发提交的大型批处理场景
"""

from .models import (
    TaskStatus,
    TaskType,
    TaskPriority,
    TaskPayload,
    TaskResult,
    JobContext,
    ScanTaskPayload,
    ParseTaskPayload,
    MergeTaskPayload,
    ReportTaskPayload,
    PersistTaskPayload,
    CleanupTaskPayload,
    generate_job_id,
    generate_task_id,
)
from .mq_abstract import MessageQueueBackend, get_message_queue
from .worker import Worker, WorkerManager, generate_worker_id
from .orchestrator import TaskOrchestrator, JobOrchestrator, get_job_orchestrator
from .tasks import (
    task_scan_files,
    task_parse_file,
    task_merge_results,
    task_generate_reports,
    task_persist_data,
    task_cleanup,
)
from .config import TaskQueueConfig, load_config

__all__ = [
    'TaskStatus',
    'TaskType',
    'TaskPriority',
    'TaskPayload',
    'TaskResult',
    'JobContext',
    'ScanTaskPayload',
    'ParseTaskPayload',
    'MergeTaskPayload',
    'ReportTaskPayload',
    'PersistTaskPayload',
    'CleanupTaskPayload',
    'generate_job_id',
    'generate_task_id',
    'generate_worker_id',
    'MessageQueueBackend',
    'get_message_queue',
    'Worker',
    'WorkerManager',
    'TaskOrchestrator',
    'JobOrchestrator',
    'get_job_orchestrator',
    'task_scan_files',
    'task_parse_file',
    'task_merge_results',
    'task_generate_reports',
    'task_persist_data',
    'task_cleanup',
    'TaskQueueConfig',
    'load_config',
]
