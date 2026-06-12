# -*- coding: utf-8 -*-
"""
工作进程管理器
负责从消息队列消费任务并执行，支持多 Worker 并发处理
"""

import os
import sys
import uuid
import time
import signal
import logging
import threading
import multiprocessing
from typing import List, Optional, Dict, Any, Callable
from dataclasses import dataclass, field
from datetime import datetime

from .config import TaskQueueConfig, load_config
from .mq_abstract import MessageQueueBackend, get_message_queue, close_message_queue
from .models import TaskPayload, TaskResult, TaskStatus, create_payload_from_dict
from .tasks import execute_task

logger = logging.getLogger('bankcheck.task_queue.worker')


def generate_worker_id() -> str:
    """生成 Worker ID"""
    return f"worker_{uuid.uuid4().hex[:12]}"


@dataclass
class WorkerStats:
    """Worker 统计信息"""
    worker_id: str
    started_at: str = field(default_factory=lambda: datetime.now().isoformat())
    tasks_processed: int = 0
    tasks_completed: int = 0
    tasks_failed: int = 0
    total_duration_ms: float = 0
    current_task_id: Optional[str] = None
    current_task_type: Optional[str] = None
    current_task_started_at: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            'worker_id': self.worker_id,
            'started_at': self.started_at,
            'tasks_processed': self.tasks_processed,
            'tasks_completed': self.tasks_completed,
            'tasks_failed': self.tasks_failed,
            'total_duration_ms': self.total_duration_ms,
            'avg_duration_ms': self.total_duration_ms / max(self.tasks_processed, 1),
            'current_task_id': self.current_task_id,
            'current_task_type': self.current_task_type,
            'current_task_started_at': self.current_task_started_at,
        }


class Worker:
    """
    单个 Worker 进程
    负责从指定队列消费任务并执行
    """

    def __init__(self,
                 config: TaskQueueConfig,
                 worker_id: Optional[str] = None,
                 queues: Optional[List[str]] = None):
        self.config = config
        self.worker_id = worker_id or generate_worker_id()
        self.queues = queues or config.worker.queues
        self.stats = WorkerStats(worker_id=self.worker_id)
        self._mq: Optional[MessageQueueBackend] = None
        self._running = False
        self._stopping = False
        self._current_payload: Optional[TaskPayload] = None
        self._task_timeout_timer: Optional[threading.Timer] = None

        self._setup_signal_handlers()

    def _setup_signal_handlers(self):
        """设置信号处理器"""
        def handle_signal(signum, frame):
            logger.info('Worker %s 收到信号 %s，准备优雅退出', self.worker_id, signum)
            self.stop()

        try:
            signal.signal(signal.SIGTERM, handle_signal)
            signal.signal(signal.SIGINT, handle_signal)
        except ValueError:
            pass

    def start(self):
        """启动 Worker"""
        logger.info('Worker %s 启动，监听队列: %s', self.worker_id, self.queues)
        self._running = True
        self._stopping = False
        self._mq = get_message_queue(config=self.config)

        while self._running:
            try:
                if self._stopping:
                    logger.info('Worker %s 正在停止，退出循环', self.worker_id)
                    break

                payload = self._mq.consume_task(
                    self.queues,
                    timeout=1
                )

                if payload:
                    self._process_task(payload)
                else:
                    time.sleep(0.1)

            except Exception as e:
                logger.error('Worker %s 主循环异常: %s', self.worker_id, e, exc_info=True)
                time.sleep(1)

        self._cleanup()
        logger.info('Worker %s 已停止，统计: %s', self.worker_id, self.stats.to_dict())

    def _process_task(self, payload: TaskPayload):
        """处理单个任务"""
        self._current_payload = payload
        self.stats.tasks_processed += 1
        self.stats.current_task_id = payload.task_id
        self.stats.current_task_type = payload.task_type
        self.stats.current_task_started_at = datetime.now().isoformat()

        logger.info('Worker %s 开始处理任务: %s (type=%s, job=%s)',
                    self.worker_id, payload.task_id, payload.task_type, payload.job_id)

        self._start_task_timeout(payload)

        try:
            result = execute_task(payload, self.worker_id)
            self._finish_task(payload, result)

        except Exception as e:
            logger.error('Worker %s 处理任务异常 %s: %s',
                         self.worker_id, payload.task_id, e, exc_info=True)
            result = TaskResult(
                task_id=payload.task_id,
                job_id=payload.job_id,
                task_type=payload.task_type,
                status=TaskStatus.FAILED.value,
                error=str(e),
                started_at=datetime.now().isoformat(),
                finished_at=datetime.now().isoformat(),
                duration_ms=0,
                worker_id=self.worker_id,
            )
            self._finish_task(payload, result)

        finally:
            self._cancel_task_timeout()
            self._current_payload = None
            self.stats.current_task_id = None
            self.stats.current_task_type = None
            self.stats.current_task_started_at = None

    def _start_task_timeout(self, payload: TaskPayload):
        """启动任务超时计时器"""
        timeout = self.config.worker.task_timeout

        def timeout_handler():
            logger.warning('Worker %s 任务 %s 执行超时 (%ds)，强制终止',
                           self.worker_id, payload.task_id, timeout)
            if self._mq:
                self._mq.update_task_status(payload.task_id, TaskStatus.TIMEOUT.value)
                result = TaskResult(
                    task_id=payload.task_id,
                    job_id=payload.job_id,
                    task_type=payload.task_type,
                    status=TaskStatus.TIMEOUT.value,
                    error=f"Task timeout after {timeout} seconds",
                    started_at=datetime.now().isoformat(),
                    finished_at=datetime.now().isoformat(),
                    duration_ms=timeout * 1000,
                    worker_id=self.worker_id,
                )
                self._mq.store_result(result)
                self._mq.reject_task(payload.task_id, requeue=True)

        self._task_timeout_timer = threading.Timer(timeout, timeout_handler)
        self._task_timeout_timer.daemon = True
        self._task_timeout_timer.start()

    def _cancel_task_timeout(self):
        """取消任务超时计时器"""
        if self._task_timeout_timer:
            self._task_timeout_timer.cancel()
            self._task_timeout_timer = None

    def _finish_task(self, payload: TaskPayload, result: TaskResult):
        """完成任务处理"""
        if self._mq is None:
            return

        self.stats.total_duration_ms += result.duration_ms

        if result.status == TaskStatus.COMPLETED.value:
            self.stats.tasks_completed += 1
            self._mq.store_result(result)
            self._mq.acknowledge_task(payload.task_id)
            logger.info('Worker %s 任务完成: %s (type=%s, duration=%.2fms)',
                        self.worker_id, payload.task_id, payload.task_type, result.duration_ms)
        else:
            self.stats.tasks_failed += 1
            self._mq.store_result(result)

            if result.status == TaskStatus.FAILED.value:
                requeue = payload.retry_count < payload.max_retries
                self._mq.reject_task(payload.task_id, requeue=requeue)
                logger.warning('Worker %s 任务失败: %s (type=%s, error=%s, retry=%d/%d)',
                             self.worker_id, payload.task_id, payload.task_type,
                             result.error, payload.retry_count, payload.max_retries)
            else:
                self._mq.reject_task(payload.task_id, requeue=False)

    def stop(self):
        """停止 Worker（优雅退出）"""
        if not self._running:
            return

        logger.info('Worker %s 收到停止请求，等待当前任务完成...', self.worker_id)
        self._stopping = True

        if self._current_payload:
            wait_start = time.time()
            while self._current_payload:
                if time.time() - wait_start > self.config.worker.graceful_stop_timeout:
                    logger.warning('Worker %s 等待超时，强制退出', self.worker_id)
                    break
                time.sleep(0.1)

        self._running = False

    def _cleanup(self):
        """清理资源"""
        self._cancel_task_timeout()
        close_message_queue()
        self._mq = None

    def get_stats(self) -> WorkerStats:
        """获取 Worker 统计信息"""
        return self.stats


class WorkerManager:
    """
    Worker 管理器
    负责启动、管理和监控多个 Worker 进程
    """

    def __init__(self,
                 config_path: Optional[str] = None,
                 config: Optional[TaskQueueConfig] = None,
                 num_workers: int = 1):
        if config is None:
            config = load_config(config_path)
        self.config = config
        self.num_workers = num_workers
        self._workers: List[Worker] = []
        self._processes: List[multiprocessing.Process] = []
        self._running = False
        self._monitor_thread: Optional[threading.Thread] = None

    def start(self):
        """启动所有 Worker 进程"""
        logger.info('WorkerManager 启动 %d 个 Worker 进程', self.num_workers)
        self._running = True

        for i in range(self.num_workers):
            worker_id = f"worker_{i:02d}_{uuid.uuid4().hex[:8]}"
            process = multiprocessing.Process(
                target=self._run_worker,
                args=(self.config, worker_id, self.config.worker.queues),
                name=f"Worker-{i:02d}",
                daemon=True
            )
            process.start()
            self._processes.append(process)
            logger.info('已启动 Worker 进程 %s (PID=%d)', worker_id, process.pid)

        self._start_monitor()

    @staticmethod
    def _run_worker(config: TaskQueueConfig, worker_id: str, queues: List[str]):
        """Worker 进程入口函数"""
        worker = Worker(config, worker_id=worker_id, queues=queues)
        worker.start()

    def _start_monitor(self):
        """启动监控线程"""
        def monitor():
            while self._running:
                try:
                    for i, process in enumerate(self._processes):
                        if not process.is_alive():
                            logger.warning('Worker 进程 %s (PID=%d) 已退出，尝试重启...',
                                         process.name, process.pid)
                            worker_id = f"worker_{i:02d}_{uuid.uuid4().hex[:8]}"
                            new_process = multiprocessing.Process(
                                target=self._run_worker,
                                args=(self.config, worker_id, self.config.worker.queues),
                                name=f"Worker-{i:02d}",
                                daemon=True
                            )
                            new_process.start()
                            self._processes[i] = new_process
                            logger.info('已重启 Worker 进程 %s (PID=%d)',
                                        worker_id, new_process.pid)

                    time.sleep(5)
                except Exception as e:
                    logger.error('Worker 监控线程异常: %s', e, exc_info=True)
                    time.sleep(5)

        self._monitor_thread = threading.Thread(target=monitor, daemon=True)
        self._monitor_thread.start()
        logger.info('Worker 监控线程已启动')

    def stop(self):
        """停止所有 Worker 进程"""
        logger.info('WorkerManager 正在停止所有 Worker 进程')
        self._running = False

        for process in self._processes:
            if process.is_alive():
                process.terminate()

        for process in self._processes:
            process.join(timeout=self.config.worker.graceful_stop_timeout)
            if process.is_alive():
                process.kill()

        logger.info('所有 Worker 进程已停止')

    def get_status(self) -> Dict[str, Any]:
        """获取所有 Worker 的状态"""
        statuses = []
        for i, process in enumerate(self._processes):
            statuses.append({
                'index': i,
                'name': process.name,
                'pid': process.pid,
                'is_alive': process.is_alive(),
            })

        return {
            'num_workers': self.num_workers,
            'running': self._running,
            'workers': statuses,
        }

    def scale_workers(self, new_num: int):
        """动态调整 Worker 数量"""
        if new_num == self.num_workers:
            return

        logger.info('调整 Worker 数量: %d -> %d', self.num_workers, new_num)

        if new_num > self.num_workers:
            for i in range(self.num_workers, new_num):
                worker_id = f"worker_{i:02d}_{uuid.uuid4().hex[:8]}"
                process = multiprocessing.Process(
                    target=self._run_worker,
                    args=(self.config, worker_id, self.config.worker.queues),
                    name=f"Worker-{i:02d}",
                    daemon=True
                )
                process.start()
                self._processes.append(process)
                logger.info('已新增 Worker 进程 %s (PID=%d)', worker_id, process.pid)
        else:
            for i in range(new_num, self.num_workers):
                process = self._processes[i]
                if process.is_alive():
                    process.terminate()
                    logger.info('已终止 Worker 进程 %s (PID=%d)', process.name, process.pid)
            self._processes = self._processes[:new_num]

        self.num_workers = new_num


def run_worker(config_path: Optional[str] = None,
               worker_id: Optional[str] = None,
               queues: Optional[List[str]] = None):
    """
    命令行方式运行单个 Worker

    Args:
        config_path: 配置文件路径
        worker_id: Worker ID
        queues: 监听的队列列表
    """
    config = load_config(config_path)
    worker = Worker(config, worker_id=worker_id, queues=queues)

    def handle_shutdown(signum, frame):
        logger.info('收到退出信号，正在停止 Worker...')
        worker.stop()

    signal.signal(signal.SIGTERM, handle_shutdown)
    signal.signal(signal.SIGINT, handle_shutdown)

    try:
        worker.start()
    except KeyboardInterrupt:
        logger.info('用户中断，正在停止 Worker...')
        worker.stop()


def run_worker_manager(config_path: Optional[str] = None,
                       num_workers: int = None):
    """
    命令行方式运行 Worker 管理器

    Args:
        config_path: 配置文件路径
        num_workers: Worker 数量，默认为 CPU 核心数
    """
    config = load_config(config_path)

    if num_workers is None:
        num_workers = multiprocessing.cpu_count()

    manager = WorkerManager(config=config, num_workers=num_workers)

    def handle_shutdown(signum, frame):
        logger.info('收到退出信号，正在停止 WorkerManager...')
        manager.stop()
        sys.exit(0)

    signal.signal(signal.SIGTERM, handle_shutdown)
    signal.signal(signal.SIGINT, handle_shutdown)

    try:
        manager.start()
        while manager._running:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info('用户中断，正在停止 WorkerManager...')
        manager.stop()
