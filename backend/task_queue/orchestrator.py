# -*- coding: utf-8 -*-
"""
任务编排器
负责协调任务之间的依赖关系，将整个处理流程编排为有向无环图（DAG）
流程：扫描 -> 并行解析 -> 合并 -> 报告 + 持久化 -> 清理
"""

import os
import sys
import json
import time
import logging
import threading
from typing import Dict, Any, Optional, List, Callable, Tuple
from dataclasses import dataclass, field, asdict
from datetime import datetime

from .config import TaskQueueConfig, load_config
from .mq_abstract import MessageQueueBackend, get_message_queue
from .models import (
    TaskStatus,
    TaskType,
    TaskPriority,
    JobContext,
    ScanTaskPayload,
    ParseTaskPayload,
    MergeTaskPayload,
    ReportTaskPayload,
    PersistTaskPayload,
    CleanupTaskPayload,
    TaskResult,
    generate_job_id,
)

logger = logging.getLogger('bankcheck.task_queue.orchestrator')

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _import_bankcheck():
    try:
        import bankcheck
        return bankcheck
    except ImportError as e:
        logger.error('导入 bankcheck 模块失败: %s', e)
        raise


def _import_batch_manager():
    try:
        import batch_manager as batch_module
        return batch_module
    except ImportError as e:
        logger.warning('导入 batch_manager 模块失败: %s', e)
        return None


@dataclass
class TaskDependency:
    """任务依赖关系"""
    task_id: str
    task_type: str
    depends_on: List[str] = field(default_factory=list)
    status: str = TaskStatus.PENDING.value


@dataclass
class JobPlan:
    """作业执行计划"""
    job_id: str
    tasks: List[TaskDependency] = field(default_factory=list)
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        return {
            'job_id': self.job_id,
            'tasks': [asdict(t) for t in self.tasks],
            'created_at': self.created_at,
        }


class TaskOrchestrator:
    """
    任务编排器
    负责发布任务到消息队列，并跟踪任务执行状态
    """

    def __init__(self,
                 config_path: Optional[str] = None,
                 config: Optional[TaskQueueConfig] = None):
        if config is None:
            config = load_config(config_path)
        self.config = config
        self._mq = get_message_queue(config=self.config)

    def publish_scan_task(self,
                          job_id: str,
                          source_folder: str,
                          priority: int = TaskPriority.HIGH.value) -> str:
        """
        发布扫描任务

        Args:
            job_id: 作业ID
            source_folder: 源文件夹路径
            priority: 任务优先级

        Returns:
            任务ID
        """
        payload = ScanTaskPayload(
            job_id=job_id,
            source_folder=source_folder,
            priority=priority,
        )
        return self._mq.publish_task('scan', payload, priority=priority)

    def publish_parse_tasks(self,
                            job_id: str,
                            file_paths: List[str],
                            lookup_file: Optional[str] = None,
                            lookup_data: Optional[Dict[str, Any]] = None,
                            priority: int = TaskPriority.HIGH.value) -> List[str]:
        """
        批量发布解析任务

        Args:
            job_id: 作业ID
            file_paths: 文件路径列表
            lookup_file: 查找表文件路径
            lookup_data: 查找表数据（优先使用）
            priority: 任务优先级

        Returns:
            任务ID列表
        """
        task_ids = []
        for file_path in file_paths:
            payload = ParseTaskPayload(
                job_id=job_id,
                file_path=file_path,
                lookup_file=lookup_file,
                lookup_data=lookup_data or {},
                priority=priority,
            )
            task_id = self._mq.publish_task('parse', payload, priority=priority)
            task_ids.append(task_id)

        logger.info('已发布 %d 个解析任务 (job=%s)', len(task_ids), job_id)
        return task_ids

    def publish_merge_task(self,
                           job_id: str,
                           script_dir: str,
                           parse_task_ids: List[str],
                           incremental: bool = True,
                           lookup_source: Optional[Dict[str, Any]] = None,
                           priority: int = TaskPriority.NORMAL.value) -> str:
        """
        发布合并任务

        Args:
            job_id: 作业ID
            script_dir: 脚本目录
            parse_task_ids: 解析任务ID列表
            incremental: 是否增量模式
            lookup_source: 查找表数据源
            priority: 任务优先级

        Returns:
            任务ID
        """
        payload = MergeTaskPayload(
            job_id=job_id,
            script_dir=script_dir,
            incremental=incremental,
            parse_task_ids=parse_task_ids,
            lookup_source=lookup_source or {},
            priority=priority,
        )
        return self._mq.publish_task('merge', payload, priority=priority)

    def publish_report_task(self,
                            job_id: str,
                            script_dir: str,
                            output_path: Optional[str],
                            records: List[Dict[str, Any]],
                            source_info: Optional[Dict[str, Any]] = None,
                            report_types: Optional[List[str]] = None,
                            priority: int = TaskPriority.LOW.value) -> str:
        """
        发布报告任务

        Args:
            job_id: 作业ID
            script_dir: 脚本目录
            output_path: 总表输出路径
            records: 记录列表
            source_info: 来源信息
            report_types: 报告类型列表
            priority: 任务优先级

        Returns:
            任务ID
        """
        payload = ReportTaskPayload(
            job_id=job_id,
            script_dir=script_dir,
            output_path=output_path,
            records=records,
            source_info=source_info or {},
            report_types=report_types or [
                'subject_summary', 'balance_check',
                'duplicate_check', 'accounting_period'
            ],
            priority=priority,
        )
        return self._mq.publish_task('report', payload, priority=priority)

    def publish_persist_task(self,
                             job_id: str,
                             script_dir: str,
                             batch_id: str,
                             records: List[Dict[str, Any]],
                             deduplicate: bool = True,
                             priority: int = TaskPriority.LOW.value) -> str:
        """
        发布持久化任务

        Args:
            job_id: 作业ID
            script_dir: 脚本目录
            batch_id: 批次ID
            records: 记录列表
            deduplicate: 是否去重
            priority: 任务优先级

        Returns:
            任务ID
        """
        payload = PersistTaskPayload(
            job_id=job_id,
            script_dir=script_dir,
            batch_id=batch_id,
            records=records,
            deduplicate=deduplicate,
            priority=priority,
        )
        return self._mq.publish_task('persist', payload, priority=priority)

    def publish_cleanup_task(self,
                             job_id: str,
                             working_folder: str,
                             excel_files: List[str],
                             processed_files: List[str],
                             error_files: List[Tuple[str, str]],
                             unprocessed_files: List[str],
                             strategy: str = 'keep_unprocessed',
                             archive_dir_name: str = '已处理归档',
                             priority: int = TaskPriority.LOW.value) -> str:
        """
        发布清理任务

        Args:
            job_id: 作业ID
            working_folder: 工作文件夹
            excel_files: 所有Excel文件
            processed_files: 已处理文件
            error_files: 错误文件列表 [(path, error), ...]
            unprocessed_files: 未处理文件
            strategy: 清理策略
            archive_dir_name: 归档目录名
            priority: 任务优先级

        Returns:
            任务ID
        """
        payload = CleanupTaskPayload(
            job_id=job_id,
            working_folder=working_folder,
            excel_files=excel_files,
            processed_files=processed_files,
            error_files=error_files,
            unprocessed_files=unprocessed_files,
            strategy=strategy,
            archive_dir_name=archive_dir_name,
            priority=priority,
        )
        return self._mq.publish_task('cleanup', payload, priority=priority)

    def get_task_status(self, task_id: str) -> Optional[str]:
        """获取任务状态"""
        return self._mq.get_task_status(task_id)

    def get_task_result(self, task_id: str) -> Optional[TaskResult]:
        """获取任务结果"""
        return self._mq.get_result(task_id)

    def cancel_task(self, task_id: str) -> bool:
        """取消任务"""
        return self._mq.cancel_task(task_id)

    def wait_for_task(self, task_id: str, timeout: Optional[float] = None) -> Optional[TaskResult]:
        """
        等待任务完成

        Args:
            task_id: 任务ID
            timeout: 超时时间（秒）

        Returns:
            TaskResult 或 None（超时）
        """
        start_time = time.time()
        while True:
            result = self._mq.get_result(task_id)
            if result and result.status in [TaskStatus.COMPLETED.value,
                                            TaskStatus.FAILED.value,
                                            TaskStatus.CANCELLED.value,
                                            TaskStatus.TIMEOUT.value]:
                return result

            if timeout is not None and (time.time() - start_time) > timeout:
                return None

            time.sleep(0.5)

    def wait_for_tasks(self, task_ids: List[str],
                       timeout: Optional[float] = None,
                       on_progress: Optional[Callable[[int, int], None]] = None) -> Dict[str, TaskResult]:
        """
        等待多个任务完成

        Args:
            task_ids: 任务ID列表
            timeout: 超时时间（秒）
            on_progress: 进度回调函数 (completed_count, total_count)

        Returns:
            任务ID到结果的映射
        """
        results: Dict[str, TaskResult] = {}
        pending = set(task_ids)
        start_time = time.time()

        while pending:
            for task_id in list(pending):
                result = self._mq.get_result(task_id)
                if result and result.status in [TaskStatus.COMPLETED.value,
                                                TaskStatus.FAILED.value,
                                                TaskStatus.CANCELLED.value,
                                                TaskStatus.TIMEOUT.value]:
                    results[task_id] = result
                    pending.remove(task_id)

                    if on_progress:
                        on_progress(len(task_ids) - len(pending), len(task_ids))

            if timeout is not None and (time.time() - start_time) > timeout:
                break

            if pending:
                time.sleep(0.5)

        return results


class JobOrchestrator:
    """
    作业编排器
    负责将整个处理流程编排为 DAG，并跟踪作业执行进度
    """

    def __init__(self,
                 config_path: Optional[str] = None,
                 config: Optional[TaskQueueConfig] = None):
        if config is None:
            config = load_config(config_path)
        self.config = config
        self._task_orchestrator = TaskOrchestrator(config=config)
        self._jobs: Dict[str, JobContext] = {}
        self._job_plans: Dict[str, JobPlan] = {}
        self._progress_callbacks: Dict[str, Callable[[JobContext], None]] = {}
        self._monitor_thread: Optional[threading.Thread] = None
        self._monitoring = False

    def submit_job(self,
                   input_folder: str,
                   script_dir: str,
                   incremental: bool = True,
                   keep_strategy: str = 'keep_unprocessed',
                   operator: str = '',
                   user_id: str = '',
                   on_progress: Optional[Callable[[JobContext], None]] = None) -> JobContext:
        """
        提交一个新的处理作业

        Args:
            input_folder: 输入文件夹路径
            script_dir: 脚本目录
            incremental: 是否增量模式
            keep_strategy: 文件保留策略
            operator: 操作员
            user_id: 用户ID
            on_progress: 进度回调函数

        Returns:
            JobContext 作业上下文
        """
        job_id = generate_job_id()
        logger.info('提交新作业: job_id=%s, folder=%s', job_id, input_folder)

        job_context = JobContext(
            job_id=job_id,
            user_id=user_id,
            operator=operator,
            input_folder=input_folder,
            script_dir=script_dir,
            incremental=incremental,
            keep_strategy=keep_strategy,
            status=TaskStatus.QUEUED.value,
        )

        self._jobs[job_id] = job_context
        if on_progress:
            self._progress_callbacks[job_id] = on_progress

        self._start_job_execution(job_context)

        return job_context

    def _start_job_execution(self, job_context: JobContext):
        """开始执行作业"""
        job_id = job_context.job_id
        job_context.status = TaskStatus.RUNNING.value
        job_context.started_at = datetime.now().isoformat()

        self._notify_progress(job_context, 'starting', 5, '作业已提交，正在初始化...')

        threading.Thread(
            target=self._execute_job_dag,
            args=(job_context,),
            daemon=True
        ).start()

    def _execute_job_dag(self, job_context: JobContext):
        """
        执行作业 DAG
        流程：扫描 -> 并行解析 -> 合并 -> 报告 + 持久化 -> 清理
        """
        job_id = job_context.job_id
        try:
            bankcheck = _import_bankcheck()

            lookup_file = bankcheck.find_lookup_file(job_context.script_dir)
            lookup_missing = lookup_file is None
            if lookup_missing:
                logger.warning('未找到主体查找表')
                lookup_data = bankcheck.load_lookup_table(None)
            else:
                logger.info('正在预加载主体查找表...')
                lookup_data = bankcheck.load_lookup_table(lookup_file)
                logger.info('主体查找表预加载完成')

            self._notify_progress(job_context, 'scanning', 10, '正在扫描文件...')

            scan_task_id = self._task_orchestrator.publish_scan_task(
                job_id, job_context.input_folder
            )
            job_context.task_ids.append(scan_task_id)

            scan_result = self._task_orchestrator.wait_for_task(scan_task_id)
            if not scan_result or scan_result.status != TaskStatus.COMPLETED.value:
                raise Exception(f"扫描任务失败: {scan_result.error if scan_result else '未知错误'}")

            scan_data = scan_result.result
            excel_files = scan_data.get('excel_files', [])
            working_folder = scan_data.get('working_folder', '')
            job_context.total_files = scan_data.get('total_files', 0)
            job_context.metadata['working_folder'] = working_folder
            job_context.metadata['excel_files'] = excel_files

            if job_context.total_files == 0:
                self._notify_progress(job_context, 'completed', 100, '文件夹中未发现任何 Excel 文件')
                job_context.status = TaskStatus.COMPLETED.value
                job_context.finished_at = datetime.now().isoformat()
                self._finalize_job(job_context, success=True)
                return

            self._notify_progress(job_context, 'parsing', 20,
                                  f'发现 {job_context.total_files} 个文件，开始并行解析...')

            parse_task_ids = self._task_orchestrator.publish_parse_tasks(
                job_id, excel_files, lookup_data=lookup_data
            )
            job_context.task_ids.extend(parse_task_ids)

            def parse_progress(completed, total):
                progress = 20 + int(completed / max(total, 1) * 50)
                self._notify_progress(
                    job_context, 'parsing', progress,
                    f'正在解析文件... ({completed}/{total})'
                )

            parse_results = self._task_orchestrator.wait_for_tasks(
                parse_task_ids, on_progress=parse_progress
            )

            processed_files: List[str] = []
            unprocessed_files: List[str] = []
            error_files: List[Tuple[str, str]] = []

            for task_id, result in parse_results.items():
                if result.status != TaskStatus.COMPLETED.value:
                    file_path = result.result.get('file_path', 'unknown')
                    error_files.append((file_path, result.error or '解析失败'))
                    continue

                status = result.result.get('status', '')
                file_path = result.result.get('file_path', '')

                if status == 'unrecognized':
                    unprocessed_files.append(file_path)
                elif status == 'success':
                    processed_files.append(file_path)

            job_context.processed_files = len(processed_files)
            job_context.unprocessed_files = len(unprocessed_files)
            job_context.error_files = len(error_files)

            self._notify_progress(job_context, 'merging', 75, '正在合并解析结果...')

            merge_task_id = self._task_orchestrator.publish_merge_task(
                job_id, job_context.script_dir, parse_task_ids,
                incremental=job_context.incremental, lookup_source=lookup_data
            )
            job_context.task_ids.append(merge_task_id)

            merge_result = self._task_orchestrator.wait_for_task(merge_task_id)
            if not merge_result or merge_result.status != TaskStatus.COMPLETED.value:
                raise Exception(f"合并任务失败: {merge_result.error if merge_result else '未知错误'}")

            merge_data = merge_result.result
            final_rows = merge_data.get('all_rows', [])
            output_path = merge_data.get('output_path')
            job_context.output_path = output_path
            job_context.total_records = merge_data.get('total_records', 0)
            job_context.new_records = merge_data.get('new_record_count', 0)
            job_context.duplicate_records = merge_data.get('duplicate_record_count', 0)
            job_context.metadata['incremental_mode'] = merge_data.get('incremental_mode', False)

            batch_module = _import_batch_manager()
            if batch_module:
                batch_manager = batch_module.get_batch_manager(job_context.script_dir)
                batch_info = batch_manager.start_batch(
                    input_folder=job_context.input_folder,
                    operator=job_context.operator,
                )
                job_context.batch_id = batch_info.batch_id
                logger.info('已创建批次: %s', batch_info.batch_id)

            self._notify_progress(job_context, 'reporting', 85, '正在生成报告和持久化数据...')

            report_task_id = self._task_orchestrator.publish_report_task(
                job_id, job_context.script_dir, output_path, final_rows,
                source_info={
                    '数据来源': '异步任务流水线',
                    '总表文件': os.path.basename(output_path) if output_path else '内存数据',
                    '记录数': len(final_rows),
                    '运行模式': '增量合并' if job_context.incremental else '全量覆盖',
                    '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                }
            )

            persist_task_id = None
            if final_rows and job_context.batch_id:
                persist_task_id = self._task_orchestrator.publish_persist_task(
                    job_id, job_context.script_dir,
                    job_context.batch_id, final_rows
                )

            job_context.task_ids.append(report_task_id)
            if persist_task_id:
                job_context.task_ids.append(persist_task_id)

            report_result = self._task_orchestrator.wait_for_task(report_task_id)
            if report_result and report_result.status == TaskStatus.COMPLETED.value:
                report_paths = report_result.result.get('report_paths', {})
                job_context.subject_summary_path = report_paths.get('subject_summary')
                job_context.balance_check_path = report_paths.get('balance_check')
                job_context.duplicate_check_path = report_paths.get('duplicate_check')
                job_context.accounting_period_path = report_paths.get('accounting_period')

            if persist_task_id:
                persist_result = self._task_orchestrator.wait_for_task(persist_task_id)
                if persist_result and persist_result.status == TaskStatus.COMPLETED.value:
                    persist_data = persist_result.result
                    job_context.metadata['db_inserted'] = persist_data.get('db_inserted', 0)
                    job_context.metadata['db_duplicates'] = persist_data.get('db_duplicates', 0)

            self._notify_progress(job_context, 'cleanup', 95, '正在清理临时文件...')

            cleanup_task_id = self._task_orchestrator.publish_cleanup_task(
                job_id, working_folder, excel_files,
                processed_files, error_files, unprocessed_files,
                strategy=job_context.keep_strategy
            )
            job_context.task_ids.append(cleanup_task_id)

            self._task_orchestrator.wait_for_task(cleanup_task_id)

            if batch_module and job_context.batch_id:
                batch_manager = batch_module.get_batch_manager(job_context.script_dir)
                result_data = {
                    'total_records': job_context.total_records,
                    'new_records': job_context.new_records,
                    'duplicate_records': job_context.duplicate_records,
                    'processed_files': processed_files,
                    'unprocessed_files': unprocessed_files,
                    'error_files': error_files,
                    'incremental_mode': job_context.incremental,
                    'output_folder': working_folder,
                    'summary_table_path': output_path,
                }
                batch_manager.finish_batch(
                    job_context.batch_id, result_data, status='success'
                )

            self._notify_progress(job_context, 'completed', 100,
                                  f'处理完成！共处理 {job_context.total_records} 条记录')
            job_context.status = TaskStatus.COMPLETED.value
            job_context.finished_at = datetime.now().isoformat()
            self._finalize_job(job_context, success=True)

        except Exception as e:
            logger.error('作业执行失败 %s: %s', job_id, e, exc_info=True)
            job_context.status = TaskStatus.FAILED.value
            job_context.error_message = str(e)
            job_context.finished_at = datetime.now().isoformat()

            if _import_batch_manager() and job_context.batch_id:
                try:
                    batch_manager = _import_batch_manager().get_batch_manager(job_context.script_dir)
                    batch_manager.finish_batch(
                        job_context.batch_id, {}, status='failed', error_message=str(e)
                    )
                except Exception:
                    pass

            self._notify_progress(job_context, 'failed', 100, f'处理失败: {e}')
            self._finalize_job(job_context, success=False)

    def _notify_progress(self, job_context: JobContext,
                         stage: str, percent: int, message: str):
        """通知作业进度"""
        job_context.update_progress(stage, percent, message)

        callback = self._progress_callbacks.get(job_context.job_id)
        if callback:
            try:
                callback(job_context)
            except Exception as e:
                logger.error('进度回调异常: %s', e)

    def _finalize_job(self, job_context: JobContext, success: bool):
        """作业完成后的清理工作"""
        if job_context.job_id in self._progress_callbacks:
            del self._progress_callbacks[job_context.job_id]

        logger.info('作业 %s 已完成，状态=%s, 总记录=%d',
                    job_context.job_id, job_context.status, job_context.total_records)

    def get_job_status(self, job_id: str) -> Optional[JobContext]:
        """获取作业状态"""
        return self._jobs.get(job_id)

    def cancel_job(self, job_id: str) -> bool:
        """取消作业"""
        job_context = self._jobs.get(job_id)
        if not job_context:
            return False

        if job_context.status in [TaskStatus.COMPLETED.value,
                                  TaskStatus.FAILED.value,
                                  TaskStatus.CANCELLED.value]:
            return False

        for task_id in job_context.task_ids:
            try:
                self._task_orchestrator.cancel_task(task_id)
            except Exception:
                pass

        job_context.status = TaskStatus.CANCELLED.value
        job_context.finished_at = datetime.now().isoformat()
        self._notify_progress(job_context, 'cancelled', 100, '作业已取消')

        if _import_batch_manager() and job_context.batch_id:
            try:
                batch_manager = _import_batch_manager().get_batch_manager(job_context.script_dir)
                batch_manager.finish_batch(
                    job_context.batch_id, {}, status='cancelled', error_message='用户取消'
                )
            except Exception:
                pass

        self._finalize_job(job_context, success=False)
        return True

    def list_jobs(self, status: Optional[str] = None,
                  limit: int = 100) -> List[JobContext]:
        """列出作业"""
        jobs = list(self._jobs.values())
        if status:
            jobs = [j for j in jobs if j.status == status]
        jobs.sort(key=lambda j: j.created_at, reverse=True)
        return jobs[:limit]

    def start_monitor(self):
        """启动监控线程（可选）"""
        if self._monitoring:
            return

        def monitor():
            while self._monitoring:
                try:
                    for job_id, job_context in self._jobs.items():
                        if job_context.status == TaskStatus.RUNNING.value:
                            pass
                    time.sleep(10)
                except Exception as e:
                    logger.error('监控线程异常: %s', e)

        self._monitoring = True
        self._monitor_thread = threading.Thread(target=monitor, daemon=True)
        self._monitor_thread.start()
        logger.info('作业编排器监控线程已启动')

    def stop_monitor(self):
        """停止监控线程"""
        self._monitoring = False
        if self._monitor_thread:
            self._monitor_thread.join(timeout=5)
            self._monitor_thread = None


_global_job_orchestrator: Optional[JobOrchestrator] = None


def get_job_orchestrator(config_path: Optional[str] = None,
                         config: Optional[TaskQueueConfig] = None) -> JobOrchestrator:
    """获取全局作业编排器实例"""
    global _global_job_orchestrator
    if _global_job_orchestrator is None:
        _global_job_orchestrator = JobOrchestrator(config_path=config_path, config=config)
    return _global_job_orchestrator
