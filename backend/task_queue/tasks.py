# -*- coding: utf-8 -*-
"""
各阶段异步任务实现
将 bankcheck.py 中的处理流程拆分为独立的可执行任务
"""

import os
import sys
import time
import shutil
import logging
import traceback
from typing import Dict, Any, List, Optional, Tuple
from datetime import datetime

from .models import (
    TaskPayload,
    TaskResult,
    TaskStatus,
    ScanTaskPayload,
    ParseTaskPayload,
    MergeTaskPayload,
    ReportTaskPayload,
    PersistTaskPayload,
    CleanupTaskPayload,
)

logger = logging.getLogger('bankcheck.task_queue.tasks')

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _import_bankcheck():
    """动态导入 bankcheck 模块"""
    try:
        import bankcheck
        return bankcheck
    except ImportError as e:
        logger.error('导入 bankcheck 模块失败: %s', e)
        raise


def _import_database():
    """动态导入 database 模块"""
    try:
        import database as db_module
        return db_module
    except ImportError as e:
        logger.warning('导入 database 模块失败: %s', e)
        return None


def task_scan_files(payload: ScanTaskPayload) -> TaskResult:
    """
    扫描任务：扫描指定文件夹中的 Excel 文件

    Args:
        payload: 扫描任务负载

    Returns:
        TaskResult 任务执行结果
    """
    start_time = time.perf_counter()
    started_at = datetime.now().isoformat()

    try:
        bankcheck = _import_bankcheck()

        if not os.path.isdir(payload.source_folder):
            raise ValueError(f"源文件夹不存在: {payload.source_folder}")

        working_folder, working_folder_is_copy = bankcheck.prepare_working_folder(
            payload.source_folder,
            strategy=payload.folder_strategy,
            output_dir=payload.folder_output_dir,
            suffix=payload.folder_suffix,
            logger=logger,
        )

        excel_files = bankcheck.scan_excel_files(working_folder)
        logger.info('扫描完成，发现 %d 个 Excel 文件', len(excel_files))

        result_data = {
            'source_folder': payload.source_folder,
            'working_folder': working_folder,
            'working_folder_is_copy': working_folder_is_copy,
            'excel_files': excel_files,
            'total_files': len(excel_files),
        }

        duration_ms = (time.perf_counter() - start_time) * 1000

        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.COMPLETED.value,
            result=result_data,
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )

    except Exception as e:
        logger.error('扫描任务失败 %s: %s', payload.task_id, e, exc_info=True)
        duration_ms = (time.perf_counter() - start_time) * 1000
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=str(e),
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )


def task_parse_file(payload: ParseTaskPayload) -> TaskResult:
    """
    解析任务：解析单个 Excel 文件，提取银行流水记录

    Args:
        payload: 解析任务负载

    Returns:
        TaskResult 任务执行结果
    """
    start_time = time.perf_counter()
    started_at = datetime.now().isoformat()

    try:
        bankcheck = _import_bankcheck()

        if not os.path.isfile(payload.file_path):
            raise ValueError(f"文件不存在: {payload.file_path}")

        lookup_data = payload.lookup_data
        if not lookup_data and payload.lookup_file:
            lookup_data = bankcheck.load_lookup_table(payload.lookup_file)

        bank = payload.bank_type or bankcheck.identify_bank(payload.file_path)
        if not bank or bank not in bankcheck.BANK_PROCESSORS:
            result_data = {
                'file_path': payload.file_path,
                'bank_type': bank,
                'records': [],
                'status': 'unrecognized',
            }
            duration_ms = (time.perf_counter() - start_time) * 1000
            return TaskResult(
                task_id=payload.task_id,
                job_id=payload.job_id,
                task_type=payload.task_type,
                status=TaskStatus.COMPLETED.value,
                result=result_data,
                started_at=started_at,
                finished_at=datetime.now().isoformat(),
                duration_ms=duration_ms,
            )

        processor = bankcheck.BANK_PROCESSORS[bank]
        rows = processor(payload.file_path, lookup_data)

        logger.info('文件解析完成: %s, 银行=%s, 记录数=%d',
                    payload.file_path, bank, len(rows))

        result_data = {
            'file_path': payload.file_path,
            'bank_type': bank,
            'records': rows,
            'record_count': len(rows),
            'status': 'success',
        }

        duration_ms = (time.perf_counter() - start_time) * 1000

        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.COMPLETED.value,
            result=result_data,
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )

    except Exception as e:
        logger.error('解析任务失败 %s (%s): %s',
                     payload.task_id, payload.file_path, e, exc_info=True)
        duration_ms = (time.perf_counter() - start_time) * 1000
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=str(e),
            result={
                'file_path': payload.file_path,
                'error': str(e),
                'status': 'error',
            },
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )


def task_merge_results(payload: MergeTaskPayload) -> TaskResult:
    """
    合并任务：合并所有解析任务的结果，处理增量合并

    Args:
        payload: 合并任务负载

    Returns:
        TaskResult 任务执行结果
    """
    from .mq_abstract import get_message_queue

    start_time = time.perf_counter()
    started_at = datetime.now().isoformat()

    try:
        bankcheck = _import_bankcheck()
        mq = get_message_queue()

        all_rows: List[Dict[str, Any]] = []
        processed_files: List[str] = []
        unprocessed_files: List[str] = []
        error_files: List[Tuple[str, str]] = []

        for parse_task_id in payload.parse_task_ids:
            result = mq.get_result(parse_task_id)
            if not result:
                logger.warning('未找到解析任务结果: %s', parse_task_id)
                continue

            if result.status == TaskStatus.FAILED.value:
                file_path = result.result.get('file_path', 'unknown')
                error_files.append((file_path, result.error or 'Unknown error'))
                continue

            file_path = result.result.get('file_path', '')
            status = result.result.get('status', '')

            if status == 'unrecognized':
                unprocessed_files.append(file_path)
            elif status == 'success':
                records = result.result.get('records', [])
                all_rows.extend(records)
                processed_files.append(file_path)

        existing_keys = set()
        existing_records: List[Dict[str, Any]] = []
        actual_incremental = False
        duplicate_count = 0
        new_record_count = 0
        output_path = None
        final_rows: List[Dict[str, Any]] = []

        if payload.incremental:
            summary_path = bankcheck.get_summary_table_path(payload.script_dir)
            existing_keys, existing_records = bankcheck.load_existing_keys(summary_path)
            actual_incremental = len(existing_records) > 0
            if actual_incremental:
                logger.info('增量合并模式已启用')

        if all_rows:
            if actual_incremental:
                incremental_rows, duplicate_count = bankcheck.filter_incremental_records(
                    all_rows, existing_keys
                )
                new_record_count = len(incremental_rows)
                output_path = bankcheck.merge_and_export_summary(
                    existing_records, incremental_rows, payload.script_dir,
                    lookup_source=payload.lookup_source
                )
                final_rows = existing_records + incremental_rows
            else:
                columns = bankcheck.get_summary_columns(all_rows, payload.lookup_source)
                df = bankcheck.pd.DataFrame(all_rows, columns=columns)
                output_path = bankcheck.get_summary_table_path(payload.script_dir)
                df.to_excel(output_path, index=False, engine='openpyxl')
                logger.info('总表输出完成: %s（共 %d 条记录）', output_path, len(all_rows))
                final_rows = all_rows
                new_record_count = len(all_rows)
        else:
            logger.warning('未提取到任何银行流水记录')
            if existing_records:
                output_path = bankcheck.merge_and_export_summary(
                    existing_records, [], payload.script_dir,
                    lookup_source=payload.lookup_source
                )
                final_rows = existing_records

        if final_rows:
            final_rows, _cp_tag_summary = bankcheck.apply_counterparty_rules(
                final_rows, payload.script_dir
            )
            if _cp_tag_summary.get('tagged_count', 0) > 0:
                logger.info('对方户名黑白名单打标完成')
                if output_path:
                    base_columns = bankcheck.get_summary_columns(final_rows, payload.lookup_source)
                    cp_extra_cols = ['黑白名单标签', '命中规则名称', '命中关键词']
                    _cp_columns = base_columns + [
                        col for col in cp_extra_cols if col not in base_columns
                    ]
                    bankcheck.pd.DataFrame(final_rows, columns=_cp_columns).to_excel(
                        output_path, index=False, engine='openpyxl'
                    )
                    logger.info('已将黑白名单打标结果回写到总表')

        result_data = {
            'all_rows': final_rows,
            'processed_files': processed_files,
            'unprocessed_files': unprocessed_files,
            'error_files': error_files,
            'output_path': output_path,
            'incremental_mode': actual_incremental,
            'existing_record_count': len(existing_records),
            'new_record_count': new_record_count,
            'duplicate_record_count': duplicate_count,
            'total_records': len(final_rows),
        }

        duration_ms = (time.perf_counter() - start_time) * 1000

        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.COMPLETED.value,
            result=result_data,
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )

    except Exception as e:
        logger.error('合并任务失败 %s: %s', payload.task_id, e, exc_info=True)
        duration_ms = (time.perf_counter() - start_time) * 1000
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=str(e),
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )


def task_generate_reports(payload: ReportTaskPayload) -> TaskResult:
    """
    报告任务：生成各类分析报告

    Args:
        payload: 报告任务负载

    Returns:
        TaskResult 任务执行结果
    """
    start_time = time.perf_counter()
    started_at = datetime.now().isoformat()

    try:
        bankcheck = _import_bankcheck()

        records = payload.records
        output_dir = payload.script_dir
        if payload.output_path:
            output_dir = os.path.dirname(payload.output_path) or payload.script_dir

        source_info = payload.source_info or {
            '数据来源': '异步任务生成',
            '总表文件': os.path.basename(payload.output_path) if payload.output_path else '内存数据',
            '记录数': len(records),
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }

        report_paths: Dict[str, Optional[str]] = {
            'subject_summary': None,
            'balance_check': None,
            'duplicate_check': None,
            'accounting_period': None,
        }

        if 'subject_summary' in payload.report_types and records:
            try:
                report_paths['subject_summary'] = bankcheck.generate_subject_summary_from_records(
                    records, output_dir, source_info
                )
                logger.info('主体维度汇总分析已生成: %s', report_paths['subject_summary'])
            except Exception as e:
                logger.error('生成主体汇总分析失败: %s', e)

        if 'balance_check' in payload.report_types and records:
            try:
                report_paths['balance_check'] = bankcheck.generate_balance_check_from_records(
                    records, output_dir, source_info
                )
                logger.info('余额连续性校验报告已生成: %s', report_paths['balance_check'])
            except Exception as e:
                logger.error('生成余额连续性校验报告失败: %s', e)

        if 'duplicate_check' in payload.report_types and records:
            try:
                report_paths['duplicate_check'] = bankcheck.generate_duplicate_check_from_records(
                    records, output_dir, source_info
                )
                logger.info('重复交易检测报告已生成: %s', report_paths['duplicate_check'])
            except Exception as e:
                logger.error('生成重复交易检测报告失败: %s', e)

        if 'accounting_period' in payload.report_types and records:
            try:
                report_paths['accounting_period'] = bankcheck.generate_accounting_period_report(
                    records, output_dir, source_info
                )
                logger.info('会计期间总表已生成: %s', report_paths['accounting_period'])
            except Exception as e:
                logger.error('生成会计期间总表失败: %s', e)

        result_data = {
            'report_paths': report_paths,
            'record_count': len(records),
            'output_dir': output_dir,
        }

        duration_ms = (time.perf_counter() - start_time) * 1000

        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.COMPLETED.value,
            result=result_data,
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )

    except Exception as e:
        logger.error('报告任务失败 %s: %s', payload.task_id, e, exc_info=True)
        duration_ms = (time.perf_counter() - start_time) * 1000
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=str(e),
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )


def task_persist_data(payload: PersistTaskPayload) -> TaskResult:
    """
    持久化任务：将记录写入数据库

    Args:
        payload: 持久化任务负载

    Returns:
        TaskResult 任务执行结果
    """
    start_time = time.perf_counter()
    started_at = datetime.now().isoformat()

    try:
        db_module = _import_database()

        if db_module is None:
            logger.warning('数据库模块不可用，跳过持久化')
            result_data = {
                'db_inserted': 0,
                'db_duplicates': 0,
                'skipped': True,
                'reason': 'database_module_not_available',
            }
            duration_ms = (time.perf_counter() - start_time) * 1000
            return TaskResult(
                task_id=payload.task_id,
                job_id=payload.job_id,
                task_type=payload.task_type,
                status=TaskStatus.COMPLETED.value,
                result=result_data,
                started_at=started_at,
                finished_at=datetime.now().isoformat(),
                duration_ms=duration_ms,
            )

        if not payload.records:
            logger.info('没有需要持久化的记录')
            result_data = {
                'db_inserted': 0,
                'db_duplicates': 0,
                'skipped': True,
                'reason': 'no_records',
            }
            duration_ms = (time.perf_counter() - start_time) * 1000
            return TaskResult(
                task_id=payload.task_id,
                job_id=payload.job_id,
                task_type=payload.task_type,
                status=TaskStatus.COMPLETED.value,
                result=result_data,
                started_at=started_at,
                finished_at=datetime.now().isoformat(),
                duration_ms=duration_ms,
            )

        db_inserted, db_duplicates = db_module.persist_transactions(
            payload.records,
            batch_id=payload.batch_id,
            deduplicate=payload.deduplicate,
            script_dir=payload.script_dir,
        )

        logger.info(
            '数据库持久化完成: 批次 %s, 插入 %d 条, 去重跳过 %d 条',
            payload.batch_id, db_inserted, db_duplicates,
        )

        result_data = {
            'batch_id': payload.batch_id,
            'db_inserted': db_inserted,
            'db_duplicates': db_duplicates,
            'total_records': len(payload.records),
        }

        duration_ms = (time.perf_counter() - start_time) * 1000

        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.COMPLETED.value,
            result=result_data,
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )

    except Exception as e:
        logger.error('持久化任务失败 %s: %s', payload.task_id, e, exc_info=True)
        duration_ms = (time.perf_counter() - start_time) * 1000
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=str(e),
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )


def task_cleanup(payload: CleanupTaskPayload) -> TaskResult:
    """
    清理任务：根据策略删除或归档已处理文件

    Args:
        payload: 清理任务负载

    Returns:
        TaskResult 任务执行结果
    """
    start_time = time.perf_counter()
    started_at = datetime.now().isoformat()

    try:
        bankcheck = _import_bankcheck()

        error_file_paths = {f for f, _ in payload.error_files}
        unprocessed_set = set(payload.unprocessed_files) | error_file_paths

        if payload.strategy == 'keep_unprocessed':
            bankcheck.delete_processed_files(
                payload.excel_files, payload.processed_files, payload.error_files,
                payload.unprocessed_files, strategy='keep_unprocessed'
            )
        elif payload.strategy == 'keep_all':
            bankcheck.delete_processed_files(
                payload.excel_files, payload.processed_files, payload.error_files,
                payload.unprocessed_files, strategy='keep_all'
            )
        elif payload.strategy == 'delete_all':
            bankcheck.delete_processed_files(
                payload.excel_files, payload.processed_files, payload.error_files,
                payload.unprocessed_files, strategy='delete_all'
            )
        elif payload.strategy == 'move_to_archive':
            bankcheck.delete_processed_files(
                payload.excel_files, payload.processed_files, payload.error_files,
                payload.unprocessed_files, strategy='move_to_archive',
                archive_dir_name=payload.archive_dir_name
            )

        logger.info('文件清理完成，策略=%s', payload.strategy)

        result_data = {
            'working_folder': payload.working_folder,
            'strategy': payload.strategy,
            'processed_files_count': len(payload.processed_files),
            'error_files_count': len(payload.error_files),
            'unprocessed_files_count': len(payload.unprocessed_files),
        }

        duration_ms = (time.perf_counter() - start_time) * 1000

        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.COMPLETED.value,
            result=result_data,
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )

    except Exception as e:
        logger.error('清理任务失败 %s: %s', payload.task_id, e, exc_info=True)
        duration_ms = (time.perf_counter() - start_time) * 1000
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=str(e),
            started_at=started_at,
            finished_at=datetime.now().isoformat(),
            duration_ms=duration_ms,
        )


TASK_HANDLERS = {
    'scan': task_scan_files,
    'parse': task_parse_file,
    'merge': task_merge_results,
    'report': task_generate_reports,
    'persist': task_persist_data,
    'cleanup': task_cleanup,
}


def execute_task(payload: TaskPayload, worker_id: str = '') -> TaskResult:
    """
    通用任务执行入口

    Args:
        payload: 任务负载
        worker_id: Worker ID

    Returns:
        TaskResult 任务执行结果
    """
    handler = TASK_HANDLERS.get(payload.task_type)
    if handler is None:
        logger.error('未知任务类型: %s', payload.task_type)
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=f"Unknown task type: {payload.task_type}",
            started_at=datetime.now().isoformat(),
            finished_at=datetime.now().isoformat(),
            duration_ms=0,
            worker_id=worker_id,
        )

    logger.info('开始执行任务: %s (type=%s, job=%s)',
                payload.task_id, payload.task_type, payload.job_id)

    try:
        result = handler(payload)
        result.worker_id = worker_id
        return result
    except Exception as e:
        logger.error('任务执行异常 %s: %s', payload.task_id, e, exc_info=True)
        return TaskResult(
            task_id=payload.task_id,
            job_id=payload.job_id,
            task_type=payload.task_type,
            status=TaskStatus.FAILED.value,
            error=str(e),
            started_at=datetime.now().isoformat(),
            finished_at=datetime.now().isoformat(),
            duration_ms=0,
            worker_id=worker_id,
        )
