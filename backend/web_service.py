# -*- coding: utf-8 -*-
"""
银行流水检验工具 - Web 上传处理服务
基于 Flask 提供浏览器端文件夹上传、异步处理与总表下载能力，
将现有桌面工具扩展为团队可共用的内网服务。

功能：
  1. 浏览器端多文件/文件夹上传（保留文件夹目录结构）
  2. 后台异步处理流水线（调用 bankcheck.run_pipeline）
  3. SSE 实时进度推送（替代轮询）
  4. 任务取消
  5. 总表文件下载
  6. 历史任务列表与状态查询
  7. 并发任务数限制
"""

import os
import sys
import re
import uuid
import json
import shutil
import logging
import threading
import time
import queue
from datetime import datetime
from typing import Dict, Any, Optional, List
from dataclasses import dataclass, field, asdict

from flask import (Flask, render_template, request, jsonify,
                   send_file, redirect, url_for, Response)

import bankcheck
import batch_manager as batch_module
import database as db_module
import workflow as workflow_module

try:
    from pii_classifier import setup_pii_aware_logging, PIILogFilter
    HAS_PII_CLASSIFIER = True
except ImportError:
    HAS_PII_CLASSIFIER = False

try:
    from task_queue import (
        get_job_orchestrator,
        JobContext,
        TaskStatus,
    )
    HAS_TASK_QUEUE = True
except ImportError as e:
    HAS_TASK_QUEUE = False


BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BACKEND_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BACKEND_DIR, 'output')
TASKS_DB_PATH = os.path.join(BACKEND_DIR, 'web_tasks.json')

MAX_CONCURRENT_TASKS = 2

USE_TASK_QUEUE = os.environ.get('USE_TASK_QUEUE', 'false').lower() == 'true'

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

_job_orchestrator = None


def _get_job_orchestrator():
    """获取作业编排器实例"""
    global _job_orchestrator
    if _job_orchestrator is None and HAS_TASK_QUEUE and USE_TASK_QUEUE:
        try:
            _job_orchestrator = get_job_orchestrator()
            logger.info('作业编排器已初始化')
        except Exception as e:
            logger.error('初始化作业编排器失败: %s', e)
    return _job_orchestrator


def _job_to_web_task(job_context: JobContext) -> WebTask:
    """将 JobContext 转换为 WebTask"""
    status_map = {
        TaskStatus.PENDING.value: 'pending',
        TaskStatus.QUEUED.value: 'pending',
        TaskStatus.RUNNING.value: 'processing',
        TaskStatus.COMPLETED.value: 'completed',
        TaskStatus.FAILED.value: 'failed',
        TaskStatus.CANCELLED.value: 'cancelled',
        TaskStatus.TIMEOUT.value: 'failed',
    }

    web_task = WebTask(
        task_id=job_context.job_id,
        status=status_map.get(job_context.status, 'pending'),
        created_at=job_context.created_at,
        started_at=job_context.started_at,
        finished_at=job_context.finished_at,
        upload_folder=job_context.input_folder,
        output_path=job_context.output_path,
        incremental=job_context.incremental,
        operator=job_context.operator,
        total_files=job_context.total_files,
        processed_files=job_context.processed_files,
        error_files=job_context.error_files,
        unprocessed_files=job_context.unprocessed_files,
        total_records=job_context.total_records,
        new_records=job_context.new_records,
        duplicate_records=job_context.duplicate_records,
        db_inserted=job_context.metadata.get('db_inserted', 0),
        db_duplicates=job_context.metadata.get('db_duplicates', 0),
        error_message=job_context.error_message,
        batch_id=job_context.batch_id,
        progress_stage=job_context.progress_stage,
        progress_percent=job_context.progress_percent,
        progress_message=job_context.progress_message,
    )

    if job_context.subject_summary_path:
        web_task.file_names.append(job_context.subject_summary_path)
    if job_context.balance_check_path:
        web_task.file_names.append(job_context.balance_check_path)
    if job_context.duplicate_check_path:
        web_task.file_names.append(job_context.duplicate_check_path)
    if job_context.accounting_period_path:
        web_task.file_names.append(job_context.accounting_period_path)

    return web_task


def _notify_sse_from_job(job_context: JobContext):
    """将作业进度通过 SSE 推送给前端"""
    web_task = _job_to_web_task(job_context)
    _save_task(web_task)
    _sse_push(job_context.job_id, {
        'type': 'progress',
        'stage': job_context.progress_stage,
        'percent': job_context.progress_percent,
        'message': job_context.progress_message,
        'status': web_task.status,
        'task_id': job_context.job_id,
    })

    if web_task.status in ('completed', 'failed', 'cancelled'):
        _sse_push(job_context.job_id, {
            'type': web_task.status,
            'status': web_task.status,
            'task_id': job_context.job_id,
            'data': web_task.to_dict(),
        })
        if job_context.job_id in _cancel_events:
            del _cancel_events[job_context.job_id]

_task_semaphore = threading.Semaphore(MAX_CONCURRENT_TASKS)
_cancel_events: Dict[str, threading.Event] = {}
_sse_queues: Dict[str, List[queue.Queue]] = {}
_sse_lock = threading.Lock()


def setup_logging():
    log_file = os.path.join(BACKEND_DIR, 'web_service.log')

    if HAS_PII_CLASSIFIER:
        logger = setup_pii_aware_logging(
            logger_name='bankcheck.web',
            log_file=log_file,
            console_level=logging.INFO,
            file_level=logging.DEBUG,
        )
    else:
        logger = logging.getLogger('bankcheck.web')
        logger.setLevel(logging.DEBUG)

        if not logger.handlers:
            fmt = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S',
            )
            fh = logging.FileHandler(log_file, encoding='utf-8')
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(fmt)
            ch = logging.StreamHandler(sys.stdout)
            ch.setLevel(logging.INFO)
            ch.setFormatter(fmt)
            logger.addHandler(fh)
            logger.addHandler(ch)

    if not HAS_TASK_QUEUE:
        logger.warning('任务队列模块不可用')

    return logger


logger = setup_logging()


app = Flask(__name__,
            template_folder=os.path.join(BACKEND_DIR, 'templates'),
            static_folder=os.path.join(BACKEND_DIR, 'static'))
app.secret_key = 'bankcheck_web_service_secret_2024'
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024


@dataclass
class WebTask:
    task_id: str
    status: str = 'pending'
    created_at: str = ''
    started_at: Optional[str] = None
    finished_at: Optional[str] = None
    upload_folder: str = ''
    output_path: Optional[str] = None
    masked_output_path: Optional[str] = None
    incremental: bool = True
    operator: str = ''
    total_files: int = 0
    processed_files: int = 0
    error_files: int = 0
    unprocessed_files: int = 0
    total_records: int = 0
    new_records: int = 0
    duplicate_records: int = 0
    db_inserted: int = 0
    db_duplicates: int = 0
    error_message: Optional[str] = None
    batch_id: Optional[str] = None
    file_names: list = field(default_factory=list)
    progress_stage: str = 'pending'
    progress_percent: int = 0
    progress_message: str = ''

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        return d


def _load_tasks() -> Dict[str, Any]:
    if os.path.exists(TASKS_DB_PATH):
        try:
            with open(TASKS_DB_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_tasks(tasks: Dict[str, Any]):
    with open(TASKS_DB_PATH, 'w', encoding='utf-8') as f:
        json.dump(tasks, f, ensure_ascii=False, indent=2)


def _save_task(task: WebTask):
    tasks = _load_tasks()
    tasks[task.task_id] = task.to_dict()
    _save_tasks(tasks)


def _get_task(task_id: str) -> Optional[WebTask]:
    tasks = _load_tasks()
    data = tasks.get(task_id)
    if data is None:
        return None
    filtered = {k: v for k, v in data.items() if k in WebTask.__dataclass_fields__}
    return WebTask(**filtered)


def _sse_register(task_id: str) -> queue.Queue:
    q = queue.Queue(maxsize=64)
    with _sse_lock:
        if task_id not in _sse_queues:
            _sse_queues[task_id] = []
        _sse_queues[task_id].append(q)
    return q


def _sse_unregister(task_id: str, q: queue.Queue):
    with _sse_lock:
        if task_id in _sse_queues:
            try:
                _sse_queues[task_id].remove(q)
            except ValueError:
                pass
            if not _sse_queues[task_id]:
                del _sse_queues[task_id]


def _sse_push(task_id: str, data: Dict[str, Any]):
    with _sse_lock:
        queues = _sse_queues.get(task_id, [])
    for q in queues:
        try:
            q.put_nowait(data)
        except queue.Full:
            pass


def _update_task_progress(task: WebTask, stage: str, percent: int, message: str):
    task.progress_stage = stage
    task.progress_percent = percent
    task.progress_message = message
    _save_task(task)
    _sse_push(task.task_id, {
        'type': 'progress',
        'stage': stage,
        'percent': percent,
        'message': message,
        'status': task.status,
        'task_id': task.task_id,
    })


def _run_pipeline_async(task: WebTask, cancel_event: threading.Event):
    acquired = _task_semaphore.acquire(blocking=True, timeout=300)
    if not acquired:
        task.status = 'failed'
        task.error_message = '等待队列已满，请稍后重试'
        task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        _save_task(task)
        _sse_push(task.task_id, {
            'type': 'error',
            'message': task.error_message,
            'status': 'failed',
            'task_id': task.task_id,
        })
        return

    try:
        if cancel_event.is_set():
            task.status = 'cancelled'
            task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            task.progress_message = '任务已取消'
            _save_task(task)
            _sse_push(task.task_id, {
                'type': 'cancelled',
                'status': 'cancelled',
                'task_id': task.task_id,
            })
            return

        task.status = 'processing'
        task.started_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        _update_task_progress(task, 'scanning', 10, '正在扫描上传文件...')

        bm = batch_module.get_batch_manager(BACKEND_DIR)
        batch_info = bm.start_batch(
            input_folder=task.upload_folder,
            operator=task.operator or 'web',
        )
        task.batch_id = batch_info.batch_id
        _save_task(task)

        if cancel_event.is_set():
            task.status = 'cancelled'
            task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            task.progress_message = '任务已取消'
            _save_task(task)
            _sse_push(task.task_id, {
                'type': 'cancelled',
                'status': 'cancelled',
                'task_id': task.task_id,
            })
            return

        _update_task_progress(task, 'processing', 30, '正在解析银行流水文件...')

        result = bankcheck.run_pipeline(
            folder=task.upload_folder,
            script_dir=BACKEND_DIR,
            incremental=task.incremental,
            batch_id=batch_info.batch_id,
        )

        if cancel_event.is_set():
            task.status = 'cancelled'
            task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            task.progress_message = '任务已取消（流水线已完成执行）'
            _save_task(task)
            _sse_push(task.task_id, {
                'type': 'cancelled',
                'status': 'cancelled',
                'task_id': task.task_id,
            })
            return

        _update_task_progress(task, 'exporting', 70, '正在导出总表...')

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        task_output_dir = os.path.join(OUTPUT_DIR, task.task_id)
        os.makedirs(task_output_dir, exist_ok=True)

        if result.output_path and os.path.exists(result.output_path):
            output_filename = f'银行流水总表_{timestamp}.xlsx'
            dest = os.path.join(task_output_dir, output_filename)
            shutil.copy2(result.output_path, dest)
            task.output_path = dest
            logger.info('总表已复制到: %s', dest)

        if result.masked_output_path and os.path.exists(result.masked_output_path):
            masked_filename = f'银行流水总表_脱敏版_{timestamp}.xlsx'
            masked_dest = os.path.join(task_output_dir, masked_filename)
            shutil.copy2(result.masked_output_path, masked_dest)
            task.masked_output_path = masked_dest
            logger.info('脱敏版总表已复制到: %s', masked_dest)

        task.total_files = len(result.processed_files)
        task.processed_files = len(result.processed_files)
        task.error_files = len(result.error_files)
        task.unprocessed_files = len(result.unprocessed_files)
        task.total_records = len(result.all_rows)
        task.new_records = result.new_record_count
        task.duplicate_records = result.duplicate_record_count
        task.db_inserted = result.db_inserted_count
        task.db_duplicates = result.db_duplicate_count

        result_data = {
            'total_records': len(result.all_rows),
            'new_records': result.new_record_count,
            'duplicate_records': result.duplicate_record_count,
            'processed_files': result.processed_files,
            'unprocessed_files': result.unprocessed_files,
            'error_files': [(f, e) for f, e in result.error_files],
            'incremental_mode': result.incremental_mode,
            'output_folder': task_output_dir,
            'summary_table_path': result.output_path or '',
            'log_file_path': os.path.join(BACKEND_DIR, 'bankcheck.log'),
        }
        bm.finish_batch(batch_info.batch_id, result_data, status='success')

        _update_task_progress(task, 'completed', 100,
                              f'处理完成：{task.processed_files} 个文件，'
                              f'{task.new_records} 条新增记录')

        task.status = 'completed'
        task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        _save_task(task)
        _sse_push(task.task_id, {
            'type': 'completed',
            'status': 'completed',
            'task_id': task.task_id,
            'data': task.to_dict(),
        })
        logger.info('任务 %s 处理完成', task.task_id)

    except Exception as e:
        logger.error('任务 %s 处理失败: %s', task.task_id, e, exc_info=True)
        task.status = 'failed'
        task.error_message = str(e)
        task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        task.progress_stage = 'failed'
        task.progress_message = f'处理失败: {str(e)}'
        _save_task(task)
        _sse_push(task.task_id, {
            'type': 'error',
            'message': str(e),
            'status': 'failed',
            'task_id': task.task_id,
        })

        if task.batch_id:
            try:
                bm = batch_module.get_batch_manager(BACKEND_DIR)
                bm._update_batch(batch_module.BatchInfo(
                    batch_id=task.batch_id,
                    start_time=task.started_at or '',
                    end_time=task.finished_at,
                    status='failed',
                    error_message=str(e),
                ))
            except Exception:
                pass

    finally:
        _task_semaphore.release()
        try:
            if os.path.exists(task.upload_folder):
                shutil.rmtree(task.upload_folder)
                logger.info('已清理上传临时目录: %s', task.upload_folder)
        except Exception:
            pass
        _cancel_events.pop(task.task_id, None)


@app.route('/')
def index():
    return redirect(url_for('upload_page'))


@app.route('/upload')
def upload_page():
    tasks = _load_tasks()
    task_list = sorted(tasks.values(), key=lambda t: t.get('created_at', ''), reverse=True)[:20]
    return render_template('upload.html', recent_tasks=task_list)


@app.route('/counterparty-rules')
def counterparty_rules_page():
    return render_template('counterparty_rules.html')


@app.route('/voucher-attachments')
def voucher_attachments_page():
    return render_template('voucher_attachments.html')


@app.route('/dashboard')
def dashboard_page():
    return render_template('dashboard.html')


@app.route('/column-mapping')
def column_mapping_page():
    return render_template('column_mapping.html')


WIZARD_UPLOAD_DIR = os.path.join(UPLOAD_DIR, 'wizard')
os.makedirs(WIZARD_UPLOAD_DIR, exist_ok=True)


def _wizard_safe_filename(filename):
    safe = filename.replace('\\', '/').replace('../', '').replace('..\\', '')
    safe = os.path.basename(safe)
    safe = re.sub(r'[\x00-\x1f\x7f]', '', safe)
    if not safe:
        safe = 'unnamed'
    return safe


@app.route('/api/wizard/upload', methods=['POST'])
def api_wizard_upload():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'}), 400
    f = request.files['file']
    if not f or not f.filename:
        return jsonify({'success': False, 'message': '未选择文件'}), 400
    fname = f.filename
    if not (fname.lower().endswith('.xlsx') or fname.lower().endswith('.xls')):
        return jsonify({'success': False, 'message': '仅支持 .xlsx / .xls 格式'}), 400
    safe_name = _wizard_safe_filename(fname)
    wizard_id = f"WZ{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6]}"
    save_dir = os.path.join(WIZARD_UPLOAD_DIR, wizard_id)
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, safe_name)
    f.save(save_path)
    logger.info('向导文件已上传: wizard_id=%s, file=%s', wizard_id, safe_name)
    return jsonify({
        'success': True,
        'wizard_id': wizard_id,
        'filename': safe_name,
        'message': '文件上传成功',
    })


@app.route('/api/wizard/preview', methods=['GET'])
def api_wizard_preview():
    wizard_id = request.args.get('wizard_id', '').strip()
    sheet_name = request.args.get('sheet')
    if not wizard_id:
        return jsonify({'success': False, 'message': '缺少 wizard_id'}), 400
    save_dir = os.path.join(WIZARD_UPLOAD_DIR, wizard_id)
    if not os.path.isdir(save_dir):
        return jsonify({'success': False, 'message': '会话已过期，请重新上传文件'}), 404
    files = [f for f in os.listdir(save_dir)
             if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    if not files:
        return jsonify({'success': False, 'message': '文件不存在'}), 404
    filepath = os.path.join(save_dir, files[0])
    try:
        data = bankcheck.read_excel_preview(
            filepath, sheet_name=sheet_name if sheet_name else None,
            max_rows=50, max_cols=30,
        )
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        logger.error('读取 Excel 预览失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/wizard/extract-preview', methods=['POST'])
def api_wizard_extract_preview():
    body = request.get_json(silent=True) or {}
    wizard_id = body.get('wizard_id', '').strip()
    rule_data = body.get('rule_data', {})
    sheet_name = body.get('sheet_name')
    if not wizard_id:
        return jsonify({'success': False, 'message': '缺少 wizard_id'}), 400
    save_dir = os.path.join(WIZARD_UPLOAD_DIR, wizard_id)
    if not os.path.isdir(save_dir):
        return jsonify({'success': False, 'message': '会话已过期，请重新上传文件'}), 404
    files = [f for f in os.listdir(save_dir)
             if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    if not files:
        return jsonify({'success': False, 'message': '文件不存在'}), 404
    filepath = os.path.join(save_dir, files[0])
    result = bankcheck.preview_extraction(
        filepath, rule_data, sheet_name=sheet_name, max_preview_rows=10,
    )
    if result.get('success'):
        return jsonify({'success': True, 'data': result})
    return jsonify({'success': False, 'message': result.get('error', '预览失败')}), 400


@app.route('/api/wizard/infer-mapping', methods=['POST'])
def api_wizard_infer_mapping():
    """自动推断银行流水列映射草案"""
    body = request.get_json(silent=True) or {}
    wizard_id = body.get('wizard_id', '').strip()
    sheet_name = body.get('sheet_name')
    if not wizard_id:
        return jsonify({'success': False, 'message': '缺少 wizard_id'}), 400
    save_dir = os.path.join(WIZARD_UPLOAD_DIR, wizard_id)
    if not os.path.isdir(save_dir):
        return jsonify({'success': False, 'message': '会话已过期，请重新上传文件'}), 404
    files = [f for f in os.listdir(save_dir)
             if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    if not files:
        return jsonify({'success': False, 'message': '文件不存在'}), 404
    filepath = os.path.join(save_dir, files[0])
    try:
        from bank_template_inferrer import scan_workbook
        result = scan_workbook(filepath, sheet_name=sheet_name)
        if result.get('success'):
            logger.info('自动推断完成: wizard_id=%s, 置信度=%.2f, 映射字段=%d',
                        wizard_id, result.get('confidence', 0),
                        len(result.get('column_map', {})))
            return jsonify({'success': True, 'data': result})
        return jsonify({
            'success': False,
            'message': result.get('error', '推断失败'),
            'data': result,
        }), 400
    except Exception as e:
        logger.error('自动推断失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/wizard/confirm-mapping', methods=['POST'])
def api_wizard_confirm_mapping():
    """确认推断映射并保存为银行配置"""
    body = request.get_json(silent=True) or {}
    wizard_id = body.get('wizard_id', '').strip()
    bank_name = (body.get('bank_name') or '').strip()
    overrides = body.get('overrides', {})
    if not wizard_id:
        return jsonify({'success': False, 'message': '缺少 wizard_id'}), 400
    if not bank_name:
        return jsonify({'success': False, 'message': '银行名称不能为空'}), 400
    save_dir = os.path.join(WIZARD_UPLOAD_DIR, wizard_id)
    if not os.path.isdir(save_dir):
        return jsonify({'success': False, 'message': '会话已过期，请重新上传文件'}), 404
    files = [f for f in os.listdir(save_dir)
             if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    if not files:
        return jsonify({'success': False, 'message': '文件不存在'}), 404
    filepath = os.path.join(save_dir, files[0])
    try:
        from bank_template_inferrer import scan_workbook, confirm_and_save
        inferred = scan_workbook(filepath,
                                 sheet_name=body.get('sheet_name'))
        if not inferred.get('success'):
            return jsonify({
                'success': False,
                'message': inferred.get('error', '推断失败'),
            }), 400
        ok, message = confirm_and_save(inferred, bank_name, overrides=overrides)
        if ok:
            logger.info('银行配置已确认保存: %s', bank_name)
            return jsonify({'success': True, 'message': message})
        return jsonify({'success': False, 'message': message}), 400
    except Exception as e:
        logger.error('确认保存失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/bank-rules', methods=['GET'])
def api_list_bank_rules():
    try:
        config = bankcheck.get_bank_config()
        rules = config.list_rules_detailed()
        return jsonify({'success': True, 'data': rules, 'total': len(rules)})
    except Exception as e:
        logger.error('获取银行规则列表失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/bank-rules', methods=['POST'])
def api_save_bank_rule():
    body = request.get_json(silent=True) or {}
    bank_name = (body.get('bank_name') or '').strip()
    account_cell = (body.get('account_cell') or '').strip()
    start_row = body.get('start_row')
    columns = body.get('columns') or {}

    if not bank_name:
        return jsonify({'success': False, 'message': '银行名称不能为空'}), 400
    if not account_cell:
        return jsonify({'success': False, 'message': '请选择账号单元格'}), 400
    if not start_row:
        return jsonify({'success': False, 'message': '请设置数据起始行'}), 400
    if not columns.get('trade_date'):
        return jsonify({'success': False, 'message': '请至少映射交易日期列'}), 400

    try:
        bankcheck.parse_cell_ref(account_cell)
    except Exception as e:
        return jsonify({'success': False, 'message': f'账号单元格格式错误: {e}'}), 400

    try:
        start_row_int = int(start_row)
        if start_row_int < 1:
            raise ValueError('起始行必须 >= 1')
    except Exception as e:
        return jsonify({'success': False, 'message': f'数据起始行无效: {e}'}), 400

    rule_data = {
        'bank_name': bank_name,
        'account_cell': account_cell,
        'start_row': start_row_int,
        'columns': {k: int(v) for k, v in columns.items() if v},
        'payment_sign': body.get('payment_sign', 'negative'),
        'enabled': bool(body.get('enabled', True)),
        'expected_headers': body.get('expected_headers') or {},
        'header_validation': body.get('header_validation', 'warn'),
        'multi_account': bool(body.get('multi_account', False)),
        'skip_sheets': body.get('skip_sheets') or [],
    }
    try:
        config = bankcheck.get_bank_config()
        ok = config.save_rule(rule_data)
        if ok:
            logger.info('银行规则已保存: %s', bank_name)
            return jsonify({'success': True, 'message': '银行规则保存成功'})
        return jsonify({'success': False, 'message': '保存失败，请检查日志'}), 500
    except Exception as e:
        logger.error('保存银行规则失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/bank-rules/<bank_name>', methods=['DELETE'])
def api_delete_bank_rule(bank_name):
    try:
        config = bankcheck.get_bank_config()
        ok = config.delete_rule(bank_name)
        if ok:
            logger.info('银行规则已删除: %s', bank_name)
            return jsonify({'success': True, 'message': '银行规则已删除'})
        return jsonify({'success': False, 'message': '银行规则不存在'}), 404
    except Exception as e:
        logger.error('删除银行规则失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


def _safe_filename(filename):
    safe = filename.replace('\\', '/').replace('../', '').replace('..\\', '')
    safe = os.path.basename(safe)
    safe = re.sub(r'[\x00-\x1f\x7f]', '', safe)
    if not safe:
        safe = 'unnamed'
    return safe


def _run_pipeline_with_task_queue(task: WebTask, cancel_event: threading.Event):
    """使用任务队列异步处理流水线"""
    try:
        orchestrator = _get_job_orchestrator()
        if orchestrator is None:
            raise Exception('任务队列不可用')

        def progress_callback(job_context: JobContext):
            if cancel_event.is_set():
                orchestrator.cancel_job(job_context.job_id)
                return
            _notify_sse_from_job(job_context)

        job_context = orchestrator.submit_job(
            input_folder=task.upload_folder,
            script_dir=BACKEND_DIR,
            incremental=task.incremental,
            keep_strategy='keep_unprocessed',
            operator=task.operator,
            user_id='web_user',
            on_progress=progress_callback,
        )

        task.task_id = job_context.job_id
        _cancel_events[job_context.job_id] = cancel_event
        _save_task(task)

        while True:
            if cancel_event.is_set():
                orchestrator.cancel_job(job_context.job_id)
                break

            current_job = orchestrator.get_job_status(job_context.job_id)
            if current_job and current_job.status in [
                TaskStatus.COMPLETED.value,
                TaskStatus.FAILED.value,
                TaskStatus.CANCELLED.value,
            ]:
                break

            time.sleep(1)

        final_job = orchestrator.get_job_status(job_context.job_id)
        if final_job:
            web_task = _job_to_web_task(final_job)
            _save_task(web_task)

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            task_output_dir = os.path.join(OUTPUT_DIR, web_task.task_id)
            os.makedirs(task_output_dir, exist_ok=True)

            if final_job.output_path and os.path.exists(final_job.output_path):
                output_filename = f'银行流水总表_{timestamp}.xlsx'
                dest = os.path.join(task_output_dir, output_filename)
                shutil.copy2(final_job.output_path, dest)
                web_task.output_path = dest

            _save_task(web_task)
            _notify_sse_from_job(final_job)

    except Exception as e:
        logger.error('任务队列处理失败 %s: %s', task.task_id, e, exc_info=True)
        task.status = 'failed'
        task.error_message = str(e)
        task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        _save_task(task)
        _sse_push(task.task_id, {
            'type': 'error',
            'message': str(e),
            'status': 'failed',
            'task_id': task.task_id,
        })


@app.route('/api/upload', methods=['POST'])
def api_upload():
    if 'files' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'}), 400

    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'success': False, 'message': '未选择文件'}), 400

    incremental = request.form.get('incremental', 'true').lower() == 'true'
    operator = request.form.get('operator', '').strip()
    use_async = request.form.get('use_async', 'false').lower() == 'true'

    task_id = f"WEB{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6]}"
    task_upload_dir = os.path.join(UPLOAD_DIR, task_id)
    os.makedirs(task_upload_dir, exist_ok=True)

    rel_paths = {}
    for key in request.form:
        if key.startswith('relpath_'):
            idx = key[len('relpath_'):]
            rel_paths[idx] = request.form[key]

    saved_count = 0
    file_names = []
    for i, f in enumerate(files):
        if f.filename == '':
            continue
        if f.filename.startswith('.'):
            continue
        if not (f.filename.lower().endswith('.xlsx') or f.filename.lower().endswith('.xls')):
            continue

        safe_name = _safe_filename(f.filename)
        rel_path = rel_paths.get(str(i), '')

        if rel_path and '/' in rel_path:
            parts = rel_path.split('/')
            sub_dirs = parts[:-1]
            file_dir = os.path.join(task_upload_dir, *sub_dirs)
            os.makedirs(file_dir, exist_ok=True)
        else:
            file_dir = task_upload_dir

        save_path = os.path.join(file_dir, safe_name)
        counter = 1
        base_name, ext = os.path.splitext(safe_name)
        while os.path.exists(save_path):
            save_path = os.path.join(file_dir, f"{base_name}_{counter}{ext}")
            counter += 1
        f.save(save_path)
        file_names.append(safe_name)
        saved_count += 1

    if saved_count == 0:
        shutil.rmtree(task_upload_dir, ignore_errors=True)
        return jsonify({'success': False, 'message': '未检测到有效的 Excel 文件（.xlsx/.xls）'}), 400

    active_count = sum(
        1 for t in _load_tasks().values()
        if t.get('status') in ('pending', 'processing')
    )
    if active_count >= MAX_CONCURRENT_TASKS * 3:
        shutil.rmtree(task_upload_dir, ignore_errors=True)
        return jsonify({'success': False, 'message': '当前排队任务过多，请稍后重试'}), 429

    cancel_event = threading.Event()
    _cancel_events[task_id] = cancel_event

    task = WebTask(
        task_id=task_id,
        status='pending',
        created_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        upload_folder=task_upload_dir,
        incremental=incremental,
        operator=operator,
        total_files=saved_count,
        file_names=file_names,
        progress_stage='pending',
        progress_percent=5,
        progress_message='已上传，等待处理...',
    )
    _save_task(task)

    if use_async and HAS_TASK_QUEUE and USE_TASK_QUEUE:
        thread = threading.Thread(
            target=_run_pipeline_with_task_queue,
            args=(task, cancel_event),
            daemon=True
        )
        mode = '异步队列模式'
    else:
        thread = threading.Thread(
            target=_run_pipeline_async,
            args=(task, cancel_event),
            daemon=True
        )
        mode = '同步线程模式'

    thread.start()

    logger.info('任务 %s 已创建，共上传 %d 个文件，模式: %s', task_id, saved_count, mode)
    return jsonify({
        'success': True,
        'task_id': task_id,
        'message': f'已上传 {saved_count} 个文件，正在后台处理（{mode}）',
        'async_mode': use_async and HAS_TASK_QUEUE and USE_TASK_QUEUE,
    })


@app.route('/api/queue/status', methods=['GET'])
def api_queue_status():
    """获取任务队列状态"""
    if not HAS_TASK_QUEUE:
        return jsonify({
            'success': False,
            'message': '任务队列模块未安装',
        }), 503

    if not USE_TASK_QUEUE:
        return jsonify({
            'success': False,
            'message': '任务队列未启用（设置 USE_TASK_QUEUE=true 启用）',
        }), 503

    try:
        orchestrator = _get_job_orchestrator()
        if orchestrator is None:
            return jsonify({
                'success': False,
                'message': '任务队列未初始化',
            }), 503

        from task_queue import get_message_queue
        mq = get_message_queue()

        queue_status = {}
        for queue_name in ['scan', 'parse', 'merge', 'report', 'persist', 'cleanup', 'default']:
            try:
                queue_status[queue_name] = mq.get_queue_length(queue_name)
            except Exception:
                queue_status[queue_name] = 0

        running_jobs = orchestrator.list_jobs(status=TaskStatus.RUNNING.value)
        pending_jobs = orchestrator.list_jobs(status=TaskStatus.PENDING.value)
        completed_jobs = orchestrator.list_jobs(status=TaskStatus.COMPLETED.value, limit=10)

        return jsonify({
            'success': True,
            'enabled': True,
            'backend': mq.config.backend,
            'queues': queue_status,
            'running_jobs': len(running_jobs),
            'pending_jobs': len(pending_jobs),
            'recent_completed': [j.to_dict() for j in completed_jobs],
        })

    except Exception as e:
        logger.error('获取队列状态失败: %s', e)
        return jsonify({
            'success': False,
            'message': str(e),
        }), 500


@app.route('/api/queue/workers', methods=['GET'])
def api_worker_status():
    """获取 Worker 状态"""
    if not HAS_TASK_QUEUE or not USE_TASK_QUEUE:
        return jsonify({
            'success': False,
            'message': '任务队列未启用',
        }), 503

    try:
        from task_queue import WorkerManager, load_config
        config = load_config()
        manager = WorkerManager(config=config, num_workers=0)

        return jsonify({
            'success': True,
            'status': manager.get_status(),
            'config': {
                'worker_queues': config.worker.queues,
                'task_timeout': config.worker.task_timeout,
                'max_retries': config.max_retries,
            },
        })
    except Exception as e:
        logger.error('获取 Worker 状态失败: %s', e)
        return jsonify({
            'success': False,
            'message': str(e),
        }), 500


@app.route('/api/jobs/<job_id>', methods=['GET'])
def api_get_job(job_id):
    """获取作业详情"""
    if not HAS_TASK_QUEUE or not USE_TASK_QUEUE:
        return jsonify({
            'success': False,
            'message': '任务队列未启用',
        }), 503

    try:
        orchestrator = _get_job_orchestrator()
        if orchestrator is None:
            return jsonify({
                'success': False,
                'message': '任务队列未初始化',
            }), 503

        job_context = orchestrator.get_job_status(job_id)
        if job_context is None:
            return jsonify({
                'success': False,
                'message': '作业不存在',
            }), 404

        return jsonify({
            'success': True,
            'job': job_context.to_dict(),
        })

    except Exception as e:
        logger.error('获取作业详情失败: %s', e)
        return jsonify({
            'success': False,
            'message': str(e),
        }), 500


@app.route('/api/jobs/<job_id>/cancel', methods=['POST'])
def api_cancel_job(job_id):
    """取消作业"""
    if not HAS_TASK_QUEUE or not USE_TASK_QUEUE:
        return jsonify({
            'success': False,
            'message': '任务队列未启用',
        }), 503

    try:
        orchestrator = _get_job_orchestrator()
        if orchestrator is None:
            return jsonify({
                'success': False,
                'message': '任务队列未初始化',
            }), 503

        success = orchestrator.cancel_job(job_id)
        if not success:
            return jsonify({
                'success': False,
                'message': '取消作业失败（作业可能已完成或不存在）',
            }), 400

        return jsonify({
            'success': True,
            'message': '作业已取消',
        })

    except Exception as e:
        logger.error('取消作业失败: %s', e)
        return jsonify({
            'success': False,
            'message': str(e),
        }), 500


@app.route('/api/tasks/<task_id>/stream', methods=['GET'])
def api_task_stream(task_id):
    task = _get_task(task_id)
    if task is None:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    q = _sse_register(task_id)

    def generate():
        try:
            if task.status in ('completed', 'failed', 'cancelled'):
                yield f"data: {json.dumps({'type': task.status, 'status': task.status, 'task_id': task_id, 'data': task.to_dict()}, ensure_ascii=False)}\n\n"
                return

            yield f"data: {json.dumps({'type': 'connected', 'task_id': task_id}, ensure_ascii=False)}\n\n"

            while True:
                try:
                    data = q.get(timeout=30)
                    yield f"data: {json.dumps(data, ensure_ascii=False)}\n\n"
                    if data.get('type') in ('completed', 'failed', 'cancelled', 'error'):
                        break
                except queue.Empty:
                    current = _get_task(task_id)
                    if current and current.status in ('completed', 'failed', 'cancelled'):
                        yield f"data: {json.dumps({'type': current.status, 'status': current.status, 'task_id': task_id, 'data': current.to_dict()}, ensure_ascii=False)}\n\n"
                        break
                    yield f": keepalive\n\n"
        except GeneratorExit:
            pass
        finally:
            _sse_unregister(task_id, q)

    return Response(
        generate(),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no',
            'Connection': 'keep-alive',
        },
    )


@app.route('/api/tasks/<task_id>/cancel', methods=['POST'])
def api_cancel_task(task_id):
    task = _get_task(task_id)
    if task is None:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    if task.status not in ('pending', 'processing'):
        return jsonify({'success': False, 'message': f'任务状态为 {task.status}，无法取消'}), 400

    cancel_event = _cancel_events.get(task_id)
    if cancel_event:
        cancel_event.set()
        task.status = 'cancelled'
        task.finished_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        task.progress_message = '任务已取消'
        task.progress_stage = 'cancelled'
        _save_task(task)
        _sse_push(task_id, {
            'type': 'cancelled',
            'status': 'cancelled',
            'task_id': task_id,
        })
        logger.info('任务 %s 已取消', task_id)
        return jsonify({'success': True, 'message': '任务已取消'})
    else:
        return jsonify({'success': False, 'message': '无法取消该任务'}), 400


@app.route('/api/tasks/<task_id>', methods=['GET'])
def api_get_task(task_id):
    task = _get_task(task_id)
    if task is None:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    return jsonify({'success': True, 'data': task.to_dict()})


@app.route('/api/tasks', methods=['GET'])
def api_list_tasks():
    tasks = _load_tasks()
    status_filter = request.args.get('status', '').strip()
    limit = min(int(request.args.get('limit', 50)), 200)
    offset = int(request.args.get('offset', 0))

    task_list = list(tasks.values())
    if status_filter:
        task_list = [t for t in task_list if t.get('status') == status_filter]

    task_list.sort(key=lambda t: t.get('created_at', ''), reverse=True)
    total = len(task_list)
    task_list = task_list[offset:offset + limit]

    return jsonify({
        'success': True,
        'data': task_list,
        'total': total,
        'limit': limit,
        'offset': offset,
    })


@app.route('/api/download/<task_id>', methods=['GET'])
def api_download(task_id):
    task = _get_task(task_id)
    if task is None:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    if task.status != 'completed':
        return jsonify({'success': False, 'message': f'任务状态为 {task.status}，无法下载'}), 400

    if not task.output_path or not os.path.exists(task.output_path):
        return jsonify({'success': False, 'message': '总表文件不存在'}), 404

    download_name = os.path.basename(task.output_path)
    return send_file(task.output_path,
                     as_attachment=True,
                     download_name=download_name,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/download/<task_id>/masked', methods=['GET'])
def api_download_masked(task_id):
    task = _get_task(task_id)
    if task is None:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    if task.status != 'completed':
        return jsonify({'success': False, 'message': f'任务状态为 {task.status}，无法下载'}), 400

    if not task.masked_output_path or not os.path.exists(task.masked_output_path):
        return jsonify({'success': False, 'message': '脱敏版总表文件不存在'}), 404

    download_name = os.path.basename(task.masked_output_path)
    return send_file(task.masked_output_path,
                     as_attachment=True,
                     download_name=download_name,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/tasks/<task_id>', methods=['DELETE'])
def api_delete_task(task_id):
    task = _get_task(task_id)
    if task is None:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    if task.status == 'processing':
        return jsonify({'success': False, 'message': '任务正在处理中，请先取消'}), 400

    if task.output_path:
        task_output_dir = os.path.dirname(task.output_path)
        if os.path.exists(task_output_dir) and task_output_dir.startswith(OUTPUT_DIR):
            shutil.rmtree(task_output_dir, ignore_errors=True)

    tasks = _load_tasks()
    tasks.pop(task_id, None)
    _save_tasks(tasks)

    return jsonify({'success': True, 'message': '任务已删除'})


@app.route('/api/stats', methods=['GET'])
def api_stats():
    tasks = _load_tasks()
    task_list = list(tasks.values())

    stats = {
        'total_tasks': len(task_list),
        'completed_tasks': sum(1 for t in task_list if t.get('status') == 'completed'),
        'processing_tasks': sum(1 for t in task_list if t.get('status') == 'processing'),
        'failed_tasks': sum(1 for t in task_list if t.get('status') == 'failed'),
        'pending_tasks': sum(1 for t in task_list if t.get('status') == 'pending'),
        'cancelled_tasks': sum(1 for t in task_list if t.get('status') == 'cancelled'),
        'total_records': sum(t.get('total_records', 0) for t in task_list),
        'total_files': sum(t.get('total_files', 0) for t in task_list),
        'max_concurrent': MAX_CONCURRENT_TASKS,
        'active_count': sum(1 for t in task_list if t.get('status') in ('pending', 'processing')),
    }
    return jsonify({'success': True, 'data': stats})


@app.route('/api/counterparty-rules', methods=['GET'])
def api_list_counterparty_rules():
    config = bankcheck.get_counterparty_rule_config(BACKEND_DIR)
    rule_type = request.args.get('rule_type')
    enabled = request.args.get('enabled')
    enabled_filter = None
    if enabled is not None:
        enabled_filter = enabled.lower() == 'true'
    rules = config.get_rules(rule_type=rule_type, enabled=enabled_filter)
    data = [vars(r) for r in rules]
    return jsonify({'success': True, 'data': data, 'total': len(data)})


@app.route('/api/counterparty-rules', methods=['POST'])
def api_create_counterparty_rule():
    body = request.get_json(silent=True) or {}
    name = body.get('name', '').strip()
    rule_type = body.get('rule_type', '').strip()
    keywords = body.get('keywords', [])
    if not name:
        return jsonify({'success': False, 'message': '规则名称不能为空'}), 400
    if rule_type not in ('blacklist', 'whitelist'):
        return jsonify({'success': False, 'message': 'rule_type 必须为 blacklist 或 whitelist'}), 400
    if not keywords:
        return jsonify({'success': False, 'message': '关键词不能为空'}), 400
    rule_id = bankcheck.add_counterparty_keyword_rule(
        name=name,
        rule_type=rule_type,
        keywords=keywords,
        match_mode=body.get('match_mode', 'contains'),
        category=body.get('category', ''),
        severity=body.get('severity', 'medium'),
        description=body.get('description'),
        script_dir=BACKEND_DIR,
        username=body.get('created_by', ''),
    )
    return jsonify({'success': True, 'rule_id': rule_id, 'message': '规则创建成功'})


@app.route('/api/counterparty-rules/<rule_id>', methods=['PUT'])
def api_update_counterparty_rule(rule_id):
    body = request.get_json(silent=True) or {}
    config = bankcheck.get_counterparty_rule_config(BACKEND_DIR)
    config.load_config()
    ok = config.update_rule(rule_id, body)
    if not ok:
        return jsonify({'success': False, 'message': '规则不存在'}), 404
    return jsonify({'success': True, 'message': '规则更新成功'})


@app.route('/api/counterparty-rules/<rule_id>', methods=['DELETE'])
def api_delete_counterparty_rule(rule_id):
    config = bankcheck.get_counterparty_rule_config(BACKEND_DIR)
    config.load_config()
    ok = config.delete_rule(rule_id)
    if not ok:
        return jsonify({'success': False, 'message': '规则不存在'}), 404
    return jsonify({'success': True, 'message': '规则删除成功'})


@app.route('/api/counterparty-rules/<rule_id>/toggle', methods=['POST'])
def api_toggle_counterparty_rule(rule_id):
    body = request.get_json(silent=True) or {}
    enabled = body.get('enabled')
    if enabled is None:
        return jsonify({'success': False, 'message': '缺少 enabled 参数'}), 400
    config = bankcheck.get_counterparty_rule_config(BACKEND_DIR)
    config.load_config()
    ok = config.toggle_rule(rule_id, bool(enabled))
    if not ok:
        return jsonify({'success': False, 'message': '规则不存在'}), 404
    return jsonify({'success': True, 'message': '规则状态已更新'})


@app.route('/api/counterparty-rules/apply', methods=['POST'])
def api_apply_counterparty_rules():
    summary_path = os.path.join(BACKEND_DIR, '银行流水总表.xlsx')
    if not os.path.exists(summary_path):
        return jsonify({'success': False, 'message': '银行流水总表不存在'}), 404
    try:
        import pandas as pd
        df = pd.read_excel(summary_path, engine='openpyxl')
        records = df.to_dict(orient='records')
        tagged_records, summary = bankcheck.apply_counterparty_rules(records, BACKEND_DIR)
        df_out = pd.DataFrame(tagged_records)
        df_out.to_excel(summary_path, index=False, engine='openpyxl')
        return jsonify({'success': True, 'data': summary})
    except Exception as e:
        logger.error('应用对方户名规则失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/counterparty-rules/export', methods=['GET'])
def api_export_counterparty_tags():
    summary_path = os.path.join(BACKEND_DIR, '银行流水总表.xlsx')
    if not os.path.exists(summary_path):
        return jsonify({'success': False, 'message': '银行流水总表不存在'}), 404
    try:
        import pandas as pd
        df = pd.read_excel(summary_path, engine='openpyxl')
        records = df.to_dict(orient='records')
        tagged_records, _summary = bankcheck.apply_counterparty_rules(records, BACKEND_DIR)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        export_path = os.path.join(OUTPUT_DIR, f'对方户名标签_{timestamp}.xlsx')
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        result_path = bankcheck.export_counterparty_tags(tagged_records, export_path)
        if not result_path:
            return jsonify({'success': False, 'message': '没有命中规则的记录'}), 404
        return send_file(result_path,
                         as_attachment=True,
                         download_name=os.path.basename(result_path),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        logger.error('导出对方户名标签失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


# ──────────────────────────────────────────────
# 凭证附件关联 API
# ──────────────────────────────────────────────

@app.route('/api/voucher-attachments', methods=['GET'])
def api_list_voucher_attachments():
    """查询凭证附件列表"""
    try:
        transaction_id = request.args.get('transaction_id', '').strip() or None
        attachment_type = request.args.get('attachment_type', '').strip() or None
        keyword = request.args.get('keyword', '').strip() or None
        limit = request.args.get('limit', type=int)
        offset = request.args.get('offset', type=int, default=0)

        result = bankcheck.list_voucher_attachments(
            transaction_id=transaction_id,
            attachment_type=attachment_type,
            keyword=keyword,
            limit=limit,
            offset=offset,
            script_dir=BACKEND_DIR,
        )
        data = [r.to_dict() for r in result.records]
        return jsonify({
            'success': True,
            'data': data,
            'total': result.total_count,
        })
    except Exception as e:
        logger.error('查询凭证附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/voucher-attachments/<int:attachment_id>', methods=['GET'])
def api_get_voucher_attachment(attachment_id):
    """获取单个凭证附件"""
    try:
        att = bankcheck.get_voucher_attachment(attachment_id, script_dir=BACKEND_DIR)
        if att is None:
            return jsonify({'success': False, 'message': '附件记录不存在'}), 404
        return jsonify({'success': True, 'data': att.to_dict()})
    except Exception as e:
        logger.error('获取凭证附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/voucher-attachments', methods=['POST'])
def api_create_voucher_attachment():
    """新增凭证附件"""
    try:
        body = request.get_json(silent=True) or {}
        transaction_id = str(body.get('transaction_id', '') or '').strip()
        attachment_path = str(body.get('attachment_path', '') or '').strip()
        attachment_type = str(body.get('attachment_type', '其他') or '其他').strip()
        remark = body.get('remark')

        if not transaction_id:
            return jsonify({'success': False, 'message': '交易流水号不能为空'}), 400
        if not attachment_path:
            return jsonify({'success': False, 'message': '附件路径不能为空'}), 400

        new_id = bankcheck.add_voucher_attachment(
            transaction_id=transaction_id,
            attachment_path=attachment_path,
            attachment_type=attachment_type,
            remark=remark,
            script_dir=BACKEND_DIR,
        )
        return jsonify({
            'success': True,
            'id': new_id,
            'message': '凭证附件创建成功',
        }), 201
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('创建凭证附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/voucher-attachments/<int:attachment_id>', methods=['PUT'])
def api_update_voucher_attachment(attachment_id):
    """更新凭证附件"""
    try:
        body = request.get_json(silent=True) or {}
        attachment_path = body.get('attachment_path')
        attachment_type = body.get('attachment_type')
        remark = body.get('remark')

        ok = bankcheck.update_voucher_attachment(
            attachment_id=attachment_id,
            attachment_path=attachment_path,
            attachment_type=attachment_type,
            remark=remark,
            script_dir=BACKEND_DIR,
        )
        if not ok:
            return jsonify({'success': False, 'message': '附件记录不存在或无更新字段'}), 404
        return jsonify({'success': True, 'message': '更新成功'})
    except Exception as e:
        logger.error('更新凭证附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/voucher-attachments/<int:attachment_id>', methods=['DELETE'])
def api_delete_voucher_attachment(attachment_id):
    """删除凭证附件"""
    try:
        ok = bankcheck.delete_voucher_attachment(attachment_id, script_dir=BACKEND_DIR)
        if not ok:
            return jsonify({'success': False, 'message': '附件记录不存在'}), 404
        return jsonify({'success': True, 'message': '删除成功'})
    except Exception as e:
        logger.error('删除凭证附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/voucher-attachments/<int:attachment_id>/open', methods=['POST'])
def api_open_voucher_attachment(attachment_id):
    """一键打开凭证附件"""
    try:
        ok, msg = bankcheck.open_voucher_attachment(
            attachment_id=attachment_id,
            script_dir=BACKEND_DIR,
        )
        if ok:
            return jsonify({'success': True, 'message': msg})
        return jsonify({'success': False, 'message': msg}), 400
    except Exception as e:
        logger.error('打开凭证附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/voucher-attachments/open-by-transaction', methods=['POST'])
def api_open_voucher_by_transaction():
    """按交易流水号一键打开所有关联附件"""
    try:
        body = request.get_json(silent=True) or {}
        transaction_id = str(body.get('transaction_id', '') or '').strip()
        if not transaction_id:
            return jsonify({'success': False, 'message': '交易流水号不能为空'}), 400

        success_count, errors = bankcheck.open_voucher_attachments_for_transaction(
            transaction_id=transaction_id,
            script_dir=BACKEND_DIR,
        )
        return jsonify({
            'success': True,
            'success_count': success_count,
            'errors': errors,
            'message': f'成功打开 {success_count} 个附件' + (f'，失败 {len(errors)} 个' if errors else ''),
        })
    except Exception as e:
        logger.error('按交易流水号打开附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/voucher-attachments/by-transaction/<transaction_id>', methods=['GET'])
def api_get_voucher_by_transaction(transaction_id):
    """获取某笔交易流水关联的所有附件"""
    try:
        attachments = bankcheck.get_voucher_attachments_for_transaction(
            transaction_id=transaction_id,
            script_dir=BACKEND_DIR,
        )
        data = [a.to_dict() for a in attachments]
        return jsonify({
            'success': True,
            'data': data,
            'total': len(data),
        })
    except Exception as e:
        logger.error('获取交易流水附件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'success': False, 'message': '上传文件总大小超过 500MB 限制'}), 413


@app.errorhandler(404)
def not_found(error):
    return jsonify({'success': False, 'message': '接口不存在'}), 404


@app.errorhandler(500)
def internal_error(error):
    logger.error('服务器内部错误: %s', error, exc_info=True)
    return jsonify({'success': False, 'message': '服务器内部错误'}), 500


# ──────────────────────────────────────────────
# 流水可视化看板 API
# ──────────────────────────────────────────────

def _get_dashboard_data_source():
    """获取看板数据源，优先从数据库读取，其次从总表文件读取"""
    summary_path = os.path.join(BACKEND_DIR, '银行流水总表.xlsx')

    try:
        db = db_module.SQLiteBackend()
        db.connect()
        stats = db.get_statistics()
        if stats.get('总记录数', 0) > 0:
            return {'type': 'database', 'db': db, 'summary_path': summary_path}
    except Exception as e:
        logger.warning('数据库连接失败: %s', e)

    if os.path.exists(summary_path):
        return {'type': 'excel', 'db': None, 'summary_path': summary_path}

    return {'type': 'none', 'db': None, 'summary_path': summary_path}


def _load_transaction_data():
    """加载交易数据，优先从数据库，其次从 Excel"""
    source = _get_dashboard_data_source()

    if source['type'] == 'database':
        result = source['db'].query_records(limit=100000)
        records = [r.to_dict() for r in result.records]
        return records

    if source['type'] == 'excel':
        try:
            import pandas as pd
            df = pd.read_excel(source['summary_path'], engine='openpyxl')
            df = df.where(pd.notnull(df), None)
            return df.to_dict(orient='records')
        except Exception as e:
            logger.error('读取总表 Excel 失败: %s', e, exc_info=True)
            return []

    return []


def _parse_date(date_val):
    """解析日期，返回 YYYY-MM 格式字符串"""
    if date_val is None:
        return None
    try:
        date_str = str(date_val).strip()
        if not date_str:
            return None

        match = re.search(r'(\d{4})[-/](\d{2})', date_str)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            if 2000 <= year <= 2100 and 1 <= month <= 12:
                return f"{year}-{month:02d}"

        if len(date_str) == 8 and date_str.isdigit():
            year = int(date_str[:4])
            month = int(date_str[4:6])
            if 2000 <= year <= 2100 and 1 <= month <= 12:
                return f"{year}-{month:02d}"

        if hasattr(date_val, 'strftime'):
            return date_val.strftime('%Y-%m')
    except Exception:
        pass
    return None


def _to_float(val):
    """安全转换为浮点数"""
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


@app.route('/api/dashboard/summary', methods=['GET'])
def api_dashboard_summary():
    """获取看板总体统计数据"""
    try:
        records = _load_transaction_data()

        if not records:
            return jsonify({
                'success': True,
                'data': {
                    'total_records': 0,
                    'total_payment': 0,
                    'total_receipt': 0,
                    'net_amount': 0,
                    'subject_count': 0,
                    'counterparty_count': 0,
                    'date_range': {'start': None, 'end': None}
                }
            })

        total_payment = 0.0
        total_receipt = 0.0
        subjects = set()
        counterparties = set()
        dates = []

        for r in records:
            payment = _to_float(r.get('付款'))
            receipt = _to_float(r.get('收款'))

            if payment < 0:
                total_payment += payment
            if receipt > 0:
                total_receipt += receipt

            if r.get('主体'):
                subjects.add(r.get('主体'))
            if r.get('对方户名'):
                counterparties.add(r.get('对方户名'))

            month = _parse_date(r.get('交易日期'))
            if month:
                dates.append(month)

        dates_sorted = sorted(dates) if dates else []

        return jsonify({
            'success': True,
            'data': {
                'total_records': len(records),
                'total_payment': round(total_payment, 2),
                'total_receipt': round(total_receipt, 2),
                'net_amount': round(total_payment + total_receipt, 2),
                'subject_count': len(subjects),
                'counterparty_count': len(counterparties),
                'date_range': {
                    'start': dates_sorted[0] if dates_sorted else None,
                    'end': dates_sorted[-1] if dates_sorted else None
                }
            }
        })

    except Exception as e:
        logger.error('获取看板统计数据失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/dashboard/monthly-trend', methods=['GET'])
def api_dashboard_monthly_trend():
    """获取按月经费趋势数据"""
    try:
        records = _load_transaction_data()

        if not records:
            return jsonify({'success': True, 'data': []})

        monthly_data = {}

        for r in records:
            month = _parse_date(r.get('交易日期'))
            if not month:
                continue

            if month not in monthly_data:
                monthly_data[month] = {'month': month, 'payment': 0.0, 'receipt': 0.0, 'count': 0}

            payment = _to_float(r.get('付款'))
            receipt = _to_float(r.get('收款'))

            if payment < 0:
                monthly_data[month]['payment'] += payment
            if receipt > 0:
                monthly_data[month]['receipt'] += receipt
            monthly_data[month]['count'] += 1

        result = sorted(monthly_data.values(), key=lambda x: x['month'])
        for item in result:
            item['payment'] = round(item['payment'], 2)
            item['receipt'] = round(item['receipt'], 2)
            item['net'] = round(item['payment'] + item['receipt'], 2)

        return jsonify({'success': True, 'data': result})

    except Exception as e:
        logger.error('获取月度趋势数据失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/dashboard/top-counterparties', methods=['GET'])
def api_dashboard_top_counterparties():
    """获取 Top 对方户名数据"""
    try:
        top_n = int(request.args.get('limit', 10))
        records = _load_transaction_data()

        if not records:
            return jsonify({'success': True, 'data': []})

        counterparty_data = {}

        for r in records:
            cp = r.get('对方户名')
            if not cp or str(cp).strip() == '':
                continue

            cp = str(cp).strip()
            if cp not in counterparty_data:
                counterparty_data[cp] = {
                    'name': cp,
                    'payment': 0.0,
                    'receipt': 0.0,
                    'count': 0
                }

            payment = _to_float(r.get('付款'))
            receipt = _to_float(r.get('收款'))

            if payment < 0:
                counterparty_data[cp]['payment'] += payment
            if receipt > 0:
                counterparty_data[cp]['receipt'] += receipt
            counterparty_data[cp]['count'] += 1

        result = list(counterparty_data.values())
        for item in result:
            item['total'] = round(abs(item['payment']) + item['receipt'], 2)
            item['payment'] = round(item['payment'], 2)
            item['receipt'] = round(item['receipt'], 2)

        result.sort(key=lambda x: x['total'], reverse=True)
        result = result[:top_n]

        return jsonify({'success': True, 'data': result})

    except Exception as e:
        logger.error('获取 Top 对方户名数据失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/dashboard/subject-breakdown', methods=['GET'])
def api_dashboard_subject_breakdown():
    """获取各主体收支占比数据"""
    try:
        records = _load_transaction_data()

        if not records:
            return jsonify({'success': True, 'data': []})

        subject_data = {}

        for r in records:
            subject = r.get('主体') or '未知主体'
            subject = str(subject).strip() or '未知主体'

            if subject not in subject_data:
                subject_data[subject] = {
                    'name': subject,
                    'payment': 0.0,
                    'receipt': 0.0,
                    'count': 0
                }

            payment = _to_float(r.get('付款'))
            receipt = _to_float(r.get('收款'))

            if payment < 0:
                subject_data[subject]['payment'] += payment
            if receipt > 0:
                subject_data[subject]['receipt'] += receipt
            subject_data[subject]['count'] += 1

        result = list(subject_data.values())
        for item in result:
            item['payment'] = round(item['payment'], 2)
            item['receipt'] = round(item['receipt'], 2)
            item['net'] = round(item['payment'] + item['receipt'], 2)
            item['total'] = round(abs(item['payment']) + item['receipt'], 2)

        result.sort(key=lambda x: x['total'], reverse=True)

        return jsonify({'success': True, 'data': result})

    except Exception as e:
        logger.error('获取主体收支占比数据失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


# ──────────────────────────────────────────────
# 银企直连目录对接 API
# ──────────────────────────────────────────────

try:
    from bank_directory_connector import BankDirectoryConnector
    HAS_DIRECTORY_CONNECTOR_WS = True
except ImportError as e:
    HAS_DIRECTORY_CONNECTOR_WS = False
    logger.warning('目录对接模块不可用: %s', e)


def _get_directory_connector():
    """获取目录连接器实例"""
    if not HAS_DIRECTORY_CONNECTOR_WS:
        return None
    try:
        return BankDirectoryConnector(script_dir=BACKEND_DIR)
    except Exception as e:
        logger.error('创建目录连接器失败: %s', e)
        return None


@app.route('/api/directory/status', methods=['GET'])
def api_directory_status():
    """获取目录对接状态"""
    try:
        connector = _get_directory_connector()
        if connector is None:
            return jsonify({
                'success': False,
                'message': '目录对接模块不可用'
            }), 503

        status = connector.get_status()
        return jsonify({
            'success': True,
            'data': status
        })
    except Exception as e:
        logger.error('获取目录状态失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/directory/run', methods=['POST'])
def api_directory_run():
    """触发一次目录处理"""
    try:
        connector = _get_directory_connector()
        if connector is None:
            return jsonify({
                'success': False,
                'message': '目录对接模块不可用'
            }), 503

        data = request.get_json(silent=True) or {}
        incremental = data.get('incremental', None)
        keep_strategy = data.get('keep_strategy', None)

        if incremental is not None:
            connector._processing_config.incremental = incremental
        if keep_strategy is not None:
            connector._processing_config.keep_strategy = keep_strategy

        task_id = f'dir_{datetime.now().strftime("%Y%m%d%H%M%S")}'

        def _run_directory_pipeline():
            try:
                result = connector.run_once()
                _sse_push(task_id, {
                    'type': 'completed',
                    'status': 'completed' if result.success else 'failed',
                    'task_id': task_id,
                    'data': {
                        'success': result.success,
                        'message': result.message,
                        'processed_files': result.processed_files,
                        'error_files': result.error_files,
                        'output_path': result.output_path,
                        'archive_dir': result.archive_dir,
                    }
                })
            except Exception as e:
                logger.exception('目录处理任务失败')
                _sse_push(task_id, {
                    'type': 'error',
                    'message': str(e),
                    'status': 'failed',
                    'task_id': task_id,
                })

        thread = threading.Thread(target=_run_directory_pipeline, daemon=True)
        thread.start()

        return jsonify({
            'success': True,
            'message': '目录处理任务已启动',
            'task_id': task_id,
        })

    except Exception as e:
        logger.error('触发目录处理失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/directory/download', methods=['POST'])
def api_directory_download():
    """触发银行流水下载"""
    try:
        connector = _get_directory_connector()
        if connector is None:
            return jsonify({
                'success': False,
                'message': '目录对接模块不可用'
            }), 503

        data = request.get_json(silent=True) or {}
        bank_name = data.get('bank_name')

        if not bank_name:
            return jsonify({
                'success': False,
                'message': '请指定银行名称'
            }), 400

        success, message = connector.trigger_download(bank_name)

        return jsonify({
            'success': success,
            'message': message,
            'bank_name': bank_name,
        })

    except Exception as e:
        logger.error('触发银行下载失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/directory/config', methods=['GET'])
def api_directory_config():
    """获取目录对接配置"""
    try:
        connector = _get_directory_connector()
        if connector is None:
            return jsonify({
                'success': False,
                'message': '目录对接模块不可用'
            }), 503

        config = {
            'root_dir': connector._directory_config.root_dir,
            'poll_interval': connector._directory_config.poll_interval,
            'file_stable_seconds': connector._directory_config.file_stable_seconds,
            'enable_lock_detection': connector._directory_config.enable_lock_detection,
            'archive_retention_days': connector._directory_config.archive_retention_days,
            'processing': {
                'incremental': connector._processing_config.incremental,
                'keep_strategy': connector._processing_config.keep_strategy,
                'generate_report': connector._processing_config.generate_report,
            },
            'banks': [
                {
                    'name': name,
                    'enabled': cfg.enabled,
                    'file_pattern': cfg.file_pattern,
                    'has_download_script': cfg.download_script is not None,
                    'download_schedule': cfg.download_script.get('schedule') if cfg.download_script else None,
                }
                for name, cfg in connector._bank_configs.items()
            ],
        }

        return jsonify({
            'success': True,
            'data': config
        })

    except Exception as e:
        logger.error('获取目录配置失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/directory/archives', methods=['GET'])
def api_directory_archives():
    """获取归档列表"""
    try:
        connector = _get_directory_connector()
        if connector is None:
            return jsonify({
                'success': False,
                'message': '目录对接模块不可用'
            }), 503

        outbox_dir = os.path.join(connector._directory_config.root_dir, 'outbox')
        archives = []

        if os.path.isdir(outbox_dir):
            for entry in sorted(os.listdir(outbox_dir), reverse=True):
                entry_path = os.path.join(outbox_dir, entry)
                if os.path.isdir(entry_path):
                    try:
                        stat = os.stat(entry_path)
                        manifest_path = os.path.join(entry_path, 'manifest.json')
                        manifest = {}
                        if os.path.exists(manifest_path):
                            import json
                            with open(manifest_path, 'r', encoding='utf-8') as f:
                                manifest = json.load(f)

                        archives.append({
                            'timestamp': entry,
                            'path': entry_path,
                            'created_at': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                            'total_records': manifest.get('total_records', 0),
                            'processed_files': len(manifest.get('processed_files', [])),
                            'error_files': len(manifest.get('error_files', [])),
                            'has_output': os.path.exists(os.path.join(entry_path, '银行流水总表.xlsx')),
                        })
                    except (OSError, ValueError):
                        continue

        return jsonify({
            'success': True,
            'data': archives
        })

    except Exception as e:
        logger.error('获取归档列表失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/directory/archive/<timestamp>/download', methods=['GET'])
def api_directory_archive_download(timestamp):
    """下载指定归档的结果文件"""
    try:
        connector = _get_directory_connector()
        if connector is None:
            return jsonify({
                'success': False,
                'message': '目录对接模块不可用'
            }), 503

        import re
        if not re.match(r'^\d{8}_\d{6}$', timestamp):
            return jsonify({
                'success': False,
                'message': '无效的时间戳格式'
            }), 400

        outbox_dir = os.path.join(connector._directory_config.root_dir, 'outbox')
        archive_dir = os.path.join(outbox_dir, timestamp)

        if not os.path.isdir(archive_dir):
            return jsonify({
                'success': False,
                'message': '归档不存在'
            }), 404

        file_type = request.args.get('type', 'summary')
        if file_type == 'summary':
            output_file = os.path.join(archive_dir, '银行流水总表.xlsx')
        elif file_type == 'report':
            output_file = os.path.join(archive_dir, '流水检验报告.md')
        else:
            return jsonify({
                'success': False,
                'message': '不支持的文件类型'
            }), 400

        if not os.path.exists(output_file):
            return jsonify({
                'success': False,
                'message': '文件不存在'
            }), 404

        return send_file(
            output_file,
            as_attachment=True,
            download_name=f'{timestamp}_{os.path.basename(output_file)}'
        )

    except Exception as e:
        logger.error('下载归档文件失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/directory')
def directory_page():
    """目录对接管理页面"""
    return render_template('directory.html',
                           has_directory_connector=HAS_DIRECTORY_CONNECTOR_WS)


@app.route('/workflow')
def workflow_page():
    """工作流管理页面"""
    return render_template('workflow.html')


# ──────────────────────────────────────────────
# 工作流 API
# ──────────────────────────────────────────────

def _get_wf_manager():
    """获取工作流管理器实例"""
    return workflow_module.get_workflow_manager(BACKEND_DIR)


@app.route('/api/workflows', methods=['GET'])
def api_list_workflows():
    """查询工作流列表"""
    try:
        wf = _get_wf_manager()
        status = request.args.get('status', '').strip() or None
        submitter = request.args.get('submitter', '').strip() or None
        approver = request.args.get('approver', '').strip() or None
        batch_id = request.args.get('batch_id', '').strip() or None
        limit = min(int(request.args.get('limit', 50)), 200)
        offset = int(request.args.get('offset', 0))

        workflows, total = wf.list_workflows(
            status=status,
            submitter=submitter,
            approver=approver,
            batch_id=batch_id,
            limit=limit,
            offset=offset
        )

        data = [w.to_dict() for w in workflows]
        return jsonify({
            'success': True,
            'data': data,
            'total': total,
            'limit': limit,
            'offset': offset,
        })

    except Exception as e:
        logger.error('查询工作流列表失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>', methods=['GET'])
def api_get_workflow(workflow_id):
    """获取工作流详情"""
    try:
        wf = _get_wf_manager()
        try:
            workflow = wf.get_workflow(workflow_id)
        except ValueError as e:
            if '不存在' in str(e):
                return jsonify({'success': False, 'message': str(e)}), 404
            raise

        data = workflow.to_dict()
        exceptions = wf.get_exception_items(workflow_id)
        actions = wf.get_action_logs(workflow_id)

        data['exceptions'] = [e.to_dict() for e in exceptions]
        data['actions'] = [a.to_dict() for a in actions]

        return jsonify({'success': True, 'data': data})

    except Exception as e:
        logger.error('获取工作流详情失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows', methods=['POST'])
def api_create_workflow():
    """创建工作流"""
    try:
        body = request.get_json(silent=True) or {}
        batch_id = body.get('batch_id', '').strip()
        title = body.get('title', '').strip()
        description = body.get('description', '').strip()
        submitter = body.get('submitter', '').strip() or None
        input_folder = body.get('input_folder', '').strip()
        output_path = body.get('output_path', '').strip() or None
        total_records = int(body.get('total_records', 0) or 0)
        exception_items = body.get('exception_items', [])

        if not batch_id:
            return jsonify({'success': False, 'message': '批次号不能为空'}), 400
        if not title:
            return jsonify({'success': False, 'message': '标题不能为空'}), 400

        wf = _get_wf_manager()
        workflow = wf.create_workflow(
            batch_id=batch_id,
            title=title,
            description=description,
            submitter=submitter,
            input_folder=input_folder,
            output_path=output_path,
            total_records=total_records,
            exception_items=exception_items
        )

        return jsonify({
            'success': True,
            'workflow_id': workflow.workflow_id,
            'data': workflow.to_dict(),
            'message': '工作流创建成功',
        }), 201

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('创建工作流失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/submit', methods=['POST'])
def api_submit_workflow(workflow_id):
    """提交审批"""
    try:
        body = request.get_json(silent=True) or {}
        operator = body.get('operator', '').strip() or None
        remark = body.get('remark', '').strip() or None

        wf = _get_wf_manager()
        workflow = wf.submit_for_approval(
            workflow_id=workflow_id,
            operator=operator,
            remark=remark
        )

        return jsonify({
            'success': True,
            'data': workflow.to_dict(),
            'message': '已提交审批',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('提交审批失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/approve', methods=['POST'])
def api_approve_workflow(workflow_id):
    """审批通过"""
    try:
        body = request.get_json(silent=True) or {}
        approver = body.get('approver', '').strip() or None
        remark = body.get('remark', '').strip() or None

        wf = _get_wf_manager()
        workflow = wf.approve_workflow(
            workflow_id=workflow_id,
            approver=approver,
            remark=remark
        )

        return jsonify({
            'success': True,
            'data': workflow.to_dict(),
            'message': '审批通过，异常清单已确认',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('审批失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/reject', methods=['POST'])
def api_reject_workflow(workflow_id):
    """驳回工作流"""
    try:
        body = request.get_json(silent=True) or {}
        approver = body.get('approver', '').strip() or None
        reject_reason = body.get('reject_reason', '').strip()
        remark = body.get('remark', '').strip() or None

        if not reject_reason:
            return jsonify({'success': False, 'message': '驳回原因不能为空'}), 400

        wf = _get_wf_manager()
        workflow = wf.reject_workflow(
            workflow_id=workflow_id,
            approver=approver,
            reject_reason=reject_reason,
            remark=remark
        )

        return jsonify({
            'success': True,
            'data': workflow.to_dict(),
            'message': '已驳回',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('驳回失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/publish', methods=['POST'])
def api_publish_workflow(workflow_id):
    """正式发布总表"""
    try:
        body = request.get_json(silent=True) or {}
        publisher = body.get('publisher', '').strip() or None
        output_path = body.get('output_path', '').strip() or None
        remark = body.get('remark', '').strip() or None

        wf = _get_wf_manager()
        workflow = wf.publish_workflow(
            workflow_id=workflow_id,
            publisher=publisher,
            output_path=output_path,
            remark=remark
        )

        return jsonify({
            'success': True,
            'data': workflow.to_dict(),
            'message': '总表已正式发布',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('发布失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/cancel', methods=['POST'])
def api_cancel_workflow(workflow_id):
    """取消工作流"""
    try:
        body = request.get_json(silent=True) or {}
        operator = body.get('operator', '').strip() or None
        remark = body.get('remark', '').strip() or None

        wf = _get_wf_manager()
        workflow = wf.cancel_workflow(
            workflow_id=workflow_id,
            operator=operator,
            remark=remark
        )

        return jsonify({
            'success': True,
            'data': workflow.to_dict(),
            'message': '工作流已取消',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('取消工作流失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/exceptions', methods=['GET'])
def api_list_exceptions(workflow_id):
    """获取异常项列表"""
    try:
        status = request.args.get('status', '').strip() or None

        wf = _get_wf_manager()
        exceptions = wf.get_exception_items(workflow_id, status=status)

        data = [e.to_dict() for e in exceptions]
        return jsonify({
            'success': True,
            'data': data,
            'total': len(data),
        })

    except Exception as e:
        logger.error('获取异常项列表失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/exceptions', methods=['POST'])
def api_add_exceptions(workflow_id):
    """添加异常项"""
    try:
        body = request.get_json(silent=True) or {}
        items = body.get('items', [])
        operator = body.get('operator', '').strip() or None

        if not items:
            return jsonify({'success': False, 'message': '异常项列表不能为空'}), 400

        wf = _get_wf_manager()
        new_ids = wf.add_exception_items(
            workflow_id=workflow_id,
            items=items,
            operator=operator
        )

        return jsonify({
            'success': True,
            'exception_ids': new_ids,
            'message': f'已添加 {len(new_ids)} 个异常项',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('添加异常项失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/exceptions/<int:exception_id>', methods=['PUT'])
def api_update_exception(exception_id):
    """更新异常项状态（复核人确认）"""
    try:
        body = request.get_json(silent=True) or {}
        status = body.get('status', '').strip() or None
        remark = body.get('remark')
        operator = body.get('operator', '').strip() or None

        if status is None and remark is None:
            return jsonify({'success': False, 'message': '缺少更新字段'}), 400

        wf = _get_wf_manager()
        success = wf.update_exception_item(
            exception_id=exception_id,
            status=status,
            remark=remark,
            operator=operator
        )

        if not success:
            return jsonify({'success': False, 'message': '更新失败，异常项不存在'}), 404

        return jsonify({
            'success': True,
            'message': '异常项已更新',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('更新异常项失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/exceptions/<int:exception_id>', methods=['DELETE'])
def api_delete_exception(exception_id):
    """删除异常项"""
    try:
        body = request.get_json(silent=True) or {}
        operator = (body or {}).get('operator', '').strip() or None

        wf = _get_wf_manager()
        success = wf.delete_exception_item(
            exception_id=exception_id,
            operator=operator
        )

        if not success:
            return jsonify({'success': False, 'message': '删除失败，异常项不存在'}), 404

        return jsonify({
            'success': True,
            'message': '异常项已删除',
        })

    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error('删除异常项失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/<workflow_id>/actions', methods=['GET'])
def api_get_action_logs(workflow_id):
    """获取操作日志"""
    try:
        wf = _get_wf_manager()
        actions = wf.get_action_logs(workflow_id)

        data = [a.to_dict() for a in actions]
        return jsonify({
            'success': True,
            'data': data,
            'total': len(data),
        })

    except Exception as e:
        logger.error('获取操作日志失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/workflows/stats', methods=['GET'])
def api_workflow_stats():
    """获取工作流统计信息"""
    try:
        wf = _get_wf_manager()
        stats = wf.get_statistics()

        return jsonify({
            'success': True,
            'data': stats,
        })

    except Exception as e:
        logger.error('获取工作流统计失败: %s', e, exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500


def main():
    host = os.environ.get('BANKCHECK_HOST', '0.0.0.0')
    port = int(os.environ.get('BANKCHECK_PORT', '5001'))
    debug = os.environ.get('BANKCHECK_DEBUG', 'false').lower() == 'true'

    bankcheck.setup_logging()

    logger.info('=' * 60)
    logger.info('银行流水检验 Web 服务启动')
    logger.info('访问地址: http://%s:%d', host, port)
    logger.info('上传目录: %s', UPLOAD_DIR)
    logger.info('输出目录: %s', OUTPUT_DIR)
    logger.info('最大并发任务数: %d', MAX_CONCURRENT_TASKS)
    logger.info('=' * 60)

    app.run(host=host, port=port, debug=debug, threaded=True)


if __name__ == '__main__':
    main()
