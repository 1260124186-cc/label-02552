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


BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BACKEND_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BACKEND_DIR, 'output')
TASKS_DB_PATH = os.path.join(BACKEND_DIR, 'web_tasks.json')

MAX_CONCURRENT_TASKS = 2

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

_task_semaphore = threading.Semaphore(MAX_CONCURRENT_TASKS)
_cancel_events: Dict[str, threading.Event] = {}
_sse_queues: Dict[str, List[queue.Queue]] = {}
_sse_lock = threading.Lock()


def setup_logging():
    log_file = os.path.join(BACKEND_DIR, 'web_service.log')
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


def _safe_filename(filename):
    safe = filename.replace('\\', '/').replace('../', '').replace('..\\', '')
    safe = os.path.basename(safe)
    safe = re.sub(r'[\x00-\x1f\x7f]', '', safe)
    if not safe:
        safe = 'unnamed'
    return safe


@app.route('/api/upload', methods=['POST'])
def api_upload():
    if 'files' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'}), 400

    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'success': False, 'message': '未选择文件'}), 400

    incremental = request.form.get('incremental', 'true').lower() == 'true'
    operator = request.form.get('operator', '').strip()

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

    thread = threading.Thread(target=_run_pipeline_async, args=(task, cancel_event), daemon=True)
    thread.start()

    logger.info('任务 %s 已创建，共上传 %d 个文件', task_id, saved_count)
    return jsonify({
        'success': True,
        'task_id': task_id,
        'message': f'已上传 {saved_count} 个文件，正在后台处理',
    })


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
