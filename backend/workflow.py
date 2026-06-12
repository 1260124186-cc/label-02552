# -*- coding: utf-8 -*-
"""
轻量工作流模块 - 处理前审批、处理后确认
支持状态流转：提交批次 → 复核人确认异常清单 → 正式发布总表
适配内控要求较高的财务团队

功能：
  1. 工作流实例管理（创建、查询、状态变更）
  2. 审批动作记录（提交、审批、驳回、发布）
  3. 异常清单管理（新增、确认、驳回）
  4. 操作留痕与审计追踪
  5. 角色权限控制（提交人、复核人）
"""

import os
import sys
import uuid
import json
import sqlite3
import logging
from dataclasses import dataclass, field, asdict
from datetime import datetime
from enum import Enum
from typing import List, Optional, Dict, Any, Tuple


WORKFLOW_DB_FILENAME = 'workflow.db'


class WorkflowStatus(str, Enum):
    """工作流状态枚举"""
    DRAFT = 'draft'
    PENDING_APPROVAL = 'pending_approval'
    APPROVED = 'approved'
    REJECTED = 'rejected'
    PUBLISHED = 'published'
    CANCELLED = 'cancelled'


class WorkflowAction(str, Enum):
    """工作流动作枚举"""
    CREATE = 'create'
    SUBMIT = 'submit'
    APPROVE = 'approve'
    REJECT = 'reject'
    PUBLISH = 'publish'
    CANCEL = 'cancel'
    UPDATE_EXCEPTION = 'update_exception'
    RESUBMIT = 'resubmit'


class ExceptionStatus(str, Enum):
    """异常项状态枚举"""
    PENDING = 'pending'
    CONFIRMED = 'confirmed'
    RESOLVED = 'resolved'
    IGNORED = 'ignored'


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_program_dir():
    return get_script_dir()


def is_writable(dir_path):
    import uuid as _uuid
    if not os.path.isdir(dir_path):
        return False
    try:
        test_file = os.path.join(dir_path, '.wf_write_test_' + _uuid.uuid4().hex[:8])
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        return True
    except (OSError, IOError):
        return False


def get_user_data_dir():
    app_name = 'bankcheck'
    if sys.platform.startswith('win'):
        base_dir = os.environ.get('APPDATA')
        if not base_dir:
            base_dir = os.path.expanduser('~\\AppData\\Roaming')
        return os.path.join(base_dir, app_name)
    elif sys.platform == 'darwin':
        return os.path.join(os.path.expanduser('~/Library/Application Support'), app_name)
    else:
        return os.path.join(os.path.expanduser('~'), '.' + app_name)


def get_writable_dir():
    program_dir = get_program_dir()
    if is_writable(program_dir):
        return program_dir
    user_data_dir = get_user_data_dir()
    os.makedirs(user_data_dir, exist_ok=True)
    return user_data_dir


def get_logger():
    return logging.getLogger('bankcheck.workflow')


def get_workflow_db_path(script_dir=None):
    """获取工作流数据库文件路径"""
    if script_dir is None:
        script_dir = get_writable_dir()
    return os.path.join(script_dir, WORKFLOW_DB_FILENAME)


def get_current_user():
    """获取当前操作用户，优先从环境变量获取"""
    import getpass
    user = os.environ.get('BANKCHECK_USER', '').strip()
    if user:
        return user
    try:
        return getpass.getuser()
    except Exception:
        return 'unknown_user'


@dataclass
class ExceptionItem:
    """异常项数据类"""
    id: Optional[int] = None
    workflow_id: str = ''
    exception_type: str = ''
    description: str = ''
    transaction_id: Optional[str] = None
    amount: Optional[float] = None
    counterparty: Optional[str] = None
    status: str = ExceptionStatus.PENDING.value
    remark: Optional[str] = None
    handled_by: Optional[str] = None
    handled_at: Optional[str] = None
    created_at: Optional[str] = None
    updated_at: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ExceptionItem':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class WorkflowInstance:
    """工作流实例数据类"""
    workflow_id: str
    batch_id: str
    title: str
    status: str = WorkflowStatus.DRAFT.value
    submitter: str = ''
    approver: Optional[str] = None
    publisher: Optional[str] = None
    description: str = ''
    total_records: int = 0
    exception_count: int = 0
    confirmed_exception_count: int = 0
    input_folder: str = ''
    output_path: Optional[str] = None
    submitted_at: Optional[str] = None
    approved_at: Optional[str] = None
    published_at: Optional[str] = None
    reject_reason: Optional[str] = None
    created_at: Optional[str] = None
    updated_at: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'WorkflowInstance':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class WorkflowActionLog:
    """工作流操作日志数据类"""
    id: Optional[int] = None
    workflow_id: str = ''
    action: str = ''
    operator: str = ''
    remark: Optional[str] = None
    old_status: Optional[str] = None
    new_status: Optional[str] = None
    created_at: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'WorkflowActionLog':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class WorkflowManager:
    """工作流管理器 - 核心业务逻辑类"""

    def __init__(self, script_dir: Optional[str] = None):
        self.script_dir = script_dir or get_writable_dir()
        self.db_path = get_workflow_db_path(self.script_dir)
        self.logger = get_logger()
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        """获取数据库连接"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def _init_db(self):
        """初始化数据库表结构"""
        conn = self._get_connection()
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS workflow_instances (
                workflow_id TEXT PRIMARY KEY,
                batch_id TEXT NOT NULL,
                title TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'draft',
                submitter TEXT NOT NULL,
                approver TEXT,
                publisher TEXT,
                description TEXT,
                total_records INTEGER DEFAULT 0,
                exception_count INTEGER DEFAULT 0,
                confirmed_exception_count INTEGER DEFAULT 0,
                input_folder TEXT,
                output_path TEXT,
                submitted_at TEXT,
                approved_at TEXT,
                published_at TEXT,
                reject_reason TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS exception_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workflow_id TEXT NOT NULL,
                exception_type TEXT NOT NULL,
                description TEXT NOT NULL,
                transaction_id TEXT,
                amount REAL,
                counterparty TEXT,
                status TEXT NOT NULL DEFAULT 'pending',
                remark TEXT,
                handled_by TEXT,
                handled_at TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (workflow_id) REFERENCES workflow_instances(workflow_id) ON DELETE CASCADE
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS workflow_action_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workflow_id TEXT NOT NULL,
                action TEXT NOT NULL,
                operator TEXT NOT NULL,
                remark TEXT,
                old_status TEXT,
                new_status TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (workflow_id) REFERENCES workflow_instances(workflow_id) ON DELETE CASCADE
            )
        ''')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_workflow_status ON workflow_instances(status)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_workflow_submitter ON workflow_instances(submitter)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_workflow_approver ON workflow_instances(approver)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_workflow_batch ON workflow_instances(batch_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_workflow_created ON workflow_instances(created_at)')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_exception_workflow ON exception_items(workflow_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_exception_status ON exception_items(status)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_exception_transaction ON exception_items(transaction_id)')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_action_workflow ON workflow_action_logs(workflow_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_action_operator ON workflow_action_logs(operator)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_action_created ON workflow_action_logs(created_at)')

        conn.commit()
        conn.close()
        self.logger.debug('工作流数据库初始化完成')

    def _log_action(self, cursor: sqlite3.Cursor, workflow_id: str, action: str,
                    operator: str, remark: Optional[str] = None,
                    old_status: Optional[str] = None, new_status: Optional[str] = None):
        """记录操作日志"""
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        cursor.execute('''
            INSERT INTO workflow_action_logs (
                workflow_id, action, operator, remark, old_status, new_status, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (workflow_id, action, operator, remark, old_status, new_status, now))

    def _update_exception_counts(self, cursor: sqlite3.Cursor, workflow_id: str):
        """更新工作流实例的异常计数"""
        cursor.execute('''
            SELECT COUNT(*) as total,
                   SUM(CASE WHEN status != 'pending' THEN 1 ELSE 0 END) as confirmed
            FROM exception_items
            WHERE workflow_id = ?
        ''', (workflow_id,))
        row = cursor.fetchone()
        total = row['total'] if row else 0
        confirmed = row['confirmed'] if row and row['confirmed'] else 0

        cursor.execute('''
            UPDATE workflow_instances
            SET exception_count = ?, confirmed_exception_count = ?, updated_at = ?
            WHERE workflow_id = ?
        ''', (total, confirmed, datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f'), workflow_id))

    def create_workflow(self, batch_id: str, title: str, description: str = '',
                       submitter: Optional[str] = None,
                       input_folder: str = '', output_path: Optional[str] = None,
                       total_records: int = 0,
                       exception_items: Optional[List[Dict[str, Any]]] = None) -> WorkflowInstance:
        """
        创建工作流实例

        Args:
            batch_id: 关联的批次号
            title: 工作流标题
            description: 描述说明
            submitter: 提交人（默认当前用户）
            input_folder: 输入文件夹路径
            output_path: 输出文件路径
            total_records: 总记录数
            exception_items: 异常项列表

        Returns:
            创建的工作流实例
        """
        submitter = submitter or get_current_user()
        workflow_id = f"WF{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

        conn = self._get_connection()
        cursor = conn.cursor()

        try:
            cursor.execute('''
                INSERT INTO workflow_instances (
                    workflow_id, batch_id, title, status, submitter, description,
                    input_folder, output_path, total_records, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                workflow_id, batch_id, title, WorkflowStatus.DRAFT.value, submitter,
                description, input_folder, output_path, total_records, now, now
            ))

            self._log_action(cursor, workflow_id, WorkflowAction.CREATE.value,
                           submitter, f'创建工作流，批次号: {batch_id}',
                           None, WorkflowStatus.DRAFT.value)

            if exception_items:
                for item in exception_items:
                    cursor.execute('''
                        INSERT INTO exception_items (
                            workflow_id, exception_type, description, transaction_id,
                            amount, counterparty, status, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        workflow_id,
                        item.get('exception_type', 'other'),
                        item.get('description', ''),
                        item.get('transaction_id'),
                        item.get('amount'),
                        item.get('counterparty'),
                        ExceptionStatus.PENDING.value,
                        now, now
                    ))

                self._update_exception_counts(cursor, workflow_id)

            conn.commit()

            self.logger.info('工作流已创建: %s, 批次: %s, 提交人: %s',
                           workflow_id, batch_id, submitter)

            return self.get_workflow(workflow_id)

        except Exception as e:
            conn.rollback()
            self.logger.error('创建工作流失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def get_workflow(self, workflow_id: str) -> WorkflowInstance:
        """获取工作流实例详情

        Raises:
            ValueError: 工作流不存在时抛出
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM workflow_instances WHERE workflow_id = ?',
                         (workflow_id,))
            row = cursor.fetchone()
            if row:
                return WorkflowInstance.from_dict(dict(row))
            raise ValueError(f'工作流不存在: {workflow_id}')
        finally:
            conn.close()

    def list_workflows(self, status: Optional[str] = None,
                      submitter: Optional[str] = None,
                      approver: Optional[str] = None,
                      batch_id: Optional[str] = None,
                      limit: int = 100, offset: int = 0) -> Tuple[List[WorkflowInstance], int]:
        """
        查询工作流列表

        Args:
            status: 状态过滤
            submitter: 提交人过滤
            approver: 审批人过滤
            batch_id: 批次号过滤
            limit: 分页大小
            offset: 分页偏移

        Returns:
            (工作流列表, 总数)
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            conditions = []
            params = []

            if status:
                conditions.append('status = ?')
                params.append(status)
            if submitter:
                conditions.append('submitter = ?')
                params.append(submitter)
            if approver:
                conditions.append('approver = ?')
                params.append(approver)
            if batch_id:
                conditions.append('batch_id = ?')
                params.append(batch_id)

            where_clause = ' WHERE ' + ' AND '.join(conditions) if conditions else ''

            count_query = f'SELECT COUNT(*) as cnt FROM workflow_instances{where_clause}'
            cursor.execute(count_query, params)
            total = cursor.fetchone()['cnt']

            query = f'''
                SELECT * FROM workflow_instances{where_clause}
                ORDER BY created_at DESC
                LIMIT ? OFFSET ?
            '''
            params.extend([limit, offset])
            cursor.execute(query, params)
            rows = cursor.fetchall()

            workflows = [WorkflowInstance.from_dict(dict(row)) for row in rows]
            return workflows, total

        finally:
            conn.close()

    def submit_for_approval(self, workflow_id: str, operator: Optional[str] = None,
                           remark: Optional[str] = None) -> WorkflowInstance:
        """
        提交审批

        Args:
            workflow_id: 工作流ID
            operator: 操作人（默认当前用户）
            remark: 备注

        Returns:
            更新后的工作流实例
        """
        operator = operator or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM workflow_instances WHERE workflow_id = ?',
                         (workflow_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'工作流不存在: {workflow_id}')

            old_status = row['status']
            if old_status not in (WorkflowStatus.DRAFT.value, WorkflowStatus.REJECTED.value):
                raise ValueError(
                    f'当前状态 {old_status} 不允许提交审批，'
                    f'仅草稿或已驳回状态可提交'
                )

            if row['submitter'] != operator:
                raise ValueError(f'仅提交人 {row["submitter"]} 可提交审批')

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            new_status = WorkflowStatus.PENDING_APPROVAL.value

            cursor.execute('''
                UPDATE workflow_instances
                SET status = ?, submitted_at = ?, updated_at = ?
                WHERE workflow_id = ?
            ''', (new_status, now, now, workflow_id))

            self._log_action(cursor, workflow_id, WorkflowAction.SUBMIT.value,
                           operator, remark or '提交审批', old_status, new_status)

            conn.commit()
            self.logger.info('工作流 %s 已提交审批，操作人: %s', workflow_id, operator)

            return self.get_workflow(workflow_id)

        except Exception as e:
            conn.rollback()
            self.logger.error('提交审批失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def approve_workflow(self, workflow_id: str, approver: Optional[str] = None,
                        remark: Optional[str] = None) -> WorkflowInstance:
        """
        复核人审批通过（确认异常清单）

        Args:
            workflow_id: 工作流ID
            approver: 审批人（默认当前用户）
            remark: 备注

        Returns:
            更新后的工作流实例
        """
        approver = approver or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM workflow_instances WHERE workflow_id = ?',
                         (workflow_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'工作流不存在: {workflow_id}')

            old_status = row['status']
            if old_status != WorkflowStatus.PENDING_APPROVAL.value:
                raise ValueError(
                    f'当前状态 {old_status} 不允许审批，'
                    f'仅待审批状态可审批'
                )

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            new_status = WorkflowStatus.APPROVED.value

            cursor.execute('''
                UPDATE workflow_instances
                SET status = ?, approver = ?, approved_at = ?, updated_at = ?
                WHERE workflow_id = ?
            ''', (new_status, approver, now, now, workflow_id))

            self._log_action(cursor, workflow_id, WorkflowAction.APPROVE.value,
                           approver, remark or '审批通过，异常清单已确认',
                           old_status, new_status)

            conn.commit()
            self.logger.info('工作流 %s 已审批通过，审批人: %s', workflow_id, approver)

            return self.get_workflow(workflow_id)

        except Exception as e:
            conn.rollback()
            self.logger.error('审批失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def reject_workflow(self, workflow_id: str, approver: Optional[str] = None,
                       reject_reason: str = '',
                       remark: Optional[str] = None) -> WorkflowInstance:
        """
        驳回工作流

        Args:
            workflow_id: 工作流ID
            approver: 审批人（默认当前用户）
            reject_reason: 驳回原因
            remark: 备注

        Returns:
            更新后的工作流实例
        """
        approver = approver or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM workflow_instances WHERE workflow_id = ?',
                         (workflow_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'工作流不存在: {workflow_id}')

            old_status = row['status']
            if old_status != WorkflowStatus.PENDING_APPROVAL.value:
                raise ValueError(
                    f'当前状态 {old_status} 不允许驳回，'
                    f'仅待审批状态可驳回'
                )

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            new_status = WorkflowStatus.REJECTED.value

            cursor.execute('''
                UPDATE workflow_instances
                SET status = ?, approver = ?, reject_reason = ?, updated_at = ?
                WHERE workflow_id = ?
            ''', (new_status, approver, reject_reason, now, workflow_id))

            action_remark = f'驳回: {reject_reason}'
            if remark:
                action_remark += f' - {remark}'

            self._log_action(cursor, workflow_id, WorkflowAction.REJECT.value,
                           approver, action_remark, old_status, new_status)

            conn.commit()
            self.logger.warning('工作流 %s 已被驳回，审批人: %s, 原因: %s',
                              workflow_id, approver, reject_reason)

            return self.get_workflow(workflow_id)

        except Exception as e:
            conn.rollback()
            self.logger.error('驳回失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def publish_workflow(self, workflow_id: str, publisher: Optional[str] = None,
                        output_path: Optional[str] = None,
                        remark: Optional[str] = None) -> WorkflowInstance:
        """
        正式发布总表

        Args:
            workflow_id: 工作流ID
            publisher: 发布人（默认当前用户）
            output_path: 发布的总表路径（可选，用于更新）
            remark: 备注

        Returns:
            更新后的工作流实例
        """
        publisher = publisher or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM workflow_instances WHERE workflow_id = ?',
                         (workflow_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'工作流不存在: {workflow_id}')

            old_status = row['status']
            if old_status != WorkflowStatus.APPROVED.value:
                raise ValueError(
                    f'当前状态 {old_status} 不允许发布，'
                    f'仅审批通过状态可发布'
                )

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            new_status = WorkflowStatus.PUBLISHED.value

            update_fields = ['status = ?', 'publisher = ?', 'published_at = ?', 'updated_at = ?']
            params = [new_status, publisher, now, now]

            if output_path:
                update_fields.append('output_path = ?')
                params.append(output_path)

            params.append(workflow_id)

            cursor.execute(f'''
                UPDATE workflow_instances
                SET {', '.join(update_fields)}
                WHERE workflow_id = ?
            ''', params)

            self._log_action(cursor, workflow_id, WorkflowAction.PUBLISH.value,
                           publisher, remark or '正式发布总表',
                           old_status, new_status)

            conn.commit()
            self.logger.info('工作流 %s 已正式发布，发布人: %s', workflow_id, publisher)

            return self.get_workflow(workflow_id)

        except Exception as e:
            conn.rollback()
            self.logger.error('发布失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def cancel_workflow(self, workflow_id: str, operator: Optional[str] = None,
                       remark: Optional[str] = None) -> WorkflowInstance:
        """
        取消工作流

        Args:
            workflow_id: 工作流ID
            operator: 操作人（默认当前用户）
            remark: 备注

        Returns:
            更新后的工作流实例
        """
        operator = operator or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('SELECT * FROM workflow_instances WHERE workflow_id = ?',
                         (workflow_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'工作流不存在: {workflow_id}')

            old_status = row['status']
            if old_status == WorkflowStatus.PUBLISHED.value:
                raise ValueError('已发布的工作流无法取消')

            if row['submitter'] != operator:
                raise ValueError(f'仅提交人 {row["submitter"]} 可取消工作流')

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            new_status = WorkflowStatus.CANCELLED.value

            cursor.execute('''
                UPDATE workflow_instances
                SET status = ?, updated_at = ?
                WHERE workflow_id = ?
            ''', (new_status, now, workflow_id))

            self._log_action(cursor, workflow_id, WorkflowAction.CANCEL.value,
                           operator, remark or '取消工作流',
                           old_status, new_status)

            conn.commit()
            self.logger.info('工作流 %s 已取消，操作人: %s', workflow_id, operator)

            return self.get_workflow(workflow_id)

        except Exception as e:
            conn.rollback()
            self.logger.error('取消工作流失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def add_exception_items(self, workflow_id: str,
                           items: List[Dict[str, Any]],
                           operator: Optional[str] = None) -> List[int]:
        """
        添加异常项

        Args:
            workflow_id: 工作流ID
            items: 异常项列表
            operator: 操作人

        Returns:
            新增异常项的ID列表
        """
        operator = operator or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('SELECT status FROM workflow_instances WHERE workflow_id = ?',
                         (workflow_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'工作流不存在: {workflow_id}')

            if row['status'] not in (WorkflowStatus.DRAFT.value, WorkflowStatus.REJECTED.value):
                raise ValueError(
                    f'当前状态 {row["status"]} 不允许添加异常项，'
                    f'仅草稿或已驳回状态可编辑'
                )

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            new_ids = []

            for item in items:
                cursor.execute('''
                    INSERT INTO exception_items (
                        workflow_id, exception_type, description, transaction_id,
                        amount, counterparty, status, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    workflow_id,
                    item.get('exception_type', 'other'),
                    item.get('description', ''),
                    item.get('transaction_id'),
                    item.get('amount'),
                    item.get('counterparty'),
                    ExceptionStatus.PENDING.value,
                    now, now
                ))
                new_ids.append(cursor.lastrowid)

            self._update_exception_counts(cursor, workflow_id)

            self._log_action(cursor, workflow_id, WorkflowAction.UPDATE_EXCEPTION.value,
                           operator, f'添加 {len(items)} 个异常项')

            conn.commit()
            self.logger.info('工作流 %s 添加了 %d 个异常项', workflow_id, len(items))

            return new_ids

        except Exception as e:
            conn.rollback()
            self.logger.error('添加异常项失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def update_exception_item(self, exception_id: int,
                             status: Optional[str] = None,
                             remark: Optional[str] = None,
                             operator: Optional[str] = None) -> bool:
        """
        更新异常项状态（复核人确认异常）

        Args:
            exception_id: 异常项ID
            status: 新状态
            remark: 备注
            operator: 操作人

        Returns:
            是否更新成功
        """
        operator = operator or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('''
                SELECT ei.*, wi.status as workflow_status
                FROM exception_items ei
                JOIN workflow_instances wi ON ei.workflow_id = wi.workflow_id
                WHERE ei.id = ?
            ''', (exception_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'异常项不存在: {exception_id}')

            if row['workflow_status'] != WorkflowStatus.PENDING_APPROVAL.value:
                raise ValueError(
                    f'工作流当前状态 {row["workflow_status"]} 不允许处理异常项，'
                    f'仅待审批状态可处理'
                )

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

            updates = []
            params = []

            if status:
                valid_statuses = {s.value for s in ExceptionStatus}
                if status not in valid_statuses:
                    raise ValueError(f'无效的异常状态: {status}')
                updates.append('status = ?')
                params.append(status)
                updates.append('handled_by = ?')
                params.append(operator)
                updates.append('handled_at = ?')
                params.append(now)

            if remark is not None:
                updates.append('remark = ?')
                params.append(remark)

            if not updates:
                return False

            updates.append('updated_at = ?')
            params.append(now)
            params.append(exception_id)

            cursor.execute(f'''
                UPDATE exception_items
                SET {', '.join(updates)}
                WHERE id = ?
            ''', params)

            self._update_exception_counts(cursor, row['workflow_id'])

            action_remark = f'更新异常项 #{exception_id}'
            if status:
                action_remark += f' 状态: {status}'
            if remark:
                action_remark += f' 备注: {remark}'

            self._log_action(cursor, row['workflow_id'],
                           WorkflowAction.UPDATE_EXCEPTION.value,
                           operator, action_remark)

            conn.commit()
            return cursor.rowcount > 0

        except Exception as e:
            conn.rollback()
            self.logger.error('更新异常项失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def get_exception_items(self, workflow_id: str,
                           status: Optional[str] = None) -> List[ExceptionItem]:
        """
        获取工作流的异常项列表

        Args:
            workflow_id: 工作流ID
            status: 状态过滤

        Returns:
            异常项列表
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            conditions = ['workflow_id = ?']
            params = [workflow_id]

            if status:
                conditions.append('status = ?')
                params.append(status)

            query = f'''
                SELECT * FROM exception_items
                WHERE {' AND '.join(conditions)}
                ORDER BY created_at DESC
            '''
            cursor.execute(query, params)
            rows = cursor.fetchall()

            return [ExceptionItem.from_dict(dict(row)) for row in rows]

        finally:
            conn.close()

    def get_exception_item(self, exception_id: int) -> Optional[ExceptionItem]:
        """获取单个异常项"""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM exception_items WHERE id = ?',
                         (exception_id,))
            row = cursor.fetchone()
            if row:
                return ExceptionItem.from_dict(dict(row))
            return None
        finally:
            conn.close()

    def delete_exception_item(self, exception_id: int,
                             operator: Optional[str] = None) -> bool:
        """
        删除异常项

        Args:
            exception_id: 异常项ID
            operator: 操作人

        Returns:
            是否删除成功
        """
        operator = operator or get_current_user()
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            cursor.execute('''
                SELECT ei.*, wi.status as workflow_status, wi.submitter
                FROM exception_items ei
                JOIN workflow_instances wi ON ei.workflow_id = wi.workflow_id
                WHERE ei.id = ?
            ''', (exception_id,))
            row = cursor.fetchone()
            if not row:
                raise ValueError(f'异常项不存在: {exception_id}')

            if row['workflow_status'] not in (
                WorkflowStatus.DRAFT.value, WorkflowStatus.REJECTED.value
            ):
                raise ValueError(
                    f'工作流当前状态 {row["workflow_status"]} 不允许删除异常项'
                )

            if row['submitter'] != operator:
                raise ValueError(f'仅提交人 {row["submitter"]} 可删除异常项')

            workflow_id = row['workflow_id']

            cursor.execute('DELETE FROM exception_items WHERE id = ?',
                         (exception_id,))

            self._update_exception_counts(cursor, workflow_id)

            self._log_action(cursor, workflow_id,
                           WorkflowAction.UPDATE_EXCEPTION.value,
                           operator, f'删除异常项 #{exception_id}')

            conn.commit()
            return cursor.rowcount > 0

        except Exception as e:
            conn.rollback()
            self.logger.error('删除异常项失败: %s', e, exc_info=True)
            raise
        finally:
            conn.close()

    def get_action_logs(self, workflow_id: str) -> List[WorkflowActionLog]:
        """获取工作流的操作日志"""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT * FROM workflow_action_logs
                WHERE workflow_id = ?
                ORDER BY created_at ASC
            ''', (workflow_id,))
            rows = cursor.fetchall()

            return [WorkflowActionLog.from_dict(dict(row)) for row in rows]

        finally:
            conn.close()

    def get_statistics(self) -> Dict[str, Any]:
        """获取工作流统计信息"""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            stats = {}

            for status in WorkflowStatus:
                cursor.execute(
                    'SELECT COUNT(*) as cnt FROM workflow_instances WHERE status = ?',
                    (status.value,)
                )
                stats[f'{status.value}_count'] = cursor.fetchone()['cnt']

            cursor.execute('SELECT COUNT(*) as cnt FROM workflow_instances')
            stats['total_count'] = cursor.fetchone()['cnt']

            cursor.execute('''
                SELECT COUNT(*) as cnt FROM exception_items
                WHERE status = 'pending'
            ''')
            stats['pending_exceptions'] = cursor.fetchone()['cnt']

            cursor.execute('''
                SELECT submitter, COUNT(*) as cnt
                FROM workflow_instances
                GROUP BY submitter
                ORDER BY cnt DESC
                LIMIT 10
            ''')
            stats['by_submitter'] = [dict(row) for row in cursor.fetchall()]

            cursor.execute('''
                SELECT date(created_at) as dt, COUNT(*) as cnt
                FROM workflow_instances
                WHERE created_at >= date('now', '-30 days')
                GROUP BY date(created_at)
                ORDER BY dt DESC
            ''')
            stats['daily_trend'] = [dict(row) for row in cursor.fetchall()]

            return stats

        finally:
            conn.close()


_global_workflow_manager: Optional[WorkflowManager] = None


def get_workflow_manager(script_dir: Optional[str] = None) -> WorkflowManager:
    """获取全局工作流管理器实例"""
    global _global_workflow_manager
    if _global_workflow_manager is None:
        _global_workflow_manager = WorkflowManager(script_dir)
    return _global_workflow_manager


def reset_workflow_manager():
    """重置全局工作流管理器实例（主要用于测试）"""
    global _global_workflow_manager
    _global_workflow_manager = None


def init_workflow_db(script_dir: Optional[str] = None) -> str:
    """初始化工作流数据库，返回数据库路径"""
    db_path = get_workflow_db_path(script_dir)
    wf = WorkflowManager(script_dir)
    return db_path
