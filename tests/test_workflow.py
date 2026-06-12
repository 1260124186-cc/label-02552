import os
import sys
import sqlite3
import tempfile
import shutil
from datetime import datetime

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import workflow as workflow_module
from workflow import (
    WorkflowStatus, WorkflowAction, ExceptionStatus,
    WorkflowInstance, ExceptionItem, WorkflowActionLog,
    WorkflowManager, get_workflow_db_path, get_current_user,
    init_workflow_db, get_workflow_manager, reset_workflow_manager
)


@pytest.fixture
def workflow_script_dir(tmp_dir):
    """创建带工作流数据库的脚本目录"""
    script_dir = os.path.join(tmp_dir, 'script_workflow')
    os.makedirs(script_dir, exist_ok=True)
    init_workflow_db(script_dir)
    return script_dir


@pytest.fixture
def wf_manager(workflow_script_dir):
    """获取工作流管理器实例"""
    return WorkflowManager(workflow_script_dir)


class TestWorkflowDBInit:
    """测试数据库初始化"""

    def test_init_workflow_db_creates_tables(self, workflow_script_dir):
        db_path = get_workflow_db_path(workflow_script_dir)
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = {row[0] for row in cursor.fetchall()}

        assert 'workflow_instances' in tables
        assert 'exception_items' in tables
        assert 'workflow_action_logs' in tables

        cursor.close()
        conn.close()

    def test_init_workflow_db_creates_indexes(self, workflow_script_dir):
        db_path = get_workflow_db_path(workflow_script_dir)
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM sqlite_master WHERE type='index'")
        indexes = {row[0] for row in cursor.fetchall()}

        assert 'idx_workflow_status' in indexes
        assert 'idx_workflow_submitter' in indexes
        assert 'idx_workflow_batch' in indexes
        assert 'idx_exception_workflow' in indexes
        assert 'idx_exception_status' in indexes
        assert 'idx_action_workflow' in indexes

        conn.close()

    def test_init_workflow_db_idempotent(self, workflow_script_dir):
        db_path = get_workflow_db_path(workflow_script_dir)
        for _ in range(3):
            init_workflow_db(workflow_script_dir)
        assert os.path.exists(db_path)

    def test_workflow_instances_schema(self, workflow_script_dir):
        db_path = get_workflow_db_path(workflow_script_dir)
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("PRAGMA table_info(workflow_instances)")
        cols = {row[1] for row in cursor.fetchall()}

        required_cols = {
            'workflow_id', 'batch_id', 'title', 'status', 'submitter',
            'approver', 'publisher', 'description', 'total_records',
            'exception_count', 'confirmed_exception_count',
            'input_folder', 'output_path', 'submitted_at', 'approved_at',
            'published_at', 'reject_reason', 'created_at', 'updated_at'
        }
        assert required_cols.issubset(cols)

        conn.close()

    def test_exception_items_schema(self, workflow_script_dir):
        db_path = get_workflow_db_path(workflow_script_dir)
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("PRAGMA table_info(exception_items)")
        cols = {row[1] for row in cursor.fetchall()}

        required_cols = {
            'id', 'workflow_id', 'exception_type', 'description',
            'transaction_id', 'amount', 'counterparty', 'status',
            'remark', 'handled_by', 'handled_at', 'created_at', 'updated_at'
        }
        assert required_cols.issubset(cols)

        conn.close()

    def test_action_logs_schema(self, workflow_script_dir):
        db_path = get_workflow_db_path(workflow_script_dir)
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("PRAGMA table_info(workflow_action_logs)")
        cols = {row[1] for row in cursor.fetchall()}

        required_cols = {
            'id', 'workflow_id', 'action', 'operator', 'remark',
            'old_status', 'new_status', 'created_at'
        }
        assert required_cols.issubset(cols)

        conn.close()

    def test_foreign_keys_enabled(self, wf_manager):
        conn = wf_manager._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("PRAGMA foreign_keys")
            result = cursor.fetchone()
            assert result[0] == 1
        finally:
            conn.close()


class TestGetCurrentUser:
    """测试用户获取"""

    def test_get_current_user_returns_string(self):
        user = get_current_user()
        assert isinstance(user, str)
        assert len(user) > 0

    def test_get_current_user_uses_env_variable(self, monkeypatch):
        monkeypatch.setenv('BANKCHECK_USER', 'test_finance_user')
        assert get_current_user() == 'test_finance_user'
        monkeypatch.delenv('BANKCHECK_USER', raising=False)

    def test_get_current_user_env_empty_falls_back(self, monkeypatch):
        monkeypatch.setenv('BANKCHECK_USER', '')
        user = get_current_user()
        assert user != ''
        monkeypatch.delenv('BANKCHECK_USER', raising=False)


class TestCreateWorkflow:
    """测试创建工作流"""

    def test_create_workflow_basic(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='BATCH001',
            title='2024年1月银行流水处理',
            description='测试工作流创建',
            submitter='accountant_a'
        )

        assert wf.workflow_id.startswith('WF')
        assert wf.batch_id == 'BATCH001'
        assert wf.title == '2024年1月银行流水处理'
        assert wf.status == WorkflowStatus.DRAFT.value
        assert wf.submitter == 'accountant_a'
        assert wf.description == '测试工作流创建'
        assert wf.created_at is not None
        assert wf.updated_at is not None

    def test_create_workflow_default_submitter(self, wf_manager, monkeypatch):
        monkeypatch.setenv('BANKCHECK_USER', 'default_user')
        wf = wf_manager.create_workflow(
            batch_id='BATCH002',
            title='测试默认提交人'
        )
        assert wf.submitter == 'default_user'
        monkeypatch.delenv('BANKCHECK_USER', raising=False)

    def test_create_workflow_with_exception_items(self, wf_manager):
        exceptions = [
            {
                'exception_type': 'amount_mismatch',
                'description': '金额与凭证不符',
                'transaction_id': 'TXN001',
                'amount': 5000.0,
                'counterparty': '供应商A'
            },
            {
                'exception_type': 'missing_subject',
                'description': '找不到对应主体',
                'transaction_id': 'TXN002',
                'amount': -3000.0,
                'counterparty': '客户B'
            }
        ]

        wf = wf_manager.create_workflow(
            batch_id='BATCH003',
            title='带异常项的工作流',
            total_records=100,
            exception_items=exceptions,
            submitter='submitter_1'
        )

        assert wf.exception_count == 2
        assert wf.confirmed_exception_count == 0
        assert wf.total_records == 100

        saved_exceptions = wf_manager.get_exception_items(wf.workflow_id)
        assert len(saved_exceptions) == 2
        assert saved_exceptions[0].exception_type == 'amount_mismatch'
        assert saved_exceptions[0].status == ExceptionStatus.PENDING.value
        assert saved_exceptions[1].exception_type == 'missing_subject'

    def test_create_workflow_generates_unique_ids(self, wf_manager):
        ids = set()
        for i in range(5):
            wf = wf_manager.create_workflow(
                batch_id=f'BATCH{i}',
                title=f'工作流{i}',
                submitter='test_user'
            )
            ids.add(wf.workflow_id)
        assert len(ids) == 5

    def test_create_workflow_creates_action_log(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='BATCH_LOG',
            title='测试操作日志',
            submitter='logger_user'
        )

        logs = wf_manager.get_action_logs(wf.workflow_id)
        assert len(logs) == 1
        assert logs[0].action == WorkflowAction.CREATE.value
        assert logs[0].operator == 'logger_user'
        assert logs[0].old_status is None
        assert logs[0].new_status == WorkflowStatus.DRAFT.value


class TestGetWorkflow:
    """测试查询工作流"""

    def test_get_workflow_exists(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='BATCH_GET',
            title='查询测试',
            submitter='user_get'
        )

        result = wf_manager.get_workflow(wf.workflow_id)
        assert result is not None
        assert result.workflow_id == wf.workflow_id
        assert result.title == '查询测试'

    def test_get_workflow_not_exists(self, wf_manager):
        with pytest.raises(ValueError, match='工作流不存在'):
            wf_manager.get_workflow('WF_NONEXISTENT')

    def test_get_workflow_returns_dataclass(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='BATCH_DC',
            title='数据类测试',
            submitter='user_dc'
        )

        result = wf_manager.get_workflow(wf.workflow_id)
        assert isinstance(result, WorkflowInstance)
        assert hasattr(result, 'to_dict')
        d = result.to_dict()
        assert isinstance(d, dict)
        assert d['workflow_id'] == wf.workflow_id


class TestListWorkflows:
    """测试工作流列表查询"""

    def test_list_workflows_empty(self, wf_manager):
        workflows, total = wf_manager.list_workflows()
        assert total == 0
        assert workflows == []

    def test_list_workflows_pagination(self, wf_manager):
        for i in range(15):
            wf_manager.create_workflow(
                batch_id=f'BATCH{i:03d}',
                title=f'工作流{i}',
                submitter='user_list'
            )

        page1, total = wf_manager.list_workflows(limit=5, offset=0)
        assert total == 15
        assert len(page1) == 5

        page2, _ = wf_manager.list_workflows(limit=5, offset=5)
        assert len(page2) == 5

        page3, _ = wf_manager.list_workflows(limit=5, offset=10)
        assert len(page3) == 5

    def test_list_workflows_filter_by_status(self, wf_manager):
        for i in range(3):
            wf = wf_manager.create_workflow(
                batch_id=f'BATCH_ST{i}',
                title=f'草稿{i}',
                submitter='user_status'
            )
            if i == 0:
                wf_manager.submit_for_approval(wf.workflow_id, operator='user_status')

        drafts, total = wf_manager.list_workflows(status=WorkflowStatus.DRAFT.value)
        assert total == 2

        pending, _ = wf_manager.list_workflows(status=WorkflowStatus.PENDING_APPROVAL.value)
        assert len(pending) == 1

    def test_list_workflows_filter_by_submitter(self, wf_manager):
        wf_manager.create_workflow(
            batch_id='BATCH_S1',
            title='用户A的工作流',
            submitter='user_a'
        )
        wf_manager.create_workflow(
            batch_id='BATCH_S2',
            title='用户B的工作流',
            submitter='user_b'
        )

        user_a_wf, total = wf_manager.list_workflows(submitter='user_a')
        assert total == 1
        assert user_a_wf[0].submitter == 'user_a'

    def test_list_workflows_filter_by_batch_id(self, wf_manager):
        wf_manager.create_workflow(
            batch_id='BATCH_FILTER',
            title='指定批次',
            submitter='user_filter'
        )
        wf_manager.create_workflow(
            batch_id='OTHER_BATCH',
            title='其他批次',
            submitter='user_filter'
        )

        filtered, total = wf_manager.list_workflows(batch_id='BATCH_FILTER')
        assert total == 1
        assert filtered[0].batch_id == 'BATCH_FILTER'

    def test_list_workflows_ordered_by_created_at(self, wf_manager):
        import time
        ids = []
        for i in range(5):
            wf = wf_manager.create_workflow(
                batch_id=f'BATCH_ORD{i}',
                title=f'顺序{i}',
                submitter='user_order'
            )
            ids.append(wf.workflow_id)
            time.sleep(0.01)

        workflows, _ = wf_manager.list_workflows()
        result_ids = [wf.workflow_id for wf in workflows]
        assert result_ids == list(reversed(ids))


class TestStateTransitions:
    """测试状态流转 - 核心功能"""

    def test_full_workflow_draft_to_published(self, wf_manager):
        """完整流程：草稿 → 待审批 → 已审批 → 已发布"""
        submitter = 'finance_clerk'
        reviewer = 'finance_manager'
        publisher = 'finance_director'

        wf = wf_manager.create_workflow(
            batch_id='FULL001',
            title='2024年1月完整流程测试',
            description='测试完整的审批发布流程',
            submitter=submitter,
            total_records=500
        )
        assert wf.status == WorkflowStatus.DRAFT.value

        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)
        assert wf.status == WorkflowStatus.PENDING_APPROVAL.value
        assert wf.submitted_at is not None

        wf = wf_manager.approve_workflow(wf.workflow_id, approver=reviewer)
        assert wf.status == WorkflowStatus.APPROVED.value
        assert wf.approver == reviewer
        assert wf.approved_at is not None

        wf = wf_manager.publish_workflow(wf.workflow_id, publisher=publisher)
        assert wf.status == WorkflowStatus.PUBLISHED.value
        assert wf.publisher == publisher
        assert wf.published_at is not None

    def test_submit_from_rejected_resubmit(self, wf_manager):
        """测试驳回后重新提交"""
        submitter = 'submitter_a'
        reviewer = 'reviewer_a'

        wf = wf_manager.create_workflow(
            batch_id='RESUBMIT001',
            title='重新提交测试',
            submitter=submitter
        )

        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)
        assert wf.status == WorkflowStatus.PENDING_APPROVAL.value

        wf = wf_manager.reject_workflow(
            wf.workflow_id,
            approver=reviewer,
            reject_reason='异常项未处理完整'
        )
        assert wf.status == WorkflowStatus.REJECTED.value
        assert wf.reject_reason == '异常项未处理完整'

        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)
        assert wf.status == WorkflowStatus.PENDING_APPROVAL.value

    def test_cancel_workflow_from_draft(self, wf_manager):
        """测试从草稿状态取消"""
        wf = wf_manager.create_workflow(
            batch_id='CANCEL001',
            title='取消测试',
            submitter='canceller'
        )

        wf = wf_manager.cancel_workflow(wf.workflow_id, operator='canceller')
        assert wf.status == WorkflowStatus.CANCELLED.value

    def test_cancel_workflow_from_pending(self, wf_manager):
        """测试从待审批状态取消"""
        submitter = 'canceller2'
        wf = wf_manager.create_workflow(
            batch_id='CANCEL002',
            title='待审批取消测试',
            submitter=submitter
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)

        wf = wf_manager.cancel_workflow(wf.workflow_id, operator=submitter)
        assert wf.status == WorkflowStatus.CANCELLED.value

    def test_cannot_submit_wrong_status(self, wf_manager):
        """测试无效状态提交"""
        wf = wf_manager.create_workflow(
            batch_id='ERR_SUBMIT',
            title='错误状态测试',
            submitter='submitter'
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')
        wf = wf_manager.approve_workflow(wf.workflow_id, approver='reviewer')

        with pytest.raises(ValueError, match='不允许提交审批'):
            wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

    def test_cannot_approve_wrong_status(self, wf_manager):
        """测试无效状态审批"""
        wf = wf_manager.create_workflow(
            batch_id='ERR_APPROVE',
            title='错误审批测试',
            submitter='submitter'
        )

        with pytest.raises(ValueError, match='不允许审批'):
            wf_manager.approve_workflow(wf.workflow_id, approver='reviewer')

    def test_cannot_reject_wrong_status(self, wf_manager):
        """测试无效状态驳回"""
        wf = wf_manager.create_workflow(
            batch_id='ERR_REJECT',
            title='错误驳回测试',
            submitter='submitter'
        )

        with pytest.raises(ValueError, match='不允许驳回'):
            wf_manager.reject_workflow(wf.workflow_id, approver='reviewer')

    def test_cannot_publish_wrong_status(self, wf_manager):
        """测试无效状态发布"""
        wf = wf_manager.create_workflow(
            batch_id='ERR_PUBLISH',
            title='错误发布测试',
            submitter='submitter'
        )

        with pytest.raises(ValueError, match='不允许发布'):
            wf_manager.publish_workflow(wf.workflow_id, publisher='publisher')

    def test_cannot_cancel_published(self, wf_manager):
        """测试已发布无法取消"""
        submitter = 'submitter'
        wf = wf_manager.create_workflow(
            batch_id='ERR_CANCEL',
            title='已发布取消测试',
            submitter=submitter
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)
        wf = wf_manager.approve_workflow(wf.workflow_id, approver='reviewer')
        wf = wf_manager.publish_workflow(wf.workflow_id, publisher='publisher')

        with pytest.raises(ValueError, match='已发布的工作流无法取消'):
            wf_manager.cancel_workflow(wf.workflow_id, operator=submitter)

    def test_submit_nonexistent_workflow(self, wf_manager):
        with pytest.raises(ValueError, match='工作流不存在'):
            wf_manager.submit_for_approval('WF_NONEXIST', operator='someone')


class TestPermissionControl:
    """测试权限控制"""

    def test_only_submitter_can_submit(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='PERM_SUBMIT',
            title='提交权限测试',
            submitter='owner'
        )

        with pytest.raises(ValueError, match='仅提交人'):
            wf_manager.submit_for_approval(wf.workflow_id, operator='other_person')

        wf_manager.submit_for_approval(wf.workflow_id, operator='owner')

    def test_only_submitter_can_cancel(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='PERM_CANCEL',
            title='取消权限测试',
            submitter='owner'
        )

        with pytest.raises(ValueError, match='仅提交人'):
            wf_manager.cancel_workflow(wf.workflow_id, operator='other_person')

        wf_manager.cancel_workflow(wf.workflow_id, operator='owner')

    def test_reviewer_can_approve_any_pending(self, wf_manager):
        """任何复核人都可以审批待审批的工作流（财务场景下的角色分配）"""
        wf = wf_manager.create_workflow(
            batch_id='PERM_APPROVE',
            title='审批权限测试',
            submitter='submitter'
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        wf = wf_manager.approve_workflow(wf.workflow_id, approver='any_reviewer')
        assert wf.approver == 'any_reviewer'


class TestExceptionItems:
    """测试异常项管理"""

    def test_add_exception_items_to_draft(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_ADD',
            title='添加异常测试',
            submitter='submitter'
        )

        items = [
            {
                'exception_type': 'amount_mismatch',
                'description': '借方贷方不平',
                'transaction_id': 'TXN001',
                'amount': 10000.0,
                'counterparty': '公司A'
            },
            {
                'exception_type': 'date_mismatch',
                'description': '日期与凭证不符',
                'transaction_id': 'TXN002',
                'amount': -5000.0,
                'counterparty': '公司B'
            }
        ]

        new_ids = wf_manager.add_exception_items(wf.workflow_id, items, operator='submitter')
        assert len(new_ids) == 2

        wf_updated = wf_manager.get_workflow(wf.workflow_id)
        assert wf_updated.exception_count == 2

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        assert len(exceptions) == 2
        assert isinstance(exceptions[0], ExceptionItem)

    def test_cannot_add_exceptions_to_pending(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_ERR1',
            title='异常添加错误测试',
            submitter='submitter'
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        items = [{'exception_type': 'test', 'description': '测试'}]

        with pytest.raises(ValueError, match='不允许添加异常项'):
            wf_manager.add_exception_items(wf.workflow_id, items, operator='submitter')

    def test_update_exception_status(self, wf_manager):
        """复核人在待审批状态下确认异常"""
        wf = wf_manager.create_workflow(
            batch_id='EXC_UPDATE',
            title='异常更新测试',
            submitter='submitter',
            exception_items=[
                {'exception_type': 'test', 'description': '测试异常1'},
                {'exception_type': 'test', 'description': '测试异常2'}
            ]
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        exc_id = exceptions[0].id

        result = wf_manager.update_exception_item(
            exc_id,
            status=ExceptionStatus.CONFIRMED.value,
            remark='已核实，确认为异常',
            operator='reviewer'
        )
        assert result is True

        updated_exc = wf_manager.get_exception_item(exc_id)
        assert updated_exc.status == ExceptionStatus.CONFIRMED.value
        assert updated_exc.handled_by == 'reviewer'
        assert updated_exc.handled_at is not None
        assert updated_exc.remark == '已核实，确认为异常'

        wf_updated = wf_manager.get_workflow(wf.workflow_id)
        assert wf_updated.confirmed_exception_count == 1

    def test_update_exception_resolved_status(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_RESOLVE',
            title='异常解决测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': '待解决'}]
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        exc_id = exceptions[0].id

        wf_manager.update_exception_item(
            exc_id,
            status=ExceptionStatus.RESOLVED.value,
            operator='reviewer'
        )

        updated_exc = wf_manager.get_exception_item(exc_id)
        assert updated_exc.status == ExceptionStatus.RESOLVED.value

    def test_update_exception_ignored_status(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_IGNORE',
            title='异常忽略测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': '可忽略'}]
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        exc_id = exceptions[0].id

        wf_manager.update_exception_item(
            exc_id,
            status=ExceptionStatus.IGNORED.value,
            remark='数据误差在容许范围内',
            operator='reviewer'
        )

        updated_exc = wf_manager.get_exception_item(exc_id)
        assert updated_exc.status == ExceptionStatus.IGNORED.value
        assert updated_exc.remark == '数据误差在容许范围内'

    def test_cannot_update_exception_in_draft(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_ERR2',
            title='异常更新错误测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': '测试'}]
        )

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        exc_id = exceptions[0].id

        with pytest.raises(ValueError, match='不允许处理异常项'):
            wf_manager.update_exception_item(
                exc_id,
                status=ExceptionStatus.CONFIRMED.value,
                operator='reviewer'
            )

    def test_delete_exception_item(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_DEL',
            title='删除异常测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': '待删除'}]
        )

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        exc_id = exceptions[0].id

        result = wf_manager.delete_exception_item(exc_id, operator='submitter')
        assert result is True

        assert wf_manager.get_exception_item(exc_id) is None

        wf_updated = wf_manager.get_workflow(wf.workflow_id)
        assert wf_updated.exception_count == 0

    def test_cannot_delete_exception_wrong_status(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_ERR3',
            title='删除异常错误测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': '测试'}]
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        exc_id = exceptions[0].id

        with pytest.raises(ValueError, match='不允许删除异常项'):
            wf_manager.delete_exception_item(exc_id, operator='submitter')

    def test_cannot_delete_exception_not_submitter(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_ERR4',
            title='删除异常权限测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': '测试'}]
        )

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        exc_id = exceptions[0].id

        with pytest.raises(ValueError, match='仅提交人'):
            wf_manager.delete_exception_item(exc_id, operator='other_user')

    def test_get_exception_items_filter_by_status(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='EXC_FILTER',
            title='异常过滤测试',
            submitter='submitter',
            exception_items=[
                {'exception_type': 't1', 'description': '待处理1'},
                {'exception_type': 't2', 'description': '待处理2'},
                {'exception_type': 't3', 'description': '待处理3'}
            ]
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        wf_manager.update_exception_item(
            exceptions[0].id,
            status=ExceptionStatus.CONFIRMED.value,
            operator='reviewer'
        )
        wf_manager.update_exception_item(
            exceptions[1].id,
            status=ExceptionStatus.RESOLVED.value,
            operator='reviewer'
        )

        pending = wf_manager.get_exception_items(
            wf.workflow_id,
            status=ExceptionStatus.PENDING.value
        )
        assert len(pending) == 1

        confirmed = wf_manager.get_exception_items(
            wf.workflow_id,
            status=ExceptionStatus.CONFIRMED.value
        )
        assert len(confirmed) == 1


class TestActionLogs:
    """测试操作日志"""

    def test_full_workflow_creates_complete_audit_trail(self, wf_manager):
        """完整流程应该生成完整的审计追踪"""
        submitter = 'auditor_sub'
        reviewer = 'auditor_rev'
        publisher = 'auditor_pub'

        wf = wf_manager.create_workflow(
            batch_id='AUDIT001',
            title='审计追踪测试',
            submitter=submitter
        )
        wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)
        wf_manager.approve_workflow(wf.workflow_id, approver=reviewer)
        wf_manager.publish_workflow(wf.workflow_id, publisher=publisher)

        logs = wf_manager.get_action_logs(wf.workflow_id)
        assert len(logs) == 4

        actions = [log.action for log in logs]
        assert WorkflowAction.CREATE.value in actions
        assert WorkflowAction.SUBMIT.value in actions
        assert WorkflowAction.APPROVE.value in actions
        assert WorkflowAction.PUBLISH.value in actions

        assert logs[0].operator == submitter
        assert logs[1].operator == submitter
        assert logs[2].operator == reviewer
        assert logs[3].operator == publisher

        status_transitions = [
            (logs[0].old_status, logs[0].new_status),
            (logs[1].old_status, logs[1].new_status),
            (logs[2].old_status, logs[2].new_status),
            (logs[3].old_status, logs[3].new_status),
        ]
        assert status_transitions == [
            (None, WorkflowStatus.DRAFT.value),
            (WorkflowStatus.DRAFT.value, WorkflowStatus.PENDING_APPROVAL.value),
            (WorkflowStatus.PENDING_APPROVAL.value, WorkflowStatus.APPROVED.value),
            (WorkflowStatus.APPROVED.value, WorkflowStatus.PUBLISHED.value),
        ]

    def test_reject_logs_reason(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='AUDIT_REJ',
            title='驳回日志测试',
            submitter='submitter'
        )
        wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')
        wf_manager.reject_workflow(
            wf.workflow_id,
            approver='reviewer',
            reject_reason='需要补充更多说明文档',
            remark='请于周五前修改'
        )

        logs = wf_manager.get_action_logs(wf.workflow_id)
        reject_log = [l for l in logs if l.action == WorkflowAction.REJECT.value][0]
        assert '需要补充更多说明文档' in reject_log.remark
        assert '请于周五前修改' in reject_log.remark

    def test_exception_update_logs(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='AUDIT_EXC',
            title='异常操作日志测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': '测试'}]
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        wf_manager.update_exception_item(
            exceptions[0].id,
            status=ExceptionStatus.CONFIRMED.value,
            operator='reviewer'
        )

        logs = wf_manager.get_action_logs(wf.workflow_id)
        update_logs = [l for l in logs if l.action == WorkflowAction.UPDATE_EXCEPTION.value]
        assert len(update_logs) >= 1
        assert update_logs[-1].operator == 'reviewer'
        assert '状态: confirmed' in update_logs[-1].remark

    def test_action_logs_ordered_by_time(self, wf_manager):
        import time
        wf = wf_manager.create_workflow(
            batch_id='AUDIT_ORDER',
            title='日志顺序测试',
            submitter='submitter'
        )
        time.sleep(0.01)
        wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')
        time.sleep(0.01)
        wf_manager.approve_workflow(wf.workflow_id, approver='reviewer')

        logs = wf_manager.get_action_logs(wf.workflow_id)
        for i in range(len(logs) - 1):
            assert logs[i].created_at <= logs[i + 1].created_at


class TestStatistics:
    """测试统计信息"""

    def test_get_statistics_empty(self, wf_manager):
        stats = wf_manager.get_statistics()
        assert isinstance(stats, dict)
        assert stats['total_count'] == 0
        assert stats[f'{WorkflowStatus.DRAFT.value}_count'] == 0
        assert stats['pending_exceptions'] == 0

    def test_get_statistics_with_workflows(self, wf_manager):
        for i in range(5):
            wf = wf_manager.create_workflow(
                batch_id=f'STAT{i}',
                title=f'统计测试{i}',
                submitter=f'user_{i % 2}',
                exception_items=[{'exception_type': 'test', 'description': '测试'}]
            )
            if i < 2:
                wf_manager.submit_for_approval(wf.workflow_id, operator=f'user_{i % 2}')
                if i == 0:
                    wf_manager.approve_workflow(wf.workflow_id, approver='reviewer')

        stats = wf_manager.get_statistics()

        assert stats['total_count'] == 5
        assert stats[f'{WorkflowStatus.DRAFT.value}_count'] == 3
        assert stats[f'{WorkflowStatus.PENDING_APPROVAL.value}_count'] == 1
        assert stats[f'{WorkflowStatus.APPROVED.value}_count'] == 1
        assert stats[f'{WorkflowStatus.PUBLISHED.value}_count'] == 0
        assert stats['pending_exceptions'] == 5

        assert 'by_submitter' in stats
        assert len(stats['by_submitter']) <= 10

        assert 'daily_trend' in stats


class TestGlobalManager:
    """测试全局管理器"""

    def test_get_workflow_manager_returns_singleton(self, tmp_dir):
        m1 = get_workflow_manager(tmp_dir)
        m2 = get_workflow_manager(tmp_dir)
        assert m1 is m2

    def test_get_workflow_manager_creates_db(self, tmp_dir):
        reset_workflow_manager()
        get_workflow_manager(tmp_dir)
        db_path = get_workflow_db_path(tmp_dir)
        assert os.path.exists(db_path)


class TestExceptionStatusEnum:
    """测试异常状态枚举"""

    def test_exception_status_has_all_values(self):
        expected = {'pending', 'confirmed', 'resolved', 'ignored'}
        actual = {s.value for s in ExceptionStatus}
        assert expected == actual

    def test_workflow_status_has_all_values(self):
        expected = {'draft', 'pending_approval', 'approved', 'rejected', 'published', 'cancelled'}
        actual = {s.value for s in WorkflowStatus}
        assert expected == actual

    def test_workflow_action_has_all_values(self):
        actions = {a.value for a in WorkflowAction}
        assert 'create' in actions
        assert 'submit' in actions
        assert 'approve' in actions
        assert 'reject' in actions
        assert 'publish' in actions
        assert 'cancel' in actions


class TestDataclassMethods:
    """测试数据类方法"""

    def test_workflow_to_dict(self):
        wf = WorkflowInstance(
            workflow_id='WF001',
            batch_id='B001',
            title='Test',
            status='draft',
            submitter='user1'
        )
        d = wf.to_dict()
        assert isinstance(d, dict)
        assert d['workflow_id'] == 'WF001'
        assert d['batch_id'] == 'B001'

    def test_workflow_from_dict(self):
        data = {
            'workflow_id': 'WF002',
            'batch_id': 'B002',
            'title': 'From Dict',
            'status': 'pending_approval',
            'submitter': 'user2',
            'extra_field': 'should_be_ignored'
        }
        wf = WorkflowInstance.from_dict(data)
        assert wf.workflow_id == 'WF002'
        assert wf.batch_id == 'B002'
        assert not hasattr(wf, 'extra_field')

    def test_exception_item_to_dict(self):
        exc = ExceptionItem(
            id=1,
            workflow_id='WF003',
            exception_type='test',
            description='Test exception'
        )
        d = exc.to_dict()
        assert d['id'] == 1
        assert d['exception_type'] == 'test'

    def test_exception_item_from_dict(self):
        data = {
            'id': 2,
            'workflow_id': 'WF004',
            'exception_type': 'amount_mismatch',
            'description': '金额不符',
            'extra': 'ignored'
        }
        exc = ExceptionItem.from_dict(data)
        assert exc.id == 2
        assert exc.exception_type == 'amount_mismatch'
        assert not hasattr(exc, 'extra')


class TestIntegrationScenarios:
    """集成测试场景"""

    def test_scenario_monthly_closing_process(self, wf_manager):
        """
        财务月结场景：
        1. 会计创建工作流并提交
        2. 系统自动识别异常项
        3. 复核人逐一确认异常
        4. 复核人审批通过
        5. 财务总监发布总表
        """
        accountant = 'accountant_zhang'
        reviewer = 'reviewer_li'
        director = 'director_wang'

        wf = wf_manager.create_workflow(
            batch_id='MONTHLY_202401',
            title='2024年1月银行流水月结',
            description='包含5个银行账户，共1200笔交易',
            submitter=accountant,
            total_records=1200,
            exception_items=[
                {
                    'exception_type': 'amount_mismatch',
                    'description': '转账金额与发票金额差0.01元',
                    'transaction_id': 'TXN20240115001',
                    'amount': 10000.00,
                    'counterparty': '供应商甲公司'
                },
                {
                    'exception_type': 'missing_subject',
                    'description': '新开户银行，未在主体映射表中',
                    'transaction_id': 'TXN20240120002',
                    'amount': -50000.00,
                    'counterparty': '新客户公司'
                },
                {
                    'exception_type': 'date_mismatch',
                    'description': '交易日期跨月，需确认归属期',
                    'transaction_id': 'TXN20240131003',
                    'amount': 8000.00,
                    'counterparty': '客户乙公司'
                }
            ]
        )
        assert wf.exception_count == 3

        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=accountant)
        assert wf.status == WorkflowStatus.PENDING_APPROVAL.value

        exceptions = wf_manager.get_exception_items(wf.workflow_id)
        wf_manager.update_exception_item(
            exceptions[0].id,
            status=ExceptionStatus.IGNORED.value,
            remark='尾差0.01元，属于正常误差范围',
            operator=reviewer
        )
        wf_manager.update_exception_item(
            exceptions[1].id,
            status=ExceptionStatus.RESOLVED.value,
            remark='已添加主体映射，对应主体为"新客户公司"',
            operator=reviewer
        )
        wf_manager.update_exception_item(
            exceptions[2].id,
            status=ExceptionStatus.CONFIRMED.value,
            remark='确认为1月交易，按权责发生制计入1月',
            operator=reviewer
        )

        wf = wf_manager.get_workflow(wf.workflow_id)
        assert wf.confirmed_exception_count == 3

        wf = wf_manager.approve_workflow(
            wf.workflow_id,
            approver=reviewer,
            remark='所有异常已处理完毕，同意发布'
        )
        assert wf.status == WorkflowStatus.APPROVED.value

        wf = wf_manager.publish_workflow(
            wf.workflow_id,
            publisher=director,
            output_path='/finance/2024_01_银行流水总表.xlsx',
            remark='2024年1月银行流水总表正式发布'
        )
        assert wf.status == WorkflowStatus.PUBLISHED.value
        assert wf.output_path == '/finance/2024_01_银行流水总表.xlsx'

        logs = wf_manager.get_action_logs(wf.workflow_id)
        assert len(logs) >= 7
        assert all(log.created_at is not None for log in logs)

    def test_scenario_reject_and_resubmit(self, wf_manager):
        """驳回后修改并重新提交的场景"""
        submitter = 'junior_accountant'
        reviewer = 'senior_accountant'

        wf = wf_manager.create_workflow(
            batch_id='Q1_REVIEW',
            title='2024年Q1银行流水复核',
            submitter=submitter,
            exception_items=[
                {'exception_type': 'test', 'description': '异常1'}
            ]
        )

        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)

        wf = wf_manager.reject_workflow(
            wf.workflow_id,
            approver=reviewer,
            reject_reason='缺少第3个银行账户的流水数据',
            remark='请补充完整后重新提交'
        )
        assert wf.status == WorkflowStatus.REJECTED.value
        assert wf.reject_reason == '缺少第3个银行账户的流水数据'

        wf_manager.add_exception_items(
            wf.workflow_id,
            [{'exception_type': 'missing_file', 'description': '补充第3个银行账户流水'}],
            operator=submitter
        )

        wf = wf_manager.submit_for_approval(wf.workflow_id, operator=submitter)
        assert wf.status == WorkflowStatus.PENDING_APPROVAL.value
        assert wf.exception_count == 2

        logs = wf_manager.get_action_logs(wf.workflow_id)
        actions = [log.action for log in logs]
        assert actions.count(WorkflowAction.SUBMIT.value) == 2
        assert WorkflowAction.REJECT.value in actions

    def test_scenario_cancel_flow(self, wf_manager):
        """创建后发现错误取消的场景"""
        submitter = 'accountant_wang'

        wf = wf_manager.create_workflow(
            batch_id='WRONG_BATCH',
            title='错误批次 - 应该是2月的数据',
            submitter=submitter
        )

        wf = wf_manager.cancel_workflow(
            wf.workflow_id,
            operator=submitter,
            remark='批次选错了，应该是2月的数据，重新创建'
        )

        assert wf.status == WorkflowStatus.CANCELLED.value

        logs = wf_manager.get_action_logs(wf.workflow_id)
        assert logs[-1].action == WorkflowAction.CANCEL.value
        assert logs[-1].remark == '批次选错了，应该是2月的数据，重新创建'

    def test_nonexistent_workflow_raises_error(self, wf_manager):
        with pytest.raises(ValueError, match='工作流不存在'):
            wf_manager.get_workflow('NONEXISTENT')

        with pytest.raises(ValueError, match='工作流不存在'):
            wf_manager.submit_for_approval('NONEXISTENT', operator='test')

    def test_nonexistent_exception_item_raises_error(self, wf_manager):
        with pytest.raises(ValueError, match='异常项不存在'):
            wf_manager.update_exception_item(99999, status='confirmed', operator='test')

        with pytest.raises(ValueError, match='异常项不存在'):
            wf_manager.delete_exception_item(99999, operator='test')

    def test_invalid_exception_status_raises_error(self, wf_manager):
        wf = wf_manager.create_workflow(
            batch_id='INVALID_STATUS',
            title='无效状态测试',
            submitter='submitter',
            exception_items=[{'exception_type': 'test', 'description': 'test'}]
        )
        wf = wf_manager.submit_for_approval(wf.workflow_id, operator='submitter')
        exceptions = wf_manager.get_exception_items(wf.workflow_id)

        with pytest.raises(ValueError, match='无效的异常状态'):
            wf_manager.update_exception_item(
                exceptions[0].id,
                status='invalid_status',
                operator='reviewer'
            )
