#!/usr/bin/env python3
"""
任务队列模块单元测试

运行测试:
    python -m pytest test_task_queue.py -v
    或: python test_task_queue.py
"""

import unittest
import sys
import os
import time
import tempfile
import json
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch, Mock
from dataclasses import asdict, is_dataclass

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from task_queue import (
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
    generate_worker_id,
    TaskQueueConfig,
    load_config,
    get_message_queue,
)
from task_queue.worker import (
    WorkerStats,
)


class TestModels(unittest.TestCase):
    """测试数据模型"""

    def test_generate_ids(self):
        """测试ID生成器"""
        job_id = generate_job_id()
        task_id = generate_task_id()
        worker_id = generate_worker_id()

        self.assertTrue(job_id.startswith('job_'))
        self.assertTrue(task_id.startswith('task_'))
        self.assertTrue(worker_id.startswith('worker_'))

        job_id2 = generate_job_id()
        self.assertNotEqual(job_id, job_id2)

    def test_task_status_enum(self):
        """测试任务状态枚举"""
        self.assertEqual(TaskStatus.PENDING.value, 'pending')
        self.assertEqual(TaskStatus.QUEUED.value, 'queued')
        self.assertEqual(TaskStatus.RUNNING.value, 'running')
        self.assertEqual(TaskStatus.COMPLETED.value, 'completed')
        self.assertEqual(TaskStatus.FAILED.value, 'failed')
        self.assertEqual(TaskStatus.CANCELLED.value, 'cancelled')
        self.assertEqual(TaskStatus.TIMEOUT.value, 'timeout')

    def test_task_type_enum(self):
        """测试任务类型枚举"""
        self.assertEqual(TaskType.SCAN.value, 'scan')
        self.assertEqual(TaskType.PARSE.value, 'parse')
        self.assertEqual(TaskType.MERGE.value, 'merge')
        self.assertEqual(TaskType.REPORT.value, 'report')
        self.assertEqual(TaskType.PERSIST.value, 'persist')
        self.assertEqual(TaskType.CLEANUP.value, 'cleanup')

    def test_task_priority_enum(self):
        """测试任务优先级枚举"""
        self.assertEqual(TaskPriority.LOW.value, 1)
        self.assertEqual(TaskPriority.NORMAL.value, 5)
        self.assertEqual(TaskPriority.HIGH.value, 8)
        self.assertEqual(TaskPriority.CRITICAL.value, 10)

    def test_job_context_serialization(self):
        """测试JobContext序列化和反序列化"""
        job_id = generate_job_id()
        context = JobContext(
            job_id=job_id,
            user_id='test_user',
            operator='tester',
            input_folder='/tmp/test',
            script_dir='/app',
            incremental=True,
            keep_strategy='keep_unprocessed',
        )

        context_dict = context.to_dict()
        self.assertEqual(context_dict['job_id'], job_id)
        self.assertEqual(context_dict['user_id'], 'test_user')
        self.assertEqual(context_dict['incremental'], True)

        context2 = JobContext.from_dict(context_dict)
        self.assertEqual(context2.job_id, job_id)
        self.assertEqual(context2.user_id, 'test_user')
        self.assertEqual(context2.incremental, True)

    def test_task_payload_serialization(self):
        """测试任务负载序列化"""
        payload = ScanTaskPayload(
            task_id=generate_task_id(),
            task_type=TaskType.SCAN.value,
            priority=TaskPriority.NORMAL.value,
            source_folder='/tmp/test',
            recursive=True,
        )

        data = payload.to_dict()
        self.assertEqual(data['task_type'], 'scan')
        self.assertEqual(data['source_folder'], '/tmp/test')

        payload2 = TaskPayload.from_dict(data)
        self.assertEqual(payload2.task_id, payload.task_id)
        self.assertIsInstance(payload2, ScanTaskPayload)

    def test_worker_stats(self):
        """测试Worker统计"""
        stats = WorkerStats(worker_id='worker-1')
        stats.tasks_processed = 10
        stats.tasks_failed = 2
        stats.current_task = 'task-123'

        stats_dict = stats.to_dict()
        self.assertEqual(stats_dict['tasks_processed'], 10)
        self.assertEqual(stats_dict['tasks_failed'], 2)


class TestConfig(unittest.TestCase):
    """测试配置模块"""

    def test_default_config(self):
        """测试默认配置"""
        config = TaskQueueConfig()
        self.assertEqual(config.backend, 'redis')
        self.assertEqual(config.max_retries, 3)
        self.assertEqual(config.result_expire_seconds, 86400)

    def test_load_config_from_file(self):
        """测试从文件加载配置"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump({
                'backend': 'rabbitmq',
                'max_retries': 5,
                'worker': {
                    'task_timeout': 7200,
                    'prefetch_count': 2,
                },
            }, f)
            temp_path = f.name

        try:
            config = load_config(temp_path)
            self.assertEqual(config.backend, 'rabbitmq')
            self.assertEqual(config.max_retries, 5)
            self.assertEqual(config.worker.task_timeout, 7200)
            self.assertEqual(config.worker.prefetch_count, 2)
        finally:
            os.unlink(temp_path)

    def test_config_to_from_dict(self):
        """测试配置序列化"""
        config = TaskQueueConfig()
        config_dict = asdict(config)
        self.assertEqual(config_dict['backend'], 'redis')
        self.assertEqual(config_dict['max_retries'], 3)


class TestTasks(unittest.TestCase):
    """测试任务函数"""

    def setUp(self):
        """测试前置准备"""
        self.temp_dir = tempfile.mkdtemp()
        self.input_dir = Path(self.temp_dir) / 'input'
        self.input_dir.mkdir()

    def tearDown(self):
        """测试后清理"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_task_scan_empty(self):
        """测试扫描空目录"""
        from task_queue.tasks import task_scan_files

        payload = ScanTaskPayload(
            task_id=generate_task_id(),
            task_type=TaskType.SCAN.value,
            priority=TaskPriority.NORMAL.value,
            source_folder=str(self.input_dir),
            recursive=True,
        )

        result = task_scan_files(payload)
        self.assertEqual(result.status, TaskStatus.COMPLETED.value)
        self.assertEqual(result.result['total_files'], 0)

    def test_task_cleanup(self):
        """测试清理任务"""
        from task_queue.tasks import task_cleanup

        test_file = self.input_dir / 'test.txt'
        test_file.write_text('test')
        self.assertTrue(test_file.exists())

        payload = CleanupTaskPayload(
            task_id=generate_task_id(),
            task_type=TaskType.CLEANUP.value,
            priority=TaskPriority.LOW.value,
            working_folder=str(self.input_dir),
            processed_files=[],
            error_files=[],
            unprocessed_files=[str(test_file)],
            strategy='keep_unprocessed',
        )

        result = task_cleanup(payload)
        self.assertEqual(result.status, TaskStatus.COMPLETED.value)


class TestMessageQueue(unittest.TestCase):
    """测试消息队列抽象层"""

    def test_redis_backend_basic(self):
        """测试Redis后端基础功能（Mock）"""
        try:
            import redis
        except ImportError:
            self.skipTest('redis 模块未安装')

        with patch('redis.Redis') as mock_redis_class:
            from task_queue.mq_abstract import RedisBackend

            mock_instance = Mock()
            mock_redis_class.return_value = mock_instance
            mock_instance.ping.return_value = True

            config = TaskQueueConfig()
            backend = RedisBackend(config)

            self.assertTrue(backend.connect())

    def test_get_message_queue_singleton(self):
        """测试消息队列工厂"""
        from task_queue.mq_abstract import get_message_queue

        mq1 = get_message_queue()
        self.assertIsNotNone(mq1)
        self.assertEqual(mq1.config.backend, 'redis')


class TestOrchestrator(unittest.TestCase):
    """测试任务编排器"""

    @patch('task_queue.orchestrator.get_message_queue')
    def test_task_orchestrator_basic(self, mock_get_mq):
        """测试任务编排器基础功能（Mock）"""
        from task_queue.orchestrator import TaskOrchestrator

        mock_mq = Mock()
        mock_mq.connect.return_value = True
        mock_mq.publish_task.return_value = 'task-123'
        mock_get_mq.return_value = mock_mq

        orchestrator = TaskOrchestrator()

        task_id = orchestrator.publish_scan_task(
            job_id='job_test',
            source_folder='/tmp/test',
            priority=TaskPriority.NORMAL.value,
        )
        self.assertIsNotNone(task_id)

    def test_job_context_update(self):
        """测试作业上下文更新"""
        job_id = generate_job_id()
        context = JobContext(
            job_id=job_id,
            user_id='test',
            input_folder='/tmp',
            script_dir='/app',
        )

        context.update_progress('scanning', 10, '开始扫描')
        self.assertEqual(context.progress_stage, 'scanning')
        self.assertEqual(context.progress_percent, 10)
        self.assertEqual(context.progress_message, '开始扫描')

        context.status = TaskStatus.COMPLETED.value
        context.finished_at = datetime.now().isoformat()
        self.assertEqual(context.status, TaskStatus.COMPLETED.value)
        self.assertIsNotNone(context.finished_at)


class TestWorker(unittest.TestCase):
    """测试Worker模块"""

    @patch('task_queue.worker.get_message_queue')
    def test_worker_init(self, mock_get_mq):
        """测试Worker初始化"""
        from task_queue.worker import Worker

        mock_mq = Mock()
        mock_mq.connect.return_value = True
        mock_get_mq.return_value = mock_mq

        config = TaskQueueConfig()
        worker = Worker(config, worker_id='test-worker', queues=['scan'])

        self.assertEqual(worker.worker_id, 'test-worker')
        self.assertEqual(worker.queues, ['scan'])

    @patch('task_queue.worker.get_message_queue')
    def test_worker_manager_init(self, mock_get_mq):
        """测试WorkerManager初始化"""
        from task_queue.worker import WorkerManager

        mock_mq = Mock()
        mock_mq.connect.return_value = True
        mock_get_mq.return_value = mock_mq

        config = TaskQueueConfig()
        config.worker.queues = ['scan', 'parse']
        manager = WorkerManager(config=config, num_workers=2)

        self.assertEqual(manager.num_workers, 2)
        self.assertEqual(manager.config.worker.queues, ['scan', 'parse'])


def run_tests():
    """运行所有测试"""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    suite.addTests(loader.loadTestsFromTestCase(TestModels))
    suite.addTests(loader.loadTestsFromTestCase(TestConfig))
    suite.addTests(loader.loadTestsFromTestCase(TestTasks))
    suite.addTests(loader.loadTestsFromTestCase(TestMessageQueue))
    suite.addTests(loader.loadTestsFromTestCase(TestOrchestrator))
    suite.addTests(loader.loadTestsFromTestCase(TestWorker))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    print("\n" + "=" * 60)
    if result.wasSuccessful():
        print("✅ 所有测试通过！")
    else:
        print(f"❌ 测试失败: {len(result.failures)} 个失败, {len(result.errors)} 个错误")
    print("=" * 60)

    return result.wasSuccessful()


if __name__ == '__main__':
    success = run_tests()
    sys.exit(0 if success else 1)
