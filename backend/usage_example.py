#!/usr/bin/env python3
"""
任务队列模块使用示例

此脚本展示了如何使用任务队列模块进行异步批处理。

运行前确保：
1. Redis 或 RabbitMQ 已启动
2. 已启动 Worker 进程: python start_worker.py --workers 4
3. 设置环境变量 USE_TASK_QUEUE=true
"""

import os
import sys
import time
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from task_queue import (
    get_job_orchestrator,
    get_message_queue,
    TaskStatus,
    TaskType,
    generate_job_id,
)


def example_basic_usage():
    """基础使用示例"""
    print("=" * 60)
    print("示例 1: 基础作业提交")
    print("=" * 60)

    orchestrator = get_job_orchestrator()
    mq = get_message_queue()

    if not mq.connect():
        print("❌ 无法连接到消息队列，请确保 Redis/RabbitMQ 已启动")
        return

    print("✅ 消息队列已连接，后端:", mq.config.backend)

    with tempfile.TemporaryDirectory() as tmpdir:
        input_dir = Path(tmpdir) / "input"
        input_dir.mkdir()

        for i in range(3):
            (input_dir / f"test_{i}.xlsx").touch()

        job_id = generate_job_id()
        print(f"📋 提交作业: {job_id}")

        def on_progress(job_context):
            print(f"  进度: [{job_context.progress_percent:3d}%] "
                  f"{job_context.progress_stage} - {job_context.progress_message}")

        job_context = orchestrator.submit_job(
            input_folder=str(input_dir),
            script_dir=os.path.dirname(os.path.abspath(__file__)),
            incremental=True,
            keep_strategy='keep_unprocessed',
            operator='test_user',
            user_id='example_user',
            on_progress=on_progress,
        )

        print(f"🚀 作业已提交，ID: {job_context.job_id}")

        while True:
            status = orchestrator.get_job_status(job_context.job_id)
            if status and status.status in [
                TaskStatus.COMPLETED.value,
                TaskStatus.FAILED.value,
                TaskStatus.CANCELLED.value,
            ]:
                break
            time.sleep(0.5)

        final_status = orchestrator.get_job_status(job_context.job_id)
        if final_status:
            print(f"\n✅ 作业完成，状态: {final_status.status}")
            print(f"   总文件数: {final_status.total_files}")
            print(f"   处理文件数: {final_status.processed_files}")
            print(f"   错误文件数: {final_status.error_files}")
            if final_status.output_path:
                print(f"   输出文件: {final_status.output_path}")


def example_queue_status():
    """队列状态查询示例"""
    print("\n" + "=" * 60)
    print("示例 2: 队列状态查询")
    print("=" * 60)

    mq = get_message_queue()

    for queue_name in ['scan', 'parse', 'merge', 'report', 'persist', 'cleanup']:
        try:
            length = mq.get_queue_length(queue_name)
            print(f"  队列 [{queue_name:10s}]: {length} 个任务")
        except Exception as e:
            print(f"  队列 [{queue_name:10s}]: 查询失败 - {e}")


def example_task_priority():
    """任务优先级示例"""
    print("\n" + "=" * 60)
    print("示例 3: 任务优先级")
    print("=" * 60)

    from task_queue import TaskPriority, ScanTaskPayload

    mq = get_message_queue()

    high_priority = ScanTaskPayload(
        task_id=f"high_{int(time.time())}",
        task_type=TaskType.SCAN.value,
        priority=TaskPriority.HIGH.value,
        source_folder="/tmp/test",
        incremental=True,
    )

    low_priority = ScanTaskPayload(
        task_id=f"low_{int(time.time())}",
        task_type=TaskType.SCAN.value,
        priority=TaskPriority.LOW.value,
        source_folder="/tmp/test",
        incremental=True,
    )

    task_id1 = mq.publish_task('scan', high_priority, priority=TaskPriority.HIGH.value)
    task_id2 = mq.publish_task('scan', low_priority, priority=TaskPriority.LOW.value)

    print(f"  高优先级任务: {task_id1}")
    print(f"  低优先级任务: {task_id2}")
    print("  (高优先级任务会被优先处理)")


def example_concurrent_jobs():
    """并发作业示例"""
    print("\n" + "=" * 60)
    print("示例 4: 多用户并发提交")
    print("=" * 60)

    orchestrator = get_job_orchestrator()
    mq = get_message_queue()

    if not mq.connect():
        print("❌ 无法连接到消息队列")
        return

    num_jobs = 5
    print(f"  同时提交 {num_jobs} 个作业...")

    job_ids = []
    for i in range(num_jobs):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_dir = Path(tmpdir) / f"input_{i}"
            input_dir.mkdir()
            (input_dir / "test.xlsx").touch()

            job_context = orchestrator.submit_job(
                input_folder=str(input_dir),
                script_dir=os.path.dirname(os.path.abspath(__file__)),
                incremental=True,
                operator=f'user_{i % 3}',
                user_id=f'user_{i}',
            )
            job_ids.append(job_context.job_id)
            print(f"  用户 {i}: 作业 {job_context.job_id} 已提交")

    print(f"\n  已提交 {len(job_ids)} 个作业")
    print("  所有作业将由 Worker 池并发处理")


def main():
    print("\n" + "🚀" * 30)
    print("任务队列模块使用示例")
    print("🚀" * 30)

    os.environ.setdefault('USE_TASK_QUEUE', 'true')

    try:
        example_queue_status()
        example_basic_usage()
        example_task_priority()
        example_concurrent_jobs()

        print("\n" + "=" * 60)
        print("所有示例运行完成！")
        print("=" * 60)
        print("\n提示:")
        print("  1. 启动 Worker: python start_worker.py --workers 4")
        print("  2. 启用队列: export USE_TASK_QUEUE=true")
        print("  3. 启动 Web: python web_service.py")
        print("  4. 查看队列状态: GET /api/queue/status")
        print("  5. 异步上传文件: POST /api/upload with use_async=true")

    except KeyboardInterrupt:
        print("\n\n用户中断")
    except Exception as e:
        print(f"\n❌ 运行出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
