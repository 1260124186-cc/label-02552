#!/usr/bin/env python3
"""
任务队列 Worker 启动脚本

Usage:
    python start_worker.py [options]

Options:
    --workers N           启动 N 个 Worker 进程 (默认: 4)
    --queues Q1 Q2 ...    指定要监听的队列 (默认: 所有队列)
    --config PATH         配置文件路径 (默认: task_queue_config.json)
    --backend redis|rabbitmq  指定消息队列后端 (默认: 从配置读取)
    --log-level LEVEL     日志级别: DEBUG, INFO, WARNING, ERROR (默认: INFO)
"""

import argparse
import logging
import os
import signal
import sys
import time
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from task_queue import WorkerManager, load_config
except ImportError as e:
    print(f"导入任务队列模块失败: {e}")
    sys.exit(1)


def setup_logging(level: str):
    """配置日志"""
    numeric_level = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(
        level=numeric_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(process)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def main():
    parser = argparse.ArgumentParser(
        description='任务队列 Worker 管理器',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        '--workers',
        type=int,
        default=None,
        help='启动 N 个 Worker 进程'
    )
    parser.add_argument(
        '--queues',
        nargs='+',
        default=None,
        help='指定要监听的队列'
    )
    parser.add_argument(
        '--config',
        type=str,
        default='task_queue_config.json',
        help='配置文件路径'
    )
    parser.add_argument(
        '--backend',
        type=str,
        choices=['redis', 'rabbitmq'],
        default=None,
        help='指定消息队列后端'
    )
    parser.add_argument(
        '--log-level',
        type=str,
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='日志级别'
    )

    args = parser.parse_args()
    setup_logging(args.log_level)
    logger = logging.getLogger('worker_manager')

    try:
        from self_check import run_self_check, format_report
        report = run_self_check(
            include_optional=False,
            script_dir=str(Path(__file__).parent),
        )
        if not report.passed:
            logger.error('启动自检失败，程序无法启动')
            logger.error('\n%s', format_report(report, verbose=False))
            sys.exit(1)
    except ImportError:
        pass

    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = Path(__file__).parent / args.config

    logger.info('加载配置文件: %s', config_path)
    try:
        config = load_config(str(config_path))
    except Exception as e:
        logger.error('加载配置文件失败: %s', e)
        sys.exit(1)

    if args.backend:
        config.backend = args.backend
        logger.info('使用指定后端: %s', args.backend)

    num_workers = args.workers or config.worker.num_workers
    queues = args.queues or config.worker.queues
    if args.queues:
        config.worker.queues = args.queues

    logger.info('=' * 60)
    logger.info('任务队列 Worker 管理器启动')
    logger.info('=' * 60)
    logger.info('后端: %s', config.backend)
    logger.info('Worker 数量: %d', num_workers)
    logger.info('监听队列: %s', ', '.join(queues))
    logger.info('任务超时: %d 秒', config.worker.task_timeout)
    logger.info('最大重试: %d 次', config.max_retries)
    logger.info('=' * 60)

    try:
        manager = WorkerManager(
            config=config,
            num_workers=num_workers,
            queues=queues,
        )
    except Exception as e:
        logger.error('初始化 Worker 管理器失败: %s', e, exc_info=True)
        sys.exit(1)

    stop_event = [False]

    def handle_signal(signum, frame):
        """处理终止信号"""
        signame = signal.Signals(signum).name
        logger.info('收到信号 %s，正在优雅停止...', signame)
        stop_event[0] = True
        try:
            manager.stop()
        except Exception as e:
            logger.error('停止 Worker 管理器时出错: %s', e)

    signal.signal(signal.SIGINT, handle_signal)
    signal.signal(signal.SIGTERM, handle_signal)

    try:
        manager.start()
        logger.info('Worker 管理器已启动')

        while not stop_event[0]:
            try:
                status = manager.get_status()
                running = sum(1 for w in status['workers'] if w['status'] == 'running')
                total = len(status['workers'])
                logger.debug(
                    'Worker 状态: 运行中 %d/%d, 已处理任务: %d, 失败: %d',
                    running, total,
                    status['total_tasks_processed'],
                    status['total_tasks_failed']
                )
            except Exception:
                pass
            time.sleep(10)

    except Exception as e:
        logger.error('Worker 管理器运行失败: %s', e, exc_info=True)
        sys.exit(1)
    finally:
        logger.info('Worker 管理器已停止')
        sys.exit(0)


if __name__ == '__main__':
    main()
