# -*- coding: utf-8 -*-
"""
任务队列配置模块
"""

import os
import json
import logging
from dataclasses import dataclass, field, asdict
from typing import Dict, Any, Optional
from pathlib import Path


def get_script_dir():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@dataclass
class RedisConfig:
    host: str = 'localhost'
    port: int = 6379
    db: int = 0
    password: Optional[str] = None
    socket_timeout: int = 5
    socket_connect_timeout: int = 5
    retry_on_timeout: bool = True
    max_connections: int = 50

    def to_url(self) -> str:
        if self.password:
            return f"redis://:{self.password}@{self.host}:{self.port}/{self.db}"
        return f"redis://{self.host}:{self.port}/{self.db}"


@dataclass
class RabbitMQConfig:
    host: str = 'localhost'
    port: int = 5672
    username: str = 'guest'
    password: str = 'guest'
    virtual_host: str = '/'
    heartbeat: int = 60
    connection_attempts: int = 3
    retry_delay: int = 5

    def to_url(self) -> str:
        return (f"amqp://{self.username}:{self.password}"
                f"@{self.host}:{self.port}/{self.virtual_host}")


@dataclass
class QueueConfig:
    name: str
    routing_key: str
    durable: bool = True
    auto_delete: bool = False
    max_priority: int = 10


@dataclass
class WorkerConfig:
    worker_id: str = ''
    queues: list = field(default_factory=lambda: ['default', 'parse', 'report'])
    max_tasks_per_child: int = 1000
    prefetch_count: int = 1
    task_timeout: int = 3600
    graceful_stop_timeout: int = 30


@dataclass
class TaskQueueConfig:
    backend: str = 'redis'
    redis: RedisConfig = field(default_factory=RedisConfig)
    rabbitmq: RabbitMQConfig = field(default_factory=RabbitMQConfig)
    queues: Dict[str, QueueConfig] = field(default_factory=dict)
    worker: WorkerConfig = field(default_factory=WorkerConfig)
    result_expire_seconds: int = 86400
    max_retries: int = 3
    retry_delay_seconds: int = 5
    default_queue: str = 'default'

    def __post_init__(self):
        if not self.queues:
            self.queues = {
                'scan': QueueConfig(name='scan', routing_key='task.scan', max_priority=10),
                'parse': QueueConfig(name='parse', routing_key='task.parse', max_priority=8),
                'merge': QueueConfig(name='merge', routing_key='task.merge', max_priority=7),
                'report': QueueConfig(name='report', routing_key='task.report', max_priority=6),
                'persist': QueueConfig(name='persist', routing_key='task.persist', max_priority=5),
                'cleanup': QueueConfig(name='cleanup', routing_key='task.cleanup', max_priority=3),
                'default': QueueConfig(name='default', routing_key='task.default', max_priority=5),
            }


def load_config(config_path: Optional[str] = None) -> TaskQueueConfig:
    """
    加载配置文件

    Args:
        config_path: 配置文件路径，默认为 script_dir/task_queue_config.json

    Returns:
        TaskQueueConfig 配置对象
    """
    logger = logging.getLogger('bankcheck.task_queue')

    if config_path is None:
        config_path = os.path.join(get_script_dir(), 'task_queue_config.json')

    config = TaskQueueConfig()

    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            if 'backend' in data:
                config.backend = data['backend']

            if 'redis' in data:
                redis_data = data['redis']
                for key, value in redis_data.items():
                    if hasattr(config.redis, key):
                        setattr(config.redis, key, value)

            if 'rabbitmq' in data:
                rabbitmq_data = data['rabbitmq']
                for key, value in rabbitmq_data.items():
                    if hasattr(config.rabbitmq, key):
                        setattr(config.rabbitmq, key, value)

            if 'queues' in data:
                for qname, qdata in data['queues'].items():
                    if qname in config.queues:
                        for key, value in qdata.items():
                            if hasattr(config.queues[qname], key):
                                setattr(config.queues[qname], key, value)
                    else:
                        config.queues[qname] = QueueConfig(**qdata)

            if 'worker' in data:
                worker_data = data['worker']
                for key, value in worker_data.items():
                    if hasattr(config.worker, key):
                        setattr(config.worker, key, value)

            if 'result_expire_seconds' in data:
                config.result_expire_seconds = data['result_expire_seconds']
            if 'max_retries' in data:
                config.max_retries = data['max_retries']
            if 'retry_delay_seconds' in data:
                config.retry_delay_seconds = data['retry_delay_seconds']
            if 'default_queue' in data:
                config.default_queue = data['default_queue']

            logger.info('已加载任务队列配置: %s', config_path)
        except Exception as e:
            logger.warning('加载配置文件失败 %s: %s，使用默认配置', config_path, e)
    else:
        logger.info('配置文件不存在 %s，使用默认配置', config_path)

    return config


def save_config(config: TaskQueueConfig, config_path: Optional[str] = None) -> str:
    """
    保存配置文件

    Args:
        config: 配置对象
        config_path: 配置文件路径

    Returns:
        配置文件路径
    """
    if config_path is None:
        config_path = os.path.join(get_script_dir(), 'task_queue_config.json')

    data = {
        'backend': config.backend,
        'redis': asdict(config.redis),
        'rabbitmq': asdict(config.rabbitmq),
        'queues': {k: asdict(v) for k, v in config.queues.items()},
        'worker': asdict(config.worker),
        'result_expire_seconds': config.result_expire_seconds,
        'max_retries': config.max_retries,
        'retry_delay_seconds': config.retry_delay_seconds,
        'default_queue': config.default_queue,
    }

    os.makedirs(os.path.dirname(config_path), exist_ok=True)
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return config_path
