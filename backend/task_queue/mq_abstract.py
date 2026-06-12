# -*- coding: utf-8 -*-
"""
消息队列抽象层
支持 Redis 和 RabbitMQ 双后端，提供统一的发布/订阅接口
"""

import os
import sys
import json
import time
import logging
import threading
from abc import ABC, abstractmethod
from typing import Dict, Any, Optional, Callable, List
from datetime import datetime

from .config import TaskQueueConfig, load_config
from .models import TaskPayload, TaskResult, TaskStatus, create_payload_from_dict


logger = logging.getLogger('bankcheck.task_queue.mq')


class MessageQueueBackend(ABC):
    """消息队列后端抽象基类"""

    def __init__(self, config: TaskQueueConfig):
        self.config = config
        self._connected = False
        self._lock = threading.Lock()

    @abstractmethod
    def connect(self) -> bool:
        """连接到消息队列"""
        pass

    @abstractmethod
    def disconnect(self) -> None:
        """断开连接"""
        pass

    @abstractmethod
    def publish_task(self, queue_name: str, payload: TaskPayload,
                     priority: Optional[int] = None) -> str:
        """
        发布任务到指定队列

        Args:
            queue_name: 队列名称
            payload: 任务负载
            priority: 优先级（可选，覆盖负载中的优先级）

        Returns:
            任务ID
        """
        pass

    @abstractmethod
    def consume_task(self, queue_names: List[str],
                     timeout: Optional[int] = None) -> Optional[TaskPayload]:
        """
        从队列消费任务

        Args:
            queue_names: 队列名称列表（按优先级排序）
            timeout: 超时时间（秒），None 表示阻塞等待

        Returns:
            TaskPayload 或 None
        """
        pass

    @abstractmethod
    def acknowledge_task(self, task_id: str) -> bool:
        """确认任务已处理完成"""
        pass

    @abstractmethod
    def reject_task(self, task_id: str, requeue: bool = False) -> bool:
        """拒绝任务"""
        pass

    @abstractmethod
    def store_result(self, result: TaskResult) -> bool:
        """存储任务执行结果"""
        pass

    @abstractmethod
    def get_result(self, task_id: str) -> Optional[TaskResult]:
        """获取任务执行结果"""
        pass

    @abstractmethod
    def update_task_status(self, task_id: str, status: str,
                           metadata: Optional[Dict[str, Any]] = None) -> bool:
        """更新任务状态"""
        pass

    @abstractmethod
    def get_task_status(self, task_id: str) -> Optional[str]:
        """获取任务状态"""
        pass

    @abstractmethod
    def cancel_task(self, task_id: str) -> bool:
        """取消任务"""
        pass

    @abstractmethod
    def get_queue_length(self, queue_name: str) -> int:
        """获取队列长度"""
        pass

    @abstractmethod
    def purge_queue(self, queue_name: str) -> bool:
        """清空队列"""
        pass

    def is_connected(self) -> bool:
        return self._connected

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()


class RedisBackend(MessageQueueBackend):
    """Redis 消息队列后端实现"""

    def __init__(self, config: TaskQueueConfig):
        super().__init__(config)
        self._client = None
        self._pubsub = None
        self._task_prefix = 'task:'
        self._result_prefix = 'result:'
        self._status_prefix = 'status:'
        self._queue_prefix = 'queue:'
        self._pending_set = 'pending_tasks'

    def connect(self) -> bool:
        try:
            import redis
            redis_config = self.config.redis
            self._client = redis.Redis(
                host=redis_config.host,
                port=redis_config.port,
                db=redis_config.db,
                password=redis_config.password,
                socket_timeout=redis_config.socket_timeout,
                socket_connect_timeout=redis_config.socket_connect_timeout,
                retry_on_timeout=redis_config.retry_on_timeout,
                max_connections=redis_config.max_connections,
                decode_responses=True,
            )
            self._client.ping()
            self._connected = True
            self._declare_queues()
            logger.info('已连接到 Redis: %s:%d', redis_config.host, redis_config.port)
            return True
        except ImportError:
            logger.error('redis 包未安装，请执行: pip install redis>=5.0.0')
            return False
        except Exception as e:
            logger.error('连接 Redis 失败: %s', e)
            return False

    def _declare_queues(self):
        """声明队列（Redis 中无需显式创建，这里用于初始化相关数据结构）"""
        for queue_name in self.config.queues.keys():
            key = f"{self._queue_prefix}{queue_name}"
            if not self._client.exists(key):
                self._client.lpush(key, '')
                self._client.lpop(key)

    def disconnect(self) -> None:
        if self._client:
            try:
                self._client.close()
            except Exception:
                pass
            self._client = None
        self._connected = False
        logger.info('已断开 Redis 连接')

    def publish_task(self, queue_name: str, payload: TaskPayload,
                     priority: Optional[int] = None) -> str:
        if not self._connected:
            self.connect()

        if priority is not None:
            payload.priority = priority

        task_key = f"{self._task_prefix}{payload.task_id}"
        queue_key = f"{self._queue_prefix}{queue_name}"
        status_key = f"{self._status_prefix}{payload.task_id}"

        with self._lock:
            pipe = self._client.pipeline()
            pipe.hset(task_key, mapping=payload.to_dict())
            pipe.expire(task_key, self.config.result_expire_seconds)

            pipe.zadd(queue_key, {payload.task_id: float(payload.priority)})
            pipe.sadd(self._pending_set, payload.task_id)

            pipe.set(status_key, TaskStatus.QUEUED.value)
            pipe.expire(status_key, self.config.result_expire_seconds)

            pipe.execute()

        logger.info('任务已发布到队列 %s: %s (type=%s, priority=%d)',
                    queue_name, payload.task_id, payload.task_type, payload.priority)
        return payload.task_id

    def consume_task(self, queue_names: List[str],
                     timeout: Optional[int] = None) -> Optional[TaskPayload]:
        if not self._connected:
            self.connect()

        start_time = time.time()
        while True:
            for queue_name in queue_names:
                queue_key = f"{self._queue_prefix}{queue_name}"
                result = self._client.zpopmax(queue_key, 1)

                if result:
                    task_id = result[0][0]
                    task_key = f"{self._task_prefix}{task_id}"
                    task_data = self._client.hgetall(task_key)

                    if task_data:
                        self.update_task_status(task_id, TaskStatus.RUNNING.value)
                        return create_payload_from_dict(task_data)

            if timeout is not None and (time.time() - start_time) > timeout:
                return None

            if timeout is None:
                time.sleep(0.1)
            else:
                return None

    def acknowledge_task(self, task_id: str) -> bool:
        if not self._connected:
            return False

        try:
            with self._lock:
                pipe = self._client.pipeline()
                pipe.srem(self._pending_set, task_id)
                pipe.delete(f"{self._task_prefix}{task_id}")
                pipe.execute()
            return True
        except Exception as e:
            logger.error('确认任务失败 %s: %s', task_id, e)
            return False

    def reject_task(self, task_id: str, requeue: bool = False) -> bool:
        if not self._connected:
            return False

        try:
            task_key = f"{self._task_prefix}{task_id}"
            task_data = self._client.hgetall(task_key)

            if not task_data:
                return False

            payload = create_payload_from_dict(task_data)
            payload.retry_count += 1

            if requeue and payload.retry_count < payload.max_retries:
                self.update_task_status(task_id, TaskStatus.RETRYING.value)
                queue_name = payload.task_type if payload.task_type in self.config.queues else self.config.default_queue
                self.publish_task(queue_name, payload)
                logger.info('任务已重新入队 %s (重试次数: %d/%d)',
                            task_id, payload.retry_count, payload.max_retries)
            else:
                self.update_task_status(task_id, TaskStatus.FAILED.value)
                self.acknowledge_task(task_id)
                logger.warning('任务已达到最大重试次数，标记为失败: %s', task_id)

            return True
        except Exception as e:
            logger.error('拒绝任务失败 %s: %s', task_id, e)
            return False

    def store_result(self, result: TaskResult) -> bool:
        if not self._connected:
            return False

        try:
            result_key = f"{self._result_prefix}{result.task_id}"
            with self._lock:
                pipe = self._client.pipeline()
                pipe.hset(result_key, mapping=result.to_dict())
                pipe.expire(result_key, self.config.result_expire_seconds)

                if result.status == TaskStatus.COMPLETED.value:
                    pipe.set(f"{self._status_prefix}{result.task_id}",
                             TaskStatus.COMPLETED.value)
                    pipe.expire(f"{self._status_prefix}{result.task_id}",
                                self.config.result_expire_seconds)
                elif result.status == TaskStatus.FAILED.value:
                    pipe.set(f"{self._status_prefix}{result.task_id}",
                             TaskStatus.FAILED.value)
                    pipe.expire(f"{self._status_prefix}{result.task_id}",
                                self.config.result_expire_seconds)

                pipe.execute()

            logger.debug('任务结果已存储: %s, status=%s', result.task_id, result.status)
            return True
        except Exception as e:
            logger.error('存储任务结果失败 %s: %s', result.task_id, e)
            return False

    def get_result(self, task_id: str) -> Optional[TaskResult]:
        if not self._connected:
            return None

        try:
            result_key = f"{self._result_prefix}{task_id}"
            result_data = self._client.hgetall(result_key)
            if result_data:
                return TaskResult.from_dict(result_data)
            return None
        except Exception as e:
            logger.error('获取任务结果失败 %s: %s', task_id, e)
            return None

    def update_task_status(self, task_id: str, status: str,
                           metadata: Optional[Dict[str, Any]] = None) -> bool:
        if not self._connected:
            return False

        try:
            status_key = f"{self._status_prefix}{task_id}"
            with self._lock:
                pipe = self._client.pipeline()
                pipe.set(status_key, status)
                pipe.expire(status_key, self.config.result_expire_seconds)
                if metadata:
                    meta_key = f"meta:{task_id}"
                    pipe.hset(meta_key, mapping=metadata)
                    pipe.expire(meta_key, self.config.result_expire_seconds)
                pipe.execute()
            return True
        except Exception as e:
            logger.error('更新任务状态失败 %s: %s', task_id, e)
            return False

    def get_task_status(self, task_id: str) -> Optional[str]:
        if not self._connected:
            return None

        try:
            status_key = f"{self._status_prefix}{task_id}"
            return self._client.get(status_key)
        except Exception as e:
            logger.error('获取任务状态失败 %s: %s', task_id, e)
            return None

    def cancel_task(self, task_id: str) -> bool:
        if not self._connected:
            return False

        try:
            self.update_task_status(task_id, TaskStatus.CANCELLED.value)

            for queue_name in self.config.queues.keys():
                queue_key = f"{self._queue_prefix}{queue_name}"
                self._client.zrem(queue_key, task_id)

            self._client.srem(self._pending_set, task_id)
            logger.info('任务已取消: %s', task_id)
            return True
        except Exception as e:
            logger.error('取消任务失败 %s: %s', task_id, e)
            return False

    def get_queue_length(self, queue_name: str) -> int:
        if not self._connected:
            return 0

        try:
            queue_key = f"{self._queue_prefix}{queue_name}"
            return self._client.zcard(queue_key)
        except Exception:
            return 0

    def purge_queue(self, queue_name: str) -> bool:
        if not self._connected:
            return False

        try:
            queue_key = f"{self._queue_prefix}{queue_name}"
            task_ids = self._client.zrange(queue_key, 0, -1)

            with self._lock:
                pipe = self._client.pipeline()
                for task_id in task_ids:
                    pipe.delete(f"{self._task_prefix}{task_id}")
                    pipe.delete(f"{self._status_prefix}{task_id}")
                    pipe.srem(self._pending_set, task_id)
                pipe.delete(queue_key)
                pipe.execute()

            logger.info('队列已清空: %s (%d 个任务)', queue_name, len(task_ids))
            return True
        except Exception as e:
            logger.error('清空队列失败 %s: %s', queue_name, e)
            return False


class RabbitMQBackend(MessageQueueBackend):
    """RabbitMQ 消息队列后端实现"""

    def __init__(self, config: TaskQueueConfig):
        super().__init__(config)
        self._connection = None
        self._channel = None
        self._exchange_name = 'bankcheck.tasks'
        self._result_store: Dict[str, TaskResult] = {}
        self._status_store: Dict[str, str] = {}
        self._task_store: Dict[str, TaskPayload] = {}

    def connect(self) -> bool:
        try:
            import pika
            rmq_config = self.config.rabbitmq

            credentials = pika.PlainCredentials(
                rmq_config.username, rmq_config.password
            )
            parameters = pika.ConnectionParameters(
                host=rmq_config.host,
                port=rmq_config.port,
                virtual_host=rmq_config.virtual_host,
                credentials=credentials,
                heartbeat=rmq_config.heartbeat,
                connection_attempts=rmq_config.connection_attempts,
                retry_delay=rmq_config.retry_delay,
            )

            self._connection = pika.BlockingConnection(parameters)
            self._channel = self._connection.channel()

            self._channel.exchange_declare(
                exchange=self._exchange_name,
                exchange_type='direct',
                durable=True
            )

            for queue_name, queue_config in self.config.queues.items():
                args = {'x-max-priority': queue_config.max_priority} if queue_config.max_priority else None
                self._channel.queue_declare(
                    queue=queue_name,
                    durable=queue_config.durable,
                    auto_delete=queue_config.auto_delete,
                    arguments=args
                )
                self._channel.queue_bind(
                    exchange=self._exchange_name,
                    queue=queue_name,
                    routing_key=queue_config.routing_key
                )

            self._connected = True
            logger.info('已连接到 RabbitMQ: %s:%d', rmq_config.host, rmq_config.port)
            return True
        except ImportError:
            logger.error('pika 包未安装，请执行: pip install pika>=1.3.0')
            return False
        except Exception as e:
            logger.error('连接 RabbitMQ 失败: %s', e)
            return False

    def disconnect(self) -> None:
        if self._channel:
            try:
                self._channel.close()
            except Exception:
                pass
            self._channel = None
        if self._connection:
            try:
                self._connection.close()
            except Exception:
                pass
            self._connection = None
        self._connected = False
        logger.info('已断开 RabbitMQ 连接')

    def publish_task(self, queue_name: str, payload: TaskPayload,
                     priority: Optional[int] = None) -> str:
        if not self._connected:
            self.connect()

        if priority is not None:
            payload.priority = priority

        queue_config = self.config.queues.get(queue_name)
        if queue_config is None:
            queue_config = self.config.queues[self.config.default_queue]

        properties = pika.BasicProperties(
            delivery_mode=2,
            priority=payload.priority,
            message_id=payload.task_id,
            timestamp=int(time.time()),
            headers={'task_type': payload.task_type, 'job_id': payload.job_id}
        )

        self._task_store[payload.task_id] = payload
        self._status_store[payload.task_id] = TaskStatus.QUEUED.value

        self._channel.basic_publish(
            exchange=self._exchange_name,
            routing_key=queue_config.routing_key,
            body=payload.to_json().encode('utf-8'),
            properties=properties
        )

        logger.info('任务已发布到队列 %s: %s (type=%s, priority=%d)',
                    queue_name, payload.task_id, payload.task_type, payload.priority)
        return payload.task_id

    def consume_task(self, queue_names: List[str],
                     timeout: Optional[int] = None) -> Optional[TaskPayload]:
        if not self._connected:
            self.connect()

        for queue_name in queue_names:
            try:
                method, properties, body = self._channel.basic_get(
                    queue=queue_name,
                    auto_ack=False
                )
                if method:
                    payload = create_payload_from_json(body.decode('utf-8'))
                    self._status_store[payload.task_id] = TaskStatus.RUNNING.value
                    self._current_delivery_tag = method.delivery_tag
                    return payload
            except Exception as e:
                logger.warning('从队列 %s 消费任务失败: %s', queue_name, e)
                continue

        if timeout is None:
            time.sleep(0.1)
            return self.consume_task(queue_names, timeout)

        return None

    def acknowledge_task(self, task_id: str) -> bool:
        if not self._connected:
            return False

        try:
            if hasattr(self, '_current_delivery_tag'):
                self._channel.basic_ack(self._current_delivery_tag)
                delattr(self, '_current_delivery_tag')
            if task_id in self._task_store:
                del self._task_store[task_id]
            return True
        except Exception as e:
            logger.error('确认任务失败 %s: %s', task_id, e)
            return False

    def reject_task(self, task_id: str, requeue: bool = False) -> bool:
        if not self._connected:
            return False

        try:
            if hasattr(self, '_current_delivery_tag'):
                self._channel.basic_nack(
                    self._current_delivery_tag,
                    requeue=requeue
                )
                delattr(self, '_current_delivery_tag')

            payload = self._task_store.get(task_id)
            if payload and requeue:
                payload.retry_count += 1
                if payload.retry_count < payload.max_retries:
                    self._status_store[task_id] = TaskStatus.RETRYING.value
                    queue_name = payload.task_type if payload.task_type in self.config.queues else self.config.default_queue
                    self.publish_task(queue_name, payload)
                else:
                    self._status_store[task_id] = TaskStatus.FAILED.value
                    if task_id in self._task_store:
                        del self._task_store[task_id]
            return True
        except Exception as e:
            logger.error('拒绝任务失败 %s: %s', task_id, e)
            return False

    def store_result(self, result: TaskResult) -> bool:
        try:
            self._result_store[result.task_id] = result
            self._status_store[result.task_id] = result.status
            logger.debug('任务结果已存储: %s, status=%s', result.task_id, result.status)
            return True
        except Exception as e:
            logger.error('存储任务结果失败 %s: %s', result.task_id, e)
            return False

    def get_result(self, task_id: str) -> Optional[TaskResult]:
        return self._result_store.get(task_id)

    def update_task_status(self, task_id: str, status: str,
                           metadata: Optional[Dict[str, Any]] = None) -> bool:
        self._status_store[task_id] = status
        return True

    def get_task_status(self, task_id: str) -> Optional[str]:
        return self._status_store.get(task_id)

    def cancel_task(self, task_id: str) -> bool:
        self._status_store[task_id] = TaskStatus.CANCELLED.value
        if task_id in self._task_store:
            del self._task_store[task_id]
        logger.info('任务已取消: %s', task_id)
        return True

    def get_queue_length(self, queue_name: str) -> int:
        if not self._connected:
            return 0

        try:
            method = self._channel.queue_declare(
                queue=queue_name,
                passive=True
            )
            return method.method.message_count
        except Exception:
            return 0

    def purge_queue(self, queue_name: str) -> bool:
        if not self._connected:
            return False

        try:
            method = self._channel.queue_purge(queue=queue_name)
            logger.info('队列已清空: %s (%d 个任务)',
                        queue_name, method.method.message_count)
            return True
        except Exception as e:
            logger.error('清空队列失败 %s: %s', queue_name, e)
            return False


_global_mq_backend: Optional[MessageQueueBackend] = None


def get_message_queue(config_path: Optional[str] = None,
                      config: Optional[TaskQueueConfig] = None) -> MessageQueueBackend:
    """
    获取消息队列后端实例（单例）

    Args:
        config_path: 配置文件路径
        config: 配置对象（优先使用）

    Returns:
        MessageQueueBackend 实例
    """
    global _global_mq_backend

    if _global_mq_backend is not None and _global_mq_backend.is_connected():
        return _global_mq_backend

    if config is None:
        config = load_config(config_path)

    if config.backend == 'rabbitmq':
        _global_mq_backend = RabbitMQBackend(config)
    else:
        _global_mq_backend = RedisBackend(config)

    _global_mq_backend.connect()
    return _global_mq_backend


def close_message_queue() -> None:
    """关闭全局消息队列连接"""
    global _global_mq_backend
    if _global_mq_backend:
        _global_mq_backend.disconnect()
        _global_mq_backend = None
