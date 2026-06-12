# -*- coding: utf-8 -*-
"""
银行流水自动下载脚本模板
========================

本模块提供各银行的下载脚本模板，与 bank_directory_connector 配合使用，
形成端到端的无人值守链路：

    银行网银/银企直连 → 下载脚本 → inbox目录 → 目录对接器 → 处理归档

使用方式：
    1. 复制对应银行的模板，填入实际的登录凭证和接口参数
    2. 配置定时任务（如 crontab）定时执行下载
    3. 下载完成后，文件会自动被目录对接器处理

文件锁定机制：
    下载开始时创建 .lock/{filename}.lock 文件
    下载完成后删除锁定文件，此时目录对接器才会处理该文件
"""

import os
import sys
import time
import logging
import tempfile
from datetime import datetime, timedelta
from typing import Tuple, Dict, Any, Callable, Optional
from pathlib import Path

logger = logging.getLogger('bankcheck.download_templates')


class LockFile:
    """文件锁定上下文管理器，确保下载完成前文件不会被处理"""

    def __init__(self, output_dir: str, filename: str):
        self.output_dir = output_dir
        self.filename = filename
        self.lock_dir = os.path.join(output_dir, '.lock')
        self.lock_file = os.path.join(self.lock_dir, f'{filename}.lock')
        self._acquired = False

    def __enter__(self):
        os.makedirs(self.lock_dir, exist_ok=True)
        with open(self.lock_file, 'w', encoding='utf-8') as f:
            f.write(f'pid: {os.getpid()}\n')
            f.write(f'start_time: {datetime.now().isoformat()}\n')
        self._acquired = True
        logger.info('已创建锁定文件: %s', self.lock_file)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._acquired and os.path.exists(self.lock_file):
            try:
                os.remove(self.lock_file)
                logger.info('已删除锁定文件: %s', self.lock_file)
            except OSError as e:
                logger.warning('删除锁定文件失败: %s', e)
        self._acquired = False


def _safe_write(output_path: str, content: bytes, lock_dir: str) -> None:
    """安全写入文件：先写到临时文件，再原子性重命名"""
    temp_path = output_path + '.tmp'
    try:
        with open(temp_path, 'wb') as f:
            f.write(content)
        os.replace(temp_path, output_path)
        logger.info('文件已写入: %s', output_path)
    except Exception:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        raise


def _generate_filename(bank_name: str) -> str:
    """生成标准的流水文件名：{银行名称}_{YYYYMMDD}_{HHMMSS}.xlsx"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f'{bank_name}_{timestamp}.xlsx'


def _get_date_range(date_range: str) -> Tuple[datetime, datetime]:
    """解析日期范围参数"""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    if date_range == 'today':
        return today, today
    elif date_range == 'yesterday':
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday
    elif date_range == 'last_7_days':
        start = today - timedelta(days=7)
        return start, today - timedelta(days=1)
    elif date_range == 'last_30_days':
        start = today - timedelta(days=30)
        return start, today - timedelta(days=1)
    elif date_range == 'this_month':
        start = today.replace(day=1)
        return start, today
    elif date_range == 'last_month':
        end = today.replace(day=1) - timedelta(days=1)
        start = end.replace(day=1)
        return start, end
    else:
        raise ValueError(f'不支持的日期范围: {date_range}')


def run_download_template(
    template_name: str,
    output_dir: str,
    **kwargs
) -> Tuple[bool, str]:
    """
    执行指定的下载模板

    Args:
        template_name: 模板名称，如 'beijing_bank'
        output_dir: 输出目录（通常是 inbox 目录）
        **kwargs: 模板参数，如 account, date_range 等

    Returns:
        (success: bool, message: str)
    """
    templates = {
        'beijing_bank': download_beijing_bank,
        'east_asia_bank': download_east_asia_bank,
        'icbc_bank': download_icbc_bank,
        'ccb_bank': download_ccb_bank,
        'cmb_bank': download_cmb_bank,
        'mock_beijing_bank': download_mock_beijing_bank,
    }

    template_func = templates.get(template_name)
    if not template_func:
        return False, f'未找到下载模板: {template_name}'

    try:
        return template_func(output_dir=output_dir, **kwargs)
    except Exception as e:
        logger.exception('执行下载模板 %s 失败', template_name)
        return False, f'下载失败: {e}'


def download_beijing_bank(
    output_dir: str,
    account: str = '',
    date_range: str = 'yesterday',
    **kwargs
) -> Tuple[bool, str]:
    """
    北京银行流水下载模板

    实际使用时需要：
    1. 安装北京银行网银客户端或银企直连驱动
    2. 配置登录凭证（建议从环境变量或密钥管理系统读取）
    3. 调用银行提供的 API 或使用自动化工具（如 Selenium）下载

    Args:
        output_dir: 输出目录
        account: 银行账号标识
        date_range: 日期范围

    Returns:
        (success, message)
    """
    logger.info('开始下载北京银行流水，账号: %s, 日期范围: %s', account, date_range)

    try:
        start_date, end_date = _get_date_range(date_range)
        filename = _generate_filename('北京银行')
        output_path = os.path.join(output_dir, filename)

        with LockFile(output_dir, filename):
            # TODO: 替换为实际的北京银行下载逻辑
            # 示例：
            # driver = BankClient(username=os.environ['BEIJING_BANK_USER'],
            #                    password=os.environ['BEIJING_BANK_PASS'])
            # driver.login()
            # data = driver.download_transactions(account, start_date, end_date)
            # _safe_write(output_path, data, os.path.join(output_dir, '.lock'))

            logger.warning('⚠️  这是模板代码，需要替换为实际的北京银行下载逻辑')
            logger.info('模拟下载: 北京银行流水，日期: %s ~ %s',
                       start_date.strftime('%Y-%m-%d'),
                       end_date.strftime('%Y-%m-%d'))
            logger.info('输出文件将保存到: %s', output_path)

            logger.warning('创建模拟文件用于测试...')
            from conftest import _create_beijing_bank_excel
            _create_beijing_bank_excel(output_path)

        logger.info('北京银行流水下载完成: %s', filename)
        return True, f'下载成功: {filename}'

    except Exception as e:
        logger.exception('北京银行下载失败')
        return False, f'下载失败: {e}'


def download_east_asia_bank(
    output_dir: str,
    account: str = '',
    date_range: str = 'yesterday',
    **kwargs
) -> Tuple[bool, str]:
    """东亚银行流水下载模板"""
    logger.info('开始下载东亚银行流水，账号: %s, 日期范围: %s', account, date_range)

    try:
        start_date, end_date = _get_date_range(date_range)
        filename = _generate_filename('东亚银行')
        output_path = os.path.join(output_dir, filename)

        with LockFile(output_dir, filename):
            # TODO: 替换为实际的东亚银行下载逻辑
            logger.warning('⚠️  这是模板代码，需要替换为实际的东亚银行下载逻辑')
            logger.info('模拟下载: 东亚银行流水，日期: %s ~ %s',
                       start_date.strftime('%Y-%m-%d'),
                       end_date.strftime('%Y-%m-%d'))

            logger.warning('创建模拟文件用于测试...')
            from conftest import _create_east_asia_bank_excel
            _create_east_asia_bank_excel(output_path)

        logger.info('东亚银行流水下载完成: %s', filename)
        return True, f'下载成功: {filename}'

    except Exception as e:
        logger.exception('东亚银行下载失败')
        return False, f'下载失败: {e}'


def download_icbc_bank(
    output_dir: str,
    account: str = '',
    date_range: str = 'yesterday',
    **kwargs
) -> Tuple[bool, str]:
    """工商银行流水下载模板"""
    logger.info('开始下载工商银行流水，账号: %s, 日期范围: %s', account, date_range)

    try:
        start_date, end_date = _get_date_range(date_range)
        filename = _generate_filename('工商银行')
        output_path = os.path.join(output_dir, filename)

        with LockFile(output_dir, filename):
            # TODO: 替换为实际的工商银行下载逻辑
            # 工商银行通常提供银企直连 API
            # 可参考: https://open.icbc.com.cn/
            logger.warning('⚠️  这是模板代码，需要替换为实际的工商银行下载逻辑')
            logger.info('模拟下载: 工商银行流水，日期: %s ~ %s',
                       start_date.strftime('%Y-%m-%d'),
                       end_date.strftime('%Y-%m-%d'))

            logger.warning('创建模拟文件用于测试...')
            from conftest import _create_beijing_bank_excel
            _create_beijing_bank_excel(output_path)

        logger.info('工商银行流水下载完成: %s', filename)
        return True, f'下载成功: {filename}'

    except Exception as e:
        logger.exception('工商银行下载失败')
        return False, f'下载失败: {e}'


def download_ccb_bank(
    output_dir: str,
    account: str = '',
    date_range: str = 'yesterday',
    **kwargs
) -> Tuple[bool, str]:
    """建设银行流水下载模板"""
    logger.info('开始下载建设银行流水，账号: %s, 日期范围: %s', account, date_range)

    try:
        start_date, end_date = _get_date_range(date_range)
        filename = _generate_filename('建设银行')
        output_path = os.path.join(output_dir, filename)

        with LockFile(output_dir, filename):
            # TODO: 替换为实际的建设银行下载逻辑
            # 建设银行提供企业网银和银企直连接口
            logger.warning('⚠️  这是模板代码，需要替换为实际的建设银行下载逻辑')
            logger.info('模拟下载: 建设银行流水，日期: %s ~ %s',
                       start_date.strftime('%Y-%m-%d'),
                       end_date.strftime('%Y-%m-%d'))

            logger.warning('创建模拟文件用于测试...')
            from conftest import _create_beijing_bank_excel
            _create_beijing_bank_excel(output_path)

        logger.info('建设银行流水下载完成: %s', filename)
        return True, f'下载成功: {filename}'

    except Exception as e:
        logger.exception('建设银行下载失败')
        return False, f'下载失败: {e}'


def download_cmb_bank(
    output_dir: str,
    account: str = '',
    date_range: str = 'yesterday',
    **kwargs
) -> Tuple[bool, str]:
    """招商银行流水下载模板"""
    logger.info('开始下载招商银行流水，账号: %s, 日期范围: %s', account, date_range)

    try:
        start_date, end_date = _get_date_range(date_range)
        filename = _generate_filename('招商银行')
        output_path = os.path.join(output_dir, filename)

        with LockFile(output_dir, filename):
            # TODO: 替换为实际的招商银行下载逻辑
            # 招商银行提供"企业网银"和"银企直连"两种方式
            # 银企直连 SDK 下载: https://www.cmbchina.com/corporate/
            logger.warning('⚠️  这是模板代码，需要替换为实际的招商银行下载逻辑')
            logger.info('模拟下载: 招商银行流水，日期: %s ~ %s',
                       start_date.strftime('%Y-%m-%d'),
                       end_date.strftime('%Y-%m-%d'))

            logger.warning('创建模拟文件用于测试...')
            from conftest import _create_beijing_bank_excel
            _create_beijing_bank_excel(output_path)

        logger.info('招商银行流水下载完成: %s', filename)
        return True, f'下载成功: {filename}'

    except Exception as e:
        logger.exception('招商银行下载失败')
        return False, f'下载失败: {e}'


def download_mock_beijing_bank(
    output_dir: str,
    account: str = '',
    date_range: str = 'yesterday',
    **kwargs
) -> Tuple[bool, str]:
    """
    模拟北京银行下载（用于测试和演示）
    生成真实的测试数据，不调用任何外部接口
    """
    logger.info('开始模拟下载北京银行流水（测试模式）')

    try:
        start_date, end_date = _get_date_range(date_range)
        filename = _generate_filename('北京银行')
        output_path = os.path.join(output_dir, filename)

        with LockFile(output_dir, filename):
            time.sleep(1)

            from conftest import _create_beijing_bank_excel
            _create_beijing_bank_excel(output_path)

            file_size = os.path.getsize(output_path)
            logger.info('已生成测试文件: %s (%d bytes)', filename, file_size)

        return True, f'模拟下载成功: {filename}'

    except Exception as e:
        logger.exception('模拟下载失败')
        return False, f'模拟下载失败: {e}'


def create_cron_script(
    bank_name: str,
    template_name: str,
    schedule: str = '0 9 * * 1-5',
    config_path: str = './bank_directories.yaml',
    log_file: str = './bank_download.log',
    bank_root: str = './bank_data',
) -> str:
    """
    生成 crontab 脚本内容，用于定时触发银行下载

    Args:
        bank_name: 银行名称
        template_name: 下载模板名称
        schedule: cron 表达式，默认工作日早上9点
        config_path: 配置文件路径
        log_file: 日志文件路径
        bank_root: 银行数据根目录

    Returns:
        crontab 条目字符串
    """
    script_content = f'''
# {bank_name} 流水自动下载 - {schedule}
# 定时触发下载，下载完成后自动由目录对接器处理
BANK_ROOT={bank_root}
{schedule} cd {os.path.dirname(config_path)} && \\
    python bank_directory_connector.py download --bank {bank_name} >> {log_file} 2>&1

# {bank_name} 流水自动处理 - 每30分钟检查一次
*/30 * * * * cd {os.path.dirname(config_path)} && \\
    BANK_ROOT={bank_root} python bank_directory_connector.py run-once >> {log_file} 2>&1
'''
    return script_content.strip()


def create_systemd_service(
    service_name: str = 'bank-directory-watcher',
    config_path: str = './bank_directories.yaml',
    bank_root: str = './bank_data',
    user: str = 'bankcheck',
) -> str:
    """
    生成 systemd service 文件，用于持续监控目录

    Args:
        service_name: 服务名称
        config_path: 配置文件路径
        bank_root: 银行数据根目录
        user: 运行用户

    Returns:
        systemd unit 文件内容
    """
    working_dir = os.path.dirname(config_path)
    unit_content = f'''[Unit]
Description=银行流水目录对接监控服务
After=network.target

[Service]
Type=simple
User={user}
WorkingDirectory={working_dir}
Environment="BANK_ROOT={bank_root}"
ExecStart=/usr/bin/python {working_dir}/bank_directory_connector.py watch --config {config_path}
Restart=always
RestartSec=10
StandardOutput=journal
StandardError=journal
SyslogIdentifier={service_name}

[Install]
WantedBy=multi-user.target
'''
    return unit_content


def main():
    """命令行入口，用于直接执行下载"""
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    parser = argparse.ArgumentParser(description='银行流水下载工具')
    parser.add_argument('--template', required=True,
                       help='下载模板名称，如 beijing_bank')
    parser.add_argument('--output-dir', required=True,
                       help='输出目录（通常是 inbox 目录）')
    parser.add_argument('--account', default='',
                       help='银行账号标识')
    parser.add_argument('--date-range', default='yesterday',
                       choices=['today', 'yesterday', 'last_7_days', 'last_30_days',
                               'this_month', 'last_month'],
                       help='下载日期范围')

    args = parser.parse_args()

    success, message = run_download_template(
        args.template,
        output_dir=args.output_dir,
        account=args.account,
        date_range=args.date_range,
    )

    print(message)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
