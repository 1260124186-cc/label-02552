# -*- coding: utf-8 -*-
"""
输出文件加密模块 - 对总表与检验报告设置打开密码或整文件加密
满足涉密财务数据在邮件、网盘传递时的基本安全要求

功能：
  1. Excel 打开密码保护（基于 msoffcrypto-python，可选依赖）
  2. 整文件 AES-256-GCM 加密（基于 cryptography，已有依赖）
  3. 密码强度校验
  4. 批量加密输出文件（总表、检验报告等）
  5. 加密结果审计日志
"""

import os
import sys
import json
import logging
import hashlib
import secrets
import base64
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import List, Optional, Dict, Any, Tuple

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    HAS_CRYPTOGRAPHY = True
except ImportError:
    HAS_CRYPTOGRAPHY = False

try:
    import msoffcrypto
    HAS_MSOFFCRYPTO = True
except ImportError:
    HAS_MSOFFCRYPTO = False

try:
    from pii_classifier import PIILogFilter
    HAS_PII_CLASSIFIER = True
except ImportError:
    HAS_PII_CLASSIFIER = False


ENCRYPTION_MARKER = b'BKENC01'
PBKDF2_ITERATIONS = 600000
AES_KEY_SIZE = 32
SALT_SIZE = 16
NONCE_SIZE = 12
ENCRYPTED_EXTENSION = '.enc'


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    logger = logging.getLogger('bankcheck.file_encryption')
    if HAS_PII_CLASSIFIER:
        pii_filter = PIILogFilter()
        for h in logger.handlers:
            if not any(isinstance(f, PIILogFilter) for f in h.filters):
                h.addFilter(pii_filter)
    return logger


@dataclass
class EncryptionResult:
    file_path: str
    encrypted_path: Optional[str] = None
    mode: str = ''
    success: bool = False
    error: Optional[str] = None
    original_size: int = 0
    encrypted_size: int = 0
    file_hash: Optional[str] = None
    timestamp: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class BatchEncryptionResult:
    results: List[EncryptionResult] = field(default_factory=list)
    total_files: int = 0
    success_count: int = 0
    failure_count: int = 0
    mode: str = ''
    timestamp: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def validate_password_strength(password: str) -> Tuple[bool, str]:
    if not password:
        return False, '密码不能为空'
    if len(password) < 6:
        return False, '密码长度不能少于6位'
    if len(password) > 128:
        return False, '密码长度不能超过128位'
    has_letter = any(c.isalpha() for c in password)
    has_digit = any(c.isdigit() for c in password)
    if not has_letter or not has_digit:
        return False, '密码必须同时包含字母和数字'
    return True, ''


def _derive_key(password: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=AES_KEY_SIZE,
        salt=salt,
        iterations=PBKDF2_ITERATIONS,
    )
    return kdf.derive(password.encode('utf-8'))


def _compute_file_hash(filepath: str) -> Optional[str]:
    if not filepath or not os.path.exists(filepath):
        return None
    sha256 = hashlib.sha256()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                sha256.update(chunk)
        return sha256.hexdigest()
    except Exception:
        return None


def encrypt_file_aes(file_path: str, password: str, output_path: Optional[str] = None) -> EncryptionResult:
    logger = get_logger()

    if not HAS_CRYPTOGRAPHY:
        return EncryptionResult(
            file_path=file_path,
            mode='aes256gcm',
            success=False,
            error='cryptography 库未安装，无法执行AES加密',
        )

    if not os.path.isfile(file_path):
        return EncryptionResult(
            file_path=file_path,
            mode='aes256gcm',
            success=False,
            error=f'文件不存在: {file_path}',
        )

    valid, msg = validate_password_strength(password)
    if not valid:
        return EncryptionResult(
            file_path=file_path,
            mode='aes256gcm',
            success=False,
            error=msg,
        )

    if output_path is None:
        output_path = file_path + ENCRYPTED_EXTENSION

    try:
        original_size = os.path.getsize(file_path)
        file_hash = _compute_file_hash(file_path)

        salt = secrets.token_bytes(SALT_SIZE)
        key = _derive_key(password, salt)
        nonce = secrets.token_bytes(NONCE_SIZE)

        aesgcm = AESGCM(key)

        with open(file_path, 'rb') as f:
            plaintext = f.read()

        ciphertext = aesgcm.encrypt(nonce, plaintext, None)

        header = ENCRYPTION_MARKER
        with open(output_path, 'wb') as f:
            f.write(header)
            f.write(salt)
            f.write(nonce)
            f.write(ciphertext)

        encrypted_size = os.path.getsize(output_path)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        logger.info('文件AES加密完成: %s -> %s (原始 %s 字节, 加密后 %s 字节)',
                     file_path, output_path, original_size, encrypted_size)

        return EncryptionResult(
            file_path=file_path,
            encrypted_path=output_path,
            mode='aes256gcm',
            success=True,
            original_size=original_size,
            encrypted_size=encrypted_size,
            file_hash=file_hash,
            timestamp=timestamp,
        )

    except Exception as e:
        logger.error('AES加密失败 %s: %s', file_path, e, exc_info=True)
        return EncryptionResult(
            file_path=file_path,
            mode='aes256gcm',
            success=False,
            error=str(e),
        )


def decrypt_file_aes(encrypted_path: str, password: str, output_path: Optional[str] = None) -> EncryptionResult:
    logger = get_logger()

    if not HAS_CRYPTOGRAPHY:
        return EncryptionResult(
            file_path=encrypted_path,
            mode='aes256gcm',
            success=False,
            error='cryptography 库未安装，无法执行AES解密',
        )

    if not os.path.isfile(encrypted_path):
        return EncryptionResult(
            file_path=encrypted_path,
            mode='aes256gcm',
            success=False,
            error=f'文件不存在: {encrypted_path}',
        )

    if output_path is None:
        if encrypted_path.endswith(ENCRYPTED_EXTENSION):
            output_path = encrypted_path[:-len(ENCRYPTED_EXTENSION)]
        else:
            output_path = encrypted_path + '.decrypted'

    try:
        with open(encrypted_path, 'rb') as f:
            header = f.read(len(ENCRYPTION_MARKER))
            if header != ENCRYPTION_MARKER:
                return EncryptionResult(
                    file_path=encrypted_path,
                    mode='aes256gcm',
                    success=False,
                    error='文件格式不正确，非本工具加密的文件',
                )

            salt = f.read(SALT_SIZE)
            nonce = f.read(NONCE_SIZE)
            ciphertext = f.read()

        key = _derive_key(password, salt)
        aesgcm = AESGCM(key)

        plaintext = aesgcm.decrypt(nonce, ciphertext, None)

        with open(output_path, 'wb') as f:
            f.write(plaintext)

        file_hash = _compute_file_hash(output_path)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        logger.info('文件AES解密完成: %s -> %s', encrypted_path, output_path)

        return EncryptionResult(
            file_path=encrypted_path,
            encrypted_path=output_path,
            mode='aes256gcm',
            success=True,
            original_size=os.path.getsize(encrypted_path),
            encrypted_size=os.path.getsize(output_path),
            file_hash=file_hash,
            timestamp=timestamp,
        )

    except Exception as e:
        logger.error('AES解密失败 %s: %s', encrypted_path, e, exc_info=True)
        if 'Tag check failed' in str(e) or 'InvalidTag' in type(e).__name__:
            return EncryptionResult(
                file_path=encrypted_path,
                mode='aes256gcm',
                success=False,
                error='密码错误或文件已损坏',
            )
        return EncryptionResult(
            file_path=encrypted_path,
            mode='aes256gcm',
            success=False,
            error=str(e),
        )


def encrypt_excel_with_password(file_path: str, password: str, output_path: Optional[str] = None) -> EncryptionResult:
    logger = get_logger()

    if not HAS_MSOFFCRYPTO:
        logger.warning('msoffcrypto-python 库未安装，Excel密码保护不可用，可运行 pip install msoffcrypto-python 安装')
        return EncryptionResult(
            file_path=file_path,
            mode='excel_password',
            success=False,
            error='msoffcrypto-python 库未安装，Excel密码保护不可用',
        )

    if not os.path.isfile(file_path):
        return EncryptionResult(
            file_path=file_path,
            mode='excel_password',
            success=False,
            error=f'文件不存在: {file_path}',
        )

    if not file_path.lower().endswith(('.xlsx', '.xls', '.xlsm')):
        return EncryptionResult(
            file_path=file_path,
            mode='excel_password',
            success=False,
            error='仅支持 Excel 文件 (.xlsx/.xls/.xlsm) 的密码保护',
        )

    valid, msg = validate_password_strength(password)
    if not valid:
        return EncryptionResult(
            file_path=file_path,
            mode='excel_password',
            success=False,
            error=msg,
        )

    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f'{base}_加密版{ext}'

    try:
        original_size = os.path.getsize(file_path)
        file_hash = _compute_file_hash(file_path)

        with open(file_path, 'rb') as f:
            file_io = msoffcrypto.OfficeFile(f)
            file_io.load_key(password=password)

            with open(output_path, 'wb') as out_f:
                file_io.encrypt(password=password, outfile=out_f)

        encrypted_size = os.path.getsize(output_path)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        logger.info('Excel密码保护完成: %s -> %s', file_path, output_path)

        return EncryptionResult(
            file_path=file_path,
            encrypted_path=output_path,
            mode='excel_password',
            success=True,
            original_size=original_size,
            encrypted_size=encrypted_size,
            file_hash=file_hash,
            timestamp=timestamp,
        )

    except Exception as e:
        logger.error('Excel密码保护失败 %s: %s', file_path, e, exc_info=True)
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        return EncryptionResult(
            file_path=file_path,
            mode='excel_password',
            success=False,
            error=str(e),
        )


def encrypt_output_files(
    file_paths: List[str],
    password: str,
    mode: str = 'excel_password',
    output_dir: Optional[str] = None,
) -> BatchEncryptionResult:
    logger = get_logger()

    valid, msg = validate_password_strength(password)
    if not valid:
        return BatchEncryptionResult(
            total_files=len(file_paths),
            failure_count=len(file_paths),
            mode=mode,
        )

    results = []
    success_count = 0
    failure_count = 0

    for file_path in file_paths:
        if not os.path.isfile(file_path):
            logger.warning('文件不存在，跳过加密: %s', file_path)
            results.append(EncryptionResult(
                file_path=file_path,
                mode=mode,
                success=False,
                error='文件不存在',
            ))
            failure_count += 1
            continue

        output_path = None
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            basename = os.path.basename(file_path)
            if mode == 'aes256gcm':
                output_path = os.path.join(output_dir, basename + ENCRYPTED_EXTENSION)
            elif mode == 'excel_password':
                base, ext = os.path.splitext(basename)
                output_path = os.path.join(output_dir, f'{base}_加密版{ext}')

        if mode == 'excel_password':
            is_excel = file_path.lower().endswith(('.xlsx', '.xls', '.xlsm'))
            if is_excel:
                result = encrypt_excel_with_password(file_path, password, output_path)
                if not result.success and HAS_CRYPTOGRAPHY:
                    logger.info('Excel密码保护不可用，自动降级为AES加密: %s', file_path)
                    result = encrypt_file_aes(file_path, password, output_path)
            else:
                if HAS_CRYPTOGRAPHY:
                    logger.info('非Excel文件，使用AES加密: %s', file_path)
                    result = encrypt_file_aes(file_path, password, output_path)
                else:
                    result = EncryptionResult(
                        file_path=file_path,
                        mode='excel_password',
                        success=False,
                        error='非Excel文件且cryptography库不可用',
                    )
        elif mode == 'aes256gcm':
            result = encrypt_file_aes(file_path, password, output_path)
        else:
            result = EncryptionResult(
                file_path=file_path,
                mode=mode,
                success=False,
                error=f'不支持的加密模式: {mode}',
            )

        results.append(result)
        if result.success:
            success_count += 1
        else:
            failure_count += 1

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    logger.info('批量加密完成: 模式=%s, 总数=%d, 成功=%d, 失败=%d',
                mode, len(file_paths), success_count, failure_count)

    return BatchEncryptionResult(
        results=results,
        total_files=len(file_paths),
        success_count=success_count,
        failure_count=failure_count,
        mode=mode,
        timestamp=timestamp,
    )


def is_encrypted_file(file_path: str) -> bool:
    if not os.path.isfile(file_path):
        return False
    try:
        with open(file_path, 'rb') as f:
            header = f.read(len(ENCRYPTION_MARKER))
            return header == ENCRYPTION_MARKER
    except Exception:
        return False


def get_encryption_info(file_path: str) -> Dict[str, Any]:
    logger = get_logger()

    if not os.path.isfile(file_path):
        return {'exists': False, 'encrypted': False}

    result = {
        'exists': True,
        'file_path': file_path,
        'file_size': os.path.getsize(file_path),
        'encrypted': False,
        'encryption_mode': None,
        'file_hash': _compute_file_hash(file_path),
    }

    if is_encrypted_file(file_path):
        result['encrypted'] = True
        result['encryption_mode'] = 'aes256gcm'
        return result

    if file_path.lower().endswith(('.xlsx', '.xls', '.xlsm')):
        if HAS_MSOFFCRYPTO:
            try:
                with open(file_path, 'rb') as f:
                    ofile = msoffcrypto.OfficeFile(f)
                    result['encrypted'] = ofile.is_encrypted()
                    if ofile.is_encrypted():
                        result['encryption_mode'] = 'excel_password'
            except Exception as e:
                logger.debug('检测Excel加密状态失败: %s', e)

    return result


def save_encryption_record(
    batch_result: BatchEncryptionResult,
    script_dir: Optional[str] = None,
) -> Optional[str]:
    logger = get_logger()

    if not batch_result.results:
        return None

    if script_dir is None:
        script_dir = get_script_dir()

    log_dir = os.path.join(script_dir, 'encryption_logs')
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'encryption_{timestamp}.json')

    try:
        log_data = {
            'timestamp': batch_result.timestamp,
            'mode': batch_result.mode,
            'total_files': batch_result.total_files,
            'success_count': batch_result.success_count,
            'failure_count': batch_result.failure_count,
            'results': [r.to_dict() for r in batch_result.results],
        }

        with open(log_file, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)

        logger.info('加密日志已保存: %s', log_file)
        return log_file

    except Exception as e:
        logger.error('保存加密日志失败: %s', e, exc_info=True)
        return None
