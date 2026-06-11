# -*- coding: utf-8 -*-
"""
主体查找表 Web 管理应用
提供 RESTful API 和 Web 界面，用于管理银行账号与主体的映射关系
"""

import os
import sys
import logging
import tempfile
from datetime import datetime

from flask import (Flask, render_template, request, jsonify,
                   send_file, redirect, url_for, flash)

import lookup_manager as lm

try:
    from i18n import t, set_language, get_language, get_available_languages, init_i18n
    HAS_I18N = True
except ImportError:
    HAS_I18N = False
    def t(key, **kwargs):
        return key
    def set_language(lang):
        return False
    def get_language():
        return 'zh_CN'
    def get_available_languages():
        return {'zh_CN': '简体中文'}
    def init_i18n(lang=None):
        return None

if HAS_I18N:
    init_i18n()


def setup_logging():
    log_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(log_dir, 'lookup_manager.log')

    logger = logging.getLogger('bankcheck')
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)

        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger


logger = setup_logging()

app = Flask(__name__,
            template_folder=os.path.join(os.path.dirname(__file__), 'templates'),
            static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.secret_key = 'bankcheck_lookup_manager_secret_key_2024'


@app.context_processor
def inject_i18n():
    """注入多语言函数到模板上下文"""
    return {
        't': t,
        'current_language': get_language(),
        'available_languages': get_available_languages(),
    }


@app.route('/')
def index():
    """主页 - 查找表管理界面"""
    keyword = request.args.get('keyword', '').strip()
    entries = lm.search_entries(keyword)
    lookup_file = lm.get_lookup_file_path()
    duplicates = lm.get_duplicate_entries()

    stats = {
        'total': len(entries),
        'subjects': len(set(e.subject for e in entries)),
        'duplicates': len(duplicates)
    }

    return render_template('index.html',
                           entries=entries,
                           keyword=keyword,
                           lookup_file=lookup_file,
                           stats=stats,
                           duplicates=duplicates)


@app.route('/api/entries', methods=['GET'])
def api_get_entries():
    """获取所有条目 API"""
    keyword = request.args.get('keyword', '').strip()
    entries = lm.search_entries(keyword)
    return jsonify({
        'success': True,
        'data': [e.to_dict() for e in entries],
        'total': len(entries)
    })


@app.route('/api/entries/<account>', methods=['GET'])
def api_get_entry(account):
    """获取单个条目 API"""
    entry = lm.get_entry_by_account(account)
    if entry:
        return jsonify({
            'success': True,
            'data': entry.to_dict()
        })
    return jsonify({
        'success': False,
        'message': t('api.entry_not_found')
    }), 404


@app.route('/api/entries', methods=['POST'])
def api_add_entry():
    """添加条目 API"""
    data = request.get_json() or request.form
    subject = data.get('subject', '').strip()
    account = data.get('account', '').strip()

    success, message = lm.add_entry(subject, account)
    if success:
        return jsonify({
            'success': True,
            'message': message
        }), 201
    return jsonify({
        'success': False,
        'message': message
    }), 400


@app.route('/api/entries/<old_account>', methods=['PUT'])
def api_update_entry(old_account):
    """更新条目 API"""
    data = request.get_json() or request.form
    new_subject = data.get('subject', '').strip()
    new_account = data.get('account', '').strip()

    success, message = lm.update_entry(old_account, new_subject, new_account)
    if success:
        return jsonify({
            'success': True,
            'message': message
        })
    return jsonify({
        'success': False,
        'message': message
    }), 400


@app.route('/api/entries/<account>', methods=['DELETE'])
def api_delete_entry(account):
    """删除条目 API"""
    success, message = lm.delete_entry(account)
    if success:
        return jsonify({
            'success': True,
            'message': message
        })
    return jsonify({
        'success': False,
        'message': message
    }), 400


@app.route('/api/export', methods=['GET'])
def api_export():
    """导出 Excel API"""
    try:
        tmp_dir = tempfile.gettempdir()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        export_filename = f'主体查找表_导出_{timestamp}.xlsx'
        export_path = os.path.join(tmp_dir, export_filename)

        success, message = lm.export_to_excel(export_path)
        if success:
            return send_file(export_path,
                           as_attachment=True,
                           download_name=export_filename,
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        return jsonify({
            'success': False,
            'message': message
        }), 500
    except Exception as e:
        logger.error('导出失败: %s', e, exc_info=True)
        return jsonify({
            'success': False,
            'message': t('api.export_failed', error=str(e))
        }), 500


@app.route('/api/import', methods=['POST'])
def api_import():
    """导入 Excel API"""
    try:
        if 'file' not in request.files:
            return jsonify({
                'success': False,
                'message': t('api.please_select_file')
            }), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({
                'success': False,
                'message': t('api.please_select_file')
            }), 400

        overwrite = request.form.get('overwrite', 'false').lower() == 'true'

        tmp_dir = tempfile.gettempdir()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        import_filename = f'import_tmp_{timestamp}_{file.filename}'
        import_path = os.path.join(tmp_dir, import_filename)
        file.save(import_path)

        success, message, stats = lm.import_from_excel(import_path, overwrite=overwrite)

        try:
            os.remove(import_path)
        except Exception:
            pass

        if success:
            return jsonify({
                'success': True,
                'message': message,
                'stats': stats
            })
        return jsonify({
            'success': False,
            'message': message,
            'stats': stats
        }), 400
    except Exception as e:
        logger.error('导入失败: %s', e, exc_info=True)
        return jsonify({
            'success': False,
            'message': t('api.import_failed', error=str(e))
        }), 500


@app.route('/api/duplicates', methods=['GET'])
def api_get_duplicates():
    """获取重复条目 API"""
    duplicates = lm.get_duplicate_entries()
    return jsonify({
        'success': True,
        'data': duplicates,
        'total': len(duplicates)
    })


@app.route('/add', methods=['POST'])
def web_add_entry():
    """Web 表单 - 添加条目"""
    subject = request.form.get('subject', '').strip()
    account = request.form.get('account', '').strip()

    success, message = lm.add_entry(subject, account)
    flash(message, 'success' if success else 'error')

    return redirect(url_for('index'))


@app.route('/update/<old_account>', methods=['POST'])
def web_update_entry(old_account):
    """Web 表单 - 更新条目"""
    new_subject = request.form.get('subject', '').strip()
    new_account = request.form.get('account', '').strip()

    success, message = lm.update_entry(old_account, new_subject, new_account)
    flash(message, 'success' if success else 'error')

    return redirect(url_for('index'))


@app.route('/delete/<account>', methods=['POST'])
def web_delete_entry(account):
    """Web 表单 - 删除条目"""
    success, message = lm.delete_entry(account)
    flash(message, 'success' if success else 'error')

    return redirect(url_for('index'))


@app.route('/import', methods=['POST'])
def web_import():
    """Web 表单 - 导入 Excel"""
    try:
        if 'file' not in request.files:
            flash('请选择要导入的文件', 'error')
            return redirect(url_for('index'))

        file = request.files['file']
        if file.filename == '':
            flash('请选择要导入的文件', 'error')
            return redirect(url_for('index'))

        overwrite = request.form.get('overwrite') == 'on'

        tmp_dir = tempfile.gettempdir()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        import_filename = f'import_tmp_{timestamp}_{file.filename}'
        import_path = os.path.join(tmp_dir, import_filename)
        file.save(import_path)

        success, message, stats = lm.import_from_excel(import_path, overwrite=overwrite)

        try:
            os.remove(import_path)
        except Exception:
            pass

        flash(message, 'success' if success else 'error')
        return redirect(url_for('index'))
    except Exception as e:
        logger.error('导入失败: %s', e, exc_info=True)
        flash(f'导入失败: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/api/language', methods=['GET'])
def api_get_language():
    """获取当前语言和可用语言"""
    return jsonify({
        'success': True,
        'current_language': get_language(),
        'available_languages': get_available_languages()
    })


@app.route('/api/language/<lang>', methods=['POST'])
def api_set_language(lang):
    """设置语言"""
    success = set_language(lang)
    if success:
        return jsonify({
            'success': True,
            'message': t('api.language_changed'),
            'language': lang
        })
    return jsonify({
        'success': False,
        'message': t('api.unsupported_language')
    }), 400


@app.errorhandler(404)
def not_found(error):
    return jsonify({
        'success': False,
        'message': t('api.endpoint_not_found')
    }), 404


@app.errorhandler(500)
def internal_error(error):
    logger.error('服务器内部错误: %s', error, exc_info=True)
    return jsonify({
        'success': False,
        'message': t('api.internal_error')
    }), 500


def main():
    host = '127.0.0.1'
    port = 5000
    debug = True

    logger.info('=' * 60)
    logger.info(t('web_ui.app_started'))
    logger.info(t('web_ui.access_url', host=host, port=port))
    logger.info(t('web_ui.lookup_file_path', path=lm.get_lookup_file_path()))
    logger.info('=' * 60)

    app.run(host=host, port=port, debug=debug)


if __name__ == '__main__':
    main()
