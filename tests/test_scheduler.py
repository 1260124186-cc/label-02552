import os
import json
import sqlite3
import tempfile
import shutil
from unittest.mock import patch, MagicMock

import pytest

from bankcheck import (
    SCHEDULER_CONFIG_FILENAME,
    PROCESSED_FILES_DB_FILENAME,
    init_processed_files_db,
    get_processed_files_db_path,
    load_scheduler_config,
    save_scheduler_config,
    add_schedule_job,
    update_schedule_job,
    remove_schedule_job,
    list_schedule_jobs,
    is_file_processed,
    mark_file_processed,
    scan_new_files,
    record_scheduler_run,
    get_scheduler_config_path,
    SimpleScheduler,
)


@pytest.fixture
def script_dir(tmp_path):
    return str(tmp_path)


@pytest.fixture
def clean_config(script_dir):
    config_path = get_scheduler_config_path(script_dir)
    if os.path.exists(config_path):
        os.remove(config_path)
    db_path = get_processed_files_db_path(script_dir)
    if os.path.exists(db_path):
        os.remove(db_path)
    yield script_dir
    if os.path.exists(config_path):
        os.remove(config_path)
    if os.path.exists(db_path):
        os.remove(db_path)


class TestSchedulerConfig:
    def test_load_default_config(self, clean_config):
        config = load_scheduler_config(clean_config)
        assert 'jobs' in config
        assert 'settings' in config
        assert config['settings']['max_concurrent_jobs'] == 1
        assert config['settings']['enable_alerts'] is True

    def test_save_and_load_config(self, clean_config):
        test_config = {
            'jobs': [{'job_id': 'test1', 'name': 'test'}],
            'settings': {'max_concurrent_jobs': 2}
        }
        result = save_scheduler_config(test_config, clean_config)
        assert result is True

        loaded = load_scheduler_config(clean_config)
        assert len(loaded['jobs']) == 1
        assert loaded['jobs'][0]['job_id'] == 'test1'
        assert loaded['settings']['max_concurrent_jobs'] == 2

    def test_add_job(self, clean_config):
        job_config = {
            'name': '测试任务',
            'watch_directory': '/tmp/test',
            'cron_expression': '0 0 * * *',
            'schedule_type': 'cron',
            'incremental': True,
        }
        job_id = add_schedule_job(job_config, clean_config)
        assert job_id is not None
        assert job_id.startswith('JOB')

        jobs = list_schedule_jobs(clean_config)
        assert len(jobs) == 1
        assert jobs[0]['name'] == '测试任务'
        assert jobs[0]['watch_directory'] == '/tmp/test'

    def test_update_job(self, clean_config):
        job_config = {
            'name': '原任务',
            'watch_directory': '/tmp/original',
        }
        job_id = add_schedule_job(job_config, clean_config)

        result = update_schedule_job(job_id, {'name': '更新后的任务'}, clean_config)
        assert result is True

        jobs = list_schedule_jobs(clean_config)
        assert jobs[0]['name'] == '更新后的任务'

    def test_update_nonexistent_job(self, clean_config):
        result = update_schedule_job('NONEXISTENT', {'name': 'test'}, clean_config)
        assert result is False

    def test_remove_job(self, clean_config):
        job_config = {'name': '待删除', 'watch_directory': '/tmp/test'}
        job_id = add_schedule_job(job_config, clean_config)

        remove_schedule_job(job_id, clean_config)
        jobs = list_schedule_jobs(clean_config)
        assert len(jobs) == 0

    def test_list_jobs_empty(self, clean_config):
        jobs = list_schedule_jobs(clean_config)
        assert jobs == []


class TestProcessedFilesDB:
    def test_init_db(self, script_dir):
        db_path = get_processed_files_db_path(script_dir)
        init_processed_files_db(db_path)

        assert os.path.exists(db_path)

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in cursor.fetchall()]
        conn.close()

        assert 'processed_files' in tables
        assert 'scheduler_runs' in tables

    def test_is_file_processed_false(self, script_dir):
        db_path = get_processed_files_db_path(script_dir)
        init_processed_files_db(db_path)

        test_file = os.path.join(script_dir, 'test.xlsx')
        with open(test_file, 'w') as f:
            f.write('test content')

        processed, file_hash = is_file_processed(test_file, 'JOB001', db_path)
        assert processed is False
        assert file_hash is not None

        os.remove(test_file)

    def test_mark_and_check_file(self, script_dir):
        db_path = get_processed_files_db_path(script_dir)
        init_processed_files_db(db_path)

        test_file = os.path.join(script_dir, 'test.xlsx')
        with open(test_file, 'w') as f:
            f.write('test content')

        _, file_hash = is_file_processed(test_file, 'JOB001', db_path)
        mark_file_processed(test_file, file_hash, 'JOB001', 10, 'success', db_path)

        processed, _ = is_file_processed(test_file, 'JOB001', db_path)
        assert processed is True

        os.remove(test_file)

    def test_scan_new_files(self, script_dir):
        watch_dir = os.path.join(script_dir, 'watch')
        os.makedirs(watch_dir)

        file1 = os.path.join(watch_dir, '北京银行_test1.xlsx')
        file2 = os.path.join(watch_dir, '东亚银行_test2.xlsx')
        with open(file1, 'w') as f:
            f.write('test1')
        with open(file2, 'w') as f:
            f.write('test2')

        db_path = get_processed_files_db_path(script_dir)
        init_processed_files_db(db_path)

        new_files = scan_new_files(watch_dir, 'JOB001', script_dir)
        assert len(new_files) == 2

        _, file_hash = is_file_processed(file1, 'JOB001', db_path)
        mark_file_processed(file1, file_hash, 'JOB001', 5, 'success', db_path)

        new_files = scan_new_files(watch_dir, 'JOB001', script_dir)
        assert len(new_files) == 1

        shutil.rmtree(watch_dir)

    def test_record_scheduler_run(self, script_dir):
        db_path = get_processed_files_db_path(script_dir)
        init_processed_files_db(db_path)

        run_data = {
            'run_id': 'SCH20240101000000TEST',
            'job_id': 'JOB001',
            'job_name': '测试任务',
            'started_at': '2024-01-01 00:00:00',
            'completed_at': '2024-01-01 00:05:00',
            'status': 'success',
            'files_scanned': 10,
            'files_new': 3,
            'files_processed': 3,
            'files_skipped': 0,
            'files_error': 0,
            'records_extracted': 150,
            'output_path': '/tmp/output.xlsx',
            'duration_ms': 300000,
        }

        record_scheduler_run(run_data, script_dir)

        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM scheduler_runs WHERE run_id = ?', (run_data['run_id'],))
        row = cursor.fetchone()
        conn.close()

        assert row is not None
        assert row['job_id'] == 'JOB001'
        assert row['status'] == 'success'
        assert row['files_new'] == 3
        assert row['records_extracted'] == 150


class TestSimpleScheduler:
    def test_scheduler_initialization(self):
        scheduler = SimpleScheduler()
        assert scheduler.running is False
        assert scheduler.jobs == []

    def test_add_job(self):
        scheduler = SimpleScheduler()
        mock_func = MagicMock()

        class MockTrigger:
            def get_interval(self):
                from datetime import timedelta
                return timedelta(minutes=60)

        job = scheduler.add_job(
            mock_func,
            trigger=MockTrigger(),
            args=['arg1'],
            kwargs={'key': 'value'},
            id='TESTJOB',
            name='Test Job'
        )

        assert job['id'] == 'TESTJOB'
        assert job['name'] == 'Test Job'
        assert len(scheduler.jobs) == 1

    def test_should_run_first_time(self):
        scheduler = SimpleScheduler()

        class MockTrigger:
            def get_interval(self):
                from datetime import timedelta
                return timedelta(minutes=60)

        job = {
            'id': 'TEST',
            'trigger': MockTrigger(),
            'last_run': None,
        }

        from datetime import datetime
        now = datetime.now()
        assert scheduler._should_run(job, now) is True

    def test_should_run_interval(self):
        scheduler = SimpleScheduler()

        class MockTrigger:
            def get_interval(self):
                from datetime import timedelta
                return timedelta(minutes=60)

        from datetime import datetime, timedelta
        now = datetime.now()

        job = {
            'id': 'TEST',
            'trigger': MockTrigger(),
            'last_run': now - timedelta(minutes=30),
        }
        assert scheduler._should_run(job, now) is False

        job['last_run'] = now - timedelta(minutes=70)
        assert scheduler._should_run(job, now) is True

    def test_start_and_shutdown(self):
        scheduler = SimpleScheduler()
        import time

        scheduler.start()
        time.sleep(0.1)
        assert scheduler.running is True

        scheduler.shutdown()
        assert scheduler.running is False


class TestCronAndWindowsScripts:
    def test_generate_cron_script(self, script_dir):
        from bankcheck import generate_cron_script

        job_config = {
            'job_id': 'JOB001',
            'name': '每日汇总',
            'watch_directory': '/data/transactions',
            'cron_expression': '0 2 * * *',
        }

        script_path = os.path.join(script_dir, 'bankcheck.py')
        with open(script_path, 'w') as f:
            f.write('# dummy')

        output_path, cron_line = generate_cron_script(job_config, script_path, script_dir)

        assert os.path.exists(output_path)
        assert 'JOB001' in cron_line
        assert '0 2 * * *' in cron_line
        assert '--run-job' in cron_line
        assert os.access(output_path, os.X_OK)

    def test_generate_windows_task_script(self, script_dir):
        from bankcheck import generate_windows_task_script

        job_config = {
            'job_id': 'JOB001',
            'name': '每日汇总',
            'watch_directory': 'C:\\data\\transactions',
            'cron_expression': '0 2 * * *',
        }

        script_path = os.path.join(script_dir, 'bankcheck.py')
        with open(script_path, 'w') as f:
            f.write('# dummy')

        ps_output, bat_output = generate_windows_task_script(job_config, script_path, script_dir)

        assert os.path.exists(ps_output)
        assert os.path.exists(bat_output)
        assert 'JOB001' in open(ps_output, encoding='utf-8').read()
        assert 'SchTasks' in open(bat_output, encoding='gbk').read()


class TestScheduledPipeline:
    def test_run_scheduled_pipeline_no_new_files(self, script_dir, tmp_path):
        from bankcheck import run_scheduled_pipeline

        watch_dir = tmp_path / 'watch'
        watch_dir.mkdir()

        job_config = {
            'job_id': 'TESTJOB',
            'name': '测试任务',
            'watch_directory': str(watch_dir),
            'incremental': True,
        }

        with patch('bankcheck.detect_and_record_lookup_change') as mock_detect, \
             patch('bankcheck.scan_new_files') as mock_scan, \
             patch('bankcheck.run_pipeline') as mock_run:

            mock_detect.return_value = MagicMock(has_changes=False)
            mock_scan.return_value = []

            result = run_scheduled_pipeline(job_config, script_dir)

            assert result['status'] == 'success'
            assert result['files_new'] == 0
            mock_run.assert_not_called()

    def test_run_scheduled_pipeline_with_new_files(self, script_dir, tmp_path):
        from bankcheck import run_scheduled_pipeline
        from bankcheck import ProcessingResult

        watch_dir = tmp_path / 'watch'
        watch_dir.mkdir()

        test_file = watch_dir / '北京银行_test.xlsx'
        test_file.write_text('test')

        job_config = {
            'job_id': 'TESTJOB',
            'name': '测试任务',
            'watch_directory': str(watch_dir),
            'incremental': True,
        }

        mock_result = ProcessingResult(
            all_rows=[{'唯一id': '1'}],
            processed_files=[str(test_file)],
            new_record_count=1,
        )

        with patch('bankcheck.detect_and_record_lookup_change') as mock_detect, \
             patch('bankcheck.AuditLogger') as mock_audit, \
             patch('bankcheck.run_pipeline', return_value=mock_result) as mock_run, \
             patch('bankcheck.mark_file_processed') as mock_mark:

            mock_detect.return_value = MagicMock(has_changes=False)
            mock_audit.return_value.__enter__.return_value = MagicMock()

            result = run_scheduled_pipeline(job_config, script_dir)

            assert result['status'] == 'success'
            assert result['files_new'] > 0
            mock_run.assert_called_once()
            mock_mark.assert_called()


class TestConfigFileCreation:
    def test_config_file_created_on_first_load(self, clean_config):
        config_path = get_scheduler_config_path(clean_config)
        assert not os.path.exists(config_path)

        load_scheduler_config(clean_config)

        assert os.path.exists(config_path)
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        assert 'jobs' in config
        assert 'settings' in config
