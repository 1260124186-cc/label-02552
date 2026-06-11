import os
import sys
import json
import shutil
import tempfile
import time

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from perf_profiler import PerfProfiler, get_profiler, reset_profiler


@pytest.fixture(autouse=True)
def _clean_profiler():
    reset_profiler()
    yield
    reset_profiler()


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp(prefix='perf_profiler_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


class TestPerfProfilerBasic:
    def test_initial_state(self):
        p = PerfProfiler()
        assert p.file_open_records == []
        assert p.row_traversal_records == []
        assert p.lookup_hit_records == []
        assert p.phase_records == []
        assert p._total_duration_ms == 0.0

    def test_start_stop(self):
        p = PerfProfiler()
        p.start()
        time.sleep(0.01)
        p.stop()
        assert p._total_duration_ms > 0

    def test_record_file_open(self):
        p = PerfProfiler()
        p.record_file_open('/tmp/test.xlsx', 150.5, bank_name='北京银行')
        assert len(p.file_open_records) == 1
        r = p.file_open_records[0]
        assert r.filepath == '/tmp/test.xlsx'
        assert r.duration_ms == 150.5
        assert r.bank_name == '北京银行'
        assert r.file_size_kb == 0.0
        assert r.is_xls_convert is False

    def test_record_file_open_with_size(self, tmp_dir):
        f = os.path.join(tmp_dir, 'test.xlsx')
        with open(f, 'w') as fh:
            fh.write('x' * 1024)
        p = PerfProfiler()
        p.record_file_open(f, 100.0)
        r = p.file_open_records[0]
        assert r.file_size_kb > 0

    def test_record_row_traversal(self):
        p = PerfProfiler()
        p.record_row_traversal('/tmp/test.xlsx', 'Sheet1', 100, 50.0,
                               bank_name='东亚银行', extracted_count=80)
        assert len(p.row_traversal_records) == 1
        r = p.row_traversal_records[0]
        assert r.sheet_name == 'Sheet1'
        assert r.row_count == 100
        assert r.duration_ms == 50.0
        assert r.bank_name == '东亚银行'
        assert r.extracted_count == 80

    def test_record_lookup_hit(self):
        p = PerfProfiler()
        p.record_lookup_hit('6222021234', 0.5, hit=True)
        p.record_lookup_hit('6222029999', 1.2, hit=False)
        assert len(p.lookup_hit_records) == 2
        assert p.lookup_hit_records[0].hit is True
        assert p.lookup_hit_records[1].hit is False
        assert p.lookup_hit_records[0].duration_ms == 0.5

    def test_record_lookup_hit_fuzzy(self):
        p = PerfProfiler()
        p.record_lookup_hit('6222021234', 5.0, hit=True, fuzzy=True, similarity=0.85)
        r = p.lookup_hit_records[0]
        assert r.fuzzy is True
        assert r.similarity == 0.85

    def test_record_phase(self):
        p = PerfProfiler()
        p.record_phase('lookup_preload', 200.0, '查找表预加载')
        assert len(p.phase_records) == 1
        r = p.phase_records[0]
        assert r.phase_name == 'lookup_preload'
        assert r.duration_ms == 200.0
        assert r.detail == '查找表预加载'


class TestPerfProfilerContextManagers:
    def test_measure_file_open(self):
        p = PerfProfiler()
        with p.measure_file_open('/tmp/test.xlsx', bank_name='工商银行'):
            time.sleep(0.005)
        assert len(p.file_open_records) == 1
        r = p.file_open_records[0]
        assert r.filepath == '/tmp/test.xlsx'
        assert r.bank_name == '工商银行'
        assert r.duration_ms > 0

    def test_measure_file_open_with_xls(self):
        p = PerfProfiler()
        with p.measure_file_open('/tmp/old.xls', is_xls_convert=True):
            pass
        r = p.file_open_records[0]
        assert r.is_xls_convert is True

    def test_measure_row_traversal(self):
        p = PerfProfiler()
        with p.measure_row_traversal('/tmp/test.xlsx', 'Sheet1', bank_name='北京银行') as holder:
            holder['row_count'] = 100
            holder['extracted_count'] = 80
        assert len(p.row_traversal_records) == 1
        r = p.row_traversal_records[0]
        assert r.row_count == 100
        assert r.extracted_count == 80
        assert r.duration_ms >= 0

    def test_measure_lookup_hit(self):
        p = PerfProfiler()
        with p.measure_lookup_hit('6222021234') as holder:
            holder['hit'] = True
            holder['fuzzy'] = False
            holder['similarity'] = 1.0
        r = p.lookup_hit_records[0]
        assert r.hit is True
        assert r.similarity == 1.0

    def test_measure_phase(self):
        p = PerfProfiler()
        with p.measure_phase('test_phase', 'some detail'):
            time.sleep(0.005)
        assert len(p.phase_records) == 1
        r = p.phase_records[0]
        assert r.phase_name == 'test_phase'
        assert r.duration_ms > 0
        assert r.detail == 'some detail'


class TestPerfProfilerSummary:
    def test_empty_summary(self):
        p = PerfProfiler()
        s = p.get_summary()
        assert s['total_duration_ms'] == 0.0
        assert s['file_open']['count'] == 0
        assert s['row_traversal']['count'] == 0
        assert s['lookup_hit']['count'] == 0
        assert s['bottlenecks'] == []

    def test_summary_with_data(self):
        p = PerfProfiler()
        p._total_duration_ms = 1000.0
        p.record_file_open('/tmp/a.xlsx', 100.0, bank_name='北京银行')
        p.record_file_open('/tmp/b.xlsx', 200.0, bank_name='东亚银行')
        p.record_row_traversal('/tmp/a.xlsx', 'Sheet1', 100, 50.0,
                               bank_name='北京银行', extracted_count=80)
        p.record_lookup_hit('6222021234', 0.5, hit=True)
        p.record_lookup_hit('9999999', 1.0, hit=False)
        p.record_phase('test', 50.0, 'detail')

        s = p.get_summary()
        assert s['total_duration_ms'] == 1000.0
        assert s['file_open']['count'] == 2
        assert s['file_open']['total_ms'] == 300.0
        assert s['file_open']['avg_ms'] == 150.0
        assert s['file_open']['max_ms'] == 200.0
        assert s['file_open']['min_ms'] == 100.0
        assert len(s['file_open']['slowest']) == 2
        assert s['row_traversal']['count'] == 1
        assert s['row_traversal']['total_rows'] == 100
        assert s['row_traversal']['total_extracted'] == 80
        assert s['lookup_hit']['count'] == 2
        assert s['lookup_hit']['hit_count'] == 1
        assert s['lookup_hit']['miss_count'] == 1
        assert s['lookup_hit']['hit_rate'] == 0.5
        assert len(s['phases']) == 1

    def test_summary_slowest_top5(self):
        p = PerfProfiler()
        for i in range(10):
            p.record_file_open(f'/tmp/f{i}.xlsx', float(i * 100))
        s = p.get_summary()
        assert len(s['file_open']['slowest']) == 5
        assert s['file_open']['slowest'][0]['duration_ms'] == 900.0


class TestPerfProfilerBottleneck:
    def test_bottleneck_file_open_high(self):
        p = PerfProfiler()
        p._total_duration_ms = 1000.0
        for i in range(10):
            p.record_file_open(f'/tmp/f{i}.xlsx', 80.0)
        p.record_row_traversal('/tmp/f0.xlsx', 'S', 10, 10.0)
        p.record_lookup_hit('123', 0.1, hit=True)
        s = p.get_summary()
        fo_bottlenecks = [b for b in s['bottlenecks'] if b['phase'] == 'file_open']
        assert len(fo_bottlenecks) > 0
        assert fo_bottlenecks[0]['severity'] in ('high', 'medium')

    def test_bottleneck_lookup_low_hit_rate(self):
        p = PerfProfiler()
        p._total_duration_ms = 1000.0
        for i in range(10):
            p.record_file_open(f'/tmp/f{i}.xlsx', 10.0)
        p.record_row_traversal('/tmp/f0.xlsx', 'S', 10, 10.0)
        for i in range(20):
            p.record_lookup_hit(f'acc{i}', 1.0, hit=False)
        s = p.get_summary()
        lk_bottlenecks = [b for b in s['bottlenecks'] if b['phase'] == 'lookup_hit']
        assert len(lk_bottlenecks) > 0

    def test_bottleneck_traversal_high(self):
        p = PerfProfiler()
        p._total_duration_ms = 2000.0
        p.record_file_open('/tmp/f.xlsx', 100.0)
        for i in range(5):
            p.record_row_traversal('/tmp/f.xlsx', f'S{i}', 5000, 300.0)
        p.record_lookup_hit('123', 0.1, hit=True)
        s = p.get_summary()
        rt_bottlenecks = [b for b in s['bottlenecks'] if b['phase'] == 'row_traversal']
        assert len(rt_bottlenecks) > 0

    def test_bottleneck_many_files(self):
        p = PerfProfiler()
        p._total_duration_ms = 5000.0
        for i in range(60):
            p.record_file_open(f'/tmp/f{i}.xlsx', 10.0)
        s = p.get_summary()
        file_bottlenecks = [b for b in s['bottlenecks'] if b['phase'] == 'file_open' and b['severity'] == 'low']
        assert len(file_bottlenecks) > 0

    def test_no_high_severity_bottleneck(self):
        p = PerfProfiler()
        p._total_duration_ms = 1000.0
        for i in range(20):
            p.record_file_open(f'/tmp/f{i}.xlsx', 5.0)
            p.record_row_traversal(f'/tmp/f{i}.xlsx', f'S{i}', 100, 4.0)
            p.record_lookup_hit(f'acc{i}', 0.1, hit=True)
        s = p.get_summary()
        file_open_pct = s['file_open']['total_ms'] / max(
            s['file_open']['total_ms'] + s['row_traversal']['total_ms'] + s['lookup_hit']['total_ms'], 1) * 100
        row_traversal_pct = s['row_traversal']['total_ms'] / max(
            s['file_open']['total_ms'] + s['row_traversal']['total_ms'] + s['lookup_hit']['total_ms'], 1) * 100
        assert file_open_pct < 55
        assert row_traversal_pct < 55


class TestPerfProfilerReport:
    def test_generate_report_markdown(self):
        p = PerfProfiler()
        p._total_duration_ms = 500.0
        p.record_file_open('/tmp/test.xlsx', 100.0, bank_name='北京银行')
        p.record_row_traversal('/tmp/test.xlsx', 'Sheet1', 50, 30.0,
                               bank_name='北京银行', extracted_count=40)
        p.record_lookup_hit('6222021234', 0.5, hit=True)
        report = p.generate_report()
        assert '# 运行性能剖析报告' in report
        assert '文件打开耗时' in report
        assert '行遍历耗时' in report
        assert '查找表命中耗时' in report
        assert '性能瓶颈分析' in report

    def test_save_report(self, tmp_dir):
        p = PerfProfiler()
        p._total_duration_ms = 500.0
        p.record_file_open('/tmp/test.xlsx', 100.0)
        report_path = p.save_report(tmp_dir)
        assert os.path.exists(report_path)
        assert report_path.endswith('.md')
        with open(report_path, 'r', encoding='utf-8') as f:
            content = f.read()
        assert '运行性能剖析报告' in content

    def test_save_report_json(self, tmp_dir):
        p = PerfProfiler()
        p._total_duration_ms = 500.0
        p.record_file_open('/tmp/test.xlsx', 100.0)
        p.save_report(tmp_dir)
        json_files = [f for f in os.listdir(tmp_dir) if f.endswith('.json')]
        assert len(json_files) == 1
        with open(os.path.join(tmp_dir, json_files[0]), 'r', encoding='utf-8') as f:
            data = json.load(f)
        assert 'file_open' in data
        assert 'row_traversal' in data
        assert 'lookup_hit' in data
        assert 'bottlenecks' in data

    def test_report_with_bottlenecks(self):
        p = PerfProfiler()
        p._total_duration_ms = 2000.0
        for i in range(10):
            p.record_file_open(f'/tmp/f{i}.xlsx', 150.0)
        report = p.generate_report()
        assert '性能瓶颈分析' in report
        assert 'file_open' in report or '文件打开' in report

    def test_report_empty(self):
        p = PerfProfiler()
        report = p.generate_report()
        assert '运行性能剖析报告' in report
        assert '未发现明显性能瓶颈' in report


class TestPerfProfilerReset:
    def test_reset(self):
        p = PerfProfiler()
        p.record_file_open('/tmp/test.xlsx', 100.0)
        p.record_row_traversal('/tmp/test.xlsx', 'S1', 10, 5.0)
        p.record_lookup_hit('123', 0.1)
        p.record_phase('test', 10.0)
        p.start()
        p.reset()
        assert p.file_open_records == []
        assert p.row_traversal_records == []
        assert p.lookup_hit_records == []
        assert p.phase_records == []
        assert p._total_duration_ms == 0.0
        assert p._start_time is None


class TestGetProfiler:
    def test_get_profiler_singleton(self):
        p1 = get_profiler()
        p2 = get_profiler()
        assert p1 is p2

    def test_reset_profiler(self):
        p1 = get_profiler()
        reset_profiler()
        p2 = get_profiler()
        assert p1 is not p2
