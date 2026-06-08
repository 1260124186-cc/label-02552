#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批次管理模块测试脚本
"""

import os
import sys
import tempfile
import shutil
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import batch_manager as bm


def test_batch_id_generation():
    """测试批次号生成"""
    print("\n" + "=" * 50)
    print("测试 1: 批次号生成")
    print("=" * 50)

    batch_id1 = bm.generate_batch_id()
    batch_id2 = bm.generate_batch_id()

    assert batch_id1.startswith('BATCH'), f"批次号应以 BATCH 开头，实际为: {batch_id1}"
    assert len(batch_id1) == 23, f"批次号长度应为 23，实际为: {len(batch_id1)}"
    assert batch_id1 != batch_id2, "连续生成的批次号应不同"

    print(f"  生成批次号 1: {batch_id1}")
    print(f"  生成批次号 2: {batch_id2}")
    print("  ✅ 批次号生成测试通过")


def test_batch_lifecycle():
    """测试批次完整生命周期"""
    print("\n" + "=" * 50)
    print("测试 2: 批次完整生命周期")
    print("=" * 50)

    with tempfile.TemporaryDirectory() as tmpdir:
        print(f"  临时目录: {tmpdir}")

        manager = bm.BatchManager(script_dir=tmpdir)

        print("\n  --- 2.1 创建批次 ---")
        batch_info = manager.start_batch(
            input_folder='/test/input',
            operator='test_user'
        )
        batch_id = batch_info.batch_id
        print(f"  创建批次: {batch_id}")
        print(f"  状态: {batch_info.status}")
        print(f"  归档目录: {batch_info.batch_dir}")

        assert batch_info.status == 'running', "新建批次状态应为 running"
        assert os.path.exists(batch_info.batch_dir), "批次目录应已创建"

        print("\n  --- 2.2 模拟归档文件 ---")
        summary_path = os.path.join(tmpdir, '银行流水总表.xlsx')
        with open(summary_path, 'w') as f:
            f.write('测试总表内容')

        log_path = os.path.join(tmpdir, 'bankcheck.log')
        with open(log_path, 'w') as f:
            f.write('测试日志内容')

        archived_summary = manager.archive_summary_table(batch_id, summary_path)
        archived_log = manager.archive_log_file(batch_id, log_path)

        print(f"  总表归档路径: {archived_summary}")
        print(f"  日志归档路径: {archived_log}")

        assert os.path.exists(archived_summary), "总表应已归档"
        assert os.path.exists(archived_log), "日志应已归档"

        print("\n  --- 2.3 完成批次并生成报告 ---")
        result_data = {
            'total_records': 100,
            'new_records': 50,
            'duplicate_records': 5,
            'processed_files': ['file1.xlsx', 'file2.xlsx'],
            'unprocessed_files': ['unknown.xlsx'],
            'error_files': [('error.xlsx', '格式错误')],
            'incremental_mode': True,
            'output_folder': '/test/output',
            'summary_table_path': archived_summary,
            'log_file_path': archived_log,
        }
        finished_batch = manager.finish_batch(batch_id, result_data, status='success')

        print(f"  批次状态: {finished_batch.status}")
        print(f"  总记录数: {finished_batch.total_records}")
        print(f"  报告路径: {finished_batch.report_path}")

        assert finished_batch.status == 'success', "批次状态应为 success"
        assert finished_batch.total_records == 100, "总记录数应为 100"
        assert os.path.exists(finished_batch.report_path), "检验报告应已生成"
        assert os.path.exists(os.path.join(batch_info.batch_dir, f"{batch_id}_metadata.json")), "元数据文件应已生成"

        print("\n  --- 2.4 查看检验报告内容 ---")
        with open(finished_batch.report_path, 'r', encoding='utf-8') as f:
            report_content = f.read()
        assert '银行流水检验报告' in report_content, "报告应包含标题"
        assert batch_id in report_content, "报告应包含批次号"
        assert '处理统计' in report_content, "报告应包含处理统计"
        print("  ✅ 检验报告内容正确")

        print("\n  ✅ 批次生命周期测试通过")


def test_query_and_backtrack():
    """测试查询和回溯功能"""
    print("\n" + "=" * 50)
    print("测试 3: 查询和回溯功能")
    print("=" * 50)

    with tempfile.TemporaryDirectory() as tmpdir:
        manager = bm.BatchManager(script_dir=tmpdir)

        print("\n  --- 3.1 创建多个测试批次 ---")
        for i in range(5):
            batch = manager.start_batch(input_folder=f'/test/input_{i}', operator=f'user_{i % 2}')
            result_data = {
                'total_records': 100 * (i + 1),
                'new_records': 50 * (i + 1),
                'duplicate_records': i,
                'processed_files': [f'file_{i}.xlsx'],
                'unprocessed_files': [],
                'error_files': [],
                'incremental_mode': i % 2 == 0,
                'summary_table_path': '',
                'log_file_path': '',
            }
            status = 'success' if i < 4 else 'failed'
            manager.finish_batch(batch.batch_id, result_data, status=status)
            print(f"  创建批次 {i+1}: {batch.batch_id}, 状态: {status}")

        print("\n  --- 3.2 查询所有批次 ---")
        all_batches = manager.query_batches(limit=10)
        print(f"  查询到 {len(all_batches)} 个批次")
        assert len(all_batches) == 5, "应查询到 5 个批次"

        print("\n  --- 3.3 按状态查询 ---")
        success_batches = manager.query_batches(status='success')
        print(f"  成功批次: {len(success_batches)} 个")
        assert len(success_batches) == 4, "应查询到 4 个成功批次"

        print("\n  --- 3.4 按操作员查询 ---")
        user0_batches = manager.query_batches(operator='user_0')
        print(f"  user_0 的批次: {len(user0_batches)} 个")
        assert len(user0_batches) >= 2, "user_0 应有至少 2 个批次"

        print("\n  --- 3.5 按最小记录数查询 ---")
        large_batches = manager.query_batches(min_records=300)
        print(f"  记录数 >= 300 的批次: {len(large_batches)} 个")
        assert len(large_batches) >= 2, "应有至少 2 个批次记录数 >= 300"

        print("\n  --- 3.6 获取批次详情 ---")
        first_batch = all_batches[0]
        detail = manager.get_batch_detail(first_batch.batch_id)
        assert detail is not None, "应能获取批次详情"
        assert detail['batch_info']['batch_id'] == first_batch.batch_id, "批次号应匹配"
        print(f"  批次 {first_batch.batch_id} 详情获取成功")
        print(f"    - 总记录数: {detail['batch_info']['total_records']}")
        print(f"    - 归档文件数: {len(detail['files'])}")

        print("\n  --- 3.7 恢复批次文件 ---")
        restore_dir = os.path.join(tmpdir, 'restore_test')
        restored = manager.restore_batch(first_batch.batch_id, restore_dir)
        print(f"  恢复到目录: {restore_dir}")
        print(f"  恢复文件数: {len(restored)}")
        assert len(restored) >= 2, "应至少恢复 2 个文件（元数据 + 报告）"
        for name, path in restored.items():
            assert os.path.exists(path), f"恢复的文件应存在: {name}"
            print(f"    - {name}")

        print("\n  --- 3.8 获取统计信息 ---")
        stats = manager.get_statistics()
        print(f"  总批次数: {stats['total_batches']}")
        print(f"  成功批次: {stats['success_batches']}")
        print(f"  失败批次: {stats['failed_batches']}")
        print(f"  累计记录: {stats['total_records']}")
        assert stats['total_batches'] == 5, "总批次数应为 5"
        assert stats['total_records'] == 1500, "累计记录数应为 1500 (100+200+300+400+500)"

        print("\n  --- 3.9 删除批次 ---")
        delete_batch_id = all_batches[-1].batch_id
        print(f"  删除批次: {delete_batch_id}")
        result = manager.delete_batch(delete_batch_id)
        assert result, "删除应成功"
        assert not os.path.exists(manager._get_batch_dir(delete_batch_id, bm.get_date_str())), "批次目录应已删除"

        remaining = manager.query_batches(limit=10)
        assert len(remaining) == 4, "删除后应剩 4 个批次"
        print("  ✅ 批次删除成功")

        print("\n  ✅ 查询和回溯测试通过")


def test_edge_cases():
    """测试边界情况"""
    print("\n" + "=" * 50)
    print("测试 4: 边界情况处理")
    print("=" * 50)

    with tempfile.TemporaryDirectory() as tmpdir:
        manager = bm.BatchManager(script_dir=tmpdir)

        print("\n  --- 4.1 查询不存在的批次 ---")
        detail = manager.get_batch_detail('BATCH-NOT-EXIST')
        assert detail is None, "不存在的批次应返回 None"
        print("  ✅ 不存在的批次查询返回 None")

        print("\n  --- 4.2 删除不存在的批次 ---")
        result = manager.delete_batch('BATCH-NOT-EXIST')
        assert result is False, "删除不存在的批次应返回 False"
        print("  ✅ 删除不存在的批次返回 False")

        print("\n  --- 4.3 归档不存在的文件 ---")
        batch = manager.start_batch()
        result = manager.archive_summary_table(batch.batch_id, '/not/exist/file.xlsx')
        assert result == '', "归档不存在的文件应返回空字符串"
        print("  ✅ 归档不存在的文件返回空字符串")

        print("\n  --- 4.4 按日期查询批次ID ---")
        today = bm.get_date_str()
        batch_ids = manager.get_batch_ids_by_date(today)
        assert len(batch_ids) >= 1, "应查询到今天的批次"
        print(f"  今日批次: {batch_ids}")

        not_exist_date = '2000-01-01'
        batch_ids_empty = manager.get_batch_ids_by_date(not_exist_date)
        assert len(batch_ids_empty) == 0, "不存在日期应返回空列表"
        print("  ✅ 边界日期查询正确")

        print("\n  ✅ 边界情况测试通过")


def test_global_manager():
    """测试全局管理器单例"""
    print("\n" + "=" * 50)
    print("测试 5: 全局管理器单例")
    print("=" * 50)

    with tempfile.TemporaryDirectory() as tmpdir:
        manager1 = bm.get_batch_manager(tmpdir)
        manager2 = bm.get_batch_manager(tmpdir)

        assert manager1 is manager2, "两次获取应返回同一实例"
        print("  ✅ 全局管理器单例测试通过")

        bm._global_batch_manager = None


def main():
    print("\n" + "=" * 60)
    print("批次管理模块测试")
    print("=" * 60)

    try:
        test_batch_id_generation()
        test_batch_lifecycle()
        test_query_and_backtrack()
        test_edge_cases()
        test_global_manager()

        print("\n" + "=" * 60)
        print("✅ 所有测试通过！")
        print("=" * 60)
        return 0
    except AssertionError as e:
        print(f"\n❌ 断言失败: {e}")
        import traceback
        traceback.print_exc()
        return 1
    except Exception as e:
        print(f"\n❌ 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
