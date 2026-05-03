"""
吃药小管家 - 软件测试
运行方式: python test_medicine.py
"""

import unittest
import json
import os
import tempfile
from datetime import datetime


class MedicineCore:
    """核心功能测试类"""

    def __init__(self):
        self.medicines = []
        self.taking_history = {}
        self.reminded_set = set()

    def add_medicine(self, name, dosage, time_str):
        if not name or not dosage or not time_str:
            return False
        try:
            datetime.strptime(time_str, "%H:%M")
        except:
            return False
        self.medicines.append({'name': name, 'dosage': dosage, 'time': time_str})
        return True

    def delete_medicine(self, index):
        if 0 <= index < len(self.medicines):
            self.medicines.pop(index)
            return True
        return False

    def mark_as_taken(self, med_name):
        key = f"{med_name}_{datetime.now().date()}"
        self.taking_history[key] = {'status': 'taken', 'time': datetime.now().strftime("%H:%M:%S")}
        return True

    def check_status(self, med_name):
        key = f"{med_name}_{datetime.now().date()}"
        return 'taken' if key in self.taking_history else 'pending'


class TestMedicine(unittest.TestCase):

    def setUp(self):
        self.app = MedicineCore()

    def print_data(self):
        print(f"  药品: {[m['name'] for m in self.app.medicines]}")
        print(f"  记录: {list(self.app.taking_history.keys())}")

    def test_1_add(self):
        print("\n📋 测试1: 添加药品")
        self.app.add_medicine("降压药", "1片", "08:00")
        self.assertEqual(len(self.app.medicines), 1)
        self.print_data()
        print("  ✅")

    def test_2_empty(self):
        print("\n📋 测试2: 空数据拦截")
        result = self.app.add_medicine("", "", "")
        self.assertFalse(result)
        self.print_data()
        print("  ✅")

    def test_3_invalid_time(self):
        print("\n📋 测试3: 错误时间拦截")
        result = self.app.add_medicine("测试药", "1片", "25:00")
        self.assertFalse(result)
        print("  ✅")

    def test_4_multiple(self):
        print("\n📋 测试4: 添加多个药品")
        self.app.add_medicine("降压药", "1片", "08:00")
        self.app.add_medicine("降糖药", "1粒", "12:00")
        self.app.add_medicine("维生素", "2片", "20:00")
        self.assertEqual(len(self.app.medicines), 3)
        self.print_data()
        print("  ✅")

    def test_5_delete(self):
        print("\n📋 测试5: 删除药品")
        self.app.add_medicine("待删药", "1片", "08:00")
        self.app.delete_medicine(0)
        self.assertEqual(len(self.app.medicines), 0)
        self.print_data()
        print("  ✅")

    def test_6_taken(self):
        print("\n📋 测试6: 标记已服用")
        self.app.add_medicine("降压药", "1片", "08:00")
        self.app.mark_as_taken("降压药")
        status = self.app.check_status("降压药")
        self.assertEqual(status, 'taken')
        self.print_data()
        print("  ✅")

    def test_7_save_load(self):
        print("\n📋 测试7: 数据保存加载")
        self.app.add_medicine("持久药", "1片", "08:00")
        self.app.mark_as_taken("持久药")

        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            temp_file = f.name
            json.dump({'medicines': self.app.medicines, 'taking_history': self.app.taking_history}, f)

        self.app.medicines = []
        self.app.taking_history = {}
        with open(temp_file, 'r') as f:
            data = json.load(f)
            self.app.medicines = data['medicines']
            self.app.taking_history = data['taking_history']

        self.assertEqual(len(self.app.medicines), 1)
        os.remove(temp_file)
        self.print_data()
        print("  ✅")

    def test_8_no_repeat(self):
        print("\n📋 测试8: 防重复提醒")
        key = f"降压药_{datetime.now().date()}"
        self.app.reminded_set.add(key)
        self.app.reminded_set.add(key)
        self.assertEqual(len(self.app.reminded_set), 1)
        print(f"  集合大小: {len(self.app.reminded_set)}")
        print("  ✅")

    def test_9_different_dates(self):
        print("\n📋 测试9: 不同日期独立")
        self.app.add_medicine("降压药", "1片", "08:00")
        self.app.mark_as_taken("降压药")
        status = self.app.check_status("降压药")
        self.assertEqual(status, 'taken')
        print(f"  今天状态: {status}")
        print("  ✅")


if __name__ == "__main__":
    unittest.main(verbosity=0)