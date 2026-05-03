"""
吃药小管家 - 运行维护工具
运行方式: python maintenance.py [选项]
"""

import os
import json
import shutil
import sys
from datetime import datetime


class Maintenance:
    def __init__(self):
        self.data_file = "medicine_data.json"
        self.backup_dir = "backups"

    def create_backup(self):
        """创建数据备份"""
        if not os.path.exists(self.data_file):
            print("❌ 数据文件不存在")
            return False

        os.makedirs(self.backup_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(self.backup_dir, f"backup_{timestamp}.json")

        shutil.copy2(self.data_file, backup_file)
        print(f"✅ 备份成功: {backup_file}")
        return True

    def list_backups(self):
        """列出所有备份"""
        if not os.path.exists(self.backup_dir):
            print("📂 暂无备份")
            return []

        files = [f for f in os.listdir(self.backup_dir) if f.endswith('.json')]
        if not files:
            print("📂 暂无备份")
            return []

        print("\n备份列表:")
        for i, f in enumerate(sorted(files, reverse=True), 1):
            size = os.path.getsize(os.path.join(self.backup_dir, f))
            print(f"  {i}. {f} ({size}字节)")
        return files

    def restore_backup(self, backup_name):
        """恢复备份"""
        backup_path = os.path.join(self.backup_dir, backup_name)
        if not os.path.exists(backup_path):
            print(f"❌ 备份不存在: {backup_name}")
            return False

        # 自动备份当前数据
        if os.path.exists(self.data_file):
            self.create_backup()

        shutil.copy2(backup_path, self.data_file)
        print(f"✅ 恢复成功: {backup_name}")
        return True

    def check_data(self):
        """检查数据完整性"""
        if not os.path.exists(self.data_file):
            print("❌ 数据文件不存在")
            return False

        try:
            with open(self.data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            required = ['medicines', 'taking_history']
            missing = [f for f in required if f not in data]

            if missing:
                print(f"⚠️ 缺少字段: {missing}")
                return False

            med_count = len(data.get('medicines', []))
            history_count = len(data.get('taking_history', {}))
            print(f"✅ 数据完整: {med_count}个药品, {history_count}条记录")
            return True

        except Exception as e:
            print(f"❌ 数据损坏: {e}")
            return False

    def repair_data(self):
        """修复损坏数据"""
        if not os.path.exists(self.data_file):
            default = {'medicines': [], 'taking_history': {}, 'emergency_phone': '', 'timeout_minutes': 20}
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(default, f, ensure_ascii=False, indent=2)
            print("✅ 已创建默认数据")
            return True

        try:
            # 备份损坏文件
            backup_name = f"corrupted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            shutil.copy2(self.data_file, backup_name)
            print(f"📦 已备份损坏文件: {backup_name}")

            # 创建空数据
            default = {'medicines': [], 'taking_history': {}, 'emergency_phone': '', 'timeout_minutes': 20}
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(default, f, ensure_ascii=False, indent=2)
            print("✅ 数据已修复")
            return True

        except Exception as e:
            print(f"❌ 修复失败: {e}")
            return False

    def clean_old_backups(self, days=30):
        """清理旧备份"""
        if not os.path.exists(self.backup_dir):
            print("📂 备份目录不存在")
            return

        cutoff = datetime.now().timestamp() - (days * 86400)
        deleted = 0

        for f in os.listdir(self.backup_dir):
            if f.endswith('.json'):
                path = os.path.join(self.backup_dir, f)
                if os.path.getmtime(path) < cutoff:
                    os.remove(path)
                    deleted += 1

        print(f"✅ 清理完成，删除{deleted}个旧备份")

    def show_logs(self):
        """查看运行日志"""
        log_dir = "logs"
        if not os.path.exists(log_dir):
            print("📂 暂无日志")
            return

        log_files = [f for f in os.listdir(log_dir) if f.endswith('.log')]
        if not log_files:
            print("📂 暂无日志")
            return

        latest = sorted(log_files)[-1]
        with open(os.path.join(log_dir, latest), 'r', encoding='utf-8') as f:
            lines = f.readlines()[-20:]
            print(f"\n最新日志 ({latest}):")
            print("-"*40)
            for line in lines:
                print(line.strip())

    def export_data(self):
        """导出数据到可读格式"""
        if not os.path.exists(self.data_file):
            print("❌ 数据文件不存在")
            return

        try:
            with open(self.data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            export_file = f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(export_file, 'w', encoding='utf-8') as f:
                f.write("="*50 + "\n")
                f.write("吃药小管家 - 数据导出\n")
                f.write(f"导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*50 + "\n\n")

                f.write("【药品列表】\n")
                for med in data.get('medicines', []):
                    f.write(f"  💊 {med.get('name', '未知')} | {med.get('dosage', '未知')} | 时间: {med.get('time', '未知')}\n")

                f.write("\n【服药记录】\n")
                for key, val in data.get('taking_history', {}).items():
                    f.write(f"  {key}: {val.get('status', '未知')} at {val.get('time', '未知')}\n")

                f.write(f"\n紧急联系人: {data.get('emergency_phone', '未设置')}\n")
                f.write(f"超时时间: {data.get('timeout_minutes', 20)}分钟\n")

            print(f"✅ 数据已导出: {export_file}")
        except Exception as e:
            print(f"❌ 导出失败: {e}")


def show_menu():
    """显示菜单"""
    print("\n" + "="*40)
    print("     吃药小管家 - 维护工具")
    print("="*40)
    print("1. 创建备份")
    print("2. 查看备份")
    print("3. 恢复备份")
    print("4. 检查数据")
    print("5. 修复数据")
    print("6. 清理旧备份")
    print("7. 查看日志")
    print("8. 导出数据")
    print("0. 退出")
    print("="*40)


def main():
    """主函数 - 支持循环操作"""
    m = Maintenance()

    # 命令行模式
    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        if cmd == '--backup':
            m.create_backup()
        elif cmd == '--check':
            m.check_data()
        elif cmd == '--repair':
            m.repair_data()
        elif cmd == '--clean':
            days = int(sys.argv[2]) if len(sys.argv) > 2 else 30
            m.clean_old_backups(days)
        elif cmd == '--export':
            m.export_data()
        else:
            print("用法: python maintenance.py [选项]")
            print("  --backup    创建备份")
            print("  --check     检查数据")
            print("  --repair    修复数据")
            print("  --clean N   清理N天前的备份")
            print("  --export    导出数据")
        return

    # 【修复】交互模式 - 循环直到用户选择退出
    while True:
        show_menu()
        choice = input("请选择 (0-8): ").strip()

        if choice == '1':
            m.create_backup()
        elif choice == '2':
            m.list_backups()
        elif choice == '3':
            backups = m.list_backups()
            if backups:
                try:
                    idx = int(input("选择备份编号: ").strip())
                    if 1 <= idx <= len(backups):
                        m.restore_backup(backups[idx - 1])
                    else:
                        print("❌ 无效编号")
                except ValueError:
                    print("❌ 请输入数字")
        elif choice == '4':
            m.check_data()
        elif choice == '5':
            confirm = input("确认修复数据？(y/n): ").strip().lower()
            if confirm == 'y':
                m.repair_data()
        elif choice == '6':
            try:
                days = input("保留多少天内的备份(默认30): ").strip()
                days = int(days) if days else 30
                m.clean_old_backups(days)
            except ValueError:
                print("❌ 请输入数字")
        elif choice == '7':
            m.show_logs()
        elif choice == '8':
            m.export_data()
        elif choice == '0':
            print("👋 再见！")
            break
        else:
            print("❌ 无效选项，请重新选择")

        # 操作完成后暂停，让用户看到结果
        input("\n按回车键继续...")


if __name__ == "__main__":
    main()