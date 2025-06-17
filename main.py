#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HR数据管理系统主程序 - 自动化版本
自动执行完整的数据合并和报告生成流程
"""

import os
import sys
import argparse
from datetime import datetime
from module.HR_manager import HRDataManager


def print_banner():
    """打印程序标题"""
    print("=" * 60)
    print("           HR数据管理系统")
    print("      自动化Excel文件合并与报告生成")
    print("=" * 60)
    print()


def get_current_month():
    """获取当前月份（默认处理上个月的数据）"""
    current_date = datetime.now()
    # 通常处理上个月的数据
    if current_date.month == 1:
        return 12  # 如果是1月，处理去年12月的数据
    else:
        return current_date.month - 1


def setup_hr_manager(base_path=None, output_folder='./output'):
    """初始化HR管理器"""
    print("🔧 系统初始化中...")
    
    # 确保输出目录存在
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        hr_manager = HRDataManager(base_path=base_path, output_folder=output_folder)
        print(f"✅ 系统初始化完成")
        print(f"📁 输出目录: {os.path.abspath(output_folder)}")
        return hr_manager
    except Exception as e:
        print(f"❌ 系统初始化失败: {e}")
        return None


def auto_process_workflow(hr_manager, month=None, apply_mappings=True):
    """自动执行完整工作流程"""
    if month is None:
        month = get_current_month()
    
    print(f"🚀 开始自动处理 {month} 月份数据...")
    print(f"📊 数据映射: {'启用' if apply_mappings else '禁用'}")
    print("=" * 50)
    
    try:
        # 执行完整工作流程
        report_path = hr_manager.process_monthly_workflow(month)
        
        if report_path:
            print("=" * 50)
            print("🎉 自动化流程执行成功!")
            print(f"📄 最终报告: {os.path.abspath(report_path)}")
            print(f"📅 处理月份: {month}月")
            print(f"⏰ 完成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            return True
        else:
            print("❌ 自动化流程执行失败")
            return False
            
    except Exception as e:
        print(f"❌ 自动化流程执行过程中发生错误: {e}")
        return False


def main():
    """主函数 - 自动化版本"""
    # 命令行参数解析
    parser = argparse.ArgumentParser(description='HR数据管理系统 - 自动化处理')
    parser.add_argument('--month', '-m', type=int, choices=range(1, 13), 
                       help='指定处理的月份 (1-12)，默认为上个月')
    parser.add_argument('--base-path', '-p', type=str, 
                       help='数据文件基础路径')
    parser.add_argument('--output', '-o', type=str, default='./output',
                       help='输出文件夹路径，默认为 ./output')
    parser.add_argument('--no-mappings', action='store_true',
                       help='禁用数据映射和清洗，仅进行原始数据合并')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='详细输出模式')
    
    args = parser.parse_args()
    
    print_banner()
    
    # 设置详细输出
    if args.verbose:
        print("🔍 详细输出模式已启用")
        print(f"📋 参数信息:")
        print(f"   - 月份: {args.month or '自动检测'}")
        print(f"   - 基础路径: {args.base_path or '默认'}")
        print(f"   - 输出路径: {args.output}")
        print(f"   - 数据映射: {'禁用' if args.no_mappings else '启用'}")
        print()
    
    # 初始化HR管理器
    hr_manager = setup_hr_manager(
        base_path=args.base_path,
        output_folder=args.output
    )
    
    if not hr_manager:
        print("❌ 程序启动失败，请检查配置")
        return 1
    
    print()
    
    # 自动执行工作流程
    try:
        success = auto_process_workflow(
            hr_manager=hr_manager,
            month=args.month,
            apply_mappings=not args.no_mappings
        )
        
        if success:
            print("\n🎯 程序执行完成!")
            return 0
        else:
            print("\n⚠️  程序执行失败!")
            return 1
            
    except KeyboardInterrupt:
        print("\n\n⏹️  程序被用户中断")
        return 1
    except Exception as e:
        print(f"\n❌ 发生未预期的错误: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def simple_auto_run():
    """简单的自动运行函数 - 无参数版本"""
    print_banner()
    print("🤖 自动模式启动...")
    print("💡 提示: 如需自定义参数，请使用命令行参数")
    print("   例如: python main.py --month 3 --verbose")
    print()
    
    # 使用默认设置
    hr_manager = setup_hr_manager()
    if not hr_manager:
        print("❌ 程序启动失败")
        return
    
    print()
    
    # 自动执行
    month = get_current_month()
    print(f"📅 自动检测处理月份: {month}月")
    
    success = auto_process_workflow(hr_manager, month)
    
    if success:
        print("\n🎯 自动化处理完成!")
    else:
        print("\n⚠️  自动化处理失败!")
    
    print("\n按回车键退出...")
    input()


if __name__ == "__main__":
    try:
        # 检查是否有命令行参数
        if len(sys.argv) > 1:
            # 有参数，使用命令行模式
            exit_code = main()
            sys.exit(exit_code)
        else:
            # 无参数，使用简单自动模式
            simple_auto_run()
            
    except Exception as e:
        print(f"❌ 程序启动失败: {e}")
        input("按回车键退出...")
        sys.exit(1)