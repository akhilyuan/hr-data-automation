from config.config import FileConfig, DataMappings
from .data_processor import DataProcessor
from .excel_merger import ExcelMerger
from .report_generator import ReportGenerator
from typing import Optional, Dict

class HRDataManager:
    def __init__(self, base_path: str = None, output_folder: str = './output'):
        self.file_config = FileConfig(base_path=base_path, output_folder=output_folder)
        self.mappings = DataMappings()
        self.processor = DataProcessor(self.mappings)
        self.merger = ExcelMerger(self.processor, self.file_config)
        self.report_generator = ReportGenerator(self.file_config)

    def merge_monthly_data(self, month: int, apply_mappings: bool = True) -> Optional[str]:
        """合并指定月份的数据"""
        print("请选择输入方式:")
        print("1. 快速模式 - 只输入月份数字")
        print("2. 完整路径模式 - 手动输入完整文件路径")
        input_choice = input("\n请选择模式（1或2）：").strip()

        if input_choice == "1":
            try:
                file1_path, file2_path, output_path = self.file_config.get_monthly_files(month)

                merged_df = self.merger.merge_files(file1_path, file2_path, month, apply_mappings)
                if merged_df is None:
                    print("合并失败，可能是文件不存在或格式不正确。")
                    return None
                
                merged_df.to_excel(output_path, index=False)
                return output_path
            except Exception as e:
                print(f"合并数据时出错: {e}")
                return None
            
        elif input_choice == "2":
            # 完整路径模式
            file1_path = input("请输入第一个Excel文件路径: ").strip().strip('"')
            file2_path = input("请输入第二个Excel文件路径: ").strip().strip('"')
            output_path = input("请输入输出文件路径（回车使用默认名称 merged_data.xlsx）: ").strip().strip('"')
            if not output_path:
                output_path = "merged_data.xlsx"
            
            try:
                merged_df = self.merger.merge_files(file1_path, file2_path, month, apply_mappings)
                if merged_df is None:
                    return None
                
                # 保存合并后的数据
                merged_df.to_excel(output_path, index=False)
                
                return output_path
                
            except Exception as e:
                print(f"合并数据时出错: {e}")
                return None
            

    def generate_monthly_report(self, merged_data_path: str) -> Optional[str]:
        try:
            df = self.report_generator.load_merge_data(merged_data_path)
            if df is None:
                print("加载合并数据失败，无法生成报告。")
                return None
            
            stats = self.report_generator.generate_summary_stats(df)
            report_path = self.report_generator.create_excel_report(stats)
            return report_path
        except Exception as e:
            print(f"生成报告时出错: {e}")
            return None
        
    def process_monthly_workflow(self, month:int) -> Optional[str]:
        
        merged_data_path = self.merge_monthly_data(month)
        if merged_data_path is None:
            print("数据合并失败，无法继续生成报告。")
            return None
        
        report_path = self.generate_monthly_report(merged_data_path)
        if report_path is None:
            print("报告生成失败。")
            return None
        
        print(f"工作流处理完成，报告已保存到: {report_path}")
        return report_path
