"""
Excel合并模块
负责将多个Excel文件读取合并处理
"""
import pandas as pd
import os
from .data_processor import DataProcessor
from config.config import FileConfig, TARGET_COLUMNS
from typing import Optional

class ExcelMerger:
    def __init__(self, processor: DataProcessor, file_config: FileConfig):
        self.processor = processor
        self.file_config = file_config
        self.target_columns = TARGET_COLUMNS

    def extract_target_columns(self, df: pd.DataFrame, file_name: str) -> pd.DataFrame:
        """提取目标列"""
        available_columns = df.columns.tolist()

        matched_columns = [col for col in self.target_columns if col in available_columns]
        missing_columns = [col for col in self.target_columns if col not in available_columns]
        
        if missing_columns:
            print(f"文件 {file_name} 缺少以下目标列: {missing_columns}")

        extracted_df = df[matched_columns].copy()
        for col in missing_columns:
            extracted_df[col] = ''

        return extracted_df[self.target_columns]
    
    def merge_files(self, file1_path: str, file2_path: str, month: int, apply_mappings: bool = True) -> Optional[pd.DataFrame]:
        try:
            for file_path in [file1_path, file2_path]:
                if not os.path.exists(file_path):
                    print(f"文件 {file_path} 不存在")
                    return None
                
            df1 = pd.read_excel(file1_path, skiprows=1)
            df2 = pd.read_excel(file2_path, skiprows=1)

            df1_extracted = self.extract_target_columns(df1, "文件1")
            df2_extracted = self.extract_target_columns(df2, "文件2")

            # 合并数据
            merge_df = pd.concat([df1_extracted, df2_extracted], ignore_index=True)

            if apply_mappings:
                merge_df = self._apply_all_mappings(merge_df, month)

            initial_count = len(merge_df)
            merge_df = merge_df.drop_duplicates()
            final_count = len(merge_df)
            if initial_count != final_count:
                print(f"去除重复数据: {initial_count-final_count} 行")

            return merge_df
        
        except Exception as e:
            print(f"合并文件时出错: {e}")
            return None

    def _apply_all_mappings(self, df: pd.DataFrame, month: int) -> pd.DataFrame:
        """应用所有映射"""
        df.insert(0, '月份', month)
        
        if '部门/区县名称' in df.columns:
            has_secondary_org = 'BU/营服名称' in df.columns

            # TODO 这段代码也不是很懂
            if has_secondary_org:
                df['部门/区县名称'] = df.apply(
                    lambda row: self.processor.apply_department_mappings(
                        row['部门/区县名称'], row.get('BU/营服名称')
                    ), axis=1
                )
            else:
                df['部门/区县名称'] = df['部门/区县名称'].apply(
                    lambda x: self.processor.apply_department_mappings(x)
                )

            df['是否一线'] = df.apply(lambda row: self.processor.is_frontline_department(row['部门/区县名称'], row['岗位名称']), axis=1)


        if '姓名' in df.columns and '部门/区县名称' in df.columns:
            for name, new_dept in self.processor.mappings.SPECIAL_STAFF_MAPPING.items():
                mask = df['姓名'] == name
                # TODO 这段代码也不是很懂
                if mask.any():
                    df.loc[mask, '部门/区县名称'] = new_dept

        if '岗位名称' in df.columns:
            df['是否一线销售人员'] = df['岗位名称'].apply(self.processor.is_frontline_staff)
            frontline_count = (df['是否一线销售人员'] == '是').sum()
            print(f"一线人员数量: {frontline_count}")

        df = self.processor.clean_and_standarize(df)
        return df
    