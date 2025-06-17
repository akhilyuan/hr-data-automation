"""
数据处理模块
负责数据清洗、转换和标准化
"""
import pandas as pd
from config.config import DataMappings
from typing import Optional

class DataProcessor:
    """数据处理核心类"""
    def __init__(self, mappings: DataMappings):
        self.mappings = mappings

    def apply_department_mappings(self, department_name: str, secondary_org_name: Optional[str] = None) -> str:
        if pd.isna(department_name) or department_name == '':
            return department_name
        
        department_str = department_name.strip()
        primary_mapping = self.mappings.DEPARTMENT_MAPPING.get(department_str, '')

        if primary_mapping == '参考二级组织':
            # TODO 其实这段代码的逻辑我没有看很懂
            secondary_key = '' if pd.isna(secondary_org_name) else str(secondary_org_name.strip())
            return self.mappings.SECONDARY_ORG_MAPPING.get(secondary_key, department_str)
        
        return self.mappings.DEPARTMENT_MAPPING.get(department_str, department_str)
    
    def is_frontline_staff(self, position: str) -> str:
        """判断是否为一线员工"""
        if pd.isna(position) or position == '':
            return '否'
        
        position_str = str(position).strip()
        if position_str in self.mappings.FRONTLINE_POSITIONS:
            return '是'
        
        # 模糊匹配
        for frontline_pos in self.mappings.FRONTLINE_POSITIONS:
            if frontline_pos in position_str or position_str in frontline_pos:
                return '是'

        return '否'
    
    def categorize_age(self, age: float) -> str:
        """年龄分组"""
        if pd.isna(age):
            return '未知'
        if age < 30:
            return '30岁以下'
        if age < 40:
            return '30-40岁'
        if age < 50:
            return '40-50岁'
        return '50岁以上'

    def standardize_education(self, education: str) -> str:
        """学历标准化"""
        if pd.isna(education) or education == '':
            return '高中以下'
        return self.mappings.EDUCATION_MAPPING.get(str(education).strip(), '高中以下')
    
    def clean_and_standarize(self, df: pd.DataFrame) -> pd.DataFrame:
        # 年龄处理
        if '年龄' in df.columns:
            df['年龄'] = pd.to_numeric(df['年龄'], errors='coerce')
            mean_age = df['年龄'].mean()
            df['年龄'] = df['年龄'].fillna(mean_age)
            df['年龄段'] = df['年龄'].apply(self.categorize_age)

        # 学历标准化
        if '最高学历' in df.columns:
            df['学历分组'] = df['最高学历'].apply(self.standardize_education)

        return df