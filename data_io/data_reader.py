"""
data_reader.py

功能: 文件读取模块
时间: 2025/10/21
版本: 1.0
"""

import pandas as pd

from utils.data_manager import DataManager
from utils.types import InsertFuncType


class DataReader:
    """文件读取控制器"""

    def __init__(self, data_manager: DataManager, insert_text: InsertFuncType):
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_read(self, detail: str, condition: str) -> None:
        """处理所有读取指令"""
        if 'excel' in detail.lower():
            self._read_excel(condition)

    def _read_excel(self, condition: str) -> None:
        """读取 Excel"""
        parts = condition.split('|')
        if not parts[1].endswith('.xlsx'):
            parts[1] += '.xlsx'

        temp_df = pd.read_excel(parts[1], sheet_name=parts[2], usecols=[parts[3]], dtype=str)
        last_non_empty_idx = temp_df.last_valid_index()
        temp_df = temp_df[:last_non_empty_idx + 1]

        self.data_manager.all_data[parts[0]] = list(temp_df[parts[3]].tolist())
        self.data_manager.all_data_index[parts[0]] = 0
        self.insert_text(f'Excel 读取成功！\n')
