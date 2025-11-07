"""
data_writer.py

功能: 文件输入模块
时间: 2025/10/21
版本: 1.0
"""

from openpyxl import load_workbook

from utils.helpers import get_data_content, get_variable_content, is_object
from utils.data_manager import DataManager
from utils.types import InsertFuncType


class DataWriter:
    """文件输入控制器"""

    def __init__(self, data_manager: DataManager, insert_text: InsertFuncType):
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_write(self, detail: str, condition: str) -> None:
        """处理所有输入指令"""
        if 'excel' in detail.lower():
            self._write_excel(condition)

    def _write_excel(self, condition: str) -> None:
        """输入 Excel"""
        parts = condition.split('|')
        wb = load_workbook(parts[1])
        sheet = wb[parts[2]]

        if is_object(parts[0], target='数组'):
            text, _ = get_data_content(parts[0][3:], self.data_manager)
        elif is_object(parts[0], target='变量'):
            text = get_variable_content(parts[0][3:], self.data_manager)
        else:
            text = parts[0]

        target_cell = (
            str(self.data_manager.all_variables[parts[3][3:]])
            if is_object(parts[3], target='变量') else parts[3]
        )

        sheet[target_cell] = str(text)
        wb.save(parts[1])
        self.insert_text(f'内容“{text}”已写入Excel！\n')
