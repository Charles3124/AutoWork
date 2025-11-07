"""
file_controller.py

功能: 文件操作模块
时间: 2025/10/21
版本: 1.0
"""

import os
import shutil

from utils.data_manager import DataManager
from utils.types import InsertFuncType


class FileController:
    """文件操作控制器"""

    def __init__(self, data_manager: DataManager, insert_text: InsertFuncType):
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_control(self, detail: str, condition: str) -> None:
        """处理所有操作指令"""
        if any(op in detail for op in ['复制', '移动']):
            self._move_and_copy_file(detail, condition)

        elif '删除' in detail:
            self._remove_file(condition)

        elif '写入' in detail:
            self._write_file(detail, condition)

    def _move_and_copy_file(self, detail: str, condition: str) -> None:
        """移动或复制文件"""
        source_file, target_folder = condition.split('|')

        if os.path.exists(source_file):
            file_name = os.path.basename(source_file)
            copy_target = os.path.join(target_folder, file_name)

            if '复制' in detail:
                shutil.copy(source_file, copy_target)
                self.insert_text(f'成功复制文件到{copy_target}\n')
            elif '移动' in detail:
                shutil.move(source_file, copy_target)
                self.insert_text(f'成功移动文件到{copy_target}\n')

        else:
            self.insert_text(f'文件{source_file}不存在，无法复制或移动！\n')

    def _remove_file(self, condition: str) -> None:
        """删除文件"""
        if os.path.exists(condition):
            os.remove(condition)
            self.insert_text(f'文件{condition}已被删除！\n')
        else:
            self.insert_text(f'文件{condition}不存在，无法删除！\n')

    def _write_file(self, detail: str, condition: str) -> None:
        """写入文件"""
        if 'txt' in detail:
            file_path, name = condition.split('|')

            # 以追加模式打开文件
            with open(file_path, 'a') as file:
                file.write(f'{self.data_manager.all_variables[name[3:]]}\n')  # 添加内容并换行
            self.insert_text(f'已将内容写入txt文件\n')
