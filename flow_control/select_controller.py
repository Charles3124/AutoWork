"""
select_controller.py

功能: 选择功能模块
时间: 2025/10/17
版本: 1.0
"""

import os

import pandas as pd

from utils.data_manager import DataManager
from utils.types import CommandType, InsertFuncType
from utils.helpers import (
    are_images_similar, is_object, get_position, get_image,
    get_data_content, get_variable_content
)


class SelectController:
    """控制选择功能"""

    def __init__(
            self, events: CommandType, details: CommandType,
            data_manager: DataManager, insert_text: InsertFuncType
    ):
        self.select_info = []
        self.events = events
        self.details = details
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_select(self, i: int, event: str, detail: str, condition: str) -> int:
        """处理循环指令"""
        # 选择开始
        if '{' in event:
            self.select_info.append([event.replace('{', ''), False])
            self.insert_text(f'第{len(self.select_info)}层选择开始...\n')

        # 选择结束
        if '}' in detail:
            self.insert_text(f'第{len(self.select_info)}层选择结束！\n')
            self.select_info.pop()

        # 已经执行过其他分支，跳过
        elif self.select_info[-1][1]:
            while self.select_info[-1][0] not in self.events[i + 1] or '}' not in self.details[i + 1]:
                i += 1
            self.insert_text(f'跳过剩余分支\n')

        # 进行判断
        else:
            self.insert_text(f'开始判断...')
            if '其他' in detail:
                judgement_result = True
            else:
                judgement_result = self._judge_condition(condition)

            if judgement_result:
                self.insert_text(f'判断为真\n')
                self.select_info[-1][1] = True
            else:
                self.insert_text(f'判断为假\n')
                while self.select_info[-1][0] not in self.events[i + 1]:
                    i += 1

        return i + 1

    def _judge_condition(self, condition: str) -> bool:
        """判断选择条件是否成立"""
        # 文件是否存在
        if condition.startswith('存在文件'):
            file_name = condition[5:]
            return (
                os.path.exists(str(get_variable_content(file_name[3:], self.data_manager)))
                if is_object(file_name, target='变量')
                else os.path.exists(file_name)
            )

        # 图片是否相似
        if is_object(condition, target='相似'):
            part1, part2, threshold = condition[3:].split('|')
            image1, image2 = get_image(part1, self.data_manager), get_image(part2, self.data_manager)
            return are_images_similar(image1, image2, threshold)

        # 文字或变量是否在数组中
        if '在数组' in condition:
            parts = condition.split('在数组')
            data_name = parts[1][1:]
            target = (
                str(get_variable_content(parts[0][3:], self.data_manager))
                if is_object(parts[0], target='变量')
                else parts[0]
            )

            for content in self.data_manager.all_data[data_name]:
                if not pd.isna(content) and target == str(content):
                    return True
            return False

        # 内容是否相同
        if is_object(condition, target='是否相同'):
            parts = condition[5:].split('|')
            results = []

            for part in parts:
                if is_object(part, target='数组'):
                    data_content, _ = get_data_content(part[3:], self.data_manager)
                    results.append(str(data_content))
                elif is_object(part, target='变量'):
                    variable_content = get_variable_content(part[3:], self.data_manager)
                    results.append(str(variable_content))
                else:
                    results.append(part)
            return all(r == results[0] for r in results)

        # 图片或文字是否在指定范围
        if '|' in condition:
            find_result, _ = get_position(condition, self.data_manager)
            return len(find_result) != 0

        return False
