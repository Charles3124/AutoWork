"""
wait_controller.py

功能: 等待功能模块
时间: 2025/10/21
版本: 1.0
"""

import time

from utils.data_manager import DataManager
from utils.types import CommandType, InsertFuncType
from utils.helpers import get_position, get_image, are_images_similar


class WaitController:
    """控制循环功能"""

    def __init__(
            self, df_length: int,
            events: CommandType, details: CommandType,
            data_manager: DataManager, insert_text: InsertFuncType
    ):
        self.df_length = df_length
        self.events = events
        self.details = details
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_wait(self, i: int, detail: str, condition: str) -> int:
        """处理等待指令"""
        self.insert_text(f'开始等待...\n')
        start_time = time.time()

        parts = detail.split('|')
        negative_flag = '0' if '不' in parts[0] else '1'
        gap_time, max_time = map(float, parts[1].split(','))
        normal_break = True

        if '相似' in parts[0]:
            part1, part2, threshold = condition.split('|')

        while True:
            time.sleep(gap_time)

            # 是否存在文字或图片
            if '直到' in parts[0]:
                find_result, _ = get_position(condition, self.data_manager)
                condition_map = {'0': len(find_result) == 0, '1': len(find_result) != 0}

            # 两张图片是否相似
            elif '相似' in parts[0]:
                image1, image2 = get_image(part1, self.data_manager), get_image(part2, self.data_manager)
                similar = are_images_similar(image1, image2, threshold)
                condition_map = {'0': not similar, '1': similar}

            end_time = time.time()

            if condition_map[negative_flag]:
                break
            if end_time - start_time >= max_time:
                normal_break = False
                break

        self.insert_text(f'等待{'结束' if normal_break else '超时'}！共等了{end_time - start_time:.2f}秒\n')

        if i < self.df_length - 1 and '超时' in self.events[i + 1]:
            # 正常结束
            if normal_break:
                while '超时' not in self.events[i] or '}' not in self.details[i]:
                    i += 1

            # 超时结束
            else:
                i += 1

        return i + 1
