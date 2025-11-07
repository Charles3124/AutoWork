"""
loop_controller.py

功能: 循环功能模块
时间: 2025/10/17
版本: 1.0
"""

from utils.data_manager import DataManager
from utils.types import CommandType, InsertFuncType


class LoopController:
    """控制循环功能"""

    def __init__(
            self, events: CommandType, details: CommandType,
            data_manager: DataManager, insert_text: InsertFuncType
    ):
        self.loop_info = []
        self.events = events
        self.details = details
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_loop(self, i: int, event: str, detail: str) -> int:
        """处理循环指令"""
        # break
        if '退出' in detail:
            self.insert_text(f'直接退出第{len(self.loop_info)}层循环！\n')
            last_info = self.loop_info.pop()
            i = last_info['end']

        # 循环结束或 continue
        elif "}" in detail or "继续" in detail:
            last_info = self.loop_info[-1]

            # 次数循环
            if last_info['times'] is not None:
                last_info['times'] -= 1

                if last_info['times'] > 0:
                    self.insert_text(f'第{len(self.loop_info)}层循环还有{last_info['times']}次\n')
                    i = last_info['begin']
                else:
                    i = last_info['end']
                    self.insert_text(f'第{len(self.loop_info)}层循环结束！\n')
                    self.loop_info.pop()

            # 数组循环
            else:
                data_name = last_info['name']
                data_index = self.data_manager.all_data_index[data_name]

                if data_index < len(self.data_manager.all_data[data_name]):
                    self.insert_text(f'循环继续\n')
                    i = last_info['begin']
                else:
                    i = last_info['end']
                    self.insert_text(f'第{len(self.loop_info)}层循环结束！\n')
                    self.loop_info.pop()

        # 循环开始
        else:
            loop_name = event.replace('{', '')
            end = i
            while loop_name not in self.events[end] or '}' not in self.details[end]:
                end += 1
            self.loop_info.append({'begin': i, 'end': end, 'times': None, 'name': None})

            # 次数循环
            if detail.isdigit():
                self.loop_info[-1]['times'] = int(detail)
                self.insert_text(f'第{len(self.loop_info)}层循环开始，共{detail}次...\n')

            # 数组循环
            else:
                self.loop_info[-1]['name'] = detail[3:]
                self.insert_text(f'第{len(self.loop_info)}层循环开始，数组名为{self.loop_info[-1]['name']}...\n')

        return i + 1
