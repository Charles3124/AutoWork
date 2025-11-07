"""
mouse_controller.py

功能: 鼠标功能模块
时间: 2025/10/21
版本: 1.0
"""

from pynput.mouse import Controller

from utils.data_manager import DataManager
from utils.types import InsertFuncType
from utils.helpers import get_position, is_object, get_variable
from utils.mappings import button_mapping


class MouseController:
    """控制鼠标功能"""

    def __init__(self, mouse: Controller, data_manager: DataManager, insert_text: InsertFuncType):
        self.mouse = mouse
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_mouse(self, event: str, detail: str, condition: str) -> str:
        """处理所有鼠标指令"""
        if '平移' in event:
            self._mouse_move(detail, condition)
            return ''

        else:
            return self._mouse_click(event, detail, condition)

    def _mouse_move(self, detail: str, condition: str) -> None:
        """鼠标移动"""
        distance = (
            int(self.data_manager.all_variables[condition[3:]])
            if is_object(condition, target='变量') else float(condition)
        )
        move_map = {
            '上': (0, -distance),
            '下': (0, distance),
            '左': (-distance, 0),
            '右': (distance, 0)
        }
        x, y = self.mouse.position
        dx, dy = move_map.get(detail, (0, 0))
        self.mouse.position = (x + dx, y + dy)

    def _mouse_click(self, event: str, detail: str, condition: str) -> str:
        """鼠标点击"""
        if condition != 'nan':
            # 给定坐标点
            if '|' not in condition:
                x, y = get_variable(condition, self.data_manager)

            # 寻找图片或文字的位置
            else:
                find_result, print_text = get_position(condition, self.data_manager)

                # 没找到结果，直接退出操作
                if len(find_result) == 0:
                    self.insert_text(f'错误：没找到指定的{print_text}，请检查指定的{print_text}和查找范围！\n')
                    return 'error'

                x, y = find_result

            self.mouse.position = (x, y)

        if '单击' in event:
            self.mouse.click(button_mapping.get(detail))

        elif '双击' in event:
            self.mouse.click(button_mapping.get(detail))
            self.mouse.click(button_mapping.get(detail))

        return ''
