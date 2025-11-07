"""
screen_shot_controller.py

功能: 屏幕截图模块
时间: 2025/10/21
版本: 1.0
"""

from utils.data_manager import DataManager
from utils.types import InsertFuncType
from utils.helpers import get_variable
from utils.funcs import capture_screen


class ScreenShotController:
    """屏幕截图控制器"""

    def __init__(self, data_manager: DataManager, insert_text: InsertFuncType):
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_shot(self, detail: str, condition: str) -> None:
        """处理屏幕截图指令"""
        image = capture_screen()

        if condition != 'nan':
            l, u, r, d = get_variable(condition, self.data_manager)
            image = image[u:d, l:r]

        self.data_manager.all_figures[detail] = image
        self.insert_text(f'截图成功！\n')
