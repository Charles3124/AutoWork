"""
keyboard_controller.py

功能: 键盘功能模块
时间: 2025/10/21
版本: 1.0
"""

from pynput.keyboard import Controller

from utils.data_manager import DataManager
from utils.types import InsertFuncType
from utils.helpers import get_data_content, is_object, get_variable_content, type_text
from utils.mappings import key_mapping


class KeyboardController:
    """控制键盘功能"""

    def __init__(self, keyboard: Controller, data_manager: DataManager, insert_text: InsertFuncType):
        self.keyboard = keyboard
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_keyboard(self, event: str, detail: str) -> None:
        """处理所有键盘指令"""
        if '输入' in event:
            if detail == '剪贴板':
                self._input_clipboard()

            elif is_object(detail, target='数组'):
                self._input_from_array(detail)

            elif is_object(detail, target='变量'):
                self._input_from_variable(detail)

            else:
                self._input_text(detail)

        elif '按键' in event:
            if '按下' in event:
                self._key_press(detail)

            else:
                self._key_release(detail)

    @staticmethod
    def _input_clipboard() -> None:
        """从剪贴板输入"""
        type_text()

    def _input_from_array(self, detail: str) -> None:
        """从数组输入"""
        parts = detail.split('|')
        text, data_name = get_data_content(parts[0][3:], self.data_manager)
        type_text(text)
        self.data_manager.all_data_index[data_name] += int(parts[-1]) if parts[-1] != '' else 0

    def _input_from_variable(self, detail: str) -> None:
        """从变量输入"""
        type_text(get_variable_content(detail[3:], self.data_manager))

    @staticmethod
    def _input_text(detail: str) -> None:
        """输入文本"""
        type_text(detail)

    def _key_press(self, detail: str) -> None:
        """按键按下"""
        if detail in key_mapping:
            self.keyboard.press(key_mapping.get(detail))
        else:
            self.keyboard.press(detail)

    def _key_release(self, detail: str) -> None:
        """按键松开"""
        if detail in key_mapping:
            self.keyboard.release(key_mapping.get(detail))
        else:
            self.keyboard.release(detail)
