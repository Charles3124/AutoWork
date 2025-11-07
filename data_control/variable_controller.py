"""
variable_controller.py

功能: 变量功能模块
时间: 2025/10/21
版本: 1.0
"""

import os
from datetime import datetime, timedelta

import cv2
from easyocr import Reader
from PIL import Image

from utils.data_manager import DataManager
from utils.types import InsertFuncType
from utils.helpers import get_variable, get_position, is_object, get_variable_content, get_data_content
from utils.funcs import capture_screen


class VariableController:
    """控制变量功能"""

    def __init__(self, readers: dict[str, Reader], data_manager: DataManager, insert_text: InsertFuncType):
        self.readers = readers
        self.data_manager = data_manager
        self.insert_text = insert_text

    def handle_variable(self, event: str, detail: str, condition: str) -> str:
        """处理所有变量指令"""
        if '获取' in event:
            return self._get_variable(detail, condition)

        elif '运算' in event:
            self._calculate_variable(detail, condition)

        elif '整数' in event or '小数' in event:
            self._assign_numeric_variable(event, detail, condition)

        elif '文字' in event:
            return self._assign_string_variable(detail, condition)

        return ''

    def _assign_numeric_variable(self, event: str, detail: str, condition: str) -> None:
        """声明数字变量"""
        if is_object(condition, target='变量'):
            self.data_manager.all_variables[detail] = get_variable_content(condition[3:], self.data_manager)
        elif is_object(condition, target='数组'):
            data_content, _ = get_data_content(condition[3:], self.data_manager)
            self.data_manager.all_variables[detail] = data_content
        else:
            self.data_manager.all_variables[detail] = condition

        if '整数' in event:
            self.data_manager.all_variables[detail] = int(self.data_manager.all_variables[detail])
        elif '小数' in event:
            self.data_manager.all_variables[detail] = float(self.data_manager.all_variables[detail])

    def _assign_string_variable(self, detail: str, condition: str) -> str:
        """声明字符串变量"""
        # 获取日期
        if condition.startswith('日期'):
            current_date = datetime.now()

            # 如果有天数差，计算相应的日期
            if '天' in condition:
                split = condition[4:-1]
                date_delta = int(split) if split.isdigit() else int(self.data_manager.all_variables[split])
                if '后' in condition:
                    current_date = current_date + timedelta(days=date_delta)
                elif '前' in condition:
                    current_date = current_date - timedelta(days=date_delta)

            if '/' in condition:
                date_format = current_date.strftime('%Y/%m/%d')
            elif '-' in condition:
                date_format = current_date.strftime('%Y-%m-%d')

            self.data_manager.all_variables[detail] = date_format

        # 读取 txt
        elif condition.startswith('txt'):
            txt_path = condition[4:]

            # 如果是变量，转换为变量中的内容
            if is_object(txt_path, target='变量'):
                txt_path = self.data_manager.all_variables[txt_path[3:]]

            # 如果文件不存在，执行结束
            if not os.path.exists(txt_path):
                self.insert_text(f'错误：文件{txt_path}不存在！\n')
                return 'error'

            with open(txt_path, 'r', encoding='gbk') as file:
                self.data_manager.all_variables[detail] = file.read()

        # 串联若干个数文字
        else:
            final_str = ''
            parts = condition.split('|')

            for part in parts:
                if is_object(part, target='变量'):
                    final_str += str(get_variable_content(part[3:], self.data_manager))
                elif is_object(part, target='数组'):
                    data_content, _ = get_data_content(part[3:], self.data_manager)
                    final_str += str(data_content)
                else:
                    final_str += part

            self.data_manager.all_variables[detail] = final_str

        return ''

    def _get_variable(self, detail: str, condition: str) -> str:
        """获取变量"""
        # 获取范围内的文字
        if condition.count('|') == 1 and not condition.startswith(('图片', '彩图')):
            parts = condition.split('|')
            image = capture_screen()
            l, u, r, d = get_variable(parts[1], self.data_manager)
            image = image[u:d, l:r]
            image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            results = self.readers[parts[0]].readtext(image)
            self.data_manager.all_variables[detail] = ''.join(text for (_, text, _) in results)

        # 获取图片或文字的位置
        else:
            find_result, print_text = get_position(condition, self.data_manager)

            if len(find_result) == 0:
                self.insert_text(f'错误：在给变量赋值时，找到的结果为空，无法赋值！\n')
                return 'error'

            parts = detail.split(',')
            self.data_manager.all_variables[parts[0]], self.data_manager.all_variables[parts[1]] = find_result

            if print_text == '图片':
                width, height = Image.open(condition.split('|')[0][3:]).size
                self.data_manager.all_variables[parts[0]] -= width // 2
                self.data_manager.all_variables[parts[1]] -= height // 2

        return ''

    def _calculate_variable(self, detail: str, condition: str) -> None:
        """计算变量"""
        parts = condition.split('|')

        if any(op in parts[0] for op in ('加', '减', '乘', '除')):
            number = int(parts[1]) if float(parts[1]).is_integer() else float(parts[1])

            if '加' in parts[0]:
                self.data_manager.all_variables[detail] += number
            elif '减' in parts[0]:
                self.data_manager.all_variables[detail] -= number
            elif '乘' in parts[0]:
                self.data_manager.all_variables[detail] *= number
            elif '除' in parts[0]:
                self.data_manager.all_variables[detail] /= number
