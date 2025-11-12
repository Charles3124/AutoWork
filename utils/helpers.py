"""
helpers.py

功能: 辅助函数模块
时间: 2025/10/17
版本: 1.0
"""

import time
import numpy as np
from typing import Optional

import cv2
import easyocr
import pyautogui
import pyperclip
import pandas as pd
from skimage.metrics import structural_similarity as ssim

from utils.funcs import capture_screen, find_text, find_image
from utils.data_manager import DataManager
from utils.types import TimeType, CommandType, RegionType, PositionType


readers = {
    '1': easyocr.Reader(lang_list=['ch_sim'], model_storage_directory='ocr_models'),
    '2': easyocr.Reader(lang_list=['en'], model_storage_directory='ocr_models'),
    '3': easyocr.Reader(lang_list=['ch_sim', 'en'], model_storage_directory='ocr_models')
}


def read_and_process_excel(target_excel_path: str) -> tuple[int, TimeType, CommandType, CommandType, CommandType]:
    """读取并预处理指令集"""
    df = pd.read_excel(target_excel_path).iloc[:, 0:4]
    df_length = len(df)

    # 处理表格内容
    for i in range(df_length):
        c0, c1, c2, c3 = df.iloc[i, 0:4].astype(str)
        df.iloc[i, 1:4] = c1, c2, c3

        if c0 == 'nan':            # 时间为空，转换为 -1.0
            df.iloc[i, 0] = -1.0
        if c1 == c2 == 'nan':      # 事件和细节都为空，合法
            continue

        df.iloc[i, 1] = c1.replace(' ', '')

    # 分割成 4 个字典
    temp_dic = {col: dict(zip(df.index, df[col])) for col in ['时间', '事件', '细节', '条件']}
    return df_length, temp_dic['时间'], temp_dic['事件'], temp_dic['细节'], temp_dic['条件']


def get_event_handlers(
        loop_controller, select_controller, wait_controller,
        data_reader, data_writer,
        file_controller, screen_shot_controller
):
    """创建事件分发器"""
    return {
        '循环': lambda i, event, detail, condition: loop_controller.handle_loop(i, event, detail),
        '选择': lambda i, event, detail, condition: select_controller.handle_select(i, event, detail, condition),
        '等待': lambda i, event, detail, condition: wait_controller.handle_wait(i, detail, condition),
        '读取': lambda i, event, detail, condition: (data_reader.handle_read(detail, condition), i + 1),
        '写入': lambda i, event, detail, condition: (data_writer.handle_write(detail, condition), i + 1),
        '文件': lambda i, event, detail, condition: (file_controller.handle_control(detail, condition), i + 1),
        '截图': lambda i, event, detail, condition: (screen_shot_controller.handle_shot(detail, condition), i + 1)
    }


def type_text(text: Optional[str] = None) -> None:
    """输入文字"""
    if text is not None:
        pyperclip.copy(text)
        time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')


def are_images_similar(image1: np.ndarray, image2: np.ndarray, threshold: str) -> bool:
    """比较图片是否相似"""
    threshold = 0.9 if threshold == '' else float(threshold)
    gray1 = cv2.cvtColor(image1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(image2, cv2.COLOR_BGR2GRAY)
    similarity_index, _ = ssim(gray1, gray2, full=True)
    return similarity_index >= threshold


def is_object(text: str, target: str) -> bool:
    """判断是否是某个对象"""
    return text.startswith((f'{target}:', f'{target}：'))


def get_position(content: str, data_manager: DataManager) -> tuple[PositionType, str]:
    """由条件获取坐标"""
    parts = content.split('|')
    region = define_region(parts[-1], data_manager)
    return find_image_or_text(parts, region, data_manager)


def define_region(content: str, data_manager: DataManager) -> RegionType:
    """判断范围类型"""
    # 没有给出筛选范围
    if content == '':
        return None

    # 筛选范围是像素元组
    if ',' in content:
        return get_variable(content, data_manager)

    # 筛选范围是字符串
    return content


def find_image_or_text(
        contents: list[str],
        region: RegionType,
        data_manager: DataManager
) -> tuple[PositionType, str]:
    """识别图片或文字"""
    if is_object(contents[0], target='图片'):
        return find_image(target=str(contents[0][3:]), image_path=None, region=region, gray=True), '图片'
    
    if is_object(contents[0], target='彩图'):
        return find_image(target=str(contents[0][3:]), image_path=None, region=region, gray=False), '彩图'
    
    reader = readers[contents[1]]

    if is_object(contents[0], target='数组'):
        text, _ = get_data_content(contents[0][3:], data_manager)
    elif is_object(contents[0], target='变量'):
        text = str(get_variable_content(contents[0][3:], data_manager))
    else:
        text = contents[0]
    
    return find_text(target=text, image_path=None, region=region, reader=reader), '文字'


def get_image(image: str, data_manager: DataManager) -> np.ndarray:
    """获取图片"""
    # 在屏幕上某个范围内截图
    if ',' in image:
        l, u, r, d = get_variable(image, data_manager)
        image = capture_screen()
        return image[u:d, l:r]

    # 读取一张图片
    if is_object(image, target='图片'):
        return cv2.imread(image[3:])

    # 使用之前定义的图片变量
    return data_manager.all_figures[image]


def get_data_content(text: str, data_manager: DataManager) -> tuple[str, str]:
    """获取数组的内容"""
    if not all(char in text for char in ['(', ':', ')']):
        data_index = data_manager.all_data_index[text]
        return data_manager.all_data[text][data_index], text
    
    parts = text.split('(')
    left, right = get_split_range(parts[1][:-1])

    data_name = parts[0]
    data_index = data_manager.all_data_index[data_name]
    text = data_manager.all_data[data_name][data_index]

    return text[left:right], data_name


def get_split_range(split_range: str) -> tuple[Optional[int], Optional[int]]:
    """获取文字切分范围"""
    left, right = split_range.split(':')
    left = int(left) if left != '' else None
    right = int(right) if right != '' else None
    return left, right


def get_variable_content(text: str, data_manager: DataManager) -> str:
    """获取变量的内容"""
    if not all(char in text for char in ['(', ':', ')']):
        return data_manager.all_variables[text]
    parts = text.split('(')
    left, right = get_split_range(parts[1][:-1])
    text = data_manager.all_variables[parts[0]]
    return text[left:right]


def get_variable(contents: str, data_manager: DataManager) -> int | tuple[int, ...]:
    """把一组由“,”隔开的内容转换成数字元组"""
    if ',' not in contents:
        return int(contents) if contents.isdigit() else int(data_manager.all_variables[contents])

    return tuple(
        int(part) if part.isdigit() else int(data_manager.all_variables[part])
        for part in contents.split(',')
    )
