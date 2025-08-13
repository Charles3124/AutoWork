import os
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

RegionType = tuple[int, int, int, int] | str | None
PositionType = tuple[int, int] | tuple[()]

readers = {
    '1': easyocr.Reader(['ch_sim'], model_storage_directory='ocr_models'),
    '2': easyocr.Reader(['en'], model_storage_directory='ocr_models'),
    '3': easyocr.Reader(['ch_sim', 'en'], model_storage_directory='ocr_models')
}


def type_text(text: Optional[str] = None) -> None:
    """输入文字"""
    if text != None:
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


def get_position(content: str, all_data: dict, all_data_index: dict, all_variables: dict) -> tuple[PositionType, str]:
    """由条件获取坐标"""
    parts = content.split('|')
    region = define_region(parts[-1], all_variables)
    return find_image_or_text(parts, region, all_data, all_data_index, all_variables)


def define_region(content: str, all_variables: dict) -> RegionType:
    """判断范围"""
    if content == '':       # 没有给出筛选范围
        return None
    if ',' in content:      # 筛选范围是像素元组
        return get_variable(content, all_variables)
    return content          # 筛选范围是字符串


def find_image_or_text(contents: list[str], region: RegionType, all_data: dict, all_data_index: dict, all_variables: dict) -> tuple[PositionType, str]:
    """识别图片或文字"""
    if is_object(contents[0], '图片'):
        return find_image(target=str(contents[0][3:]), image_path=None, region=region, gray=True), "图片"
    
    elif is_object(contents[0], '彩图'):
        return find_image(target=str(contents[0][3:]), image_path=None, region=region, gray=False), "图片"
    
    reader = readers[contents[1]]

    if is_object(contents[0], '数组'):
        text, _ = get_data_content(contents[0][3:], all_data, all_data_index)
    elif is_object(contents[0], '变量'):
        text = str(get_variable_content(contents[0][3:], all_variables))
    else:
        text = contents[0]
    
    return find_text(target=text, image_path=None, region=region, reader=reader), "文字"


def get_image(image: str, all_figures: dict, all_variables: dict) -> np.ndarray:
    """获取图片"""
    if ',' in image:                   # 在屏幕上某个范围内截图
        l, u, r, d = get_variable(image, all_variables)
        image = capture_screen()
        return image[u:d, l:r]
    if is_object(image, '图片'):       # 读取一张图片
        return cv2.imread(image[3:])
    return all_figures[image]          # 使用之前定义的图片变量


def get_data_content(text: str, all_data: dict, all_data_index: dict) -> tuple[str, str]:
    """获取数组的内容"""
    if not all(char in text for char in ['(', ':', ')']):
        data_index = all_data_index[text]
        return all_data[text][data_index], text
    
    parts = text.split('(')
    left, right = get_split_range(parts[1][:-1])

    data_name = parts[0]
    data_index = all_data_index[data_name]
    text = all_data[data_name][data_index]

    return text[left:right], data_name


def get_split_range(range: str) -> tuple[Optional[int], Optional[int]]:
    """获取文字切分范围"""
    left, right = range.split(':')
    left = int(left) if left != '' else None
    right = int(right) if right != '' else None
    return left, right


def get_variable_content(text: str, all_variables: dict) -> str:
    """获取变量的内容"""
    if not all(char in text for char in ['(', ':', ')']):
        return all_variables[text]
    parts = text.split('(')
    left, right = get_split_range(parts[1][:-1])
    text = all_variables[parts[0]]
    return text[left:right]


def get_variable(contents: str, all_variables: dict) -> int | tuple[int]:
    """把一组由“,”隔开的内容转换成数字元组"""
    if ',' not in contents:
        return int(contents) if contents.isdigit() else int(all_variables[contents])
    return tuple(int(part) if part.isdigit() else int(all_variables[part]) for part in contents.split(','))


def judge_condition(condition: str, all_data: dict, all_data_index: dict, all_variables: dict, all_figures: dict) -> bool:
    """判断选择条件是否成立"""
    if condition.startswith('存在文件'):      # 文件是否存在
        file_name = condition[5:]
        return os.path.exists(str(get_variable_content(file_name[3:], all_variables))) if is_object(file_name, '变量') else os.path.exists(file_name)

    if is_object(condition, '相似'):
        part1, part2, threshold = condition[3:].split('|')
        image1, image2 = get_image(part1, all_figures, all_variables), get_image(part2, all_figures, all_variables)
        return are_images_similar(image1, image2, threshold)

    if '在数组' in condition:                 # 文字或变量是否在数组中
        parts = condition.split('在数组')
        data_name = parts[1][1:]
        target = str(get_variable_content(parts[0][3:], all_variables)) if is_object(parts[0], '变量') else target == parts[0]

        for content in all_data[data_name]:
            if not pd.isna(content) and target == str(content):
                return True
        return False

    if is_object(condition, '是否相同'):      # 内容是否相同
        parts = condition[5:].split('|')
        results = []

        for part in parts:
            if is_object(part, '数组'):
                data_content, _ = get_data_content(part[3:], all_data, all_data_index)
                results.append(str(data_content))
            elif is_object(part, '变量'):
                variable_content = get_variable_content(part[3:], all_variables)
                results.append(str(variable_content))
            else:
                results.append(part)
        return all(r == results[0] for r in results)
    
    if '|' in condition:                     # 图片或文字是否在指定范围
        find_result, _ = get_position(condition, all_data, all_data_index, all_variables)
        return len(find_result) != 0
    
    return False     # 默认返回 False
