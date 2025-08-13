# funcs.py
import cv2
import easyocr
import pyautogui
import numpy as np
from screeninfo import get_monitors
from typing import Optional

RegionType = tuple[int, int, int, int] | str | None
PositionType = tuple[int, int] | tuple[()]

# 裁剪区域
CROP_REGION = {
    "left": (0, 0.5, 0, 1),
    "right": (0.5, 1, 0, 1),
    "up": (0, 1, 0, 0.5),
    "down": (0, 1, 0.5, 1),
    "q1": (0, 0.5, 0, 0.5),
    "q2": (0.5, 1, 0, 0.5),
    "q3": (0, 0.5, 0.5, 1),
    "q4": (0.5, 1, 0.5, 1)
}

def capture_screen() -> np.ndarray:
    """
    捕捉整个屏幕的截图并返回 BGR 格式的图像
    """
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)
    screenshot_bgr = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2BGR)
    return screenshot_bgr

def capture_screen_by_monitor(monitor_index: int | None = None) -> np.ndarray:
    """
    捕捉整个屏幕的截图或指定显示器的截图并返回 BGR 格式的图像。
    :param monitor_index: 指定显示器的索引，默认为 None 表示截取所有显示器的屏幕
    :return: BGR 格式的图像
    """
    monitors = get_monitors()

    if monitor_index is None:
        # 截取所有显示器的屏幕
        screenshot = pyautogui.screenshot()
    else:
        # 截取指定显示器的屏幕
        monitor = monitors[monitor_index]
        screenshot = pyautogui.screenshot(region=(monitor.x, monitor.y, monitor.width, monitor.height))
    
    screenshot_np = np.array(screenshot)
    screenshot_bgr = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2BGR)
    return screenshot_bgr

def find_text(
    target: str,
    image_path: Optional[str] = None,
    region: RegionType = None,
    reader: Optional[easyocr.Reader] = None
) -> PositionType:
    """
    在图像中寻找指定的文字，并返回第一个符合的文字的位置。
    :param target: 目标文字
    :param image_path: 图片路径，如果为 None，则进行屏幕截图
    :param region: 区域范围，字符串或像素值的元组
    :param reader: easyocr 读取器，默认为 None，自动创建
    :return: 返回文字位置的元组 (x, y)，如果未找到文字，则返回空元组
    """
    # 确定语言
    if reader is None:
        reader = easyocr.Reader(['ch_sim', 'en'], model_storage_directory=r'ocr_models')
    
    # 确定图像
    if image_path is None:   # 对屏幕截图
        image = capture_screen()
    else:                    # 读取图片
        image = cv2.imread(image_path)

    # 确定范围
    if region is not None:              # 如果有指定范围
        if isinstance(region, str):     # 范围是一个字符串
            height, width, _ = image.shape
            crop_para = CROP_REGION[region]
            l, r, u, d = [int(ratio * length) for ratio, length in zip(crop_para, (width, width, height, height))]
        else:                           # 范围是像素值
            l, u, r, d = region         # region 格式是 (x1, y1, x2, y2)
        image = image[u:d, l:r]
        res_x, res_y = l, u
    else:
        res_x, res_y = 0, 0
    
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)    # 转换为灰度图
    results = reader.readtext(gray_image)                   # 寻找文字

    # 返回匹配位置
    for (bbox, text, _) in results:
        if target in text:              # 找到就返回结果
            res_x += (bbox[0][0] + bbox[2][0]) // 2
            res_y += (bbox[0][1] + bbox[2][1]) // 2
            return (res_x, res_y)
    return ()

def find_image(
    target: str,
    image_path: Optional[str] = None,
    region: RegionType = None,
    gray: bool = True
) -> PositionType:
    """
    在图像中寻找指定模板的位置，并返回第一个符合的图片的中心坐标。
    :param target: 模板图片路径
    :param image_path: 图片路径，如果为 None，则进行屏幕截图
    :param region: 区域范围，字符串或像素值的元组
    :param gray: 读取彩图或者灰度图
    :return: 返回匹配位置的元组 (x, y)，如果未找到匹配项，则返回空元组
    """
    # 读取模板图片
    if gray:
        template = cv2.imread(target, cv2.IMREAD_GRAYSCALE)
        h, w = template.shape
    else:
        template = cv2.imread(target)
        h, w, _ = template.shape
    
    # 确定图像
    if image_path is None:   # 对屏幕截图
        image = capture_screen()
    else:                    # 读取图片
        image = cv2.imread(image_path)

    # 确定范围
    if region is not None:              # 如果有指定范围
        if isinstance(region, str):     # 范围是一个字符串
            height, width, _ = image.shape
            crop_para = CROP_REGION[region]
            l, r, u, d = [int(ratio * length) for ratio, length in zip(crop_para, (width, width, height, height))]
        else:                           # 范围是像素值
            l, u, r, d = region         # region 格式是 (x1, y1, x2, y2)
        image = image[u:d, l:r]
        res_x, res_y = l, u
    else:
        res_x, res_y = 0, 0

    if gray:
        gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)                     # 转换为灰度图
        result = cv2.matchTemplate(gray_image, template, cv2.TM_CCOEFF_NORMED)   # 使用模板匹配
    else:
        result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)        # 使用模板匹配
    
    threshold = 0.9                                                              # 设置匹配阈值
    loc = np.where(result >= threshold)

    # 返回匹配位置
    for pt in zip(*loc[::-1]):  # zip(*loc[::-1]) 获取匹配到的左上角坐标
        res_x += pt[0] + w // 2
        res_y += pt[1] + h // 2
        return (res_x, res_y)
    return ()
