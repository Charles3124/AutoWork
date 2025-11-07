"""
recur.py

功能: 根据 Excel 指令集自动化操作电脑
时间: 2025/10/17
版本: 1.0
"""

import os
import time
import threading
from threading import Event
from typing import Optional, Union

import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from playsound import playsound
from pynput.mouse import Controller as MController
from pynput.keyboard import Controller as KController, Key, KeyCode, Listener as KeyboardListener

from utils.helpers import readers, read_and_process_excel, get_event_handlers
from utils.data_manager import DataManager
from utils.colors import *
from flow_control import LoopController, SelectController, WaitController
from data_io import DataReader, DataWriter
from file_manager import FileController
from image_control import ScreenShotController
from device_control import MouseController, KeyboardController
from data_control import VariableController


# ------- 准备阶段 -------
mouse = MController()       # 创建鼠标控制器
keyboard = KController()    # 创建键盘控制器
pause_event = Event()       # 键盘监听器
pause_event.set()           # 默认不暂停
exit_flag = False           # 初始化退出标志
rectangle_color = GREEN     # 矩形颜色
min_time_gap = 0.1          # 最小执行间距


def on_press(key: Union[Key, KeyCode]) -> Optional[bool]:
    """监听键盘按键函数"""
    global exit_flag

    if key == Key.esc:
        insert_text(f'控制台：Esc 按下，等待程序退出...\n')
        exit_flag = True    # 退出标志
        return False        # 返回 False 停止监听器

    if key == Key.space:
        # 如果当前是运行状态，则暂停
        if pause_event.is_set():
            insert_text(f'控制台：程序暂停\n')
            pause_event.clear()

        # 如果当前是暂停状态，则继续
        else:
            insert_text(f'控制台：程序继续\n')
            pause_event.set()

    return None


def insert_text(text: str) -> None:
    """实时显示进程"""
    info_text.tag_configure('big', font=('宋体', 14), spacing3=7)
    info_text.insert(tk.END, text, 'big')
    info_text.yview(tk.END)


# ------- 主程序 -------
def run(given_excel_path: Optional[str] = None) -> None:
    """读取 Excel 指令并依次执行"""
    # 设置置顶状态
    if checkbox_var.get():
        root.attributes('-topmost', False)
    else:
        root.attributes('-topmost', True)

    # 检测输入路径是否存在
    target_excel_path = excel_combobox.get() if given_excel_path is None else given_excel_path
    if not os.path.exists(target_excel_path):
        insert_text(f'给定的 Excel 路径不存在！\n')
        return

    # 设置执行状态
    insert_text(f'执行开始！\n')
    global exit_flag
    global rectangle_color
    rectangle_color = RED
    initialize_rectangle(rectangle_color)

    # 启动键盘监听器
    keyboard_listener = KeyboardListener(on_press=on_press)
    keyboard_listener.start()

    # 获取指令集信息
    df_length, times, events, details, conditions = read_and_process_excel(target_excel_path)

    # 创建所有控制器
    data_manager = DataManager()

    loop_controller = LoopController(events, details, data_manager, insert_text)
    select_controller = SelectController(events, details, data_manager, insert_text)
    wait_controller = WaitController(df_length, events, details, data_manager, insert_text)

    data_reader = DataReader(data_manager, insert_text)
    data_writer = DataWriter(data_manager, insert_text)
    data_controller = VariableController(readers, data_manager, insert_text)

    file_controller = FileController(data_manager, insert_text)
    screen_shot_controller = ScreenShotController(data_manager, insert_text)

    mouse_controller = MouseController(mouse, data_manager, insert_text)
    keyboard_controller = KeyboardController(keyboard, data_manager, insert_text)

    handlers = get_event_handlers(
        loop_controller, select_controller, wait_controller,
        data_reader, data_writer, file_controller, screen_shot_controller
    )

    # 主循环开始
    i = 0
    time.sleep(max(times[0], min_time_gap))

    while i < df_length:
        if exit_flag:
            insert_text(f'控制台：程序已结束！\n')
            break

        # 如果暂停，等待直到恢复
        pause_event.wait()

        # 读取时间、事件、细节、条件
        cur_time: float = times[i]
        event: str = events[i]
        detail: str = details[i]
        condition: Optional[str] = conditions[i]

        if event == detail == 'nan':
            i += 1
            continue

        # 等待时间差
        if i > 0:
            if cur_time == -1.0 or times[i - 1] == -1.0:
                time.sleep(0.1)
            else:
                time.sleep(max(cur_time - times[i - 1], min_time_gap))

        # 判断是哪种指令
        matched = False
        for key, func in handlers.items():
            if key in event:
                result = func(i, event, detail, condition)
                if isinstance(result, tuple):
                    i = result[1]
                else:
                    i = result
                matched = True
                break

        if matched:
            continue

        if '鼠标' in event:
            mouse_result = mouse_controller.handle_mouse(event, detail, condition)
            if mouse_result == 'error':
                break

        elif '输入' in event or '按键' in event:
            keyboard_controller.handle_keyboard(event, detail)

        elif '变量' in event or '整数' in event or '小数' in event or '文字' in event:
            variable_result = data_controller.handle_variable(event, detail, condition)
            if variable_result == 'error':
                break

        elif '数组' in event:
            data_manager.all_data_index[detail] += int(condition)

        elif '提示音' in event:
            playsound(f'sounds/{detail}.wav')

        elif 'i' in event:
            i = int(detail) - 3
            insert_text(f'执行第{detail}行指令\n')

        else:
            insert_text(f'没看出这一行是什么指令呢，跳过！\n')
            i += 1
            continue

        insert_text(f'时间：{cur_time if cur_time != -1.0 else ' '} 执行：{event} {detail}\n')
        i += 1

    keyboard_listener.stop()
    exit_flag = False
    rectangle_color = GREEN
    initialize_rectangle(rectangle_color)
    insert_text(f'执行完毕！\n\n')


# ------- 创建 GUI 界面 -------
root = tk.Tk()
root.title('自动化办公系统')
gui_width, gui_height = 550, 480

rate = 1.25
gui_width = int(rate * gui_width)
gui_height = int(rate * gui_height)

root.geometry(f'{gui_width}x{gui_height}')
root.resizable(width=False, height=False)

rect_width, rect_height = 150, 40
button_font = ('宋体', 14)

# ------- 加载背景 -------
image = Image.open('backgrounds/1.png')
image = image.resize((gui_width, gui_height), Image.Resampling.LANCZOS)
photo = ImageTk.PhotoImage(image)

# 创建 Canvas 组件，设置为窗口的背景
canvas = tk.Canvas(root, width=image.width, height=image.height)
canvas.pack(fill='both', expand=True)
canvas.create_image(0, 0, anchor=tk.NW, image=photo)

# 处理窗口大小变化
def resize_image(event: tk.Event) -> None:
    width, height = root.winfo_width(), root.winfo_height()                   # 获取窗口大小
    resized_image = image.resize((width, height), Image.Resampling.LANCZOS)   # 根据新的大小调整图片
    resized_photo = ImageTk.PhotoImage(resized_image)                         # 转换成 Canvas 可以显示的图片格式
    canvas.delete('background')                                               # 删除旧背景
    canvas.create_image(0, 0, anchor=tk.NW, image=resized_photo)        # 在 Canvas 上更新背景图片
    canvas.image = resized_photo                                              # 更新 photo 防止垃圾回收
    initialize_rectangle(rectangle_color)                                     # 初始化矩形颜色

# 绑定窗口大小变化事件
root.bind('<Configure>', resize_image)

def initialize_rectangle(color: str) -> None:
    """生成矩形"""
    rect = canvas.create_rectangle(
        (gui_width - rect_width) // 2, 30,
        (gui_width + rect_width) // 2, 30 + rect_height,
        fill=color, tags='rectangle'
    )

# ------- Excel 路径输入框 -------
def get_excel_files() -> Optional[list[str]]:
    """获取 my_programs 文件夹中的 Excel 文件列表"""
    excel_files = []
    my_programs_path = os.path.join(os.getcwd(), 'my_programs')
    if os.path.exists(my_programs_path):
        for file in os.listdir(my_programs_path):
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(f'my_programs/{file}')
    return excel_files

def update_excel_dropdown() -> None:
    """初始化下拉菜单选项"""
    excel_files = get_excel_files()
    excel_combobox['values'] = excel_files

def on_excel_selected(event: tk.Event) -> None:
    """当选择下拉菜单项时的处理函数"""
    selected = excel_combobox.get()
    excel_combobox.set(selected)

excel_combobox = ttk.Combobox(root, width=43, font=button_font)
excel_combobox.place(relx=0.5, y=90, anchor='n')
update_excel_dropdown()
excel_combobox.bind('<<ComboboxSelected>>', on_excel_selected)

# ------- 运行按钮 -------
def run_in_thread() -> None:
    """创建子进程"""
    threading.Thread(target=run).start()

save_button = tk.Button(root, text='运行', command=run_in_thread, font=button_font)
save_button.place(relx=0.55, y=140, anchor='n')

# ------- 清除按钮 -------
def clear() -> None:
    """清楚输入框"""
    excel_combobox.set('')

clear_button = tk.Button(root, text='清除', command=clear, font=button_font)
clear_button.place(relx=0.7, y=140, anchor='n')

# ------- 复选框 -------
checkbox_var = tk.BooleanVar()
checkbox = tk.Checkbutton(root, text='运行后隐藏界面', variable=checkbox_var, font=button_font)
checkbox.place(relx=0.2, y=142, anchor='n')

# ------- 信息输出框 -------
info_text = tk.Text(root, height=16, width=60)
info_text.place(relx=0.5, y=205, anchor='n')

# ------- 启动 GUI -------
root.mainloop()
