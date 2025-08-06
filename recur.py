import os
import time
import shutil
import threading
from threading import Event
from datetime import datetime, timedelta

import cv2
import pandas as pd
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from playsound import playsound
from openpyxl import load_workbook
from pynput.mouse import Controller as MouseController, Button
from pynput.keyboard import Controller as KeyboardController, Key, Listener as KeyboardListener

from utils.funcs import capture_screen
from utils.mappings import button_mapping, key_mapping
from utils.helpers import *


# ----------准备阶段----------
mouse = MouseController()         # 创建鼠标控制器
keyboard = KeyboardController()   # 创建键盘控制器
pause_event = Event()             # 键盘监听器
pause_event.set()                 # 默认不暂停
exit_flag = False                 # 初始化退出标志
rectangle_color = 'green'         # 矩形颜色


# ----------监听键盘按键函数----------
def on_press(key):
    global exit_flag
    if key == Key.esc:
        insert_text(f'控制台：Esc 按下，等待程序退出...\n')
        exit_flag = True           # 退出标志
        return False               # 返回 False 停止监听器
    elif key == Key.space:
        if pause_event.is_set():   # 如果当前是运行状态，则暂停
            insert_text(f'控制台：程序暂停\n')
            pause_event.clear()
        else:                      # 如果当前是暂停状态，则继续
            insert_text(f'控制台：程序继续\n')
            pause_event.set()


# ----------实时显示进程----------
def insert_text(text: str):
    info_text.tag_configure("big", font=("宋体", 14), spacing3=7)
    info_text.insert(tk.END, text, "big")
    info_text.yview(tk.END)


# ----------主程序----------
def run(given_excel_path=None):

    # 设置置顶状态
    if checkbox_var.get():
        root.attributes('-topmost', False)
    else:
        root.attributes('-topmost', True)

    # 读取 Excel 中的指令
    target_excel_path = excel_combobox.get() if given_excel_path == None else given_excel_path
    if not os.path.exists(target_excel_path):
        insert_text(f'给定的Excel路径不存在！\n')
        return
    df = pd.read_excel(target_excel_path).iloc[:, 0:4]
    df_length = len(df)

    # 设置执行状态
    insert_text(f'执行开始！\n')
    global exit_flag
    global rectangle_color
    rectangle_color = 'red'
    initialize_rectangle(rectangle_color)

    # 启动键盘监听器
    keyboard_listener = KeyboardListener(on_press=on_press)
    keyboard_listener.start()

    # 处理表格内容
    for i in range(df_length):
        c0, c1, c2, c3 = df.iloc[i, 0:4].astype(str)
        df.iloc[i, 1:4] = c1, c2, c3

        if c0 == 'nan':            # 时间为空，转换为 -1.0
            df.iloc[i, 0] = -1.0
        if c1 == c2 == 'nan':      # 事件和细节都为空，合法
            continue
        df.iloc[i, 1] = c1.replace(' ', '')

    temp_dic = {col: dict(zip(df.index, df[col])) for col in ['时间', '事件', '细节', '条件']}                    # 转换成字典
    times, events, details, conditions = temp_dic['时间'], temp_dic['事件'], temp_dic['细节'], temp_dic['条件']   # 分割成 4 个字典

    # 主循环开始
    i = 0
    l_info, j_info = [], []
    all_data, all_data_index, all_variables, all_figures = {}, {}, {}, {}
    min_time_gap = 0.1
    time.sleep(max(times[0], min_time_gap))

    while i < df_length:
        if exit_flag:
            insert_text(f'控制台：程序已结束！\n')
            break
        pause_event.wait()   # 如果暂停，等待直到恢复

        # 读取时间、事件、细节、条件
        cur_time, event, detail, condition = times[i], events[i], details[i], conditions[i]
        if event == detail == 'nan':
            i += 1
            continue

        # 等待时间差
        if i > 0:
            if cur_time == -1.0 or times[i - 1] == -1.0:
                time.sleep(0.1)
            else:
                time.sleep(max(cur_time - times[i - 1], min_time_gap))

        if "循环" in event:
            if '退出' in detail:
                last_info = l_info[-1]
                i = last_info['end']
                insert_text(f'直接退出第{len(l_info)}层循环！\n')
                l_info.pop()

            elif "}" in detail or "继续" in detail:
                last_info = l_info[-1]

                if last_info['times'] is not None:     # 次数循环
                    last_info['times'] -= 1

                    if last_info['times'] > 0:
                        insert_text(f'第{len(l_info)}层循环还有{last_info['times']}次\n')
                        i = last_info['begin']
                    else:
                        i = last_info['end']
                        insert_text(f'第{len(l_info)}层循环结束！\n')
                        l_info.pop()
                else:                                  # 数组循环
                    data_name = last_info['name']
                    data_index = all_data_index[data_name]

                    if data_index < len(all_data[data_name]):
                        insert_text(f'循环继续\n')
                        i = last_info['begin']
                    else:
                        i = last_info['end']
                        insert_text(f'第{len(l_info)}层循环结束！\n')
                        l_info.pop()
            
            else:
                loop_name = event.replace('{', '')
                end = i
                while loop_name not in events[end] or '}' not in details[end]:
                    end += 1
                l_info.append({'begin': i, 'end': end, 'times': None, 'name': None})
                
                if detail.isdigit():     # 次数循环
                    l_info[-1]['times'] = int(detail)
                    insert_text(f'第{len(l_info)}层循环开始，共{detail}次...\n')
                else:                    # 数组循环
                    l_info[-1]['name'] = detail[3:]
                    insert_text(f'第{len(l_info)}层循环开始，数组名为{l_info[-1]['name']}...\n')
            i += 1
            continue

        if "选择" in event:
            if "{" in event:      # 选择开始
                j_info.append([event.replace('{', ''), False])
                insert_text(f'第{len(j_info)}层选择开始...\n')

            if "}" in detail:     # 选择结束
                insert_text(f'第{len(j_info)}层选择结束！\n')
                j_info.pop()

            elif j_info[-1][1]:   # 已经执行过其他分支，跳过
                while j_info[-1][0] not in events[i + 1] or "}" not in details[i + 1]:
                    i += 1
                insert_text(f'跳过剩余分支\n')

            else:                 # 开始判断
                insert_text(f'开始判断...')
                judgement_result = True if "其他" in detail else judge_condition(condition, all_data, all_data_index, all_variables, all_figures)

                if judgement_result:
                    insert_text(f'判断为真\n')
                    j_info[-1][1] = True
                else:
                    insert_text(f'判断为假\n')
                    while j_info[-1][0] not in events[i + 1]:
                        i += 1
            i += 1
            continue

        if '读取' in event:
            if 'excel' in detail.lower():
                parts = condition.split('|')
                if not parts[1].endswith('.xlsx'):
                    parts[1] += '.xlsx'
                temp_df = pd.read_excel(parts[1], sheet_name=parts[2], usecols=[parts[3]], dtype=str)

                last_non_empty_idx = temp_df.last_valid_index()
                temp_df = temp_df[:last_non_empty_idx + 1]

                all_data[parts[0]] = list(temp_df[parts[3]].tolist())
                all_data_index[parts[0]] = 0
                insert_text(f'Excel读取成功！\n')
            i += 1
            continue

        if '写入' in event:
            if 'excel' in detail.lower():
                parts = condition.split('|')
                wb = load_workbook(parts[1])
                sheet = wb[parts[2]]
                
                if is_object(parts[0], '数组'):
                    text, _ = get_data_content(parts[0][3:], all_data, all_data_index)
                elif is_object(parts[0], '变量'):
                    text = get_variable_content(parts[0][3:], all_variables)
                else:
                    text = parts[0]
                
                target_cell = str(all_variables[parts[3][3:]]) if is_object(parts[3], '变量') else parts[3]

                sheet[target_cell] = str(text)
                wb.save(parts[1])
                insert_text(f'内容“{text}”已写入Excel！\n')
            i += 1
            continue

        if '文件' in event:
            if any(op in detail for op in ['复制', '移动']):          # 移动或复制文件
                source_file, target_folder = condition.split('|')
                if os.path.exists(source_file):                      # 文件存在
                    file_name = os.path.basename(source_file)        # 获取文件名
                    copy_target = os.path.join(target_folder, file_name)
                    if '复制' in detail:
                        shutil.copy(source_file, copy_target)
                        insert_text(f'成功复制文件到{copy_target}\n')
                    elif '移动' in detail:
                        shutil.move(source_file, copy_target)
                        insert_text(f'成功移动文件到{copy_target}\n')
                else:                                                # 文件不存在
                    insert_text(f'文件{source_file}不存在，无法复制或移动！\n')
            
            elif '删除' in detail:                                    # 删除文件
                if os.path.exists(condition):
                    os.remove(condition)
                    insert_text(f'文件{condition}已被删除！\n')
                else:
                    insert_text(f'文件{condition}不存在，无法删除！\n')
            
            elif '写入' in detail:                                    # 生成 txt 文件
                if 'txt' in detail:
                    file_path, name = condition.split('|')
                    with open(file_path, 'a') as file:                # 以追加模式打开文件
                        file.write(f'{all_variables[name[3:]]}\n')    # 添加内容并换行
                    insert_text(f'已将内容写入txt文件\n')
            
            i += 1
            continue

        if '截图' in event:
            image = capture_screen()
            if condition != 'nan':
                l, u, r, d = get_variable(condition, all_variables)
                image = image[u:d, l:r]
            all_figures[detail] = image
            insert_text(f'截图成功！\n')
            i += 1
            continue

        if '等待' in event:
            insert_text(f'开始等待...\n')
            start_time = time.time()

            parts = detail.split('|')
            negative_flag = '0' if '不' in parts[0] else '1'
            gap_time, max_time = map(float, parts[1].split(','))
            normal_break = True

            if '相似' in parts[0]:
                part1, part2, threshold = condition.split('|')

            while True:
                time.sleep(gap_time)

                if '直到' in parts[0]:         # 是否存在文字或图片
                    find_result, _ = get_position(condition, all_data, all_data_index, all_variables)
                    condition_map = {'0': len(find_result) == 0, '1': len(find_result) != 0}

                elif '相似' in parts[0]:       # 两张图片是否相似
                    image1, image2 = get_image(part1, all_figures, all_variables), get_image(part2, all_figures, all_variables)
                    similar = are_images_similar(image1, image2, threshold)
                    condition_map = {'0': not similar, '1': similar}

                end_time = time.time()
                if condition_map[negative_flag]:
                    break
                if end_time - start_time >= max_time:
                    normal_break = False
                    break

            insert_text(f'等待{'结束' if normal_break else '超时'}！共等了{end_time - start_time:.2f}秒\n')

            if i < df_length - 1 and '超时' in events[i + 1]:
                if normal_break:    # 正常结束
                    while '超时' not in events[i] or '}' not in details[i]:
                        i += 1
                else:               # 超时结束
                    i += 1
            i += 1
            continue

        if '鼠标' in event:
            if '平移' in event:
                distance =  int(all_variables[condition[3:]]) if is_object(condition, '变量') else float(condition)
                move_map = {"上": (0, -distance), "下": (0, distance), "左": (-distance, 0), "右": (distance, 0)}
                x, y = mouse.position
                dx, dy = move_map.get(detail, (0, 0))
                mouse.position = (x + dx, y + dy)
            
            else:
                if condition != 'nan':
                    if '|' not in condition:         # 给定坐标点
                        x, y = get_variable(condition, all_variables)
                    else:                            # 寻找图片或文字的位置
                        find_result, print_text = get_position(condition, all_data, all_data_index, all_variables)
                        if len(find_result) == 0:    # 没找到结果，直接退出操作
                            insert_text(f"错误：没找到指定的{print_text}，请检查指定的{print_text}和查找范围！\n")
                            break
                        x, y = find_result
                    mouse.position = (x, y)
                if '单击' in event:
                    mouse.click(button_mapping.get(detail))
                if '双击' in event:
                    mouse.click(button_mapping.get(detail)); mouse.click(button_mapping.get(detail))

        elif '输入' in event:
            if detail == '剪贴板':
                type_text()
            elif is_object(detail, '数组'):
                parts = detail.split('|')
                text, data_name = get_data_content(parts[0][3:], all_data, all_data_index)
                type_text(text)
                all_data_index[data_name] += int(parts[-1]) if parts[-1] != '' else 0
            elif is_object(detail, '变量'):
                type_text(get_variable_content(detail[3:], all_variables))
            else:
                type_text(detail)

        elif '按键' in event:
            if '按下' in event:
                if detail in key_mapping:
                    keyboard.press(key_mapping.get(detail))
                else:
                    keyboard.press(detail)
            else:
                if detail in key_mapping:
                    keyboard.release(key_mapping.get(detail))
                else:
                    keyboard.release(detail)

        elif '变量' in event:
            if '获取' in event:
                if condition.count('|') == 1 and not condition.startswith('图片') and not condition.startswith('彩图'):      # 获取范围内的文字
                    parts = condition.split('|')
                    image = capture_screen()
                    l, u, r, d = get_variable(parts[1], all_variables)
                    image = image[u:d, l:r]
                    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                    results = readers[parts[0]].readtext(image)
                    all_variables[detail] = ''.join(text for (_, text, _) in results)
                
                else:          # 获取图片或文字的位置
                    find_result, print_text = get_position(condition, all_data, all_data_index, all_variables)

                    if len(find_result) == 0:
                        insert_text(f'错误：在给变量赋值时，找到的结果为空，无法赋值！\n')
                        break
                    parts = detail.split(',')
                    all_variables[parts[0]], all_variables[parts[1]] = find_result

                    if print_text == '图片':
                        width, height = Image.open(condition.split('|')[0][3:]).size
                        all_variables[parts[0]] -= width // 2
                        all_variables[parts[1]] -= height // 2
            
            elif '运算' in event:              # 变量运算
                parts = condition.split('|')
                if any(op in parts[0] for op in ['加', '减', '乘', '除']):
                    number = int(parts[1]) if float(parts[1]).is_integer() else float(parts[1])
                    if   '加' in parts[0]: all_variables[detail] += number
                    elif '减' in parts[0]: all_variables[detail] -= number
                    elif '乘' in parts[0]: all_variables[detail] *= number
                    elif '除' in parts[0]: all_variables[detail] /= number

        elif any(char in event for char in ['整数', '小数']):
            if is_object(condition, '变量'):
                all_variables[detail] = get_variable_content(condition[3:], all_variables)
            elif is_object(condition, '数组'):
                data_content, _ = get_data_content(condition[3:], all_data, all_data_index)
                all_variables[detail] = data_content
            else:
                all_variables[detail] = condition
            
            if '整数' in event:
                all_variables[detail] = int(all_variables[detail])
            elif '小数' in event:
                all_variables[detail] = float(all_variables[detail])

        elif '文字' in event:
            if condition.startswith('日期'):       # 获取日期
                current_date = datetime.now()
                if '天' in condition:              # 如果有天数差，计算相应的日期
                    split = condition[4:-1]
                    date_delta = int(split) if split.isdigit() else int(all_variables[split])
                    if '后' in condition:
                        current_date = current_date + timedelta(days=date_delta)
                    elif '前' in condition:
                        current_date = current_date - timedelta(days=date_delta)
                if '/' in condition:
                    date_format = current_date.strftime('%Y/%m/%d')
                elif '-' in condition:
                    date_format = current_date.strftime('%Y-%m-%d')
                all_variables[detail] = date_format
            
            elif condition.startswith('txt'):
                txt_path = condition[4:]
                if is_object(txt_path, '变量'):      # 如果是变量，转换为变量中的内容
                    txt_path = all_variables[txt_path[3:]]
                if not os.path.exists(txt_path):    # 如果文件不存在，执行结束
                    insert_text(f"错误：文件{txt_path}不存在！\n")
                    break
                with open(txt_path, 'r', encoding='gbk') as file:
                    all_variables[detail] = file.read()

            else:               # 串联若干个数文字
                final_str = ''
                parts = condition.split('|')
                for part in parts:
                    if is_object(part, '变量'):
                        final_str += str(get_variable_content(part[3:], all_variables))
                    elif is_object(part, '数组'):
                        data_content, _ = get_data_content(part[3:], all_data, all_data_index)
                        final_str += str(data_content)
                    else:
                        final_str += part
                all_variables[detail] = final_str

        elif '数组' in event:
            all_data_index[detail] += int(condition)

        elif '提示音' in event:
            playsound(f'sounds/{detail}.wav')

        elif 'i' in event:
            i = int(detail) - 3
            insert_text(f"执行第{detail}行指令\n")

        else:
            insert_text(f"没看出这一行是什么指令呢，跳过！\n")
            i += 1
            continue

        insert_text(f"时间：{cur_time if cur_time != -1.0 else ' '} 执行：{event} {detail}\n")
        i += 1

    keyboard_listener.stop()
    exit_flag = False
    rectangle_color = 'green'
    initialize_rectangle(rectangle_color)
    insert_text(f'执行完毕！\n\n')


# --------------------创建 GUI 界面--------------------
root = tk.Tk()
root.title("自动执行程序")
gui_width, gui_height = 550, 480

rate = 1.5
gui_width = int(rate * gui_width)
gui_height = int(rate * gui_height)

root.geometry(f'{gui_width}x{gui_height}')
root.resizable(False, False)

rect_width, rect_height = 150, 40
button_font_big = ('宋体', 14)
button_font_small = ('宋体', 12)


# -------加载背景-------
image = Image.open("backgrounds/1.png")
image = image.resize((gui_width, gui_height), Image.Resampling.LANCZOS)
photo = ImageTk.PhotoImage(image)

# 创建 Canvas 组件，设置为窗口的背景
canvas = tk.Canvas(root, width=image.width, height=image.height)
canvas.pack(fill="both", expand=True)
canvas.create_image(0, 0, anchor=tk.NW, image=photo)

# 处理窗口大小变化
def resize_image(event):
    width, height = root.winfo_width(), root.winfo_height()                   # 获取窗口大小
    resized_image = image.resize((width, height), Image.Resampling.LANCZOS)   # 根据新的大小调整图片
    resized_photo = ImageTk.PhotoImage(resized_image)                         # 转换成 Canvas 可以显示的图片格式
    canvas.delete("background")                                               # 删除旧背景
    canvas.create_image(0, 0, anchor=tk.NW, image=resized_photo)              # 在 Canvas 上更新背景图片
    canvas.image = resized_photo                                              # 更新 photo 防止垃圾回收
    initialize_rectangle(rectangle_color)                                     # 初始化矩形颜色

# 绑定窗口大小变化事件
root.bind("<Configure>", resize_image)


# -------生成矩形-------
def initialize_rectangle(color):
    rect = canvas.create_rectangle((gui_width - rect_width) // 2, 30, (gui_width + rect_width) // 2, 30 + rect_height, fill=color, tags="rectangle")


# -------获取 my_program 文件夹中的 Excel 文件列表-------
def get_excel_files():
    excel_files = []
    my_program_path = os.path.join(os.getcwd(), "my_programs")
    if os.path.exists(my_program_path):
        for file in os.listdir(my_program_path):
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(f"my_programs/{file}")
    return excel_files

# -------更新下拉菜单选项-------
def update_excel_dropdown():
    excel_files = get_excel_files()
    excel_combobox['values'] = excel_files

# -------当选择下拉菜单项时的处理函数-------
def on_excel_selected(event):
    selected = excel_combobox.get()
    excel_combobox.set(selected)

# -------Excel 路径输入框-------
excel_combobox = ttk.Combobox(root, width=43, font=button_font_big)
excel_combobox.place(relx=0.5, y=90, anchor="n")
update_excel_dropdown()    # 初始化下拉菜单选项
excel_combobox.bind("<<ComboboxSelected>>", on_excel_selected)


# -------运行按钮-------
def run_in_thread():
    threading.Thread(target=run).start()
save_button = tk.Button(root, text="运行", command=run_in_thread, font=button_font_big)
save_button.place(relx=0.55, y=140, anchor="n")

# -------清除按钮-------
def clear():
    excel_combobox.set('')
clear_button = tk.Button(root, text="清除", command=clear, font=button_font_big)
clear_button.place(relx=0.7, y=140, anchor="n")

# -------复选框-------
checkbox_var = tk.BooleanVar()
checkbox = tk.Checkbutton(root, text="运行后隐藏界面", variable=checkbox_var, font=button_font_big)
checkbox.place(relx=0.2, y=142, anchor="n")


# -------信息输出-------
info_text = tk.Text(root, height=16, width=60)
info_text.place(relx=0.5, y=205, anchor="n")


# -------启动 GUI-------
root.mainloop()
