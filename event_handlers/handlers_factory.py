"""
handlers_factory.py

功能: 管理事件分发器
时间: 2025/11/16
版本: 1.0
"""

from flow_control import (LoopController, SelectController, WaitController)
from data_io import (DataReader, DataWriter)
from file_manager import FileController
from image_control import ScreenShotController


def get_event_handlers(
        loop_controller: LoopController, select_controller: SelectController, wait_controller: WaitController,
        data_reader: DataReader, data_writer: DataWriter,
        file_controller: FileController, screen_shot_controller: ScreenShotController
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
