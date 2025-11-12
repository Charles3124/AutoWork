"""
time_controller.py

功能: 时间功能模块
时间: 2025/11/13
版本: 1.0
"""

import time


class TimeController:
    """控制时间功能"""

    def __init__(self, times):
        self.times = times
        self.min_time_gap = 0.1    # 最小执行间距

    def sleep_start(self) -> None:
        """初始时延时"""
        time.sleep(max(self.times[0], self.min_time_gap))

    def sleep_loop(self, i: int) -> None:
        """循环中延时"""
        cur_time = self.times[i]
        if cur_time == -1.0 or self.times[i - 1] == -1.0:
            time.sleep(0.1)
        else:
            time.sleep(max(cur_time - self.times[i - 1], self.min_time_gap))
