"""
data_manager.py

功能: 集中封装和操作所有数据
时间: 2025/10/17
版本: 1.0
"""

class DataManager:
    """管理脚本执行期间的所有共享状态"""

    def __init__(self):
        self.all_data = {}         # 数组
        self.all_data_index = {}   # 数据索引
        self.all_variables = {}    # 变量
        self.all_figures = {}      # 图像

    def reset(self) -> None:
        """清空数据"""
        self.all_data.clear()
        self.all_data_index.clear()
        self.all_variables.clear()
        self.all_figures.clear()
