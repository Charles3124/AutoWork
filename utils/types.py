"""
types.py

功能: 类型常量模块
时间: 2025/10/19
版本: 1.0
"""

from typing import Callable


TimeType = dict[int, float]
CommandType = dict[int, str]
RegionType = tuple[int, int, int, int] | str | None
PositionType = tuple[int, int] | tuple[()]
InsertFuncType = Callable[[str], None]
