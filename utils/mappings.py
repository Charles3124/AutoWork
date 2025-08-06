from pynput.mouse import Controller as MouseController, Button
from pynput.keyboard import Controller as KeyboardController, Key

# 组合键映射
combine_mapping = {
    '/x01': 'a',
    '/x02': 'b',
    '/x03': 'c',
    '/x04': 'd',
    '/x05': 'e',
    '/x06': 'f',
    '/x07': 'g',
    '/x08': 'h',
    '/t': 'i',
    '/n': 'j',
    '/x0b': 'k',
    '/x0c': 'l',
    '/r': 'm',
    '/x0e': 'n',
    '/x0f': 'o',
    '/x10': 'p',
    '/x11': 'q',
    '/x12': 'r',
    '/x13': 's',
    '/x14': 't',
    '/x15': 'u',
    '/x16': 'v',
    '/x17': 'w',
    '/x18': 'x',
    '/x19': 'y',
    '/x1a': 'z'
}

# 鼠标映射
button_mapping = {
    '左键': Button.left,
    '右键': Button.right
}

# 键盘映射
key_mapping = {
    '空格': Key.space,
    '回车': Key.enter,

    '上': Key.up,
    '下': Key.down,
    '左': Key.left,
    '右': Key.right,

    '删除': Key.backspace,
    'delete': Key.delete,

    'tab': Key.tab,
    'caps': Key.caps_lock,
    '大写': Key.caps_lock,
    'shift': Key.shift,

    'ctrl': Key.ctrl_l,
    '左ctrl': Key.ctrl_l,
    '右ctrl': Key.ctrl_r,

    'alt': Key.alt_l,
    '左alt': Key.alt_l,
    '右alt': Key.alt_gr,

    '上一页': Key.page_up,
    '下一页': Key.page_down,
    'home': Key.home,

    'f1': Key.f1,
    'f2': Key.f2,
    'f3': Key.f3,
    'f4': Key.f4,
    'f5': Key.f5,
    'f6': Key.f6,
    'f7': Key.f7,
    'f8': Key.f8,
    'f9': Key.f9,
    'f10': Key.f10,
    'f11': Key.f11,
    'f12': Key.f12
}
