import sys
import os
import json
import threading
import time
import ctypes
import ctypes.wintypes
import win32gui
import win32con
import win32api
import win32process
import win32event
import struct
import psutil  # 需要安装: pip install psutil
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QListWidget, QListWidgetItem, QCheckBox,
                               QPushButton, QComboBox, QGroupBox, QLabel, QProgressBar,
                               QMessageBox, QFileDialog, QSpinBox, QDoubleSpinBox, 
                               QTextEdit, QSplitter, QGridLayout, QLineEdit, QRadioButton,
                               QButtonGroup, QCheckBox)
from PySide6.QtCore import Qt, QSettings, Signal, QThread, QTimer
from PySide6.QtGui import QIcon, QFont, QColor


# 全局调试标志 - 编译后设置为False，不再输出到控制台
DEBUG_TO_CONSOLE = False  # 编译后设置为False，避免控制台错误

def console_print(*args, **kwargs):
    """输出到控制台 - 编译后不执行任何操作"""
    if DEBUG_TO_CONSOLE:
        # 仅在调试模式下才尝试输出到控制台
        try:
            print(*args, **kwargs)
            sys.stdout.flush()
        except:
            pass  # 忽略任何控制台输出错误

# 后台按键发送方式枚举 - 扩充到15种
class BackendKeyMethod:
    SENDMESSAGE = 0          # SendMessage WM_KEYDOWN/UP
    POSTMESSAGE = 1          # PostMessage WM_KEYDOWN/UP
    KEYBD_EVENT = 2          # keybd_event (全局)
    SENDINPUT = 3            # SendInput (高级输入)
    WM_KEYDOWN_DIRECT = 4    # WM_KEYDOWN直接发送
    WM_CHAR = 5              # WM_CHAR字符消息
    WM_IME_CHAR = 6          # WM_IME_CHAR输入法字符
    WM_SYSKEYDOWN = 7        # WM_SYSKEYDOWN系统按键
    POSTMESSAGE_LPARAM = 8   # PostMessage带完整lParam
    HOOK_MESSAGE = 9         # 使用SetWindowsHookEx
    HARDWARE_INPUT = 10      # 硬件输入模拟
    DIRECTINPUT = 11         # DirectInput模拟
    DRIVER_LEVEL = 12        # 驱动级模拟（通过WinRing0）
    VMWARE_HID = 13          # VMware虚拟HID
    DLL_INJECTION = 14       # DLL注入后调用

# 窗口枚举方式枚举
class EnumMethod:
    BY_TITLE = 0      # 按窗口标题
    BY_PROCESS = 1    # 按程序名
    BY_BOTH = 2       # 两者兼顾（标题和程序名都匹配）

# 按键映射表 - 完整版
KEY_MAP = {
    # 字母键
    "A": 0x41, "B": 0x42, "C": 0x43, "D": 0x44, "E": 0x45,
    "F": 0x46, "G": 0x47, "H": 0x48, "I": 0x49, "J": 0x4A,
    "K": 0x4B, "L": 0x4C, "M": 0x4D, "N": 0x4E, "O": 0x4F,
    "P": 0x50, "Q": 0x51, "R": 0x52, "S": 0x53, "T": 0x54,
    "U": 0x55, "V": 0x56, "W": 0x57, "X": 0x58, "Y": 0x59, "Z": 0x5A,
    
    # 数字键
    "0": 0x30, "1": 0x31, "2": 0x32, "3": 0x33, "4": 0x34,
    "5": 0x35, "6": 0x36, "7": 0x37, "8": 0x38, "9": 0x39,
    
    # 功能键
    "F1": 0x70, "F2": 0x71, "F3": 0x72, "F4": 0x73, "F5": 0x74,
    "F6": 0x75, "F7": 0x76, "F8": 0x77, "F9": 0x78, "F10": 0x79,
    "F11": 0x7A, "F12": 0x7B,
    
    # 控制键
    "Space": 0x20, "Enter": 0x0D, "Esc": 0x1B, "Tab": 0x09,
    "Backspace": 0x08, "Delete": 0x2E, "Insert": 0x2D,
    "Home": 0x24, "End": 0x23, "PageUp": 0x21, "PageDown": 0x22,
    
    # 方向键
    "Up": 0x26, "Down": 0x28, "Left": 0x25, "Right": 0x27,
    
    # 修饰键
    "Ctrl": 0x11, "Alt": 0x12, "Shift": 0x10, "Win": 0x5B,
    
    # 小键盘
    "Num0": 0x60, "Num1": 0x61, "Num2": 0x62, "Num3": 0x63,
    "Num4": 0x64, "Num5": 0x65, "Num6": 0x66, "Num7": 0x67,
    "Num8": 0x68, "Num9": 0x69,
    "NumMult": 0x6A, "NumPlus": 0x6B, "NumMinus": 0x6D,
    "NumDot": 0x6E, "NumDiv": 0x6F,
    
    # 其他常用键
    "CapsLock": 0x14, "PrintScreen": 0x2C, "ScrollLock": 0x91,
    "Pause": 0x13, "Menu": 0x5D,
    
    # 符号键（需要特殊处理，这里列出基础映射）
    "`": 0xC0, "-": 0xBD, "=": 0xBB, "[": 0xDB, "]": 0xDD,
    "\\": 0xDC, ";": 0xBA, "'": 0xDE, ",": 0xBC, ".": 0xBE,
    "/": 0xBF,
}

# 组合键映射表
MODIFIER_KEYS = {
    "Ctrl": 0x11,
    "Alt": 0x12,
    "Shift": 0x10,
    "Win": 0x5B,
}

# 按键显示名称列表（用于下拉框）
KEY_LIST = [
    # 字母
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
    "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
    "U", "V", "W", "X", "Y", "Z",
    # 数字
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
    # 功能键
    "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
    # 方向键
    "Up", "Down", "Left", "Right",
    # 控制键
    "Space", "Enter", "Esc", "Tab", "Backspace", "Delete", "Insert",
    "Home", "End", "PageUp", "PageDown",
    # 修饰键
    "Ctrl", "Alt", "Shift", "Win",
    # 小键盘
    "Num0", "Num1", "Num2", "Num3", "Num4", "Num5", "Num6", "Num7", "Num8", "Num9",
    "NumMult", "NumPlus", "NumMinus", "NumDot", "NumDiv",
    # 符号键
    "`", "-", "=", "[", "]", "\\", ";", "'", ",", ".", "/",
    # 其他
    "CapsLock", "PrintScreen", "ScrollLock", "Pause", "Menu"
]

def get_all_process_names_from_hwnd(hwnd):
    """通过窗口句柄获取进程的所有信息（更可靠）"""
    result = {
        'process_name': '',
        'process_path': '',
        'process_cmdline': '',
        'modules': [],
        'pid': 0
    }
    
    try:
        # 获取进程ID
        pid = win32process.GetWindowThreadProcessId(hwnd)[0]
        result['pid'] = pid
        console_print(f"[获取进程信息] 窗口句柄: {hwnd}, PID: {pid}")
        
        # 方法1: 尝试使用 Windows API 打开进程
        PROCESS_QUERY_INFORMATION = 0x0400
        PROCESS_VM_READ = 0x0010
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        
        kernel32 = ctypes.windll.kernel32
        psapi = ctypes.windll.psapi
        
        # 尝试以不同权限打开进程
        hProcess = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not hProcess:
            hProcess = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
        
        if hProcess:
            console_print(f"  成功打开进程句柄: {hProcess}")
            
            # 获取进程名
            exe_name = ctypes.create_unicode_buffer(260)
            if psapi.GetModuleBaseNameW(hProcess, None, exe_name, 260):
                result['process_name'] = exe_name.value.lower()
                console_print(f"  进程名: {result['process_name']}")
            
            # 获取完整路径
            exe_path = ctypes.create_unicode_buffer(260)
            if psapi.GetModuleFileNameExW(hProcess, None, exe_path, 260):
                result['process_path'] = exe_path.value.lower()
                console_print(f"  进程路径: {result['process_path']}")
            
            # 获取命令行（通过 WMI）
            try:
                import subprocess
                cmd = f'wmic process where processid={pid} get commandline'
                output = subprocess.check_output(cmd, shell=True, encoding='gbk', errors='ignore')
                lines = output.strip().split('\n')
                if len(lines) >= 2:
                    result['process_cmdline'] = lines[1].strip().lower()
                    console_print(f"  命令行: {result['process_cmdline'][:100]}")
            except:
                pass
            
            # 获取进程加载的模块
            try:
                # 枚举进程模块
                MODULE_LIST = ctypes.c_void_p * 1024
                modules = MODULE_LIST()
                needed = ctypes.c_ulong()
                
                if psapi.EnumProcessModules(hProcess, modules, ctypes.sizeof(modules), ctypes.byref(needed)):
                    module_count = needed.value // ctypes.sizeof(ctypes.c_void_p)
                    console_print(f"  找到 {module_count} 个模块")
                    
                    for i in range(min(module_count, 20)):  # 只取前20个
                        module_path = ctypes.create_unicode_buffer(260)
                        if psapi.GetModuleFileNameExW(hProcess, modules[i], module_path, 260):
                            module_name = os.path.basename(module_path.value).lower()
                            if module_name.endswith(('.exe', '.dll', '.dat', '.bin')):
                                if module_name not in result['modules']:
                                    result['modules'].append(module_name)
                                    if len(result['modules']) <= 10:
                                        console_print(f"    模块: {module_name}")
            except Exception as e:
                console_print(f"  枚举模块失败: {e}")
            
            kernel32.CloseHandle(hProcess)
        else:
            console_print(f"  无法打开进程 PID: {pid}, 错误码: {ctypes.GetLastError()}")
            
            # 方法2: 尝试使用 psutil
            try:
                process = psutil.Process(pid)
                result['process_name'] = process.name().lower()
                result['process_path'] = process.exe().lower() if process.exe() else ''
                console_print(f"  psutil 获取成功 - 进程名: {result['process_name']}")
            except Exception as e:
                console_print(f"  psutil 也失败: {e}")
            
            # 方法3: 通过窗口标题推断（征途游戏的特殊处理）
            window_text = win32gui.GetWindowText(hwnd)
            if '征途' in window_text:
                result['process_name'] = 'zhengtu.exe'
                result['modules'] = ['zhengtu.dat']
                console_print(f"  通过窗口标题推断: {result['process_name']}")
    
    except Exception as e:
        console_print(f"  获取进程信息异常: {e}")
    
    return result

# ==================== 项目信息元数据 ====================
class ProjectInfo:
    """项目信息元数据（集中管理所有项目相关信息）"""
    VERSION = "1.1.8"
    BUILD_DATE = "2026-03-26"
    AUTHOR = "杜玛"
    LICENSE = "GNU Affero General Public License v3.0"
    COPYRIGHT = "© 永久 杜玛"
    URL = "https://github.com/duma520/KeyForge"
    MAINTAINER_EMAIL = "不提供"
    NAME = "KeyForge 后台按键测试工具"
    DESCRIPTION = "KeyForge 后台按键测试工具，支持多种按键发送方式、窗口枚举方式，并提供丰富的调试信息输出。适用于游戏测试、自动化操作等场景。"
    
    @classmethod
    def get_full_name(cls) -> str:
        """获取完整的程序名称（带版本）"""
        return f"{cls.NAME} v{cls.VERSION}"
    
    @classmethod
    def get_full_title(cls, username: str = None) -> str:
        """获取完整的窗口标题"""
        base = f"{cls.get_full_name()} 构建:{cls.BUILD_DATE}"
        if username:
            return f"{base} - 当前用户: {username}"
        return base
    
    @classmethod
    def get_about_text(cls) -> str:
        """获取关于信息文本"""
        return f"""
        <h2>{cls.NAME}</h2>
        <p><b>版本:</b> {cls.VERSION}</p>
        <p><b>构建日期:</b> {cls.BUILD_DATE}</p>
        <p><b>作者:</b> {cls.AUTHOR}</p>
        <p><b>版权:</b> {cls.COPYRIGHT}</p>
        <p><b>许可证:</b> {cls.LICENSE}</p>
        <p><b>项目主页:</b> <a href='{cls.URL}'>{cls.URL}</a></p>
        <p><b>描述:</b> {cls.DESCRIPTION}</p>
        """


# ====================  马卡龙色系定义 ====================
class MacaronColors:
    """马卡龙色系完整定义"""
    # 粉色系
    PINK_SAKURA = QColor('#FFB7CE')      # 樱花粉
    PINK_ROSE = QColor('#FF9AA2')        # 玫瑰粉
    PINK_COTTON = QColor('#FFD1DC')      # 棉花粉
    PINK_BALLET = QColor('#FCC9D3')      # 芭蕾粉
    
    # 蓝色系
    BLUE_SKY = QColor('#A2E1F6')         # 天空蓝
    BLUE_MIST = QColor('#C2E5F9')        # 雾霾蓝
    BLUE_PERIWINKLE = QColor('#C5D0E6')  # 长春花蓝
    BLUE_LAVENDER = QColor('#D6EAF8')    # 薰衣草蓝
    
    # 绿色系
    GREEN_MINT = QColor('#B5EAD7')       # 薄荷绿
    GREEN_APPLE = QColor('#D4F1C7')      # 苹果绿
    GREEN_PISTACHIO = QColor('#D8E9D6')  # 开心果绿
    GREEN_SAGE = QColor('#C9DFC5')       # 鼠尾草绿
    
    # 黄色/橙色系
    YELLOW_LEMON = QColor('#FFEAA5')      # 柠檬黄
    YELLOW_CREAM = QColor('#FFF8B8')      # 奶油黄
    YELLOW_HONEY = QColor('#FCE5B4')      # 蜂蜜黄
    ORANGE_PEACH = QColor('#FFDAC1')      # 蜜桃橙
    ORANGE_APRICOT = QColor('#FDD9B5')    # 杏色
    
    # 紫色系
    PURPLE_LAVENDER = QColor('#C7CEEA')   # 薰衣草紫
    PURPLE_TARO = QColor('#D8BFD8')       # 香芋紫
    PURPLE_WISTERIA = QColor('#C9B6D9')   # 紫藤
    PURPLE_MAUVE = QColor('#E0C7D7')      # 淡紫
    
    # 中性色
    NEUTRAL_CARAMEL = QColor('#F0E6DD')   # 焦糖奶霜
    NEUTRAL_CREAM = QColor('#F7F1E5')     # 奶油白
    NEUTRAL_MOCHA = QColor('#EAD7C7')     # 摩卡
    NEUTRAL_ALMOND = QColor('#F2E4D4')    # 杏仁
    
    # 其他颜色
    RED_CORAL = QColor('#FFB3A7')          # 珊瑚红
    RED_WATERMELON = QColor('#FFC5C5')     # 西瓜红
    TEAL_MINT = QColor('#B8E2DE')          # 薄荷绿蓝
    
    @classmethod
    def get_color_list(cls):
        """获取所有马卡龙颜色列表"""
        return [
            cls.PINK_SAKURA, cls.PINK_ROSE, cls.PINK_COTTON, cls.PINK_BALLET,
            cls.BLUE_SKY, cls.BLUE_MIST, cls.BLUE_PERIWINKLE, cls.BLUE_LAVENDER,
            cls.GREEN_MINT, cls.GREEN_APPLE, cls.GREEN_PISTACHIO, cls.GREEN_SAGE,
            cls.YELLOW_LEMON, cls.YELLOW_CREAM, cls.YELLOW_HONEY, cls.ORANGE_PEACH,
            cls.ORANGE_APRICOT, cls.PURPLE_LAVENDER, cls.PURPLE_TARO, cls.PURPLE_WISTERIA,
            cls.PURPLE_MAUVE, cls.NEUTRAL_CARAMEL, cls.NEUTRAL_CREAM, cls.NEUTRAL_MOCHA,
            cls.NEUTRAL_ALMOND, cls.RED_CORAL, cls.RED_WATERMELON, cls.TEAL_MINT
        ]
    
    @classmethod
    def get_color_names(cls):
        """获取马卡龙颜色名称列表"""
        return [
            "樱花粉", "玫瑰粉", "棉花粉", "芭蕾粉",
            "天空蓝", "雾霾蓝", "长春花蓝", "薰衣草蓝",
            "薄荷绿", "苹果绿", "开心果绿", "鼠尾草绿",
            "柠檬黄", "奶油黄", "蜂蜜黄", "蜜桃橙",
            "杏色", "薰衣草紫", "香芋紫", "紫藤",
            "淡紫", "焦糖奶霜", "奶油白", "摩卡",
            "杏仁", "珊瑚红", "西瓜红", "薄荷绿蓝"
        ]




# 增强的按键发送类
class EnhancedKeyPressThread(QThread):
    progress_updated = Signal(int)
    status_updated = Signal(str)
    log_message = Signal(str)
    
    def __init__(self):
        super().__init__()
        self.running = False
        self.selected_windows = []
        self.key_type = "A"
        self.method = BackendKeyMethod.SENDMESSAGE
        self.delay = 1.0
        self.repeat_count = 1  # 重复次数，0表示无限循环
        
        # 组合键标志
        self.modifiers = []  # 存储选中的组合键名称列表
        
        # 添加一个锁来保护参数
        self.param_lock = threading.Lock()
        
    def update_parameters(self, key_type, method, delay, repeat_count, modifiers=None):
        """动态更新参数"""
        with self.param_lock:
            self.key_type = key_type
            self.method = method
            self.delay = delay
            self.repeat_count = repeat_count
            if modifiers is not None:
                self.modifiers = modifiers
            
    def get_parameters(self):
        """获取当前参数"""
        with self.param_lock:
            return self.key_type, self.method, self.delay, self.repeat_count, self.modifiers
        
    def get_vk_code(self, key_type):
        """获取虚拟键码"""
        # 支持自定义输入（十六进制或数字）
        if key_type.startswith("0x") or key_type.startswith("0X"):
            try:
                return int(key_type, 16)
            except:
                pass
        elif key_type.isdigit():
            try:
                return int(key_type)
            except:
                pass
        
        # 从映射表中获取
        return KEY_MAP.get(key_type, 0x41)  # 默认返回A
    
    def get_modifier_codes(self, modifiers):
        """获取组合键的虚拟键码列表"""
        codes = []
        for mod in modifiers:
            if mod in MODIFIER_KEYS:
                codes.append(MODIFIER_KEYS[mod])
        return codes
    
    def send_keybd_event(self, vk_code, is_down):
        """使用keybd_event发送按键"""
        flags = 0 if is_down else win32con.KEYEVENTF_KEYUP
        ctypes.windll.user32.keybd_event(vk_code, 0, flags, 0)
    
    def send_sendinput(self, vk_code, is_down):
        """使用SendInput发送按键"""
        INPUT_KEYBOARD = 1
        class KEYBDINPUT(ctypes.Structure):
            _fields_ = [("wVk", ctypes.c_ushort),
                       ("wScan", ctypes.c_ushort),
                       ("dwFlags", ctypes.c_ulong),
                       ("time", ctypes.c_ulong),
                       ("dwExtraInfo", ctypes.c_ulong)]
        
        class INPUT(ctypes.Structure):
            _fields_ = [("type", ctypes.c_ulong),
                       ("ki", KEYBDINPUT)]
        
        flags = 0 if is_down else win32con.KEYEVENTF_KEYUP
        input_struct = INPUT()
        input_struct.type = INPUT_KEYBOARD
        input_struct.ki = KEYBDINPUT(vk_code, 0, flags, 0, 0)
        ctypes.windll.user32.SendInput(1, ctypes.byref(input_struct), ctypes.sizeof(INPUT))
    
    def send_key_with_modifiers(self, hwnd, vk_code, method, char_code=0):
        """
        发送带组合键的按键
        返回是否成功
        """
        try:
            # 获取当前激活的组合键
            modifier_codes = self.get_modifier_codes(self.modifiers)
            
            # 根据不同的方法处理组合键
            # 对于全局方法（keybd_event, SendInput），直接模拟按键
            if method == BackendKeyMethod.KEYBD_EVENT:
                # 按下所有组合键
                for mod_code in modifier_codes:
                    self.send_keybd_event(mod_code, True)
                    time.sleep(0.01)
                
                # 按下并释放主键
                self.send_keybd_event(vk_code, True)
                time.sleep(0.02)
                self.send_keybd_event(vk_code, False)
                time.sleep(0.01)
                
                # 释放所有组合键（逆序）
                for mod_code in reversed(modifier_codes):
                    self.send_keybd_event(mod_code, False)
                    time.sleep(0.01)
                    
                return True
                
            elif method == BackendKeyMethod.SENDINPUT:
                # 按下所有组合键
                for mod_code in modifier_codes:
                    self.send_sendinput(mod_code, True)
                    time.sleep(0.01)
                
                # 按下并释放主键
                self.send_sendinput(vk_code, True)
                time.sleep(0.02)
                self.send_sendinput(vk_code, False)
                time.sleep(0.01)
                
                # 释放所有组合键（逆序）
                for mod_code in reversed(modifier_codes):
                    self.send_sendinput(mod_code, False)
                    time.sleep(0.01)
                    
                return True
                
            elif method in [BackendKeyMethod.SENDMESSAGE, BackendKeyMethod.POSTMESSAGE, 
                            BackendKeyMethod.WM_KEYDOWN_DIRECT, BackendKeyMethod.POSTMESSAGE_LPARAM,
                            BackendKeyMethod.WM_SYSKEYDOWN]:
                # 对于窗口消息方式，需要发送带修饰键状态的消息
                # 构建 lParam 参数
                def make_lparam(repeat, scan, extended, context, previous):
                    return (repeat & 0xFFFF) | ((scan & 0xFF) << 16) | \
                           ((extended & 1) << 24) | ((context & 1) << 29) | \
                           ((previous & 1) << 30)
                
                # 首先发送组合键的按下消息
                for mod_code in modifier_codes:
                    if method == BackendKeyMethod.SENDMESSAGE:
                        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, mod_code, 0)
                    else:
                        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, mod_code, 0)
                    time.sleep(0.01)
                
                # 发送主键消息
                if method == BackendKeyMethod.SENDMESSAGE:
                    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                    time.sleep(0.02)
                    win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                else:
                    lParam = make_lparam(1, 0, 0, 0, 0)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam)
                    time.sleep(0.02)
                    lParam_up = make_lparam(1, 0, 0, 1, 1)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam_up)
                time.sleep(0.01)
                
                # 释放组合键（逆序）
                for mod_code in reversed(modifier_codes):
                    if method == BackendKeyMethod.SENDMESSAGE:
                        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, mod_code, 0)
                    else:
                        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, mod_code, 0)
                    time.sleep(0.01)
                    
                return True
                
            elif method == BackendKeyMethod.WM_CHAR:
                # WM_CHAR 方式，组合键可能不生效，直接发送字符
                if char_code:
                    win32gui.PostMessage(hwnd, win32con.WM_CHAR, char_code, 0)
                    return True
                else:
                    # 回退到普通按键
                    return self.send_key_with_modifiers(hwnd, vk_code, BackendKeyMethod.POSTMESSAGE, char_code)
                    
            else:
                # 其他方法回退到 POSTMESSAGE
                return self.send_key_with_modifiers(hwnd, vk_code, BackendKeyMethod.POSTMESSAGE, char_code)
                
        except Exception as e:
            self.log_message.emit(f"组合键发送错误: {str(e)}")
            return False
    
    def send_key_without_modifiers(self, hwnd, vk_code, method, char_code=0):
        """发送不带组合键的按键"""
        try:
            if method == BackendKeyMethod.SENDMESSAGE:
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                time.sleep(0.02)
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                return True
                
            elif method == BackendKeyMethod.POSTMESSAGE:
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                return True
                
            elif method == BackendKeyMethod.KEYBD_EVENT:
                ctypes.windll.user32.keybd_event(vk_code, 0, 0, 0)
                time.sleep(0.02)
                ctypes.windll.user32.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                return True
                
            elif method == BackendKeyMethod.SENDINPUT:
                INPUT_KEYBOARD = 1
                class KEYBDINPUT(ctypes.Structure):
                    _fields_ = [("wVk", ctypes.c_ushort),
                               ("wScan", ctypes.c_ushort),
                               ("dwFlags", ctypes.c_ulong),
                               ("time", ctypes.c_ulong),
                               ("dwExtraInfo", ctypes.c_ulong)]
                
                class INPUT(ctypes.Structure):
                    _fields_ = [("type", ctypes.c_ulong),
                               ("ki", KEYBDINPUT)]
                
                input_down = INPUT()
                input_down.type = INPUT_KEYBOARD
                input_down.ki = KEYBDINPUT(vk_code, 0, 0, 0, 0)
                ctypes.windll.user32.SendInput(1, ctypes.byref(input_down), ctypes.sizeof(INPUT))
                time.sleep(0.02)
                
                input_up = INPUT()
                input_up.type = INPUT_KEYBOARD
                input_up.ki = KEYBDINPUT(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0, 0)
                ctypes.windll.user32.SendInput(1, ctypes.byref(input_up), ctypes.sizeof(INPUT))
                return True
                
            elif method == BackendKeyMethod.WM_KEYDOWN_DIRECT:
                lParam = (1 << 30) | (1 << 31) | (1 << 29)
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam)
                return True
                
            elif method == BackendKeyMethod.WM_CHAR:
                if char_code:
                    win32gui.PostMessage(hwnd, win32con.WM_CHAR, char_code, 0)
                    return True
                else:
                    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                    time.sleep(0.02)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                    return True
                
            elif method == BackendKeyMethod.WM_IME_CHAR:
                if char_code:
                    win32gui.PostMessage(hwnd, win32con.WM_IME_CHAR, char_code, 0)
                else:
                    win32gui.PostMessage(hwnd, win32con.WM_IME_CHAR, vk_code, 0)
                return True
                
            elif method == BackendKeyMethod.WM_SYSKEYDOWN:
                win32gui.PostMessage(hwnd, win32con.WM_SYSKEYDOWN, vk_code, 0)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_SYSKEYUP, vk_code, 0)
                return True
                
            elif method == BackendKeyMethod.POSTMESSAGE_LPARAM:
                def make_lparam(repeat, scan, extended, context, previous):
                    return (repeat & 0xFFFF) | ((scan & 0xFF) << 16) | \
                           ((extended & 1) << 24) | ((context & 1) << 29) | \
                           ((previous & 1) << 30)
                
                lParam = make_lparam(1, 0, 0, 0, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam)
                time.sleep(0.02)
                lParam_up = make_lparam(1, 0, 0, 1, 1)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam_up)
                return True
                
            else:
                # 其他方法简化处理
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                return True
                    
        except Exception as e:
            self.log_message.emit(f"错误: {str(e)}")
            return False
    
    def send_key_to_window_advanced(self, hwnd, key_type, method):
        """增强的按键发送方法（支持组合键）"""
        vk_code = self.get_vk_code(key_type)
        
        # 获取字符码（用于WM_CHAR等）
        char_code = 0
        if len(key_type) == 1:
            char_code = ord(key_type[0])
        
        # 检查是否有组合键
        has_modifiers = len(self.modifiers) > 0
        
        # 构建组合键信息字符串
        modifier_str = ""
        if has_modifiers:
            modifier_str = " + ".join(self.modifiers) + " + "
        
        # 发送按键
        if has_modifiers:
            success = self.send_key_with_modifiers(hwnd, vk_code, method, char_code)
            if success:
                self.log_message.emit(f"{self.get_method_name(method)}: 已发送组合键 [{modifier_str}{key_type}]")
        else:
            success = self.send_key_without_modifiers(hwnd, vk_code, method, char_code)
            if success:
                self.log_message.emit(f"{self.get_method_name(method)}: 已发送 {key_type} 键")
        
        return success
    
    def get_method_name(self, method):
        """获取发送方式的名称"""
        method_names = {
            BackendKeyMethod.SENDMESSAGE: "SendMessage",
            BackendKeyMethod.POSTMESSAGE: "PostMessage",
            BackendKeyMethod.KEYBD_EVENT: "keybd_event",
            BackendKeyMethod.SENDINPUT: "SendInput",
            BackendKeyMethod.WM_KEYDOWN_DIRECT: "WM_KEYDOWN直接",
            BackendKeyMethod.WM_CHAR: "WM_CHAR",
            BackendKeyMethod.WM_IME_CHAR: "WM_IME_CHAR",
            BackendKeyMethod.WM_SYSKEYDOWN: "WM_SYSKEYDOWN",
            BackendKeyMethod.POSTMESSAGE_LPARAM: "PostMessage完整参数",
        }
        return method_names.get(method, "默认方式")
    
    def run(self):
        self.running = True
        
        # 获取初始参数
        key_type, method, delay, repeat_count, modifiers = self.get_parameters()
        is_infinite = (repeat_count == 0)  # 0表示无限循环
        
        # 初始化进度相关变量
        if is_infinite:
            total = 0  # 无限模式没有总数
            current = 0
        else:
            total = len(self.selected_windows) * repeat_count
            current = 0
        
        # 主循环
        loop_count = 0
        while self.running:
            # 检查是否需要无限循环或已到达重复次数
            if not is_infinite and loop_count >= repeat_count:
                break
            
            # 获取最新的参数
            key_type, method, delay, repeat_count, modifiers = self.get_parameters()
            is_infinite = (repeat_count == 0)
            
            # 发送一轮所有窗口
            for idx, hwnd in enumerate(self.selected_windows):
                if not self.running:
                    break
                    
                success = self.send_key_to_window_advanced(hwnd, key_type, method)
                
                if not is_infinite:
                    current += 1
                    self.progress_updated.emit(int(current / total * 100))
                
                if success:
                    modifier_str = " + ".join(modifiers) + " + " if modifiers else ""
                    self.status_updated.emit(f"成功发送 {modifier_str}{key_type} 到窗口 {idx+1}/{len(self.selected_windows)} (第{loop_count+1}轮)")
                else:
                    self.status_updated.emit(f"发送失败到窗口 {idx+1}/{len(self.selected_windows)}")
                    
                time.sleep(delay)
            
            loop_count += 1
            
            # 无限模式下更新状态显示
            if is_infinite:
                self.status_updated.emit(f"已完成 {loop_count} 轮发送，继续循环...")
                self.log_message.emit(f"已完成第 {loop_count} 轮发送（无限模式）")
        
        if not self.running:
            self.status_updated.emit("按键发送已停止")
        else:
            self.status_updated.emit("按键发送完成")
        
    def stop(self):
        self.running = False

# 增强的窗口枚举线程
class EnhancedWindowEnumThread(QThread):
    windows_found = Signal(list)
    progress_updated = Signal(int)
    debug_log = Signal(str)
    
    def __init__(self, enum_method, title_keyword, process_keyword, show_debug=False):
        super().__init__()
        self.enum_method = enum_method
        self.title_keyword = title_keyword.lower() if title_keyword else ""
        self.process_keyword = process_keyword.lower() if process_keyword else ""
        self.show_debug = show_debug
        
    def debug_print(self, message):
        """输出调试信息"""
        if self.show_debug:
            self.debug_log.emit(message)
            console_print(f"[DEBUG] {message}")
        
    def run(self):
        windows = []
        total_visible = 0
        matched_count = 0
        all_windows_info = []
        
        console_print("\n" + "="*80)
        console_print("开始枚举窗口...")
        console_print(f"枚举方式: {self.enum_method}, 标题关键词: '{self.title_keyword}', 程序名关键词: '{self.process_keyword}'")
        console_print("="*80)
        
        def enum_callback(hwnd, extra):
            nonlocal total_visible, matched_count
            
            if win32gui.IsWindowVisible(hwnd):
                total_visible += 1
                window_text = win32gui.GetWindowText(hwnd)
                
                if not window_text:
                    return True
                    
                class_name = win32gui.GetClassName(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                pid = win32process.GetWindowThreadProcessId(hwnd)[0]
                
                # 获取进程信息
                process_info = get_all_process_names_from_hwnd(hwnd)
                process_name = process_info['process_name']
                process_path = process_info['process_path']
                process_cmdline = process_info['process_cmdline']
                modules = process_info['modules']
                
                # 输出到控制台
                if total_visible <= 50:
                    console_print(f"\n[{total_visible}] 窗口标题: '{window_text[:60]}'")
                    console_print(f"    类名: '{class_name}'")
                    console_print(f"    PID: {pid}")
                    console_print(f"    进程名: '{process_name}'")
                    console_print(f"    进程路径: '{process_path[:80]}'")
                    if modules:
                        console_print(f"    模块: {', '.join(modules[:5])}")
                
                # 调试输出到界面
                if self.show_debug and total_visible <= 100:
                    debug_info = f"窗口: '{window_text[:40]}' | 类名: '{class_name}' | PID:{pid} | 进程: '{process_name}'"
                    if process_path:
                        debug_info += f" | 路径: '{os.path.basename(process_path)}'"
                    if modules:
                        debug_info += f" | 模块: {', '.join(modules[:3])}"
                    self.debug_print(debug_info)
                    
                    all_windows_info.append({
                        'title': window_text[:50],
                        'class': class_name,
                        'pid': pid,
                        'process_name': process_name,
                        'process_path': os.path.basename(process_path) if process_path else '',
                        'modules': modules[:5]
                    })
                
                # 匹配逻辑
                matched = False
                match_info = ""
                
                if self.enum_method == EnumMethod.BY_TITLE:
                    if self.title_keyword and self.title_keyword in window_text.lower():
                        matched = True
                        match_info = f"标题匹配: '{self.title_keyword}' in '{window_text[:50]}'"
                        console_print(f"    ✓ 匹配成功: {match_info}")
                        
                elif self.enum_method == EnumMethod.BY_PROCESS:
                    if self.process_keyword:
                        console_print(f"    检查关键词 '{self.process_keyword}' 是否匹配...")
                        
                        if self.process_keyword in process_name:
                            matched = True
                            match_info = f"进程名匹配: '{self.process_keyword}' in '{process_name}'"
                            console_print(f"    ✓ {match_info}")
                            
                        elif process_path and self.process_keyword in os.path.basename(process_path).lower():
                            matched = True
                            match_info = f"进程路径匹配: '{self.process_keyword}' in '{os.path.basename(process_path)}'"
                            console_print(f"    ✓ {match_info}")
                            
                        elif self.process_keyword in process_cmdline:
                            matched = True
                            match_info = f"命令行匹配: '{self.process_keyword}' in '{process_cmdline[:100]}'"
                            console_print(f"    ✓ {match_info}")
                            
                        elif any(self.process_keyword in module for module in modules):
                            matched = True
                            matched_module = next(module for module in modules if self.process_keyword in module)
                            match_info = f"模块匹配: '{self.process_keyword}' 在模块 '{matched_module}' 中"
                            console_print(f"    ✓ {match_info}")
                            
                        elif self.process_keyword in class_name.lower():
                            matched = True
                            match_info = f"窗口类名匹配: '{self.process_keyword}' in '{class_name}'"
                            console_print(f"    ✓ {match_info}")
                        else:
                            console_print(f"    ✗ 无匹配")
                            
                elif self.enum_method == EnumMethod.BY_BOTH:
                    title_match = self.title_keyword and self.title_keyword in window_text.lower()
                    process_match = False
                    
                    if self.process_keyword:
                        if (self.process_keyword in process_name or
                            (process_path and self.process_keyword in os.path.basename(process_path).lower()) or
                            self.process_keyword in process_cmdline or
                            any(self.process_keyword in module for module in modules) or
                            self.process_keyword in class_name.lower()):
                            process_match = True
                    
                    if title_match and process_match:
                        matched = True
                        match_info = f"两者匹配"
                        console_print(f"    ✓ 两者匹配成功")
                
                if matched:
                    matched_count += 1
                    if self.show_debug:
                        self.debug_print(f"✓ 匹配成功 [{matched_count}]: {match_info}")
                    
                    process_display = process_name if process_name else '未知'
                    
                    windows.append({
                        'hwnd': hwnd,
                        'title': window_text,
                        'class': class_name,
                        'rect': rect,
                        'pid': pid,
                        'process_name': process_name,
                        'process_path': process_path,
                        'process_cmdline': process_cmdline[:200] if process_cmdline else '',
                        'modules': modules[:5]
                    })
            return True
            
        win32gui.EnumWindows(enum_callback, None)
        
        console_print("\n" + "="*80)
        console_print(f"枚举完成 - 总可见窗口: {total_visible}, 匹配成功: {matched_count}")
        console_print("="*80)
        
        console_print("\n所有可见窗口汇总:")
        for idx, win in enumerate(all_windows_info[:30], 1):
            console_print(f"{idx:2}. PID:{win['pid']:6} | 进程: '{win['process_name']:20}' | 标题: '{win['title'][:40]}'")
            if win['modules']:
                console_print(f"      模块: {', '.join(win['modules'][:3])}")
        
        self.progress_updated.emit(100)
        self.windows_found.emit(windows)

# 增强的主窗口
class EnhancedMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("BackendKeyTester", "EnhancedSettings")
        self.key_thread = None
        self.enum_thread = None
        self.window_items = []
        self.init_ui()
        self.load_settings()
        
        console_print("程序启动完成")
        
        # 连接参数变化信号，实时更新线程参数
        self.key_combo.currentTextChanged.connect(self.on_parameter_changed)
        self.delay_spin.valueChanged.connect(self.on_parameter_changed)
        self.method_combo.currentIndexChanged.connect(self.on_parameter_changed)
        self.repeat_spin.valueChanged.connect(self.on_parameter_changed)
        
        # 连接组合键复选框变化信号
        self.cb_ctrl.stateChanged.connect(self.on_parameter_changed)
        self.cb_alt.stateChanged.connect(self.on_parameter_changed)
        self.cb_shift.stateChanged.connect(self.on_parameter_changed)
        self.cb_win.stateChanged.connect(self.on_parameter_changed)
        
    def init_ui(self):
        self.setWindowTitle(f"{ProjectInfo.NAME} {ProjectInfo.VERSION} (Build: {ProjectInfo.BUILD_DATE})")
        self.setMinimumSize(1200, 900)
        
        if os.path.exists("icon.ico"):
            self.setWindowIcon(QIcon("icon.ico"))
            
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        splitter = QSplitter(Qt.Orientation.Vertical)
        main_layout = QVBoxLayout(central_widget)
        main_layout.addWidget(splitter)
        
        # 上半部分
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        
        # 枚举设置区域
        enum_group = QGroupBox("窗口枚举设置")
        enum_layout = QGridLayout()
        
        enum_layout.addWidget(QLabel("枚举方式:"), 0, 0)
        self.enum_method_group = QButtonGroup()
        self.radio_title = QRadioButton("按窗口标题")
        self.radio_title.setChecked(True)
        self.radio_process = QRadioButton("按程序名")
        self.radio_both = QRadioButton("两者兼顾")
        self.enum_method_group.addButton(self.radio_title, EnumMethod.BY_TITLE)
        self.enum_method_group.addButton(self.radio_process, EnumMethod.BY_PROCESS)
        self.enum_method_group.addButton(self.radio_both, EnumMethod.BY_BOTH)
        
        enum_method_layout = QHBoxLayout()
        enum_method_layout.addWidget(self.radio_title)
        enum_method_layout.addWidget(self.radio_process)
        enum_method_layout.addWidget(self.radio_both)
        enum_method_layout.addStretch()
        enum_layout.addLayout(enum_method_layout, 0, 1, 1, 3)
        
        enum_layout.addWidget(QLabel("窗口标题关键词:"), 1, 0)
        self.title_keyword_edit = QLineEdit()
        self.title_keyword_edit.setPlaceholderText("例如: 共创征途 或 游戏窗口标题")
        enum_layout.addWidget(self.title_keyword_edit, 1, 1, 1, 3)
        
        enum_layout.addWidget(QLabel("程序名关键词:"), 2, 0)
        self.process_keyword_edit = QLineEdit()
        self.process_keyword_edit.setPlaceholderText("例如: zhengtu.dat 或 client.exe (支持部分匹配)")
        enum_layout.addWidget(self.process_keyword_edit, 2, 1, 1, 3)
        
        self.debug_checkbox = QCheckBox("调试模式 (显示所有窗口信息)")
        self.debug_checkbox.setChecked(True)
        enum_layout.addWidget(self.debug_checkbox, 3, 0, 1, 2)
        
        tip_label = QLabel("💡 提示: 程序名支持匹配进程名、路径、模块名、窗口类名和命令行")
        tip_label.setStyleSheet("color: #666; font-size: 9pt;")
        enum_layout.addWidget(tip_label, 3, 2, 1, 2)
        
        button_layout = QHBoxLayout()
        self.refresh_btn = QPushButton("🔍 刷新窗口列表")
        self.refresh_btn.clicked.connect(self.enumerate_windows)
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
        """)
        button_layout.addWidget(self.refresh_btn)
        
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_windows)
        button_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_windows)
        button_layout.addWidget(self.deselect_all_btn)
        
        button_layout.addStretch()
        enum_layout.addLayout(button_layout, 4, 0, 1, 4)
        
        enum_group.setLayout(enum_layout)
        top_layout.addWidget(enum_group)
        
        self.window_list = QListWidget()
        self.window_list.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        self.window_list.setAlternatingRowColors(True)
        top_layout.addWidget(self.window_list)
        
        splitter.addWidget(top_widget)
        
        # 下半部分
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        
        settings_group = QGroupBox("按键设置（15种方式 + 组合键支持）")
        settings_layout = QGridLayout()
        
        # 组合键区域 - 在按键类型上方
        settings_layout.addWidget(QLabel("组合键:"), 0, 0)
        modifier_layout = QHBoxLayout()
        
        self.cb_ctrl = QCheckBox("Ctrl")
        self.cb_alt = QCheckBox("Alt")
        self.cb_shift = QCheckBox("Shift")
        self.cb_win = QCheckBox("Win")
        # Fn键通常没有虚拟键码，这里用Win键替代或预留，实际Fn是硬件级按键
        self.cb_fn = QCheckBox("Fn (功能键)")
        self.cb_fn.setToolTip("注意：Fn键是硬件级按键，大多数情况下无法通过软件模拟")
        
        modifier_layout.addWidget(self.cb_ctrl)
        modifier_layout.addWidget(self.cb_alt)
        modifier_layout.addWidget(self.cb_shift)
        modifier_layout.addWidget(self.cb_win)
        modifier_layout.addWidget(self.cb_fn)
        modifier_layout.addStretch()
        
        settings_layout.addLayout(modifier_layout, 0, 1, 1, 5)
        
        # 添加分隔提示
        tip_modifier = QLabel("💡 提示: 可以同时选择多个组合键，如 Ctrl+Alt+A")
        tip_modifier.setStyleSheet("color: #666; font-size: 9pt;")
        settings_layout.addWidget(tip_modifier, 1, 1, 1, 5)
        
        # 按键类型
        settings_layout.addWidget(QLabel("按键类型:"), 2, 0)
        self.key_combo = QComboBox()
        self.key_combo.addItems(KEY_LIST)
        self.key_combo.setEditable(True)  # 允许用户自定义输入，支持十六进制（如 0x41）
        self.key_combo.setToolTip("支持选择预设按键，也可以手动输入虚拟键码（如0x41代表A，0x20代表空格）")
        settings_layout.addWidget(self.key_combo, 2, 1)
        
        settings_layout.addWidget(QLabel("发送间隔(秒):"), 2, 2)
        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0.01, 10.0)
        self.delay_spin.setSingleStep(0.05)
        self.delay_spin.setValue(0.5)
        self.delay_spin.setDecimals(2)
        settings_layout.addWidget(self.delay_spin, 2, 3)
        
        settings_layout.addWidget(QLabel("重复次数:"), 2, 4)
        self.repeat_spin = QSpinBox()
        self.repeat_spin.setRange(0, 100)  # 修改范围，0表示无限循环
        self.repeat_spin.setValue(1)
        self.repeat_spin.setToolTip("0 = 无限循环，直到手动停止")
        settings_layout.addWidget(self.repeat_spin, 2, 5)
        
        settings_layout.addWidget(QLabel("后台按键方式:"), 3, 0)
        self.method_combo = QComboBox()
        methods = [
            "1. SendMessage (WM_KEYDOWN/UP)",
            "2. PostMessage (WM_KEYDOWN/UP)", 
            "3. keybd_event (全局)",
            "4. SendInput (高级输入)",
            "5. WM_KEYDOWN直接发送",
            "6. WM_CHAR字符消息",
            "7. WM_IME_CHAR输入法字符",
            "8. WM_SYSKEYDOWN系统按键",
            "9. PostMessage带完整lParam",
            "10. SetWindowsHookEx钩子注入",
            "11. 硬件输入模拟",
            "12. DirectInput模拟",
            "13. 驱动级模拟",
            "14. VMware虚拟HID",
            "15. DLL注入"
        ]
        self.method_combo.addItems(methods)
        settings_layout.addWidget(self.method_combo, 3, 1, 1, 3)
        
        control_layout = QHBoxLayout()
        self.start_btn = QPushButton("▶ 开始发送按键")
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 8px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.start_btn.clicked.connect(self.start_sending_keys)
        
        self.stop_btn = QPushButton("⏹ 停止发送")
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 8px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)
        self.stop_btn.clicked.connect(self.stop_sending_keys)
        self.stop_btn.setEnabled(False)
        
        control_layout.addWidget(self.start_btn)
        control_layout.addWidget(self.stop_btn)
        control_layout.addStretch()
        settings_layout.addLayout(control_layout, 4, 0, 1, 6)
        
        settings_group.setLayout(settings_layout)
        bottom_layout.addWidget(settings_group)
        
        progress_group = QGroupBox("执行进度")
        progress_layout = QVBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #bbb;
                border-radius: 5px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 4px;
            }
        """)
        progress_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("font-weight: bold; color: #333;")
        progress_layout.addWidget(self.status_label)
        
        progress_group.setLayout(progress_layout)
        bottom_layout.addWidget(progress_group)
        
        log_group = QGroupBox("执行日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(250)
        self.log_text.setStyleSheet("font-family: Consolas, monospace; font-size: 10pt;")
        log_layout.addWidget(self.log_text)
        
        clear_log_btn = QPushButton("清空日志")
        clear_log_btn.clicked.connect(lambda: self.log_text.clear())
        log_layout.addWidget(clear_log_btn)
        
        log_group.setLayout(log_layout)
        bottom_layout.addWidget(log_group)
        
        splitter.addWidget(bottom_widget)
        splitter.setSizes([500, 400])
        
        self.radio_title.toggled.connect(self.on_enum_method_changed)
        self.radio_process.toggled.connect(self.on_enum_method_changed)
        self.radio_both.toggled.connect(self.on_enum_method_changed)
        
        self.log_message("就绪 - 请设置枚举条件后点击刷新窗口列表")
        self.log_message("提示: 重复次数设置为0时，将无限循环发送，直到手动停止")
        self.log_message("提示: 发送过程中可以随时修改参数，修改后立即生效")
        self.log_message(f"提示: 按键类型支持{len(KEY_LIST)}种预设按键，也支持手动输入虚拟键码(如0x41)")
        self.log_message("提示: 支持组合键发送，可同时选择Ctrl、Alt、Shift、Win等多个修饰键")
        
    def get_selected_modifiers(self):
        """获取选中的组合键列表"""
        modifiers = []
        if self.cb_ctrl.isChecked():
            modifiers.append("Ctrl")
        if self.cb_alt.isChecked():
            modifiers.append("Alt")
        if self.cb_shift.isChecked():
            modifiers.append("Shift")
        if self.cb_win.isChecked():
            modifiers.append("Win")
        # Fn键是硬件级按键，无法通过软件模拟，这里只做提示
        if self.cb_fn.isChecked():
            self.log_message("⚠ 注意: Fn键是硬件级按键，大多数情况下无法通过软件模拟，已忽略")
        return modifiers
        
    def on_parameter_changed(self):
        """参数变化时，如果正在发送，则实时更新线程参数"""
        if self.key_thread and self.key_thread.isRunning():
            key_type = self.key_combo.currentText()
            method_index = self.method_combo.currentIndex()
            delay = self.delay_spin.value()
            repeat_count = self.repeat_spin.value()
            modifiers = self.get_selected_modifiers()
            
            self.key_thread.update_parameters(key_type, method_index, delay, repeat_count, modifiers)
            
            modifier_str = " + ".join(modifiers) if modifiers else "无"
            self.log_message(f"参数已实时更新 - 组合键:{modifier_str}, 按键:{key_type}, 方式:{method_index+1}, 间隔:{delay}秒, 重复:{repeat_count}")
        
    def on_enum_method_changed(self):
        """枚举方式改变时更新输入框状态"""
        if self.radio_title.isChecked():
            self.title_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setEnabled(False)
            self.process_keyword_edit.setPlaceholderText("按窗口标题模式，程序名关键词已禁用")
        elif self.radio_process.isChecked():
            self.title_keyword_edit.setEnabled(False)
            self.title_keyword_edit.setPlaceholderText("按程序名模式，窗口标题关键词已禁用")
            self.process_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setPlaceholderText("例如: zhengtu.dat 或 client.exe (支持部分匹配)")
        else:
            self.title_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setPlaceholderText("例如: zhengtu.dat 或 client.exe (支持部分匹配)")
            
    def get_enum_method(self):
        """获取当前枚举方式"""
        if self.radio_title.isChecked():
            return EnumMethod.BY_TITLE
        elif self.radio_process.isChecked():
            return EnumMethod.BY_PROCESS
        else:
            return EnumMethod.BY_BOTH
        
    def enumerate_windows(self):
        """枚举窗口"""
        console_print("\n" + "="*80)
        console_print("用户点击了刷新按钮")
        console_print("="*80)
        
        enum_method = self.get_enum_method()
        title_keyword = self.title_keyword_edit.text().strip()
        process_keyword = self.process_keyword_edit.text().strip()
        
        console_print(f"枚举方式: {enum_method}")
        console_print(f"标题关键词: '{title_keyword}'")
        console_print(f"程序名关键词: '{process_keyword}'")
        
        if enum_method == EnumMethod.BY_TITLE and not title_keyword:
            QMessageBox.warning(self, "警告", "请填写窗口标题关键词")
            return
        elif enum_method == EnumMethod.BY_PROCESS and not process_keyword:
            QMessageBox.warning(self, "警告", "请填写程序名关键词")
            return
        elif enum_method == EnumMethod.BY_BOTH and (not title_keyword or not process_keyword):
            QMessageBox.warning(self, "警告", "两者兼顾模式需要同时填写窗口标题关键词和程序名关键词")
            return
        
        self.refresh_btn.setEnabled(False)
        self.status_label.setText("正在枚举窗口...")
        self.log_text.clear()
        
        method_name = "窗口标题" if enum_method == EnumMethod.BY_TITLE else \
                     "程序名" if enum_method == EnumMethod.BY_PROCESS else "两者兼顾"
        self.log_message(f"开始枚举窗口 - 方式:{method_name}")
        self.log_message(f"  标题关键词: '{title_keyword}'")
        self.log_message(f"  程序名关键词: '{process_keyword}'")
        
        if self.debug_checkbox.isChecked():
            self.log_message("🔧 调试模式已开启，将显示所有窗口的详细信息")
        
        self.enum_thread = EnhancedWindowEnumThread(
            enum_method, 
            title_keyword, 
            process_keyword,
            show_debug=self.debug_checkbox.isChecked()
        )
        self.enum_thread.windows_found.connect(self.on_windows_found)
        self.enum_thread.progress_updated.connect(self.on_enum_progress)
        self.enum_thread.debug_log.connect(self.log_message)
        self.enum_thread.start()
        
    def on_windows_found(self, windows):
        """窗口枚举完成"""
        console_print(f"\n找到 {len(windows)} 个匹配窗口")
        
        self.window_list.clear()
        self.window_items.clear()
        
        for window in windows:
            rect = window['rect']
            size_info = f"{rect[2]-rect[0]}x{rect[3]-rect[1]}"
            process_name = window.get('process_name', '未知')
            modules = window.get('modules', [])
            
            console_print(f"匹配窗口: '{window['title'][:50]}' | PID:{window['pid']} | 进程:{process_name} | 模块:{modules[:2]}")
            
            if modules:
                module_info = f" 模块: {', '.join(modules[:2])}"
            else:
                module_info = ""
                
            item = QListWidgetItem(
                f"{window['title']} | [{window['class']}] | PID:{window['pid']} | 进程:{process_name}{module_info} | 尺寸:{size_info}"
            )
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            item.setData(Qt.ItemDataRole.UserRole, window['hwnd'])
            item.setData(Qt.ItemDataRole.UserRole + 1, window['pid'])
            self.window_list.addItem(item)
            self.window_items.append({
                'hwnd': window['hwnd'],
                'title': window['title'],
                'class': window['class'],
                'pid': window['pid'],
                'rect': window['rect'],
                'process_name': process_name,
                'modules': modules,
                'item': item
            })
        
        if windows:
            self.status_label.setText(f"找到 {len(windows)} 个窗口")
            self.log_message(f"✓ 找到 {len(windows)} 个符合条件的窗口")
        else:
            self.status_label.setText("未找到符合条件的窗口")
            self.log_message("✗ 未找到符合条件的窗口")
            self.log_message("  建议:")
            self.log_message("    1. 检查关键词是否正确")
            self.log_message("    2. 开启调试模式查看所有窗口的进程名和模块信息")
            self.log_message("    3. 尝试使用更短的关键词，如 'zhengtu' 而不是 'zhengtu.dat'")
            self.log_message("    4. 确认目标程序正在运行且窗口可见")
            self.log_message("    5. 建议以管理员权限运行程序")
            
        self.refresh_btn.setEnabled(True)
        
    def on_enum_progress(self, progress):
        """枚举进度更新"""
        self.progress_bar.setValue(progress)
        
    def select_all_windows(self):
        """全选窗口"""
        for i in range(self.window_list.count()):
            item = self.window_list.item(i)
            item.setCheckState(Qt.CheckState.Checked)
        self.log_message("已全选所有窗口")
        
    def deselect_all_windows(self):
        """取消全选"""
        for i in range(self.window_list.count()):
            item = self.window_list.item(i)
            item.setCheckState(Qt.CheckState.Unchecked)
        self.log_message("已取消全选")
        
    def get_selected_windows(self):
        """获取选中的窗口句柄"""
        selected = []
        for i in range(self.window_list.count()):
            item = self.window_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                hwnd = item.data(Qt.ItemDataRole.UserRole)
                selected.append(hwnd)
        return selected
        
    def start_sending_keys(self):
        """开始发送按键"""
        selected_windows = self.get_selected_windows()
        
        if not selected_windows:
            QMessageBox.warning(self, "警告", "请至少选择一个窗口")
            return
            
        key_type = self.key_combo.currentText()
        delay = self.delay_spin.value()
        method_index = self.method_combo.currentIndex()
        repeat = self.repeat_spin.value()
        modifiers = self.get_selected_modifiers()
        
        self.log_text.clear()
        
        modifier_str = " + ".join(modifiers) if modifiers else "无组合键"
        if repeat == 0:
            self.log_message(f"开始测试 - 组合键:{modifier_str}, 按键:{key_type}, 方式:{method_index+1}, 间隔:{delay}秒, 重复:无限循环")
        else:
            self.log_message(f"开始测试 - 组合键:{modifier_str}, 按键:{key_type}, 方式:{method_index+1}, 间隔:{delay}秒, 重复:{repeat}次")
        self.log_message(f"目标窗口数: {len(selected_windows)}")
        self.log_message("提示: 发送过程中可以随时修改参数，修改后立即生效")
        if modifiers:
            self.log_message("提示: 组合键模式下，按键将以组合键形式发送（如 Ctrl+C）")
        
        self.key_thread = EnhancedKeyPressThread()
        self.key_thread.selected_windows = selected_windows
        self.key_thread.update_parameters(key_type, method_index, delay, repeat, modifiers)
        
        self.key_thread.progress_updated.connect(self.on_key_progress)
        self.key_thread.status_updated.connect(self.on_key_status)
        self.key_thread.log_message.connect(self.log_message)
        self.key_thread.finished.connect(self.on_key_finished)
        
        self.key_thread.start()
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.refresh_btn.setEnabled(False)
        # 不锁定控件，允许用户在发送过程中修改参数
        
    def stop_sending_keys(self):
        """停止发送按键"""
        if self.key_thread and self.key_thread.isRunning():
            self.key_thread.stop()
            self.key_thread.wait()
            self.log_message("用户手动停止发送")
            self.status_label.setText("已停止")
            
    def on_key_progress(self, progress):
        """按键进度更新"""
        # 无限循环模式下，进度条显示当前循环轮数
        if self.key_thread and self.key_thread.repeat_count == 0:
            # 无限模式，进度条作为脉冲显示
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)  # 设置为脉冲模式
        else:
            if self.progress_bar.maximum() != 100:
                self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(progress)
        
    def on_key_status(self, status):
        """按键状态更新"""
        self.status_label.setText(status)
        
    def on_key_finished(self):
        """按键完成"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.refresh_btn.setEnabled(True)
        
        # 恢复进度条为正常模式
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        
        if self.key_thread and not self.key_thread.running:
            self.log_message("按键发送已停止")
            self.status_label.setText("发送已停止")
        else:
            self.progress_bar.setValue(100)
            self.log_message("按键发送完成！")
            self.status_label.setText("发送完成")
            
    def log_message(self, message):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        
    def save_settings(self):
        """保存设置"""
        self.settings.setValue("key_type", self.key_combo.currentText())
        self.settings.setValue("delay", self.delay_spin.value())
        self.settings.setValue("method_index", self.method_combo.currentIndex())
        self.settings.setValue("repeat_count", self.repeat_spin.value())
        self.settings.setValue("enum_method", self.get_enum_method())
        self.settings.setValue("title_keyword", self.title_keyword_edit.text())
        self.settings.setValue("process_keyword", self.process_keyword_edit.text())
        self.settings.setValue("debug_mode", self.debug_checkbox.isChecked())
        
        # 保存组合键状态
        self.settings.setValue("modifier_ctrl", self.cb_ctrl.isChecked())
        self.settings.setValue("modifier_alt", self.cb_alt.isChecked())
        self.settings.setValue("modifier_shift", self.cb_shift.isChecked())
        self.settings.setValue("modifier_win", self.cb_win.isChecked())
        self.settings.setValue("modifier_fn", self.cb_fn.isChecked())
        
        window_states = []
        for item_data in self.window_items:
            window_states.append({
                'hwnd': item_data['hwnd'],
                'checked': item_data['item'].checkState() == Qt.CheckState.Checked
            })
        self.settings.setValue("window_states", json.dumps(window_states))
        
    def load_settings(self):
        """加载设置"""
        if self.settings.contains("key_type"):
            self.key_combo.setCurrentText(self.settings.value("key_type"))
        if self.settings.contains("delay"):
            self.delay_spin.setValue(float(self.settings.value("delay")))
        if self.settings.contains("method_index"):
            self.method_combo.setCurrentIndex(int(self.settings.value("method_index")))
        if self.settings.contains("repeat_count"):
            self.repeat_spin.setValue(int(self.settings.value("repeat_count")))
        if self.settings.contains("title_keyword"):
            self.title_keyword_edit.setText(self.settings.value("title_keyword"))
        if self.settings.contains("process_keyword"):
            self.process_keyword_edit.setText(self.settings.value("process_keyword"))
        if self.settings.contains("debug_mode"):
            self.debug_checkbox.setChecked(self.settings.value("debug_mode") == "true" or self.settings.value("debug_mode") == True)
        
        # 加载组合键状态
        if self.settings.contains("modifier_ctrl"):
            self.cb_ctrl.setChecked(self.settings.value("modifier_ctrl") == "true" or self.settings.value("modifier_ctrl") == True)
        if self.settings.contains("modifier_alt"):
            self.cb_alt.setChecked(self.settings.value("modifier_alt") == "true" or self.settings.value("modifier_alt") == True)
        if self.settings.contains("modifier_shift"):
            self.cb_shift.setChecked(self.settings.value("modifier_shift") == "true" or self.settings.value("modifier_shift") == True)
        if self.settings.contains("modifier_win"):
            self.cb_win.setChecked(self.settings.value("modifier_win") == "true" or self.settings.value("modifier_win") == True)
        if self.settings.contains("modifier_fn"):
            self.cb_fn.setChecked(self.settings.value("modifier_fn") == "true" or self.settings.value("modifier_fn") == True)
        
        if self.settings.contains("enum_method"):
            enum_method = int(self.settings.value("enum_method"))
            if enum_method == EnumMethod.BY_TITLE:
                self.radio_title.setChecked(True)
            elif enum_method == EnumMethod.BY_PROCESS:
                self.radio_process.setChecked(True)
            else:
                self.radio_both.setChecked(True)
        
        self.on_enum_method_changed()
        self.window_states_json = self.settings.value("window_states") if self.settings.contains("window_states") else None
        
    def apply_window_states(self):
        """应用窗口状态"""
        if not self.window_states_json:
            return
            
        try:
            window_states = json.loads(self.window_states_json)
            state_map = {state['hwnd']: state['checked'] for state in window_states}
            
            for item_data in self.window_items:
                if item_data['hwnd'] in state_map:
                    item_data['item'].setCheckState(
                        Qt.CheckState.Checked if state_map[item_data['hwnd']] 
                        else Qt.CheckState.Unchecked
                    )
        except Exception as e:
            self.log_message(f"加载窗口状态失败: {e}")
            
    def closeEvent(self, event):
        """关闭事件"""
        self.save_settings()
        if self.key_thread and self.key_thread.isRunning():
            self.key_thread.stop()
            self.key_thread.wait()
        event.accept()

def main():
    console_print("="*80)
    console_print("程序启动")
    console_print("="*80)
    
    # 检查管理员权限
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        console_print(f"管理员权限: {is_admin}")
        if not is_admin:
            console_print("⚠ 提示: 建议以管理员权限运行程序，以便更好地获取进程信息")
    except:
        pass
    
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    font = QFont("Microsoft YaHei", 9)
    app.setFont(font)
    
    window = EnhancedMainWindow()
    window.show()
    
    console_print("窗口已显示，等待用户操作...")
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()