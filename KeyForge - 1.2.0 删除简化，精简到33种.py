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
                               QButtonGroup)
from PySide6.QtCore import Qt, QSettings, Signal, QThread, QTimer
from PySide6.QtGui import QIcon, QFont, QColor


# 全局调试标志 - 编译后设置为False，不再输出到控制台
DEBUG_TO_CONSOLE = False  # 编译后设置为False，避免控制台错误

def console_print(*args, **kwargs):
    """输出到控制台 - 编译后不执行任何操作"""
    if DEBUG_TO_CONSOLE:
        try:
            print(*args, **kwargs)
            sys.stdout.flush()
        except:
            pass

# 后台按键发送方式枚举 - 只保留完整实现的方法
class BackendKeyMethod:
    # 基础消息方式 (0-9)
    SENDMESSAGE = 0          # SendMessage WM_KEYDOWN/UP
    POSTMESSAGE = 1          # PostMessage WM_KEYDOWN/UP
    KEYBD_EVENT = 2          # keybd_event (全局)
    SENDINPUT = 3            # SendInput (高级输入)
    WM_KEYDOWN_DIRECT = 4    # WM_KEYDOWN直接发送
    WM_CHAR = 5              # WM_CHAR字符消息
    WM_IME_CHAR = 6          # WM_IME_CHAR输入法字符
    WM_SYSKEYDOWN = 7        # WM_SYSKEYDOWN系统按键
    POSTMESSAGE_LPARAM = 8   # PostMessage带完整lParam
    
    # 高级消息方式 (10-19)
    WM_HOTKEY = 10           # WM_HOTKEY热键消息
    WM_APPCOMMAND = 11       # WM_APPCOMMAND应用程序命令
    SENDMESSAGE_TIMEOUT = 12 # SendMessageTimeout
    BROADCAST_MSG = 13       # BroadcastSystemMessage
    
    # 输入模拟方式 (20-29)
    ATTACH_THREAD = 20       # AttachThreadInput方式
    SCAN_CODE = 21           # 扫描码方式
    EXTENDED_KEY = 22        # 扩展键标志
    
    # 特殊应用方式 (30-39)
    WM_COPYDATA = 30         # WM_COPYDATA数据复制
    CLIPBOARD = 31           # 剪贴板方式
    COM_OBJECT = 32          # COM对象调用

    # 获取方式总数
    @classmethod
    def get_total_count(cls):
        """获取所有方式的数量"""
        return 33
    
    @classmethod
    def get_method_name(cls, method):
        """获取发送方式的名称"""
        method_names = {
            cls.SENDMESSAGE: "SendMessage (WM_KEY)",
            cls.POSTMESSAGE: "PostMessage (WM_KEY)",
            cls.KEYBD_EVENT: "keybd_event (全局)",
            cls.SENDINPUT: "SendInput (高级输入)",
            cls.WM_KEYDOWN_DIRECT: "WM_KEYDOWN直接发送",
            cls.WM_CHAR: "WM_CHAR字符消息",
            cls.WM_IME_CHAR: "WM_IME_CHAR输入法字符",
            cls.WM_SYSKEYDOWN: "WM_SYSKEYDOWN系统按键",
            cls.POSTMESSAGE_LPARAM: "PostMessage完整参数",
            cls.WM_HOTKEY: "WM_HOTKEY热键消息",
            cls.WM_APPCOMMAND: "WM_APPCOMMAND命令",
            cls.SENDMESSAGE_TIMEOUT: "SendMessageTimeout",
            cls.BROADCAST_MSG: "BroadcastSystemMessage",
            cls.ATTACH_THREAD: "AttachThreadInput",
            cls.SCAN_CODE: "扫描码方式",
            cls.EXTENDED_KEY: "扩展键标志",
            cls.WM_COPYDATA: "WM_COPYDATA",
            cls.CLIPBOARD: "剪贴板方式",
            cls.COM_OBJECT: "COM对象调用",
        }
        return method_names.get(method, f"未知方式({method})")

# 窗口枚举方式枚举
class EnumMethod:
    BY_TITLE = 0      # 按窗口标题
    BY_PROCESS = 1    # 按程序名
    BY_BOTH = 2       # 两者兼顾（标题和程序名都匹配）
    BY_CLASS = 3      # 按窗口类名
    BY_PID = 4        # 按进程ID
    BY_HANDLE = 5     # 按窗口句柄

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
    "F11": 0x7A, "F12": 0x7B, "F13": 0x7C, "F14": 0x7D, "F15": 0x7E,
    "F16": 0x7F, "F17": 0x80, "F18": 0x81, "F19": 0x82, "F20": 0x83,
    "F21": 0x84, "F22": 0x85, "F23": 0x86, "F24": 0x87,
    
    # 控制键
    "Space": 0x20, "Enter": 0x0D, "Esc": 0x1B, "Tab": 0x09,
    "Backspace": 0x08, "Delete": 0x2E, "Insert": 0x2D,
    "Home": 0x24, "End": 0x23, "PageUp": 0x21, "PageDown": 0x22,
    "Print": 0x2C, "Pause": 0x13, "Cancel": 0x03, "Clear": 0x0C,
    "Select": 0x29, "Execute": 0x2B, "Help": 0x2F,
    
    # 方向键
    "Up": 0x26, "Down": 0x28, "Left": 0x25, "Right": 0x27,
    
    # 修饰键
    "Ctrl": 0x11, "Alt": 0x12, "Shift": 0x10, "Win": 0x5B,
    "RWin": 0x5C, "Apps": 0x5D,
    
    # 小键盘
    "Num0": 0x60, "Num1": 0x61, "Num2": 0x62, "Num3": 0x63,
    "Num4": 0x64, "Num5": 0x65, "Num6": 0x66, "Num7": 0x67,
    "Num8": 0x68, "Num9": 0x69,
    "NumMult": 0x6A, "NumPlus": 0x6B, "NumMinus": 0x6D,
    "NumDot": 0x6E, "NumDiv": 0x6F,
    "NumEnter": 0x0D, "NumLock": 0x90,
    
    # 多媒体键
    "VolumeUp": 0xAF, "VolumeDown": 0xAE, "VolumeMute": 0xAD,
    "MediaNext": 0xB0, "MediaPrev": 0xB1, "MediaStop": 0xB2,
    "MediaPlayPause": 0xB3, "LaunchMail": 0xB4, "LaunchMedia": 0xB5,
    "LaunchApp1": 0xB6, "LaunchApp2": 0xB7,
    
    # 浏览器键
    "BrowserBack": 0xA6, "BrowserForward": 0xA7, "BrowserRefresh": 0xA8,
    "BrowserStop": 0xA9, "BrowserSearch": 0xAA, "BrowserFavorites": 0xAB,
    "BrowserHome": 0xAC,
    
    # 其他常用键
    "CapsLock": 0x14, "PrintScreen": 0x2C, "ScrollLock": 0x91,
    "PauseBreak": 0x13, "Menu": 0x5D, "Sleep": 0x5F,
    
    # 符号键
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
    "RWin": 0x5C,
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
    "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", "F24",
    # 方向键
    "Up", "Down", "Left", "Right",
    # 控制键
    "Space", "Enter", "Esc", "Tab", "Backspace", "Delete", "Insert",
    "Home", "End", "PageUp", "PageDown", "Print", "Pause", "Cancel", "Clear",
    "Select", "Execute", "Help",
    # 修饰键
    "Ctrl", "Alt", "Shift", "Win", "RWin", "Apps",
    # 小键盘
    "Num0", "Num1", "Num2", "Num3", "Num4", "Num5", "Num6", "Num7", "Num8", "Num9",
    "NumMult", "NumPlus", "NumMinus", "NumDot", "NumDiv", "NumEnter", "NumLock",
    # 多媒体键
    "VolumeUp", "VolumeDown", "VolumeMute", "MediaNext", "MediaPrev",
    "MediaStop", "MediaPlayPause", "LaunchMail", "LaunchMedia", "LaunchApp1", "LaunchApp2",
    # 浏览器键
    "BrowserBack", "BrowserForward", "BrowserRefresh", "BrowserStop",
    "BrowserSearch", "BrowserFavorites", "BrowserHome",
    # 符号键
    "`", "-", "=", "[", "]", "\\", ";", "'", ",", ".", "/",
    # 其他
    "CapsLock", "PrintScreen", "ScrollLock", "PauseBreak", "Menu", "Sleep"
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
        pid = win32process.GetWindowThreadProcessId(hwnd)[0]
        result['pid'] = pid
        
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        kernel32 = ctypes.windll.kernel32
        psapi = ctypes.windll.psapi
        
        hProcess = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        
        if hProcess:
            exe_name = ctypes.create_unicode_buffer(260)
            if psapi.GetModuleBaseNameW(hProcess, None, exe_name, 260):
                result['process_name'] = exe_name.value.lower()
            
            exe_path = ctypes.create_unicode_buffer(260)
            if psapi.GetModuleFileNameExW(hProcess, None, exe_path, 260):
                result['process_path'] = exe_path.value.lower()
            
            kernel32.CloseHandle(hProcess)
        else:
            try:
                process = psutil.Process(pid)
                result['process_name'] = process.name().lower()
                result['process_path'] = process.exe().lower() if process.exe() else ''
            except:
                pass
    except Exception as e:
        console_print(f"获取进程信息异常: {e}")
    
    return result


# ==================== 项目信息元数据 ====================
class ProjectInfo:
    """项目信息元数据（集中管理所有项目相关信息）"""
    VERSION = "1.2.0"
    BUILD_DATE = "2026-03-26"
    AUTHOR = "杜玛"
    LICENSE = "GNU Affero General Public License v3.0"
    COPYRIGHT = "© 永久 杜玛"
    URL = "https://github.com/duma520/KeyForge"
    MAINTAINER_EMAIL = "不提供"
    NAME = "KeyForge 后台按键测试工具"
    DESCRIPTION = "KeyForge 后台按键测试工具，支持多种按键发送方式、多种窗口枚举方式，并提供丰富的调试信息输出。"
    
    @classmethod
    def get_full_name(cls) -> str:
        return f"{cls.NAME} v{cls.VERSION}"
    
    @classmethod
    def get_full_title(cls, username: str = None) -> str:
        base = f"{cls.get_full_name()} 构建:{cls.BUILD_DATE}"
        if username:
            return f"{base} - 当前用户: {username}"
        return base
    
    @classmethod
    def get_about_text(cls) -> str:
        return f"""
        <h2>{cls.NAME}</h2>
        <p><b>版本:</b> {cls.VERSION}</p>
        <p><b>构建日期:</b> {cls.BUILD_DATE}</p>
        <p><b>作者:</b> {cls.AUTHOR}</p>
        <p><b>版权:</b> {cls.COPYRIGHT}</p>
        <p><b>许可证:</b> {cls.LICENSE}</p>
        <p><b>项目主页:</b> <a href='{cls.URL}'>{cls.URL}</a></p>
        <p><b>描述:</b> {cls.DESCRIPTION}</p>
        <p><b>支持按键方式:</b> {BackendKeyMethod.get_total_count()} 种</p>
        """


# 增强的按键发送类
class EnhancedKeyPressThread(QThread):
    progress_updated = Signal(int)
    status_updated = Signal(str)
    log_message = Signal(str)
    loop_count_updated = Signal(int)
    
    def __init__(self):
        super().__init__()
        self.running = False
        self.selected_windows = []
        self.key_type = "A"
        self.method = BackendKeyMethod.SENDMESSAGE
        self.delay = 1.0
        self.repeat_count = 1
        
        self.modifiers = []
        
        self.param_lock = threading.Lock()
        self.current_loop = 0
        self.total_windows = 0
        
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
        return KEY_MAP.get(key_type, 0x41)
    
    def get_modifier_codes(self, modifiers):
        """获取组合键的虚拟键码列表"""
        codes = []
        for mod in modifiers:
            if mod in MODIFIER_KEYS:
                codes.append(MODIFIER_KEYS[mod])
        return codes
    
    def get_modifier_flags(self):
        """获取组合键的MOD_标志"""
        flags = 0
        if "Alt" in self.modifiers:
            flags |= 0x0001
        if "Ctrl" in self.modifiers:
            flags |= 0x0002
        if "Shift" in self.modifiers:
            flags |= 0x0004
        if "Win" in self.modifiers:
            flags |= 0x0008
        return flags
    
    def send_keybd_event(self, vk_code, is_down):
        """使用keybd_event发送按键"""
        flags_val = 0 if is_down else win32con.KEYEVENTF_KEYUP
        ctypes.windll.user32.keybd_event(vk_code, 0, flags_val, 0)
    
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
    
    def send_attach_thread_input(self, hwnd, vk_code, is_down):
        """使用AttachThreadInput方式"""
        try:
            current_thread = win32api.GetCurrentThreadId()
            target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
            
            win32process.AttachThreadInput(current_thread, target_thread, True)
            self.send_keybd_event(vk_code, is_down)
            win32process.AttachThreadInput(current_thread, target_thread, False)
            return True
        except Exception as e:
            self.log_message.emit(f"AttachThreadInput错误: {e}")
            return False
    
    def send_wm_hotkey(self, hwnd, vk_code, is_down):
        """使用WM_HOTKEY消息"""
        if is_down:
            hotkey_id = 1
            try:
                win32gui.RegisterHotKey(hwnd, hotkey_id, self.get_modifier_flags(), vk_code)
                win32gui.PostMessage(hwnd, win32con.WM_HOTKEY, hotkey_id, 0)
                win32gui.UnregisterHotKey(hwnd, hotkey_id)
            except:
                pass
        return True
    
    def send_wm_appcommand(self, hwnd, vk_code, is_down):
        """使用WM_APPCOMMAND消息"""
        if is_down:
            win32gui.PostMessage(hwnd, win32con.WM_APPCOMMAND, 0, vk_code << 16)
        return True
    
    def send_broadcast_msg(self, vk_code, is_down):
        """广播系统消息"""
        if is_down:
            win32gui.BroadcastSystemMessage(
                win32con.BSF_IGNORECURRENTTASK | win32con.BSF_POSTMESSAGE,
                0, win32con.WM_KEYDOWN, vk_code, 0)
        return True
    
    def send_scan_code(self, hwnd, vk_code, is_down):
        """扫描码方式"""
        try:
            scan_code = ctypes.windll.user32.MapVirtualKeyW(vk_code, 0)
            lParam = (scan_code << 16) | (1 if is_down else 0xC0000000)
            if is_down:
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam)
            else:
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam)
            return True
        except Exception as e:
            self.log_message.emit(f"扫描码方式错误: {e}")
            return False
    
    def send_extended_key(self, hwnd, vk_code, is_down):
        """扩展键标志方式"""
        try:
            scan_code = ctypes.windll.user32.MapVirtualKeyW(vk_code, 0)
            extended = 1
            lParam = (1 & 0xFFFF) | ((scan_code & 0xFF) << 16) | ((extended & 1) << 24)
            if is_down:
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam)
            else:
                lParam |= (1 << 30) | (1 << 31)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam)
            return True
        except Exception as e:
            self.log_message.emit(f"扩展键标志错误: {e}")
            return False
    
    def send_wm_copydata(self, hwnd, vk_code, is_down):
        """使用WM_COPYDATA发送"""
        if is_down:
            class COPYDATASTRUCT(ctypes.Structure):
                _fields_ = [
                    ("dwData", ctypes.c_ulong),
                    ("cbData", ctypes.c_ulong),
                    ("lpData", ctypes.c_void_p)
                ]
            
            data = struct.pack('I', vk_code)
            cds = COPYDATASTRUCT()
            cds.dwData = 0x4B455900
            cds.cbData = len(data)
            cds.lpData = ctypes.addressof(ctypes.create_string_buffer(data))
            
            win32gui.SendMessage(hwnd, win32con.WM_COPYDATA, 0, cds)
        return True
    
    def send_clipboard(self, key_type):
        """通过剪贴板发送文本"""
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(key_type)
            win32clipboard.CloseClipboard()
            
            for hwnd in self.selected_windows:
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, 0x11, 0)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, 0x56, 0)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, 0x56, 0)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, 0x11, 0)
            return True
        except Exception as e:
            self.log_message.emit(f"剪贴板方式错误: {e}")
            return False
    
    def send_com_object(self, vk_code, is_down):
        """COM对象调用"""
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            if is_down:
                if 32 <= vk_code <= 126:
                    shell.SendKeys(chr(vk_code))
                else:
                    shell.SendKeys('{' + self.key_type.upper() + '}')
            return True
        except Exception as e:
            self.log_message.emit(f"COM对象错误: {e}")
            return False
    
    # ==================== 不带组合键的发送方法 ====================
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
                self.send_keybd_event(vk_code, True)
                time.sleep(0.02)
                self.send_keybd_event(vk_code, False)
                return True
                
            elif method == BackendKeyMethod.SENDINPUT:
                self.send_sendinput(vk_code, True)
                time.sleep(0.02)
                self.send_sendinput(vk_code, False)
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
                
            elif method == BackendKeyMethod.WM_HOTKEY:
                return self.send_wm_hotkey(hwnd, vk_code, True)
                
            elif method == BackendKeyMethod.WM_APPCOMMAND:
                return self.send_wm_appcommand(hwnd, vk_code, True)
                
            elif method == BackendKeyMethod.SENDMESSAGE_TIMEOUT:
                ctypes.windll.user32.SendMessageTimeoutW(
                    hwnd, win32con.WM_KEYDOWN, vk_code, 0, 0x0002, 1000, None)
                time.sleep(0.02)
                ctypes.windll.user32.SendMessageTimeoutW(
                    hwnd, win32con.WM_KEYUP, vk_code, 0, 0x0002, 1000, None)
                return True
                
            elif method == BackendKeyMethod.BROADCAST_MSG:
                return self.send_broadcast_msg(vk_code, True) and \
                       self.send_broadcast_msg(vk_code, False)
                
            elif method == BackendKeyMethod.ATTACH_THREAD:
                return self.send_attach_thread_input(hwnd, vk_code, True) and \
                       self.send_attach_thread_input(hwnd, vk_code, False)
                       
            elif method == BackendKeyMethod.SCAN_CODE:
                return self.send_scan_code(hwnd, vk_code, True) and \
                       self.send_scan_code(hwnd, vk_code, False)
                       
            elif method == BackendKeyMethod.EXTENDED_KEY:
                return self.send_extended_key(hwnd, vk_code, True) and \
                       self.send_extended_key(hwnd, vk_code, False)
                       
            elif method == BackendKeyMethod.WM_COPYDATA:
                return self.send_wm_copydata(hwnd, vk_code, True)
                
            elif method == BackendKeyMethod.CLIPBOARD:
                return self.send_clipboard(chr(vk_code) if 32 <= vk_code <= 126 else str(vk_code))
                
            elif method == BackendKeyMethod.COM_OBJECT:
                return self.send_com_object(vk_code, True) and \
                       self.send_com_object(vk_code, False)
                
            else:
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                time.sleep(0.02)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                return True
                    
        except Exception as e:
            self.log_message.emit(f"错误: {str(e)}")
            return False
    
    # ==================== 带组合键的发送方法 ====================
    def send_key_with_modifiers(self, hwnd, vk_code, method, char_code=0):
        """发送带组合键的按键"""
        try:
            modifier_codes = self.get_modifier_codes(self.modifiers)
            
            if not modifier_codes:
                return self.send_key_without_modifiers(hwnd, vk_code, method, char_code)
            
            # 全局输入模拟方法
            if method in [BackendKeyMethod.KEYBD_EVENT, BackendKeyMethod.SENDINPUT]:
                for mod_code in modifier_codes:
                    self.send_keybd_event(mod_code, True)
                    time.sleep(0.005)
                
                self.send_keybd_event(vk_code, True)
                time.sleep(0.02)
                self.send_keybd_event(vk_code, False)
                time.sleep(0.005)
                
                for mod_code in reversed(modifier_codes):
                    self.send_keybd_event(mod_code, False)
                    time.sleep(0.005)
                return True
            
            # 窗口消息发送方法
            elif method in [BackendKeyMethod.SENDMESSAGE, BackendKeyMethod.POSTMESSAGE,
                            BackendKeyMethod.WM_KEYDOWN_DIRECT, BackendKeyMethod.POSTMESSAGE_LPARAM,
                            BackendKeyMethod.WM_SYSKEYDOWN, BackendKeyMethod.SENDMESSAGE_TIMEOUT]:
                
                def make_lparam(repeat, scan, extended, context, previous):
                    return (repeat & 0xFFFF) | ((scan & 0xFF) << 16) | \
                           ((extended & 1) << 24) | ((context & 1) << 29) | \
                           ((previous & 1) << 30)
                
                for mod_code in modifier_codes:
                    if method == BackendKeyMethod.SENDMESSAGE:
                        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, mod_code, 0)
                    elif method == BackendKeyMethod.SENDMESSAGE_TIMEOUT:
                        ctypes.windll.user32.SendMessageTimeoutW(
                            hwnd, win32con.WM_KEYDOWN, mod_code, 0, 0x0002, 1000, None)
                    else:
                        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, mod_code, 0)
                    time.sleep(0.005)
                
                if method == BackendKeyMethod.SENDMESSAGE:
                    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                    time.sleep(0.02)
                    win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
                elif method == BackendKeyMethod.SENDMESSAGE_TIMEOUT:
                    ctypes.windll.user32.SendMessageTimeoutW(
                        hwnd, win32con.WM_KEYDOWN, vk_code, 0, 0x0002, 1000, None)
                    time.sleep(0.02)
                    ctypes.windll.user32.SendMessageTimeoutW(
                        hwnd, win32con.WM_KEYUP, vk_code, 0, 0x0002, 1000, None)
                elif method == BackendKeyMethod.BROADCAST_MSG:
                    win32gui.BroadcastSystemMessage(
                        win32con.BSF_IGNORECURRENTTASK | win32con.BSF_POSTMESSAGE,
                        0, win32con.WM_KEYDOWN, vk_code, 0)
                    time.sleep(0.02)
                    win32gui.BroadcastSystemMessage(
                        win32con.BSF_IGNORECURRENTTASK | win32con.BSF_POSTMESSAGE,
                        0, win32con.WM_KEYUP, vk_code, 0)
                else:
                    lParam = make_lparam(1, 0, 0, 0, 0)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam)
                    time.sleep(0.02)
                    lParam_up = make_lparam(1, 0, 0, 1, 1)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam_up)
                
                time.sleep(0.005)
                
                for mod_code in reversed(modifier_codes):
                    if method == BackendKeyMethod.SENDMESSAGE:
                        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, mod_code, 0)
                    elif method == BackendKeyMethod.SENDMESSAGE_TIMEOUT:
                        ctypes.windll.user32.SendMessageTimeoutW(
                            hwnd, win32con.WM_KEYUP, mod_code, 0, 0x0002, 1000, None)
                    else:
                        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, mod_code, 0)
                    time.sleep(0.005)
                return True
            
            # 扫描码/扩展键方式
            elif method in [BackendKeyMethod.SCAN_CODE, BackendKeyMethod.EXTENDED_KEY]:
                scan_code = ctypes.windll.user32.MapVirtualKeyW(vk_code, 0)
                extended = 1 if method == BackendKeyMethod.EXTENDED_KEY else 0
                
                def make_lparam_scan(repeat, scan, extended, context, previous):
                    return (repeat & 0xFFFF) | ((scan & 0xFF) << 16) | \
                           ((extended & 1) << 24) | ((context & 1) << 29) | \
                           ((previous & 1) << 30)
                
                for mod_code in modifier_codes:
                    mod_scan = ctypes.windll.user32.MapVirtualKeyW(mod_code, 0)
                    lParam_mod = make_lparam_scan(1, mod_scan, 0, 0, 0)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, mod_code, lParam_mod)
                    time.sleep(0.005)
                
                lParam = make_lparam_scan(1, scan_code, extended, 0, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam)
                time.sleep(0.02)
                lParam_up = make_lparam_scan(1, scan_code, extended, 1, 1)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam_up)
                time.sleep(0.005)
                
                for mod_code in reversed(modifier_codes):
                    mod_scan = ctypes.windll.user32.MapVirtualKeyW(mod_code, 0)
                    lParam_mod_up = make_lparam_scan(1, mod_scan, 0, 1, 1)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, mod_code, lParam_mod_up)
                    time.sleep(0.005)
                return True
            
            # 线程附加方式
            elif method == BackendKeyMethod.ATTACH_THREAD:
                try:
                    current_thread = win32api.GetCurrentThreadId()
                    target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
                    
                    win32process.AttachThreadInput(current_thread, target_thread, True)
                    old_focus = win32gui.GetFocus()
                    win32gui.SetFocus(hwnd)
                    
                    for mod_code in modifier_codes:
                        self.send_keybd_event(mod_code, True)
                        time.sleep(0.005)
                    
                    self.send_keybd_event(vk_code, True)
                    time.sleep(0.02)
                    self.send_keybd_event(vk_code, False)
                    time.sleep(0.005)
                    
                    for mod_code in reversed(modifier_codes):
                        self.send_keybd_event(mod_code, False)
                        time.sleep(0.005)
                    
                    if old_focus:
                        win32gui.SetFocus(old_focus)
                    
                    win32process.AttachThreadInput(current_thread, target_thread, False)
                    return True
                except Exception as e:
                    self.log_message.emit(f"线程附加方式错误: {e}")
                    return False
            
            # 字符消息方式
            elif method == BackendKeyMethod.WM_CHAR:
                if char_code:
                    win32gui.PostMessage(hwnd, win32con.WM_CHAR, char_code, 0)
                    return True
                else:
                    return self.send_key_with_modifiers(hwnd, vk_code, BackendKeyMethod.POSTMESSAGE, char_code)
            
            elif method == BackendKeyMethod.WM_IME_CHAR:
                if char_code:
                    win32gui.PostMessage(hwnd, win32con.WM_IME_CHAR, char_code, 0)
                else:
                    win32gui.PostMessage(hwnd, win32con.WM_IME_CHAR, vk_code, 0)
                return True
            
            # 热键方式
            elif method == BackendKeyMethod.WM_HOTKEY:
                hotkey_id = 1
                try:
                    win32gui.RegisterHotKey(hwnd, hotkey_id, self.get_modifier_flags(), vk_code)
                    win32gui.PostMessage(hwnd, win32con.WM_HOTKEY, hotkey_id, 0)
                    win32gui.UnregisterHotKey(hwnd, hotkey_id)
                    return True
                except:
                    return self.send_key_with_modifiers(hwnd, vk_code, BackendKeyMethod.POSTMESSAGE, char_code)
            
            # COM对象方式
            elif method == BackendKeyMethod.COM_OBJECT:
                try:
                    import win32com.client
                    shell = win32com.client.Dispatch("WScript.Shell")
                    
                    modifier_map = {"Ctrl": "^", "Alt": "%", "Shift": "+", "Win": "#"}
                    key_str = ""
                    for mod in self.modifiers:
                        if mod in modifier_map:
                            key_str += modifier_map[mod]
                    
                    if 32 <= vk_code <= 126:
                        key_str += chr(vk_code)
                    else:
                        key_str += "{" + self.key_type.upper() + "}"
                    
                    shell.SendKeys(key_str)
                    return True
                except Exception as e:
                    self.log_message.emit(f"COM对象方式错误: {e}")
                    return False
            
            else:
                return self.send_key_with_modifiers(hwnd, vk_code, BackendKeyMethod.POSTMESSAGE, char_code)
                    
        except Exception as e:
            self.log_message.emit(f"组合键发送错误: {str(e)}")
            return False
    
    # ==================== 主发送方法 ====================
    def send_key_to_window_advanced(self, hwnd, key_type, method):
        """增强的按键发送方法（支持组合键）"""
        vk_code = self.get_vk_code(key_type)
        char_code = 0
        if len(key_type) == 1:
            char_code = ord(key_type[0])
        
        has_modifiers = len(self.modifiers) > 0
        
        modifier_str = ""
        if has_modifiers:
            modifier_str = " + ".join(self.modifiers) + " + "
        
        if has_modifiers:
            success = self.send_key_with_modifiers(hwnd, vk_code, method, char_code)
            if success:
                self.log_message.emit(f"{BackendKeyMethod.get_method_name(method)}: 已发送组合键 [{modifier_str}{key_type}]")
        else:
            success = self.send_key_without_modifiers(hwnd, vk_code, method, char_code)
            if success:
                self.log_message.emit(f"{BackendKeyMethod.get_method_name(method)}: 已发送 {key_type} 键")
        
        return success
    
    def run(self):
        self.running = True
        
        key_type, method, delay, repeat_count, modifiers = self.get_parameters()
        is_infinite = (repeat_count == 0)
        
        self.total_windows = len(self.selected_windows)
        if is_infinite:
            total_operations = 0
        else:
            total_operations = self.total_windows * repeat_count
            current_operation = 0
        
        self.current_loop = 0
        while self.running:
            if not is_infinite and self.current_loop >= repeat_count:
                break
            
            key_type, method, delay, repeat_count, modifiers = self.get_parameters()
            is_infinite = (repeat_count == 0)
            
            for idx, hwnd in enumerate(self.selected_windows):
                if not self.running:
                    break
                    
                success = self.send_key_to_window_advanced(hwnd, key_type, method)
                
                if not is_infinite:
                    current_operation += 1
                    progress_percent = int(current_operation / total_operations * 100)
                    self.progress_updated.emit(progress_percent)
                
                if success:
                    modifier_str = " + ".join(modifiers) + " + " if modifiers else ""
                    self.status_updated.emit(f"成功发送 {modifier_str}{key_type} 到窗口 {idx+1}/{self.total_windows} (第{self.current_loop+1}轮)")
                else:
                    self.status_updated.emit(f"发送失败到窗口 {idx+1}/{self.total_windows}")
                    
                time.sleep(delay)
            
            self.current_loop += 1
            
            if is_infinite:
                self.status_updated.emit(f"已完成 {self.current_loop} 轮发送，继续循环...")
                self.log_message.emit(f"已完成第 {self.current_loop} 轮发送（无限模式）")
                self.progress_updated.emit(self.current_loop)
                self.loop_count_updated.emit(self.current_loop)
        
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
        if self.show_debug:
            self.debug_log.emit(message)
        
    def run(self):
        windows = []
        total_visible = 0
        matched_count = 0
        
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
                
                process_info = get_all_process_names_from_hwnd(hwnd)
                process_name = process_info['process_name']
                process_path = process_info['process_path']
                
                if self.show_debug and total_visible <= 100:
                    debug_info = f"窗口: '{window_text[:40]}' | 类名: '{class_name}' | PID:{pid} | 进程: '{process_name}'"
                    self.debug_print(debug_info)
                
                matched = False
                
                if self.enum_method == EnumMethod.BY_TITLE:
                    if self.title_keyword and self.title_keyword in window_text.lower():
                        matched = True
                        
                elif self.enum_method == EnumMethod.BY_PROCESS:
                    if self.process_keyword:
                        if self.process_keyword in process_name:
                            matched = True
                        elif process_path and self.process_keyword in os.path.basename(process_path).lower():
                            matched = True
                        elif self.process_keyword in class_name.lower():
                            matched = True
                            
                elif self.enum_method == EnumMethod.BY_BOTH:
                    title_match = self.title_keyword and self.title_keyword in window_text.lower()
                    process_match = False
                    
                    if self.process_keyword:
                        if (self.process_keyword in process_name or
                            (process_path and self.process_keyword in os.path.basename(process_path).lower()) or
                            self.process_keyword in class_name.lower()):
                            process_match = True
                    
                    if title_match and process_match:
                        matched = True
                
                if matched:
                    matched_count += 1
                    windows.append({
                        'hwnd': hwnd,
                        'title': window_text,
                        'class': class_name,
                        'rect': rect,
                        'pid': pid,
                        'process_name': process_name,
                        'process_path': process_path,
                    })
            return True
            
        win32gui.EnumWindows(enum_callback, None)
        
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
        
        self.key_combo.currentTextChanged.connect(self.on_parameter_changed)
        self.delay_spin.valueChanged.connect(self.on_parameter_changed)
        self.method_combo.currentIndexChanged.connect(self.on_parameter_changed)
        self.repeat_spin.valueChanged.connect(self.on_parameter_changed)
        
        self.cb_ctrl.stateChanged.connect(self.on_parameter_changed)
        self.cb_alt.stateChanged.connect(self.on_parameter_changed)
        self.cb_shift.stateChanged.connect(self.on_parameter_changed)
        self.cb_win.stateChanged.connect(self.on_parameter_changed)
        
    def init_ui(self):
        self.setWindowTitle(f"{ProjectInfo.NAME} {ProjectInfo.VERSION} (Build: {ProjectInfo.BUILD_DATE})")
        self.setMinimumSize(1200, 800)
        
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
        self.title_keyword_edit.setPlaceholderText("例如: 游戏窗口标题")
        enum_layout.addWidget(self.title_keyword_edit, 1, 1, 1, 3)
        
        enum_layout.addWidget(QLabel("程序名关键词:"), 2, 0)
        self.process_keyword_edit = QLineEdit()
        self.process_keyword_edit.setPlaceholderText("例如: zhengtu.dat 或 client.exe")
        enum_layout.addWidget(self.process_keyword_edit, 2, 1, 1, 3)
        
        self.debug_checkbox = QCheckBox("调试模式 (显示所有窗口信息)")
        self.debug_checkbox.setChecked(True)
        enum_layout.addWidget(self.debug_checkbox, 3, 0, 1, 2)
        
        tip_label = QLabel("提示: 程序名支持匹配进程名、路径和窗口类名")
        tip_label.setStyleSheet("color: #666; font-size: 9pt;")
        enum_layout.addWidget(tip_label, 3, 2, 1, 2)
        
        button_layout = QHBoxLayout()
        self.refresh_btn = QPushButton("刷新窗口列表")
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
        
        settings_group = QGroupBox(f"按键设置（{BackendKeyMethod.get_total_count()}种方式 + 组合键支持）")
        settings_layout = QGridLayout()
        
        # 组合键区域
        settings_layout.addWidget(QLabel("组合键:"), 0, 0)
        modifier_layout = QHBoxLayout()
        
        self.cb_ctrl = QCheckBox("Ctrl")
        self.cb_alt = QCheckBox("Alt")
        self.cb_shift = QCheckBox("Shift")
        self.cb_win = QCheckBox("Win")
        
        modifier_layout.addWidget(self.cb_ctrl)
        modifier_layout.addWidget(self.cb_alt)
        modifier_layout.addWidget(self.cb_shift)
        modifier_layout.addWidget(self.cb_win)
        modifier_layout.addStretch()
        
        settings_layout.addLayout(modifier_layout, 0, 1, 1, 5)
        
        tip_modifier = QLabel("提示: 可以同时选择多个组合键，如 Ctrl+Alt+A")
        tip_modifier.setStyleSheet("color: #666; font-size: 9pt;")
        settings_layout.addWidget(tip_modifier, 1, 1, 1, 5)
        
        # 按键类型
        settings_layout.addWidget(QLabel("按键类型:"), 2, 0)
        self.key_combo = QComboBox()
        self.key_combo.addItems(KEY_LIST)
        self.key_combo.setEditable(True)
        self.key_combo.setToolTip("支持选择预设按键，也可以手动输入虚拟键码（如0x41代表A）")
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
        self.repeat_spin.setRange(0, 100)
        self.repeat_spin.setValue(1)
        self.repeat_spin.setToolTip("0 = 无限循环，直到手动停止")
        settings_layout.addWidget(self.repeat_spin, 2, 5)
        
        settings_layout.addWidget(QLabel("后台按键方式:"), 3, 0)
        self.method_combo = QComboBox()
        methods = []
        for i in range(BackendKeyMethod.get_total_count()):
            method_name = BackendKeyMethod.get_method_name(i)
            methods.append(f"{i+1:2d}. {method_name}")
        self.method_combo.addItems(methods)
        self.method_combo.setMaxVisibleItems(20)
        settings_layout.addWidget(self.method_combo, 3, 1, 1, 3)
        
        method_tip = QLabel(f"共 {BackendKeyMethod.get_total_count()} 种后台按键方式，可通过滚动条查看所有方式")
        method_tip.setStyleSheet("color: #666; font-size: 9pt;")
        settings_layout.addWidget(method_tip, 3, 4, 1, 2)
        
        control_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始发送按键")
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
        
        self.stop_btn = QPushButton("停止发送")
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
        
        self.loop_label = QLabel("发送轮数: 0")
        self.loop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loop_label.setStyleSheet("color: #666; font-size: 10pt;")
        progress_layout.addWidget(self.loop_label)
        
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
        self.log_message(f"提示: 按键类型支持{len(KEY_LIST)}种预设按键，也支持手动输入虚拟键码")
        self.log_message("提示: 支持组合键发送，可同时选择Ctrl、Alt、Shift、Win等多个修饰键")
        self.log_message(f"提示: 当前支持 {BackendKeyMethod.get_total_count()} 种后台按键发送方式")
        
    def get_selected_modifiers(self):
        modifiers = []
        if self.cb_ctrl.isChecked():
            modifiers.append("Ctrl")
        if self.cb_alt.isChecked():
            modifiers.append("Alt")
        if self.cb_shift.isChecked():
            modifiers.append("Shift")
        if self.cb_win.isChecked():
            modifiers.append("Win")
        return modifiers
        
    def on_parameter_changed(self):
        if self.key_thread and self.key_thread.isRunning():
            key_type = self.key_combo.currentText()
            method_index = self.method_combo.currentIndex()
            delay = self.delay_spin.value()
            repeat_count = self.repeat_spin.value()
            modifiers = self.get_selected_modifiers()
            
            self.key_thread.update_parameters(key_type, method_index, delay, repeat_count, modifiers)
            
            modifier_str = " + ".join(modifiers) if modifiers else "无"
            method_name = BackendKeyMethod.get_method_name(method_index)
            self.log_message(f"参数已实时更新 - 组合键:{modifier_str}, 按键:{key_type}, 方式:{method_name}, 间隔:{delay}秒, 重复:{repeat_count}")
        
    def on_enum_method_changed(self):
        if self.radio_title.isChecked():
            self.title_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setEnabled(False)
            self.process_keyword_edit.setPlaceholderText("按窗口标题模式，程序名关键词已禁用")
        elif self.radio_process.isChecked():
            self.title_keyword_edit.setEnabled(False)
            self.title_keyword_edit.setPlaceholderText("按程序名模式，窗口标题关键词已禁用")
            self.process_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setPlaceholderText("例如: zhengtu.dat 或 client.exe")
        else:
            self.title_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setEnabled(True)
            self.process_keyword_edit.setPlaceholderText("例如: zhengtu.dat 或 client.exe")
            
    def get_enum_method(self):
        if self.radio_title.isChecked():
            return EnumMethod.BY_TITLE
        elif self.radio_process.isChecked():
            return EnumMethod.BY_PROCESS
        else:
            return EnumMethod.BY_BOTH
        
    def enumerate_windows(self):
        enum_method = self.get_enum_method()
        title_keyword = self.title_keyword_edit.text().strip()
        process_keyword = self.process_keyword_edit.text().strip()
        
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
            self.log_message("调试模式已开启，将显示所有窗口的详细信息")
        
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
        self.window_list.clear()
        self.window_items.clear()
        
        for window in windows:
            rect = window['rect']
            size_info = f"{rect[2]-rect[0]}x{rect[3]-rect[1]}"
            process_name = window.get('process_name', '未知')
                
            item = QListWidgetItem(
                f"{window['title']} | [{window['class']}] | PID:{window['pid']} | 进程:{process_name} | 尺寸:{size_info}"
            )
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            item.setData(Qt.ItemDataRole.UserRole, window['hwnd'])
            self.window_list.addItem(item)
            self.window_items.append({
                'hwnd': window['hwnd'],
                'title': window['title'],
                'item': item
            })
        
        if windows:
            self.status_label.setText(f"找到 {len(windows)} 个窗口")
            self.log_message(f"找到 {len(windows)} 个符合条件的窗口")
        else:
            self.status_label.setText("未找到符合条件的窗口")
            self.log_message("未找到符合条件的窗口")
            self.log_message("建议:")
            self.log_message("  1. 检查关键词是否正确")
            self.log_message("  2. 开启调试模式查看所有窗口信息")
            self.log_message("  3. 尝试使用更短的关键词")
            self.log_message("  4. 确认目标程序正在运行且窗口可见")
            
        self.refresh_btn.setEnabled(True)
        
    def on_enum_progress(self, progress):
        self.progress_bar.setValue(progress)
        
    def select_all_windows(self):
        for i in range(self.window_list.count()):
            item = self.window_list.item(i)
            item.setCheckState(Qt.CheckState.Checked)
        self.log_message("已全选所有窗口")
        
    def deselect_all_windows(self):
        for i in range(self.window_list.count()):
            item = self.window_list.item(i)
            item.setCheckState(Qt.CheckState.Unchecked)
        self.log_message("已取消全选")
        
    def get_selected_windows(self):
        selected = []
        for i in range(self.window_list.count()):
            item = self.window_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                hwnd = item.data(Qt.ItemDataRole.UserRole)
                selected.append(hwnd)
        return selected
        
    def start_sending_keys(self):
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
        self.loop_label.setText("发送轮数: 0")
        
        modifier_str = " + ".join(modifiers) if modifiers else "无组合键"
        method_name = BackendKeyMethod.get_method_name(method_index)
        if repeat == 0:
            self.log_message(f"开始测试 - 组合键:{modifier_str}, 按键:{key_type}, 方式:{method_name}, 间隔:{delay}秒, 重复:无限循环")
        else:
            self.log_message(f"开始测试 - 组合键:{modifier_str}, 按键:{key_type}, 方式:{method_name}, 间隔:{delay}秒, 重复:{repeat}次")
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
        self.key_thread.loop_count_updated.connect(self.on_loop_count_updated)
        
        self.key_thread.start()
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.refresh_btn.setEnabled(False)
        
    def on_loop_count_updated(self, loop_count):
        self.loop_label.setText(f"发送轮数: {loop_count}")
        
    def stop_sending_keys(self):
        if self.key_thread and self.key_thread.isRunning():
            self.key_thread.stop()
            self.key_thread.wait()
            self.log_message("用户手动停止发送")
            self.status_label.setText("已停止")
            
    def on_key_progress(self, progress):
        if self.key_thread and self.key_thread.repeat_count == 0:
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)
                if progress > 0:
                    self.progress_bar.setFormat(f"已发送 {progress} 轮")
        else:
            if self.progress_bar.maximum() != 100:
                self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(progress)
            self.progress_bar.setFormat("%p%")
        
    def on_key_status(self, status):
        self.status_label.setText(status)
        
    def on_key_finished(self):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.refresh_btn.setEnabled(True)
        
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        
        if self.key_thread and not self.key_thread.running:
            self.log_message("按键发送已停止")
            self.status_label.setText("发送已停止")
        else:
            self.progress_bar.setValue(100)
            self.log_message("按键发送完成！")
            self.status_label.setText("发送完成")
            
    def log_message(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        
    def save_settings(self):
        self.settings.setValue("key_type", self.key_combo.currentText())
        self.settings.setValue("delay", self.delay_spin.value())
        self.settings.setValue("method_index", self.method_combo.currentIndex())
        self.settings.setValue("repeat_count", self.repeat_spin.value())
        self.settings.setValue("enum_method", self.get_enum_method())
        self.settings.setValue("title_keyword", self.title_keyword_edit.text())
        self.settings.setValue("process_keyword", self.process_keyword_edit.text())
        self.settings.setValue("debug_mode", self.debug_checkbox.isChecked())
        
        self.settings.setValue("modifier_ctrl", self.cb_ctrl.isChecked())
        self.settings.setValue("modifier_alt", self.cb_alt.isChecked())
        self.settings.setValue("modifier_shift", self.cb_shift.isChecked())
        self.settings.setValue("modifier_win", self.cb_win.isChecked())
        
        window_states = []
        for item_data in self.window_items:
            window_states.append({
                'hwnd': item_data['hwnd'],
                'checked': item_data['item'].checkState() == Qt.CheckState.Checked
            })
        self.settings.setValue("window_states", json.dumps(window_states))
        
    def load_settings(self):
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
        
        if self.settings.contains("modifier_ctrl"):
            self.cb_ctrl.setChecked(self.settings.value("modifier_ctrl") == "true" or self.settings.value("modifier_ctrl") == True)
        if self.settings.contains("modifier_alt"):
            self.cb_alt.setChecked(self.settings.value("modifier_alt") == "true" or self.settings.value("modifier_alt") == True)
        if self.settings.contains("modifier_shift"):
            self.cb_shift.setChecked(self.settings.value("modifier_shift") == "true" or self.settings.value("modifier_shift") == True)
        if self.settings.contains("modifier_win"):
            self.cb_win.setChecked(self.settings.value("modifier_win") == "true" or self.settings.value("modifier_win") == True)
        
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
        self.save_settings()
        if self.key_thread and self.key_thread.isRunning():
            self.key_thread.stop()
            self.key_thread.wait()
        event.accept()

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    font = QFont("Microsoft YaHei", 9)
    app.setFont(font)
    
    window = EnhancedMainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()