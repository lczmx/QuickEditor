import os
import time
import base64
import logging
import webbrowser
from typing import Optional
from threading import Thread
from queue import Queue, Full
import tkinter as tk
from tkinter.messagebox import askyesno

import pyautogui
import keyboard
import win32api
import win32con
import win32gui
import win32gui_struct

PROJECT_TITLE = "快捷编辑"
PROJECT_URL = r"https://www.baidu.com"
VERSION = 1.0

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0
IsOpen = True
Status_Str = None

logging.basicConfig(
    filename=os.path.join(os.getcwd(), 'quickEditor.log'),
    level=logging.DEBUG,
    format="[%(asctime)s] - %(levelname)s - %(lineno)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


class ProcessEvent:
    """
    处理热键对应的事件
    """

    def __init__(self):
        # 键映射，方便后续处理
        # 提示： 用pyautogui.KEYBOARD_KEYS，可以查看pyautogui可以模拟的键
        self._key_maps = {
            "k": "up",
            "j": "down",
            "h": "left",
            "l": "right",
            "a": "home",
            "e": "end",
            "o": "backspace",
            "p": "delete",
            ",": "pageup",
            ".": "pagedown",
            "u": "ctrl-left",
            "i": "ctrl-right",
            "n": "ctrl-backspace",
            "m": "ctrl-delete"

        }

    @staticmethod
    def process_one_key(key: str):
        """
        处理单个键
        :param key:
        :return:
        """
        pyautogui.press(key)

    @staticmethod
    def process_hotkey(hot_key):
        """
        处理hot_key
        :param key:
        :return:
        """
        pyautogui.hotkey(*hot_key)

    @staticmethod
    def process_write(content):
        """
        处理直接写入
        :param content: 要写入的内容
        :return:
        """
        pyautogui.typewrite(content)

    def move_cursor(self, key):
        """
        移动光标

        :param key: _key_maps的key
        :return:
        """
        direction = self._key_maps[key]
        self.process_one_key(direction)

    def delete(self, key):
        """
        删除
        分别向前与后删除
        :param key: _key_maps的key
        :return:
        """
        self.process_one_key(self._key_maps[key])

    def change_page(self, key):
        """
        翻页
        PageUp/PageDown
        :param key: _key_maps的key
        :return:
        """
        page_direction = self._key_maps[key]
        self.process_one_key(page_direction)

    def multiple_space(self, key):
        """
        写入多个空格
        :param key: _key_maps的key
        :return:
        """
        num = int(key)
        self.process_write(" " * num)

    def move_cursor_base_word(self, key):
        """
        根据单词移动光标
        :param key: _key_maps的key
        :return:
        """
        hotkeys = self._key_maps[key].split("-")
        self.process_hotkey(hotkeys)

    def delete_word(self, key):
        """
        移动光标到下一行
        :param key: _key_maps的key
        :return:
        """
        hotkeys = self._key_maps[key].split("-")
        self.process_hotkey(hotkeys)

    @staticmethod
    def switch(key):
        """
        打开或关闭程序
        :return:
        """
        global IsOpen
        IsOpen = not IsOpen
        status = "正在运行" if IsOpen else "已暂停"
        message = "已经打开" if IsOpen else "已经关闭" + PROJECT_TITLE
        logging.info(message)
        # 异步在 gui中更新
        if Main.queue.empty():
            try:
                Main.queue.put(status, block=False)
            except Full:
                pass


class Monitor:
    """
    监听处理函数类
    """

    def __init__(self):
        # 定义热键字典
        # 用于反射

        """
        _hotKeys
        key: 键盘上的键, 多个键时使用两个空格表示
        val: 对应的处理函数
        """
        self._hotKeys = {
            "k  j  h  l  a  e": "move_cursor",  # 通过单个键移动光标
            "u  i": "move_cursor_base_word",  # 通过热键实现
            "o  p": "delete",
            ",  .": "change_page",
            "2  3  4  5  6  7  8  9": "multiple_space",
            "1": "switch",
            "n  m": "delete_word",
        }

        # 用于存放hook对像，后面用于删除
        self._event_hook_obj = []

        # 判断是否是想输入热键，processed会在process中被修改
        self._processed = False
        self.processEvent = ProcessEvent()
        #
        self._space_num = 0

    def process(self, key_obj):
        """
         执行除space外的其他键的处理函数
         具体处理函数在ProcessEvent类中
        :param key_obj:
        :return:
        """
        for k, v in self._hotKeys.items():
            v = v.strip()
            if str(key_obj.name) in k:
                # 反射执行
                if hasattr(self.processEvent, v):
                    func = getattr(self.processEvent, v)
                    func(key_obj.name)
                    logging.info(f"processed {v}({key_obj.name})")
                    break

        self._processed = True

    # space按下
    def key_down_callback(self, key_obj):
        try:
            global IsOpen
            print("key_down_callback: %s" % IsOpen)
            # 一直按着时，不忽略就会报错
            if not self._event_hook_obj and IsOpen:
                key_string = "  ".join([k for k in self._hotKeys])
                keys = [i for i in key_string.split("  ") if i]
                for w in keys:
                    # 按下space的同时, 按下要监听的键

                    down_obj = keyboard.on_press_key(w, self.process, suppress=True)
                    self._event_hook_obj.append(down_obj)

            # 用于恢复正常的激活状态
            if not IsOpen:
                try:
                    # 可能重复添加
                    if not self._event_hook_obj:
                        self.clear_listen()
                    self._event_hook_obj.append(keyboard.on_press_key("1", self.process, suppress=True))
                    # 释放除space外的监听
                except Exception as e:
                    logging.error(str(e))
                    IsOpen = True

        except Exception as e:
            logging.error(str(e))
            # 执行错误时, 默认输出空格, 所以需要except异常
            self._processed = False

    def key_up_callback(self, key_obj):
        self.clear_listen()
        # 只想输入空格
        if not self._processed:
            pyautogui.press('space')

        self._processed = False

    def clear_listen(self):
        """清空监听列表, 并在keyboard模块中取消监听"""
        while self._event_hook_obj:
            # 清空监听
            obj = self._event_hook_obj.pop()
            keyboard.unhook(obj)


class SysTrayIcon(object):
    """
    系统托盘
    """
    QUIT = 'QUIT'
    TOGGLE = "TOGGLE"
    SPECIAL_ACTIONS = [QUIT, TOGGLE]
    FIRST_ID = 1314

    def __init__(s,
                 icon,
                 hover_text,
                 menu_options,
                 on_quit=None,
                 default_menu_index=None,
                 window_class_name=None, ):

        s.icon = icon
        s.hover_text = hover_text
        s.on_quit = on_quit
        menu_options = menu_options + (("切换程序运行状态", None, s.TOGGLE), ('退出', None, s.QUIT))
        s._next_action_id = s.FIRST_ID
        s.menu_actions_by_id = set()
        s.menu_options = s._add_ids_to_menu_options(list(menu_options))
        s.menu_actions_by_id = dict(s.menu_actions_by_id)
        del s._next_action_id

        s.default_menu_index = (default_menu_index or 0)
        s.window_class_name = window_class_name or "SysTrayIconPy"

        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): s.refresh_icon,
                       win32con.WM_DESTROY: s.destroy,
                       win32con.WM_COMMAND: s.command,
                       win32con.WM_USER + 20: s.notify, }
        # 注册窗口类。
        window_class = win32gui.WNDCLASS()
        window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = s.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map  # 也可以指定wndproc.
        s.classAtom = win32gui.RegisterClass(window_class)

    def show_icon(s):
        # 创建窗口。
        hinst = win32gui.GetModuleHandle(None)
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        s.hwnd = win32gui.CreateWindow(s.classAtom,
                                       s.window_class_name,
                                       style,
                                       0,
                                       0,
                                       win32con.CW_USEDEFAULT,
                                       win32con.CW_USEDEFAULT,
                                       0,
                                       0,
                                       hinst,
                                       None)
        win32gui.UpdateWindow(s.hwnd)
        s.notify_id = None
        s.refresh_icon()

        win32gui.PumpMessages()

    def show_menu(s):
        menu = win32gui.CreatePopupMenu()
        s.create_menu(menu, s.menu_options)
        # win32gui.SetMenuDefaultItem(menu, 1000, 0)

        pos = win32gui.GetCursorPos()
        # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
        win32gui.SetForegroundWindow(s.hwnd)
        win32gui.TrackPopupMenu(menu,
                                win32con.TPM_LEFTALIGN,
                                pos[0],
                                pos[1],
                                0,
                                s.hwnd,
                                None)
        win32gui.PostMessage(s.hwnd, win32con.WM_NULL, 0, 0)

    def destroy(s, hwnd, msg, wparam, lparam):
        if s.on_quit: s.on_quit(s)  # 运行传递的on_quit
        nid = (s.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # 退出托盘图标

    def notify(s, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:  # 双击左键
            pass  # s.execute_menu_option(s.default_menu_index + s.FIRST_ID)
        elif lparam == win32con.WM_RBUTTONUP:  # 单击右键
            s.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:  # 单击左键
            nid = (s.hwnd, 0)
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
            win32gui.PostQuitMessage(0)  # 退出托盘图标
            if Main:
                Main.root.deiconify()
        return True

    def _add_ids_to_menu_options(s, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if callable(option_action) or option_action in s.SPECIAL_ACTIONS:
                s.menu_actions_by_id.add((s._next_action_id, option_action))
                result.append(menu_option + (s._next_action_id,))
            else:
                result.append((option_text,
                               option_icon,
                               s._add_ids_to_menu_options(option_action),
                               s._next_action_id))
            s._next_action_id += 1
        return result

    def refresh_icon(s, **data):
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(s.icon):  # 尝试找到自定义图标
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst,
                                       s.icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
        else:  # 找不到图标文件 - 使用默认值
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if s.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        s.notify_id = (s.hwnd,
                       0,
                       win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                       win32con.WM_USER + 20,
                       hicon,
                       s.hover_text)
        win32gui.Shell_NotifyIcon(message, s.notify_id)

    def create_menu(s, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = s.prep_menu_icon(option_icon)

            if option_id in s.menu_actions_by_id:
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                s.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

    def prep_menu_icon(s, icon):
        # 首先加载图标。
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hdcBitmap = win32gui.CreateCompatibleDC(0)
        hdcScreen = win32gui.GetDC(0)
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)
        # 填满背景。
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
        # "GetSysColorBrush返回缓存的画笔而不是分配新的画笔。"
        #  - 暗示没有DeleteObject
        # 画出图标
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        win32gui.SelectObject(hdcBitmap, hbmOld)
        win32gui.DeleteDC(hdcBitmap)

        return hbm

    def command(s, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        s.execute_menu_option(id)

    def execute_menu_option(s, id):
        """
        点击菜单, 然后执行对应函数
        :param id:
        :return:
        """
        menu_action = s.menu_actions_by_id[id]
        if menu_action == s.QUIT:
            win32gui.DestroyWindow(s.hwnd)
        elif menu_action == s.TOGGLE:
            # 切换状态
            ProcessEvent.switch("1")

        else:
            menu_action(s)


class App:
    """GUI程序"""

    def __init__(self):
        self.queue = Queue(maxsize=1)  # "正在运行" or "已暂停"

    def main(self):
        icon = os.path.join(os.getcwd(), r'icon.ico')
        # 初始化图标
        if not os.path.isfile(icon):
            img = b'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Ev/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhH825cQMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA25cRlduWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9qVEuXalg8iAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pUSdNuWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/2pYR39mZDRQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2JMPNNuWEvbblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uW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mP/d3d3/3d3d/6+vr/8cHBz/09PT/93d3f/d3d3/bW1t/zs7O//d3d3/3d3d/93d3f8yMjL/eXl5/93d3f/d3d3/w8PD/xMTE/+tra3/3d3d/93d3f+Pj4//HR0d/9ra2v/d3d3/3d3d/93d3f95eXn/kWYT/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhLUAAAAANyVEjrblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/y4mFP/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/b29v/3d3d/93d3f/d3d3/3d3d/9ra2v/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/a2tr/3d3d/93d3f/d3d3/2tra/93d3f/d3d3/3d3d/93d3f/b29v/3d3d/93d3f/d3d3/3d3d/9ra2v/d3d3/3d3d/93d3f/d3d3/eXl5/5FmE//blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS1AAAAADemBAv25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv8uJhT/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/3l5ef+RZhP/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEb8AAAAA1pgOJduWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/LiYU/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f95eXn/kWYT/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhKqAAAAAM+PEBDblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/y4mFP/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/eXl5/5FmE//blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YSjwAAAAAAAAAA25US79uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv8uJhT/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/93d3f/d3d3/3d3d/3l5ef+RZhP/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9qVEm8AAAAAAAAAANuVEc/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/RzYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/ywkFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/LiYU/y4mFP8uJhT/nW4S/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blBBPAAAAAAAAAADalRKf25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv9gRhT/qnYS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/2JMPIQAAAAAAAAAA25URW9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/YEYU/6p2Ev/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/2pYR3wAAAAAAAAAAAAAAANqWDyLblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/2BGFP+qdhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uVEaQAAAAAAAAAAAAAAAAAAAAA25YRz9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv9gRhT/qnYS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/ZlhBQAAAAAAAAAAAAAAAAAAAAANuWEXrblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/YEYU/6p2Ev/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhHu6KIACwAAAAAAAAAAAAAAAAAAAADdmA4l25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/2BGFP+qdhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/2pYRpQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANuWEbLblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv9gRhT/qnYS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9eWDzMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADXlg8z25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/YEYU/6p2Ev/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEbIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANuVEbDblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/3pXE/+JYRP/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEf3blA8yAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADYkA0n25US9tuWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/KixL/Oi0U/2xOE/+YaxP/rXgS/76DEv+3fhL/rHcS/7qAEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blRCdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqWEYrblhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/82MEv+NYxP/aUwT/1dAFP9JNxT/RzYU/0c2FP9CMhT/OSwU/59vE//blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blRHv2JMKGgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADXlA0T2pYR39uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/8KGEv8nIRT/voMS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/2ZUScwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANiTDzTblhL225YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/n28S/2VJE//blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/2pUSqwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pUSdNuWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9GPEv87LhT/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/2pYR39mZDRQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADblxGV25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uWEv/blhL/25YS/9uW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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADMmQAF2JMKGt2ZEQ8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=='

            tmp = open(icon, "wb+")
            tmp.write(base64.b64decode(img))
            tmp.close()
        # ########################      tkinter界面设定      #####################################
        window = tk.Tk()
        window.protocol('WM_DELETE_WINDOW', self.close_window)
        window.title(PROJECT_TITLE)
        window.iconbitmap(icon)
        window.geometry('250x300')
        window.resizable(0, 0)

        status = "程序当前状态: " + "正在运行" if IsOpen else "已暂停"
        self.status_Str = tk.StringVar(value=status)
        tk.Label(window, textvariable=self.status_Str, font=('微软雅黑', 10, ''), padx=10, pady=10, justify="left").pack()

        tk.Label(window, text="关于", font=('微软雅黑', 20, 'bold'), padx=10, pady=10, justify="left").pack()
        tk.Message(window, text="这是一个避免手在键盘上的上下左右等键区域左右移动，并提高码字的质量的工具",
                   font=('微软雅黑', 10, ''), padx=10, pady=30, justify="left").pack(side="top")

        def open_url(event):
            webbrowser.open_new(PROJECT_URL)

        label = tk.Label(window, text="使用方式见GitHub (点击这里)", font=('微软雅黑', 10, ''), padx=10, pady=10, justify="left")

        label.bind("<Button-1>", open_url)

        label.pack(side="top")
        tk.Label(window, text=f"版本: {VERSION}", font=('微软雅黑', 9, ''),
                 padx=10, pady=30, justify="left").pack(side="top")
        # ##########################     开始托盘程序嵌入     #####################################
        self.root = window

        hover_text = PROJECT_TITLE  # 悬浮于图标上方时的提示
        menu_options = ()
        self.sysTrayIcon = SysTrayIcon(icon, hover_text, menu_options, on_quit=self.exit, default_menu_index=1)

        self.root.bind("<Unmap>", lambda event: self.unmap() if self.root.state() == 'iconic' else False)
        # s.root.protocol('WM_DELETE_WINDOW', s.exit)
        # 默认最小化
        self.root.resizable(0, 0)
        self.root.state("iconic")
        self.root.mainloop()

    def switch_icon(self, _sysTrayIcon, icons="./icon.ico"):
        _sysTrayIcon.icon = icons
        _sysTrayIcon.refresh_icon()
        # 点击右键菜单项目会传递SysTrayIcon自身给引用的函数，所以这里的_sysTrayIcon = s.sysTrayIcon

    def unmap(self):
        self.root.withdraw()
        self.sysTrayIcon.show_icon()

    def exit(self, _sysTrayIcon=None):
        self.root.destroy()
        logging.info('exit...')

    def close_window(self):
        ans = askyesno(title='提示', message='你确定要退出吗?')
        if ans:
            self.exit()
        else:
            return


def toggle_string():
    """
    最小化时会阻塞, 所以只能这样了
    异步更新文本
    :return:
    """
    logging.debug("listening toggle string")
    while True:
        if Main:
            status = Main.queue.get()
            while True:
                # 最小化时抛出异常 main thread is not in main loop
                # 所以需要重复执行代码
                try:

                    Main.status_Str.set(f"程序当前状态: {status}")
                    break
                except Exception as e:
                    logging.error(str(e))

            time.sleep(.5)


def start_monitor():
    mon = Monitor()
    # 监听space按下
    keyboard.on_press_key("space", mon.key_down_callback, suppress=True)
    # 监听space释放
    keyboard.on_release_key("space", mon.key_up_callback, suppress=True)
    keyboard.wait()


def start_gui():
    global Main
    # 打开GUI
    Main = App()
    Main.main()


if __name__ == '__main__':
    Main: Optional[App] = None
    monitor = Thread(target=start_monitor)
    monitor.setDaemon(True)
    monitor.start()

    tg_str = Thread(target=toggle_string)
    tg_str.setDaemon(True)
    tg_str.start()

    gui = Thread(target=start_gui)
    gui.start()
    gui.join()
