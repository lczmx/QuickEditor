import pyautogui
import keyboard

pyautogui.FAILSAFE = False
ISOPEN = True

class ProcessEvent:
    """
    处理热键对应的事件
    space + i 上（up），以下省略space
          + k 下（down）
          + j 左（left）
          + l 右（right）
          + a 行首（Home）
          + f 行尾（End）
          + u 左删除（Backspace）
          + o 右删除（Delete）
          + e 上一页（PageUp）
          + d 下一页（PageDown）
          + 2~9 一次输入多少个空格
          + 1 关闭热键,长按空格5秒左右可以重新打开
    """

    def __init__(self):
        # 键映射，方便后续处理
        # 提示： 用pyautogui.KEYBOARD_KEYS，可以查看pyautogui可以模拟的键
        self._key_maps = {
            "i": "up",
            "k": "down",
            "j": "left",
            "l": "right",
            "a": "home",
            "f": "end",
            "u": "backspace",
            "o": "delete",
            "e": "pageup",
            "d": "pagedown",


        }

    def move_cursor(self, key, distance=1):
        """
        :param key: 键盘上的键
        :param distance: 移动的距离
        :return:
        """
        direction = self._key_maps[key]
        for i in range(distance):
            pyautogui.press(direction)

    def delete(self, key):
        direction = self._key_maps[key]
        pyautogui.press(direction)

    def change_page(self, key):
        direction = self._key_maps[key]
        pyautogui.press(direction)

    def multiple_space(self, key):
        num = int(key)
        pyautogui.typewrite(" " * num)
    def switch(self, key):
        global ISOPEN
        ISOPEN = not ISOPEN
        print("已经关闭")


class Monitor:
    def __init__(self):
        # 定义热键字典
        # 用于反射
        self._hotKeys = {
            "ikjlaf": "move_cursor",
            "uo": "delete",
            "ed": "change_page",
            "23456789": "multiple_space",
            "1": "switch",
        }

        # 用于存放hook对像，后面用于删除
        self._event_hook_obj = []

        # 判断是否是想输入热键，processed会在process中被修改
        self._processed = False
        self.processEvent = ProcessEvent()
        self._space_num = 0

    def process(self, key_obj):
        for k, v in self._hotKeys.items():
            if key_obj.name in k:
                # 反射执行
                if hasattr(self.processEvent, v):
                    func = getattr(self.processEvent, v)
                    func(key_obj.name)
                    print("processed %s" % v)
                    break

        self._processed = True

    # space按下
    def key_down_callback(self, key_obj):
        global ISOPEN
        # 一直按着时，不忽略就会报错
        if not self._event_hook_obj and ISOPEN:
            for w in "".join([k for k in self._hotKeys]):
                down_obj = keyboard.on_press_key(w, self.process, suppress=True)
                self._event_hook_obj.append(down_obj)
        # 用于恢复正常的激活状态
        if not ISOPEN and self._space_num > 150:
            ISOPEN = True
            self._space_num = 0
            self._processed = True
            print("已经激活")
        else:
            self._space_num += 1



    # space弹起
    # 释放除space外的监听
    def key_up_callback(self, key_obj):

        while self._event_hook_obj:
            # 清空监听
            obj = self._event_hook_obj.pop()
            keyboard.unhook(obj)
        # 只想输入空格
        if not self._processed:
            pyautogui.press('space')

        self._processed = False


if __name__ == '__main__':
    mon = Monitor()
    keyboard.on_press_key("space", mon.key_down_callback, suppress=True)
    keyboard.on_release_key("space", mon.key_up_callback, suppress=True)
    keyboard.wait()
