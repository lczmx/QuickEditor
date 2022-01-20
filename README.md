# 概述
这是一个避免手在键盘上的上下左右等键区域左右移动，并提高码字的质量的工具  
利用通过空格键+其他键实现`Home`、`End`、`PageUp`、`PageDown`等功能, 自以为最大的作用是，避免右手在键盘上的左右移动，提高码字的质量  


# 键位图
通过空格键+其他键实现, 键位图如下:
![示意图](https://gitee.com/lczmx/Note/raw/master/pictures/QuickEditor文档/1642689980521.png)

# 如何自定义?
主要需要用到两个类: `Monitor`和`ProcessEvent`, 前者用于监听键盘, 后者用于定义监听的回调函数  
- `Monitor`类的初始化`__init__()`中的`_hotKeys`属性中添加要监听的键+回调函数
    如:
    ```python
    class Monitor:
        def __init__(self):
            """
            _hotKeys
            key: 键盘上的键, 多个键时使用两个空格表示
            val: 对应的处理函数
            """
            self._hotKeys = {
                "k  j": "move_cursor",  # 通过单个键移动光标
            }
    ```
    > 注：回调函数要写在`ProcessEvent`类中

- `ProcessEvent`类中定义回调函数, 需要接收一个`key`参数
    如:
    ```python

    class ProcessEvent:
        def __init__(self):
            # 键映射，方便后续处理
            # 提示： 用pyautogui.KEYBOARD_KEYS，可以查看pyautogui可以模拟的键
            self._key_maps = {
                "k": "up",
                "j": "down",
            }

        @staticmethod
        def process_one_key(key: str):
            """
            处理单个键
            :param key:
            :return:
            """
            pyautogui.press(key)

        def move_cursor(self, key):
            """
            移动光标

            :param key: _key_maps的key
            :return:
            """
            direction = self._key_maps[key]
            self.process_one_key(direction)
    ```

# 最后
假如有什么想法。欢迎提交你们的代码
