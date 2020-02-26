# HotKey
自定义快捷键实现Home、End、PageUp、PageDown等功能。<br />
自以为最大的作用是，避免右手在键盘上的左右移动，提高码字的质量。当然你也可以实现其他功能

## 功能
处理热键对应的事件<br />
    space + i 上（up），以下省略space<br />
         + k 下（down）<br />
          + j 左（left）<br />
          + l 右（right）<br />
          + a 行首（Home）<br />
          + f 行尾（End）<br />
          + u 左删除（Backspace）<br />
          + o 右删除（Delete）<br />
          + e 上一页（PageUp）<br />
          + d 下一页（PageDown）<br />
          + 2~9 一次输入多少个空格<br />
          + 1 关闭热键,长按空格5秒左右可以重新打开<br /><br />
 <b>支持自定义功能</b><br />
## 使用
为了减少秃头的可能性我没有写GUI<br />
所以使用方法<br />
cd HotKey.py后<br />
python HotKey.py<br />
<b>嗯？？？</b><br />
没错我也没有编译，主要是我是在另一个脚本上打开这个脚本的，设置好后cd+打开，一个命令完成，实在顶不住的人可以自行编译<br />
## 环境
pyautogui库和keyboard库<br />
没有安装的人需要手动pip安装：<br />
pip install pyautogui<br />
pip install keyboard<br />
建议国内-i使用豆瓣源
## 如何自定义?
1、Monitor类中的初始化__init__()中的_hotKeys属性中添加要监听的键+回调函数<br />
<b>注：回调函数要写在ProcessEvent类中，且有两个形参</b><br />
2、在ProcessEvent类中写回调函数，一般在函数中利用pyautogui模拟鼠标或键盘输入<br />
以下给出一个大佬的pyautogui博客<br />
[博客地址](https://asyncfor.com/posts/doc-pyautogui.html "点击查看")  <br />
### 最后
个人觉得还有一些功能没有实现，但又觉得没有什么想法。欢迎提交你们的代码<br />
由于不经常上GitHub，所以可以加我的qq：2691948831
