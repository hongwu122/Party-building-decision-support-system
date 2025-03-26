
import random
import linecache
# 调用系统自带的语音识别模块
# import win32com.client
import win32com.client as win32

from tkinter import *
import threading
from tkinter import messagebox

# class MyThread(threading.Thread):
#     def __init__(self, func, *args):
#         super().__init__()
#
#         self.func = func
#         self.args = args
#
#         self.setDaemon(True)
#         self.start()  # 在这里开始
#
#     def run(self):
#         self.func(*self.args)


#自定义的线程函数类
def thread_it(func, *args):
  '''将函数放入线程中执行'''
  # 创建线程
  t = threading.Thread(target=func, args=args)
  # 守护线程
  t.setDaemon(True)
  # 启动线程
  t.start()



def xi_sayings():
    def find():
        txt = open(r'sayings.txt', 'rb')
        data = txt.read().decode('utf-8')  # python3一定要加上这句不然会编码报错！
        txt.close()

        # 获取txt的总行数！
        n = data.count('\n')
        print("总行数", n)
        # 选取随机的数
        i = random.randint(1, (n + 1))
        print("本次使用的行数", i)
        ###得到对应的i行的数据
        line = linecache.getline(r'sayings.txt', i)
        print(line)
        return line

    def say(line):
        try:
            # 1.创建一个播报对象
            speaker = win32.Dispatch("SAPI.SpVoice")
            # 2.通过这个播报器对象，直接播放对应的语音字符串就可以
            speaker.Speak(line)
        except:
            messagebox.showinfo('小提示', '很遗憾，语音播报模块不可用')

    line = find()
    window_xijp = Toplevel()
    window_xijp.title('学习党性语录')

    b = len (line) * 19 + 44
    print("len:  %s" % b)
    if b >= 1500:
        b = len (line)/2 * 19 + 44
        # root.geometry("%dx1+0+70" % b)
        window_xijp.geometry("%dx140+0+112" % b)   # x乘以y 加上坐标 x要等于b的字符串长度，y一行取80，两行取120
        print("新:  %d" % b)
        l = int(len(line)/2)
        line1 = line[:l]
        line2 = line[l:]
        a1 = Label(window_xijp, text=line1, font=("华文中宋", 14, "normal"), fg='red')
        a1.place(x=20, y=20)
        a2 = Label(window_xijp, text=line2, font=("华文中宋", 14, "normal"), fg='red')
        a2.place(x=20, y=60)
        button = Button(window_xijp, text="点击播放语录", width="10", command=lambda : thread_it(say(line)))
        button.place(x=20,y=100)
        label = Label(window_xijp,text='To YOU',font=('华文中宋',13)).place(x=int(int("%d"% b) - 93),y=110)
        # print(int(int("%d"% b)-90))
    else:
        # root.geometry("%dx1+0+20" % b)
        window_xijp.geometry("%dx97+0+112" % b)
        # 为了区别root和tl，我们向tl中添加了一个Label
        a = Label(window_xijp, text=line, font=("华文中宋", 14, "normal"), fg='red')
        a.place(x=20, y=20)
        button = Button(window_xijp, text="点击播放语录", width="10", command=lambda : thread_it(say(line)))
        button.place(x=20,y=60)
        label = Label(window_xijp, text='To YOU', font=('华文中宋', 13)).place(x=int(int("%d" % b) - 93), y=60)
        # print(int(int("%d" % b) - 90))
    window_xijp.mainloop()

    # root.mainloop()



# def xi_sayings():
#     root = Tk()
#     root.geometry("250x50+0+30")
#     root.title("主窗口(测试版)")
#     button = Button(root, text="点击刷新语录", width="10", command = lambda: thread_it(xijp) ,anchor = 'w',relief='raised').pack()
#     root.mainloop()


if __name__ == '__main__':
    xi_sayings()


'''
linecache模块常用函数及功能
函数基本格式	功能
linecache.getline(filename, lineno, module_globals=None)	读取指定模块中指定文件的指定行（仅读取指定文件时，无需指定模块）。其中，filename 参数用来指定文件名，lineno 用来指定行号，module_globals 参数用来指定要读取的具体模块名。注意，当指定文件以相对路径的方式传给 filename 参数时，该函数以按照 sys.path 规定的路径查找该文件。
linecache.clearcache()	如果程序某处，不再需要之前使用 getline() 函数读取的数据，则可以使用该函数清空缓存。
linecache.checkcache(filename=None)	检查缓存的有效性，即如果使用 getline() 函数读取的数据，其实在本地已经被修改，而我们需要的是新的数据，此时就可以使用该函数检查缓存的是否为新的数据。注意，如果省略文件名，该函数将检车所有缓存数据的有效性。
'''