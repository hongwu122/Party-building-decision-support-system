
# 私人通话
from tkinter import *
from tkinter.filedialog import *
from tkinter import messagebox
import win32com.client
import time
import yagmail

def private_chat():
    def sending():
        a = text.get('1.0',END)
        # print('全文:\n',a)
        content = '全文:\n{}'.format(a)
        print(content)
        # print(type(content))        #  <class 'str'>
        try:
            # 1.创建一个播报对象
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
        except:
            print('语音播报模块不可用')
            messagebox.showinfo(title='飞鸽传书', message="很遗憾，语音播报模块不可用，但不影响发送功能")
        try:
            print('正在运行发送邮件模块')
            messagebox.showinfo(title='飞鸽传书', message="点击确定，开始运行发送邮件模块，期间软件可能会有点卡，请耐心等待！")

            # 连接服务器
            # 用户名、授权码、服务器地址
            yag_server = yagmail.SMTP(user='h13902688308@163.com', password='NCXGETOKOEQRQYVK', host='smtp.163.com')
            '''接着，通过 send() 函数，将邮件发送出去'''
            # 发送对象列表
            email_to = ['1228815090@qq.com', ]
            email_title = '党建使用者私信'
            email_content = content
            # 附件列表
            # email_attachments = ['./attachments/report.png', ]
            # 发送邮件
            yag_server.send(email_to, email_title, email_content)  # , email_attachments
            '''邮件发送完毕之后，关闭连接即可'''
            # 关闭连接
            yag_server.close()

            time.sleep(1)
            try:
                # 1.创建一个播报对象
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                speaker.Speak("发送成功")
            except:
                pass
            messagebox.showinfo(title='飞鸽传书', message="感谢来信，发送成功！")
        except:
            try:
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                speaker.Speak("发送失败")
            except:
                pass
            messagebox.showinfo(title='飞鸽传书', message="哎呀，小鸽子发送失败了，请检查下网络连接，或者其他问题功能不可用")


    def help():
        messagebox.showinfo('飞鸽小提示','点击在这个窗口文本框里，输入内容，最后点击发送。\n跟电脑自带的文本文档编辑器差不多操作，就是多了个发送按钮\n显示界面不是很好，条件简陋了点，和记事本差不多，应该是不限字数的。\n写完就点击发送.注意在此之前别关闭窗口哦，否则内容就没了哦。')

    def myopen():   # 打开本地
        global filename
        filename = askopenfilename(defaultextension='.txt')
        if filename == '':
            filename = None
        else:
            chat_Toplevel.title('飞鸽传书  ' + os.path.basename(filename))
            text.delete(1.0, END)
            f = open(filename, 'r')
            text.insert(1.0, f.read())
            f.close()

    def new():   # 新建文件
        global chat_Toplevel, filename, text
        chat_Toplevel.title('纸笺')
        filename = None
        text.delete(1.0, END)

    def save():   # 保存
        global filename
        try:
            f = open(filename, 'w')
            msg = text.get(1.0, 'end')
            f.write(msg)
            f.close()
        except:
            saveas()


    def saveas():   # 另存为
        f = asksaveasfilename(initialfile='纸笺.txt', defaultextension='.txt')
        global filename
        filename = f
        fh = open(f, 'w')
        msg = text.get(1.0, END)
        fh.write(msg)
        fh.close()
        chat_Toplevel.title('飞鸽传书  ' + os.path.basename(f))


    def cut(): # 剪切
        global text
        text.event_generate('<<Cut>>')
    def copy(): # 复制
        global text
        text.event_generate('<<Copy>>')
    def paste(): # 粘贴
        global text
        text.event_generate('<<Paste>>')
    def select_all(): # 全选
        global text
        # text.event_generate('sel', '1.0', 'end')
        # text.event_generate(" SEL('1.0', 'end')")
        # text.selection('1.0', 'end')
        text.event_generate("<<SelectAll>>")

    def find():   # 查找
        global chat_Toplevel
        t = Toplevel(chat_Toplevel)
        t.title('查找')
        # 设置窗口大小
        t.geometry('260x60+200+250')
        t.transient(chat_Toplevel)
        Label(t, text='查找:').grid(row=0, column=0, sticky='e')
        v = StringVar()
        e = Entry(t, width=20, textvariable=v)
        e.grid(row=0, column=1, padx=2, pady=2, sticky='we')
        e.focus_set()
        c = IntVar()
        Checkbutton(t, text='不区分大小写', variable=c).grid(row=1, column=1, sticky='e')
        Button(t, text='查找所有', command=lambda: search(v.get(), c.get(), text, t, e)).grid(row=0, column=2,
                                                                                             sticky='e' + 'w', padx=2,
                                                                                             pady=2)

        def close_search():
            text.tag_remove('match', '1.0', END)
            t.destroy()

        t.protocol('WM_DELETE_WINDOW', close_search)


    def search(needle, cssnstv, text, t, e):    # 搜索
        text.tag_remove('match', '1.0', END)
        count = 0
        if needle:
            pos = '1.0'
            while True:
                pos = text.search(needle, pos, nocase=cssnstv, stopindex=END)
                if not pos: break
                lastpos = pos + str(len(needle))
                text.tag_add('match', pos, lastpos)
                count += 1
                pos = lastpos
            text.tag_config('match', foreground='red', background='yellow')
            e.focus_set()
            t.title(str(count) + '个被匹配')

    def font():    # 调字体
        def change_a(self):    #  不加self 会报错 TypeError: change_a() takes 0 positional arguments but 1 was given
            global a,text,chat_Toplevel
            print('v.get',v.get())
            text['font'] = ('微软雅黑', '{}'.format(v.get()))
            text.pack(expand=YES, side=LEFT, fill=BOTH)

        global chat_Toplevel
        t = Toplevel(chat_Toplevel)
        t.title('字体')
        # 设置窗口大小
        t.geometry('260x60+200+280')
        v = StringVar()
        Scale(t,
              from_=10,  # 设置最大值
              to=30,  # 设置最小值
              resolution=1,  # 设置步距值
              orient=HORIZONTAL,  # 设置水平方向
              variable=v,  # 绑定变量
              command=change_a  # 设置回调函数
              ).pack()


    global chat_Toplevel,text,a
    chat_Toplevel = Toplevel()
    chat_Toplevel.geometry('500x500+480+270')
    chat_Toplevel.title('联系作者（请在文末留下您的联系方式）')
    chat_Toplevel.attributes("-topmost", 1)  # 1==True 处于顶层
    # 创建滚动条和文本框
    scroll = Scrollbar(chat_Toplevel)                         # 创建滚动条
    text = Text(chat_Toplevel, font=('微软雅黑', '15'))    # 创建文本框
    # 设置文本框初始内容
    text.insert('insert', 'From All Eternal Cute My Summer:\n')

    # 将滚动条和文本框分别填充
    scroll.pack(side=RIGHT, fill=Y)     # side指定Scrollbar为居右；fill指定填充满整个剩余区域     # side是滚动条放置的位置，上下左右。fill是将滚动条沿着y轴填充
    text.pack(expand=YES,side=LEFT,fill=BOTH)         # 将文本框填充进窗口的左侧         # expand=YES 支持扩张yes   fill=BOTH 填充XY
    # 将滚动条与文本框互相关联
    scroll.config(command=text.yview)        # 指定Scrollbar的command的回调函数是Listbar的yview     # 将文本框关联到滚动条上，滚动条滑动，文本框跟随滑动
    text.config(yscrollcommand=scroll.set)   # 将滚动条关联到文本框

    # 发送按钮
    # button = Button(chat_Toplevel,text='发送',command=sending)
    # button.place(x=220,y=460,width=60,height=30)

    # 菜单栏
    menubar=Menu(chat_Toplevel,tearoff = False)

    filemenu = Menu(menubar,tearoff = False)
    filemenu.add_command(label='新建纸条', command=new)       # , accelerator='Ctrl+N'
    filemenu.add_command(label='打开', command=myopen)    # , accelerator='Ctrl+O'
    filemenu.add_command(label='保存', command=save)      # , accelerator='Ctrl+S'
    filemenu.add_command(label='另存为', command=saveas)  # , accelerator='Ctrl+Shift+S'
    menubar.add_cascade(label=' 文件 ', menu=filemenu)

    editmenu = Menu(menubar,tearoff = False)
    editmenu.add_command(label='剪切', accelerator='Ctrl+X', command=cut)
    editmenu.add_command(label='复制', accelerator='Ctrl+C', command=copy)
    editmenu.add_command(label='粘贴', accelerator='Ctrl+V', command=paste)
    editmenu.add_command(label='全选', accelerator='Ctrl+A', command=select_all)
    menubar.add_cascade(label=' 编辑 ', menu=editmenu)

    menubar.add_command(label=' 字体 ', command=font)
    menubar.add_command(label=' 查找 ', command=find)
    menubar.add_command(label=' 帮助 ',command=help)
    menubar.add_command(label=' 点击发送 ',command=sending)

    def popup(event):
        menubar.post(event.x_chat_Toplevel,event.y_chat_Toplevel)
    #绑定鼠标右键
    text.bind('<Button-3>',popup)

    # 放在菜单栏
    chat_Toplevel.config(menu=menubar)

    chat_Toplevel.mainloop()


# private_chat()