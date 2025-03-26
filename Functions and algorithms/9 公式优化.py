from tkinter import messagebox
from tkinter.filedialog import *
from tkinter import ttk
from tkinter.ttk import *
from tkinter import scrolledtext

def gongshi():
    global list_gongshi
    def gongshi_save():
        global list_gongshi
        # list_gongshi = [gongshi1.get(),gongshi2.get(),gongshi3.get(),gongshi4.get(),gongshi5.get()]
        panduan_type = True
        panduan_int_list = [textvariable_year[0].get(),textvariable_day[0].get(),
                            textvariable_year[1].get(), textvariable_day[1].get(),
                            textvariable_year[2].get(), textvariable_day[2].get(),
                            textvariable_year[3].get(), textvariable_day[3].get(),
                            textvariable_year[4].get(), textvariable_day[4].get()]
        for i in range(len(panduan_int_list)):  # 正向遍历
            print(panduan_int_list[i])
            if panduan_int_list[i] == '':
                messagebox.showinfo('错误提示','{}值为空，不符合规范！请重新输入！'.format('有'))
                panduan_type = False
                break
            else:
                for j in panduan_int_list[i]:
                    if '0' <= j <= '9':  # 判断是不是数字
                        pass
                    else:
                        messagebox.showinfo('错误提示','{}值为非法字符，不符合规范！请重新输入！'.format('有'))
                        panduan_type = False
                        break
        if panduan_type == True:
            list_gongshi = ['int(first_value) - int(birth_value) - {}0000 -0{} <{} 0'.format(textvariable_year[0].get(),textvariable_day[0].get(),'=' if textvariable_baokuo[0].get()=='包括当天' else '' ),
                        'int(positive_value) - int(first_value) - {}0000 -0{} <{} 0'.format(textvariable_year[1].get(),textvariable_day[1].get(),'=' if textvariable_baokuo[1].get()=='包括当天' else '' ),
                        'int(object_value) - int(positive_value) - {}0000 -0{} <{} 0'.format(textvariable_year[2].get(),textvariable_day[2].get(),'=' if textvariable_baokuo[2].get()=='包括当天' else '' ),
                        'int(ready_value) - int(object_value) - {}0000 -0{} <{} 0'.format(textvariable_year[3].get(),textvariable_day[3].get(),'=' if textvariable_baokuo[3].get()=='包括当天' else '' ),
                        'int(become_value) - int(ready_value) - {}0000 -0{} <{} 0'.format(textvariable_year[4].get(),textvariable_day[4].get(),'=' if textvariable_baokuo[4].get()=='包括当天' else '' )]
            # scr_output(scr_5,'\n公式1：{}\n公式2：{}\n公式3：{}\n公式4：{}\n公式5：{}\n\n公式保存成功！\n\n\n'.format(gongshi1.get(),gongshi2.get(),gongshi3.get(),gongshi4.get(),gongshi5.get()))
            print(list_gongshi)
            gongshi.destroy()

    def gongshi_default():
        global list_gongshi
        # gongshi1.set('int(first_value) - int(birth_value) - 180000 < 0')
        # gongshi2.set('int(positive_value) - int(first_value) -15 <= 0')
        # gongshi3.set('int(object_value) - int(positive_value) - 10000 <= 0')
        # gongshi4.set('int(ready_value) - int(object_value) <= 0')
        # gongshi5.set('int(become_value) - int(ready_value) <= 0')
        textvariable_year[0].set('18')
        textvariable_year[1].set('0')
        textvariable_year[2].set('1')
        textvariable_year[3].set('0')
        textvariable_year[4].set('0')
        textvariable_day[0].set('0')
        textvariable_day[1].set('15')
        textvariable_day[2].set('0')
        textvariable_day[3].set('0')
        textvariable_day[4].set('0')
        textvariable_baokuo[0].set('不包括当天')
        textvariable_baokuo[1].set('包括当天')
        textvariable_baokuo[2].set('包括当天')
        textvariable_baokuo[3].set('包括当天')
        textvariable_baokuo[4].set('包括当天')

        list_gongshi = ['int(first_value) - int(birth_value) - {}0000 -{} <{} 0'.format(textvariable_year[0],textvariable_day[0],textvariable_baokuo[0]),
                        'int(positive_value) - int(first_value) - {}0000 -{} <{} 0'.format(textvariable_year[1],textvariable_day[1],textvariable_baokuo[1].get()),
                        'int(object_value) - int(positive_value) - {}0000 -{} <{} 0'.format(textvariable_year[2],textvariable_day[2],textvariable_baokuo[2].get()),
                        'int(ready_value) - int(object_value) - {}0000 -{} <{} 0'.format(textvariable_year[3],textvariable_day[3],textvariable_baokuo[3].get()),
                        'int(become_value) - int(ready_value) - {}0000 -{} <{} 0'.format(textvariable_year[4],textvariable_day[4],textvariable_baokuo[4].get())]
        # scr_output(scr_5,'\n公式已恢复默认！\n')

    gongshi = Tk()
    # gongshi = Toplevel(window)
    gongshi.geometry("500x190+700+310")
    try:
        gongshi.iconbitmap('mould\ico.ico')
    except:pass
    # 窗口置顶
    gongshi.attributes("-topmost", 1)  # 1==True 处于顶层
    # 禁止窗口的拉伸
    gongshi.resizable(0, 0)
    # 窗口的标题
    gongshi.title("公式编辑窗口")

    # 定义变量
    textvariable_year = [StringVar(),StringVar(),StringVar(),StringVar(),StringVar()]
    textvariable_day = [StringVar(),StringVar(),StringVar(),StringVar(),StringVar()]
    textvariable_baokuo = [StringVar(),StringVar(),StringVar(),StringVar(),StringVar()]
    text = ['公式1：出生日期-->申请入党  间隔', '公式2：申请入党-->积极分子  间隔', '公式3：积极分子-->发展对象  间隔',
            '公式4：发展对象-->预备党员  间隔', '公式5：预备党员-->转正时间  间隔']
    text2 = ['年 +','天']
    list_introduce = ['出生日期 和 首次递交入党申请书两个时间点之间关系','判断首次递交入党申请书 和 确认为入党积极分子两个时间点之间关系',
                      '确认为入党积极分子 到 列为发展对象两个时间点之间关系','列为发展对象 到 发展为预备党员两个时间点之间关系',
                      '发展为预备党员 到 预备党员转正两个时间点之间关系']
    for i in range(0, 5):
        label_gongshi = ttk.Label(gongshi, text=text[i])#标题
        label_gongshi.place(x=10, y=10 + 30 * i)
        # 多少年
        entry2_gongshi = ttk.Entry(gongshi, textvariable=textvariable_year[i])  # 输入框    # entry不能和grid连写，否则会报错
        entry2_gongshi.place(x=222, y=10 + 30 * i, width=50)
        label2_gongshi = ttk.Label(gongshi, text=text2[0])
        label2_gongshi.place(x=280, y=10 + 30 * i)
        # 多少天
        entry3_gongshi = ttk.Entry(gongshi, textvariable=textvariable_day[i])  # 输入框    # entry不能和grid连写，否则会报错
        entry3_gongshi.place(x=312, y=10 + 30 * i, width=50)
        label3_gongshi = ttk.Label(gongshi, text=text2[1])
        label3_gongshi.place(x=370, y=10 + 30 * i)
        # 包不包括当天
        Combobox_gongshi = ttk.Combobox(gongshi, width=8, textvariable=textvariable_baokuo[i])
        Combobox_gongshi['values'] = ['包括当天','不包括当天']
        Combobox_gongshi.place(x=400, y=10 + 30 * i, width=80)
        Combobox_gongshi.current(0)  # 设置初始显示值，值为元组['values']的下标
        Combobox_gongshi.config(state='readonly')  # 设为只读模式

        textvariable_year[0].set('18')
        textvariable_year[1].set('0')
        textvariable_year[2].set('1')
        textvariable_year[3].set('0')
        textvariable_year[4].set('0')
        textvariable_day[0].set('0')
        textvariable_day[1].set('15')
        textvariable_day[2].set('0')
        textvariable_day[3].set('0')
        textvariable_day[4].set('0')
        textvariable_baokuo[0].set('不包括当天')
        textvariable_baokuo[1].set('包括当天')
        textvariable_baokuo[2].set('包括当天')
        textvariable_baokuo[3].set('包括当天')
        textvariable_baokuo[4].set('包括当天')
        # createToolTip(entry_gongshi, '这里是一条判断{}的公式'.format(list_introduce[i]))  # Add Tooltip

    button_gongshi = ttk.Button(gongshi, text="保存参数", command=gongshi_save)
    button_gongshi.place(x=250, y=160)

    button_gongshi = ttk.Button(gongshi, text="恢复默认", command=gongshi_default)
    button_gongshi.place(x=120, y=160)

    # 显示窗口(消息循环)
    gongshi.mainloop()

list_gongshi = ['int(first_value) - int(birth_value) - 180000 < 0',
                'int(positive_value) - int(first_value) -15 <= 0',
                'int(object_value) - int(positive_value) - 10000 <= 0',
                'int(ready_value) - int(object_value) <= 0',
                'int(become_value) - int(ready_value) <= 0']
gongshi()