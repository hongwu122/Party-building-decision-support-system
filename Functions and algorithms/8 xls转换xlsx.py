import os
import win32com.client as win32

print('当前工作路径', os.path.abspath('.'))  # 打印当前目录
path = r'C:\Users\鸿武\Desktop\2021_09_08  大创'
path2 = r'C:\Users\鸿武\Desktop\2021_09_08  大创\（第二课堂补录）2021年大学生创新创业计划名单模板.xlsx'
xlsx_files = [x for x in os.listdir(path) if os.path.splitext(x)[1] == '.doc']
print(xlsx_files)
print(os.listdir(path)) # 罗列文件夹下面的所有文件
print(0,os.path.splitext(path))
print(1,os.path.splitext(path)[0])
# print(2,os.path.splitext(path)[1])

# xlsx_files = [x for x in os.listdir(path2) if os.path.splitext(x)[1] == '.xlsx']
# print(xlsx_files)
# print(os.listdir(path2))
print(0,os.path.splitext(path2)) # 把文件名字分成两部分，名字和后缀
print(1,os.path.splitext(path2)[0])
print(2,os.path.splitext(path2)[1])
'''
当前工作路径 C:\\Users\鸿武\Desktop\Python3.7代码\0 实战-项目\2022_02_04 党建决策支持系统 V2
['（第二课堂补录）2021年大学生创新创业计划名单模板.xlsx']
['1 申报阶段', '2 论文初稿（高校党建决策支持系统的应用研究）', '20200923 经济管理与法学学院推免生指标分配及综合素质测评办法.pdf', 
'2021_12_22 大创分享', '3 论文次稿（面向党员考察工作的三维指派模型及其求解算法研究）', '4 论文投稿系列', '5 论文次稿2', 
'~$20大创申报书模板.doc', '~$新申报内容.docx', '~$省大学生创新创业训练计划项目申报表2.0.doc', '~WRL3149.tmp', 
'参考 精选文献', '参考 英文文献', '参考文献1', '参考文献2', '参考文献3', '立项文件', '经济管理与法学学院推免生指标分配及综合素质测评办法（2020）.pdf', 
'考研回忆录.docx', '通知文件', '（第二课堂补录）2021年大学生创新创业计划名单模板.xlsx', '（答辩各组信息）2021大创申报信息核对.xls']
0 ('C:\\Users\\鸿武\\Desktop\\2021_09_08  大创', '')
1 C:\\Users\鸿武\Desktop\2021_09_08  大创
0 ('C:\\Users\\鸿武\\Desktop\\2021_09_08  大创\\（第二课堂补录）2021年大学生创新创业计划名单模板', '.xlsx')
1 C:\\Users\鸿武\Desktop\2021_09_08  大创\（第二课堂补录）2021年大学生创新创业计划名单模板
2 
'''

def xls_to_xlsx(path,sole=True):# 默认单个
    # try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    if sole == True: # 单个文件
        path = eval(repr(path).replace(r'\\\\', r'/'))  # 把 \\\\ 替换成 /  不然会报错  一根是转义的\
        path = eval(repr(path).replace('/', r'\\'))  # 把 / 替换成 \\  不然会报错，
        wb = excel.Workbooks.Open(path)
        wb.SaveAs(path+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
        wb.Close()                               #FileFormat = 56 is for .xls extension
        excel.Application.Quit()
    if sole != True: # 整个文件夹
        xls_files = [x for x in os.listdir(path) if os.path.splitext(x)[1] == '.xls']
        for x in xls_files:
            sole_path = path + r'\\' + x
            print(sole_path)
            sole_path = eval(repr(sole_path).replace(r'\\\\',r'/')) #把 \\\\ 替换成 /  不然会报错  一根是转义的\
            print(sole_path)
            sole_path = eval(repr(sole_path).replace('/',r'\\')) #把 / 替换成 \\  不然会报错，
            # 初步认定，win32用win的单个\，其他\\和/不识别。且需要绝对路径
            print(sole_path)
            wb = excel.Workbooks.Open(sole_path)
            wb.SaveAs(sole_path + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
            wb.Close()  # FileFormat = 56 is for .xls extension
            excel.Application.Quit()
    print('成功修改')
    # except Exception as error:
    #         error = str(error)
    #         print('错误提示', error)
    #         # messagebox.showinfo('错误提示', '尝试把xls文件改成xlsx文件 失败！\n请自行另存为xlsx文件类型。\n错误信息：\n{}'.format(error))

xls_to_xlsx(path=r'C:\\Users\\鸿武\Desktop\Python3.7代码\0 实战-项目\2022_02_04 党建决策支持系统 V2',sole=False)



'''文件夹内的xls处理
xls_files = [x for x in os.listdir(path) if os.path.splitext(x)[1] == '.xls']
if xls_files != []:# 说明有xls文件
    xls_to_xlsx(path=path, sole=False) # 给路径，让其自己转换成xlsx的
    scr_output(scr_10, '\n\n检测到有{}个xls格式文件，已经自动在原路径转换生成可读取的xlsx文件类型：\n'.format(len(xls_files)))
'''

'''单个xls处理
if os.path.splitext(path)[1] == '.xls':# 说明是xls文件
    xls_to_xlsx(path=path, sole=True) # 给路径，让其自己转换成xlsx的
    scr_output(scr_10, '\n\n检测到本文件是xls格式文件，已经自动在原路径转换生成可读取的xlsx文件类型：\n')
    path = os.path.splitext(path)[1] + '.xlsx'

'''