from pdf2docx import Converter
# from configparser import ConfigParser
import os
import win32com.client as win32

# 将pdf转换成word（.docx）格式
def pdf2word(path,sole=True):
    path = eval(repr(path).replace(r'\\\\', r'/'))  # 把 \\ 替换成 /  不然会报错  一根是转义的\
    path = eval(repr(path).replace('/', r'\\'))  # 把 / 替换成 \  不然会报错，
    if sole==True:
        # os.path.splitext(path)把文件名字分成两部分，名字和后缀
        if os.path.splitext(path)[1] != '.pdf':
            print('给定文件不是pdf文件')
            return
        pdf_file = path
        word_file = os.path.splitext(path)[0] + '.docx'
        cv = Converter(pdf_file)# 也支持相对路径
        try:
            cv.convert(word_file)
        except Exception as error:
            error = str(error)
            print('错误提示', error)
            if error == 'No parsed pages. Please parse page first.':
                print('错误是：pdf2docx.converter.ConversionException: No parsed pages. Please parse page first.')
                print('用word转pdf的PDF文件，再回来pdf转word会报这个错误！')
        cv.close()
    if sole==False:
        # 判断有没有文件
        if os.listdir(path) == []:
            print("文件夹为空，请检查！")
            return
        for file in os.listdir(path): #  # os.listdir(path) 罗列文件夹下面的所有文件
            if os.path.splitext(file)[1] != '.pdf':
                continue
            file_name = os.path.splitext(file)[0]
            pdf_file = path + '\\' + file
            word_file = path + '\\' + file_name + '.docx'
            cv = Converter(pdf_file)
            try:
                cv.convert(word_file)
            except Exception as error:
                error = str(error)
                print('错误提示', error)
                if error == 'No parsed pages. Please parse page first.':
                    print('错误是：pdf2docx.converter.ConversionException: No parsed pages. Please parse page first.')
                    print('用word转pdf的PDF文件，再回来pdf转word会报这个错误！')
            cv.close()

# 将doc和docx文件转换成pdf格式
def word2pdf(path,sole=True):
    path = eval(repr(path).replace(r'\\\\', r'/'))  # 把 \\ 替换成 /  不然会报错  一根是转义的\
    path = eval(repr(path).replace('/', r'\\'))  # 把 / 替换成 \  不然会报错，
    # 注意：word文件路径和生成pdf文件路径一定要使用绝对路径
    # word = win32.Dispatch('Word.Application')
    word = win32.gencache.EnsureDispatch('Word.Application')
    if sole==True:
        if (os.path.splitext(path)[1] != '.doc') and (os.path.splitext(path)[1] != '.docx'):
            print('给定文件不是.doc或.docx文件')
            return
        doc = word.Documents.Open(path)
        pdf_file = os.path.splitext(path)[0] + ".pdf"  # 生成pdf文件路径名称
        doc.SaveAs(pdf_file, FileFormat=17)
        print("文件{}完成.docx到.pdf的转换！".format(path))
        doc.Close()
        word.Quit()
    if sole==False:
        for dirpath, dirnames, filenames in os.walk(path): # path是文件夹地址
            # dirpath是文件夹路径，dirnames为空，filenames是文件夹下面所有的文件名字
            # 判断有没有文件
            if filenames==[]:
                print("文件夹为空，请检查！")
                return
            # 判断是不是含有.doc或者.docx文件
            elif ".doc" or ".docx" in filenames:
                for file in filenames:
                    if file.lower().endswith(".docx"):
                        pdf_file = file.replace(".docx", ".pdf")
                        word_file =(dirpath + '/'+ file)
                        pdf_file =(dirpath + '/' + pdf_file)
                        doc = word.Documents.Open(word_file)
                        doc.SaveAs(pdf_file, FileFormat = 17)
                        print("文件{}完成.docx到.pdf的转换！".format(word_file))
                        doc.Close()
                    elif file.lower().endswith(".doc"):
                        pdf_file = file.replace(".doc", ".pdf")
                        word_file =(dirpath +'\\' + file)
                        pdf_file =(dirpath +'\\' + pdf_file)
                        doc = word.Documents.Open(word_file)
                        doc.SaveAs(pdf_file, FileFormat = 17)
                        print("文件{}完成.doc到.pdf的转换！".format(word_file))
                        doc.Close()
        word.Quit()







# pdf2docx.converter.ConversionException: No parsed pages. Please parse page first.

# pdf2word(r'C:/Users/鸿武/Desktop/测试文件夹',sole=False)
#
# pdf2word(r'C:\\Users\\鸿武\Desktop\\测试文件夹\\p面向智慧党建党员考察的三维指派模型及其求解算法研究（投稿版0205）.pdf',sole=True)
#
# word2pdf(r'C:\\Users\鸿武\Desktop\测试文件夹',sole=False)
#
# word2pdf(r'C:\\Users\鸿武\Desktop\测试文件夹\pcikit-opt-v0.6.2-合并.docx',sole=True)
#
# word2pdf('C:/Users/鸿武\Desktop/测试文件夹/w 0220203 Xie Editing 党员指派问题研究.doc',sole=True)

# pdf2word(r'C:\\Users\\鸿武\Desktop\\测试文件夹\\q经管法党员、预备党员、发展对象学员册（2021下半年） 10.19（审核标注） 2022_02_09.pdf',sole=True)

# word2pdf(r'C:\\Users\鸿武\Desktop\测试文件夹\q经管法党员、预备党员、发展对象学员册（2021下半年） 10.19（审核标注） 2022_02_09.docx',sole=True)





