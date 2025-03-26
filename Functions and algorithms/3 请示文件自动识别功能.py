# people_sheet = '''
# hyj juk  j uask a ka sk fksk fk saf ka ks  ksa k
# '''.split()
# import os
# import openpyxl
# path = 'E:/最近处理的文件/1部门事务/党建组织工作部 2/2021_12_14 报院党委请示情况表/接收预备党员、党员转正请示情况一览表'
#
# xlsx_files = [x for x in os.listdir(path) if os.path.splitext(x)[1] == '.xlsx']  # 罗列当前目录内所有xlsx文件
# # scr_output(scr_6, '\n\n需要提取名字的表格{}个'.format(len(xlsx_files)))
# # scr_output(scr_6, '\n\n需要提取名字的表格有：\n{}'.format(xlsx_files))
# print('需要提取', len(xlsx_files), '个表格')
# print('提取表格有：\n', xlsx_files)  # 本目录下的xlsx文件名字列表
# list_names = []
# for p in xlsx_files:
#     r,c = None,None
#     xlsx_file = path + '/' + p
#     workbook = openpyxl.load_workbook(filename=xlsx_file)
#     worksheet = workbook.worksheets[0]
#     # 获取名字信息
#     # print(worksheet[1:3])
#     for row in tuple(worksheet[1:3]):
#         for cell in row:
#             print(cell.value)
#             if cell.value == ('姓名' or '名字' or '名称'):
#                 r = cell.row
#                 c = cell.column
#                 break
#
#     if r != None and c != None:
#         print(r,c)
#         # print(worksheet[c])
#         list_name = list(cell.value for cell in [col for col in worksheet.columns][c-1])[r:]
#         list_names.append(list_name)



# a = ['a','a','ab','abc']
# b = ' '.join(i for i in a)
# print(b)

# a = 'asd ads,,a fvsa fs'
# b = a.split(',')
# print(b)