# 前言
# 我们在写论文的时候，为了避免论文重复，可以使用第三方的库进行查重。但是，有时候在写论文的时候，只是引用自己之前的资料，在查重前想对自己的论文两篇文章进行查重。网上找了一下，没找到合适的工具，于是就自己用Python写了一个。
#
# 基本思路
# 两篇论文查重的方法相对比较简单，即将文章拆分成小句，然后小句间进行两两对比。主要实现基本可以分为以下三步：
#
# 读取
# 使用Python的python-docx库，可以非常方便的读取Word的内容，具体可以参见官方文档，网上也有很多不错的文章请自行查询参考。
# 原文拆分
# 对比的基本思想是按小句进行比较，所以拆分以是标点，即，。？！等进行拆分。拆分完成以后，可以有很多的小段。本文中为了便于定位，先根据原始段落进行拆分，然后再将每段根据标点拆分成若干小句，即一个word文档 = [[段落1], [段落2], [段落3], ...,[段落n]]，而每个段落= [[小句1],[小句2],[小句3],...,[小句m],]。
# 循环对比输出
# 第三步就是根据段落，两两进行对比，遇到匹配输出结果。
# 在对比中，有几点要注意：
#
# 如果子句过短（长度<5）则忽略，因为这种情况都是名词或术语，允许重复。
# 两个子句比较时，并不是用等号，而用包括，即一个子句是否包含另一个子句。
# coding=utf-8

from docx import Document
import re, sys, datetime, os


def getText(wordname):
    d = Document(wordname)
    texts = []
    for para in d.paragraphs:
        texts.append(para.text)
    return texts


def is_Chinese(word):
    for ch in word:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False


def msplit(s, seperators=',|\.|\?|，|。|？|！'):
    # 这些符号来断句
    return re.split(seperators, s)


def readDocx(docfile):
    # print('*' * 80)
    # print('文件', docfile, '加载中……')
    t1 = datetime.datetime.now()
    paras = getText(docfile)
    segs = []
    for p in paras:
        temp = []
        for s in msplit(p):
            if len(s) > 2:
                temp.append(s.replace(' ', ""))
        if len(temp) > 0:
            segs.append(temp)
    t2 = datetime.datetime.now()
    # print('加载完成，用时: ', t2 - t1)
    showInfo(segs, docfile)
    return segs


def showInfo(doc, filename='filename'):
    chars = 0
    segs = 0
    for p in doc:
        for s in p:
            segs = segs + 1
            chars = chars + len(s)
    # print('段落数: {0:>8d} 个。'.format(len(doc)))
    # print('短句数: {0:>8d} 句。'.format(segs))
    # print('字符数: {0:>8d} 个。'.format(chars))


def compareParagraph(doc1, i, doc2, j, filepath, min_segment=5):
    # print("正在compareParagraph……")
    """
    功能为比较两个段落的相似度，返回结果为两个段落中相同字符的长度与较短段落长度的比值。
    :param p1: 行
    :param p2: 列
    :param min_segment = 5: 最小段的长度
    """
    p1 = doc1[i]
    p2 = doc2[j]
    len1 = sum([len(s) for s in p1])
    len2 = sum([len(s) for s in p2])
    if len1 < 10 or len2 < 10:
        return []

    list = []
    for s1 in p1:
        if len(s1) < min_segment:
            continue
        for s2 in p2:
            if len(s2) < min_segment:
                continue
            if s2 in s1:
                list.append(s2)
            elif s1 in s2:
                list.append(s1)

    # 取两个字符串的最短的一个进行比值计算
    count = sum([len(s) for s in list])
    ratio = float(count) / min(len1, len2)
    if count > 20 and ratio > 0.25:
        print(' 发现相同内容 '.center(80, '*'))
        print('与{}段落 存在相似'.format(filepath))
        print('文件1第{0:0>4d}段内容：{1}'.format(i + 1, p1))
        print('文件2第{0:0>4d}段内容：{1}'.format(j + 1, p2))
        print('相同内容：', list)
        print('相同字符比：{1:.2f}%\n相同字符数： {0}\n'.format(count, ratio * 100))
    return list




# doc1 = readDocx('C:\\Users\\鸿武\\Desktop\\Test\\x1.docx')
# doc2 = readDocx('C:\\Users\\鸿武\\Desktop\\Test\\x4.docx')


def all_files_path(rootDir):
    for root, dirs, files in os.walk(rootDir):     # 分别代表根目录、文件夹、文件
        for file in files:                         # 遍历文件
            file_path = os.path.join(root, file)   # 获取文件绝对路径
            if os.path.splitext(file_path)[1] == '.docx':
                filepaths.append(file_path)            # 将文件路径添加进列表
        for dir in dirs:                           # 遍历目录下的子目录
            dir_path = os.path.join(root, dir)     # 获取子目录路径
            all_files_path(dir_path)               # 递归调用

dirpath = 'C:\\Users\\鸿武\\Desktop\\Test\\心得体会'
filepaths = []                                     # 初始化文件列表用来
all_files_path(dirpath)
with open('dir.txt', 'a') as f:
    for filepath in filepaths:
        f.write(filepath + '\n')
print(filepaths)

print('开始比对...'.center(80, '*'))
t1 = datetime.datetime.now()
for fi in range(len(filepaths)):
    if fi % 1000 == 0:
        print('处理进行中，已处理文章数量 {0:>4d} (总数 {1:0>4d} ） '.format(fi, len(filepaths)))
    doc1 = readDocx('C:\\Users\\鸿武\\Desktop\\Test\\李晶晶心得体会.docx')
    doc2 = readDocx(filepaths[fi])
    for i in range(len(doc1)):
        for j in range(len(doc2)):
            compareParagraph(doc1, i, doc2, j, filepaths[fi])

t2 = datetime.datetime.now()
print('\n比对完成，总用时: ', t2 - t1)




