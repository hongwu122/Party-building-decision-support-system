# 检测证书编号（结业证号是否合法）
def zhengshu_bianhao_legal(data):
    data = str(data)
    if len(data) != 9: #九位数
        return False
    for i in data:  # 正向遍历
        if '0' <= i <= '9':  # 判断是不是数字
            pass
        else:
            return False
    if str(data[0]+data[1]) != '20':
        return False
    if data[4] == data[5] == data[6] == data[7] == data[8]:
        return False
    return True

# while True:
#     data = input()
#     print(zhengshu_bianhao_legal(data))


zhengshu_bianhao_value_list = ['1','2','3','4','3','1','2','3','5','10','2']
zhengshu_bianhao_value_repitition_list = []
from collections import Counter
dict = Counter(zhengshu_bianhao_value_list)
print(dict) # Counter({'1': 2, '2': 1, '3': 1, '4': 1})
a = sorted(dict.items(), key= lambda item:item[1], reverse=True)
print(a) # [('2', 3), ('3', 3), ('1', 2), ('4', 1), ('5', 1), ('10', 1)]
for i in range(len(a)):
    if a[i][1] >= 2:
        zhengshu_bianhao_value_repitition_list.append(a[i][0])