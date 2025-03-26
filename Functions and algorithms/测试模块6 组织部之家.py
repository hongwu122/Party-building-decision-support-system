

zu_zhi_bu = Toplevel(window)
zu_zhi_bu.geometry("1000x500+200+200")
try:
    zu_zhi_bu.iconbitmap('mould\ico.ico')
except:pass
# 窗口置顶
# zu_zhi_bu.attributes("-topmost", 1)  # 1==True 处于顶层
# 禁止窗口的拉伸
zu_zhi_bu.resizable(0, 0)
# 窗口的标题
zu_zhi_bu.title("欢迎来到组织部")

# 定义变量
zu_zhi_bu_var= StringVar()

one_zuzhibu = '''本部门旨在统一协调学生党建、党支部工作，严格把关党员发展工作，培育组织能力强、熟悉党员材料，发展程序的高素质党员。\n
组织工作部主要职能为协助书记及副书记协调党建委员会及学院各支部的日常工作。具体工作包括党员发展工作、入党积极分子培训、发展对象培训、预备党员培训和党员培训名单的收集审核，
以及每期入党积极分子、发展对象培训、预备党员转正培训和党员培训的证书制作、党费收缴工作、各部门活动审批工作以及其他日常性工作。
'''
two_zuzhibu = '''
第一届学生党建工作委员会组织工作部
组织工作部
部长：
刘佳
副部长：
齐双悦
王灿

第二届学生党建工作委员会组织工作部
部长：
修梓洋 
副部长：
工管181班 赵慧
经济183班 江南
干事：
电商192班 谭周豪
法学182班 易也博
国贸181班 彭思雅
国贸192班 左玉冰
商类198班 黄永健
会计182班 蒋茜
会计185班 陈婉溶
经济181班 杨明嘉
营销182班 贺娜

第三届学生党建工作委员会组织工作部
部长：
工管181班 赵慧 
部长助理：
工企192班 黄永健
温琳艳
张祯婧
干事：
会计ACCA201班 王翊成
商类207班 何泓霖
电商192班 田梦凡
会计ACCA201班 严雨婷
商类213班 邓琴
会计192班 周国燕
法学191班 宋辰
法学201班 刘星雨
        
第四届学生党建工作委员会组织工作部
部长助理：
黄永健 
副部长：
会计204班 邓琴
经济201班 刘聪颖 
副部长助理：
会计ACCA201班 王翊成
法学201班 刘星雨
干事：
会计202班 赵碧
商类2105班 彭琼卉
会计ACCA201班 赵妍珂  
物流211班 张许兰 
国贸212班 石琳 
经济212班 陈晨
国贸211班 王静
经济202班 唐明敏
国贸212班 孔洁滢
商类2106班 李晶晶


'''

scr_zzb = scrolledtext.ScrolledText(zu_zhi_bu, wrap=WORD)
scr_zzb.place(x=10, y=10, width=500,height=250)
scr_output(scr=scr_zzb,information=one_zuzhibu)
scr_zzb.config(state=DISABLED)  # 关闭可写入模式

scr_zzb2 = scrolledtext.ScrolledText(zu_zhi_bu2, wrap=WORD)
scr_zzb2.place(x=250, y=10, width=500,height=250)
scr_output(scr=scr_zzb2,information=two_zuzhibu)
scr_zzb2.config(state=DISABLED)  # 关闭可写入模式

# 显示窗口(消息循环)
zu_zhi_bu.mainloop()