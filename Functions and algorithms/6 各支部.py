# 支部管理 生成各支部的请示和批复
# 支部请示文件cookie的模板的识别（后续开发做准备）
def zhibu_qingshi_model_cookie(cookie,party_name,year,month,day,zd_year,zd_month,zd_day,first_people,people_num,people_sheet,
                               yeardu,year_up,dy_sum,dy_true,dy_wait,dy_in,dy_true_in):
    if cookie == '100':  # 发展对象的请示
        a = "关于建议将{}等{}人列为中共党员发展对象的请示".format(first_people, people_num)
        b = "尊敬的学院党委："
        c = "{}等{}人，自    年  月至    年  月期间先后递交入党申请书，经各支部支委会讨论研究，" \
            "于    年  月至    年  月期间先后确定为入党积极分子并参加党校培训，学习优秀，获得结业证书。".format(first_people,people_num)
        d = "该{}人自递交入党申请书以来，以实际行动向党组织靠拢，以党员标准严格要求自己。政治上，认真学习党的理论，" \
            "坚决拥护党的领导、方针和政策，与党中央保持一致，入党动机端正；思想上，树立正确的人生观、价值观和世界观，" \
            "坚定共产主义信念，热爱祖国，热爱人民，严格要求自己，做到身未入党思想先入党；工作上，充分发挥了不怕苦、" \
            "不怕累、乐于奉献的精神，有强烈的责任心和集体荣誉感，起到了先锋模范作用；学习上，态度端正，成绩优良，" \
            "既牢固掌握本专业的基础知识和技能，又广泛学习其他学科的知识；在生活上，勤俭节约，诚实守信；作风上，" \
            "求真务实，言行一致，廉洁自律；纪律上，自觉遵守校纪校规，无任何违法违纪违规情况。经较长时间的培养和教育，" \
            "该{}人进步明显，对党的认识深刻，各方面表现突出，党员和群众对{}人评价良好，已基本符合入党条件。" \
            "".format(people_num,people_num,people_num)
        e = "鉴于以上表现，经支委会讨论研究，确认{}等{}人为{}年{}党员发展对象人选，建议院党委将其列为中共党员发展对象。" \
            "名单如下（排名以班级为序）：".format(first_people,people_num,yeardu,year_up)
        f = "请批示。"
        g = "党支部书记签字：______________"
        h = "中共南华大学经济管理与法学学院"
        i = "{}".format(party_name)
        j = "{}年{}月{}日".format(year, month, day)
    if cookie == '010':  # 预备党员的请示
        a = "关于建议接收{}等{}名同志为中共预备党员的请示".format(first_people, people_num)
        b = "尊敬的学院党委："
        c = "{}等{}人向党组织提出了入党申请，该{}人主动接受党的入党积极分子、发展对象阶段的培养和考察，" \
            "在各方面都能以共产党员的标准严格要求自己，群众反映良好。党支部对该{}人进行了严格考察，认真审核材料。" \
            "于{}年{}月{}日召开支部大会讨论，认为{}等{}名同志符合党员的条件，提请学院党委接收其为预备党员，" \
            "名单如下（排名以班级为序）：".format(first_people, people_num,people_num,people_num,
                                    zd_year,zd_month,zd_day,first_people, people_num)
        d = None
        e = None
        f = "请批示。"
        g = "党支部书记签字：______________"
        h = "中共南华大学经济管理与法学学院"
        i = "{}".format(party_name)
        j = "{}年{}月{}日".format(year, month, day)
    if cookie == '001':  # 预备党员转正的请示
        a = "支部同意{}等{}名同志转为正式党员呈报院党委审批的请示".format(first_people, people_num)
        b = "尊敬的学院党委："
        c = "经济管理与法学学院{}于{}年{}月{}日召开支部大会，讨论{}等共{}名同志的入党转正申请。" \
            "大会认为{}等共{}名同志在预备期间，能以党员标准严格要求自己，重视政治理论学习，学习刻苦，成绩优异，" \
            "积极参加学校的各项活动，有较强的组织纪律观念，发挥了一个共产党员应有的作用。 " \
            "".format(party_name,zd_year,zd_month,zd_day,first_people, people_num,first_people,people_num)
        d = "本支部共有党员{}名，其中正式党员{}名，预备党员{}名。到会党员{}名，其中正式党员{}名，" \
            "有表决权的到会人数超过应到会人数的半数，符合人数要求。经无记名表决，" \
            "{}名正式党员一致同意{}共{}名同志按期转为正式党员。名单如下（排名以班级为序）：" \
            "".format(dy_sum,dy_true,dy_wait,dy_in,dy_true_in,dy_true_in,first_people, people_num)
        e = None
        f = None
        g = "党支部书记签字：______________"
        h = "中共南华大学经济管理与法学学院"
        i = "{}".format(party_name)
        j = "{}年{}月{}日".format(year, month, day)
    return a,b,c,d,e,f,g,h,i,j
# 支部请示文件的写入
def write_zhibu_qingshi(cookie,party_name,year,month,day,zd_year,zd_month,zd_day,first_people,people_num,people_sheet,
                               yeardu,year_up,dy_sum,dy_true,dy_wait,dy_in,dy_true_in):
    try:
        if type(people_sheet) is str: people_sheet = people_sheet.split()
        if cookie == '000':
            messagebox.showinfo('错误提示', '未选中支部请示的类型，请检查！')
            return
        if people_num != len(people_sheet):
            scr_output(scr_11, '\n生成支部请示文件 失败！\n错误信息：支部同志人数{}与支部人名数量{}不匹配，请检查！\n'.format(people_num,len(people_sheet)))
            messagebox.showinfo('错误提示', '生成支部请示文件 失败！\n错误信息：同志人数{}与人名数量{}不匹配，请检查！'.format(people_num,len(people_sheet)))
            return
        a,b,c,d,e,f,g,h,i,j = zhibu_qingshi_model_cookie(cookie,party_name,year,month,day,zd_year,zd_month,zd_day,first_people,people_num,people_sheet,
                               yeardu,year_up,dy_sum,dy_true,dy_wait,dy_in,dy_true_in)
        doc = Document()
        # 判断人数，来设置表格
        if 0 <= people_num<= 64: # 小四字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [2.43,1.9]
            row_height = 1
        if 64 < people_num<= 104: # 小四字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE # 1.5倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [2.15,1.8]
            row_height = 0.9
        if 104 < people_num <= 120:  # 小四字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [2.15, 1.8]
            row_height = 0.8
        if 120 < people_num<= 168: # 小四字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST # 最小倍倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [1.98,1.8]
            row_height = 0.55
        if 168 < people_num<= 184: # 五号字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST # 最小倍倍行距
            doc.styles['Normal'].font.size = Pt(10.5)
            col_width = [1.98,1.8]
            row_height = 0.55
        if 184 < people_num:
            doc.styles['Normal'].font.size = Pt(10)
            col_width = [1.98, 1.8]
            row_height = 0.55
            print('支部人数太多（大于184），请自行调整word中存在的格式问题。')
            scr_output(scr_11, '支部人数太多（大于184），请自行调整word中存在的格式问题。')
        # 标题样式
        doc.styles['Header'].font.name = 'Times New Roman'  # 设置英文字体
        doc.styles['Header']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 设置中文字体
        doc.styles['Header'].font.bold = True  # 加粗
        doc.styles['Header'].font.size = Pt(14)
        doc.styles['Header'].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER  # 居中对齐
        doc.styles['Header'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE # 1.5倍行距
        doc.styles['Header'].paragraph_format.space_before = Pt(0)  # 段前
        doc.styles['Header'].paragraph_format.space_after = Pt(0)  # 段后
        # 普通正文央视
        doc.styles['Footer'].font.name = 'Times New Roman'  # 设置英文字体
        doc.styles['Footer']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 设置中文字体
        doc.styles['Footer'].font.size = Pt(12)
        doc.styles['Footer'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY # 两端对齐
        doc.styles['Footer'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE # 1.5倍行距
        doc.styles['Footer'].paragraph_format.space_before = Pt(0)  # 段前
        doc.styles['Footer'].paragraph_format.space_after = Pt(0)  # 段后
        doc.styles['Footer'].paragraph_format.first_line_indent = doc.styles['Footer'].font.size * 2  #首行缩进2字符 1 英寸=2.54 厘米
        # 表格样式
        doc.styles['Normal'].font.name = 'Times New Roman'  # 设置英文字体
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 设置中文字体
        # doc.styles['Normal'].font.size = Pt(12)
        doc.styles['Normal'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE # 分散对齐
        # doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST # 最小倍倍行距
        doc.styles['Normal'].paragraph_format.space_before = Pt(0)  # 段前
        doc.styles['Normal'].paragraph_format.space_after = Pt(0)  # 段后
        doc.styles['Normal'].paragraph_format.first_line_indent = Inches(0)  #首行缩进2字符 1 英寸=2.54 厘米

        # 标题 一段
        doc.add_paragraph(a,style='Header')
        # 称呼 一段 （首不设两字符）
        doc.add_paragraph(b,style='Footer').paragraph_format.first_line_indent=Inches(0) # 1 英寸=2.54 厘米
        # 正文
        doc.add_paragraph(c,style='Footer')
        if d != None:
            doc.add_paragraph(d,style='Footer')
        if e != None:
            doc.add_paragraph(e,style='Footer')

        table = doc.add_table(people_num//8 if people_num%8 == 0 else people_num//8+1, 8)
        table.autofit = True   # if is True 按窗口大小自动调整
        count = 0

        for row in range(len(table.rows)):
            table.rows[row].height = Cm(row_height)  # 调整行高
            for col in range(len(table.columns)):
                # print(行, 列)  # 可以查看表格输出结果
                table.cell(row, col).text = people_sheet[count]    # 写入人名
                # table.cell(行, 列).width = doc.styles['Normal'].font.size * len(people_sheet[count])
                # table.cell(行, 列).height = doc.styles['Normal'].font.size * 5
                table.cell(row, col).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER  # 上下居中（中部居中对齐）
                # table.cell(行, 列).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 水平居中（中部居中对齐）
                count += 1
                if count == people_num:
                    break
            if count == people_num:
                break
        # 修正列宽
        for col in range(len(table.columns)):
            maxlist = []
            for r in range(len(table.rows)):
                try:
                    maxlist.append(len(people_sheet[8*r + col]))
                    # print(people_sheet[8*r + col])
                except:pass
            if maxlist != []: maxnum = max(maxlist) # 每一列的最大值
            else: maxnum = 3  # 每一列的最大值
            table.cell(len(table.rows)-1, col).width = Cm( col_width[0] if maxnum==4 else col_width[1] ) # 调整列宽 2字:1.3 3字:1.8 4字:2.1
            # 要在最后一行设置列宽度，因为设置前面的，后面一行出现空格，前面设置的宽度就不生效了

        table.alignment = WD_TABLE_ALIGNMENT.CENTER  # 设置整个表格为居中对齐
        # table.autofit = True
        # 结束语
        if f != None:
            doc.add_paragraph(f,style='Footer')
        # 落款二段和时间一段
        doc.add_paragraph(g,style='Footer').alignment = WD_ALIGN_PARAGRAPH.RIGHT  # 右对齐
        doc.add_paragraph(h,style='Footer').alignment = WD_ALIGN_PARAGRAPH.RIGHT  # 右对齐
        doc.add_paragraph(i,style='Footer').alignment = WD_ALIGN_PARAGRAPH.RIGHT
        doc.add_paragraph(j,style='Footer').alignment = WD_ALIGN_PARAGRAPH.RIGHT

        index = None
        for i in range(len(zhibu_allname)):
            if zhibu_allname[i] == party_name:
                index = i
        doc.save( party_name if index == None else zhibu_list[i] + ' ' + a + '.docx')
        messagebox.showinfo('小提示', '生成支部请示文件 成功！请注意检查word文件格式！')
        scr_output(scr_11, '\n\n生成支部请示文件 成功！请注意检查word文件格式！\n')

    except Exception as error:
        error = str(error)
        print('错误提示', error)
        scr_output(scr_11, '\n生成支部请示文件 失败！\n\n\n本次错误信息：{}\n\n'.format(error))
        scr_output(scr_11, '\n--------文件没有成功保存--------\n\n\n\n\n\n\n')
        messagebox.showinfo('错误提示', '生支部成请示文件 失败！\n错误信息：\n{}'.format(error))
# 支部请示管理 自动检测姓名列 更新多个变量值
def auto_zhibu_qingshi_read():
    if scr_sheet1_11.get(1.0, 'end').split() != []:
        messagebox.showinfo('小提示', '已经识别到文本中已有人名，请勿重复生成，请检查！'+'\n'
                            +'如需重新生成，请记得Ctrl+A清除不需要的人名，以防出错！'+'\n'
                            +'注意：本提示只是温馨提示，是不会停止继续执行自动检测的')
    # print(pathin_6.get())
    if pathin1_11.get() != '':
        path = pathin1_11.get()
        scr_output(scr_11, '\n\n用于提取名单的表格为：\n{}'.format(path))
        r, c = None, None
        workbook = openpyxl.load_workbook(filename=path)
        worksheet = workbook.worksheets[0]
        # 获取名字信息
        for row in tuple(worksheet[1:3]):
            for cell in row:
                # print(cell.value)
                if cell.value == ('姓名' or '名字' or '名称'):
                    r = cell.row
                    c = cell.column
                    break
        if r != None and c != None:
            # print(r, c)
            # print(worksheet[c])
            list_name = list(cell.value for cell in [col for col in worksheet.columns][c - 1])[r:]
            scr_output(scr_11, '\n\n提取出来的名单：\n{}'.format(list_name))
            # print('\n\n提取出来的名单：\n{}'.format(list_name))
        else:
            # print('找不到名字，请手动输入！')
            scr_output(scr_11, '\n找不到名字，请手动输入！\n')
            return
        # print(list_names)
        scr_sheet1_11.insert('insert', ' '.join(i for i in list_name)) # 插入名字
        scr_sheet1_11.update()  # 插入后及时的更新
        scr_sheet1_11.see(END)  # 使得聊天记录text默认显示底端
    else:
        print('路径为空！')
        scr_output(scr_11, '\n路径为空！\n')
    people_sheet = scr_sheet1_11.get(1.0, 'end').split()
    if people_sheet != []:
        number5_11.set(people_sheet[0])
        number6_11.set(len(people_sheet))
    else: print('名单为空，请检查！')
# 支部请示模板修改，未完善
def zhibu_qingshi_model_alter():
    messagebox.showinfo('小提示', '本版本只支持查看支部请示模板，暂不支持修改')
    cookie = str(number_11_1.get()) + str(number_11_2.get()) + str(number_11_3.get())
    if cookie == '000':
        messagebox.showinfo('错误提示', '未选中支部请示的类型，请检查！')
        return
    if cookie == '100':  # 发展对象的请示
        name = '发展对象'
        a = "关于建议将{}等{}人列为中共党员发展对象的请示"
        b = "尊敬的学院党委："
        c = "{}等{}人，自{}年{}月至{}年{}月期间先后递交入党申请书，经各支部支委会讨论研究，" \
            "于{}年{}月至{}年{}月期间先后确定为入党积极分子并参加党校培训，学习优秀，获得结业证书。"
        d = "该{}人自递交入党申请书以来，以实际行动向党组织靠拢，以党员标准严格要求自己。政治上，认真学习党的理论，" \
            "坚决拥护党的领导、方针和政策，与党中央保持一致，入党动机端正；思想上，树立正确的人生观、价值观和世界观，" \
            "坚定共产主义信念，热爱祖国，热爱人民，严格要求自己，做到身未入党思想先入党；工作上，充分发挥了不怕苦、" \
            "不怕累、乐于奉献的精神，有强烈的责任心和集体荣誉感，起到了先锋模范作用；学习上，态度端正，成绩优良，" \
            "既牢固掌握本专业的基础知识和技能，又广泛学习其他学科的知识；在生活上，勤俭节约，诚实守信；作风上，" \
            "求真务实，言行一致，廉洁自律；纪律上，自觉遵守校纪校规，无任何违法违纪违规情况。经较长时间的培养和教育，" \
            "该{}人进步明显，对党的认识深刻，各方面表现突出，党员和群众对{}人评价良好，已基本符合入党条件。"
        e = "鉴于以上表现，经支委会讨论研究，确认{}等{}人为{}年{}党员发展对象人选，建议院党委将其列为中共党员发展对象。" \
            "名单如下（排名以班级为序）："
        f = "请批示。"
        g = "党支部书记签字：______________"
        h = "中共南华大学经济管理与法学学院"
        i = "{支部全称}"
        j = "{}年{}月{}日"
    if cookie == '010':  # 预备党员的请示
        name = '预备党员'
        a = "关于建议接收{}等{}名同志为中共预备党员的请示"
        b = "尊敬的学院党委："
        c = "{}等{}人向党组织提出了入党申请，该{}人主动接受党的入党积极分子、发展对象阶段的培养和考察，" \
            "在各方面都能以共产党员的标准严格要求自己，群众反映良好。党支部对该{}人进行了严格考察，认真审核材料。" \
            "于{}年{}月{}日召开支部大会讨论，认为{}等{}名同志符合党员的条件，提请学院党委接收其为预备党员，" \
            "名单如下（排名以班级为序）："
        d = ''
        e = ''
        f = "请批示。"
        g = "党支部书记签字：______________"
        h = "中共南华大学经济管理与法学学院"
        i = "{支部全称}"
        j = "{}年{}月{}日"
    if cookie == '001':  # 预备党员转正的请示
        name = '预备党员转正'
        a = "支部同意{}等{}同志转为正式党员呈报院党委审批的请示"
        b = "尊敬的学院党委："
        c = "经济管理与法学学院{}于{}年{}月{}日召开支部大会，讨论{}等共{}名同志的入党转正申请。" \
            "大会认为{}等共{}名同志在预备期间，能以党员标准严格要求自己，重视政治理论学习，学习刻苦，成绩优异，" \
            "积极参加学校的各项活动，有较强的组织纪律观念，发挥了一个共产党员应有的作用。 "
        d = "本支部共有党员{}名，其中正式党员{}名，预备党员{}名。到会党员{}名，其中正式党员{}名，" \
            "有表决权的到会人数超过应到会人数的半数，符合人数要求。经无记名表决，" \
            "{}名正式党员一致同意{}共{}名同志按期转为正式党员。"
        e = ''
        f = ''
        g = "党支部书记签字：______________"
        h = "中共南华大学经济管理与法学学院"
        i = "{支部全称}"
        j = "{}年{}月{}日"

    list_zhibu_qingshi_model = [a,b,c,d,e,f,g,h,i,j]
    def zhibu_qingshi_model_save():
        scr_output(scr_11,'\n{}\n支部请示模板保存失败！，本版本模板不支持修改！\n')
        zhibu_qingshi_model.destroy()

    def zhibu_qingshi_model_default():
        scr_output(scr_11,'\n模板已经是默认！\n')

    zhibu_qingshi_model = Toplevel(window)
    zhibu_qingshi_model.geometry("500x290+700+270")
    zhibu_qingshi_model.iconbitmap('mould\ico.ico')
    # 窗口置顶
    zhibu_qingshi_model.attributes("-topmost", 1)  # 1==True 处于顶层
    # 禁止窗口的拉伸
    zhibu_qingshi_model.resizable(0, 0)
    # 窗口的标题
    zhibu_qingshi_model.title("内置-{}-支部请示模板-修改窗口".format(name))

    # 定义变量
    qingshi_model_var= StringVar()
    scr_zhibu_qingshi_model = scrolledtext.ScrolledText(zhibu_qingshi_model, wrap=WORD)
    scr_zhibu_qingshi_model.place(x=10, y=10, width=480,height=245)
    scr_zhibu_qingshi_model.config(state=DISABLED)  # 关闭可写入模式
    for i in list_zhibu_qingshi_model:
        scr_output(scr_zhibu_qingshi_model, str(i) + '\n')

    button_zhibu_qingshi_model = ttk.Button(zhibu_qingshi_model, text="保存参数", command=zhibu_qingshi_model_save)
    button_zhibu_qingshi_model.place(x=250, y=260)

    button_zhibu_qingshi_model = ttk.Button(zhibu_qingshi_model, text="恢复默认", command=zhibu_qingshi_model_default)
    button_zhibu_qingshi_model.place(x=120, y=260)

    # 显示窗口(消息循环)
    zhibu_qingshi_model.mainloop()


















# 支部批复文件cookie的模板的识别（后续开发做准备）
def zhibu_pifu_model_cookie(cookie, party_name,qs_year,qs_month,qs_day,year, month, day,first_people, people_num,people_sheet):
    if cookie == '100':  # 发展对象的批复
        a = "关于同意将{}等{}人列为中共党员发展对象的批复".format(first_people,people_num)
        b = "中共南华大学经济管理与法学学院"
        c = "{}：".format(party_name)
        d = "收到了贵支部{}年{}月{}日“关于建议将{}等{}人列为中共党员发展对象的请示”，且公示无异议。" \
            "认为你们按照党员标准对入党积极分子进行了有效的培养和教育。根据《中国共产党发展党员工作细则》的要求，" \
            "院党委于{}年{}月{}日召开党委会，认真讨论和审核材料，现将{}等" \
            "{}名同志列为中共党员发展对象，名单如下（排名以班级为序）：" \
            "".format(qs_year,qs_month,qs_day,first_people,people_num,year,month,day,first_people,people_num)
        e = "望你们继续加强对发展对象的培养和考察。"
        f = "特此批复。"
        g = "党委书记签名：_______________"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日".format(year, month, day)
    if cookie == '010':  # 预备党员的批复
        a = "关于同意接收{}等{}名同志为中共预备党员的批复".format(first_people,people_num)
        b = "中共南华大学经济管理与法学学院"
        c = "{}：".format(party_name)
        d = "收到了贵支部{}年{}月{}日“关于建议将{}等{}人列为中共预备党员的请示”，且公示无异议。" \
            "认为你们按照党员标准对发展对象进行了有效的培养和教育。根据《中国共产党发展党员工作细则》的要求，" \
            "院党委于{}年{}月{}日召开党委会，认真讨论和审核材料，现确定" \
            "{}等{}名同志为中共预备党员，名单如下（排名以班级为序）：" \
            "".format(qs_year,qs_month,qs_day,first_people,people_num,year,month,day,first_people,people_num)
        e = "望你们继续加强对预备党员的培养和考察。"
        f = "特此批复。"
        g = "党委书记签名：_______________"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日".format(year, month, day)
    if cookie == '001':  # 预备党员转正的批复
        a = "学院党委对支部同意{}等{}名同志转为正式党员决议的审批意见".format(first_people,people_num)
        b = "中共南华大学经济管理与法学学院"
        c = "{}：".format(party_name)
        d = "{}等{}名同志向党支部提出了转为正式党员的书面申请。学院党委在{}年{}月{}日召开党委会，" \
            "讨论通过你支部关于{}等{}名同志预备党员转为正式党员的决议。" \
            "{}等{}名同志从预备期满之日起成为中国共产党正式党员，党龄从即日算起。名单如下" \
            "（排名以班级为序）：".format(first_people,people_num,year,month,day,first_people,people_num,first_people,people_num)
        e = None
        f = None
        g = "党委书记签名：_______________"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日".format(year, month, day)
    return a, b, c, d, e, f, g, h, i
# 支部批复文件的写入
def write_zhibu_pifu(cookie, party_name,qs_year,qs_month,qs_day,year, month, day,first_people, people_num,people_sheet):
    try:
        if type(people_sheet) is str: people_sheet = people_sheet.split()
        if cookie == '000':
            messagebox.showinfo('错误提示', '未选中支部批复的类型，请检查！')
            return
        if people_num != len(people_sheet):
            scr_output(scr_11, '\n生成支部批复文件 失败！\n错误信息：同志人数{}与人名数量{}不匹配，请检查！\n'.format(people_num, len(people_sheet)))
            messagebox.showinfo('错误提示', '生成支部批复文件 失败！\n错误信息：同志人数{}与人名数量{}不匹配，请检查！'.format(people_num, len(people_sheet)))
            return
        a, b, c, d, e, f, g, h, i = pifu_model_cookie(cookie, yeardu, pici, year_up, qs_year,qs_month,qs_day,qingshi_name,
                                                year, month, day, party_name, party_num,first_people, people_num, people_sheet)  # 执行下面注释代码的函数
        doc = Document()
        # 判断人数，来设置表格
        if 0 <= people_num <= 64:  # 小四号字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [2.43, 1.9]
            row_height = 1
        if 64 < people_num <= 80:  # 小四字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [2.15, 1.8]
            row_height = 0.9
        if 80 < people_num <= 88:  # 小四字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [2.15, 1.8]
            row_height = 0.8
        if 88 < people_num <= 120:  # 小四字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST  # 最小倍倍行距
            doc.styles['Normal'].font.size = Pt(12)
            col_width = [1.98, 1.8]
            row_height = 0.55
        if 120 < people_num <= 136:  # 五号字体
            doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST  # 最小倍倍行距
            doc.styles['Normal'].font.size = Pt(10.5)
            col_width = [1.98, 1.8]
            row_height = 0.55
        if 136 < people_num:
            doc.styles['Normal'].font.size = Pt(10)
            col_width = [1.98, 1.8]
            row_height = 0.55
            print('支部人数太多（大于184），请自行调整word中存在的格式问题。')
            scr_output(scr_11, '支部人数太多（大于184），请自行调整word中存在的格式问题。')
        # 标题样式
        doc.styles['Header'].font.name = 'Times New Roman'  # 设置英文字体
        doc.styles['Header']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 设置中文字体
        doc.styles['Header'].font.bold = True  # 加粗
        doc.styles['Header'].font.size = Pt(14)
        doc.styles['Header'].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER  # 居中对齐
        doc.styles['Header'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5倍行距
        doc.styles['Header'].paragraph_format.space_before = Pt(0)  # 段前
        doc.styles['Header'].paragraph_format.space_after = Pt(0)  # 段后
        # 普通正文央视
        doc.styles['Footer'].font.name = 'Times New Roman'  # 设置英文字体
        doc.styles['Footer']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 设置中文字体
        doc.styles['Footer'].font.size = Pt(12)
        doc.styles['Footer'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # 两端对齐
        doc.styles['Footer'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 1.5倍行距
        doc.styles['Footer'].paragraph_format.space_before = Pt(0)  # 段前
        doc.styles['Footer'].paragraph_format.space_after = Pt(0)  # 段后
        doc.styles['Footer'].paragraph_format.first_line_indent = doc.styles[
                                                                      'Footer'].font.size * 2  # 首行缩进2字符 1 英寸=2.54 厘米
        # 表格样式
        doc.styles['Normal'].font.name = 'Times New Roman'  # 设置英文字体
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 设置中文字体
        # doc.styles['Normal'].font.size = Pt(12)
        doc.styles['Normal'].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE  # 分散对齐
        # doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST # 最小倍倍行距
        doc.styles['Normal'].paragraph_format.space_before = Pt(0)  # 段前
        doc.styles['Normal'].paragraph_format.space_after = Pt(0)  # 段后
        doc.styles['Normal'].paragraph_format.first_line_indent = Inches(0)  # 首行缩进2字符 1 英寸=2.54 厘米

        # 标题 一段
        doc.add_paragraph(a, style='Header')
        # 称呼两段（首不设两字符）
        doc.add_paragraph(b, style='Footer').paragraph_format.first_line_indent = Inches(0)  # 1 英寸=2.54 厘米
        doc.add_paragraph(c, style='Footer').paragraph_format.first_line_indent = Inches(0)  # 1 英寸=2.54 厘米
        # 正文
        doc.add_paragraph(d, style='Footer')

        table = doc.add_table(people_num // 8 if people_num % 8 == 0 else people_num // 8 + 1, 8)
        table.autofit = True  # if is True 按窗口大小自动调整
        count = 0

        for row in range(len(table.rows)):
            table.rows[row].height = Cm(row_height)  # 调整行高
            for col in range(len(table.columns)):
                # print(行, 列)  # 可以查看表格输出结果
                table.cell(row, col).text = people_sheet[count]  # 写入人名
                # table.cell(行, 列).width = doc.styles['Normal'].font.size * len(people_sheet[count])
                # table.cell(行, 列).height = doc.styles['Normal'].font.size * 5
                table.cell(row, col).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER  # 上下居中（中部居中对齐）
                # table.cell(行, 列).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 水平居中（中部居中对齐）
                count += 1
                if count == people_num:
                    break
            if count == people_num:
                break
        # 修正列宽
        for col in range(len(table.columns)):
            maxlist = []
            for r in range(len(table.rows)):
                try:
                    maxlist.append(len(people_sheet[8 * r + col]))
                    # print(people_sheet[8*r + col])
                except:
                    pass
            if maxlist != []:
                maxnum = max(maxlist)  # 每一列的最大值
            else:
                maxnum = 3  # 每一列的最大值
            table.cell(len(table.rows) - 1, col).width = Cm(
                col_width[0] if maxnum == 4 else col_width[1])  # 调整列宽 2字:1.3 3字:1.8 4字:2.1
            # 要在最后一行设置列宽度，因为设置前面的，后面一行出现空格，前面设置的宽度就不生效了

        table.alignment = WD_TABLE_ALIGNMENT.CENTER  # 设置整个表格为居中对齐
        # table.autofit = True
        # 结束语
        if e != None:
            doc.add_paragraph(e, style='Footer')
        if f != None:
            doc.add_paragraph(f, style='Footer')
        doc.add_paragraph("", style='Footer')
        # 落款和时间
        doc.add_paragraph(g, style='Footer').alignment = WD_ALIGN_PARAGRAPH.RIGHT  # 右对齐
        doc.add_paragraph(h, style='Footer').alignment = WD_ALIGN_PARAGRAPH.RIGHT  # 右对齐
        doc.add_paragraph(i, style='Footer').alignment = WD_ALIGN_PARAGRAPH.RIGHT

        index = None
        for i in range(len(zhibu_allname)):
            if zhibu_allname[i] == party_name:
                index = i
        doc.save( party_name if index == None else zhibu_list[i] + a + '.docx')

        messagebox.showinfo('小提示', '生成批复文件 成功！请注意检查word文件格式！')
        scr_output(scr_7, '\n\n生成批复文件 成功！请注意检查word文件格式！\n')

    except Exception as error:
        error = str(error)
        print('错误提示', error)
        scr_output(scr_11, '\n生成批复文件 失败！\n\n\n本次错误信息：{}\n\n'.format(error))
        scr_output(scr_11, '\n--------文件没有成功保存--------\n\n\n\n\n\n\n')
        messagebox.showinfo('错误提示', '生成批复文件 失败！\n错误信息：\n{}'.format(error))
# 支部批复管理 自动检测姓名列 更新多个变量值
def auto_zhibu_pifu_read():
    if scr_sheet2_11.get(1.0, 'end').split() != []:
        messagebox.showinfo('小提示', '已经识别到文本中已有人名，请勿重复生成，请检查！' + '\n'
                            + '如需重新生成，请记得Ctrl+A清除不需要的人名，以防出错！' + '\n'
                            + '注意：本提示只是温馨提示，是不会停止继续执行自动检测的')
    # print(pathin_7.get())
    # 如果路径不为空，写入名单
    if pathin2_11.get() != '':
        path = pathin2_11.get()
        scr_output(scr_11, '\n\n用于提取名单的表格为：\n{}'.format(path))
        r, c = None, None
        workbook = openpyxl.load_workbook(filename=xlsx_file)
        worksheet = workbook.worksheets[0]
        # 获取名字信息
        for row in tuple(worksheet[1:3]):
            for cell in row:
                # print(cell.value)
                if cell.value == ('姓名' or '名字' or '名称'):
                    r = cell.row
                    c = cell.column
                    break
        if r != None and c != None:
            # print(r, c)
            # print(worksheet[c])
            list_name = list(cell.value for cell in [col for col in worksheet.columns][c - 1])[r:]
            scr_output(scr_11, '\n\n提取出来的名单：\n{}'.format(list_name))
            # print('\n\n提取出来的名单：\n{}'.format(list_name))
        else:
            # print('找不到名字，请手动输入！')
            scr_output(scr_11, '\n找不到名字，请手动输入！\n')
            return
        scr_sheet2_11.insert('insert', ' '.join(i for i in list_name))  # 插入名字
        scr_sheet2_11.update()  # 插入后及时的更新
        scr_sheet2_11.see(END)  # 使得聊天记录text默认显示底端
    else:
        print('路径为空！')
        scr_output(scr_11, '\n路径为空！\n')
    # 获取名单
    people_sheet = scr_sheet2_11.get(1.0, 'end').split()
    if people_sheet != []:
        number13_11.set(people_sheet[0])
        number14_11.set(len(people_sheet))

        number15_11.set(number2_11.get()) # 更新支部请示时间
        number16_11.set(number3_11.get())
        number17_11.set(number4_11.get())

    else:
        print('名单为空，请检查！')
        scr_output(scr_11, '\n名单为空，请检查！\n')
# 支部批复模板修改，未完善
def zhibu_pifu_model_alter():
    messagebox.showinfo('小提示', '本版本只支持查看批复模板，暂不支持修改')
    cookie = str(number_11_4.get()) + str(number_11_5.get()) + str(number_11_6.get())
    if cookie == '000':
        messagebox.showinfo('错误提示', '未选中批复的类型，请检查！')
        return
    if cookie == '100':  # 发展对象的批复
        name = '发展对象'
        a = "关于同意将{}等{}人列为中共党员发展对象的批复"
        b = "中共南华大学经济管理与法学学院"
        c = "{支部全称}："
        d = "收到了贵支部{}年{}月{}日“关于建议将{}等{}人列为中共党员发展对象的请示”，且公示无异议。" \
            "认为你们按照党员标准对入党积极分子进行了有效的培养和教育。根据《中国共产党发展党员工作细则》的要求，" \
            "院党委于{}年{}月{}日召开党委会，认真讨论和审核材料，现将{}等" \
            "{}名同志列为中共党员发展对象，名单如下（排名以班级为序）："
        e = "望你们继续加强对发展对象的培养和考察。"
        f = "特此批复。"
        g = "党委书记签名：_______________"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日"
    if cookie == '010':  # 预备党员的批复
        name = '预备党员'
        a = "关于同意接收{}等{}名同志为中共预备党员的批复"
        b = "中共南华大学经济管理与法学学院"
        c = "{支部全称}："
        d = "收到了贵支部{}年{}月{}日“关于建议将{}等{}人列为中共预备党员的请示”，且公示无异议。" \
            "认为你们按照党员标准对发展对象进行了有效的培养和教育。根据《中国共产党发展党员工作细则》的要求，" \
            "院党委于{}年{}月{}日召开党委会，认真讨论和审核材料，现确定" \
            "{}等{}名同志为中共预备党员，名单如下（排名以班级为序）："
        e = "望你们继续加强对预备党员的培养和考察。"
        f = "特此批复。"
        g = "党委书记签名：_______________"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日"
    if cookie == '001':  # 预备党员转正的批复
        name = '预备党员转正'
        a = "学院党委对支部同意{}等{}名同志转为正式党员决议的审批意见"
        b = "中共南华大学经济管理与法学学院"
        c = "{支部全称}："
        d = "{}等{}名同志向党支部提出了转为正式党员的书面申请。学院党委在{}年{}月{}日召开党委会，" \
            "讨论通过你支部关于{}等{}名同志预备党员转为正式党员的决议。" \
            "{}等{}名同志从预备期满之日起成为中国共产党正式党员，党龄从即日算起。名单如下" \
            "（排名以班级为序）："
        e = ''
        f = ''
        g = "党委书记签名：_______________"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日"
    list_pifu_model = [a, b, c, d, e, f, g, h, i]

    def zhibu_pifu_model_save():
        scr_output(scr_11, '\n{}\n批复模板保存失败！，本版本模板不支持修改！\n'.format(pifu_model_var.get()))
        pifu_model.destroy()

    def zhibu_pifu_model_default():
        scr_output(scr_11, '\n模板已经是默认！\n')

    zhibu_pifu_model = Toplevel(window)
    zhibu_pifu_model.geometry("500x290+700+270")
    zhibu_pifu_model.iconbitmap('mould\ico.ico')
    # 窗口置顶
    zhibu_pifu_model.attributes("-topmost", 1)  # 1==True 处于顶层
    # 禁止窗口的拉伸
    zhibu_pifu_model.resizable(0, 0)
    # 窗口的标题
    zhibu_pifu_model.title("内置-{}-支部批复模板-修改窗口".format(name))

    # 定义变量
    zhibu_pifu_model_var = zhibu_StringVar()
    scr_zhibu_pifu_model = scrolledtext.ScrolledText(zhibu_pifu_model, wrap=WORD)
    scr_zhibu_pifu_model.place(x=10, y=10, width=480, height=245)
    scr_zhibu_pifu_model.config(state=DISABLED)  # 关闭可写入模式
    for i in list_pifu_model:
        scr_output(scr_pifu_model, str(i) + '\n')

    button_zhibu_pifu_model = ttk.Button(zhibu_pifu_model, text="保存参数", command=zhibu_pifu_model_save)
    button_zhibu_pifu_model.place(x=250, y=260)

    button_zhibu_pifu_model = ttk.Button(zhibu_pifu_model, text="恢复默认", command=zhibu_pifu_model_default)
    button_zhibu_pifu_model.place(x=120, y=260)

    # 显示窗口(消息循环)
    zhibu_pifu_model.mainloop()