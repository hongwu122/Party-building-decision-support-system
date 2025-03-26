# 批复文件cookie的模板的识别（后续开发）
def pifu_model_cookie(cookie, yeardu, pici, year_up, qs_year,qs_month,qs_day,qingshi_name, year, month, day,party_name, party_num, first_people, people_num,people_sheet):

    if cookie == '100':  # 发展对象的批复
        a = "关于同意接收{}年度经济管理与法学学院".format(yeardu)
        b = "{}发展对象的批复".format(pici)
        c = "中共南华大学经济管理与法学学院"
        d = "{}等{}个学生党支部：".format(party_name,party_num)
        e = "收到了贵支部{}年{}月{}日“{}”，且公示无异议。" \
            "认为你们按照党员标准对入党积极分子进行了有效的培养和教育。根据《中国共产党发展党员工作细则》的要求，" \
            "院党委于{}年{}月{}日召开党委会，认真讨论和审核材料，现将{}等" \
            "{}名同志列为中共党员发展对象，名单如下（排名以班级为序）：" \
            "".format(qs_year,qs_month,qs_day,qingshi_name,year,month,day,first_people,people_num)
        f = "望你们继续加强对发展对象的培养和考察。"
        g = "特此批复。"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日".format(year, month, day)
    if cookie == '010':  # 预备党员的批复
        a = "关于同意接收{}年度经济管理与法学学院".format(yeardu)
        b = "{}预备党员的批复".format(pici)
        c = "中共南华大学经济管理与法学学院"
        d = "{}等{}个学生党支部：".format(party_name,party_num)
        e = "收到了贵支部{}年{}月{}日“{}”，且公示无异议。" \
            "认为你们按照党员标准对发展对象进行了有效的培养和教育。根据《中国共产党发展党员工作细则》的要求，" \
            "院党委于{}年{}月{}日召开党委会，认真讨论和审核材料，现确定" \
            "{}等{}名同志为中共预备党员，名单如下（排名以班级为序）：" \
            "".format(qs_year,qs_month,qs_day,qingshi_name,year,month,day,first_people,people_num)
        f = "望你们继续加强对预备党员的培养和考察。"
        g = "特此批复。"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日".format(year, month, day)
    if cookie == '001':  # 预备党员转正的批复
        a = "关于同意接收{}年度经济管理与法学学院".format(yeardu)
        b = "{}预备党员转正的批复".format(pici)
        c = "中共南华大学经济管理与法学学院"
        d = "{}等{}个学生党支部：".format(party_name,party_num)
        e = "收到了贵支部{}年{}月{}日“{}”，且公示无异议。" \
            "院党委于{}年{}月{}日召开党委会，讨论通过{{first_people}}等{{people_num}}名同志预备党员转为正式党员的决议，名单如下" \
            "（排名以班级为序）：".format(qs_year,qs_month,qs_day,qingshi_name,year,month,day,first_people,people_num)
        f = None
        g = "特此批复。"
        h = "中共南华大学经济管理与法学学院委员会（盖章）"
        i = "{}年{}月{}日".format(year, month, day)
    return a, b, c, d, e, f, g, h, i
# 批复文件的写入




    


