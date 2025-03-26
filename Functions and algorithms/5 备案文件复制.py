def beian_model_cookie(cookie,yeardu,pici,year_up, year,month,day, dw_year,dw_month,dw_day, first_people,people_num,people_sheet):
    if cookie == '10':  # 预备党员的备案报告
        a = "经济管理与法学学院预备党员报组织部备案报告"
        b = "校党委组织部："
        c = "学院党委于{}年{}月{}日召开党委会，现确定{}等{}名同志为中共党员预备党员，" \
            "名单如下（排名以班级为序）：".format(dw_year,dw_month,dw_day,first_people,people_num)
        d = "学院将继续加强对预备党员的培养和考察。"
        e = "特此报告。"
        f = "中共南华大学经济管理与法学学院委员会（盖章）"
        g = "{}年{}月{}日".format(year, month, day)
    if cookie == '01':  # 预备党员转正的备案报告
        a = "经济管理与法学学院转正党员报组织部备案报告"
        b = "校党委组织部："
        c = "学院党委于{}年{}月{}日召开党委会，现确定{}等{}名同志为中共党员，" \
            "名单如下（排名以班级为序）：".format(dw_year,dw_month,dw_day,first_people,people_num)
        d = "学院将继续加强对党员的培养和考察。"
        e = "特此报告。"
        f = "中共南华大学经济管理与法学学院委员会（盖章）"
        g = "{}年{}月{}日".format(year, month, day)

    return a,b,c,d,e,f,g