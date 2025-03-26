import datetime
def is_half_year(start_date, end_date):
    """
    判断两个时间是否满半年
    :param start_date: 开始时间，格式为YYYYMMDD
    :param end_date: 结束时间，格式为YYYYMMDD
    :return: True表示满半年，False表示不满半年
    """
    start_year, start_month, start_day = int(start_date[:4]), int(start_date[4:6]), int(start_date[6:])
    end_year, end_month, end_day = int(end_date[:4]), int(end_date[4:6]), int(end_date[6:])
    start = datetime.date(start_year, start_month, start_day)
    end = datetime.date(end_year, end_month, end_day)
    interval = end - start
    half_year = datetime.timedelta(days=365/2)
    return interval >= half_year
# 测试
print()  # True
print(is_half_year('20200910', '20201231'))  # False