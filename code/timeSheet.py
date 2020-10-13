import pandas as pd
import math
import os
import calendar
import logging
import time
from chinese_calendar import is_workday, is_holiday


def num_cal(num):
    """
    超过半小时，按半小时统计
    :param num: 时长，带小数点
    :return: 时长
    """
    xs, zs = math.modf(num)
    # 超过半小时，按半小时统计
    if xs >= 0.5:
        xs = 0.5
    else:
        xs = 0
    return zs + xs


def over_time(end_time, start_time):
    """
    :param end_time: 实际结束时间(20点之后打卡的时间)
    :param start_time: 开始时间
    :return: 加班时长
    """
    result = (end_time - pd.to_datetime(str(start_time)[0:10] + " 18:00:00")) / pd.Timedelta(1, 'm') / 60
    return result


def p_end_time(end_time, end_time_1):
    """
    判断实际结束时间（实际下班打卡时间,主要处理加班到0点之后打卡的情况）
    :param end_time:
    :param end_time_1:
    :return:
    """
    act_end_time = end_time
    # 前后考勤日期相隔多少天
    days = end_time_1.day - end_time.day
    if '00:00:01' <= str(end_time_1)[11:] < '02:00:00' and days == 1:
        act_end_time = end_time_1
    return act_end_time


def is_true_false(tf):
    if tf is True:
        rs = '是'
    else:
        rs = '否'
    return rs


def weekday_cn(weekday_name):
    """
    将英文星期转换成中文星期
    :param weekday_name:英文星期
    :return:中文星期
    """
    name = ''
    if weekday_name == 'Monday':
        name = '一'
    elif weekday_name == 'Tuesday':
        name = '二'
    elif weekday_name == 'Wednesday':
        name = '三'
    elif weekday_name == 'Thursday':
        name = '四'
    elif weekday_name == 'Friday':
        name = '五'
    elif weekday_name == 'Saturday':
        name = '六'
    elif weekday_name == 'Sunday':
        name = '日'
    else:
        name = 'unKnow'
    return name


def cal_diff(start_time, end_time):
    """
    :param start_time: 进入时间
    :param end_time: 结束时间
    :return: delta + overtime: 总时长
            delta: 正常工时
            overtime: 加班总时长
            exp: 考勤异常总时长
    """
    # print('@type start_time, end_time :{0},{1}'.format(type(start_time), type(end_time)))
    # 正常考勤总时长
    delta = 0
    # 加班总时长
    overtime = 0
    # 考勤异常总时长
    exp = 0
    days = end_time.day - start_time.day
    # print('@input time :{0},{1}'.format(start_time, end_time))
    # 周末节假日加班
    if is_holiday(start_time):
        if '02:00:00' <= str(start_time)[11:]:
            # print(start_time, '是否节假日:', is_holiday(start_time))
            overtime = num_cal((end_time - start_time) / pd.Timedelta(1, 'm') / 60)
            # print('@0_1:{0},{1},{2}'.format(delta, overtime, exp))
        else:
            overtime = 0
    # 工作日考勤
    else:
        # print(start_time, '是否节假日:', is_holiday(start_time))
        # 上午正常打卡
        if str(start_time)[11:] < '09:15:00':
            if '17:30:00' <= str(end_time)[11:] < '20:00:00':
                delta = num_cal((end_time - start_time) / pd.Timedelta(1, 'm') / 60)
                # print('@0_2:{0}'.format(over_time(end_time, start_time)))
                if delta >= 8:
                    delta = 8
                # print('@1:{0},{1},{2}'.format(delta, overtime, exp))
            if '02:00:00' <= str(end_time)[11:] < '17:30:00':
                delta = 4
                exp = 4
                # print('@2:{0},{1},{2}'.format(delta, overtime, exp))
            elif '20:00:00' <= str(end_time)[11:] < '23:59:59' or (str(end_time)[11:] < '02:00:00' and days == 1):
                delta = 8
                overtime = num_cal(over_time(end_time, start_time))
                # print('@3:{0},{1},{2}'.format(delta, overtime, exp))
        # 上午异常打卡
        elif '09:15:00' <= str(start_time)[11:]:
                if '17:30:00' <= str(end_time)[11:] < '20:00:00':
                    delta = 4
                    exp = 4
                    # print('@3:{0},{1},{2}'.format(delta, overtime, exp))
                elif '20:00:00' <= str(end_time)[11:] < '23:59:59' or (str(end_time)[11:] < '02:00:00' and days == 1):
                    delta = 4
                    exp = 4
                    overtime = num_cal(over_time(end_time, start_time))
                    # print('@4:{0},{1},{2}'.format(delta, overtime, exp))
        # print('@5:{0},{1},{2}'.format(delta, overtime, exp))
    return delta + overtime, delta, overtime, exp


def prefix(_file_url):
    """
    去掉考勤记录的文件格式
    :param _file_url: 考勤记录的文件路径
    :return: 去掉考勤记录的文件格式的路径
    """
    return _file_url.replace('.xlsx', '').replace('.xls', '')


def emp_st2(_kqmd_url):
    """
    获得考勤名单中Sheet2的内容
    :param _kqmd_url: 考勤名单的路径
    :return: 人员名单的dataframe
    """
    if os.path.exists(_kqmd_url):
        print(_kqmd_url + '文件存在，程序将自动读取此文件的名单。')
    else:
        print('考勤名单.xlsx文件不存在，程序将自动生成名单文件' + _kqmd_url)
        # emp_jk_df.to_csv(emp_jk_file, encoding='gbk', index=_kqmd_url
        emp_writer = pd.ExcelWriter(_kqmd_url)
        emp_jk = (['测试'])
        emp_jk_df = pd.DataFrame(columns=['人员'], data=emp_jk).rename_axis('序号').reset_index()
        emp_jk_df.to_excel(excel_writer=emp_writer, sheet_name='Sheet2', index=False, index_label=None)
        emp_writer.save()
        emp_writer.close()

    # df: 考勤人员名单sheet2的dataframe
    df_st2 = pd.read_excel(_kqmd_url, sheet_name='Sheet2', encoding='gbk', header=0)
    df_st2['value'] = 0
    df_st2.columns = ['id', '职工姓名', 'value']
    # print('df_st2\n', df_st2)
    return df_st2


def emp_st1(_kqmd_url):
    """
    获得考勤名单中Sheet1的内容
    :param _kqmd_url: 考勤名单的路径
    :return: 人员名单的dataframe
    """
    df_st1 = pd.read_excel(_kqmd_url, sheet_name='Sheet1', header=0).reset_index()
    df_st1['value'] = 0
    df_st1.columns = ['id', '参与项目', '职工姓名', '考勤编号', '角色', '办公场地', 'value']
    # print('df_st1\n', df_st1)
    return df_st1


def time_sheet(_sheet_context, list_emp, timesheet_cate):
    """
    处理考勤记录，获得最小时间、最大时间、工作时长的dataframe
    :param _sheet_context: 考勤记录dataframe
    :param list_emp: 考勤名单的list
    :param timesheet_cate: 考勤记录类型
    :return: 处理完成的考勤记录dataframe
    """
    if timesheet_cate == '1':
        # 过滤出人员名单的timesheet考勤记录
        timesheet_df = _sheet_context[_sheet_context['职工姓名'].isin(list_emp)]\
            .groupby(['进入日期', '职工姓名'], as_index=False).agg({'attend_datetime': ['min',  'max']})
        if timesheet_df.empty:
            return None
        else:
            timesheet_df.columns = ['进入日期', '职工姓名', '开始时间', '结束时间']
            # 重新排序的目的是为了timesheet_df.shift(-1)处理跨天的问题
            timesheet_df.sort_values(by=['职工姓名', '开始时间'], ascending=[True, True], inplace=True)
            timesheet_df.reset_index(drop=True, inplace=True)
            # 跨天打卡时间(超过12点打卡)
            timesheet_df['结束时间_1'] = timesheet_df['开始时间'].shift(-1).fillna(pd.to_datetime('1900-12-31 23:59:59'))
            timesheet_df['实际结束时间'] = timesheet_df.apply(lambda x: p_end_time(x['结束时间'], x['结束时间_1']), axis=1)
            timesheet_df['总时长'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[0], axis=1)
            timesheet_df['正常工时'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[1], axis=1)
            timesheet_df['加班工时'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[2], axis=1)
            timesheet_df['异常工时'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[3], axis=1)

    elif timesheet_cate == '2':
        # 过滤出人员名单的timesheet考勤记录
        timesheet_df = _sheet_context[_sheet_context['职工姓名'].isin(list_emp)].copy()
        if timesheet_df.empty:
            return None
        else:
            timesheet_df['总时长'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[0], axis=1)
            timesheet_df['正常工时'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[1], axis=1)
            timesheet_df['加班工时'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[2], axis=1)
            timesheet_df['异常工时'] = timesheet_df.apply(lambda x: cal_diff(x['开始时间'], x['实际结束时间'])[3], axis=1)
    return timesheet_df


if __name__ == '__main__':

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)  # Log等级总开关

    timesheet_cate = input("请确认考勤记录格式，金 科输入1(默认)，涛 飞输入2，请输入:")
    if len(timesheet_cate) == 0:
        timesheet_cate = '1'
    while timesheet_cate not in ('1', '2'):
        input('输入有误，需要输入数字1，2 。金 科输入1(默认)，涛 飞输入2，请重新输入:')
    timesheet_cate_name = '金 科' if timesheet_cate == '1' else '涛 飞'
    if len(timesheet_cate) == 1:
        print('输入的考勤记录为: {0}, 考勤记录格式：{1}'.format(timesheet_cate, timesheet_cate_name))

    args_file_url = input("请输入考勤记录文件路径:")
    print('输入的考勤记录文件路径为: ', args_file_url)
    file_url = args_file_url
    while len(file_url) == 0:
        print('路径为空。')
        args = input("请输入考勤记录文件路径:")
        print('输入的考勤记录文件路径为: ', args_file_url)
        file_url = args_file_url

    print('开始统计')
    time.sleep(1)

    excel_file_name = '.xlsx'
    out_date = '_' + pd.datetime.now().strftime('%Y%m%d%H%M%S')
    
    # 读取考勤记录
    try:
        sheet_context = pd.read_excel(file_url, header=0)
    except:
        sheet_context = pd.read_excel(file_url, header=0)

    # sheet1的考勤人员名单属性(参与参与项目/人员/考勤编号/角色/办公场地)
    df_emp_st1 = emp_st1(kqmd_url)
    # print('df_emp_st1\n', df_emp_st1)
    list_emp_st1 = df_emp_st1['职工姓名'].tolist()
    # sheet2的考勤人员名单
    df_emp_st2 = emp_st2(kqmd_url)
    list_emp_st2 = df_emp_st2['职工姓名'].tolist()

    # 处理考勤记录
    # 进入日期 和 进入时间字段合并
    # print('@@sheet_context = \n', sheet_context)
    if timesheet_cate == '1':
        sheet_context['attend_datetime'] = pd.to_datetime(sheet_context['进入日期'].astype(str).str[:10]
                                                          + " " + sheet_context['进入时间'].map(str))
    elif timesheet_cate == '2':
        sheet_context.columns = ["部门名称", "人员编号", "职工姓名", "进入日期", "最早打卡时间", "最晚打卡时间"]
        sheet_context['进入日期'] = pd.to_datetime(sheet_context['进入日期'], format="%Y-%m-%d", errors='raise')
        sheet_context['开始时间'] = pd.to_datetime(sheet_context['最早打卡时间'], format="%Y-%m-%d %H:%M:%S", errors='raise')
        sheet_context['实际结束时间'] = pd.to_datetime(sheet_context['最晚打卡时间'], format="%Y-%m-%d %H:%M:%S", errors='raise')

    # 过滤出人员名单的timesheet考勤记录
    try:
        time_sheet_st1 = time_sheet(sheet_context, list_emp_st1, timesheet_cate)

        time_sheet_st2 = time_sheet(sheet_context, list_emp_st2, timesheet_cate)
    except (SystemExit, KeyboardInterrupt):
        raise
    except Exception as e:
        logger.error('处理考勤记录异常', e.errno, e.strerror)

    # 月份所有日期
    att_days = sheet_context['进入日期'].min()
    att_year = sheet_context['进入日期'].min().year
    att_month = sheet_context['进入日期'].min().month

    last_days = calendar.monthrange(att_year, att_month)[1]
    idx = pd.date_range(att_days, freq='D', periods=last_days)
    dim_date = pd.DataFrame(pd.Series(0, idx)).reset_index()
    dim_date.columns = ['日期', 'value']
    # dim_date['日期'] = dim_date['日期'].astype(datetime.date)
    dim_date['weekday_name'] = dim_date['日期'].dt.weekday_name
    dim_date['星期'] = dim_date.apply(lambda x: weekday_cn(x['weekday_name']), axis=1)
    dim_date['是否节假日'] = dim_date['日期'].apply(lambda x: is_true_false(is_holiday(x)))
    dim_date['weekofyear'] = dim_date['日期'].dt.weekofyear
    # print('dim_date dtypes: \n', dim_date.dtypes)
    # print('time_sheet_st1 dtypes: \n', time_sheet_st1.dtypes)
    # ######################## sheet1 ########################################begin
    if time_sheet_st1 is not None:
        try:
            ts1 = pd.merge(dim_date, df_emp_st1, on='value')
            df1 = pd.merge(ts1, time_sheet_st1, how='left', left_on=['日期', '职工姓名'], right_on=['进入日期', '职工姓名'])
            # print('@@_1: df1\n', df1)
            df1['日期'] = df1['日期'].astype(str).str[:10]
            df1['正常工时'] = df1['正常工时']/8
            df1['加班工时'] = round(df1['加班工时'] / 8, 4)
            df1['总时长'] = round(df1['总时长'] / 8, 4)
            df1['异常工时'] = df1['异常工时'] / 8
            df1['公司名'] = '恒格'
            df1['年月'] = (df1['日期'].astype(str).str[:7]).apply(lambda x: int(x.replace("-", "")))
            df1['备注'] = ''
            attend1 = df1[['公司名', '年月', '参与项目', '职工姓名', '考勤编号', '角色', '日期', '开始时间', '实际结束时间',
                           '正常工时', '加班工时', '是否节假日', '备注', '异常工时', '总时长']].fillna(0)
            attend1.sort_values(by=['职工姓名', '日期'], ascending=[True, True], inplace=True)
            attend1.reset_index(drop=True, inplace=True)
            attend1.rename(columns={'开始时间': '当天最早打卡时间', '实际结束时间': '当天最晚打卡时间'}, inplace=True)
            print('===============')
            print('Sheet1统计结果:')
            print('===============')
            print(attend1.head(5))
            writer1 = pd.ExcelWriter(prefix(file_url) + out_date + '_Sheet1生成' + excel_file_name)
            attend1.to_excel(writer1, index=False, index_label=None)
            writer1.save()
            writer1.close()
        except Exception as e:
            logger.error('生成考勤统计记录1异常', e.errno, e.strerror)

    # ######################## sheet1 ########################################end

    # ######################## sheet2 ########################################begin
    # print('---')
    if time_sheet_st2 is not None:
        try:
            ts2 = pd.merge(dim_date, df_emp_st2, on='value')
            df2 = pd.merge(ts2, time_sheet_st2, how='left', left_on=['日期', '职工姓名'], right_on=['进入日期', '职工姓名'])
            # print('@@_1: df2\n', df2)
            df2['日期'] = df2['日期'].astype(str).str[:10]
            attend2 = df2[['日期', '星期', 'weekofyear', 'id', '职工姓名', '总时长']].fillna(0)
            attend_pivot2 = pd.pivot_table(attend2, index=['id', '职工姓名'], columns=['weekofyear', '日期', '星期'], values='总时长')
            print('===============')
            print('Sheet2统计结果:')
            print('===============')
            print(attend_pivot2.head(5))

            writer2 = pd.ExcelWriter(prefix(file_url) + out_date + '_Sheet2生成' + excel_file_name)
            attend_pivot2.to_excel(writer2)
            writer2.save()
            writer2.close()
        except Exception as e:
            logger.error('生成考勤统计记录1异常', e.errno, e.strerror)
    # ######################## sheet2 ########################################end
    print('完成！')
    time.sleep(3)

