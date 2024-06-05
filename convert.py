import pandas
import xlwt
from xmindparser import xmind_to_dict
import os


def handle_xmind(filename):
    out = xmind_to_dict(filename)
    return out[0]['topic']['topics']


def handle_path(topic, topics_lists, title):
    """
    遍历解析后的xmind数据
    :param topic:xmind解析后的数据
    :param topics_lists:从topic提取的数据存放的列表
    :param title:从topic提取的title
    :return:
    """
    # title去除首尾空格
    title = title.strip()
    # 如果调用本方法，则concatTitle赋值为第一个层级的名称（所属模块名称）
    if len(title) == 0:
        concatTitle = topic['title'].strip()
    else:
        try:
            concatTitle = title + '|' + topic['title'].strip() + topic['makers'][0]
        except:
            concatTitle = title + '|' + topic['title'].strip()
    # 如果第一层下没有数据了，那么就把第一层级的名称填加到topics_lists里
    if topic.__contains__('topics') == False:
        topics_lists.append(concatTitle)
    # 如果有，那么就遍历下一个层级......这样就把所有title都写到topics_lists里了
    else:
        for d in topic['topics']:
            handle_path(d, topics_lists, concatTitle)


def handle_title(topics):
    """
    判断是title的类别：所属模块，所属子模块，用例，操作步骤，预期结果，备注
    :param topics: 传入topic的列表
    :return: 经过处理的字典，{'model': '...','case': '...', 'step': '...', 'expect': '...'}

    """
    list = []
    for l in topics:
        dict = {}
        for i in l:
            if "模块" in i.split("_"):
                dict["model"] = i[3:]
            elif ("需求：" in i) or ("需求:" in i):
                dict["story"] = i[3:]
            elif ("功能细项：" in i) or ("功能细项:" in i):
                dict["function"] = i[5:]
            elif ("用例：" in i) or ("用例:" in i):
                dict["case"] = i[3:]
            elif ("步骤：" in i) or ("步骤:" in i):
                dict["step"] = i[3:]
            elif ("预期：" in i) or ("预期:" in i):
                dict["expect"] = i[3:]
            elif "priority" in i:
                dict["step"] = i.split("priority-")[0]
                dict["expect"] = i.split("priority-")[0]
                dict["priority"] = i.split("priority-")[1]
            else:
                try:
                    try:
                        dict["case"] = dict["case"] + "_" + i
                    except:
                        dict["case"] = i
                except:
                    pass
        list.append(dict)
    return list


def handle_topics(topics):
    index = 0
    title_lists = []
    # 遍历第一层级的分支（所属模块）
    for h in topics:
        topics_lists = []
        handle_path(h, topics_lists, '')
        # print("这个是lists,", topics_lists)
        # 取出topics下所有的title放到1个列表中
        for j in range(0, len(topics_lists)):
            title_lists.append(topics_lists[j].split('|'))
    return title_lists


def write_to_temp1(list, excelname):
    """
    把解析的xmind文件写到xls文件里
    使用禅道模板
    :param list: 经过处理的字典，格式 [{'model': '...', 'sub_model': '...', 'case': '...', 'step': '...','expect': '.'},{...},...]
    :param excelname: 输出的excel文件路径，文件名固定为：xmind文件名.xls
    :return:
    """
    f = xlwt.Workbook()
    # 生成excel文件
    sheet = f.add_sheet('测试用例', cell_overwrite_ok=True)
    row0 = ['所属模块', '用例标题', '步骤', '预期', '优先级', '用例类型']
    # 生成第一行中固定表头内容
    for i in range(0, len(row0)):
        sheet.write(0, i, row0[i])
    # 把title写入xls
    for index, d in enumerate(list):
        # 第二列及之后的用例数据
        try:
            sheet.write(index + 1, 0, d["model"])
        except:
            pass
        try:
            sheet.write(index + 1, 1, d["case"])
        except:
            pass
        try:
            sheet.write(index + 1, 2, d["step"])
        except:
            pass
        try:
            sheet.write(index + 1, 3, d["expect"])
        except:
            pass
        try:
            sheet.write(index + 1, 4, d["priority"])
        except:
            sheet.write(index + 1, 4, "3")
        sheet.write(index + 1, 5, "功能测试")
    f.save(excelname)
    csvname = excelname.split(".xls")[0] + ".csv"
    ex = pandas.read_excel(excelname)
    ex.to_csv(csvname, encoding="gbk", index=False)


def write_to_temp_jira(list, excelname):
    """
    把解析的xmind文件写到xls文件里
    使用JIRA模板
    :param list: 经过处理的字典，格式 [{'model': '...', 'sub_model': '...', 'case': '...', 'step': '...','expect': '.'},{...},...]
    :param excelname: 输出的excel文件路径，文件名固定为：xmind文件名.xls
    :return:
    """
    f = xlwt.Workbook()
    # 生成excel文件
    sheet = f.add_sheet('测试用例', cell_overwrite_ok=True)
    row0 = ['Team', 'TCID', 'Pre_Condition', '任务分类-TOC', 'Issue_type', 'Test-Set6', 'Summary', 'Component',
            'Description',
            'Fix_Version', 'Priority', 'Lables', 'Test_type', 'assignee', 'Step', 'Data', 'Expected_Result']
    # 生成第一行中固定表头内容
    for i in range(0, len(row0)):
        sheet.write(0, i, row0[i])
    # 把title写入xls
    for index, d in enumerate(list):
        # 第二列及之后的用例数据
        sheet.write(index + 1, 3, "用例设计")
        sheet.write(index + 1, 4, "Test")
        try:
            sheet.write(index + 1, 6, d["case"])
        except:
            pass

        if d.get("model") and d.get("sub_model"):
            sheet.write(index + 1, 11, d["model"] + '_' + d["sub_model"])
        elif d.get("model") and d.get("sub_model") == None:
            sheet.write(index + 1, 11, d["model"])
        elif d.get("model") == None and d.get("sub_model"):
            sheet.write(index + 1, 11, d["sub_model"])
        else:
            pass
    try:
        sheet.write(index + 1, 12, "Manual")
    except:
        pass
    try:
        sheet.write(index + 1, 14, d["step"])
    except:
        pass
    try:
        sheet.write(index + 1, 16, d["expect"])
    except:
        pass

    f.save(excelname)


def write_to_temp2(list, excelname):
    """
    把解析的xmind文件写到xls文件里
    使用集成测试用例模板
    :param list: 经过处理的字典，格式 [{'model': '...', 'sub_model': '...', 'case': '...', 'step': '...','expect': '.'},{...},...]
    :param excelname: 输出的excel文件路径，文件名固定为：xmind文件名.xls
    :return:
    """
    f = xlwt.Workbook()
    # 生成excel文件
    sheet = f.add_sheet('测试用例', cell_overwrite_ok=True)
    row0 = ['用例目录', '所属需求', '功能细项',  '用例名称', '前置条件', '用例步骤', '预期结果', '用例类型', '用例状态', '用例等级', '创建人']
    # 生成第一行中固定表头内容
    for i in range(0, len(row0)):
        sheet.write(0, i, row0[i])
    # 把title写入xls
    for index, d in enumerate(list):
        # 第二列及之后的用例数据
        try:
            sheet.write(index + 1, 0, d["model"])
        except:
            pass
        try:
            sheet.write(index + 1, 1, d["story"])
        except:
            pass
        try:
            sheet.write(index + 1, 2, d["function"])
        except:
            pass
        try:
            sheet.write(index + 1, 3, d["case"])
        except:
            pass
        try:
            sheet.write(index + 1, 5, d["step"])
        except:
            pass
        try:
            sheet.write(index + 1, 6, d["expect"])
        except:
            pass
        sheet.write(index + 1, 7, "功能测试")
        sheet.write(index + 1, 8, "正常")
        sheet.write(index + 1, 9, "高")
        sheet.write(index + 1, 10, "王沁")
    f.save(excelname)
