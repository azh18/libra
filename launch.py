from cat import format_date_value
import xlrd
import xlwt
import datetime

# ["基金代码", "证券简称", "认购起始日", "认购截止日", "基金成立日（排序项）", "发行总份额(亿份）", "管理人", "托管人",
#  "投资类型(wind一级分类)", "投资类型(wind二级分类)", "一级分类", "二级分类", "批文时间", "证券全称"]

launch_header = ["基金代码", "证券简称", "认购起始日", "认购截止日", "基金成立日（排序项）", "发行总份额(亿份）", "管理人", "托管人",
                 "投资类型(wind一级分类)", "投资类型(wind二级分类)", "一级分类", "二级分类", "批文时间", "证券全称"]


def load_launch_data(path):
    launch_dict = {}  # code -> launch_row
    sheet = xlrd.open_workbook(path).sheets()[2]
    for i in range(sheet.nrows):
        row = sheet.row_values(i)
        if not isinstance(row[0], str) or len(row[0]) < 7 or not row[0][:6].isdigit():
            continue
        code = row[0]

        lrow = ['' for _ in range(14)]
        lrow[0:2] = row[0:2]
        lrow[2] = format_date_value(row[2])
        lrow[3] = format_date_value(row[3])
        lrow[4] = format_date_value(row[4])
        lrow[5:12] = row[5:12]
        lrow[12] = format_date_value(row[12])
        lrow[13] = row[13]
        launch_dict[code] = lrow
    return launch_dict


wd_header = ["证券代码", "证券简称", "基金管理人", "基金托管人", "基金获批注册日期", "发行公告日", "发行日期", "个人投资者认购终止日",
             "机构投资者设立认购终止日", "基金成立日", "发行总份额[单位] 亿份", "上市日期", "投资类型(一级分类)", "投资类型(二级分类)",
             "基金全称", "基金简称", "基金简称", "基金简称"]


def load_wd_dict(path):
    wd_dict = {}  # code -> launch_row
    sheet = xlrd.open_workbook(path).sheets()[0]
    for i in range(sheet.nrows):
        row = sheet.row_values(i)
        if not isinstance(row[0], str) or len(row[0]) < 7 or not row[0][:6].isdigit():
            continue
        code = row[0]

        lrow = ['' for _ in range(17)]
        lrow[0:3] = row[0:3]
        for j in range(4, 10):
            lrow[j] = format_date_value(row[j])
        lrow[10] = row[10]
        lrow[11] = format_date_value(row[11])
        lrow[12:18] = row[12:18]
        wd_dict[code] = lrow
    return wd_dict


# 用最新的万德数据补充发行数据
# wd_header = ["证券代码", "证券简称", "基金管理人", "基金托管人", "基金获批注册日期", "发行公告日", "发行日期", "个人投资者认购终止日",
#              "机构投资者设立认购终止日", "基金成立日", "发行总份额[单位] 亿份", "上市日期", "投资类型(一级分类)", "投资类型(二级分类)",
#              "基金全称", "基金简称", "基金简称", "基金简称"]
# launch_header = ["基金代码", "证券简称", "认购起始日", "认购截止日", "基金成立日（排序项）", "发行总份额(亿份）", "管理人", "托管人",
#                  "投资类型(wind一级分类)", "投资类型(wind二级分类)", "一级分类", "二级分类", "批文时间", "证券全称"]
def fulfill_launch_data(wd_dict, launch_dict):
    for wd_row in wd_dict.values():
        wd_code = wd_row[0]
        if wd_code not in launch_dict:
            launch_dict[wd_code] = ["" for _ in range(14)]
        # else:
        #     print(launch_dict[wd_code])
        launch_dict[wd_code][0] = wd_row[0]
        launch_dict[wd_code][1] = wd_row[1]
        launch_dict[wd_code][2] = wd_row[6]
        launch_dict[wd_code][3] = max(wd_row[7], wd_row[8])
        launch_dict[wd_code][4] = wd_row[9]
        launch_dict[wd_code][5] = wd_row[10]
        launch_dict[wd_code][6:8] = wd_row[2:4]
        launch_dict[wd_code][8:10] = wd_row[12:14]
        launch_dict[wd_code][12] = wd_row[4]
        launch_dict[wd_code][13] = wd_row[14]
        # todo: 自动填充一级分类&二级分类？
    launch_rows = launch_dict.values()
    settled_rows = filter(lambda x: len(x[4]) > 5, launch_rows)
    unsettled_rows = filter(lambda x: len(x[4]) <= 5, launch_rows)
    target_rows = sorted(settled_rows, key=lambda x: x[4], reverse=False)
    target_rows = target_rows + list(unsettled_rows)
    return target_rows


def launch_row_filter_rule(x):
    if "资产管理计划" in x[13]:
        return False
    if x[1].strip()[-1] == "E":
        return False
    if x[2] < "2021-01-01":
        return False
    return x[4] >= format_date_value("2021-01-01") or len(x[4]) < 6


def filter_launch_data(launch_rows):
    return list(filter(launch_row_filter_rule, launch_rows))


def write_launch_data(launch_rows, path):
    f = xlwt.Workbook()
    easy_sheet = f.add_sheet('发行数据', cell_overwrite_ok=True)
    print('发行数据条目：%d' % len(launch_rows))
    temp_rows = [launch_header] + launch_rows
    for i, row in enumerate(temp_rows):
        wd_code = row[0]
        if wd_code == "010471.OF":
            print(row)
        print('add %d' % i)
        for j, item in enumerate(row):
            if wd_code == "010471.OF":
                print("%s: %s" % (launch_header[j], item))
            easy_sheet.write(i, j, item)
    f.save(path)


if __name__ == "__main__":
    wd_path = "wd.xlsx"
    launch_path = "基金行业数据--20210818.xlsx"
    new_launch_path = "发行数据--{}.xls".format(datetime.date.today().strftime("%Y%m%d"))

    cur_launch_dict = load_launch_data(launch_path)
    cur_wd_dict = load_wd_dict(wd_path)
    sorted_launch_rows = fulfill_launch_data(cur_wd_dict, cur_launch_dict)
    filtered_launch_rows = filter_launch_data(sorted_launch_rows)
    write_launch_data(filtered_launch_rows, new_launch_path)
