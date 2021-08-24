from WindPy import w
import xlrd
import datetime
import xlwt

w.start()
w.isconnected()


# input: like 000300.SH
def get_index_weight_data(index_code):
    data = w.wset("indexconstituent", "windcode=%s" % index_code)
    if data.ErrorCode != 0:
        raise Exception("error_code={}".format(data.ErrorCode))
    return data.Data


def format_date_value(v):
    if not v:
        return v
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        dt = xlrd.xldate_as_tuple(v, 0)
        return "%d-%02d-%02d" % (dt[0], dt[1], dt[2])
    if isinstance(v, str):
        dt = v.strip('）').strip(')').strip('受理').strip('（').strip('(')
        dt = dt.split('-')
        return "%d-%02d-%02d" % (int(dt[0]), int(dt[1]), int(dt[2]))
    raise Exception("unexpected type: {}".format(type(v)))


def gen_index_weight_file(filename, index_code_list):
    f = xlwt.Workbook()
    header = ["时间", "成分股代码", "成分股名称", "权重", "分类"]

    for index_code in index_code_list:
        print("gen weight for {}".format(index_code))
        sheet = f.add_sheet(index_code, cell_overwrite_ok=True)
        for i, item in enumerate(header):
            sheet.write(0, i, item)
        data = get_index_weight_data(index_code)
        for j, column in enumerate(data):
            for i, value in enumerate(column):
                if j == 0:
                    value = format_date_value(value)
                sheet.write(i + 1, j, value)

    f.save(filename)


if __name__ == "__main__":
    # 这里修改输出的文件名
    output_file = "指数权重测试.xls"
    # 这里修改指数列表
    index_code_list = ["000300.SH", "000985.CSI"]
    gen_index_weight_file(output_file, index_code_list)
