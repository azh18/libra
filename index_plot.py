from WindPy import w
import xlrd
import datetime
import xlwt
import numpy as np

w.start()
w.isconnected()


def format_date_value(v):
    if not v:
        return v
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        dt = xlrd.xldate_as_tuple(v, 0)
        return "%d-%02d-%02d" % (dt[0], dt[1], dt[2])
    if isinstance(v, str):
        dt = v.strip('）').strip(')').strip('受理').strip('（').strip('(')
        dt = dt.split('-')
        return "%d-%02d-%02d" % (int(dt[0]), int(dt[1]), int(dt[2]))
    raise Exception("unexpected type: {}".format(type(v)))


class Series:
    def __init__(self, date_series, value_series):
        self.date_series = date_series
        self.value_series = value_series


# 获取收盘价序列
# 要对齐时间，空的用前一天的补齐
def get_index_close_price_series(index_code_list, start, end):
    code, df = w.wsd(','.join(index_code_list), "close", start, end, "ruleType=10;PriceAdj=F", usedf=True)
    if code != 0:
        raise Exception("return non-zero code: {}".format(code))
    return df


# 将序列归一化，以第一个元素为标准
def normalize_series(series):
    value_series = series.value_series
    first_elem = value_series[0]
    normalized_value_series = np.array(value_series) / first_elem
    return Series(series.date_series, normalized_value_series)


# 作图：归一化各指数的值
def plot_series(plotname, filename, series_dict):
    pass


def write_series_to_sheet(sheet, series_dict, value_style=None):
    header = ["日期"]
    date_series = None
    for index_code in series_dict:
        if not date_series:
            date_series = series_dict[index_code].date_series
        header.append(index_code)

    for i, item in enumerate(header):
        sheet.write(0, i, item)

    for j, date in enumerate(date_series):
        v = format_date_value(date)
        sheet.write(j + 1, 0, v)

    column_idx = 1
    for index_code in series_dict:
        series = series_dict[index_code].value_series
        sheet.write(0, column_idx, index_code)
        for i, v in enumerate(series):
            if value_style:
                sheet.write(i + 1, column_idx, v, value_style)
            else:
                sheet.write(i + 1, column_idx, v)
        column_idx += 1


# 将各指数价格和归一化价格写入xls
def write_normalized_series(filename, series_dict, normalized_series_dict):
    f = xlwt.Workbook()
    index_close_price_sheet = f.add_sheet("指数收盘价", cell_overwrite_ok=True)
    normalized_price_sheet = f.add_sheet("归一化指数收盘价", cell_overwrite_ok=True)
    write_series_to_sheet(index_close_price_sheet, series_dict)
    write_series_to_sheet(normalized_price_sheet, normalized_series_dict)
    f.save(filename)


# 根据序列计算动态回撤序列
def compute_dynamic_retracement(series):
    value_series = series.value_series
    retracement_series = []
    for i, v in enumerate(value_series):
        if i == 0:
            retracement = 0
        else:
            retracement = 1 - v / (np.max(value_series[:i]))
        retracement_series.append(-retracement)
    return Series(series.date_series, np.array(retracement_series))


# 将回撤信息写入xls
def write_retracement_series(filename, series_dict):
    f = xlwt.Workbook()
    sheet = f.add_sheet("动态回撤", cell_overwrite_ok=True)
    write_series_to_sheet(sheet, series_dict, value_style=xlwt.easyxf(num_format_str='0.00%'))
    f.save(filename)


if __name__ == "__main__":
    code_list = ["892400.MI", "931463.CSI", "931476.CSI", "931465.CSI", "000970.CSI", "931148.CSI", "931466.CSI"]
    start_date = "2017-06-30"
    end_date = "2021-08-06"
    normalized_price_plot_filename = "normalized_index_price.png"
    index_price_filename = "index_price.xls"
    dynamic_retracement_plot_filename = "动态回撤.png"
    index_retracement_filename = "index_retracement.xls"

    close_price_series_dict = {}
    normalized_series_dict = {}
    dynamic_retracement_dict = {}

    close_price_series_df = get_index_close_price_series(code_list, start_date, end_date)
    date_series = list(close_price_series_df.index)

    for code in code_list:
        close_price_series = Series(date_series, np.array(close_price_series_df[code]))
        normalized_series = normalize_series(close_price_series)
        close_price_series_dict[code] = close_price_series
        normalized_series_dict[code] = normalized_series
        dynamic_retracement_dict[code] = compute_dynamic_retracement(close_price_series)

    plot_series("归一收盘价", normalized_price_plot_filename, normalized_series_dict)
    write_normalized_series(index_price_filename, close_price_series_dict, normalized_series_dict)
    write_retracement_series(index_retracement_filename, dynamic_retracement_dict)
    plot_series("动态回撤", dynamic_retracement_plot_filename, dynamic_retracement_dict)
