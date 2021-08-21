# 基金名称匹配（v2.0）
import os
import xlrd
import xlwt
import multiprocessing
from multiprocessing import Manager, Pool, RLock, Lock
import time
import re

word_replace_dict = {
    "国海富兰克林": "国海",
    "富兰克林国海": "国海",
    "富兰克林": "国海",
    "中国银河": "银河",
    "银河证券": "银河",
    "中国国际金融": "中金公司",
    "上海浦发银行": "浦发银行",
    "浦东发展银行": "浦发银行",
    "上海浦东银行": "浦发银行",
    "东方红": "东方证券"
}


def regulate_string(s):
    s = s.strip().replace('（', '(').replace('）', ')')
    for k in word_replace_dict:
        s = s.replace(k, word_replace_dict[k])
    return s


match_regex_dict = {}
chinese_re = re.compile(r'[^\x00-\x2f,\x3a-\xff]')


def match_string(a, b):
    short, long = regulate_string(a), regulate_string(b)
    if len(short) > len(long):
        short, long = long, short
    if short in match_regex_dict:
        regex = match_regex_dict[short]
    else:
        short_ch = chinese_re.findall(short)
        pattern = r"\S*".join(short_ch)
        # print(pattern)
        regex = re.compile(pattern)
        match_regex_dict[short] = regex
    if regex.search(long):
        return True
    else:
        return False

        # long_idx = 0
        # for i in range(len(short)):
        #     while short[i] != long[long_idx]:
        #         long_idx += 1
        #         if long_idx >= len(long):
        #             return False
        #         continue
        # return True


# match_string('大成成长回报六个月持有期混合型证券投资基金', '惠弘定期开放纯债债券型')


# 拉取万德数据
# ('管理人', '托管人', '基金全称', 发行公告日期', '认购起始日',' 认购截止日', '基金成立日', '发行总份额',
# 'Wind一级分类', 'Wind二级分类')
def load_wd_data(wd_file):
    wd_rows = []
    sheet = xlrd.open_workbook(wd_file).sheets()[0]
    for i in range(sheet.nrows):
        row = sheet.row_values(i)
        if not isinstance(row[0], str) or len(row[0]) < 7 or not row[0][:6].isdigit():
            continue
        wd_row = ['' for j in range(10)]
        wd_row[0] = row[2]
        wd_row[1] = row[3]
        wd_row[2] = row[14]
        wd_row[3:5] = row[5:7]
        wd_row[5] = row[7]
        if wd_row[5] > row[8]:
            wd_row[5] = row[8]
        wd_row[6] = row[9]
        wd_row[7] = row[10]
        wd_row[8:10] = row[12:14]
        for j in range(len(wd_row)):
            if isinstance(wd_row[j], str):
                wd_row[j] = wd_row[j].strip()
        wd_rows.append(wd_row)
    return wd_rows


def find_wd_row(wd_rows, guanli, tuoguan, quancheng):
    for wd_row in wd_rows:
        if match_string(guanli, wd_row[0]) and match_string(tuoguan, wd_row[1]) and match_string(quancheng, wd_row[2]):
            return wd_row
    return None


# wd_db = load_wd_data('wd.xlsx')
# row = find_wd_row(wd_db, '广发', '建设银行', '广发安泰稳健养老目标一年持有期发起式基金中基金（FOF）')
# print(row)


# load data into memory
def load_db(result_xls_path):
    easy_db = []
    normal_db = []
    if os.path.exists(result_xls_path):
        result_easy_sheet = xlrd.open_workbook(result_xls_path).sheets()[0]
        result_normal_sheet = xlrd.open_workbook(result_xls_path).sheets()[1]
        nrows_easy = result_easy_sheet.nrows
        nrows_normal = result_normal_sheet.nrows
        for i in range(nrows_easy):
            if i < 1:
                continue
            easy_db_row = result_easy_sheet.row_values(i)[:20]
            for j in range(len(easy_db_row)):
                if isinstance(easy_db_row[j], str):
                    easy_db_row[j] = easy_db_row[j].strip()
            easy_db.append(easy_db_row)
        for i in range(nrows_normal):
            if i < 1:
                continue
            normal_db_row = result_normal_sheet.row_values(i)[:20]
            for j in range(len(normal_db_row)):
                if isinstance(normal_db_row[j], str):
                    normal_db_row[j] = normal_db_row[j].strip()
            normal_db.append(normal_db_row)

    return easy_db, normal_db


easy_title = ['基金管理人', '基金托管人', '申请事项', '申请日', '受理日（排序项）', '决定日', '发行公告日期', '认购起始日', ' 认购截止日',
              '基金成立日', '发行总份额', 'Wind一级分类', 'Wind二级分类', '一级分类', '二级分类', '是否为发起式', '是否为定期开放', '是否为港股通/香港主题',
              '是否为变更注册', '若为变更注册，原申请事项名称/代码']


def store_db(easy_db, normal_db, file):
    f = xlwt.Workbook()
    easy_sheet = f.add_sheet('简易程序', cell_overwrite_ok=True)
    print('简易程序条目：%d' % len(easy_db))
    temp_rows = [easy_title] + easy_db
    for i, row in enumerate(temp_rows):
        print('add %d' % i)
        for j, item in enumerate(row):
            easy_sheet.write(i, j, item)
    normal_sheet = f.add_sheet('普通程序', cell_overwrite_ok=True)
    print('普通程序条目：$%d' % len(normal_db))
    temp_rows = [easy_title] + normal_db
    for i, row in enumerate(temp_rows):
        print('add %d' % i)
        for j, item in enumerate(row):
            normal_sheet.write(i, j, item)
    f.save(file)


def format_date_value(v):
    if not v:
        return v
    if isinstance(v, float):
        dt = xlrd.xldate_as_tuple(v, 0)
        return "%d-%02d-%02d" % (dt[0], dt[1], dt[2])
    if isinstance(v, str):
        dt = v.strip('）').strip(')').strip('受理').strip('（').strip('(')
        dt = dt.split('-')
        if len(dt) != 3:
            return v
        return "%d-%02d-%02d" % (int(dt[0]), int(dt[1]), int(dt[2]))
    raise Exception("unexpected type: {}".format(type(v)))

def format_easy_row(easy_row):
    for i in range(3, 10):
        easy_row[i] = format_date_value(easy_row[i])
    return easy_row


# 填充 一级分类
def classify_easy_row_c1(easy_row):
    full_name = easy_row[2].lower()
    if "etf" in full_name:
        return "指数型"
    if "指数" in full_name:
        return "指数型"
    if "fof" in full_name:
        return "FOF"
    if "mom" in full_name:
        return "MOM"
    if "reits" in full_name:
        return "REITs"
    if "混合" in full_name:
        return "混合型"
    if "债券" in full_name:
        return "债券型"
    if "商品" in full_name:
        return "商品型"
    if "股票" in full_name:
        return "股票型"
    return "未知"


def classify_easy_row_c2(easy_row):
    full_name = easy_row[2].lower()
    if "交易型" in full_name and "开放式" in full_name and "指数" in full_name:
        if "联接" in full_name:
            return "ETF联接"
        return "ETF"
    return ""


# 填充 '是否为发起式' ,'是否为定期开放', '是否为港股通/香港主题'
def autofill_easy_row(easy_row):
    full_name = easy_row[2].lower()
    if not easy_row[13]:
        easy_row[13] = classify_easy_row_c1(easy_row)
    if not easy_row[14]:
        easy_row[14] = classify_easy_row_c2(easy_row)
    easy_row[15] = '否'
    easy_row[16] = '否'
    easy_row[17] = '否'
    if '发起式' in full_name:
        easy_row[15] = '是'
    if '定开' in full_name or '定期开放' in full_name:
        easy_row[16] = '是'
    if '港' in full_name and '港口' not in full_name:
        easy_row[17] = '是'
    return easy_row


# 使用万德数据填充： wd_data: ('管理人', '托管人', '基金全称', 发行公告日期', '认购起始日',' 认购截止日', '基金成立日', '发行总份额',
# 'Wind一级分类', 'Wind二级分类')
def fulfill_row_with_wd_data_easy(easy_row, wd_data):
    wd_row = find_wd_row(wd_data, easy_row[0], easy_row[1], easy_row[2])
    if wd_row is None and easy_row[19]:
        # try use old name
        wd_row = find_wd_row(wd_data, easy_row[0], easy_row[1], easy_row[19])
    if wd_row is None:
        # no wd data
        return easy_row
    print('%s match %s' % (wd_row, easy_row))
    if not easy_row[6]:
        easy_row[6] = wd_row[3]
    if not easy_row[7]:
        easy_row[7] = wd_row[4]
    if not easy_row[8]:
        easy_row[8] = wd_row[5]
    if not easy_row[9]:
        easy_row[9] = wd_row[6]
    if not easy_row[10]:
        easy_row[10] = wd_row[7]
    if not easy_row[11]:
        easy_row[11] = wd_row[8]
    if not easy_row[12]:
        easy_row[12] = wd_row[9]
    return easy_row


def complete_row(idx, db_row, wd_data, zjh_row):
    t0 = time.time()
    db_row[0:6] = zjh_row[0:6]
    db_row[18:20] = zjh_row[6:8]
    db_row = fulfill_row_with_wd_data_easy(db_row, wd_data)
    db_row = autofill_easy_row(db_row)
    db_row = format_easy_row(db_row)
    t1 = time.time()
    print("process {} elapsed: {}".format(idx, t1 - t0))
    return db_row


CPU_NUM = multiprocessing.cpu_count()
print("CPU NUM=%d" % CPU_NUM)


# 返回新的db
def fulfill_db_with_zjh_easy(db, zjh_rows, wd_data):
    zjh_dict = {}
    db_dict = {}
    for row in db:
        db_dict[row[2]] = row
    for row in zjh_rows:
        zjh_dict[row[2]] = row

    n_tasks = len(db) + len(zjh_rows)
    n_fin = 0

    new_easy_db = []
    pool = Pool(processes=CPU_NUM)
    with Manager() as ma:
        #         row_list = ma.list([])
        v_list = []
        for db_row in db:
            n_fin += 1
            print('%d/%d' % (n_fin, n_tasks))
            if db_row[2] not in zjh_dict:
                continue
            zjh_row = zjh_dict[db_row[2]]
            v = pool.apply_async(complete_row, args=(n_fin, db_row, wd_data, zjh_row))
            v_list.append(v)
        #         db_row[0:6] = zjh_row[0:6]
        #         db_row[18:20] = zjh_row[6:8]
        #         db_row = fulfill_row_with_wd_data_easy(db_row, wd_data)
        #         db_row = autofill_easy_row(db_row)
        #         db_row = format_easy_row(db_row)
        #             new_easy_db.append(db_row)
        pool.close()
        pool.join()
        new_easy_db += [v.get() for v in v_list]
    #         new_easy_db += row_list
    print("len of db: %d" % len(new_easy_db))

    pool = Pool(processes=CPU_NUM)
    with Manager() as ma:
        #         row_list = ma.list([])
        v_list = []
        for zjh_row in zjh_rows:
            n_fin += 1
            print('%d/%d' % (n_fin, n_tasks))
            if zjh_row[2] in db_dict:
                continue
            db_row = ['' for i in range(20)]
            #             print("process")
            #         db_row[0:6] = zjh_row[0:6]
            #         db_row[18:20] = zjh_row[6:8]
            #         db_row = fulfill_row_with_wd_data_easy(db_row, wd_data)
            #         db_row = autofill_easy_row(db_row)
            #         db_row = format_easy_row(db_row)
            #         new_easy_db.append(db_row)
            v = pool.apply_async(complete_row, args=(n_fin, db_row, wd_data, zjh_row))
            v_list.append(v)
        pool.close()
        pool.join()
        new_easy_db += [v.get() for v in v_list]
    #         new_easy_db += row_list
    print("len of db: %d" % len(new_easy_db))
    new_easy_db = sorted(new_easy_db, key=lambda x: x[4], reverse=True)
    return new_easy_db


# 简易程序: zjh_row -> (管理人，托管人，申请事项，申请日，受理日，决定日，是否为变更注册，变更注册代码)
def extract_row_from_zjh_easy_add(zjh_row):
    row = ['' for _ in range(8)]
    row[0:4] = zjh_row[1:5]
    row[4] = format_date_value(zjh_row[7])
    row[5] = zjh_row[13]
    row[6] = '否'
    return row


# 变更注册（简易）: zjh_row -> (管理人，托管人，申请事项，申请日，受理日，决定日，是否为变更注册，变更注册代码)
def extract_row_from_zjh_easy_change(zjh_row):
    row = ['' for _ in range(8)]
    row[0:2] = zjh_row[1:3]
    row[2] = zjh_row[4]
    row[3] = zjh_row[5]
    row[4] = format_date_value(zjh_row[8])
    row[5] = zjh_row[14]
    row[6] = '是'
    row[7] = zjh_row[3]
    return row


# 普通程序：zjh_row -> (管理人，托管人，申请事项，申请日，受理日，决定日，是否为变更注册，变更注册代码)
def extract_row_from_zjh_normal_add(zjh_row):
    row = ['' for _ in range(8)]
    row[0:4] = zjh_row[1:5]
    row[4] = format_date_value(zjh_row[7])
    row[5] = zjh_row[13]
    row[6] = '否'
    return row


# 变更注册（普通）：zjh_row -> (管理人，托管人，申请事项，申请日，受理日，决定日，是否为变更注册，变更注册代码)
def extract_row_from_zjh_normal_change(zjh_row):
    row = ['' for _ in range(8)]
    row[0:2] = zjh_row[1:3]
    row[2] = zjh_row[4]
    row[3] = zjh_row[5]
    row[4] = format_date_value(zjh_row[8])
    row[5] = zjh_row[14]
    row[6] = '是'
    row[7] = zjh_row[3]
    return row


# zjh row -> data row
def extract_rows_from_zjh_easy(zjh_filename):
    easy_rows = []
    easy_add_sheet = xlrd.open_workbook(zjh_filename).sheets()[0]
    for i in range(easy_add_sheet.nrows):
        zjh_row = easy_add_sheet.row_values(i)
        if not isinstance(zjh_row[0], float):
            continue
        easy_row = extract_row_from_zjh_easy_add(zjh_row)
        for j in range(len(easy_row)):
            if isinstance(easy_row[j], str):
                easy_row[j] = easy_row[j].strip()
        easy_rows.append(easy_row)
    easy_change_sheet = xlrd.open_workbook(zjh_filename).sheets()[2]
    for i in range(easy_change_sheet.nrows):
        zjh_row = easy_change_sheet.row_values(i)
        if not isinstance(zjh_row[0], float):
            continue
        easy_row = extract_row_from_zjh_easy_change(zjh_row)
        for j in range(len(easy_row)):
            if isinstance(easy_row[j], str):
                easy_row[j] = easy_row[j].strip()
        easy_rows.append(easy_row)

    # normal
    normal_rows = []
    normal_add_sheet = xlrd.open_workbook(zjh_filename).sheets()[1]
    for i in range(normal_add_sheet.nrows):
        zjh_row = normal_add_sheet.row_values(i)
        if not isinstance(zjh_row[0], float):
            continue
        print(zjh_row)
        row = extract_row_from_zjh_normal_add(zjh_row)
        for j in range(len(row)):
            if isinstance(row[j], str):
                row[j] = row[j].strip()
        normal_rows.append(row)

    normal_change_sheet = xlrd.open_workbook(zjh_filename).sheets()[3]
    for i in range(normal_change_sheet.nrows):
        zjh_row = normal_change_sheet.row_values(i)
        if not isinstance(zjh_row[0], float):
            continue
        row = extract_row_from_zjh_normal_change(zjh_row)
        for j in range(len(row)):
            if isinstance(row[j], str):
                row[j] = row[j].strip()
        normal_rows.append(row)

    return easy_rows, normal_rows


if __name__ == "__main__":

    t0 = time.time()
    result_xls = 'result.xls'
    easy_db, normal_db = load_db(result_xls)
    wd_data = load_wd_data('wd.xlsx')
    zjh_easy_rows, zjh_normal_rows = extract_rows_from_zjh_easy('zjh.xls')
    new_easy_db = fulfill_db_with_zjh_easy(easy_db, zjh_easy_rows, wd_data)
    new_normal_db = fulfill_db_with_zjh_easy(normal_db, zjh_normal_rows, wd_data)
    store_db(new_easy_db, new_normal_db, result_xls)
    t1 = time.time()
    print("all process elapsed: {}".format(t1-t0))

    # match_string("华宝中证1000指数证券投资基金", "华宝中证金融科技主题交易型开放式指数证券投资基金发起式联接基金")
