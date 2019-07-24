'''
对客户提供的表格，进行数据清理，拆分并输出json文件
input: ROOT/raw/*.xlsx
output: ROOT/src/data/*.json
'''
import pandas as pd
import os
import shutil
import json
import re

ROOT = os.getcwd()
RAW = 'raw'
RAW_PRODS = os.path.join(ROOT, RAW, '2019年产品手册汇总.xlsx')
RAW_SRVS = os.path.join(ROOT, RAW, '终端厂商服务网点收集.xlsx')
DIR_TEMP = os.path.join(ROOT, 'temp')
DIR_OP = os.path.join(ROOT, 'output')
DIR_TAR = os.path.join(ROOT, 'src', 'data')


def export_sheets_to_xlsx(ori, dst):
    df = pd.read_excel(ori, None)
    sheets = df.keys()
    if not os.path.exists(dst):
        os.mkdir(dst)
    i = 1
    for sheet in sheets:
        tempsheet = pd.read_excel(ori, sheet_name=sheet)
        index = "{:0>2}".format(str(i))
        new_sheet_name = index + sheet + ".xlsx"
        if ori == RAW_SRVS:
            new_sheet_name = "30服务网点" + index + sheet + ".xlsx"
        new_xlsx = os.path.join(ROOT, dst, new_sheet_name)
        writer = pd.ExcelWriter(new_xlsx)
        tempsheet.to_excel(writer, sheet_name=sheet, index=False)
        writer.close()
        i += 1


def split_by_area(src_folder, op_folder, ignore_list):
    file_list = os.listdir(src_folder)
    if not os.path.exists(op_folder):
        os.mkdir(op_folder)
    for file in file_list:
        full_file_name = os.path.join(src_folder, file)
        df = pd.read_excel(full_file_name, sheet_name=0, header=None)
        temp = df.dropna(how="all", axis=1)
        if temp.at[0, 1] in ignore_list:
            op_file = os.path.join(op_folder, file.split(
                '.')[0][:2] + "%" + file.split('.')[0][2:].strip() + ".xlsx")
            shutil.copyfile(full_file_name, op_file)
        else:
            for i in range(len(temp.columns)):
                if i % 3 == 0:
                    s1 = temp[temp.columns[i]]
                    s2 = temp[temp.columns[i + 1]]
                    new_data = pd.DataFrame(list(zip(s1, s2)))
                    name = os.path.join(op_folder, file.split(
                        '.')[0] + "-{:0>2}".format(str(i // 3 + 1)) + ".xlsx")
                    writer = pd.ExcelWriter(name)
                    new_data.to_excel(writer, index=False)
                    writer.close()
    shutil.rmtree(src_folder)


def deal_2char_ppl_name(ppl):
    ppl = ppl.replace(" ", "")
    if len(ppl) == 2:
        ppl = ppl[0] + (" " * 4) + ppl[1]
    return ppl


def deal_400_tel(tel):
    tel = tel.strip()
    if tel == "" or tel == "无":
        tel = "无"
    elif "-" not in tel:
        tel = "-".join([tel[:3], tel[3:6], tel[6:]])
    return tel


def deal_comp(comp):
    comp = str(comp)
    comp = comp.strip()
    if "）" == comp[-1]:
        comp = comp[:-6]
    return comp


def deal_common(sth):
    sth = str(sth)
    sth = sth.strip()
    sth = sth.replace('nan', '  ')
    sth = sth.replace('/4008008855', '')
    return sth


def deal_feature(sth):
    sth = sth.strip('\n').strip()
    pat = re.compile(r'\t')
    sth = pat.sub('', sth)
    pat = re.compile(r'(?<=[\n])( ?\d[、\.] *)|(?<=^)(\d[、\.] *)')
    sth = pat.sub('', sth)
    pat = re.compile(r'[；。]+\n')
    sth = pat.sub('；\n', sth)
    pat = re.compile('[；。]+$')
    sth = pat.sub('。', sth)
    pat = re.compile('(?<=[^；。])$')
    sth = pat.sub('。', sth)
    return sth.split('\n')


def deal_appl(sth):
    pat = re.compile('(?<![；。])(?=\\n)')
    sth = pat.sub('；', sth)
    pat1 = re.compile('(?<![。])$')
    sth = pat1.sub('。', sth)
    pat2 = re.compile(r'\n')
    sth = pat2.sub('', sth)
    return sth


def deal_k(sth):
    for i in range(len(sth)):
        pat = re.compile(r'尺英寸')
        sth[i] = pat.sub('尺寸', sth[i])
        pat1 = re.compile(r'型号\.\d')
        sth[i] = pat1.sub('型号', sth[i])
    return sth


def deal_v(sth):
    for i in range(len(sth)):
        pat = re.compile('\s+|\\\\n|\\\\uf06c ?')
        sth[i] = pat.sub(' ', sth[i])
    return sth


def get_brand(model):
    model = model.split("+")[0].strip().upper()
    db = {
        "QR300": "世麦",
        "I9000S": "优博讯",
        "I9100": "优博讯",
        "JYT6868A": "佳友通",
        "JYT6868W": "佳友通",
        "JYT9666D": "佳友通",
        "VQR800": "信雅达",
        "K9": "升腾",
        "Q50": "升腾",
        "C1": "华智融",
        "C2": "华智融",
        "NEW7210": "华智融",
        "NEW8210": "华智融",
        "NEW9220": "华智融",
        "CASH80AWI-33": "恒银",
        "X970": "惠尔丰",
        "Z300M": "惠尔丰",
        "CPOS X5": "新大陆",
        "ME50C": "新大陆",
        "ME50H": "新大陆",
        "ME62": "新大陆",
        "ME65": "新大陆",
        "ME68S": "新大陆",
        "N850": "新大陆",
        "N910": "新大陆",
        "N920": "新大陆",
        "SP50": "新大陆",
        "SP60": "新大陆",
        "SP600": "新大陆",
        "SKT-D500-P": "旭子",
        "SKT-T9005": "旭子",
        "SKT-T9006": "旭子",
        "A910": "百富",
        "A920": "百富",
        "A930": "百富",
        "E500": "百富",
        "E800": "百富",
        "QR10": "百富",
        "QR56": "百富",
        "QR65": "百富",
        "QR68": "百富",
        "S300": "百富",
        "S500": "百富",
        "S800": "百富",
        "S90": "百富",
        "LF-282": "神思朗方",
        "AECR C10": "联迪",
        "AECR C7": "联迪",
        "AECR C9": "联迪",
        "APOS A7": "联迪",
        "APOS A8": "联迪",
        "APOS A9": "联迪",
        "E350": "联迪",
        "E630": "联迪",
        "E820": "联迪",
        "E830": "联迪",
        "E850": "联迪",
        "QM10": "联迪",
        "QM30": "联迪",
        "QM50": "联迪",
        "QM500": "联迪",
        "QM800": "联迪",
        "Q160": "艾体威尔",
        "V60 SE": "艾体威尔",
        "V80 SE": "艾体威尔",
        "MF69": "魔方"
    }
    return db[model]


def export_prod_json(xlsx, dst):
    op_dict = {}
    if "%" not in xlsx:
        df = pd.read_excel(xlsx, sheet_name=0)
        model = df.iloc[0, 1]
        brand = get_brand(model)
        pic_src = "{}{}.png".format(brand, model.split("+")[0].strip().upper())
        df.iloc[1, 1] = pic_src
        feature = df.iloc[2, :][1]
        appl = df.iloc[3, :][1]
        df.dropna(how="any", inplace=True)
        k = list(df.iloc[4:, 0])
        v = list(df.iloc[4:, 1])
        op_dict["model"] = model
        op_dict["brand"] = brand
        op_dict["title"] = op_dict["brand"] + op_dict["model"]
        op_dict["series"] = os.path.basename(xlsx)[2:-8]
        op_dict["pic_src"] = pic_src
        op_dict["feature"] = deal_feature(feature)
        op_dict["appl"] = deal_appl(appl)
        op_dict["k"] = deal_k(k)
        op_dict["v"] = deal_v(v)
    else:
        op_dict["series"] = os.path.basename(xlsx)[3:-5]
    json_str = json.dumps(op_dict, ensure_ascii=False, indent=2)
    with open(dst, "w", encoding='utf-8') as fp:
        fp.write(json_str)
    fp.close()


def export_srv_json(xlsx, dst):
    df = pd.read_excel(xlsx, sheet_name=0, header=None)
    company = deal_comp(df.iloc[0, 0])
    tel = deal_400_tel(df.iloc[1, 0][9:])
    srv_area = list(df.iloc[3:, 0].apply(deal_common))
    srv_ppl = list(df.iloc[3:, 1].apply(deal_2char_ppl_name))
    srv_tel = list(df.iloc[3:, 2].apply(deal_common))
    srv_addr = list(df.iloc[3:, 3].apply(deal_common))
    op_dict = {}
    op_dict["company"] = company
    # op_dict["c_tel"] = tel
    op_dict["tel"] = "全国售后服务热线： " + tel
    op_dict["srv_area"] = srv_area
    op_dict["srv_ppl"] = srv_ppl
    op_dict["srv_tel"] = srv_tel
    op_dict["srv_addr"] = srv_addr
    json_str = json.dumps(op_dict, ensure_ascii=False, indent=2)
    with open(dst, "w", encoding='utf-8') as fp:
        fp.write(json_str)
    fp.close()


def walk_xlsx(src, dst):
    if not os.path.exists(dst):
        os.mkdir(dst)
    f_list = os.listdir(src)
    for file in f_list:
        if file[-4:] == "xlsx":
            if int(file[:2]) >= 30:
                fx = export_srv_json
            else:
                fx = export_prod_json
            dst_file = os.path.join(dst, file[:-5] + ".json")
            fx(os.path.join(src, file), dst_file)

# 产品xlsx文件工作簿分割为单工作表的文件
export_sheets_to_xlsx(RAW_PRODS, DIR_TEMP)

# 产品单工作表文件按列分割为每个产品一个文件
# 忽略不需要的列（空列）
ignore_list = ["Unnamed: 1", "三合一键盘"]
split_by_area(DIR_TEMP, DIR_OP, ignore_list)

# 售后网点xlsx文件工作簿分割为单工作表的文件
export_sheets_to_xlsx(RAW_SRVS, DIR_OP)

# 所有单工作表文件进行规范化处理，输出为json文件
walk_xlsx(DIR_OP, DIR_TAR)
