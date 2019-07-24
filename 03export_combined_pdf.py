'''
生成目录ai文件
按顺序添加页码
根据页码切换左右页布局
组合成册
input: ROOT/dist/*.ai, ROOT/src/data/, ROOT/src/tmpl/tmpl_toc.ait
output: ROOT/output.pdf
'''
import os
import win32com
from illustrator import AI
from PyPDF2 import PdfFileMerger

ROOT = os.getcwd()
DIR_SRC = os.path.join(ROOT, "dist")
DIR_DATA = os.path.join(ROOT, "src", "data")
TMPL_TOC = os.path.join(ROOT, "src", "tmpl", "tmpl_toc.ait")
TOC = os.path.join(ROOT, "src", "目录.ai")
COVER = os.path.join(ROOT, "src", "封面.pdf")
BACKCOVER = os.path.join(ROOT, "src", "封底.pdf")
FILE_OUTPUT = os.path.join(ROOT, "output.pdf")

a = AI()
opts = win32com.client.gencache.EnsureDispatch(
    "Illustrator.IllustratorSaveOptions")
a.open(TMPL_TOC)
series = {}

for i, j in enumerate(os.listdir(DIR_DATA)):
    n = j[:-5]
    if r"%" in n:
        s = n[3:]
    elif r"服务网点" in n:
        s = r"服务网点"
    else:
        s = n[2:-3]
    if s not in series.keys():
        series[s] = ("{:0>2}".format(i + 1), j[:2])


def fill_toc(l, cat, subcat, gt, lt):
    tmp_txt = ""
    txt_cat = a.get_item_by_name_and_layer_name("txt_cat", l)
    txt_cat.Contents = cat
    txt_subcat = a.get_item_by_name_and_layer_name("txt_subcat", l)
    txt_subcat.Contents = subcat
    txt_series = a.get_item_by_name_and_layer_name("series", l)
    for i, j in series.items():
        if gt < int(j[1]) < lt:
            tmp_txt += i + "." * (40 - len(i)) * 3 + str(j[0]) + "\n"
    txt_series.Contents = tmp_txt.strip()
    tmp_txt = ""

fill_toc("t1", "商用支付终端⸺智能终端", "智能POS终端", 0, 3)
fill_toc("t2", "商用支付终端⸺传统终端", "POS终端", 2, 8)
fill_toc("t3", "自助终端", "助农金融自助终端", 7, 9)
fill_toc("t3s", "自助终端", "泛非接模块", 8, 10)
fill_toc("t4", "服务网点", "", 29, 31)
a.app.ActiveDocument.SaveAs(TOC, opts)
a.close()

pdf_output = PdfFileMerger()
pdf_output.append(COVER)
pdf_output.append(TOC)
for (num, file) in enumerate(os.listdir(DIR_SRC)):
    file = os.path.join(DIR_SRC, file)
    a.open(file)
    a.update_pg_num_for_single_page(num, "pg_num")
    a.determine_layout(num)
    a.save()
    a.close()
    pdf_output.append(file)
pdf_output.append(BACKCOVER)
pdf_output.write(open(FILE_OUTPUT, 'wb'))
pdf_output.close()
