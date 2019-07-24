'''
根据json数据文件打开对应的模板
将json数据写入模板
保存并存为ai文件
dependency: ./illustrator.py
input: ROOT/src/data/*.json, ROOT/src/tmpl/*.ait
output: ROOT/dist/*.ai
'''
import os
import re
import win32com
from illustrator import AI

a = AI()
opts = win32com.client.gencache.EnsureDispatch(
    "Illustrator.IllustratorSaveOptions")
ds = {}

for f in os.listdir(a.DIR_DATA):
    n = f[:-5]  # 去除扩展名(.json)
    ds[n] = a.import_data(f)  # 加载数据

    # 选择模板(tmpl)，填入数据
    if int(n[:2]) == 5:
        tmpl = "tmpl_blank"
        a.fill_data(ds[n]['ds'], tmpl, n)

    elif int(n[:2]) in range(3, 5):
        tmpl = "tmpl_prods_multi"
        a.fill_data(ds[n]['ds'], tmpl, n)

    elif int(n[:2]) >= 30:
        tmpl = "tmpl_srvs"
        a.fill_data(ds[n]['ds'], tmpl, n)

    else:
        tmpl = "tmpl_prods"
        a.fill_data(ds[n]['ds'], tmpl, n)
        # 对该模板做特殊处理：根据feature字段数据的行数，
        # 删除名为 "d"后一位数字 的GroupItem中的多余圆点
        pat = re.compile(r"d\d")
        layer_info = a.get_layer_by_name("info")
        to_del = []
        for g in layer_info.GroupItems:
            if pat.match(g.Name):
                if int(g.Name[1]) > ds[n]['lcount']['feature']-1:
                    to_del.append(g.Name)
        for i in to_del:
            g = a.get_item_by_name(i)
            g.Delete()

    dest = os.path.join(a.DIR_ROOT, a.DIR_DST, n + ".ai")
    # print(n, "-->", tmpl)
    a.app.ActiveDocument.SaveAs(dest, opts)
    a.close()
    print(dest, "DONE!")
