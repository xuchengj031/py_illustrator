import os
import json
from win32com.client import constants
import win32com


class AI():

    def __init__(self, dir=os.getcwd()):
        self.app = win32com.client.gencache.EnsureDispatch(
            "Illustrator.Application")
        self.DIR_ROOT = dir
        self.DIR_DATA = os.path.join(self.DIR_ROOT, "src", "data")
        self.DIR_TMPL = os.path.join(self.DIR_ROOT, "src", "tmpl")
        self.DIR_IMG = os.path.join(self.DIR_ROOT, "src", "images")
        self.DIR_DST = os.path.join(self.DIR_ROOT, "dist")
        self.LOG_ERR = os.path.join(self.DIR_ROOT, "err.log")
        self.PLACEHOLD = os.path.join(self.DIR_TMPL, "placehold.png")
        self.OLD_STR = bytes(
            'D:\\projects\\py_catalog\\src\\tmpl\\placehold.png'.encode('ISO-8859-1'))
        self.NEW_STR = bytes(self.PLACEHOLD.encode('ISO-8859-1'))
        if not os.path.exists(self.DIR_DST):
            os.mkdir(self.DIR_DST)
        for tmpl in os.listdir(self.DIR_TMPL):
            if tmpl[-4:] == ".ait":
                tmpl_path = os.path.join(self.DIR_TMPL, tmpl)
                with open(tmpl_path, 'rb') as f:
                    s = f.read()
                s = s.replace(self.OLD_STR, self.NEW_STR)
                with open(tmpl_path, 'wb') as f:
                    f.write(s)

    def open(self, filename):
        self.app.Open(filename)

    def save(self):
        self.app.Application.ActiveDocument.Save()

    def close(self):
        self.app.Application.ActiveDocument.Close(constants.aiDoNotSaveChanges)

    def close_all(self):
        while self.app.Application.Documents.Count > 0:
            self.app.Application.Documents.Item(
                1).Close(constants.aiDoNotSaveChanges)

    def get_layer_by_name(self, layer_name):
        for l in self.app.ActiveDocument.Layers:
            if l.Name == layer_name:
                return l

    def get_item_by_name(self, pageitem_name):
        for i in self.app.ActiveDocument.PageItems:
            if i.Name == pageitem_name:
                return i

    def get_item_by_name_and_layer_name(self, pageitem_name, layer_name):
        for l in self.app.ActiveDocument.Layers:
            if l.Name == layer_name:
                for i in l.PageItems:
                    if i.Name == pageitem_name:
                        return i

    def select_all(self, unlock=0, unhide=0):
        for l in self.app.ActiveDocument.Layers:
            if unlock:
                self.unlock_all()
            if unhide:
                self.unhide_all()
            self.select_all_in_layer(l)

    def select_all_in_layer(self, layer):
        prop_map = layer._prop_map_get_
        for prop_name in prop_map:
            prop = getattr(layer, prop_name)
            if hasattr(prop, 'Count'):
                for i in prop:
                    i.Selected = True

    def select_all_txt(self, unlock=False, unhide=False):
        for l in self.app.ActiveDocument.Layers:
            if unlock:
                self.unlock_all(10)
            if unhide:
                self.unhide_all(10)
            for tf in l.TextFrames:
                tf.Selected = True

    def unlock_all(self, itemtype=0):
        hiddens = self.unhide_all()
        lockeds = [[], []]
        for l in self.app.ActiveDocument.Layers:
            if l.Locked:
                lockeds[0].append(l)
                l.Locked = False
        for i in self.app.ActiveDocument.PageItems:
            if not itemtype:
                if i.Locked:
                    lockeds[0].append(i)
                    i.Locked = False
            else:
                if i.PageItemType == itemtype and i.Locked:
                    lockeds[1].append(i)
                    i.Locked = False
        self.restore_hidden_state(hiddens)
        return lockeds

    def restore_locked_state(self, lockeds):
        hiddens = self.unhide_all()
        for l in lockeds[0]:
            l.Locked = True
        for i in lockeds[1]:
            i.Locked = True
        self.restore_hidden_state(hiddens)

    def unhide_all(self, itemtype=0):
        hiddens = [[], []]
        for l in self.app.ActiveDocument.Layers:
            if not l.Visible:
                hiddens[0].append(l)
                l.Visible = True
        for i in self.app.ActiveDocument.PageItems:
            if not itemtype:
                if i.Hidden:
                    hiddens[1].append(i)
                    i.Hidden = False
            else:
                if i.PageItemType == itemtype and i.Hidden:
                    hiddens[1].append(i)
                    i.Hidden = False
        return hiddens

    def restore_hidden_state(self, hiddens):
        if not (type(hiddens) == list and
                type(hiddens[0]) == list and
                type(hiddens[1]) == list):
            print("Must be a 2d list")
            return
        for l in hiddens[0]:
            l.Visible = False
        for i in hiddens[1]:
            i.Hidden = True

    def fill_data(self, data, tmpl, target, ignore_field=[]):
        if not os.path.isabs(tmpl):
            tmpl = os.path.join(self.DIR_TMPL, tmpl + ".ait")
        self.open(tmpl)
        for item in self.app.ActiveDocument.PageItems:
            if item.Name in data:
                if item.PageItemType == constants.aiPlacedItem:
                    if os.path.isabs(data[item.Name]):
                        data[item.Name] = os.path.split(data[item.Name])[1]
                    file_path = os.path.join(self.DIR_IMG, data[item.Name])
                    item.File = file_path
                if item.PageItemType == constants.aiTextFrame:
                    item.Contents = str(data[item.Name])

    def import_data(self, file, ignore_field=[]):
        data = {}
        if not os.path.isabs(file):
            file = os.path.join(self.DIR_DATA, file)
        with open(file, "r", encoding="utf-8") as fp:
            ds = data['ds'] = json.load(fp)
        fp.close()
        data['lcount'] = {}
        for k in ds.items():
            if k[0] in ignore_field:
                pass
            elif type(k[1]) == list:
                data['lcount'][k[0]] = len(k[1])
                if type(k[1][0]) == str:
                    ds[k[0]] = "\n".join(k[1])
                else:
                    self.add_log(
                        "\n{} : \'{}\' not a string\n".format(file, k[0]))
        return data

    def iter_folder_ai(self, path, func, ignore, args=None):
        files = os.listdir(path)
        for f in files:
            if f[-3:] != ".ai" or f in ignore:
                continue
            self.open(os.path.join(path, f))
            func(self, args)
            self.save()
            self.close()

    def update_pg_num_for_single_page(self, num, target_field):
        pg_num = str(num + 1) if num > 8 else "0" + str(num + 1)
        for i in self.app.ActiveDocument.PageItems:
            if i.PageItemType == constants.aiTextFrame and i.Name == target_field:
                i.Contents = pg_num
        self.save()

    def determine_layout(self, num):
        # 显示所有层，但隐藏其下的项目
        flag = bool(num % 2)
        for l in self.app.ActiveDocument.Layers:
            l.Visible = True
            l.Locked = False
            if l.Name.endswith("_r"):
                for i in l.PageItems:
                    i.Hidden = (not flag)
            elif l.Name.endswith("_l"):
                for i in l.PageItems:
                    i.Hidden = flag

        # # 只隐藏层
        # flag = bool(num % 2)
        # for l in self.app.ActiveDocument.Layers:
        #     if l.Name.endswith("_r"):
        #         l.Visible = flag
        #     if l.Name.endswith("_l"):
        #         l.Visible = flag

        # # 隐藏层和项目
        # isRight = False if num % 2 == 0 else True
        #     if l.Name.endswith("_r"):
        #         l.Visible = True
        #         for i in l.PageItems:
        #             if i.Hidden == isRight:
        #                 i.Hidden = not isRight
        #     if l.Name.endswith("_l"):
        #         l.Visible = True
        #         for i in l.PageItems:
        #             if i.Hidden == not isRight:
        #                 i.Hidden = isRight
        self.save()

    def add_log(self, e):
        with open(self.LOG_ERR, "a+", encoding="utf-8") as fp:
            fp.write(e)
