# coding: utf-8
import configparser
import os

import xlwings as xw

cf = configparser.ConfigParser()
app = xw.App(visible=False, add_book=False)


def delete_exist():
    cf.read("conf.conf", "utf-8")
    change_file = cf.get("dir", "change_file_dir")
    data_dir = cf.get("dir", "data_dir")
    l = []
    wb = app.books.open(change_file)
    sheet = wb.sheets.active
    origin = sheet.range('A1').expand().value
    for file in os.listdir(data_dir):
        dir = data_dir + "/" + file
        print("处理文件：", dir)
        wb = app.books.open(dir)
        sheet = wb.sheets.active
        a = sheet.range('A1').expand().value
        if not isinstance(a, list):
            continue
        for v in a:
            if not isinstance(v, str):
                continue
            if len(v) > 40:
                l.append(v)
        # exit()
        wb.save()
        wb.close()


    print(l)
    exit()
    app.quit()


if __name__ == '__main__':
    delete_exist()
