# coding: utf-8
import configparser
import os
import time
import traceback

import xlwings as xw

cf = configparser.ConfigParser()
app = xw.App(visible=False, add_book=False)


def delete_exist():
    cf.read("conf.conf", "utf-8")
    change_file = cf.get("dir", "change_file_dir")
    data_dir = cf.get("dir", "data_dir")
    l = []
    for file in os.listdir(data_dir):
        _dir = data_dir + "/" + file
        print("处理文件：", _dir)
        w = app.books.open(_dir)
        try:
            sht = w.sheets.active
            a = sht.range('A1').expand().value
            if not isinstance(a, list):
                if len(str(a)) > 40:
                    l.append(a)
                continue
            for v in a:
                if not isinstance(v, str):
                    continue
                if len(v) > 40:
                    l.append(v)
            # exit()
        except:
            traceback.print_exc()
        finally:
            w.save()
            w.close()
    print(l)
    wb = app.books.open(change_file)
    try:
        sheet = wb.sheets.active
        origin = sheet.range('A1').expand().value
        print('origin:', origin)
        for i in l:
            if isinstance(origin, list):
                if i in origin:
                    origin.remove(i)
            else:
                if i == origin:
                    origin = ''
        print(origin)
        sheet_name = time.strftime('%Y-%m-%d %H\'%M\"%S', time.localtime(time.time()))
        wb.sheets.add(sheet_name)
        sheet2 = wb.sheets[sheet_name]
        sheet2.range('A1').options(transpose=True).value = origin
    except:
        traceback.print_exc()
    finally:
        wb.save()
        wb.close()
        exit()
        app.quit()


if __name__ == '__main__':
    delete_exist()
