import os
import re
from itertools import groupby

import pyexcel as pe


def sorted_nicely(l):
    convert = lambda text: int(text) if text.isdigit() else text
    alphanum_key = lambda key: [convert(c) for c in re.split('([0-9]+)', str(key[1]))]
    return sorted(l, key=alphanum_key)


def main():
    excel_rows_groupby = []
    excel_rows_append = []
    electricity = float(input("Hesaplama için çarpan giriniz"))
    block_number = input("Blok numarasını giriniz")
    date_range = ""
    for file in os.listdir('ExcelFiles'):
        excel_rows = pe.get_array(file_name="ExcelFiles" + "/" + file, start_row=2)
        date_range = excel_rows[0][0]
        excel_rows.pop(0)
        excel_rows.pop(0)
        excel_rows.pop(0)
        excel_rows.pop()
        excel_rows_append.extend(excel_rows)
        # for row in excel_rows_groupby:
        #     print(f"{row[0]} - {row[1]} - {row[2]}")

    for i, g in groupby(sorted_nicely(excel_rows_append), key=lambda x: x[1]):
        total = sum(float(str(v[6]).replace(',', '.')) for v in g)
        excel_rows_groupby.append([date_range, i, total, str(total * electricity) + "TL"])
    excel_rows = sorted_nicely(excel_rows_groupby)
    excel_rows.insert(0, ["Tarih aralığı", "Daire Numarası", "Harcama", "Hesaplanan harcama"])
    excel_rows_filtered = filter(lambda excel: str(excel[1]).startswith(block_number), excel_rows)

    pe.save_as(array=excel_rows_filtered,
               dest_file_name=block_number + "_" + date_range + "_" + "pay_olcer_hesaplama.xls")

    pe.save_as(array=excel_rows, dest_file_name="pay_olcer_hesaplama.xls")


if __name__ == '__main__':
    main()
