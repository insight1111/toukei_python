# %%
import pprint
import sqlite3

import openpyxl
import yaml
from retry import retry
from logging import getLogger

logger = getLogger(__name__)
pp = pprint.PrettyPrinter(indent=4)

# %% 各種設定ファイル読み込み
try:
    with open("toukei_formatu.ini", encoding="utf-8") as f:
        config = yaml.safe_load(f)
except Exception as e:
    print(e)
    print("toukei_formatu.iniファイルがみつかりません")
    exit()

try:
    with open("database.ini") as f:
        dbname = yaml.safe_load(f)
except Exception as e:
    print(e)
    print("database.iniファイルが見つかりません")
    exit()

conn = sqlite3.connect(dbname)

try:
    with open("excel.ini", encoding="utf-8") as f:
        excelname = yaml.safe_load(f)
except Exception as e:
    print(e)
    print("excel.iniファイルが見つかりません")
    exit()

wb = openpyxl.load_workbook(excelname)


# %%
@retry(logger=logger)
def get_year_and_month():
    year_and_month = input("年月を入力してください(例：2008年1月->200801)")
    if len(year_and_month) != 6:
        raise ValueError("入力値は数字6桁で入力してください")
    year = int(year_and_month[0:4])
    month = int(year_and_month[4:6])
    nendo = year
    if month >= 1 and month <= 3:
        nendo -= 1
    if month < 1 or month > 12:
        raise ValueError("入力数値が不正です")
    return (year, month, nendo)


(year, month, nendo) = get_year_and_month()
print(year, month, nendo)
# %% データベーステスト
cur = conn.cursor()
cur.execute("select * from data where nendo=2023 and month=11 and cont='合計'")
pp.pprint(cur.fetchall())

# %%
cur.close()
conn.close()

# %%
