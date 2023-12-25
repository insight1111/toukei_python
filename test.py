# %%
import pprint
import sqlite3

import openpyxl
import yaml

pp = pprint.PrettyPrinter(indent=4)
# %%
with open("toukei_formatu.ini", encoding="utf-8") as f:
    config = yaml.safe_load(f)

# %% データベース読み込み
with open("database.ini") as f:
    dbname = yaml.safe_load(f)
conn = sqlite3.connect(dbname)

# %% データベーステスト
cur = conn.cursor()
cur.execute("select * from data where nendo=2023 and month=11 and cont='合計'")
pp.pprint(cur.fetchall())

# %% Excelを読み込み
with open("excel.ini", encoding="utf-8") as f:
    excelname = yaml.safe_load(f)
wb = openpyxl.load_workbook(excelname)

# %%
cur.close()
conn.close()

# %%
