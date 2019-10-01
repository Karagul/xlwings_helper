
import sys
# the mock-0.3.1 dir contains testcase.py, testutils.py & mock.py
sys.path.append('''C:/Users/james/Dropbox/Python/xlwings_helper''')
import xlwings_package as xwp


sheet = xwp.get_ws('test.xlsx')
rows = xwp.get_rows(sheet)

df = xwp.df_from_rows(rows)

print (df)

df['e'] = 1

xwp.write_df_to_ws(sheet, df)
