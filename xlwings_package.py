import xlwings as xw
import pandas as pd



def base_import():

    import sys
    # the mock-0.3.1 dir contains testcase.py, testutils.py & mock.py
    sys.path.append('''C:/Users/james/Dropbox/Python/xlwings_helper''')
    import xlwings_package as xwp


def get_ws(book, sheet = 'Sheet1'):

    wb = xw.Book(book)
    ws = wb.sheets[sheet]
    return ws

def full_range(ws):

    rng = ws.range('A1').expand()
    return rng

def get_rows(ws, top_left = (1,1), bottom_right = None):

    if bottom_right == None:
        bottom_right = (full_range(ws).last_cell.row, full_range(ws).last_cell.column)
    print (bottom_right)
    rows = ws.range( top_left, bottom_right ).options(ndim = 2).value
    return rows

def df_from_rows(rows):

    df = pd.DataFrame(rows[1:], columns = rows[0])
    return df

def write_2d(ws, rows, top_left = (1,1)):

    ws.range(top_left).expand().value = rows

def row_to_col(row):

    if len(row) == 1:
        return row

    col = []
    for i in row:
        col.append( [i])
    return col

def write_df(ws, df):

    col = []

    for i in range(len(df.columns)):
        column = [df.columns[i]] + df[ df.columns[i] ].tolist()
        col = row_to_col(column)
        write_2d(ws, col, top_left = (1, (i+1)))

def change_cell_color(ws, cell, color):

    ws.range(cell).color = color
