import xlwings as xw
import pandas as pd

import sys
#this is a test
try:
    sys.path.append('''C:/Users/james/Dropbox/Python/Analytics''')
except:
    pass

try:
    sys.path.append('''C:/Users/e150445/Documents/Analytics''')
except:
    pass

import CustomPandas as cpd

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

def get_rows(ws, top_left_cell = (1,1), bottom_right = None):

    if bottom_right == None:
        bottom_right = (full_range(ws).last_cell.row, full_range(ws).last_cell.column)
    rows = ws.range( top_left_cell, bottom_right ).options(ndim = 2).value
    return rows

def df_from_rows(rows):

    df = pd.DataFrame(rows[1:], columns = rows[0])
    return df

def write_2d(ws, rows, top_left_cell = (1,1)):

    ws.range(top_left_cell).expand().value = rows

def row_to_col(row):

    if len(row) == 1:
        return row

    col = []
    for i in row:
        col.append( [i])
    return col

def keep_these_rows(df, locs):

    df = df.loc(locs)
    return df

def get_df_from_ws(ws):

    df = df_from_rows( get_rows(ws))
    return df



def write_df_col_to_ws(ws, df, col_index, col_name):

    '''writes a df column to a certain column number in the ws'''

    values = df[col_name]
    try:
        values = values.tolist()
    except:
        values = []
    values.insert(0, col_name)
    #print (values)
    values_col = row_to_col(values)
    #print (values_col)

    write_2d(ws, values_col, top_left_cell = (1, col_index + 1))



def alpha_from_index(integer):

    '''Takes a 0-based index (integer) and returns the corresponding column header'''

    lengths = [1,2,3,4,5,6,7]
    contained_in_lengths = []
    for i in lengths:
        contained_in_lengths.append(26 ** i)

    integer += 1

    integer_copy = integer
    num_digits = lengths[-1]
    for i in range(len(contained_in_lengths)):
        integer_copy -= contained_in_lengths[i]

        if integer_copy <= 0:
            num_digits = i + 1
            break

    digits = ['',] * num_digits
    alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    breakdown = [0,] * num_digits
    digits = [0,] * num_digits
    for i in range(num_digits):

        breakdown[i] = integer % (contained_in_lengths[i])
        if breakdown[i] == 0:
            breakdown[i] = contained_in_lengths[i]
        integer -= breakdown[i]
        digits[i] = breakdown[i] / (26**i)

    string = ''
    digits.reverse()
    for i in digits:
        string += alpha[int(i) - 1]
    return string

def alphas_from_index_list(ints):

    final = []
    for i in ints:
        final.append(alpha_from_index(i))
    return final

def column_index_from_alphas(string):

    list = []
    for i in string:
        list.append(i.upper())

    alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    multiplier = 0
    final = 0
    for i in range(len(list)):

        new_mult = 26**i
        #1, 26, 676

        place = (i+1) * -1
        index = alpha.index( list[place] )
        index += 1
        final += (index * new_mult)

    return (final - 1)

def get_column_headers_from_alpha(df, list_of_alphas):

    inds = []
    for i in list_of_alphas:
        inds.append(column_index_from_alphas(i))

    df_headers = []
    cols = df.columns.tolist()
    #print (cols)
    #print (len(cols))

    for i in range(len(list_of_alphas)):
        df_headers.append(cols[ inds[i] ])
        #print (inds[i])

    return df_headers

def sort_ws(ws, column_alphas):

    '''takes active ws and list of column alphas and sorts worksheet'''

    df = df_from_rows( get_rows(ws) )
    headers = get_column_headers_from_alpha(df, column_alphas)
    df = cpd.sort_df(df, headers)
    #print (headers)
    #print (df)
    write_df_to_ws(ws, df)

def check_column_in_list(df, column, list, new_column):
    '''returns a dataframe with a boolean value in new column if the row had one of those value or not'''

    df[new_column] = df[column].isin(list)
    return df

def add_sheet(sheet_name, work_book):

    '''adds sheet to workbook'''
    try:
        work_book.sheets.add(sheet_name)
    except:
        pass

def combine_string_columns(df, col1, col2, new_column):

    '''Returns df with new column that has a compiled string of col1 and col2'''

    df[new_column] = df[col1].map(str) + df[col2].map(str)
    return df

def get_wb(book_name):

    return xw.Book(book_name)

def write_df_to_ws(ws, df):


    header = df.columns.tolist()
    write_2d(ws, header)

    for col_num in range(len(df.columns)):
        col = header[col_num]
        write_df_col_to_ws(ws, df, col_num, col)


def alpha_from_column_names(df, strings):

    cols = df.columns.tolist()

    alphas = []
    for i in strings:
        a = cols.index(i)
        alphas.append( alpha_from_index(a) )

    return alphas

def change_cell_color(ws, top_left_cell, cell_color, bottom_right_cell = None):

    '''changes a range of cells a certain color'''
    if bottom_right_cell == None:
        bottom_right_cell = top_left_cell
    ws.range(top_left_cell, bottom_right_cell).color = cell_color
