# -*- coding: utf-8 -*-
"""
Created on Wed Mar 21 16:53:28 2018

@author: omer
"""

# -*- coding: utf-8 -*-
import dash
import dash_core_components as dcc
import dash_html_components as html

import xlrd
import pandas as pd
import numpy as np
import datetime

def get_header_names(sheet, row_id_headers):
    all_headers = []
    for column in range(sheet.ncols):
        all_headers.append(sheet.cell(row_id_headers,column).value)
    return all_headers

def get_rounded_dates(selected_dates_slice, date_mode):
    selected_dates = convert_dates_col_slice_to_datetime(selected_dates_slice, date_mode)
    df = pd.DataFrame({'selected_dates': selected_dates})
    rounded_dates = pd.DatetimeIndex(df['selected_dates']).normalize()
    return rounded_dates

def get_selected_dates_slice(data_sheet, dates_header_name, row_id_headers, all_headers):
    selected_date_col_id = find_first_header_id(all_headers, dates_header_name)
    selected_dates_slice = data_sheet.col_slice(selected_date_col_id, start_rowx=row_id_headers+1)
    return selected_dates_slice

def find_header_id(lst, val):
    return  [i for i,x in enumerate(lst) if x==val]

def find_first_header_id(lst, val):
    return  find_header_id(lst, val)[0]

def convert_dates_col_slice_to_datetime(dates_slice, date_mode):
    return [xlrd.xldate.xldate_as_datetime(date_cell.value, date_mode) for date_cell in dates_slice]

def filter_status(status_slice, status):
    return  [i for i,x in enumerate(status_slice) if x.value==status]

#for i in range(len(dates_slice)):
#    print(i)
#    print(xlrd.xldate.xldate_as_datetime(dates_slice[i].value, date_mode))

# Assign spreadsheet filename to `file`
filename = '../Upload_Sheet_12_Mar_V2.xlsx'

# Open a workbook 
workbook = xlrd.open_workbook(filename)
data_sheet = workbook.sheet_by_name('DATA')

# col headers
row_id_headers = 1
all_headers = get_header_names(data_sheet, row_id_headers)

status_col_id = find_first_header_id(all_headers, 'Status')
status_slice = data_sheet.col_slice(status_col_id, start_rowx=row_id_headers+1)

closed_status_id = filter_status(status_slice, 'Closed')

created_dates_slice = get_selected_dates_slice(data_sheet, 'Created Date', row_id_headers, all_headers)
resolved_dates_slice = get_selected_dates_slice(data_sheet, 'Resolved Date', row_id_headers, all_headers)

resolved_dates_slice_filtered = [resolved_dates_slice[idx] for idx in closed_status_id]
#created_date_col_id = find_first_header_id(all_headers, 'Created Date')
#created_dates_slice = data_sheet.col_slice(created_date_col_id, start_rowx=row_id_headers+1)
#
#created_dates = convert_dates_col_slice_to_datetime(created_dates_slice, workbook.datemode)
#
#df = pd.DataFrame({'created_dates': created_dates})
#
#rounded_dates = pd.DatetimeIndex(df['created_dates']).normalize()
#dates_bins = pd.date_range(rounded_dates[0],rounded_dates[-1]+pd.DateOffset(1))
#
#created_date_histogram = np.histogram(rounded_dates.values.astype('int64'),dates_bins.values.astype('int64'))
    
created_dates_rounded = get_rounded_dates(created_dates_slice, workbook.datemode)
resolved_dates_rounded = get_rounded_dates(resolved_dates_slice_filtered, workbook.datemode)

firstDate = min([created_dates_rounded[0], resolved_dates_rounded[0]])
lastDate = max([created_dates_rounded[-1], resolved_dates_rounded[-1]])


dates_bins = pd.date_range(firstDate, lastDate+pd.DateOffset(1))
created_date_histogram = np.histogram(created_dates_rounded.values.astype('int64'), dates_bins.values.astype('int64'))
resolved_date_histogram = np.histogram(resolved_dates_rounded.values.astype('int64'), dates_bins.values.astype('int64'))

unresolved_queries = np.cumsum(created_date_histogram[0] - resolved_date_histogram[0])

app = dash.Dash()

app.layout = html.Div(children=[
    html.H1(children='Analytics Dashboard'),

    html.Div(children='''
        based on "Dash: A web application framework for Python".
    '''),

    dcc.Graph(
        id='example-graph',
        figure={
            'data': [
                {'x': dates_bins[:-1], 'y': created_date_histogram[0], 'type': 'bar', 'name': 'Created dates'},
                {'x': dates_bins[:-1], 'y': resolved_date_histogram[0], 'type': 'bar', 'name': 'Resolved dates'},
                {'x': dates_bins[:-1], 'y': unresolved_queries[0], 'type': 'line', 'name': 'Unresolved dates'},
                #{'x': [1, 2, 3], 'y': [2, 4, 5], 'type': 'bar', 'name': u'Montréal'},
            ],
            'layout': {
                'title': 'Created and resolved dates and unresolved queries'
            }
        }
    )#,
    
#        dcc.Graph(
#        id='example-graph_1',
#        figure={
#            'data': [
#                {'x': dates_bins[:-1], 'y': resolved_date_histogram[0], 'type': 'bar', 'name': 'Resolved dates'},
#                #{'x': [1, 2, 3], 'y': [2, 4, 5], 'type': 'bar', 'name': u'Montréal'},
#            ],
#            'layout': {
#                'title': 'Resolved dates'
#            }
#        }
#    )
])

if __name__ == '__main__':
    app.run_server(debug=True)
