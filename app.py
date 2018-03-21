# -*- coding: utf-8 -*-
"""
Created on Wed Mar 21 16:53:28 2018

@author: omer
"""

import dash
import dash_core_components as dcc
import dash_html_components as html

import pandas as pd
import numpy as np

# Spreadsheet filename
filename = '../data/Upload_Sheet_12_Mar_V2.xlsx'

# Read spreadsheet
row_id_headers = 1
skiprows = 0
data_sheet = pd.read_excel(filename, sheet_name='DATA', skiprows=skiprows, header=row_id_headers)

# Created and resolved issues
created_dates_slice = data_sheet['Created Date']
resolved_dates_slice_filtered = data_sheet[u'Resolution Date®'][data_sheet['Status'] == 'Closed']
    
created_dates_rounded = pd.DatetimeIndex(created_dates_slice).normalize()
resolved_dates_rounded = pd.DatetimeIndex(resolved_dates_slice_filtered).normalize()

firstDate = min([created_dates_rounded[0], resolved_dates_rounded[0]])
lastDate = max([created_dates_rounded[-1], resolved_dates_rounded[-1]])

# daily
dates_bins = pd.date_range(firstDate, lastDate+pd.DateOffset(1))
createdQ_daily_histogram = np.histogram(created_dates_rounded.values.astype('int64'), dates_bins.values.astype('int64'))
resolvedQ_daily_histogram = np.histogram(resolved_dates_rounded.values.astype('int64'), dates_bins.values.astype('int64'))
unresolved_queries_daily = np.cumsum(createdQ_daily_histogram[0] - resolvedQ_daily_histogram[0])

# weekly
weeks_bins = np.array(range(firstDate.week, lastDate.week+2))
createdQ_weekly_histogram = np.histogram(created_dates_rounded.week.values, weeks_bins)
resolvedQ_weekly_histogram = np.histogram(resolved_dates_rounded.week.values, weeks_bins)
unresolved_queries_weekly = np.cumsum(createdQ_weekly_histogram[0] - resolvedQ_weekly_histogram[0])

# dashboard app
app = dash.Dash()

app.layout = html.Div(children=[
    html.H1(children='Analytics Dashboard'),

    html.Div(children='''
        based on "Dash: A web application framework for Python".
    '''),

    dcc.Graph(
        id='daily-graph',
        figure={
            'data': [
                {'x': dates_bins[:-1], 'y': createdQ_daily_histogram[0], 'type': 'bar', 'name': 'Incoming queries'},
                {'x': dates_bins[:-1], 'y': resolvedQ_daily_histogram[0], 'type': 'bar', 'name': 'Resolved queries'},
                {'x': dates_bins[:-1], 'y': unresolved_queries_daily, 'type': 'line', 'name': 'Unresolved queries'},
            ],
            'layout': {
                'title': 'Created, resolved and unresolved queries (daily)',
                'xaxis':{
                    'title':'dates'
                },
                'yaxis':{
                     'title':'Number of issues'
                }
            }
        }
    ),
    dcc.Graph(
        id='weekly-graph',
        figure={
            'data': [
                {'x': weeks_bins[:-1], 'y': createdQ_weekly_histogram[0], 'type': 'bar', 'name': 'Incoming queries'},
                {'x': weeks_bins[:-1], 'y': resolvedQ_weekly_histogram[0], 'type': 'bar', 'name': 'Resolved queries'},
                {'x': weeks_bins[:-1], 'y': unresolved_queries_weekly, 'type': 'line', 'name': 'Unresolved queries'},
            ],
            'layout': {
                'title': 'Created, resolved and unresolved queries (weekly)',
                'xaxis':{
                    'title':'week numbers'
                },
                'yaxis':{
                     'title':'Number of issues'
                }
            }
        }
    )
])

if __name__ == '__main__':
    app.run_server(debug=True)
