import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import pandas as pd

# Get data
data_dir = os.path.join('..', 'data')
input_file = os.path.join(data_dir, 'Linelist example_input file.xlsx')
linelist = pd.read_excel(input_file, sheet_name='Linelist')[['Date of Assessment', 'Village']]
linelist = linelist[~pd.isnull(linelist).transpose().any()]
epiweek = pd.read_excel(input_file, sheet_name='Epiweeks', header=2)[['Epi week', 'Month', 'First day in week']]
geo = pd.read_excel(input_file, sheet_name='GEO')[['Camp/village', 'GEOID (pcode)?']]
geo.columns = ['Village', 'CODE_LOCATION']

# Transform data
linelist['Date of Assessment'] = pd.to_datetime(linelist['Date of Assessment'])
linelist['Week'] = linelist['Date of Assessment'].apply(lambda t: t.isocalendar()[1])
linelist['Year'] = linelist['Date of Assessment'].apply(lambda t: t.year)
linelist['Week'] = linelist['Date of Assessment'].apply(lambda t: t.isocalendar()[1])
# TODO : Ensure Village is correct before
linelist_agg = linelist.groupby(['Village', 'Week', 'Year']).count()
linelist_agg.columns = ['Num of cases']
linelist_agg = linelist_agg.reset_index()
epiweek['Year'] = -1
for i in range(len(epiweek)):
    if type(epiweek.loc[i, 'First day in week']) is str:
        epiweek.loc[i, 'First day in week'] = datetime.datetime.strptime(epiweek.loc[i, 'First day in week'], '%d/%m/%Y')
    epiweek.loc[i, 'Year'] = epiweek.loc[i, 'First day in week'].year
    if epiweek.loc[i, 'Epi week'] == 1 and epiweek.loc[i, 'First day in week'].month == 12:
        epiweek.loc[i, 'Year'] += 1
epiweek = epiweek[['Epi week', 'Month', 'Year']]
epiweek.columns = ['Week', 'Month', 'Year']

# Merge data
linelist_agg_geo = pd.merge(linelist_agg, geo, how='left', on='Village')
res = pd.DataFrame()
# TODO : Test this for loop is not too much CPU heavy
for code_loc in sorted(linelist_agg_geo['CODE_LOCATION'].unique()):
    code_loc_linelist = linelist_agg_geo[linelist_agg_geo['CODE_LOCATION'] == code_loc]
    code_loc_cases = pd.merge(epiweek, code_loc_linelist, how='left', on=['Year', 'Week'])
    code_loc_cases['CODE_LOCATION'] = code_loc
    res = pd.concat([res, code_loc_cases], axis=0, ignore_index=True)

# Export result
wb = Workbook()
ws = wb.active
for r in dataframe_to_rows(res[['CODE_LOCATION', 'Year', 'Month', 'Week', 'Num of cases']], index=False, header=True):
    ws.append(r)
# Apply color and border
thin = Side(border_style="thin", color="000000")
for col in ['A', 'B', 'C', 'D']:
    for index in range(1, len(res) + 2):
        ws[col + str(index)].fill = PatternFill("solid", fgColor="D0CECE")
        ws[col + str(index)].border = Border(top=thin, left=thin, right=thin, bottom=thin)
for index in range(1, len(res) + 2):
    ws['E' + str(index)].fill = PatternFill("solid", fgColor="C6E0B4")
    ws['E' + str(index)].border = Border(top=thin, left=thin, right=thin, bottom=thin)
# Apply bold
for col in ['A', 'B', 'C', 'D', 'E']:
    ws[col + '1'].font = Font(bold=True)
for index in range(1, len(res) + 2):
    ws['A' + str(index)].font = Font(bold=True)
# Apply column width
ws.column_dimensions['A'].width = 21.71
ws.column_dimensions['B'].width = 8.43
ws.column_dimensions['C'].width = 14.57
ws.column_dimensions['D'].width = 9
ws.column_dimensions['E'].width = 26.57
wb.save(os.path.join(data_dir, 'res_V2.xlsx'))
