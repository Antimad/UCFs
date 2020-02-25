import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from numpy import nan

file = 'Report215.xlsx'

ExcelFile = pd.ExcelFile(file)
wb = Workbook()
ws = wb.active
"""
Since there are multiple sheets, I felt it was best to just read the entire workbook from the start.
Puts the slowest part of the program at the start
"""

POI = ['Store', 'Status', 'Supplier#', 'ETA', 'Container Received', 'Port of Discharge', 'ETA to US Port',
       'Supplier Stuffing',
       'ETD Origin Port', 'ETD Mother Vessel', 'Planned ETA to Port of discharge', 'ETA to Door']

AbvLocations = ['ALE', 'ASH', 'AUS', 'BAT', 'BIR', 'BUC', 'CHA', 'CHAT', 'CHI', 'CIN', 'COL', 'DAL', 'DET', 'HOU',
                'HUN', 'IND', 'KAN', 'KNO', 'LA', 'LR', 'MAR', 'MEM', 'MIA', 'MIN', 'MTP', 'NAS', 'NOLA', 'ORL',
                'PAR', 'PIT', 'PDX', 'RAL', 'SAN', 'SAV', 'TAM']


class MoveOn(Exception):
    pass


# Colors
RED = 'FF0000'
Bold = Font(bold=True)
BoldRed = Font(bold=True, color=RED)
Normal = Font(bold=False)
Hor_Center = Alignment(horizontal='center', vertical='bottom')
Hor_Left = Alignment(horizontal='left', vertical='bottom')
TextWrap = Alignment(wrap_text=True, horizontal='center', vertical='bottom')
TitleBorder = Border(top=Side(border_style='thin', color='000000'),
                     bottom=Side(border_style='thin', color='000000'))

"""
FormulaRule to be used. 

O - On Order:       Blue Font   #0000FF
W - On Water:       Pink Font   #FF00FF
R - Received:       Black Font  #000000
N - New Product:    Black Font  #000000
"""


def conditional_formatting():
    Blue = '0000FF'
    Pink = 'FF00FF'
    Black = '000000'
    apply_format = 'A3:' + ws.dimensions.split(':')[1]
    O_Rule = FormulaRule(formula=['=$B3="O"'], font=Font(color=Blue))
    W_Rule = FormulaRule(formula=['=$B3="W"'], font=Font(color=Pink))
    R_Rule = FormulaRule(formula=['=$B3="R"'], font=Font(color=Black))
    N_Rule = FormulaRule(formula=['$B3="N"'], font=Font(color=Black))
    ws.conditional_formatting.add(apply_format, O_Rule)
    ws.conditional_formatting.add(apply_format, W_Rule)
    ws.conditional_formatting.add(apply_format, R_Rule)
    ws.conditional_formatting.add(apply_format, N_Rule)


def labels():
    for col_num, title in enumerate(POI):
        cell = ws.cell(row=1, column=col_num + 1)

        cell.value = title
        cell.font = Font(size=15, bold=True)
        ws.column_dimensions[cell.column_letter].width = len(title) + 7

    ws.append([])


labels()


def extractor(item, d_time=False, latest_shipment=False):
    edit = item
    if d_time:
        edit = pd.to_datetime(item, errors='coerce').date()
        if edit is pd.NaT:
            edit = ''

    if latest_shipment:
        try:
            edit = str(item.split(' ')[-1])
        except AttributeError:
            edit = pd.to_datetime(item, errors='coerce').date()
            edit = str(edit.month) + '/' + str(edit.day)

    return edit


def page_information(page):
    if ExcelFile.sheet_names[page] not in AbvLocations:
        raise MoveOn

    Store = ExcelFile.parse(sheet_name=page).columns[0]
    book_page = ExcelFile.parse(sheet_name=page, skiprows=2)
    book_page.rename(columns={'Unnamed: 0': 'Status'}, inplace=True)
    book_page.columns = book_page.columns.str.rstrip()
    book_page.replace(r'^\s*$', nan, regex=True, inplace=True)
    book_page.dropna(how='all', inplace=True)
    book_page['Store'] = Store
    updated_book = book_page[POI].copy()
    updated_book.fillna('', inplace=True)
    updated_book.rename_axis(None, inplace=True)

    for row_addition in dataframe_to_rows(updated_book, index=False, header=False):
        row_addition[4] = extractor(item=row_addition[4], d_time=True)
        row_addition[6] = extractor(item=row_addition[6], d_time=True)
        row_addition[7] = extractor(item=row_addition[7], d_time=True)
        row_addition[11] = extractor(item=row_addition[11], d_time=True)

        row_addition[8] = extractor(item=row_addition[8], latest_shipment=True)
        row_addition[9] = extractor(item=row_addition[9], latest_shipment=True)
        row_addition[10] = extractor(item=row_addition[10], latest_shipment=True)
        ws.append(row_addition)


for x in range(len(AbvLocations)):
    try:
        page_information(x)
    except MoveOn:
        pass

conditional_formatting()
wb.save('test.xlsx')
