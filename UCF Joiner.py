import pandas as pd
from sys import argv
from os.path import dirname, basename
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from numpy import nan
from PyQt5.QtWidgets import (QApplication, QFileDialog, QWidget, QPushButton, QGridLayout, QLabel)
from PyQt5 import QtCore


FileLocations = {'File Name': [], 'Location': []}


class FileSelector(QWidget):
    def __init__(self):
        # noinspection PyArgumentList
        super(FileSelector, self).__init__()
        self.title = 'Purchase Order File'
        self.left = 900
        self.top = 500
        self.width = 520
        self.height = 200
        self.greeting()

    def greeting(self):
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setWindowTitle(self.title)
        grid_layout = QGridLayout()

        hello = QLabel('Please select a file', self)
        hello.setAlignment(QtCore.Qt.AlignCenter)
        hello.move(QtCore.Qt.AlignCenter-25, 50)
        hello.setStyleSheet('font-size:18pt; font-weight:600')
        grid_layout.addWidget(hello)

        btn = QPushButton('Search', self)
        btn.clicked.connect(self.search_file)
        btn.move(QtCore.Qt.AlignCenter+100, 150)
        btn.setStyleSheet('font-size:10pt')
        grid_layout.addWidget(btn)
        # noinspection PyTypeChecker
        btn.clicked.connect(self.close)

    def search_file(self):
        options = QFileDialog.Options()
        # noinspection PyCallByClass
        find_file, _ = QFileDialog.getOpenFileName(self, 'UCF Source File', '',
                                                   'Excel Files (*.xlsx *xls)',
                                                   options=options)
        FileLocations['Location'].append(find_file)


if __name__ == '__main__':
    app = QApplication(argv)
    app.setStyle('Fusion')
    window = FileSelector()
    window.show()
    app.exec_()

File = FileLocations['Location'][0]

full_path = dirname(File) + '/'

OriginalName = basename(File).split('.')[0]


file = File

ExcelFile = pd.ExcelFile(file)
wb = Workbook()
ws = wb.active
"""
Since there are multiple sheets, I felt it was best to just read the entire workbook from the start.
Puts the slowest part of the program at the start
"""


POI = ['Supplier#', 'Supplier Stuffing', 'ETD Mother Vessel', 'ETD Origin Port', 'Planned ETA to Port of discharge',
       'ETA to US Port', 'ETA to Door', 'Container Received', 'Store', 'Status', 'ETA', 'Ship Line',
       'Port of Discharge']

AbvLocations = ['ALE', 'ASH', 'AUS', 'BAT', 'BOS', 'BIR', 'BUC', 'CHA', 'CHAT', 'CHI', 'CIN', 'COL', 'DAL', 'DET',
                'HOU', 'HUN', 'IND', 'KAN', 'KNO', 'LA', 'LR', 'MAR', 'MEM', 'MIA', 'MIN', 'MTP', 'NAS', 'NOLA', 'ORL',
                'PAR', 'PIT', 'PDX', 'RAL', 'SAN', 'SAV', 'TAM', 'GRN']


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
    blue = '0000FF'
    pink = 'FF00FF'
    black = '000000'
    apply_format = 'A3:' + ws.dimensions.split(':')[1]

    o_rule = FormulaRule(formula=['=$J3="O"'], font=Font(color=blue))
    w_rule = FormulaRule(formula=['=$J3="W"'], font=Font(color=pink))
    r_rule = FormulaRule(formula=['=$J3="R"'], font=Font(color=black))
    n_rule = FormulaRule(formula=['$J3="N"'], font=Font(color=black))

    ws.conditional_formatting.add(apply_format, o_rule)
    ws.conditional_formatting.add(apply_format, w_rule)
    ws.conditional_formatting.add(apply_format, r_rule)
    ws.conditional_formatting.add(apply_format, n_rule)

    ws.auto_filter.ref = ws.dimensions


def labels():
    for col_num, title in enumerate(POI):
        cell = ws.cell(row=1, column=col_num + 1)

        cell.value = title
        cell.font = Font(size=10, bold=True)
        ws.column_dimensions[cell.column_letter].width = len(title) + 7


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
    if page >= len(ExcelFile.sheet_names):
        return
    if ExcelFile.sheet_names[page] not in AbvLocations:
        raise MoveOn

    store = ExcelFile.parse(sheet_name=page).columns[0]
    book_page = ExcelFile.parse(sheet_name=page, skiprows=2)
    book_page.rename(columns={'Unnamed: 0': 'Status'}, inplace=True)
    book_page.columns = book_page.columns.str.rstrip()
    book_page.replace(r'^\s*$', nan, regex=True, inplace=True)
    book_page.dropna(how='all', inplace=True)
    book_page['Store'] = store
    updated_book = book_page[POI].copy()
    updated_book.fillna('', inplace=True)
    updated_book.rename_axis(None, inplace=True)

    for row_addition in dataframe_to_rows(updated_book, index=False, header=False):
        row_addition[6] = extractor(item=row_addition[6], d_time=True)
        row_addition[7] = extractor(item=row_addition[7], d_time=True)
        row_addition[5] = extractor(item=row_addition[5], d_time=True)
        row_addition[1] = extractor(item=row_addition[1], d_time=True)

        row_addition[10] = extractor(item=row_addition[10], latest_shipment=True)  # ETA
        row_addition[4] = extractor(item=row_addition[4], latest_shipment=True)
        row_addition[8] = extractor(item=row_addition[8])
        row_addition[3] = extractor(item=row_addition[3], latest_shipment=True)
        row_addition[2] = extractor(item=row_addition[2], latest_shipment=True)
        ws.append(row_addition)


for x in range(len(AbvLocations)):
    try:
        page_information(x)
    except MoveOn:
        pass

conditional_formatting()

wb.save(full_path + str(OriginalName) + ' - UCF.xlsx')
