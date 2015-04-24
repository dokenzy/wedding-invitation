# -*- coding: utf-8 -*-

"""
Python: 3.4.2
author: dokenzy
date: 2015. 4. 21
license: MIT License
"""

import sys
import xlrd

EXCEL_NAME = 'list.xlsx'
HEADER = ['이름', '주소', '세부주소', '우편번호']
NUM_COLS = 4 # 첫번째 4개의 컬럼만 사용. 나머지 컬럼은 무시

guests = []

def get_guests(filename):
    """
    엑셀파일에서 손님 정보를 guests 리스트에 저장하는 메인 함수
    """
    workbook = xlrd.open_workbook(filename)
    _get_sheet(workbook)
    data = ''
    for guest in guests:
        row = '\guest{%s}{%s %s 귀하}{%s}\n' % (guest[0], guest[1], guest[2], guest[3])
        data += row
    with open('data.tex', 'w') as dat:
        dat.write(data)

def _get_sheet(workbook):
    """
    각 시트의 헤더를 검사하고, 손님 정보를 가져온다.
    """
    sheet_names = workbook.sheet_names()
    for sheetname in sheet_names:
        sheet = workbook.sheet_by_name(sheetname)
        header = sheet.row(0)
        _chk_header(header)
        _get_info(sheet)


def _chk_header(header):
    """
    헤더를 검사해서 HEADER와 다르면 종료한다.
    """
    for idx, head in enumerate(header):
        if head.value != HEADER[idx]:
            print('헤더 오류')
            print('프로그램을 종료합니다.')
            sys.exit(1)


def _get_info(sheet):
    """
    주어진 시트에서 손님 정보를 가져와서 guests 리스트에 바로 저장한다.
    """
    NUM_ROWS = sheet.nrows
    for row_idx in range(1, NUM_ROWS):
        juso = sheet.cell(row_idx, 1).value
        detail_juso = sheet.cell(row_idx, 2).value
        name = sheet.cell(row_idx, 0).value
        zipcode = sheet.cell(row_idx, 3).value
        guests.append((juso, detail_juso, name, zipcode))

if __name__ == '__main__':
    get_guests(EXCEL_NAME)
