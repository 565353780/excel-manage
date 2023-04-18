#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os

import xlrd
import xlwt

from excel_manage.Config.title import TITLE_LIST
from excel_manage.Method.path import createFileFolder, removeFile, renameFile


def createExcel(excel_file_path, sheet_name='Sheet 0', title_list=TITLE_LIST):
    workbook = xlwt.Workbook(encoding='utf-8')
    #  workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(sheet_name)
    for i, title in enumerate(title_list):
        worksheet.write(1, i, title)

    createFileFolder(excel_file_path)

    workbook.save(excel_file_path)
    return True


def removeExcel(excel_file_path):
    removeFile(excel_file_path)
    return True


def readExcel(excel_file_path):
    assert os.path.exists(excel_file_path)
    return True
