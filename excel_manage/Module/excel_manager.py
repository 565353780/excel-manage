#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os

from excel_manage.Method.io import \
    createExcel, removeExcel, readExcel


class ExcelManager(object):

    def __init__(self, excel_folder_path=None):
        self.excel_folder_path = None

        if excel_folder_path is not None:
            assert self.loadExcelFolder(excel_folder_path)
        return

    def loadExcelFolder(self, excel_folder_path):
        if excel_folder_path[-1] != "/":
            excel_folder_path += "/"

        os.makedirs(excel_folder_path, exist_ok=True)

        self.excel_folder_path = excel_folder_path
        return True

    def getExcelFilePath(self, excel_file_name):
        if '.xlsx' not in excel_file_name:
            excel_file_name += '.xlsx'

        excel_file_path = self.excel_folder_path + excel_file_name

        return excel_file_path

    def createExcel(self, excel_file_name):
        assert self.excel_folder_path is not None

        excel_file_path = self.getExcelFilePath(excel_file_name)
        if os.path.exists(excel_file_path):
            print("[WARN][ExcelManager::createExcel]")
            print("\t this excel file already exist!")
            print("\t", excel_file_path)
            return False

        createExcel(excel_file_path)
        return True

    def removeExcel(self, excel_file_name):
        excel_file_path = self.getExcelFilePath(excel_file_name)

        removeExcel(excel_file_path)
        return True
