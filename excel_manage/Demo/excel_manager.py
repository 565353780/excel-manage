#!/usr/bin/env python
# -*- coding: utf-8 -*-

from excel_manage.Module.excel_manager import ExcelManager


def demo():
    excel_folder_path = "./test/"
    excel_file_name = "李常颢"

    excel_manager = ExcelManager(excel_folder_path)
    excel_manager.createExcel(excel_file_name)
    return True
