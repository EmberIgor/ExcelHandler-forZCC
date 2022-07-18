import openpyxl
import os
import re
from enum import Enum

from openpyxl.cell import MergedCell


class Fields(Enum):
    idField = '申请番号'
    automaticExtractionField = '数据A'
    dataField = '数据B'
    differencesField = '差异'
    reasonField = '原因'


def get_excel_list():
    """
    获取当前目录下excel文件列表
    :return: 当前目录下excel文件列表
    """
    all_item_list = os.listdir('.')
    excel_list = []
    for item in all_item_list:
        item_matcher = re.match('^(?!~\\$).+.xlsx$|^(?!~\\$).+.xlsm$', item)
        if item_matcher is not None:
            excel_list.append({
                'name': item_matcher.group(),
                'path': f'{os.getcwd()}\\{item_matcher.group()}'
            })
    return excel_list


def load_excel(excel_name):
    """
    加载excel数据
    :param excel_name: excel文件全名
    :return: excel概况、各条目信息
    """
    wb = openpyxl.load_workbook(filename=excel_name)
    excel_detail = {
        "excelName": excel_name,
        "sheetList": []
    }
    for sheet_item in wb.sheetnames:
        if re.match('^\\d+$', sheet_item) is not None:
            excel_detail["sheetList"].append(load_sheet(sheet_item, wb))
    return excel_detail


def load_sheet(sheet_name, wb):
    """
    加载sheet数据
    :param sheet_name: sheet全名
    :param wb: workbook
    :return: sheet概况、各条目信息
    """
    sheet = wb[sheet_name]
    differences_list = []
    automatic_extraction_field_info = None
    id_field_info = None
    data_field_info = None
    differences_field_info = None
    reason_field_info = None
    # 下面这个循环用于获取表头坐标
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        if None not in [automatic_extraction_field_info, id_field_info, data_field_info, reason_field_info,
                        differences_field_info]:
            break
        for cell in row:
            if cell is not None and not isinstance(cell, MergedCell):
                cell_info = {
                    'column_letter': cell.column_letter,
                    'col_idx': cell.col_idx,
                    'row': cell.row
                }
                if cell.value == Fields.automaticExtractionField.value:
                    automatic_extraction_field_info = cell_info
                elif cell.value == Fields.idField.value:
                    id_field_info = cell_info
                elif cell.value == Fields.dataField.value:
                    data_field_info = cell_info
                elif cell.value == Fields.differencesField.value:
                    differences_field_info = cell_info
                elif cell.value == Fields.reasonField.value:
                    reason_field_info = cell_info
                if None not in [automatic_extraction_field_info, id_field_info, data_field_info, reason_field_info,
                                differences_field_info]:
                    break
    # 获取条目属性
    header_row = sheet[f"{automatic_extraction_field_info['row'] + 1}"]
    raw_data_props = []
    for cell in header_row[automatic_extraction_field_info['col_idx'] - 1:data_field_info['col_idx'] - 1]:
        raw_data_props.append(cell.value)
    data_props = []
    for cell in header_row[data_field_info['col_idx'] - 1:differences_field_info['col_idx'] - 1]:
        data_props.append(cell.value)
    differences_props = []
    for cell in header_row[differences_field_info['col_idx'] - 1:reason_field_info['col_idx'] - 1]:
        differences_props.append(cell.value)
    # 获取原数据
    current_idx = 0
    current_row = sheet[f"{automatic_extraction_field_info['row'] + 2 + current_idx}"]
    while re.match('\\d+', str(current_row[id_field_info['col_idx'] - 1].value)) is not None:
        differences_item = {
            Fields.idField.value: str(current_row[id_field_info['col_idx'] - 1].value),
            'rawData': {},
            'data': {},
            'different': [],
            'reason': '',
            'reasonCell': None
        }
        # 写入原数据
        prop_idx = 0
        for prop in raw_data_props:
            differences_item['rawData'][prop] = current_row[
                prop_idx + automatic_extraction_field_info['col_idx'] - 1].value
            prop_idx += 1
        # 写入现数据
        prop_idx = 0
        for prop in data_props:
            differences_item['data'][prop] = current_row[prop_idx + data_field_info['col_idx'] - 1].value
            prop_idx += 1
        # 写入差异点
        prop_idx = 0
        for prop in differences_props:
            if current_row[prop_idx + differences_field_info['col_idx'] - 1].value == 'X':
                differences_item['different'].append(header_row[prop_idx + differences_field_info['col_idx'] - 1].value)
            prop_idx += 1
        # 写入reasonCell
        differences_item['reasonCell'] = current_row[reason_field_info['col_idx'] - 1]
        current_idx += 1
        current_row = sheet[f"{automatic_extraction_field_info['row'] + 2 + current_idx}"]
    differences_list.append(differences_item)
    return {
        'sheetName': sheet_name,
        'differencesList': differences_list
    }
