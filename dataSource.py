import openpyxl
import os
import re
import gc
from enum import Enum

from openpyxl.cell import MergedCell


class Target_Fields(Enum):
    """待修改Excel表头名称"""
    idField = '申請番号'
    automaticExtractionField = '申請書に記載された内容（マクロ自動抽出）'
    dataField = 'SAP仕入先一覧'
    differencesField = '相違箇所自動判定'
    reasonField = '再鑑者コメント'
    supplierField = '仕入先'


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
                'path': f'{os.getcwd()}\\{item_matcher.group()}',
            })
    return excel_list


def load_excel(excel_name):
    """
    分拣并读取excel
    :param excel_name: excle名称
    :return: excel详情
    """
    excelDetail = {
        "workBook": None,
        "type": "other",
        "excelName": excel_name,
        "reappraisalResult": {}
    }
    wb = openpyxl.load_workbook(filename=excel_name)
    for sheet_item in wb.sheetnames:
        if sheet_item == '再鑑結果':
            excelDetail = load_target_excel(excel_name, wb)
            break
        elif sheet_item == '【国内】':
            break
    return excelDetail


def load_target_excel(excel_name, wb):
    """
    加载待修改excel数据
    :param wb: workBook
    :param excel_name: excel文件全名
    :return: excel概况、再鑑結果sheet内容
    """
    # wb = openpyxl.load_workbook(filename=excel_name)
    excel_detail = {
        "workBook": wb,
        "type": "reappraisalResult",
        "excelName": excel_name,
        "reappraisalResult": load_reappraisal_result_sheet(wb)
    }
    return excel_detail


def load_reappraisal_result_sheet(wb):
    """
    加载再鑑結果sheet数据
    :param wb: workbook
    :return: sheet概况、各条目信息
    """
    sheet_name = '再鑑結果'
    sheet = wb[sheet_name]
    differences_list = []
    requestItemList = []
    fieldsInfo = {
        "automaticExtractionField": None,
        "idField": None,
        "dataField": None,
        "differencesField": None,
        "reasonField": None,
        "supplierField": None
    }
    # 下面这个循环用于获取表头坐标
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        if None not in fieldsInfo.values():
            break
        for cell in row:
            if cell is not None and not isinstance(cell, MergedCell):
                cell_info = {
                    'column_letter': cell.column_letter,
                    'col_idx': cell.col_idx,
                    'row': cell.row
                }
                for key, value in fieldsInfo.items():
                    if cell.value == Target_Fields[key].value:
                        fieldsInfo[key] = cell_info
                if None not in fieldsInfo.values():
                    break
    if None in fieldsInfo.values():
        return {
            'sheetName': sheet_name,
            'differencesList': [],
            'requestList': requestItemList
        }
    # 获取条目属性
    header_row = sheet[f"{fieldsInfo['automaticExtractionField']['row'] + 1}"]
    raw_data_props = []
    for cell in header_row[
                fieldsInfo['automaticExtractionField']['col_idx'] - 1:fieldsInfo['dataField']['col_idx'] - 1]:
        raw_data_props.append(cell.value)
    data_props = []
    for cell in header_row[fieldsInfo['dataField']['col_idx'] - 1:fieldsInfo['differencesField']['col_idx'] - 1]:
        data_props.append(cell.value)
    differences_props = []
    for cell in header_row[fieldsInfo['differencesField']['col_idx'] - 1:fieldsInfo['reasonField']['col_idx'] - 1]:
        differences_props.append(cell.value)
    # 获取差异数据
    current_idx = 0
    current_row = sheet[f"{fieldsInfo['automaticExtractionField']['row'] + 2 + current_idx}"]
    while re.match('\\d+', str(current_row[fieldsInfo['idField']['col_idx'] - 1].value)) is not None:
        differences_item = {
            Target_Fields.idField.value: str(current_row[fieldsInfo['idField']['col_idx'] - 1].value),
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
                prop_idx + fieldsInfo['automaticExtractionField']['col_idx'] - 1].value
            prop_idx += 1
        # 写入现数据
        prop_idx = 0
        for prop in data_props:
            differences_item['data'][prop] = current_row[prop_idx + fieldsInfo['dataField']['col_idx'] - 1].value
            prop_idx += 1
        # 写入差异点
        prop_idx = 0
        for prop in differences_props:
            if current_row[prop_idx + fieldsInfo['differencesField']['col_idx'] - 1].value == '×':
                differences_item['different'].append(
                    header_row[prop_idx + fieldsInfo['differencesField']['col_idx'] - 1].value)
            prop_idx += 1
        # 写入reasonCell
        differences_item['reasonCell'] = current_row[fieldsInfo['reasonField']['col_idx'] - 1]
        # 迭代
        current_idx += 1
        current_row = sheet[f"{fieldsInfo['automaticExtractionField']['row'] + 2 + current_idx}"]
        differences_list.append(differences_item)
    # 获取“本日が更新日の取引先の申請件数”
    current_idx = 0
    current_row = sheet[f"{fieldsInfo['supplierField']['row'] + 1 + current_idx}"]
    print(fieldsInfo['idField'])
    while re.match('\\d+', str(current_row[fieldsInfo['idField']['col_idx'] - 1].value)) is not None:
        requestItem = {
            "id": str(current_row[fieldsInfo['idField']['col_idx'] - 1].value),
            "type": str(current_row[fieldsInfo['idField']['col_idx']].value),
            "reason": current_row[fieldsInfo['idField']['col_idx'] + 1]
        }
        requestItemList.append(requestItem)
        current_idx += 1
        current_row = sheet[f"{fieldsInfo['supplierField']['row'] + 1 + current_idx}"]
    return {
        'sheetName': sheet_name,
        'differencesList': differences_list,
        'requestList': requestItemList
    }
