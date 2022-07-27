import openpyxl
import dataSource


def check_comp(raw_data, data):
    temp_raw_data_str = ''
    temp_data_str = ''
    raw_data_str = ''
    data_str = ''
    for item in raw_data.split('株式会社'):
        temp_raw_data_str = temp_raw_data_str + item
    for item in temp_raw_data_str.split('(株)'):
        raw_data_str = raw_data_str + item
    for item in data.split('株式会社'):
        temp_data_str = temp_data_str + item
    for item in temp_data_str.split('(株)'):
        data_str = data_str + item
    if raw_data_str == data_str:
        return '株式会社₌(株)'
    else:
        return ''


def check_space(raw_data, data):
    if raw_data.strip().replace(" ", "") == data.strip().replace(" ", ""):
        return "スペースがある"
    else:
        return ""


def check_number(raw_data, data):
    if int(raw_data) == int(data):
        return f"{data}₌{raw_data}"
    else:
        return ""


def handle_reappraisal_result_list(excel_detail):
    result_list = excel_detail['reappraisalResult']['differencesList']
    for resultItem in result_list:
        reasons = ""
        for differentItem in resultItem['different']:
            raw_data = resultItem['rawData'][differentItem]
            data = resultItem['data'][differentItem]
            if differentItem == '住所':
                reasons = reasons + (';' if reasons != "" else "") + check_space(raw_data, data)
            if differentItem == '取引先名称':
                reasons = reasons + (';' if reasons != "" else "") + check_comp(raw_data, data)
            if differentItem == '預金種別':
                reasons = reasons + (';' if reasons != "" else "") + check_number(raw_data, data)
            if differentItem == '口座番号':
                reasons = reasons + (';' if reasons != "" else "") + check_number(raw_data, data)
        resultItem['reasonCell'].value = reasons
        excel_detail['workBook'].save(filename=excel_detail['excelName'])


def handle_target_excel(excel_list):
    for excel_list_item in excel_list:
        excel_detail = dataSource.load_excel(excel_list_item['name'])
        print(excel_detail)
        if excel_detail['type'] == 'reappraisalResult':
            if excel_detail['reappraisalResult'] is not {}:
                handle_reappraisal_result_list(excel_detail)
        elif excel_detail['type'] == 'manage':
            pass
        else:
            pass


if __name__ == '__main__':
    handle_target_excel(dataSource.get_excel_list())
