import openpyxl
import dataSource


def check_comp(rawData, data):
    tempRawDataStr = ''
    tempDataStr = ''
    rawDataStr = ''
    dataStr = ''
    for item in rawData.split('株式会社'):
        tempRawDataStr = tempRawDataStr + item
    for item in tempRawDataStr.split('(株)'):
        rawDataStr = rawDataStr + item
    for item in data.split('株式会社'):
        tempDataStr = tempDataStr + item
    for item in tempDataStr.split('(株)'):
        dataStr = dataStr + item
    if rawDataStr == dataStr:
        return '株式会社₌(株)'
    else:
        return ''


def check_space(rawData, data):
    if rawData.strip().replace(" ", "") == data.strip().replace(" ", ""):
        return "スペースがある"
    else:
        return ""


def check_number(rawData, data):
    if int(rawData) == int(data):
        return f"{data}₌{rawData}"
    else:
        return ""


def handle_reappraisal_result_list(excelDetail):
    resultList = excelDetail['reappraisalResult']['differencesList']
    for resultItem in resultList:
        reasons = ""
        for differentItem in resultItem['different']:
            rawData = resultItem['rawData'][differentItem]
            data = resultItem['data'][differentItem]
            if differentItem == '住所':
                reasons = reasons + (';' if reasons != "" else "") + check_space(rawData, data)
            if differentItem == '取引先名称':
                reasons = reasons + (';' if reasons != "" else "") + check_comp(rawData, data)
            if differentItem == '預金種別':
                reasons = reasons + (';' if reasons != "" else "") + check_number(rawData, data)
            if differentItem == '口座番号':
                reasons = reasons + (';' if reasons != "" else "") + check_number(rawData, data)
        resultItem['reasonCell'].value = reasons
        excelDetail['workBook'].save(filename=excelDetail['excelName'])


def handle_target_excel(excel_list):
    for excel_list_item in excel_list:
        excelDetail = dataSource.load_excel(excel_list_item['name'])
        print(excelDetail)
        if excelDetail['reappraisalResult'] is not {}:
            handle_reappraisal_result_list(excelDetail)


if __name__ == '__main__':
    handle_target_excel(dataSource.get_excel_list())
