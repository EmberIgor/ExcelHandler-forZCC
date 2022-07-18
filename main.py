import openpyxl
import dataSource

if __name__ == '__main__':
    excel_list = dataSource.get_excel_list()
    for excel_list_item in excel_list:
        excelDetail = dataSource.load_excel(excel_list_item['name'])
        print(excelDetail['sheetList'][0]['differencesList'][0]['reasonCell'].value)
