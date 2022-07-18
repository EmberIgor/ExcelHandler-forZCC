import openpyxl
import dataSource

if __name__ == '__main__':
    excel_list = dataSource.get_excel_list()
    for excel_list_item in excel_list:
        print(dataSource.load_excel(excel_list_item['name']))
