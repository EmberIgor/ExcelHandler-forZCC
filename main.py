import openpyxl
import dataSource

if __name__ == '__main__':
    dataSource.get_excel_list()
    dataSource.load_excel('testExcel.xlsx')
