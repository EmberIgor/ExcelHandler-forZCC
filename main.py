import os

import dataSource
import time
from tqdm import tqdm

# 管理表数据
manage_excel_data = {
    "domesticData": [],
    "foreignData": [],
    "surroundingData": []
}
# 待处理的目标表
target_excel_list = []

# 被跳过的文件
skip_file_list = []


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


# 检查"保留の件を処理したケース"和"当日で、同じ取引先コードを不同申請書で申請された"
def check_reserved_case(request_item):
    # 检查结果
    check_result = {
        "isReserved": False,
        "reason": "",
    }
    # 查询到的次数
    query_count = 0
    for domestic_item in manage_excel_data['domesticData']:
        if int(domestic_item['取引先番号\n必須']) == int(request_item['id']):
            if domestic_item['受付日'] is not None:
                if domestic_item['受付日'].strftime("%Y-%m-%d") != time.strftime("%Y-%m-%d", time.localtime()):
                    check_result['isReserved'] = True
                    check_result['reason'] = "保留の件を処理したケース"
                    return check_result
                else:
                    query_count = query_count + 1
    for foreign_item in manage_excel_data['foreignData']:
        if int(foreign_item['取引先番号\n必須']) == int(request_item['id']):
            if foreign_item['受付日'] is not None:
                if foreign_item['受付日'].strftime("%Y-%m-%d") != time.strftime("%Y-%m-%d", time.localtime()):
                    check_result['isReserved'] = True
                    check_result['reason'] = "保留の件を処理したケース"
                    return check_result
                else:
                    query_count = query_count + 1
    if query_count == 2:
        check_result['isReserved'] = True
        check_result['reason'] = "当日で、同じ取引先コードを不同申請書で申請された"
        return check_result
    else:
        return check_result


# 检查"申請書通りに取引先名を変更する場合、枝番の取引先名も変更した"
def check_branch_name(request_item, request_list):
    # 检查结果
    check_result = {
        "isReserved": False,
        "reason": "申請書通りに取引先名を変更する場合、枝番の取引先名も変更した",
    }
    # 编号列表
    number_list = []
    current_id = int(request_item['id'])
    for request_list_item in request_list:
        number_list.append(int(request_list_item['id']))
    number_list.sort()
    current_id_index = number_list.index(current_id)
    if current_id_index != 0:
        if number_list[current_id_index - 1] == current_id - 1:
            check_result['isReserved'] = True
            return check_result
    if current_id_index != len(number_list) - 1:
        if number_list[current_id_index + 1] == current_id + 1:
            check_result['isReserved'] = True
            return check_result
    return check_result


# 检查"変更の指示に基づき処理したケース（依頼元：経理課　山中さん）"
def check_change_case(request_item):
    # 检查结果
    check_result = {
        "isReserved": False,
        "reason": "変更の指示に基づき処理したケース（依頼元：経理課　山中さん）",
    }
    for surrounding_item in manage_excel_data['surroundingData']:
        if int(surrounding_item['申請番号必須']) == int(request_item['id']):
            check_result['isReserved'] = True
            return check_result
    return check_result


# 检查"申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり否認します）"和"申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり保留します）"
def check_preliminary(request_item):
    # 检查结果
    check_result = {
        "isReserved": False,
        "reason": "",
    }
    # 检查国内数据
    for domestic_item in manage_excel_data['domesticData']:
        if int(domestic_item['取引先番号\n必須']) == int(request_item['id']):
            if domestic_item['初鑑\nステータス'] == '否認済':
                check_result['isReserved'] = True
                check_result['reason'] = "申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり否認します）"
                return check_result
            elif domestic_item['初鑑\nステータス'] == '保留':
                check_result['isReserved'] = True
                check_result['reason'] = "申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり保留します）"
                return check_result
    # 检查国外数据
    for foreign_item in manage_excel_data['foreignData']:
        if int(foreign_item['取引先番号\n必須']) == int(request_item['id']):
            if foreign_item['初鑑\nステータス'] == '否認済':
                check_result['isReserved'] = True
                check_result['reason'] = "申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり否認します）"
                return check_result
            elif foreign_item['初鑑\nステータス'] == '保留':
                check_result['isReserved'] = True
                check_result['reason'] = "申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり保留します）"
                return check_result
    # 检查周边数据
    for surrounding_item in manage_excel_data['surroundingData']:
        if int(surrounding_item['申請番号必須']) == int(request_item['id']):
            if surrounding_item['初鑑\nステータス'] == '否認済':
                check_result['isReserved'] = True
                check_result['reason'] = "申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり否認します）"
                return check_result
            elif surrounding_item['初鑑\nステータス'] == '保留':
                check_result['isReserved'] = True
                check_result['reason'] = "申請書依頼の指示に基づき処理したケース（一回目処理しますが、チェックする時、問題あり保留します）"
                return check_result
    return check_result


def handle_reappraisal_result_list(excel_detail):
    global skip_file_list
    result_list = excel_detail['reappraisalResult']['differencesList']
    request_list = excel_detail['reappraisalResult']['requestList']
    # 处理再鑑結果列表
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
        try:
            excel_detail['workBook'].save(filename=excel_detail['excelName'])
        except PermissionError:
            skip_file_list.append(excel_detail['excelName'])
            return
    # 处理再鑑要求列表
    check_list = [check_reserved_case, check_branch_name, check_change_case, check_preliminary]
    for request_item in request_list:
        try:
            if request_item['type'] != '1' and request_item['type'] != 1:
                for check in check_list:
                    if check.__code__.co_argcount == 1:
                        check_result = check(request_item)
                    else:
                        check_result = check(request_item, request_list)
                    if check_result['isReserved']:
                        request_item['reason'].value = check_result['reason']
                        try:
                            excel_detail['workBook'].save(filename=excel_detail['excelName'])
                        except PermissionError:
                            skip_file_list.append(excel_detail['excelName'])
                            break
                        continue
        except PermissionError:
            skip_file_list.append(excel_detail['excelName'])
            return


# 初始化各表数据
def init_data(excel_list):
    global manage_excel_data
    global target_excel_list
    print(f"加载Excel文件(xlsx、xlsm):")
    pbar = tqdm(total=len(excel_list))
    for excel_list_item in excel_list:
        excel_detail = dataSource.load_excel(excel_list_item['name'])
        if excel_detail['type'] == 'reappraisalResult':
            if excel_detail['reappraisalResult'] is not {}:
                target_excel_list.append(excel_detail)
        elif excel_detail['type'] == 'manage':
            manage_excel_data = excel_detail
        pbar.update(1)
    pbar.close()


if __name__ == '__main__':
    init_data(dataSource.get_excel_list())
    print("\n处理Excel文件:")
    total_bar = tqdm(total=len(target_excel_list))
    for target_excel_detail in target_excel_list:
        handle_reappraisal_result_list(target_excel_detail)
        total_bar.update(1)
    total_bar.close()
    print('\n')
    for slip_file in skip_file_list:
        print(f"文件[{slip_file}]被占用，已跳过，请关闭文件后再试")
    print("处理完成,可以关闭程序")
    input("按回车键退出")
