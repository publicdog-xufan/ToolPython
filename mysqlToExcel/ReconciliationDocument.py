# import pandas as pd
# import os
import pymysql
import xlrd
from xlutils.copy import copy


class ReconciliationDocument(object):
    def __init__(self):
        self.HOST = '127.0.0.1'
        self.PORT = 3306
        self.USER = 'root'
        self.PASSWORD = 'ozJOgV17Bey'

    # def Read_SQL(self, SQL):
    #     conn = pymysql.connect(host=self.HOST, port=self.PORT, user=self.USER, password=self.PASSWORD,
    #                            db='openbank_p3', charset='utf8')
    #     cursor = conn.cursor()
    #     cursor.execute(SQL)
    #     Result = cursor.fetchone()
    #     cursor.close()
    #     conn.close()
    #     return Result

    def Create_File(self, fire, numb):
        workbook = xlrd.open_workbook('D://write1.csv')
        sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
        numRow = sheet1.nrows  # 获取测试文件行数
        # 对文件的追加操作
        r_xls = xlrd.open_workbook("D://write1.csv")  # 读取excel文件
        numRow1 = r_xls.sheets()[0].nrows  # 获取处理文件已有的行数，注意此处以上面获取行号的差异
        excel = copy(r_xls)  # 将xlrd的对象转化为xlwt的对象
        table = excel.get_sheet(0)  # 获取要操作的sheet
        # 对excel表追加一行内容
        for r in range(len(fire)):
            table.write(numb, r, fire[r])  # 括号内分别为行数、列数、内容
            print('写入第', numb, '行,第', r,'列')
        excel.save('D://write1.csv')  # 保存并覆盖文件

    def Add_Value(self, numb):
        # 数据拼接
        Add_Value = ReconciliationDocument()
        ConsentID_SQL = 'SELECT consent_id from authorize_info WHERE bank_consent_id =%s' % numb
        CONSENT_ID = Add_Value.Read_SQL(ConsentID_SQL)[0]
        # print(CONSENT_ID)
        Permission_List = "ReadAccountAvailability|ReadAccountStatus|ReadAccountBalance|ReadAccountTransaction"
        Expiration_Date_SQL = 'SELECT expiration_date from authorize_info WHERE bank_consent_id =%s' % numb
        Expiration_Date = Add_Value.Read_SQL(Expiration_Date_SQL)[0]
        # print(Expiration_Date)
        transaction_from_date_SQL = 'SELECT transaction_from_date from authorize_info WHERE bank_consent_id =%s' % numb
        transaction_from_date = Add_Value.Read_SQL(transaction_from_date_SQL)[0]
        transaction_to_date_SQL = 'SELECT transaction_to_date from authorize_info WHERE bank_consent_id =%s' % numb
        transaction_to_date = Add_Value.Read_SQL(transaction_to_date_SQL)[0]
        creation_date_SQL = 'SELECT creation_date from authorize_info WHERE bank_consent_id =%s' % numb
        creation_date = Add_Value.Read_SQL(creation_date_SQL)[0]
        refresh_period_date_SQL = 'SELECT refresh_period_date from authorize_info WHERE bank_consent_id =%s' % numb
        refresh_period_date = Add_Value.Read_SQL(refresh_period_date_SQL)[0]
        STATUS_SQL = 'SELECT status from authorize_info WHERE bank_consent_id =%s' % numb
        STATUS = Add_Value.Read_SQL(STATUS_SQL)[0]
        status_update_date_SQL = 'SELECT status_update_date from authorize_info WHERE bank_consent_id =%s' % numb
        status_update_date = Add_Value.Read_SQL(status_update_date_SQL)[0]
        account_id_SQL = 'SELECT account_id from authorize_account WHERE consent_id ="%s"' % CONSENT_ID
        account_id = Add_Value.Read_SQL(account_id_SQL)[0]
        BANK_USER_NAME_SQL = 'SELECT rm_num_key from authorize_info WHERE consent_id ="%s"' % CONSENT_ID
        BANK_USER_NAME = Add_Value.Read_SQL(BANK_USER_NAME_SQL)[0]
        tsp_user_id_SQL = 'SELECT tsp_user_id from authorize_info WHERE consent_id ="%s"' % CONSENT_ID
        tsp_user_id = Add_Value.Read_SQL(tsp_user_id_SQL)[0]
        Organization_Name_SQL = 'SELECT org from authorize_info WHERE consent_id ="%s"' % CONSENT_ID
        Organization_Name = Add_Value.Read_SQL(Organization_Name_SQL)[0]
        Organization_ID_SQL = 'SELECT org_id from authorize_info WHERE consent_id ="%s"' % CONSENT_ID
        Organization_ID = Add_Value.Read_SQL(Organization_ID_SQL)[0]
        Organization_title_SQL = 'SELECT org_title from authorize_info WHERE consent_id ="%s"' % CONSENT_ID
        Organization_title = Add_Value.Read_SQL(Organization_title_SQL)[0]
        value = (
            CONSENT_ID, Permission_List, Expiration_Date, transaction_from_date, transaction_to_date, creation_date,
            refresh_period_date, STATUS,
            status_update_date, account_id, BANK_USER_NAME, tsp_user_id, Organization_Name, Organization_ID,
            Organization_title)
        print('查询完成')
        return value


if __name__ == '__main__':
    a = ReconciliationDocument()
    i = 0
    for i in range(1, 50001):
        c = 1111100000 + i
        b = a.Add_Value(str(c))
        a.Create_File(b, i)
        i = i + 1
    # 第一、如何进行写入数据累加

    # 第二、如何失效高效循环
