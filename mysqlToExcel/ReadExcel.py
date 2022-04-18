import datetime

import pymysql
import xlrd

import uuid


def Consentid_UUID(consentid):
    uuid_value = uuid.uuid3(uuid.NAMESPACE_DNS, consentid)
    return uuid_value


class ReadExcel(object):

    def __init__(self):
        # 数据库连接参数
        self.HOST = '127.0.0.1'
        self.PORT = 3306
        self.USER = 'root'
        self.PASSWORD = 'ozJOgV17Bey'
        self.sql = "INSERT INTO `reconcile_authorize_info` " \
                   "(" \
                   "id," \
                   "consent_id," \
                   "expiration_date," \
                   "last_refresh_consent_date," \
                   "transaction_from_date," \
                   "transaction_to_date," \
                   "creation_date," \
                   "status," \
                   "status_update_date," \
                   "tsp_user_id," \
                   "org," \
                   "org_id," \
                   "org_title," \
                   "update_by," \
                   "update_date," \
                   "create_by," \
                   "create_date," \
                   "reconcile_date," \
                   "reconcile_result," \
                   "reconcile_comments," \
                   "remark," \
                   "customer_type," \
                   "rm_num_key," \
                   "resend_timestamp" \
                   ")" \
                   "VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"

    def Read_Excel(self):
        # 创建数据库链接
        conn = pymysql.connect(host=self.HOST, port=self.PORT, user=self.USER, password=self.PASSWORD,
                               db='openbank_channel',
                               charset='utf8')
        cur = conn.cursor()
        book = xlrd.open_workbook("work.csv")
        # 打开需要导入数据库的excel表
        sheet = book.sheet_by_name("Sheet1")
        for r in range(1, sheet.nrows):
            consent_id = sheet.cell(r, 0).value
            Permission = sheet.cell(r, 1).value
            List = sheet.cell(r, 2).value
            expiration_date = sheet.cell(r, 3).value
            transaction_from_date = sheet.cell(r, 4).value
            transaction_to_date = sheet.cell(r, 5).value
            creation_date = sheet.cell(r, 6).value
            refresh_period_date = sheet.cell(r, 7).value
            status = sheet.cell(r, 8).value
            status_update_date = sheet.cell(r, 9).value
            account_id = sheet.cell(r, 10).value
            rm_num_key = sheet.cell(r, 11).value
            tsp_user_id = sheet.cell(r, 12).value
            org = sheet.cell(r, 13).value
            org_id = sheet.cell(r, 14).value
            org_title = sheet.cell(r, 15).value
            id = Consentid_UUID(consent_id)
            update_by = 'opiuser'
            update_date = datetime.datetime()
            create_by = 'opiuser'
            create_date = datetime.datetime()
            reconcile_date = datetime.datetime()
            reconcile_result = 2
            reconcile_comments = 2
            customer_type = 1
            remark = ''
            resend_timestamp = datetime.datetime()
            value = (id, consent_id, expiration_date, refresh_period_date, transaction_from_date, creation_date, status,
                     status_update_date, tsp_user_id, org, org_id, org_title, update_by, update_date, create_by,
                     reconcile_date, reconcile_result, reconcile_comments, remark, customer_type, rm_num_key,
                     resend_timestamp)
            cur.execute(self.query, value)
        cur.close()
        conn.commit()
        conn.close()
        columns = str(sheet.ncols)
        rows = str(sheet.nrows)


if __name__ == '__main__':
    a = ReadExcel()
    a.Read_Excel()
