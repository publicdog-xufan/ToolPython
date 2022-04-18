import pandas as pd
import pymysql


def Datetime_Format_Conversion(DateKey):
    NewDateKey = pd.to_datetime(DateKey).dt.strftime('%Y-%m-%dT %H:%M:%SZ')
    return NewDateKey


class ReconciliationDocument(object):
    def __init__(self):
        # 数据库连接参数
        self.HOST = '127.0.0.1'
        self.PORT = 3306
        self.USER = 'root'
        self.PASSWORD = 'ozJOgV17Bey'

    def Add_Value(self, BankConsentId):
        # 创建数据库链接
        conn = pymysql.connect(host=self.HOST, port=self.PORT, user=self.USER, password=self.PASSWORD, db='openbank_p3',
                               charset='utf8')
        # 生成sql语句 利用传入的BankConsentId进行条件查询
        authorize_sql = "SELECT " \
                        "ai.consent_id," \
                        "ai.expiration_date," \
                        "ai.transaction_from_date," \
                        "ai.transaction_to_date," \
                        "ai.creation_date," \
                        "ai.refresh_period_date," \
                        "ai.`status`," \
                        "ai.status_update_date," \
                        "aa.account_id," \
                        "ai.rm_num_key," \
                        "ai.tsp_user_id," \
                        "ai.org," \
                        "ai.org_id," \
                        "ai.org_title " \
                        "FROM authorize_info ai " \
                        "INNER JOIN authorize_account aa " \
                        "ON ai.consent_id = aa.consent_id " \
                        "WHERE ai.bank_consent_id > %s;" % BankConsentId

        # pandas保存数据
        data = pd.read_sql(authorize_sql, con=conn)
        # 修改时间格式，从 yyyy/mm/dd hh:mm:ss 修改为yyyy-mm-ddT hh:mm:ssZ
        data["expiration_date"] = Datetime_Format_Conversion(data["expiration_date"])
        data["transaction_from_date"] = Datetime_Format_Conversion(data["transaction_from_date"])
        data["transaction_to_date"] = Datetime_Format_Conversion(data["transaction_to_date"])
        data["creation_date"] = Datetime_Format_Conversion(data["creation_date"])
        data["refresh_period_date"] = Datetime_Format_Conversion(data["refresh_period_date"])
        data["status_update_date"] = Datetime_Format_Conversion(data["status_update_date"])

        # 拼接授权类型（授权类型目前假数据均是四种类型，此部分直接取四类型进行拼接）
        # loc标识插入新列的位置， column为列名 value 为值
        data.insert(loc=1, column="Permission List", value="ReadAccountAvailability|ReadAccountBalance"
                                                           "|ReadAccountStatus|ReadAccountTransaction")
        # 默认直接在末尾添加一列
        # data["Permission List"] = "ReadAccountAvailability|ReadAccountBalance|ReadAccountStatus
        # |ReadAccountTransaction"
        conn.close()
        return data


if __name__ == '__main__':
    a = ReconciliationDocument()
    b = a.Add_Value(100000)
    print("******** 开始执行 ********")
    print(b)
    b.to_csv('work.csv', index=False)

    print("******** 输入完成 ********")
