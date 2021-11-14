import pymysql.cursors
import pandas as pd
import sys
import os


def create_table(table_name, mysql_cursur):
    sql_create_tb = f'''
        CREATE TABLE IF NOT EXISTS {table_name}(
            股票代號 varchar(20) NOT NULL PRIMARY KEY,
            公司名稱 varchar(20) NOT NULL
    )
    '''
    mysql_cursur.execute(sql_create_tb)
    
    
    

def insert_data(csv_file_path, table_name, mysql_cursor):
    data = pd.read_csv(csv_file_path)
    total_rows_num = data.index.stop
    
    for row in range(total_rows_num):
        stock_num = str(data['股票代號'][row].replace('/',''))
        stock_name = str(data['公司名稱'][row])
        
        sql_insert_data = "INSERT IGNORE INTO %s (股票代號, 公司名稱) VALUES (\'%s\',\'%s\')" %(table_name,stock_num,stock_name)
        
        print(sql_insert_data)
        
        mysql_cursor.execute(sql_insert_data)


password = input('請輸入密碼\n')
try:
    conn = pymysql.connect(
        host='127.0.0.1',
        port=3306,
        user='root',
        password=password,
        charset='utf8',
        database='TW_stocks',
        cursorclass=pymysql.cursors.DictCursor
    )
except pymysql.err.OperationalError:
    print('無法建立MySQL連線')
    sys.exit()
except ConnectionRefusedError:
    print('連線被拒絕')
    sys.exit()

cr = conn.cursor()
create_table('stocks_data', cr)
dir = './stocks_data_file'
csv_file_name = ''
for folder, subfolder, fileNames in os.walk(dir):
    for fileName in fileNames:
        if fileName.__contains__('stocks_data') and fileName.endswith('.csv'):
            csv_file_name = fileName

insert_data(csv_file_name, 'stocks_data', cr)
conn.commit()
conn.close()

