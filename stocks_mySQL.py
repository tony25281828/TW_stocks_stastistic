import datetime
import sys
import pymysql.cursors
import pandas as pd
import pyinputplus as pyip
import os


# 新增table
def create_table(table_name, mysql_cursor):
    sql_create_tb = f'''
        CREATE TABLE IF NOT EXISTS {table_name}(
            日期 DATE NOT NULL PRIMARY KEY,
            成交股數 INT,
            成交金額 BIGINT,
            開盤價 FLOAT,
            最高價 FLOAT,
            最低價 FLOAT,
            收盤價 FLOAT,
            漲跌價差 FLOAT,
            成交筆數 INT
            ); 
        '''

    mysql_cursor.execute(sql_create_tb)



# 將資料上傳到mysql
def insert_data(csv_file_path, table_name, mysql_cursor):
    data = pd.read_csv(csv_file_path)
    total_rows_num = data.index.stop

    for row in range(total_rows_num):
        date = str(data['日期'][row])
        trading_volume = int(data['成交股數'][row])
        turnover = int(data['成交金額'][row])
        opening_price = float(data['開盤價'][row])
        highest_price = float(data['最高價'][row])
        lowest_price = float(data['最低價'][row])
        closing_price = float(data['收盤價'][row])
        if data['漲跌價差'][row].__contains__('-'):
            price_difference = 0 - float(data['漲跌價差'][row].replace('-', ''))
        elif data['漲跌價差'][row].__contains__('X'):
            price_difference = float(data['漲跌價差'][row].replace('X',''))
        else:
            price_difference = float(data['漲跌價差'][row])
        transactions = int(data['成交筆數'][row])

        sql_insert_data = f'''
            INSERT IGNORE INTO {table_name} (日期, 成交股數, 成交金額, 開盤價, 最高價, 最低價, 收盤價, 漲跌價差, 成交筆數) 
            VALUES ("{date}", {trading_volume}, {turnover}, {opening_price}, {highest_price}, {lowest_price}, {closing_price}, {price_difference}, {transactions})
        '''

        mysql_cursor.execute(sql_insert_data)



# 刪除表格
def drop_table(mysql_cursor):
    show_tables = '''SHOW TABLES;'''
    mysql_cursor.execute(show_tables)
    tables = mysql_cursor.fetchall()
    for tableNum in range(len(tables)):
        print(str(tableNum+1)+'. ' + tables[tableNum]['Tables_in_tw_stocks'])
    input_table = pyip.inputStr('請輸入想要刪除的表格\n。如果想離開請輸入 "exit"\n')
    if input_table == 'exit':
        sys.exit()
    confirm = pyip.inputYesNo('請確認要刪除的表格是'+ input_table + '(y/n)\n')
    if confirm == 'yes':
        drop_table = f'''DROP TABLE {input_table}'''
        try:
            mysql_cursor.execute(drop_table)
        except:
            print('資料庫裡無此表格')
    else:
        sys.exit()




def exist_system():
    input = pyip.inputYesNo('要繼續服務(y)或是離開(n)?\n')
    if input == 'no':
        sys.exit()




print('''使用此服務之前請先開啟mySQL Server的連線''')
print('='*100)

input_menu = pyip.inputMenu(['建立表格並且上傳資料','僅建立表格','僅上傳資料','刪除表格','輸入MySQL指令','離開程式'],
                               prompt='請選擇想要的服務\n', numbered=True)
if input_menu != '離開程式':
    try:
        password = input('請輸入密碼')
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
        print('無法建立mySQL連線')
        sys.exit()
    except ConnectionRefusedError:
        print('連線被拒絕')
        sys.exit()
    cr = conn.cursor()

if input_menu == '建立表格並且上傳資料':
    CSV_FILE_NAME = pyip.inputStr('請輸入檔案名稱或是絕對路徑:\n')
    if not os.path.exists(CSV_FILE_NAME):
        print('找不到此檔案')
        sys.exit()
    csv_file_basename = os.path.basename(CSV_FILE_NAME)
    STOCK_NUM = csv_file_basename[:4]
    TABLE_NAME = 'stocks_' + STOCK_NUM
    create_table(TABLE_NAME, cr)
    insert_data(CSV_FILE_NAME, TABLE_NAME, cr)
    conn.commit()
    print('已成功建立表格及將資料上傳到MySQL')

elif input_menu == '僅建立表格':
    CSV_FILE_NAME = pyip.inputStr('請輸入檔案名稱或是絕對路徑\n')
    if not os.path.exists(CSV_FILE_NAME):
        print('找不到此檔案')
        sys.exit()
    csv_file_basename = os.path.basename(CSV_FILE_NAME)
    STOCK_NUM = csv_file_basename[:4]
    STOCK_NUM = CSV_FILE_NAME[:4]
    TABLE_NAME = 'stocks_' + STOCK_NUM
    create_table(TABLE_NAME, cr)
    print('已成功建立表格')

elif input_menu == '僅上傳資料':
    CSV_FILE_NAME = pyip.inputStr('請輸入檔案名稱或是絕對路徑\n')
    if not os.path.exists(CSV_FILE_NAME):
        print('找不到此檔案')
        sys.exit()

    csv_file_basename = os.path.basename(CSV_FILE_NAME)
    STOCK_NUM = csv_file_basename[:4]
    #STOCK_NUM = CSV_FILE_NAME[:4]
    TABLE_NAME = 'stocks_' + STOCK_NUM
    print(STOCK_NUM)
    print(TABLE_NAME)
    insert_data(CSV_FILE_NAME, TABLE_NAME, cr)
    conn.commit()
    print('已成功將檔案上傳到MySQL')

elif input_menu == '刪除表格':
    drop_table(cr)

elif input_menu == '輸入MySQL指令':
    print('=== 使用MySQL語言來與MySQL互動 ===')
    while True:
        mysql_command = input('請輸入MySQL指令: \n')
        if mysql_command.lower().__contains__('select') or \
            mysql_command.lower().__contains__('describe') or \
             mysql_command.lower().__contains__('show'):
            cr.execute(mysql_command)
            all_data = cr.fetchall()
            for data in all_data:
                print(data)
        else:
            cr.execute(mysql_command)
            conn.commit()
        exist_system()

elif input_menu == '離開程式':
    sys.exit()

conn.close()



