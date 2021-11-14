'''
此程式可以搜尋想要的股票資訊，並且將資料匯入csv file裏面
可搜索的資訊包含
1. 查詢股票代號或是公司名稱
2. 下載資料到數據檔案裡
3. 某日的股票訊息
4. 某月的股票訊息，並且繪製圖表
5. 某年的股票訊息，並且繪製圖表
6. 一段時間的股票訊息，並且繪製圖表
7. 儲存到excel

儲存到mysql在stocks_mySQL.py裏面
'''


import sys
import stocks_func as sf
import os
import pyinputplus as pyip
import calendar
import datetime
import send2trash
from os.path import expanduser


STOCK_FILE_DIR = './stocks_data_file'
HEAD_URL = 'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date='
TAIL_URL = '&stockNo='
EXCEL_SAVE_PATH = expanduser('~') + '/Desktop/'


while True:
    print('=' * 50)
    input_menu = pyip.inputMenu(['利用代號查詢公司名稱',
                                 '將資料儲存到csv檔案',
                                 '查詢個股某日的資料',
                                 '查訊個股某年某月份的資料以及繪製圖表',
                                 '查詢某個股某年份的資料以及繪製圖表',
                                 '查詢個股某段時間的資料以及繪製圖表',
                                 '刪除資料檔案',
                                 '離開程式'],
                                 prompt='請選擇下列服務\n', numbered=True)

    if input_menu == '利用代號查詢公司名稱':
        print('='*50)
        STOCK_NUM_OR_NAME = pyip.inputStr('請輸入想尋找的股票代號或是公司名稱: ')
        sf.search_stocks(STOCK_NUM_OR_NAME)
        sf.exist_system()

    elif input_menu == '將資料儲存到csv檔案':
        print('='*50)
        print('*** 由於網路限制，最多輸入兩年資料 ***')
        print('*** 如果只需要新增一年的資料，請將開始年份以及結束年份輸入同一年 ***')
        # 設定最大年份為今年，獲取現在的年份
        year_now = datetime.datetime.now().year
        STOCK_NUM = pyip.inputStr('請輸入想儲存的股票代號: ')
        START_YEAR = pyip.inputNum('請輸入想儲存的開始年份: ', min=2000, max=year_now)
        END_YEAR = pyip.inputNum('請輸入想儲存的結束年份: ', min=START_YEAR, lessThan=START_YEAR+2)
        if END_YEAR > year_now:
            print('*** 查詢年份超過今年年份，請重新操作 ***')
        # 初始化檔案路徑
        CSV_FILE_PATH = ''
        for folderName, subfolder, fileNames in os.walk(STOCK_FILE_DIR):
            for fileName in fileNames:
                if fileName.endswith('.csv') and len(fileName) >= 18 and fileName[:-14].__contains__(STOCK_NUM):
                    CSV_FILE_PATH = STOCK_FILE_DIR + '/' + fileName
        START_DATE = str(START_YEAR) + '0101'
        END_DATE = str(END_YEAR) + '1231'
        # 如果有檔案存在，檢查檔案是有我們想要的資料
        isInfoExists = sf.is_info_exists(CSV_FILE_PATH, START_DATE, END_DATE)
        if isInfoExists == 5:
            continue
        # 開始下載檔案
        sf.download_stocks_data(isInfoExists, STOCK_FILE_DIR, STOCK_NUM, START_YEAR, END_YEAR, CSV_FILE_PATH)
        # 選擇是否要繼續服務或是離開
        sf.exist_system()

    elif input_menu == '查詢個股某日的資料':
        print('=' * 50)
        print('*** 如果查詢的日期為今年的資料且查無資料，請重新回到主選單，選擇2，將今年資料更新到最新，如再無資料，代表當天未開盤 ***')
        STOCK_NUM = pyip.inputStr('請輸入想查詢的股票代號: ')
        date = pyip.inputDate('請輸入想查詢的年份月份日期 (YYYYMMDD)\n',formats=['%Y%m%d'])
        START_DATE = datetime.datetime.strftime(date, '%Y%m%d')

        # 初始化檔案路徑
        CSV_FILE_PATH = ''
        for folderName, subfolder, fileNames in os.walk(STOCK_FILE_DIR):
            for fileName in fileNames:
                if fileName.endswith('.csv') and len(fileName) >= 18 and fileName[:-14].__contains__(STOCK_NUM):
                    CSV_FILE_PATH = STOCK_FILE_DIR + '/' + fileName

        sf.get_data_by_date(CSV_FILE_PATH, START_DATE)
        sf.exist_system()

    elif input_menu == '查訊個股某年某月份的資料以及繪製圖表':
        print('=' * 50)
        STOCK_NUM = pyip.inputStr('請輸入想查詢的股票代號: ')
        startMonth = pyip.inputDate('請輸入想查詢的年份月份 (YYYYMM)\n', formats=['%Y%m'])
        year_month = datetime.datetime.strftime(startMonth, '%Y%m')
        START_YEAR = year_month[:4]
        START_MONTH = year_month[4:]
        START_DATE = year_month + '01'
        END_DATE = year_month + str(calendar.monthrange(int(START_YEAR), int(START_MONTH))[1])

        # 初始化檔案路徑
        CSV_FILE_PATH = ''
        for folderName, subfolder, fileNames in os.walk(STOCK_FILE_DIR):
            for fileName in fileNames:
                if fileName.endswith('.csv') and len(fileName) >= 18 and fileName[:-14].__contains__(STOCK_NUM):
                    CSV_FILE_PATH = STOCK_FILE_DIR + '/' + fileName

        # 如果有檔案存在，檢查檔案是有我們想要的資料
        isInfoExists = sf.is_info_exists(CSV_FILE_PATH, START_DATE, END_DATE)

        # 是否要下載檔案
        if isInfoExists == 5:
            continue
        elif isInfoExists != 1:
            print('*** 資料檔案中無資料或是尚未更新 ***')
            response = pyip.inputYesNo('要下載資料到資料檔案裡面嗎？(y/n)')
            if response == 'no':
                continue
            else:
                CSV_FILE_PATH = sf.download_stocks_data(isInfoExists, STOCK_FILE_DIR, STOCK_NUM, START_YEAR, START_YEAR, CSV_FILE_PATH)

        # 選擇以下想要的服務
        save_print_option = pyip.inputMenu(['顯示在畫面','儲存到excel','顯示並處存'],
                                           prompt='要將資料顯示出來還是儲存到excel裡面？\n',
                                           numbered=True)
        # 根據想要的內容讀取檔案
        df = sf.get_data_by_date(CSV_FILE_PATH, START_DATE, END_DATE)
        if save_print_option == '顯示在畫面':
            sf.plot_dynamic_chart(df, STOCK_NUM)
        elif save_print_option == '儲存到excel':
            sf.save_to_excel(df, STOCK_NUM, START_DATE, END_DATE, EXCEL_SAVE_PATH)
        else:
            sf.plot_dynamic_chart(df, STOCK_NUM)
            sf.save_to_excel(df, STOCK_NUM, START_DATE, END_DATE, EXCEL_SAVE_PATH)
        sf.exist_system()

    elif input_menu == '查詢某個股某年份的資料以及繪製圖表':
        print('=' * 50)
        STOCK_NUM = pyip.inputStr('請輸入想查詢的股票代號: ')
        startMonth = pyip.inputDate('請輸入想查詢的年份(YYYY)\n', formats=['%Y'])
        START_YEAR = startMonth.year
        START_DATE = str(START_YEAR) + '0101'
        END_DATE = str(START_YEAR) + '1231'

        # 初始化檔案路徑
        CSV_FILE_PATH = ''
        for folderName, subfolder, fileNames in os.walk(STOCK_FILE_DIR):
            for fileName in fileNames:
                if fileName.endswith('.csv') and fileName.__contains__(STOCK_NUM) and len(fileName) == 18:
                    CSV_FILE_PATH = STOCK_FILE_DIR + '/' + fileName

        # 如果有檔案存在，檢查檔案是有我們想要的資料
        isInfoExists = sf.is_info_exists(CSV_FILE_PATH, START_DATE, END_DATE)

        # 開始下載檔案
        if isInfoExists == 5:
            continue
        elif isInfoExists != 1:
            print('*** 資料並未儲存在資料檔案裡面 ***')
            response = pyip.inputYesNo('要下載資料到資料檔案裡面嗎？(y/n)')
            if response == 'no':
                continue
            else:
                CSV_FILE_PATH = sf.download_stocks_data(isInfoExists, STOCK_FILE_DIR, STOCK_NUM, START_YEAR, START_YEAR,
                                                        CSV_FILE_PATH)
        # 選擇以下動作
        save_print_option = pyip.inputMenu(['顯示在畫面','儲存到excel','顯示並處存'],
                                           prompt='要將資料顯示出來還是儲存到excel裡面？\n',
                                           numbered=True)
        df = sf.get_data_by_date(CSV_FILE_PATH, START_DATE, END_DATE)
        if save_print_option == '顯示在畫面':
            sf.plot_dynamic_chart(df, STOCK_NUM)
        elif save_print_option == '儲存到excel':
            sf.save_to_excel(df, STOCK_NUM, START_DATE, END_DATE, EXCEL_SAVE_PATH)
        else:
            sf.plot_dynamic_chart(df, STOCK_NUM)
            sf.save_to_excel(df, STOCK_NUM, START_DATE, END_DATE, EXCEL_SAVE_PATH)
        sf.exist_system()

    elif input_menu == '查詢個股某段時間的資料以及繪製圖表':
        print('=' * 50)
        STOCK_NUM = pyip.inputStr('請輸入想查詢的股票代號: ')
        startDate = pyip.inputDate('請輸入想查詢的開始年份月份日期(YYYYMMDD)\n', formats=['%Y%m%d'])
        endDate = pyip.inputDate('請輸入想查詢的結束年份月份日期(YYYYMMDD)\n', formats=['%Y%m%d'])

        START_DATE = datetime.datetime.strftime(startDate, '%Y%m%d')
        END_DATE = datetime.datetime.strftime(endDate, '%Y%m%d')
        # 初始化檔案路徑
        CSV_FILE_PATH = ''
        for folderName, subfolder, fileNames in os.walk(STOCK_FILE_DIR):
            for fileName in fileNames:
                if fileName.endswith('.csv') and len(fileName) >= 18 and fileName[:-14].__contains__(STOCK_NUM):
                    CSV_FILE_PATH = STOCK_FILE_DIR + '/' + fileName

        # 如果有檔案存在，檢查檔案是有我們想要的資料
        isInfoExists = sf.is_info_exists(CSV_FILE_PATH, START_DATE, END_DATE)
        # 如果資料不存在，會到第一項
        if isInfoExists == 5:
            continue
        elif isInfoExists != 1:
            print('*** 資料並未儲存在資料檔案裡面，請回到主選單，選擇第2項，將資料加入資料檔案裡 ***\n')
            continue

        save_print_option = pyip.inputMenu(['顯示在畫面', '儲存到excel', '顯示並處存'],
                                           prompt='要將資料顯示出來還是儲存到excel裡面？\n',
                                           numbered=True)
        df = sf.get_data_by_date(CSV_FILE_PATH, START_DATE, END_DATE)
        if save_print_option == '顯示在畫面':
            sf.plot_dynamic_chart(df, STOCK_NUM)
        elif save_print_option == '儲存到excel':
            sf.save_to_excel(df, STOCK_NUM, START_DATE, END_DATE, EXCEL_SAVE_PATH)
        else:
            sf.plot_dynamic_chart(df, STOCK_NUM)
            sf.save_to_excel(df, STOCK_NUM, START_DATE, END_DATE, EXCEL_SAVE_PATH)
        sf.exist_system()

    elif input_menu == '刪除資料檔案':
        print('=' * 50)
        print('*** 此選項適用於資料檔案有誤、資料檔案不齊全、資料檔案出現錯誤選項等等，可以刪除資料檔案，再重新下載 ***')
        csv_file_lists = []
        for folderName, subfolder, fileNames in os.walk(STOCK_FILE_DIR):
            for fileName in fileNames:
                if fileName.endswith('.csv') and len(fileName) >= 18 and fileName[-9] == '_' and fileName[-14] == '_':
                    CSV_FILE_PATH = STOCK_FILE_DIR + '/' + fileName
                    csv_file_lists.append(CSV_FILE_PATH)
        csv_file_lists.append('離開')
        input_menu = pyip.inputMenu(csv_file_lists, prompt='請選擇要刪除的檔案\n', numbered=True)
        if input_menu == '離開':
            print('離開')
        else:
            send2trash.send2trash(input_menu)
            print('已刪除 ' + input_menu)
        sf.exist_system()


    elif input_menu == '離開程式':
        sys.exit()




