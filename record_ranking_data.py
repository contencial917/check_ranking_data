import os
import re
import csv
import json
import datetime
import requests
import gspread
import codecs
import configparser
from time import sleep
from oauth2client.service_account import ServiceAccountCredentials

# Logger setting
from logging import getLogger, FileHandler, DEBUG
logger = getLogger(__name__)
today = datetime.datetime.now()
os.makedirs('./log', exist_ok=True)
handler = FileHandler(f'log/{today.strftime("%Y-%m-%d")}_result.log', mode='a')
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False

### functions ###
def sendChatworkNotification(message):
    try:
        url = f'https://api.chatwork.com/v2/rooms/{os.environ["CHATWORK_ROOM_ID"]}/messages'
        headers = { 'X-ChatWorkToken': os.environ["CHATWORK_API_TOKEN"] }
        params = { 'body': message }
        requests.post(url, headers=headers, params=params)
    except Exception as err:
        logger.error(f'Error: sendChatworkNotification: {err}')
        exit(1)

def getRankingCsvData(csvPath):
    with open(csvPath, newline='', encoding='utf-8') as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        for row in buf:
            yield row

def recordRankingData(project, datas, sheet):
    try:
        global append_list
        global cnt_append_row
        global last_row_num
        ja = 0
        sh = 0
        mb = 0
        dja = 0
        dsh = 0
        dmb = 0

        for index, data in enumerate(datas):
            if index == 0:
                for i, d in enumerate(data):
                    if re.search('Mobile', d) and re.search('Rank', d):
                        ja = int(i)
                    elif re.search('Shibuya', d) and re.search('Rank', d):
                        sh = int(i) 
                    elif re.search('Japan', d) and re.search('Rank', d):
                        mb = int(i)
                    elif re.search('Mobile', d) and re.search('Date', d):
                        dja = int(i)
                    elif re.search('Shibuya', d) and re.search('Date', d):
                        dmb = int(i)
                    elif re.search('Japan', d) and re.search('Date', d):
                        dsh = int(i)
                continue

            url = 'https://docs.google.com/spreadsheets/d/1LyDPP7Nz0WYm4PnuqOxygG_x3yRGMEzV31eu6JlX2ys/edit'
            cnt_append_row += 1
            row = last_row_num + cnt_append_row

            name = f'=FILTER(IMPORTRANGE("{url}", "サイト一覧!C$2:C$300"), IMPORTRANGE("{url}", "サイト一覧!D$2:D$300")=B{row})'
            rja = datetime.datetime.strptime(data[dja], '%b %d, %Y').strftime('%Y/%m/%d')
            rsh = datetime.datetime.strptime(data[dsh], '%b %d, %Y').strftime('%Y/%m/%d')
            rmb = datetime.datetime.strptime(data[dmb], '%b %d, %Y').strftime('%Y/%m/%d')
            result = [name, project, data[0], data[ja], rja, data[sh], rsh, data[mb], rmb]
            append_list.extend(result)
        
        return True
    except Exception as err:
        logger.debug(f'Error: recordRankingData: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':

    try:
        rankDataDirPath = os.environ["RANK_DATA_DIR"]
        dateDirPath = f'{rankDataDirPath}/{today.strftime("%Y-%m-%d")}'

        config = configparser.ConfigParser()
        config.read_file(codecs.open("clientInfo.ini", "r", "utf8"))
        projects = config.sections()

        SPREADSHEET_ID = os.environ['RANK_DATA_SSID']
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
        gc = gspread.authorize(credentials)
        sheet = gc.open_by_key(SPREADSHEET_ID).worksheet('Rank Data')
        append_list = []
        cnt_append_row = 0
        last_row_num = len([i for i in sheet.col_values(1) if i])

        for project in projects:
            if project == 'aimplace.co.jp':
                continue
            datas = list(getRankingCsvData(f'{dateDirPath}/{project}.txt'))
            recordRankingData(project, datas, sheet)
            logger.debug(f'recordRankingData: {project}')
        
        print(cnt_append_row)

        sheet.add_rows(cnt_append_row)
        cell_list = sheet.range(f'A{last_row_num + 1}:I{last_row_num + cnt_append_row}')
        for index, cell in enumerate(cell_list):
            cell.value = append_list[index]
        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

        last_row_num = len([i for i in sheet.col_values(1) if i])
        sheet.set_basic_filter(name=(f'A1:I{last_row_num}'))

        message = '[info][title]順位計測データ取込[/title]\n'
        message += '本日の順位計測データを取り込みました。\n'
        message += '下記リンクから順位計測データをご確認ください。\n\n'
        message += 'https://docs.google.com/spreadsheets/d/16WCFv8z9ufGtNl_F6I8r9noj8ernC6HQwJwHSc9C3ck/edit?usp=sharing \n'
        message += '[/info]'

        sendChatworkNotification(message)
 
        logger.info("record_ranking_data: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'record_ranking_data: {err}')
        exit(1)
