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
#from gspread_formatting import *

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

def recordRankingData(datas, sheet, day):
    try:
        key = 0
        ja = 0
        sh = 0
        row_count = sheet.row_count

        keywords = sheet.col_values(1)
        ranking = sheet.range(1, day + 1, row_count, day + 1)
        for index, data in enumerate(datas):
            if index == 0:
                for i, d in enumerate(data):
                    if re.search('Keyword|キーワード', d):
                        key = int(i)
                    elif re.search('Japan', d) and re.search('Rank|ランキング', d):
                        ja = int(i)
                    elif re.search('Shibuya', d) and re.search('Rank|ランキング', d):
                        sh = int(i)
                continue
            for i, d in enumerate(keywords):
                if d == data[key]:
                    if data[sh] == 'トップ圏外 100':
                        ranking[i].value = '-'
                    else:
                        ranking[i].value = data[sh]
                    break
        sheet.update_cells(ranking, value_input_option="USER_ENTERED")
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

        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
        gc = gspread.authorize(credentials)

        for project in projects:
            SPREADSHEET_ID = config[project]['SSID']
            sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(today.strftime("%Y%m"))
            datas = list(getRankingCsvData(f'{dateDirPath}/{project}.txt'))
            recordRankingData(datas, sheet, int(today.strftime("%d")))
#            last_row_num = len(list(sheet.col_values(1)))
#            format_cell_range(sheet, f'A2:AF{last_row_num}', cellFormat(horizontalAlignment='CENTER'))
            logger.debug(f'recordRankingData: {project}')
            sleep(2)
        
        message = '[info][title]順位計測データ取込[/title]\n'
        message += '本日の順位計測データを取り込みました。\n'
        message += '下記リンクから順位計測データをご確認ください。\n\n'
        message += 'https://drive.google.com/drive/folders/1AV70yHGUsYkkbOu1MdxsEzig2aFmYaTb?usp=sharing'
        message += '[/info]'

        sendChatworkNotification(message)
 
        logger.info("record_ranking_data: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'record_ranking_data: {err}')
        exit(1)
