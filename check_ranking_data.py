import os
import re
import csv
import json
import datetime
import requests
import gspread
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
def getLatestDownloadedDirName(downloadsDirPath):
    if len(os.listdir(downloadsDirPath)) == 0:
        return None
    return max (
        [downloadsDirPath + '/' + f for f in os.listdir(downloadsDirPath)],
        key=os.path.getctime
    )

def sendChatworkNotification(message):
    try:
        url = f'https://api.chatwork.com/v2/rooms/{os.environ["CHATWORK_ROOM_ID"]}/messages'
        headers = { 'X-ChatWorkToken': os.environ["CHATWORK_API_TOKEN"] }
        params = { 'body': message }
        requests.post(url, headers=headers, params=params)
    except Exception as err:
        logger.error(f'Error: sendChatworkNotification: {err}')
        exit(1)

def num2alpha(num):
    if num<=26:
        return chr(64+num)
    elif num%26==0:
        return num2alpha(num//26-1)+chr(90)
    else:
        return num2alpha(num//26)+chr(64+num%26)

### Google ###
def getRankingCsvData(csvPath):
    with open(csvPath, newline='', encoding='utf-8') as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        next(buf)
        for row in buf:
            yield row

def checkRankingData(folder, datas, message):
    try:
        SPREADSHEET_ID = os.environ['RANK_DATA_SSID']
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
        gc = gspread.authorize(credentials)
        sheet = gc.open_by_key(SPREADSHEET_ID).worksheet('Rank Data')

        for data in datas:
            rdate = datetime.datetime.strptime(data[7], '%b %d, %Y').strftime('%Y/%m/%d')
            if rdate != today.strftime('%Y/%m/%d'):
                logger.debug(f'checkRankingData: {folder}: NG')
                return False
        
        return True
    except Exception as err:
        logger.debug(f'Error: checkUploadData: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':

    try:
        rankDataDirPath = os.environ["RANK_DATA_DIR"]
        dateDirPath = getLatestDownloadedDirName(rankDataDirPath)
        message = f'[info][title]本日の計測結果 [@{today.strftime("%H:%M")}][/title]'

        if dateDirPath != f'{rankDataDirPath}/{today.strftime("%Y-%m-%d")}':
            message += "本日のデータフォルダが生成されておりません。\n担当者は本日の順位計測に問題がないかご確認ください。[/info]"
            sendChatworkNotification(message)
            logger.debug(f'check_ranking_data: 本日のデータフォルダが生成されておりません。')
            exit(0)

        config = configparser.ConfigParser()
        config.read('clientInfo.ini')
        folders = config.sections()

        for folder in folders:
            if folder == "aimplace.co.jp":
                continue
            datas = list(getRankingCsvData(f'{dateDirPath}/{folder}.txt'))
            if checkRankingData(folder, datas, message):
                message += f'{folder} ✅\n'
            else:
                message += f'{folder} 🔥\n'
        
        message += '[/info]'
        sendChatworkNotification(message)
        logger.info("check_ranking_data: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'check_ranking_data: {err}')
        exit(1)
