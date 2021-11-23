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
def checkLatestDownloadedDirName(downloadsDirPath):
    if len(os.listdir(downloadsDirPath)) == 0:
        return None
    for f in os.listdir(downloadsDirPath):
        if f == today.strftime("%Y-%m-%d"):
            return True
    return False

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
        next(buf)
        for row in buf:
            yield row

def checkRankingData(folder, datas):
    try:
        for data in datas:
            rdate = datetime.datetime.strptime(data[7], '%b %d, %Y').strftime('%Y/%m/%d')
            if rdate != today.strftime('%Y/%m/%d'):
                logger.debug(f'checkRankingData: {folder}: NG')
                return False
        
        return True
    except Exception as err:
        logger.debug(f'Error: checkRankingData: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':

    try:
        rankDataDirPath = os.environ["RANK_DATA_DIR"]
        dateDirPath = f'{rankDataDirPath}/{today.strftime("%Y-%m-%d")}'

        if not checkLatestDownloadedDirName(rankDataDirPath):
            message = f'[info][title]本日の順位計測結果 [@{today.strftime("%H:%M")}][/title]'
            message += "本日のデータフォルダが生成されておりません。\n担当者は本日の順位計測に問題がないかご確認ください。[/info]"
            sendChatworkNotification(message)
            logger.debug(f'check_ranking_data: 本日のデータフォルダが生成されておりません。')
            exit(0)

        config = configparser.ConfigParser()
        config.read_file(codecs.open("clientInfo.ini", "r", "utf8"))
        folders = config.sections()
        NGprojects = []

        for folder in folders:
            datas = list(getRankingCsvData(f'{dateDirPath}/{folder}.txt'))
            if not checkRankingData(folder, datas):
                NGprojects.append(folder)

        total = len(folders)
        ng = len(NGprojects)
        if ng == 0:
            message = f'[info][title]【本日の順位計測結果】@ {today.strftime("%H:%M")}[/title]'
            message += f'[ {total} / {total} ] 完了 ✅\n'
            message += 'パーフェクトです。[/info]'
        else:
            message = f'[info][title]【本日の順位計測結果】@ {today.strftime("%H:%M")}[/title]'
            message += f'[ {total - ng} / {total} ] 完了 🔥\n'
            message += f'担当者は再計測対応を行ってください。\n\n'
            message += ' ▼ 未計測プロジェクト一覧\n'
            message += ' 🔥\n'.join(NGprojects)
            message += ' 🔥\n'
            message += '[/info]'

        sendChatworkNotification(message)
        logger.info("check_ranking_data: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'check_ranking_data: {err}')
        exit(1)
