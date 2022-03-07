import os
import re
import csv
import sys
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
def getLatestDownloadedFileName(downloadsDirPath):
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

def getRankingCsvData(csvPath):
    with open(csvPath, newline='', encoding='utf-8') as csvfile:
        buf = csv.reader(csvfile, delimiter=',', lineterminator='\r\n', skipinitialspace=True)
        for row in buf:
            yield row

def checkRankingData(datas):
    try:
        date = -1
        for index, data in enumerate(datas):
            if index == 0:
                for i, d in enumerate(data):
                    if re.search('Keyword|キーワード', d):
                        key = int(i)
                    elif re.search('Shibuya', d):
                        if re.search('Rank|ランキング', d):
                            rank = int(i)
                        elif re.search('URL', d):
                            url = int(i)
                    elif re.search('Date|日付', d):
                        date = int(i)
                continue

            if data[rank].isdecimal() and int(data[rank]) <= 10:
                yield [data[key], data[rank], data[url], data[date]]
    except Exception as err:
        logger.debug(f'Error: checkRankingData: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':
    if len(sys.argv) > 1:
        domain = sys.argv[1]
    else:
        logger.debug("Error: No parameter")
        exit(1)

    try:
        rankDataDirPath = os.environ["RANK_DATA_DIR"]
        folderPath = getLatestDownloadedFileName(rankDataDirPath)
        dataFilePath = f'{folderPath}/{domain}.txt'

        datas = list(getRankingCsvData(dataFilePath))
        result = list(checkRankingData(datas))

        date = folderPath.replace(f'{rankDataDirPath}/', '')
        if len(result) == 0:
            message = f'[info][title]【{domain}】10位以内アラート@{date}[/title]'
            message += '順位が10以内のKWはありません。[/info]'
        else:
            message = f'[info][title]【{domain}】10位以内アラート@{date}[/title]'
            message += '下記のKWが10位以内にランクインしました！\n'
            for e in result:
                message += f'\n＋＋＋\n\n'
                message += f'KW： {e[0]}\n'
                message += f'順位： {e[1]}\n'
                message += f'URL： {e[2]}\n'
                message += f'日付： {e[3]}\n'
            message += '[/info]'

        sendChatworkNotification(message)
        logger.info("alert_ranking_data: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'alert_ranking_data: {err}')
        exit(1)
