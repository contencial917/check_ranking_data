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

def createPages(domain, year, month, data):
    date = f'{year}年{month}月'
    pages = ""
    total_pages = len(data)
    for index, item in enumerate(data):
        out_of_range_flag = False
        ranking_table = ""
        ranking_data = ""
        out_of_range = ""
        labels = ""
        best = 101
        worst = -1
        max_size = 0
        step_size = 0
        sum_rank = 0
        sum_days = 0
        keyword = item.pop(0)

        for i, d in enumerate(item):
            ranking_table += '<tr class="keyword__row">\n'\
                            f'<td class="keyword__td">{year}/{month}/{str(i + 1)}</td>\n'\
                            f'<td class="keyword__td">{d}</td>\n'\
                            '</tr>\n'

            if d == "-":
                ranking_data += '150,'
                out_of_range = '<p class="keyword__supplement">ー...範囲外 </p>'
                sum_rank += 101
                worst = 101
            else:
                ranking_data += f'{d},'
                rank = int(d)
                if rank < best:
                    best = rank
                if rank > worst:
                    worst = rank
                sum_rank += rank

            if i == 0 or (i + 1) % 5 == 0:
                labels += f'"{month}/{i + 1}",'
            else:
                labels += '"",'

            sum_days += 1

        ranking_table += '<tr class="keyword__row">\n'\
                        '<td class="keyword__td">最高</td>\n'
        if best == 101:
            ranking_table += '<td class="keyword__td">-</td>\n'
        else:
            ranking_table += f'<td class="keyword__td">{best}</td>\n'
        ranking_table += '</tr>\n'\
                        '<tr class="keyword__row">\n'\
                        '<td class="keyword__td">最低</td>\n'
        if worst == -1 or worst == 101:
            ranking_table += '<td class="keyword__td">-</td>\n'
        else:
            ranking_table += f'<td class="keyword__td">{worst}</td>\n'
        ranking_table += '</tr>\n'\
                        '<tr class="keyword__row">\n'\
                        '<td class="keyword__td">平均</td>\n'
        average = int(sum_rank / sum_days)
        if average >= 100:
            ranking_table += '<td class="keyword__td">-</td>\n'
        else:
            ranking_table += f'<td class="keyword__td">{average}</td>\n'
        ranking_table += '</tr>\n'

        if worst == -1 or worst > 50:
            max_size = 100
            step_size = 20
        elif worst <= 10:
            max_size = 10
            step_size = 1
        elif worst <= 20:
            max_size = 20
            step_size = 5
        elif worst <= 50:
            max_size = 50
            step_size = 10

        with open('./template/section.tpl', 'r', encoding='utf_8') as f:
            page = f.read()

        page = page.replace('{ domain }', domain)
        page = page.replace('{ keyword }', keyword)
        page = page.replace('{ ranking_table }', ranking_table)
        page = page.replace('{ out_of_range }', out_of_range)
        page = page.replace('{ date }', date)
        page = page.replace('{ labels }', labels.strip(','))
        page = page.replace('{ ranking_data }', ranking_data.strip(','))
        page = page.replace('{ max_size }', str(max_size))
        page = page.replace('{ step_size }', str(step_size))
        page = page.replace('{ total_pages }', str(total_pages))
        page = page.replace('{ page_no }', str(index + 1))

        pages += page

    return pages

def createFile(pages, output, domain):
    with open('./template/body.tpl', 'r', encoding='utf_8') as f:
        html = f.read()
    html = html.replace('{ pages }', pages)
    with open(f'{output}/{domain}.html', 'w', newline='', encoding='utf_8') as f:
        f.write(html)

def createReport(domain, sheet, year, month, output):
    try:
        data = sheet.get_all_values()
        data.pop(0)
        data.pop(0)
        data.pop(0)

        if domain == "wakigacenter.com":
            lst = data.pop(0)
            pages = createPages(domain, year, month, [lst])
            createFile(pages, output, f'{domain}_kodomo-wakiga')

        pages = createPages(domain, year, month, data)
        createFile(pages, output, domain)

    except Exception as err:
        logger.debug(f'Error: create_report: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':

    try:
        thismonth = datetime.datetime(today.year, today.month, 1)
        lastmonth = thismonth + datetime.timedelta(days=-1)
        year = lastmonth.strftime("%Y")
        month = lastmonth.strftime("%m")

        output_path = os.environ['RANK_REPORT_PATH']
        os.makedirs(f'{output_path}/{year}/{month}', exist_ok=True)

        rankDataDirPath = os.environ["RANK_DATA_DIR"]

        config = configparser.ConfigParser()
        config.read_file(codecs.open("clientInfo.ini", "r", "utf8"))
        projects = config.sections()

        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
        gc = gspread.authorize(credentials)

        for project in projects:
            SPREADSHEET_ID = config[project]['SSID']
            sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(f'{year}{month}')
            createReport(project, sheet, int(year), int(month), f'{output_path}/{year}/{month}')
            logger.debug(f'create_report: {project}')
            sleep(3)
        
        message = '[info][title]順位計測結果 出力完了[/title]\n'
        message += '今月の順位計測結果の出力が完了しました。\n'
        message += 'Dropboxの下記ディレクトリから順位計測結果をご確認ください。\n\n'
        message += f'/順位計測結果レポート/{year}/{month}/'
        message += '[/info]'

        sendChatworkNotification(message)
 
        logger.info("create_report: Finish")
        exit(0)
    except Exception as err:
        logger.debug(f'Error: create_report: {err}')
        exit(1)