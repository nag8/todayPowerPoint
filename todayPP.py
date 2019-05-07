# coding: UTF-8

from pptx import Presentation
from bs4 import BeautifulSoup
import configparser
import requests
import feedparser
import datetime


# メイン処理
def main():

    # 参照
    # https://python-pptx.readthedocs.io/en/latest/user/table.html
    # https://qiita.com/hujuu/items/b0339404b8b0460087f9

    # 設定ファイル取得
    iniFile = getIniFile()

    # PowerPointファイルを生成
    createPP(iniFile)


# 設定ファイル取得
def getIniFile():
    iniFile = configparser.ConfigParser()
    iniFile.read('./config.ini', 'UTF-8')
    return iniFile


# パワーポイントを生成
def createPP(iniFile):
    prs = Presentation(iniFile.get('settings', 'IN'))

    # htmlから情報を取得
    strSchedule = getSchedule(iniFile)

    # 表を変更
    editTable(prs.slides[0].shapes[0].table, prs.slides[1].shapes[1].table, strSchedule)

    # ファイルを保存
    prs.save(iniFile.get('settings', 'OUT') + judgeFileName() + '.pptx')


# スケジュールを取得
def getSchedule(iniFile):

    # 保存したhtmlを取得
    with open(iniFile.get('settings', 'HTML'), encoding="shift_JIS", errors='ignore') as f:
        html = f.read()

    #要素を抽出
    soup = BeautifulSoup(html, 'lxml')
    
    # テーブルを指定
    table = soup.findAll("table")[0]
    trs = table.findAll("tr")
    print(trs)

    return True


# テーブルを修正
def editTable(table1, table2, strSchedule):
    # 展示ギャラリー
    table1.cell(1, 2).text = 'test'

    # 講座研修室
    table1.cell(2, 2).text = 'test21'


    table2.cell(5, 3).text = 'test55'
    table2.cell(1, 3).merge(table2.cell(2, 3))


# ファイルネームを生成
def judgeFileName():

    # 曜日
    yobi = ["月","火","水","木","金","土","日"]

    # 明日を取得
    tomorrow = datetime.datetime.now() + datetime.timedelta(days = 1)
    # 整形して返却
    return '{}月{}日（{}）'.format(tomorrow.month, tomorrow.day, yobi[tomorrow.weekday()])



if __name__ == '__main__':
    main()
