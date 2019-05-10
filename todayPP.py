# coding: UTF-8

from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from bs4 import BeautifulSoup
import configparser
import requests
import feedparser
import datetime
import csv


# メイン処理
def main():

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
    inputTable = getInputTable(iniFile)

    # 表を変更
    editPPTable(iniFile, prs.slides[0].shapes[0].table, prs.slides[1].shapes[1].table, inputTable)

    # ファイルを保存
    prs.save(iniFile.get('settings', 'OUT') + createFileName() + '.pptx')


# データ元の情報を取得
def getInputTable(iniFile):

    # 保存したhtmlを取得
    with open(iniFile.get('settings', 'HTML'), encoding="shift_JIS", errors='ignore') as f:
        html = f.read()

    #要素を抽出
    soup = BeautifulSoup(html, 'lxml')

    # テーブルを指定
    return soup.findAll("table")[0]

# セルのフォントサイズを変更して、中央揃えにする
def changeLayout(cell, size):
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)

        # 中央揃えにもする
        paragraph.alignment = PP_ALIGN.CENTER


# パワーポイントのテーブルを修正
def editPPTable(iniFile, table1, table2, inputTable):
    # 要素を取得
    tdList = inputTable.findAll("td", attrs = {"class": "p11pa2"})

    # 設定に必要なcsvを取得
    directory = getDirectory(iniFile.get('settings', 'CSV'))

    for td in tdList:

        # 行番号を取得。なければ飛ばす
        if directory[td.text[:3]][1] != '':
            rowNum = int(directory[td.text[:3]][1])

            contents = td.parent.findAll("td", attrs = {"class": "p11"})
            for i,content in enumerate(contents):

                # テーブル番号、列番号を設定
                if directory[td.text[:3]][2] == '1':
                    changeTable = table1
                    columnNum   = 2 + i

                else:
                    changeTable = table2
                    columnNum   = 1 + i

                # テキストを設定
                changeCell = changeTable.cell(rowNum, columnNum)
                changeCell.text = changeStr(content.get_text('.').split('.'))

                # レイアウトを修正
                changeLayout(changeCell, 10)

    # いつかやる
    # table2.cell(1, 3).merge(table2.cell(2, 3))

# 文字を取得
def changeStr(strList):

    if len(strList) >= 2:
        return strList[1] + '\n（' + strList[0] + '　様）'

    elif len(strList) == 1:
        return strList[0]

    else:
        return 'error'


# CSVを取得
def getDirectory(csvPath):

    directory = {}

    with open(csvPath, 'r') as f:
        reader = csv.reader(f)

        for row in reader:
            directory[row[0]] = [row[1],row[2],row[3]]

    return directory



# ファイルネームを生成
def createFileName():

    # 曜日
    yobi = ["月","火","水","木","金","土","日"]

    # 明日を取得
    tomorrow = datetime.datetime.now() + datetime.timedelta(days = 1)
    # 整形して返却
    return '{}月{}日（{}）'.format(tomorrow.month, tomorrow.day, yobi[tomorrow.weekday()])



if __name__ == '__main__':
    main()
