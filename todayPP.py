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

    print('処理開始！')

    # 設定ファイル取得
    iniFile = getIniFile()

    # PowerPointファイルを生成
    createPP(iniFile)

    print('処理終了！')

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


        print('----------------------------------')

        # どこの表に渡せばいいかリスト
        dicList = directory[td.text[:3]]
        print(dicList[0])

        # 行番号を取得。なければ飛ばす
        if dicList[1] != '':
            rowNum = int(dicList[1])


            targetTdList = td.parent.findAll("td")

            # 闇の魔術
            targetTdList.pop(0)

            cellNum = 0

            for i, targetTd in enumerate(targetTdList):
                # コマの順番

                # 対象内にテーブル要素がある場合
                if targetTd.find("table"):

                    # 闇の魔術。tableの後ろにtdがあるので削除。findAllでそもそも取得されないようにしたい・・・・
                    targetTdList.pop(i + 1)

                    # セルの大きさを追加
                    cellNum += int(targetTd.get('colspan'))

                    # 表1の設定を初期設定にする
                    changeTable = table1
                    columnNum   = 2

                    # セルの種類。元のテンプレートが複数種類あるので、それに対応
                    tdId = dicList[2]

                    # 展示ギャラリーの場合
                    if tdId in ['1','2','3']:
                        # 1日貸出のため、特に考慮する必要なし
                        pass

                    # 通常の施設の場合（会議室など）
                    elif tdId == '2':

                        # 取りうるパターン。これに空セルの場合も入る。戦略的後回し
                        # 3,1,3,1,2
                        # 7,1,2
                        # 3,1,6
                        # 10
                        targetList = getNormalList()

                        pass

                    # 音楽スタジオの場合
                    elif tdId == '3':

                        # 取りうるパターン

                        pass


                    else:
                        changeTable = table2
                        columnNum   = 1

                        # 始まりセルを設定
                        startNum = cellNum - int(targetTd.get('colspan')) + 1

                        if targetTd.get('colspan') != '1':
                            print('merge → ')

                            print('rowNum' + str(rowNum) + 'startNum' + str(startNum) + 'cellNum' + str(cellNum))
                            # 通常の予定入力の場合

                            changeCell = changeTable.cell(rowNum, startNum)
                            changeCell.text = getStr(targetTd.find("td", attrs = {"class": "p11"}))
                            changeLayout(changeCell, 12)

                            if changeCell.is_merge_origin:
                                changeCell.split()

                            # changeCell.merge(changeTable.cell(rowNum, cellNum))

                            pass



                    print('順序 → ' + str(cellNum) + ' ;長さ → ' + targetTd.get('colspan') + ' ; ' + targetTd.find("td", attrs = {"class": "p11"}).get_text(' ; '))

                # 空セルの場合
                else:
                    cellNum += 1
                    print("順序 → " + str(cellNum) + ' ; 空セル')
                    # print(changeCell.is_merge_origin)



# 文字を取得
def getStr(content):

    # 内容を取得
    strList = content.get_text(';').split(';')

    # 未入金の場合があるので、それを削除
    if strList[0] == '未':
        strList.pop(0)

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
