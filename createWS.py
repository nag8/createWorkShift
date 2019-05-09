# coding: UTF-8

from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell.cell import MergedCell
import configparser
from datetime import datetime, date, timedelta
import calendar
import copy




# 使わせてもらったコード
# https://qiita.com/TakayukiKiyohara/items/83469bbf9d3786333f48



# メイン処理
def main():

    # 設定ファイル取得
    iniFile = getIniFile()

    # Excelファイルを読み込み
    wb = openpyxl.load_workbook(iniFile.get('settings', 'IN'))


    # 表のタイトルを修正
    sheet = wb[iniFile.get('settings', 'TARGET')]
    sheet['A1'].value = createTitle()

    # 初期化
    initTable(sheet,wb[iniFile.get('settings', 'TEMPLATE')])

    # 保存
    wb.save(iniFile.get('settings', 'OUT') + createTitle() + '.xlsx')


# 設定ファイル取得
def getIniFile():
    iniFile = configparser.ConfigParser()
    iniFile.read('./config.ini', 'UTF-8')
    return iniFile

#タイトルを生成
def createTitle():

    # 再来月のやつとか必要かも
    return str(nextMonth().month) + '月シフト'

# 来月の1日を取得
def nextMonth():

    # とりあえず翌月固定
    monthNum = 1

    # 来月の1日を取得
    return datetime.today().replace(day = 1) + relativedelta(months = monthNum)

# 表を翌月で初期化
def initTable(sheet,templateSheet):

    writeSchedule(sheet)



    return ''


# 表に日付を記入
def writeSchedule(sheet):

    day    = nextMonth()
    rowNum = 2

    # 曜日と列番号の辞書
    dic = {0 : 6 , 1 : 11 , 2 : 16 , 3 : 21 , 4 : 26 , 5 : 31 , 6 : 36}

    for i in range(calendar.monthrange(day.year, day.month)[1]):

        colNum = dic[day.weekday()]

        # 日付の記入
        sheet.cell(row = rowNum, column = colNum).value = day

        # シフトのコピー
        RangeCopyCell(sheet, colNum, 104, colNum + 4, 104 + 15, 0, -(104 - 2 - rowNum))

        day += relativedelta(days = 1)

        if day.weekday() == 0:
            # TODO
            rowNum += 20


# 範囲をコピー
def RangeCopyCell( sheet, min_col, min_row, max_col, max_row, shift_col, shift_row ):
    #コピー元の結合されたセルの結合を解除していく。
    merged_cells = copy.deepcopy(sheet.merged_cells);
    for merged_cell in merged_cells :
        if merged_cell.min_row >= min_row \
            and merged_cell.min_col >= min_col \
            and merged_cell.max_row <= max_row \
            and merged_cell.max_col <= max_col :
                #結合を解除していく。
                sheet.unmerge_cells(merged_cell.coord);

    #全セルをコピー
    for col in range( min_col, max_col + 1):
        for row in range( min_row, max_row + 1):
            fmt = "{min_col}{min_row}"
            #コピー元のコードを作成。
            copySrcCord = fmt.format(
                min_col = get_column_letter(col),
                min_row = row );

            #コピー先のコードを作成。
            copyDstCoord = fmt.format(
                min_col = get_column_letter(col + shift_col),
                min_row = row + shift_row );
            #コピー先に値をコピー。
            if type(sheet[copySrcCord]) != MergedCell :
                sheet[copyDstCoord].value = sheet[copySrcCord].value;
                #書式があったら、書式もコピー。
                if sheet[copySrcCord].has_style :
                    sheet[copyDstCoord]._style = sheet[copySrcCord]._style;

    #結合解除したセルを再結合していく。
    for merged_cell in merged_cells :
        if merged_cell.min_row >= min_row \
            and merged_cell.min_col >= min_col \
            and merged_cell.max_row <= max_row \
            and merged_cell.max_col <= max_col :
                #結合していく。
                sheet.merge_cells(merged_cell.coord);
                #コピー先のセルの結合範囲情報を作成する。
                newCellRange = copy.copy(merged_cell);
                #コピー先のセルの結合範囲情報をずらす。ここではshiftRow分ずらしている。
                newCellRange.shift(shift_col, shift_row);
                sheet.merge_cells(newCellRange.coord);
    return 0;

if __name__ == '__main__':
    main()
