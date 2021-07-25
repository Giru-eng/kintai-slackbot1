# -*- coding: utf-8 -*-
import pandas as pd

from slackbot.bot import respond_to
from datetime import datetime, timedelta

import gspread

from oauth2client.service_account import ServiceAccountCredentials

@respond_to('出勤')
def punch_in(message):
    print('出勤時刻を登録します')
    real_name = message.user['real_name']
    timestamp = datetime.now() + timedelta(hours=9)
    date = timestamp.strftime('%Y/%m/%d')
    punch_in = timestamp.strftime('%H:%M')
    message.send(f'出勤時刻は、 {punch_in}です。')
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    df = df.append({'Real_name': real_name, 'Date': date, 'Punch_in': punch_in, 'Punch_out': '99:99', 'Break_start': '99:99', 'Break_end': '99:99'}, ignore_index=True)
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    print('登録完了しました。')
    message.send('勤怠登録完了しました。')

@respond_to('休憩開始')
def break_start(message):
    print('休憩開始時刻を登録します')
    real_name = message.user['real_name']
    timestamp = datetime.now() + timedelta(hours=9)
    break_start = timestamp.strftime('%H:%M')
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    cells = worksheet.findall(real_name)
    for cell in cells:
        cell = cell.row
    if df.iloc[cell-2, 4] == '99:99':
        message.send(f'休憩開始時刻は、 {break_start}です。')
        df.iloc[cell-2, 4] = break_start
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
        print('登録完了しました。')
        message.send('勤怠登録完了しました。')
    else:
        message.send('休憩開始時刻は出勤時間登録後でなければ登録できません。\nまずは現時点で”出勤”と入力し、仮の出勤時刻の登録を完了してから休憩開始時刻の登録をお願いします。\nまた仮の出勤打刻は修正が必要になります。必ず退勤までに所定の勤怠申請を行い正しい出勤時刻へ訂正をお願い致します。')

@respond_to('休憩終了')
def break_end(message):
    print('休憩終了時刻を登録します')
    real_name = message.user['real_name']
    timestamp = datetime.now() + timedelta(hours=9)
    break_end = timestamp.strftime('%H:%M')
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    cells = worksheet.findall(real_name)
    for cell in cells:
        cell = cell.row
    if df.iloc[cell-2, 5] == '99:99':
        message.send(f'休憩終了時刻は、 {break_end}です。')
        df.iloc[cell-2, 5] = break_end
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
        print('登録完了しました。')
        worksheet = auth2()
        df = pd.DataFrame(worksheet.get_all_records())
        cells = worksheet.findall(real_name)
        for cell in cells:
            cell = cell.row
        if df.iloc[cell-2, 5] == "More break?":
            bt = df.iloc[-1, 4]
            message.send(f'休憩時間が{bt}です。1時間を超えてますので勤怠申請をお願いします。')
        elif df.iloc[cell-2, 5] == "Less break?":
            bt = df.iloc[-1, 4]
            message.send(f'休憩時間が{bt}です。45分に満たないようですので勤怠申請をお願いします。')
        else:
            message.send('勤怠登録完了しました。')
    else:
        message.send('休憩終了時刻は出勤時間登録後でなければ登録できません。\nまずは現時点で”出勤”と入力し、仮の出勤時刻の登録を完了してから休憩終了時刻の登録をお願いします。\nまた仮の出勤打刻は修正が必要になります。必ず退勤までに所定の勤怠申請を行い正しい出勤時刻へ訂正をお願い致します。')

@respond_to('退勤')
def punch_out(message):
    print('退勤時刻を登録します')
    real_name = message.user['real_name']
    timestamp = datetime.now() + timedelta(hours=9)
    punch_out = timestamp.strftime('%H:%M')
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    cells = worksheet.findall(real_name)
    for cell in cells:
        cell = cell.row
    if df.iloc[cell-2, 3] == '99:99':
        message.send(f'退勤時刻は、 {punch_out}です。')
        df.iloc[cell-2, 3] = punch_out
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
        print('登録完了しました。')
        worksheet = auth2()
        df = pd.DataFrame(worksheet.get_all_records())
        cells = worksheet.findall(real_name)
        for cell in cells:
            cell = cell.row
        if df.iloc[cell-2, 2] == 'Overtime?':
            wt = df.iloc[-1, 1]
            message.send(f'勤務時間が{wt}です。8時間5分を超えているようですので勤怠申請をお願いします。')
        elif df.iloc[cell-2, 3] == 'Tardy? / Leaving early?':
            wt = df.iloc[-1, 1]
            message.send(f'勤務時間が{wt}です。8時間に満たないようですので勤怠申請をお願いします。')
        else:
            message.send('勤怠登録完了しました。')
    else:
        message.send('退勤時刻は出勤時間登録後でなければ登録できません。\nまずは現時点で”出勤”と入力し、仮の出勤時刻の登録を完了してから退勤時刻の登録をお願いします。\nまた仮の出勤打刻は修正が必要になります。必ず本日中に所定の勤怠申請を行い正しい出勤時刻へ訂正をお願い致します。')


def auth():
    SP_CREDENTIAL_FILE = 'plugins/secret.json'
    SP_SCOPE = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    SP_SHEET_KEY = '1bUOyaGOL2eBEoIrP8lVdxeGWPrdTY7jzRYGbtN9HY6k'
    SP_SHEET = 'timecard'

    credentials = ServiceAccountCredentials.from_json_keyfile_name(SP_CREDENTIAL_FILE, SP_SCOPE)
    gc = gspread.authorize(credentials)

    worksheet = gc.open_by_key(SP_SHEET_KEY).worksheet(SP_SHEET)
    return worksheet

def auth2():
    SP_CREDENTIAL_FILE = 'plugins/secret.json'
    SP_SCOPE = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    SP_SHEET_KEY = '1bUOyaGOL2eBEoIrP8lVdxeGWPrdTY7jzRYGbtN9HY6k'
    SP_SHEET = 'alerts'

    credentials = ServiceAccountCredentials.from_json_keyfile_name(SP_CREDENTIAL_FILE, SP_SCOPE)
    gc = gspread.authorize(credentials)

    worksheet = gc.open_by_key(SP_SHEET_KEY).worksheet(SP_SHEET)
    return worksheet
