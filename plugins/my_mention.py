
import re
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from slackbot.bot import respond_to
from slackbot.bot import listen_to


target_channels = ['grp_timecard', 'api_test']
def auth():
    SP_CREDENTTIAL_FILE = 'secret.json'
    SP_SCOPE = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    SP_SHEET_KEY = '1DuUUu-MAk_N6a84-OOv-b_801AWYWvVRZ6Ugpjo1IiY'
    SP_SHEET = 'timesheet'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        SP_CREDENTTIAL_FILE, SP_SCOPE)
    gc = gspread.authorize(credentials)

    worksheet = gc.open_by_key(SP_SHEET_KEY).worksheet(SP_SHEET)
    return worksheet
# 出勤


@listen_to('^(出勤|s)\n?(.*)$')
def punch_in(message, state='', note=''):
    if note == '':
        note = '記録なし'
    real_name = message.user['real_name']
    display_name = message.user['profile']['display_name']
    user_id = message.user['id']
    message.reply('勤務簿を確認しています。少々お待ちください。', in_thread=True)
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())

    if df[(df["退勤時刻"] == "記録なし") & (df["display_name"] == display_name)].empty:
        timestamp = datetime.now()
        date = timestamp.strftime('%Y/%m/%d')
        time = timestamp.strftime('%H:%M')
        df = df.append({'display_name': display_name, '名前': real_name, '出勤日付': date, '出勤時刻': time,
                        '退勤日付': '記録なし', '退勤時刻': '記録なし', '備考': note}, ignore_index=True)
        worksheet.update([df.columns.values.tolist()]+df.values.tolist())
        message.reply(
            f'{real_name}(<@{user_id}>)さんの出勤を{date}の{time}で記録しました\n備考:{note}\n勤務中はdiscoodにつなぎましょう', in_thread=True)
        message.reply(
            f'<https://discord.gg/sNzuhap4eT|discoodはこちら>', in_thread=True)
        message.react('+1')
        message.react('いい_仕事')
    else:
        message.reply('退勤時刻が記録されていません。記録は以下の通りです。', in_thread=True)
        message.reply(str(df[(df["退勤時刻"] == "記録なし") & (
            df["display_name"] == display_name)]), in_thread=True)
        message.reply(
            'このチャンネルで「退勤」と言うかスプレッドシートを編集して自分の出勤記録を変更してください', in_thread=True)
        message.react('-1')
    message.reply(
        f'<https://docs.google.com/spreadsheets/d/1DuUUu-MAk_N6a84-OOv-b_801AWYWvVRZ6Ugpjo1IiY/edit#gid=0|出勤記録はこちら>', in_thread=True)


@listen_to('^(退勤|t)\n?(?s:.*)$')
def punch_out(message, state='', note=''):
    # if note == '':
    #     note = '記録なし'
    # チャンネル名を取得
    channel = message.channel._body['name']
    # 対象チャンネルでなければ終了する
    if channel not in target_channels:
        return

    real_name = message.user['real_name']
    display_name = message.user['profile']['display_name']
    user_id = message.user['id']
    message.reply('勤務簿を確認しています。少々お待ちください。', in_thread=True)
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    if not (df[(df["退勤時刻"] == "記録なし") & (df["display_name"] == display_name)]).empty:
        timestamp = datetime.now()
        date = timestamp.strftime('%Y/%m/%d')
        time = timestamp.strftime('%H:%M')
        applicable_df = (df["退勤時刻"] == "記録なし") & (
            df["display_name"] == display_name)
        df.loc[applicable_df, "退勤日付"] = date
        df.loc[applicable_df, "退勤時刻"] = time
        # if (df.loc[applicable_df, "備考"].item() != "記録なし"):
        #     df.loc[applicable_df, "備考"] = note
        worksheet.update([df.columns.values.tolist()]+df.values.tolist())
        message.reply(
            f'{real_name}(<@{user_id}>)さんの退勤を{date}の{time}で記録しました\n備考:{note}', in_thread=True)
        message.react('+1')
        message.react('いい_仕事')

    else:
        message.reply('出勤時刻が記録されていません。', in_thread=True)
        message.reply(
            'このチャンネルで「出勤」と言うかスプレッドシートを編集して自分の出勤記録を変更してください', in_thread=True)
        message.react('-1')
    message.reply(
        f'<https://docs.google.com/spreadsheets/d/1DuUUu-MAk_N6a84-OOv-b_801AWYWvVRZ6Ugpjo1IiY/edit#gid=0|出勤記録はこちら>', in_thread=True)


@listen_to('^勤務中$')
def working_people(message):
    # チャンネル名を取得
    channel = message.channel._body['name']
    # 対象チャンネルでなければ終了する
    if channel not in target_channels:
        return

    # チャンネル名を取得
    channel = message.channel._body['name']
    # 対象チャンネルでなければ終了する
    if channel not in target_channels:
        return

    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    message.reply('現在勤務中の方のリストはこちらです', in_thread=True)
    message.reply(str(df[df['退勤日付'] == '記録なし']), in_thread=True)
    message.react('+1')


@respond_to('^(アイス|りんご)(.*)$')
def greeting_1(message, group1='あ', group2='い'):
    # グループの文字列を表示する
    message.reply(group1+''+group2)


# =============================
