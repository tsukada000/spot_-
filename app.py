import streamlit as st

st.write("社保新適リスト作成")

st.caption('下記顧客分析スプレッドに本日付けでシートを追加します。')
st.caption('既に同じ日付のシートが存在しないことを確認してください。')
st.caption('同じ日付のシートが存在している場合は、事前に該当シートを削除してください。')

https://docs.google.com/spreadsheets/d/1bC0Hs9C0rAS33BavF-E1yNwwGwqe8vsmRzodQt04esI/edit#gid=0

if st.button("リスト作成"):

  import gspread
  import pandas as pd
  import os
  import datetime
  from google.colab import auth
  from google.auth import default
  from google.oauth2 import service_account
  from gspread_dataframe import set_with_dataframe
  from gspread_dataframe import get_as_dataframe
  from google.colab import drive
  drive.mount('/content/drive') 
  os.chdir(f'/content/drive/MyDrive/スポット社労士/【スポット社労士】顧客分析/')

#認証情報
  service_account_key = {
    "type": "service_account",
    "project_id": "spot-customeranalysis",
    "private_key_id": "130c20cdc874d1dc2fa6ff953a4c4ee118baa48d",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCidjr4qpb4jIax\ngiqxZgfCs/Hxh+K/YYY4oVrSf0U8NDyfsaHN6fzcW3lVUEwcy6o61TUhBfoI12QJ\n8Rp3F0u7unjNSUCK58t6j/wRfHS6qGVpEfBgB4LDmTII48zWnpETVFvIVlJIAqFv\nQC2gwCix5UuYqh5QinCTaCAPLZf5jiwiyWeOHo+fXsTD6tntuk2uK3m5Tb+H9DYF\nhJPoXrqE2Zm37sh8aQ/LSUZRmAgsolIEu/D5BWhJgy5iHeeIS8sttRwPudIV5DSA\nBi0Bv5RWK+3EUucVGIU2t2MFSTBc0RpzEUZ6PhOjR1fYzSRCgABihYc3TjQtQpqz\nMCdwaJu3AgMBAAECggEABTsFPwhPAAWWOJTWRwvf6BbfDUWwuTSsm5omUGykkAGm\nigWwDe8gov+W8nY5XRv5iRdgNIX6vNoicGeA0KDBYXPpe8KF/3LjDDkihnWZVIEj\ncrSPJqhXP3DpOXwMFkTYquSmZ9bPo13iULCiR3CoXsHhIzMm8rTiVvydO5/eKryp\n7CBqCHOCt9ebVGZGPXkltTbMLoxh4xTlC4AhPexPEV9YJPhNx2Rb1rtOm8SwkMt9\njXkeYcXlcRRnlUa/ah9tB7IEyYjvy9nPhd7TV2SrY3uaHAuWLcM82E1ntvQkGgaY\niVDjkNTwmQNhO6Rw3KsbbID0LyPGOPN6YYNzIed/QQKBgQDXndmljOACcQM8NH0s\ntCI4YDnyvz5r7yoYpp0s6j7NzCzj3xJw/xAZ6os8m37Btu95Nt4oO26VvAvgutsy\n/x1DllGxPRDkVBwp8t66dsei5di5oAQDE4jMQWbGj1vg+NyFxEPA7jwC1/Szg7B8\nxz/JiLkP/wOj4QzrAUWnbIZLRwKBgQDA48YAgS2qnujbyjLgdKIhos+hBSdL5Z2x\nlZnn1eYNhLJZmbdeFATaU2DwWEbaio9gktHpZAQcr7LqclJk1vjDZA+koPME9MlU\nmXy28do+lZLUF/g8aT7YUQH9PG98vxvSkf8opPCp6iSB/oYlxQV7FXIdbTHj7OIN\nPThQSmuEEQKBgFOQ4iG/j7JiipZy4XDJ/9lJsiva4x6B+xbCvHgD8YNhdqR6eHNC\n58KjnINI4L/DXtzj3wZIwntV/mSDByGkrnrbb535xOo5jxDTCG/MSWNhIbYPxn5K\nu+IuFt8uALYYvZ86iefkbW3MtRI+H9C8iIRbcR45//cr6g3K2GwjK4lRAoGAZ4HS\n3rJzLvvXORpn8sqjtikIAgAh9jhhRspgrGe768Upb6ttGq7ja8USX+b/Hob8KXaf\n7f4dtscR2309eZ9iHnezbURxJFe3Mg6rPgDKfIsHH4k9TC2t66aMyreDnA1xgK2X\nntfjzUo4DQnoMpxnNIVtlxzhiM21ACW58lv9FgECgYBrRB4cE24lVXn17G/jfLN9\nUkbNkvxFQ7nfNFa/q6gxdHF8Xsaf6k7M1hKJwZQXW/Q0vFXbiMizErCgF9oT6Vm7\nvhxrYSnBiFmcSaFN5tUUiQvmDWLW3eoekf5/fIYNV6Sbe0bBJpcgu9yaQD8oWsrp\njP5fuh13SQH34+2Fh55DxA==\n-----END PRIVATE KEY-----\n",
    "client_email": "spot-customeranalysis@spot-customeranalysis.iam.gserviceaccount.com",
    "client_id": "114998122524979325747",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/spot-customeranalysis%40spot-customeranalysis.iam.gserviceaccount.com"
  }
  credentials = service_account.Credentials.from_service_account_info(service_account_key)
  scoped_credentials = credentials.with_scopes(
    [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ])

  gs = gspread.authorize(scoped_credentials)

  SPREADSHEET_KEY1 = '1jNJrULTvCycVWMyQXXk8j7nzzxiq4fWuTNEftGvPe0E'
  workbook1 = gs.open_by_key(SPREADSHEET_KEY1)
  datasheet = workbook1.worksheet("Data")

  basedata_df = get_as_dataframe(datasheet, usecols=[0,1,2,3,4,5,6,7,8,9,10,11], header=0).dropna(how='all')

#パターン①「初回：社会保険新規適用、２回目以降：社会保険関係の手続きのみ」
  pdata1_df = basedata_df.loc[(basedata_df["品目"].str.contains("社会保険（健康保険・厚生年金）新規適用")) & (basedata_df["購入回数"] == 1)]
  pdata1_df = pdata1_df.groupby(['取引先'], as_index=False).max()
  pdata2_df = basedata_df.loc[(basedata_df['品目'].str.contains("社保取得")) | (basedata_df['品目'].str.contains("社保喪失"))  & (basedata_df["購入回数"] > 1)]
  pdata2_df = pdata2_df.groupby(['取引先'], as_index=False).max()
  pdata3_df = basedata_df.loc[(~basedata_df['品目'].str.contains("社保取得")) & (~basedata_df['品目'].str.contains("社保喪失")) & (basedata_df["購入回数"] > 1)]
  pdata3_df = pdata3_df.groupby(['取引先'], as_index=False).max()
  pdata4_df = basedata_df.loc[(basedata_df['品目'].str.contains("労働保険（労災・雇用）新規適用")) & (basedata_df["購入回数"] == 1)]
  pdata4_df = pdata4_df.groupby(['取引先'], as_index=False).max()
  p1out_df = pdata1_df[~pdata1_df["取引先"].isin(pdata4_df["取引先"])]
  p1out_df = p1out_df[~p1out_df["取引先"].isin(pdata3_df["取引先"])]
  p1out_df = p1out_df[p1out_df["取引先"].isin(pdata2_df["取引先"])]
  p1out_df = basedata_df[basedata_df["取引先"].isin(p1out_df["取引先"])]
  p1out_df = p1out_df.sort_values(['取引先','管理番号']) 
  p1out_df.insert(0, 'パターン', '初回：社保新適、２回目以降：社保関係のみ')

#パターン②「初回：他サービス、２回目以降：社会保険新規適用(労働者が関係する手続きの発生記録なし)」
  pdata21_df = basedata_df.loc[(~basedata_df["品目"].str.contains("社会保険（健康保険・厚生年金）新規適用")) & (basedata_df["購入回数"] == 1)]
  pdata21_df = pdata21_df.drop(columns=['品目', '購入回数']).groupby(['取引先'], as_index=False).max()
  pdata22_df = basedata_df.loc[(basedata_df["品目"].str.contains("社会保険（健康保険・厚生年金）新規適用")) & (basedata_df["購入回数"] > 1)]
  pdata22_df = pdata22_df.drop(columns=['品目', '購入回数']).groupby(['取引先'], as_index=False).max()
  pdata23_df = basedata_df.loc[(basedata_df["品目"].str.contains("労働保険（労災・雇用）新規適用")) | (basedata_df["品目"].str.contains("雇用取得")) | (basedata_df["品目"].str.contains("雇用喪失")) | (basedata_df["品目"].str.contains("時間外・休日労働に関する協定（36協定）届")) | (basedata_df["品目"].str.contains("離職票")) | (basedata_df["品目"].str.contains("年度更新")) | (basedata_df["品目"].str.contains("労務相談")) | (basedata_df["顧客分析区分"].str.contains("助成金"))]
  pdata23_df = pdata23_df.drop(columns=['品目', '購入回数']).groupby(['取引先'], as_index=False).max()
  p2out_df = pdata21_df[pdata21_df["取引先"].isin(pdata22_df["取引先"])]
  p2out_df = p2out_df[~p2out_df["取引先"].isin(pdata1_df["取引先"])]
  p2out_df = p2out_df[~p2out_df["取引先"].isin(pdata23_df["取引先"])]
  p2out_df = basedata_df[basedata_df["取引先"].isin(p2out_df["取引先"])]
  p2out_df = p2out_df.sort_values(['取引先','管理番号']) 
  p2out_df.insert(0, 'パターン', '初回：他サービス、２回目以降：社保新適')

#パターン③「初回：社会保険新規適用1件のみ」
  pdata30_df = basedata_df.loc[basedata_df["購入回数"] > 1]
  p3out_df = pdata20_df[~pdata20_df["取引先"].isin(pdata21_df["取引先"])]
  p3out_df = p3out_df[~p3out_df["取引先"].isin(pdata30_df["取引先"])]
  p3out_df = basedata_df[basedata_df["取引先"].isin(p3cnt_df["取引先"])]
  p3out_df = p3out_df.sort_values(['取引先','管理番号']) 
  p3out_df.insert(0, 'パターン', '初回：社保新適１件のみ')

  out_df = pd.concat([p1out_df, p3out_df, p2out_df])

  SPREADSHEET_KEY2 = '1bC0Hs9C0rAS33BavF-E1yNwwGwqe8vsmRzodQt04esI'
  workbook2 = gs.open_by_key(SPREADSHEET_KEY2)
  today = datetime.date.today()
  today = format(today, '%Y%m%d')
  wstitle = "社保新適リスト_" + today
  worksheet = workbook2.add_worksheet(title=wstitle, rows="100",cols="20")
  set_with_dataframe(workbook2.worksheet(wstitle), out_df)
