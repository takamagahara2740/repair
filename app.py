import pandas as pd
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
import altair as alt
import matplotlib.pyplot as plt
import plotly.figure_factory as ff
import plotly.graph_objects as go
import pydeck as pdk
import streamlit as st
import datetime
import time
from dateutil.relativedelta import relativedelta

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
credentials = Credentials.from_service_account_file(
    'service_account.json',
    scopes=scopes
)
gc = gspread.authorize(credentials)

SP_SHEET_KEY = '1UZ0SX4xeej3b0LdHbt6Ui1j-GB3xcnv6wP5f_ot7RN8'
sh = gc.open_by_key(SP_SHEET_KEY)

def get_df_M(sh):
  SP_SHEET = '修理受付'
  worksheet = sh.worksheet(SP_SHEET)
  data = worksheet.get_all_values()
  df_M = pd.DataFrame(data[1:], columns=data[0])
  return df_M
df_M = get_df_M(sh)

def get_df_G(sh):
  SP_SHEET = 'ゲーム修理受付'
  worksheet = sh.worksheet(SP_SHEET)
  data = worksheet.get_all_values()
  df_G = pd.DataFrame(data[1:], columns=data[0])
  return df_G
df_G = get_df_G(sh)

def get_update_time(sh):
  SP_SHEET = '更新日時'
  worksheet = sh.worksheet(SP_SHEET)
  update_time = worksheet.acell('A1').value
  return update_time
update_time = get_update_time(sh)

st.title('スマホ修理工房ゆめタウン光の森店')

st.sidebar.write(f"""
                 ## データ抽出期間
                 ### 更新日時：{update_time}
                 """)
reload =  st.sidebar.button('データの更新')

start_day = st.sidebar.date_input('開始日', datetime.date.today() + relativedelta(day=1), min_value=datetime.date(2022, 9, 1), max_value=datetime.date.today())
set_start_day = datetime.datetime(start_day.year, start_day.month, start_day.day)
end_day = st.sidebar.date_input('終了日', datetime.date.today(), min_value=start_day, max_value=datetime.date.today())
set_end_day = datetime.datetime(end_day.year, end_day.month, end_day.day)

spans = {
    '日別実績': 'D',
    '週間実績': 'W',
    '月間実績': 'M',
    '四半期実績': 'Q',
    '年間実績': 'Y',
}
span = st.sidebar.selectbox("集計期間", spans) 

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(['売上高', '来客住所', '機種区分', 'バッテリー', 'フロントパネル', '再修理'])

with tab1:
  st.subheader('売上高')
  def get_Earnings(df_M, df_G):
    # モバイルの実績の抽出（売上高）
    df_M = df_M[['受付日::fm883__free24','総額（税込）::fm883__free64']]
    df_M = df_M.rename(columns={'受付日::fm883__free24': '受付日', '総額（税込）::fm883__free64': 'モバイル'})
    df_M['受付日'] = pd.to_datetime(df_M['受付日'])
    df_M = df_M.dropna()
    df_M['モバイル'] = df_M['モバイル'].astype(int)
    df_M = df_M.groupby('受付日').sum()
    # ゲーム機の実績の抽出（売上高）
    df_G = df_G[['受付日::fm886__free24','総額（税込）::fm886__free206']]
    df_G = df_G.rename(columns={'受付日::fm886__free24': '受付日', '総額（税込）::fm886__free206': 'ゲーム機'})
    df_G['受付日'] = pd.to_datetime(df_G['受付日'])
    df_G = df_G.dropna()
    df_G['ゲーム機'] = df_G['ゲーム機'].astype(int)
    df_G = df_G.groupby('受付日').sum()
    # モバイルとゲーム機の実績を結合する
    df_Earnings = pd.concat([df_M, df_G], axis=1, join='outer')
    df_Earnings = df_Earnings.fillna(0)
    df_Earnings['ゲーム機'] = df_Earnings['ゲーム機'].astype(int)
    df_Earnings['売上合計'] = df_Earnings['モバイル'] + df_Earnings['ゲーム機']
    # 集計期間の実績の抽出
    df_Earnings = df_Earnings[(df_Earnings.index >= set_start_day) & (df_Earnings.index <= set_end_day)]
    df_Earnings = df_Earnings.resample(spans[span]).sum()
    # 集計期間に合わせて日付をフォーマットする関数を設定
    def change_index():
      if span == '年間実績':
        df_Earnings.index = df_Earnings.index.strftime('%Y')
        return df_Earnings.index
      elif span == '四半期実績':
        df_Earnings.index = df_Earnings.index.strftime('%Y/%m')
        return df_Earnings.index
      elif span == '月間実績':
        df_Earnings.index = df_Earnings.index.strftime('%Y/%m')
        return df_Earnings.index
      else:
        df_Earnings.index = df_Earnings.index.strftime('%m/%d')
        return df_Earnings.index
    # 集計期間に合わせて日付をフォーマットする関数を起動する
    change_index()
    # グラフ化するデータを抽出
    chart = df_Earnings[['ゲーム機', 'モバイル']]
    # グラフとデータフレームをフロントサイドに表示する
    st.bar_chart(chart)
    st.write(df_Earnings)
  get_Earnings(df_M, df_G)

with tab2:
  st.subheader('来客者住所')
  def get_Adress(df_M, df_G):
    # モバイルの実績の抽出（来客住所）
    df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25', '郵便番号::fm883__free163']]
    df_M = df_M.rename(columns={'受付日::fm883__free24': '受付日', '受付種別::fm883__free25': '受付種別', '郵便番号::fm883__free163': '郵便番号'})
    df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
    df_M = df_M.dropna()
    df_M['郵便番号'] = df_M['郵便番号'].replace(['〒','-','ー',' ','　','〜'],'', regex=True)
    df_M = df_M[df_M['受付種別'].isin(['来店'])]
    df_M = df_M[['受付日','郵便番号']]
    df_M = df_M[df_M['郵便番号'] != '']
    df_M['郵便番号'] = df_M['郵便番号'].astype(int)
    # ゲーム機のの実績の抽出（来客住所）
    df_G = df_G[['受付日::fm886__free24','受付種別::fm886__free25', '郵便番号::fm886__free163']]
    df_G = df_G.rename(columns={'受付日::fm886__free24': '受付日', '受付種別::fm886__free25': '受付種別', '郵便番号::fm886__free163': '郵便番号'})
    df_G['受付日'] = pd.to_datetime(df_G['受付日'], format='%Y/%m/%d')
    df_G = df_G.dropna()
    df_G['郵便番号'] = df_G['郵便番号'].replace(['〒','-','ー',' ','　','〜'],'', regex=True)
    df_G = df_G[df_G['受付種別'].isin(['来店'])]
    df_G = df_G[['受付日','郵便番号']]
    df_G = df_G[df_G['郵便番号'] != '']
    df_G['郵便番号'] = df_G['郵便番号'].astype(int)
    # モバイルとゲーム機の実績を結合する
    df_M_G = pd.concat([df_M, df_G], axis=0)
    df_M_G = df_M_G[(df_M_G['受付日'] >= set_start_day) & (df_M_G['受付日'] <= set_end_day)]
    # 郵便番号・緯度経度・住所表記のデータを読み込む
    df_Adress_Data = pd.read_csv("Zip2Geoc.csv")
    df_Adress_Data = df_Adress_Data.rename(columns={'Zip':'郵便番号','Lon':'lon','Lat':'lat'})
    df_Adress_Data['住所'] = df_Adress_Data['Pre'] + df_Adress_Data['City'] + df_Adress_Data['Addr']
    # 来客住所のデータと緯度経度のデータを結合する
    df_Lon_Lat = df_Adress_Data[['郵便番号','lon','lat']]
    df_Adress = pd.merge(df_M_G, df_Lon_Lat, on='郵便番号', how='left')
    df_Adress = df_Adress.dropna()
    # グラフ化するデータを抽出
    df_Adress = df_Adress[['lon','lat']]
    # グラフ化してフロントサイドに表示する
    st.pydeck_chart(pdk.Deck(
        map_style=None,
        initial_view_state=pdk.ViewState(
            latitude=32.82,
            longitude=130.75,
            zoom=10,
            pitch=60,
        ),
        layers=[
            pdk.Layer(
              'HexagonLayer',
              data=df_Adress,
              get_position='[lon, lat]',
              radius=1000,
              elevation_scale=4,
              elevation_range=[0, 2000],
              pickable=True,
              extruded=True,
            ),
            pdk.Layer(
                'ScatterplotLayer',
                data=df_Adress,
                get_position='[lon, lat]',
                get_color='[200, 30, 0, 160]',
                get_radius=2000,
            ),
        ],
    ))
    # 来客住所のデータと住所表記のデータを結合する
    df_Adrr = df_Adress_Data[['郵便番号','住所']]
    df_Adrr = pd.merge(df_M_G, df_Adrr, on='郵便番号', how='left')
    df_Adrr = df_Adrr.groupby('住所').count()
    df_Adrr = df_Adrr[['受付日']].rename(columns={'受付日':'来客数'}).sort_values('来客数', ascending=False)
    # データフレームをフロントサイドに表示する
    st.write(df_Adrr)
  get_Adress(df_M, df_G)

with tab3:
  tabA, tabB, tabC, tabD = st.tabs(['全体', 'iPhone', 'iPad', 'Android'])
  with tabA:
    st.subheader('【機種区分】')
    def get_Prodact(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種区分']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種区分'], values='機種区分', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # 集計するカテゴリを選択する
      df_M = df_M.T
      count = 100
      category = st.multiselect('カテゴリーを選択してください。',list(df_M.index),list(df_M.index), key=count)
      count += 1
      if not category:
        st.error('少なくともひとつはカテゴリを選んで下さい。')
      else:
        df_M = df_M.loc[category]
        df_M = df_M.T
        # グラフとデータフレームをフロントサイドに表示する
        st.bar_chart(df_M)
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_Prodact(df_M)
  
  with tabB:
    def get_iPhone(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M = df_M[df_M['機種区分'].isin(['iPhone'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【iPhone】機種別販売数')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【iPhone】 集計期間別推移')
      count = 110
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_iPhone(df_M)
  
  with tabC:
    def get_iPad(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M = df_M[df_M['機種区分'].isin(['iPad'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【iPad】機種別販売数')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【iPad】 集計期間別推移')
      count = 120
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_iPad(df_M)
  
  with tabD:
    def get_Android(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M = df_M[df_M['機種区分'].isin(['Android'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【Android】機種別販売数')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【Android】 集計期間別推移')
      count = 130
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_Android(df_M)

with tab4:
  tabA, tabB, tabC, tabD = st.tabs(['全体', 'iPhone', 'iPad', 'Android'])
  with tabA:
    st.subheader('バッテリー交換件数')
    def get_BT(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'バッテリー：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'バッテリー：件数':'バッテリー'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['バッテリー'] = df_M['バッテリー'].astype(int)
      df_M = df_M[(df_M['バッテリー'] == 1)]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種区分']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種区分'], values='機種区分', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # 集計するカテゴリを選択する
      df_M = df_M.T
      count = 200
      category = st.multiselect('カテゴリーを選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not category:
        st.error('少なくともひとつはカテゴリを選んで下さい。')
      else:
        df_M = df_M.loc[category]
        df_M = df_M.T
        # グラフとデータフレームをフロントサイドに表示する
        st.bar_chart(df_M)
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_BT(df_M)
  
  with tabB:
    def get_BT_iPhone(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'バッテリー：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'バッテリー：件数':'バッテリー'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['バッテリー'] = df_M['バッテリー'].astype(int)
      df_M = df_M[(df_M['バッテリー'] == 1)]
      df_M = df_M[df_M['機種区分'].isin(['iPhone'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【iPhone】機種別受注数（バッテリー交換）')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【iPhone】 集計期間別推移（バッテリー交換）')
      count = 210
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_BT_iPhone(df_M)
  
  with tabC:
    def get_BT_iPad(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'バッテリー：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'バッテリー：件数':'バッテリー'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['バッテリー'] = df_M['バッテリー'].astype(int)
      df_M = df_M[(df_M['バッテリー'] == 1)]
      df_M = df_M[df_M['機種区分'].isin(['iPad'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【iPad】機種別受注数（バッテリー交換）')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【iPad】 集計期間別推移（バッテリー交換）')
      count = 220
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_BT_iPad(df_M)
  
  with tabD:
    def get_BT_Android(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'バッテリー：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'バッテリー：件数':'バッテリー'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['バッテリー'] = df_M['バッテリー'].astype(int)
      df_M = df_M[(df_M['バッテリー'] == 1)]
      df_M = df_M[df_M['機種区分'].isin(['Android'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【Android】機種別受注数（バッテリー交換）')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【Android】 集計期間別推移（バッテリー交換）')
      count = 230
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_BT_Android(df_M)
  
with tab5:
  tabA, tabB, tabC, tabD = st.tabs(['全体', 'iPhone', 'iPad', 'Android'])
  with tabA:
    st.subheader('画面修理件数')
    def get_FP(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'フロントパネル：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'フロントパネル：件数':'フロントパネル'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['フロントパネル'] = df_M['フロントパネル'].astype(int)
      df_M = df_M[(df_M['フロントパネル'] == 1)]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種区分']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種区分'], values='機種区分', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # 集計するカテゴリを選択する
      df_M = df_M.T
      count = 300
      category = st.multiselect('カテゴリーを選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not category:
        st.error('少なくともひとつはカテゴリを選んで下さい。')
      else:
        df_M = df_M.loc[category]
        df_M = df_M.T
        # グラフとデータフレームをフロントサイドに表示する
        st.bar_chart(df_M)
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_FP(df_M)
  
  with tabB:
    def get_FP_iPhone(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'フロントパネル：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'フロントパネル：件数':'フロントパネル'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['フロントパネル'] = df_M['フロントパネル'].astype(int)
      df_M = df_M[(df_M['フロントパネル'] == 1)]
      df_M = df_M[df_M['機種区分'].isin(['iPhone'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【iPhone】機種別受注数（画面修理件数）')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【iPhone】 集計期間別推移（画面修理件数）')
      count = 310
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_FP_iPhone(df_M)
  
  with tabC:
    def get_FP_iPad(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'フロントパネル：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'フロントパネル：件数':'フロントパネル'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['フロントパネル'] = df_M['フロントパネル'].astype(int)
      df_M = df_M[(df_M['フロントパネル'] == 1)]
      df_M = df_M[df_M['機種区分'].isin(['iPad'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【iPad】機種別受注数（画面修理件数）')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【iPad】 集計期間別推移（画面修理件数）')
      count = 320
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_FP_iPad(df_M)
  
  with tabD:
    def get_FP_Android(df_M):
      df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', 'フロントパネル：件数']]
      df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別', 'フロントパネル：件数':'フロントパネル'})
      df_M = df_M[df_M['受付種別'].isin(['来店'])]
      df_M['フロントパネル'] = df_M['フロントパネル'].astype(int)
      df_M = df_M[(df_M['フロントパネル'] == 1)]
      df_M = df_M[df_M['機種区分'].isin(['Android'])]
      df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
      df_M = df_M.dropna()
      df_M = df_M[['受付日','機種名']]
      df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
      df_M = df_M.pivot_table(index=['受付日'], columns=['機種名'], values='機種名', aggfunc=len)
      df_M = df_M.fillna(0)
      df_M = df_M.astype('int')
      df_M = df_M.resample(spans[span]).sum()
      # 集計期間に合わせて日付をフォーマットする関数を設定
      def change_index():
        if span == '年間実績':
          df_M.index = df_M.index.strftime('%Y')
          return df_M.index
        elif span == '四半期実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return ddf_M
        elif span == '月間実績':
          df_M.index = df_M.index.strftime('%Y/%m')
          return df_M.index
        else:
          df_M.index = df_M.index.strftime('%m/%d')
          return df_M.index
      # 集計期間に合わせて日付をフォーマットする関数を起動する
      change_index()
      # モデル別販売数
      st.subheader('【Android】機種別受注数（画面修理件数）')
      df_M = df_M.T
      st.bar_chart(df_M)
      # 集計する機種を選択する
      st.subheader('【Android】 集計期間別推移（画面修理件数）')
      count = 330
      model = st.multiselect('機種を選択してください。',list(df_M.index),list(df_M.index),key=count)
      count += 1
      if not model:
        st.error('少なくともひとつは機種を選んで下さい。')
      else:
        # グラフとデータフレームをフロントサイドに表示する
        df_M = df_M.loc[model]    
        df_M = df_M.T
        st.line_chart(df_M)
        df_M.loc['合計'] = df_M.sum()
        df_M = df_M.T
        st.write(df_M)
    get_FP_Android(df_M)
    
with tab6:
  df_M = df_M[['受付日::fm883__free24','受付種別::fm883__free25','機種区分', '機種名', '受付内容【1】::fm883__free28','処理内容【1】::fm883__free49','フロントパネル：件数','バッテリー：件数']]
  df_M = df_M.rename(columns={'受付日::fm883__free24':'受付日','受付種別::fm883__free25':'受付種別','受付内容【1】::fm883__free28':'受付内容','処理内容【1】::fm883__free49':'処理内容'})
  df_M = df_M[df_M['受付種別'].isin(['再修理'])]
  df_M['受付日'] = pd.to_datetime(df_M['受付日'], format='%Y/%m/%d')
  df_M = df_M.dropna()
  df_M = df_M[(df_M['受付日'] >= set_start_day) & (df_M['受付日'] <= set_end_day)]
  
  tabA, tabB, tabC = st.tabs(['全体', 'バッテリー', 'フロントパネル'])
  
  with tabA:
    st.subheader('再修理案件一覧')
    pv = df_M[['機種名','処理内容']]
    pv = pv.pivot_table(index=['機種名'], columns=['処理内容'], values='処理内容', aggfunc=len)
    pv = pv.fillna(0)
    pv = pv.astype('int')
    pv = pv.T
    pv.loc['合計'] = pv.sum()
    pv = pv.T
    pv = pv.sort_values('合計', ascending=False)
    st.write(pv)
  
  with tabB:
    try:
      st.subheader('【バッテリー】再修理件数')
      df_M['バッテリー：件数'] = df_M['バッテリー：件数'].astype(int)
      df_BT = df_M[(df_M['バッテリー：件数'] == 1)]
      df_BT = df_BT[['機種名','処理内容']]
      df_BT = df_BT.pivot_table(index=['機種名'], columns=['処理内容'], values='処理内容', aggfunc=len)
      df_BT = df_BT.fillna(0)
      df_BT = df_BT.astype('int')
      df_BT = df_BT.T
      df_BT.loc['合計'] = df_BT.sum()
      df_BT = df_BT.T
      df_BT = df_BT.sort_values('合計', ascending=False)
      st.write(df_BT)
    except:
      st.write('### No Data')
  
  with tabC:
    try:
      st.subheader('【フロントパネル】再修理件数')
      df_M['フロントパネル：件数'] = df_M['フロントパネル：件数'].astype(int)
      df_FP = df_M[(df_M['フロントパネル：件数'] == 1)]
      df_FP = df_FP[['機種名','処理内容']]
      df_FP = df_FP.pivot_table(index=['機種名'], columns=['処理内容'], values='処理内容', aggfunc=len)
      df_FP = df_FP.fillna(0)
      df_FP = df_FP.astype('int')
      df_FP = df_FP.T
      df_FP.loc['合計'] = df_FP.sum()
      df_FP = df_FP.T
      df_FP = df_FP.sort_values('合計', ascending=False)
      st.write(df_FP)
    except:
      st.write('### No Data')
