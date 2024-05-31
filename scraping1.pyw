import PySimpleGUI as sg
import requests, pandas as pd
import bs4
import re

#【設定用】
title = "試験.comスクレイピング用(ITパスポート～応用情報)"
savingFolder = "C:\\"
loadingURL, value1 = "読み込むURL", "https://www.ap-siken.com/kakomon/06_haru/"
# ----------------------------

# Webスクレイピング
def ScrapingText(loadingURL, savefolder):
  # サイトからHTMLデータを取得
  html = requests.get(loadingURL)
  soup = bs4.BeautifulSoup(html.text, 'lxml')

  # 問題番号と問題を格納するリストを作成
  problem_numbers = []
  problem_titles = []
  problem_subject = []

  TitleTag = GetTitle(soup.find('title').text)

  # テーブルから問題番号と問題を抽出し、リストに格納
  rows = soup.find('table', class_='qtable').find_all('td')
  for i in range(0, len(rows), 4):
      problem_numbers.append(rows[i].text.strip())  # 問題番号
      problem_titles.append(rows[i+1].text.strip())  # 問題タイトル
      problem_subject.append(rows[i+2].text.strip())

  # pandas DataFrameにデータを格納
  df = pd.DataFrame({
      '問題番号': problem_numbers,
      '問題': problem_titles,
      '科目': problem_subject
  })

  # Excel出力
  df.to_excel(savefolder + '/' + TitleTag[0] + '_' + TitleTag[1] + '_午前.xlsx', index=False)
#--------------------
def GetTitle(text):
  # 正規表現にマッチする部分文字列を抜き出す
  examName = text.split(' ')[0]
  year = re.search(r'\((.*?)\)', text)

  if year:
    return (examName, year.group(1))

#アプリのレイアウト
layout = [[sg.Text(loadingURL, size=(14,1)), sg.Input(value1, key="loadingURL")],
    		  [sg.Text("保存先のフォルダ", size=(14,1)),
           sg.Input(savingFolder, key="savingFolder"),sg.FolderBrowse("選択")],
          [sg.Button("実行", size=(20,1), pad=(5,15), bind_return_key=True)]]

#アプリの実行処理
window = sg.Window(title, layout, font=(None,14))
while True:
    event, values = window.read()
    if event == None:
        break
    if event == "実行":
      ScrapingText(values["loadingURL"], values["savingFolder"])
window.close()