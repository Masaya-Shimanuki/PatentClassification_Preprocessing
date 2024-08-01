# Databricks notebook source
!pip install openpyxl

# COMMAND ----------

import pandas as pd
import matplotlib.pyplot as plt

# ①エクセルを読み出し、pandasでデータフレーム形式とする
file_path = '/Volumes/main/default/public/personal/0199667/PatentClassification_Preprocessing/MUIS特許リスト_Japanall.xlsx'
df = pd.read_excel(file_path, engine="openpyxl")

# ②カラムの「重要度」を読み出す
if '重要度' not in df.columns:
    raise KeyError("カラムネームをチェックしてください")

# ③各アルファベットの個数をカウントして表示する
importance_counts = df['重要度'].value_counts()
print(importance_counts)

# ④カウントした結果をグラフにして表示する
importance_counts.plot(kind='bar')
plt.title('Count of Classes')
plt.xlabel('Importance Level')
plt.ylabel('Count')
plt.show()

# COMMAND ----------

required_columns = ['重要度', '請求の範囲', '実施例']
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"エクセルファイルに「{col}」列が含まれていません。")

# 文字数を計算
df['請求の範囲文字数'] = df['請求の範囲'].astype(str).map(len)
df['実施例文字数'] = df['実施例'].astype(str).map(len)

# 重要度の各ラベルごとの文字数統計を計算し出力
importance_stats = df.groupby('重要度').agg({
    '請求の範囲文字数': ['mean', 'min', 'max'],
    '実施例文字数': ['mean', 'min', 'max']
})

print(importance_stats)

# データ全体の「請求の範囲」と「実施例」を合計した文字数のヒストグラムを表示
df['合計文字数'] = df['請求の範囲文字数'] + df['実施例文字数']

plt.hist(df['合計文字数'], bins=20, edgecolor='black')
plt.title('histgram of total characters')
plt.xlabel('total characters')
plt.ylabel('frequency')
plt.show()
