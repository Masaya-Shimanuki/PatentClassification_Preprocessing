# Databricks notebook source
#このノートでは自動的なデータ拡張を実施します

# COMMAND ----------

!pip install openpyxl

# COMMAND ----------

import pandas as pd
import matplotlib.pyplot as plt

# エクセルを読み出し、pandasでデータフレーム形式とする
file_path = '/Volumes/main/default/public/personal/0199667/PatentClassification_Preprocessing/MUIS特許リスト_Japanall軽量版.xlsx'
df = pd.read_excel(file_path, engine="openpyxl")

# カラムの「重要度」を読み出す
if '重要度' not in df.columns:
    raise KeyError("カラムネームをチェックしてください")

# 各ランクの個数をカウントして表示する
importance_counts = df['重要度'].value_counts()
print(importance_counts)

# カウントした結果をグラフにして表示する
importance_counts.plot(kind='bar')
plt.title('Count of Classes')
plt.xlabel('Importance Level')
plt.ylabel('Count')
plt.show()

# COMMAND ----------

# カラムの「重要度」「請求の範囲」「実施例」を読み出す
required_columns = ['重要度', '請求の範囲', '実施例']
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"エクセルファイルに「{col}」列が含まれていません。")

# 「請求の範囲」「実施例」の文字を結合し、「請求の範囲」のカラムに格納する
df['請求の範囲'] = df['請求の範囲'].astype(str) + df['実施例'].astype(str)

# 「重要度」に含まれる各ランクの個数をカウントして表示する
importance_counts = df['重要度'].value_counts()
print(importance_counts)

# 「重要度」に含まれる各ランクの個数の最大値をAug_maxとする
Aug_max = importance_counts.max()
print(f"Aug_max: {Aug_max}")

# 「重要度」に含まれるA,B,Cの個数をカウント
Aug_A = importance_counts['A']
Aug_B = importance_counts['B']
Aug_C = importance_counts['C']
print("A:",Aug_A, "B:",Aug_B, "C:",Aug_C)

# COMMAND ----------

# 操作実行(Aを増やす操作)
import os
import shutil
import datetime

increment = 300  # 削る文字数
th_inc = increment
added_count_A = Aug_A  # 追加されたAの数のカウント(元のデータ数からカウント)

additional_rows = []

# '重要度'が'A'の行を選別
a_rows = df[df['重要度'] == 'A']

for _, row in a_rows.iterrows():
    current_text = row['請求の範囲']
    
    while len(current_text) > th_inc:
        current_text = current_text[increment:]  # 冒頭から300文字削除

        # 新しい行を追加
        new_row = {'重要度': row['重要度'], '請求の範囲': current_text, '実施例': ''}
        additional_rows.append(new_row)
        added_count_A = added_count_A + 1

        # 追加されたAの数がAug_maxを超えた場合、処理を中断
        if added_count_A > Aug_max:
            print("Aug_maxを超えました。処理を中断します。")
            break
      
    # Aug_maxを超えた場合の処理中断
    if added_count_A > Aug_max:
        break

# データフレームに新しい行を追加
df2 = df.append(additional_rows, ignore_index=True)

# ちょっと中身の確認
print(df2.head(5))
print(df2.tail(5))


#DAしたデータ個数の表示

# カラムの「重要度」を読み出す
if '重要度' not in df2.columns:
    raise KeyError("カラムネームをチェックしてください")

# 各アルファベットの個数をカウントして表示する
importance_counts = df2['重要度'].value_counts()
print(importance_counts)

# カウントした結果をグラフにして表示する
importance_counts.plot(kind='bar')
plt.title('Count of Classes')
plt.xlabel('Importance Level')
plt.ylabel('Count')
plt.show()

"""
output_filename = '/Volumes/main/default/public/personal/0199667/PatentClassification_Preprocessing/A増やしてみた.xlsx'
now_= datetime.datetime.now()
timestamp_ = datetime.datetime.timestamp(now_)
tmp_path = f'temp_{timestamp_}.xlsx' # I save it with the timestamp to avoid errors
writer = pd.ExcelWriter(tmp_path)
df2.to_excel(writer,'sheet_name')
writer.save()
shutil.copy(tmp_path, output_filename) # Copy the temp file to mount

os.remove(tmp_path) #仮作成したxlsxを削除
"""

# COMMAND ----------

# 操作実行(Aに引き続き、Bを増やす操作)
"""
# この場合はAを増やしたデータ(df2)を用いている為、Bのみの拡張の場合は上のAを変更ください
# 必要であれば下記を参考に
file_path ='/Volumes/main/default/public/personal/0199667/PatentClassification_Preprocessing/A増やしてみた.xlsx'
df2 = pd.read_excel(file_path, engine="openpyxl")
"""
# '重要度'が'B'の行を選別
a_rows = df2[df2['重要度'] == 'B']

increment = 300  # 削る文字数
th_inc = increment
added_count_B = Aug_B  # 追加されたBの数のカウント

additional_rows = []

for _, row in a_rows.iterrows():
    current_text = row['請求の範囲']
    
    while len(current_text) > th_inc:
        current_text = current_text[increment:]  # 冒頭から300文字削除

        # 新しい行を追加
        new_row = {'重要度': row['重要度'], '請求の範囲': current_text, '実施例': ''}
        additional_rows.append(new_row)
        added_count_B = added_count_B + 1

        # 追加されたAの数がAug_maxを超えた場合、処理を中断
        if added_count_B > Aug_max:
            print("Aug_maxを超えました。処理を中断します。")
            break
      
    # Aug_maxを超えた場合の処理中断
    if added_count_B > Aug_max:
        break

# データフレームに新しい行を追加
df3 = df2.append(additional_rows, ignore_index=True)

# ちょっと中身の確認
print(df3.head(5))
print(df3.tail(5))

#DAしたデータ個数の表示

# カラムの「重要度」を読み出す
if '重要度' not in df3.columns:
    raise KeyError("カラムネームをチェックしてください")

# 各アルファベットの個数をカウントして表示する
importance_counts = df3['重要度'].value_counts()
print(importance_counts)

# カウントした結果をグラフにして表示する
importance_counts.plot(kind='bar')
plt.title('Count of Classes')
plt.xlabel('Importance Level')
plt.ylabel('Count')
plt.show()

"""
output_filename = '/Volumes/main/default/public/personal/0199667/PatentClassification_Preprocessing/AB増やしてみた.xlsx'
now_= datetime.datetime.now()
timestamp_ = datetime.datetime.timestamp(now_)
tmp_path = f'temp_{timestamp_}.xlsx' # I save it with the timestamp to avoid errors
writer = pd.ExcelWriter(tmp_path)
#要確認
df3.to_excel(writer,'sheet_name')
writer.save()
shutil.copy(tmp_path, output_filename) # Copy the temp file to mount

os.remove(tmp_path) #仮作成したxlsxを削除
"""

# COMMAND ----------

# 操作実行(ABに引き続き、Cを増やす操作)
"""
# この場合はABを増やしたデータ(df3)を用いている為、Cのみの拡張の場合は上のAとBを変更ください
# 必要であれば下記を参考に
file_path ='/Volumes/main/default/public/personal/0199667/PatentClassification_Preprocessing/A増やしてみた.xlsx'
df3 = pd.read_excel(file_path, engine="openpyxl")
"""
# '重要度'が'C'の行を選別
a_rows = df3[df3['重要度'] == 'C']

increment = 300  # 削る文字数
th_inc = increment
added_count_C = Aug_C  # 追加されたCの数のカウント

additional_rows = []

for _, row in a_rows.iterrows():
    current_text = row['請求の範囲']
    
    while len(current_text) > th_inc:
        current_text = current_text[increment:]  # 冒頭から300文字削除

        # 新しい行を追加
        new_row = {'重要度': row['重要度'], '請求の範囲': current_text, '実施例': ''}
        additional_rows.append(new_row)
        added_count_C = added_count_C + 1

        # 追加されたAの数がAug_maxを超えた場合、処理を中断
        if added_count_C > Aug_max:
            print("Aug_maxを超えました。処理を中断します。")
            break
      
    # Aug_maxを超えた場合の処理中断
    if added_count_C > Aug_max:
        break

# データフレームに新しい行を追加
df4 = df3.append(additional_rows, ignore_index=True)

# ちょっと中身の確認
print(df4.head(5))
print(df4.tail(5))

#DAしたデータ個数の表示

# カラムの「重要度」を読み出す
if '重要度' not in df4.columns:
    raise KeyError("カラムネームをチェックしてください")

# 各アルファベットの個数をカウントして表示する
importance_counts = df4['重要度'].value_counts()
print(importance_counts)

# カウントした結果をグラフにして表示する
importance_counts.plot(kind='bar')
plt.title('Count of Classes')
plt.xlabel('Importance Level')
plt.ylabel('Count')
plt.show()



# COMMAND ----------

#最終的に出力したデータをエクセルファイルとして出力。Databricksの仕様上、一旦tempファイルに書き出してコピーします

output_filename = '/Volumes/main/default/public/personal/0199667/PatentClassification_Preprocessing/ABC増やしてみた.xlsx'
now_= datetime.datetime.now()
timestamp_ = datetime.datetime.timestamp(now_)
tmp_path = f'temp_{timestamp_}.xlsx' # I save it with the timestamp to avoid errors
writer = pd.ExcelWriter(tmp_path)
#要確認
df4.to_excel(writer,'sheet_name')
writer.save()
shutil.copy(tmp_path, output_filename) # Copy the temp file to mount

os.remove(tmp_path) #仮作成したxlsxを削除

