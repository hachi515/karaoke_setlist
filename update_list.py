import pandas as pd
import requests
import datetime

# 集約したいポート番号のリスト
target_ports = [11021, 11066,11058,11069,11074,11068,11106,11088, 11101] 

all_data_frames = []

for port in target_ports:
    url = f"http://Ykr.moe:{port}/simplelist.php"
    try:
        # データ取得
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        # HTMLの表を読み込み
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0]
            # 識別用の列を先頭に追加（列名: Source）
            df.insert(0, 'Source', f'No.{port}')
            all_data_frames.append(df)
            print(f"Port {port}: OK")
    except Exception as e:
        print(f"Port {port}: Error - {e}")

# HTMLの生成
html_content = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>統合リスト</title>
    <style>
        body { font-family: sans-serif; padding: 20px; }
        table { border-collapse: collapse; width: 100%; font-size: 14px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .update-time { color: #666; font-size: 0.9em; margin-bottom: 10px; }
    </style>
</head>
<body>
    <h1>統合リスト</h1>
"""

# 更新時刻の追加
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))) # 日本時間
html_content += f'<div class="update-time">最終更新: {now.strftime("%Y/%m/%d %H:%M:%S")}</div>'

if all_data_frames:
    final_df = pd.concat(all_data_frames, ignore_index=True)
    # pandasの機能でHTMLテーブルに変換 (classを指定しない素のHTML)
    table_html = final_df.to_html(index=False, border=0, escape=False)
    html_content += table_html
else:
    html_content += "<p>データの取得に失敗しました。</p>"

html_content += "</body></html>"

# index.html として保存
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)