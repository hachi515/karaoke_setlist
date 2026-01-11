import pandas as pd
import requests
import datetime
import re

# --- 時刻設定（データ取得日として使用） ---
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
current_date_str = now.strftime("%Y/%m/%d") # 日付のみ（例: 2026/01/11）
current_datetime_str = now.strftime("%Y/%m/%d %H:%M") # 日時（例: 2026/01/11 14:00）

# --- ① 設定: ポート番号と部屋主の名前の対応表 ---
room_map = {
    11000: "ゆーふうりん部屋",
    11001: "ゆーふうりん部屋",
    11002: "ゆーふうりん部屋",
    11003: "ゆーふうりん部屋",
    11004: "ゆーふうりん部屋",
    11005: "ゆーふうりん部屋",
    11006: "ゆーふうりん部屋",
    11007: "ゆーふうりん部屋",
    11008: "ゆーふうりん部屋",
    11009: "ゆーふうりん部屋",
    11021: "成田部屋",
    11022: "成田部屋",
    11028: "タマ部屋",
    11058: "すみた部屋",
    11059: "つぼはち部屋",
    11063: "なぎ部屋",
    11064: "naoo部屋",
    11066: "芝ちゃん部屋",
    11067: "crom部屋",
    11068: "けんしん部屋",
    11069: "けんちぃ部屋",
    11070: "黒河部屋",
    11071: "黒河部屋",
    11074: "tukinowa部屋",
    11077: "v3部屋",
    11078: "のんでるん部屋",
    11079: "まどか部屋",
    11084: "タカヒロ部屋",
    11085: "タカヒロ部屋",
    11086: "タカヒロ部屋",
    11088: "ほっしー部屋",
    11101: "えみち部屋",
    11102: "るえ部屋",
    11103: "ながし部屋",
    11106: "冨塚部屋"
}

target_ports = list(room_map.keys())
all_data_frames = []

for port in target_ports:
    # URL設定: requestlist_only.php
    url = f"http://Ykr.moe:{port}/requestlist_only.php"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0]
            df = df.fillna("") # 空欄処理
            
            # 項目追加
            df['部屋主'] = room_map[port]
            df['取得日'] = current_date_str
            
            all_data_frames.append(df)
            print(f"Port {port} ({room_map[port]}): OK")
            
    except Exception as e:
        # エラーが出ても止まらずに次のポートへ
        print(f"Port {port}: Error - {e}")

# HTML生成
html_content = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Karaoke setlist all</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body { font-family: "Helvetica Neue", Arial, sans-serif; padding: 20px; color: #333; }
        h1 { margin-bottom: 10px; }
        
        /* 検索ボックスエリア */
        .search-container { margin-bottom: 15px; padding: 15px; background: #f8f9fa; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
        .search-box {
            width: 100%; max-width: 400px; padding: 10px; font-size: 16px;
            border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;
        }
        .btn {
            padding: 10px 20px; font-size: 16px; cursor: pointer;
            background-color: #007bff; color: white; border: none; border-radius: 4px;
            margin-left: 5px;
        }
        .btn:hover { background-color: #0056b3; }
        .btn-reset { background-color: #6c757d; }
        .btn-reset:hover { background-color: #545b62; }

        /* スクロール表示用のコンテナ */
        .table-wrapper {
            max-height: 80vh; /* 画面の高さの80%まで表示、それ以上はスクロール */
            overflow-y: auto; /* 縦スクロールを有効化 */
            border: 1px solid #ddd;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }

        table { border-collapse: collapse; width: 100%; font-size: 14px; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        
        /* ヘッダー固定 */
        th { 
            background-color: #f2f2f2; 
            position: sticky; 
            top: 0; 
            z-index: 10; 
            cursor: pointer;
            box-shadow: 0 2px 2px -1px rgba(0, 0, 0, 0.1);
        }
        th:hover { background-color: #e2e2e2; }
        tr:nth-child(even) { background-color: #fff; }

        .update-time { color: #666; font-size: 0.9em; margin-bottom: 10px; }
        
        @media (max-width: 600px) {
            .btn { width: 100%; margin: 5px 0 0 0; }
            .table-wrapper { max-height: 70vh; }
        }
    </style>
</head>
<body>
    <h1>Karaoke setlist all</h1>
"""

html_content += f'<div class="update-time">最終集計: {current_datetime_str}</div>'

html_content += f'''
    <div class="search-container">
        <input type="text" id="searchInput" class="search-box" placeholder="キーワード・日付で検索 (例: 2026/01/11 King)...">
        <button onclick="filterTable()" class="btn"><i class="fas fa-search"></i> 検索</button>
        <button onclick="resetFilter()" class="btn btn-reset">リセット</button>
    </div>
'''

if all_data_frames:
    final_df = pd.concat(all_data_frames, ignore_index=True)

    # 順番でソート
    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
        final_df = final_df.sort_values(by=['順番'], ascending=True)

    # カラム並び替え: 部屋主を先頭へ
    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    # テーブル作成（スクロール枠で囲む）
    html_content += '<div class="table-wrapper"><table id="setlistTable">'
    
    # ヘッダー
    html_content += '<thead><tr>'
    for col in final_df.columns:
        html_content += f'<th onclick="sortTable({list(final_df.columns).index(col)})">{col} <i class="fas fa-sort"></i></th>'
    html_content += '</tr></thead>'
    
    # ボディ
    html_content += '<tbody>'
    for _, row in final_df.iterrows():
        html_content += '<tr>'
        for val in row:
            html_content += f'<td>{val}</td>'
        html_content += '</tr>'
    html_content += '</tbody></table></div>'

else:
    html_content += "<p>データの取得に失敗しました。</p>"

# JavaScript (検索・ソート)
html_content += """
<script>
    function filterTable() {
        var input = document.getElementById("searchInput");
        var filter = input.value.toUpperCase();
        var keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        
        var table = document.getElementById("setlistTable");
        var tr = table.getElementsByTagName("tr");

        for (var i = 1; i < tr.length; i++) {
            var td = tr[i].getElementsByTagName("td");
            var rowText = "";
            for (var j = 0; j < td.length; j++) {
                rowText += td[j].textContent || td[j].innerText;
            }
            rowText = rowText.toUpperCase();
            
            var isMatch = true;
            for (var k = 0; k < keywords.length; k++) {
                if (rowText.indexOf(keywords[k]) === -1) {
                    isMatch = false;
                    break;
                }
            }
            if (isMatch || keywords.length === 0) {
                tr[i].style.display = "";
            } else {
                tr[i].style.display = "none";
            }
        }
    }

    document.getElementById("searchInput").addEventListener("keyup", function(event) {
        if (event.key === "Enter") filterTable();
    });

    function resetFilter() {
        document.getElementById("searchInput").value = "";
        filterTable();
    }

    function sortTable(n) {
        var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
        table = document.getElementById("setlistTable");
        switching = true;
        dir = "asc"; 
        while (switching) {
            switching = false;
            rows = table.rows;
            for (i = 1; i < (rows.length - 1); i++) {
                shouldSwitch = false;
                x = rows[i].getElementsByTagName("TD")[n];
                y = rows[i + 1].getElementsByTagName("TD")[n];
                var xContent = x.innerText.toLowerCase();
                var yContent = y.innerText.toLowerCase();
                if (!isNaN(xContent) && !isNaN(yContent) && xContent !== "" && yContent !== "") {
                     if (dir == "asc") { if (Number(xContent) > Number(yContent)) { shouldSwitch = true; break; } }
                     else if (dir == "desc") { if (Number(xContent) < Number(yContent)) { shouldSwitch = true; break; } }
                } else {
                    if (dir == "asc") { if (xContent > yContent) { shouldSwitch = true; break; } }
                    else if (dir == "desc") { if (xContent < yContent) { shouldSwitch = true; break; } }
                }
            }
            if (shouldSwitch) {
                rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                switching = true;
                switchcount ++;      
            } else {
                if (switchcount == 0 && dir == "asc") { dir = "desc"; switching = true; }
            }
        }
    }
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
