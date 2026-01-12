import pandas as pd
import requests
import datetime
import os

# --- 時刻設定 ---
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
current_date_str = now.strftime("%Y/%m/%d") 
current_datetime_str = now.strftime("%Y/%m/%d %H:%M")

# --- 設定: ポート番号と部屋主の名前の対応表 ---
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
    11087: "MiO部屋",
    11088: "ほっしー部屋",
    11101: "えみち部屋",
    11102: "るえ部屋",
    11103: "ながし部屋",
    11106: "冨塚部屋"
}

# --- 1. 過去のデータ(history.csv)を読み込む ---
history_file = "history.csv"
if os.path.exists(history_file):
    try:
        history_df = pd.read_csv(history_file, encoding='utf-8-sig')
        # nanを空欄に変換
        history_df = history_df.fillna("")
        if '順番' in history_df.columns:
            history_df['順番'] = pd.to_numeric(history_df['順番'], errors='coerce')
    except Exception as e:
        print(f"履歴ファイルの読み込みエラー: {e}")
        history_df = pd.DataFrame()
else:
    history_df = pd.DataFrame()

# --- 2. 新しいデータを取得する ---
target_ports = list(room_map.keys())
new_data_frames = []

for port in target_ports:
    url = f"http://Ykr.moe:{port}/simplelist.php"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0]
            df = df.fillna("") 
            
            # 項目追加
            df['部屋主'] = room_map[port]
            df['取得日'] = current_date_str
            
            new_data_frames.append(df)
            print(f"Port {port} ({room_map[port]}): OK")
            
    except Exception as e:
        print(f"Port {port}: Error - {e}")

# 新しいデータがある場合のみ処理
if new_data_frames:
    new_df = pd.concat(new_data_frames, ignore_index=True)
    if '順番' in new_df.columns:
        new_df['順番'] = pd.to_numeric(new_df['順番'], errors='coerce')

    # --- 3. 過去データと結合 & 重複排除 ---
    combined_df = pd.concat([history_df, new_df], ignore_index=True)
    
    # 重複判定
    subset_cols = ['取得日', '部屋主', '曲名（ファイル名）', '歌った人', '作品名', '歌手名']
    existing_cols = [c for c in subset_cols if c in combined_df.columns]
    
    final_df = combined_df.drop_duplicates(subset=existing_cols, keep='last')
    final_df = final_df.fillna("")

    # ソート
    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
        
    sort_cols = ['取得日', '順番']
    final_df = final_df.sort_values(by=sort_cols, ascending=[False, False])
    
    # 列の整理
    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    # --- 4. 履歴ファイル(CSV)を更新保存 ---
    # ここでは全てのデータを保存します
    final_df.to_csv(history_file, index=False, encoding='utf-8-sig')
    print("履歴ファイルを更新しました。")

else:
    final_df = history_df
    print("新しいデータが取得できませんでした。過去のデータを使用します。")


# --- 5. HTML表示用データの準備 ---
# ★ここで「コメント」列を除外設定しました★
columns_to_hide = ['コメント'] 

if not final_df.empty:
    # 指定された列を削除したHTML用のデータを作成（CSV用のfinal_dfはそのまま）
    html_df = final_df.drop(columns=columns_to_hide, errors='ignore')
else:
    html_df = pd.DataFrame()


# --- HTML生成 ---
# CSSの {} は {{ }} にエスケープし、変数 {val} は一重の { } で記述します。

html_content = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Karaoke setlist all</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        /* ベース設定 */
        html, body {{
            height: 100%;
            margin: 0;
            padding: 0;
            overflow: hidden;
            font-family: "Helvetica Neue", "Arial", "Hiragino Kaku Gothic ProN", "Hiragino Sans", "Meiryo", sans-serif;
            background-color: #f8f9fa;
            color: #333;
            font-size: 14px; 
            display: flex;
            flex-direction: column;
        }}

        /* ヘッダー */
        .header-area {{
            flex: 0 0 auto;
            padding: 12px 20px;
            background-color: #fff;
            border-bottom: 1px solid #ddd;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            z-index: 20;
        }}

        .title-row {{
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            flex-wrap: wrap;
            margin-bottom: 10px;
        }}

        h1 {{ 
            margin: 0; 
            font-size: 1.4rem; 
            font-weight: bold;
            color: #2c3e50;
        }}
        
        .update-time {{ color: #7f8c8d; font-size: 0.9em; }}

        /* コントロール */
        .controls-row {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: center;
        }}

        .search-container {{
            display: flex;
            flex: 1; 
            gap: 8px;
            align-items: center;
            max-width: 600px;
        }}

        .search-box {{
            flex: 1;
            padding: 8px 12px;
            font-size: 14px;
            border: 1px solid #ccc;
            border-radius: 4px;
            background-color: #fff;
        }}
        .search-box:focus {{ outline: 2px solid #007bff; border-color: transparent; }}

        .btn {{
            padding: 8px 16px;
            font-size: 14px;
            cursor: pointer;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            white-space: nowrap;
            font-weight: bold;
            transition: opacity 0.2s;
        }}
        .btn:hover {{ opacity: 0.9; }}
        .btn-reset {{ background-color: #6c757d; }}

        .count-display {{
            margin-left: auto;
            font-weight: bold;
            color: #555;
            white-space: nowrap;
        }}

        /* テーブルラッパー */
        .table-wrapper {{
            flex: 1 1 auto;
            overflow: auto;
            position: relative;
            background-color: #fff;
            -webkit-overflow-scrolling: touch;
            padding-bottom: 20px;
        }}

        /* テーブル設定 */
        table {{ 
            border-collapse: separate; 
            border-spacing: 0; 
            width: 100%; 
            table-layout: fixed; 
            min-width: 900px; 
        }}
        
        th, td {{ 
            padding: 8px 10px;
            text-align: left; 
            border-bottom: 1px solid #eee;
            vertical-align: middle;
            line-height: 1.3;
            white-space: normal; 
            word-break: break-all; 
            overflow-wrap: break-word;
        }}

        /* ヘッダー固定 */
        th {{ 
            background-color: #f1f3f5;
            color: #444;
            font-weight: bold;
            position: sticky; 
            top: 0; 
            z-index: 10; 
            cursor: pointer;
            border-bottom: 2px solid #ddd;
        }}
        th:hover {{ background-color: #e9ecef; }}

        /* --- 列幅設定 --- */
        /* 列の削除に伴い、nth-childの番号が変わるので幅指定も調整しました */

        /* 1. 部屋主 */
        th:nth-child(1), td:nth-child(1) {{ width: 12%; min-width: 90px; }}
        
        /* 2. 順番 */
        th:nth-child(2), td:nth-child(2) {{ width: 6%; min-width: 40px; text-align: center; }}
        
        /* 3. 曲名 */
        th:nth-child(3), td:nth-child(3) {{ width: 25%; min-width: 200px; }}

        /* 4. 作品名 */
        th:nth-child(4), td:nth-child(4) {{ width: 18%; min-width: 150px; }}

        /* 5. 歌手名 */
        th:nth-child(5), td:nth-child(5) {{ width: 18%; min-width: 150px; }}

        /* 6. 歌った人 */
        th:nth-child(6), td:nth-child(6) {{ width: 12%; min-width: 100px; }}

        /* 7. 取得日 (コメントが消えたのでこれが7番目になります) */
        th:nth-child(7), td:nth-child(7) {{ width: 10%; min-width: 90px; }}

        /* 行装飾 */
        tr:nth-child(even) {{ background-color: #fafafa; }}
        tr:hover {{ background-color: #f1f8ff; }}

        /* スマホ向け */
        @media (max-width: 600px) {{
            .header-area {{ padding: 8px; }}
            .controls-row {{ flex-direction: column; align-items: stretch; }}
            .search-container {{ max-width: 100%; }}
            .count-display {{ margin-left: 0; text-align: right; margin-top: 5px; }}
            
            th, td {{ padding: 6px; font-size: 12px; }}
        }}
    </style>
</head>
<body>

    <div class="header-area">
        <div class="title-row">
            <h1>Karaoke setlist all</h1>
            <div class="update-time">最終集計: {current_datetime_str}</div>
        </div>
        
        <div class="controls-row">
            <div class="search-container">
                <input type="text" id="searchInput" class="search-box" placeholder="キーワード・日付 (例: 2026/01/11)...">
                <button onclick="filterTable()" class="btn"><i class="fas fa-search"></i> 検索</button>
                <button onclick="resetFilter()" class="btn btn-reset">リセット</button>
            </div>
            <div class="count-display" id="countDisplay">読み込み中...</div>
        </div>
    </div>
"""

if not html_df.empty:
    initial_count = len(html_df)
    
    html_content += '<div class="table-wrapper"><table id="setlistTable">'
    
    html_content += '<thead><tr>'
    for col in html_df.columns:
        html_content += f'<th onclick="sortTable({list(html_df.columns).index(col)})">{col} <i class="fas fa-sort"></i></th>'
    html_content += '</tr></thead>'
    
    html_content += '<tbody>'
    for _, row in html_df.iterrows():
        html_content += '<tr>'
        for val in row:
            html_content += f'<td>{val}</td>'
        html_content += '</tr>'
    html_content += '</tbody></table></div>'
    
else:
    initial_count = 0
    html_content += '<div style="padding:20px; text-align:center; color:#666;">データがありません。</div>'

# JavaScript
html_content += f"""
<script>
    document.getElementById('countDisplay').innerText = '全 {initial_count} 件';

    function filterTable() {{
        const input = document.getElementById("searchInput");
        const filter = input.value.toUpperCase();
        const keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        
        const table = document.getElementById("setlistTable");
        const trs = table.getElementsByTagName("tr");
        let visibleCount = 0;

        for (let i = 1; i < trs.length; i++) {{
            const tr = trs[i];
            let rowText = "";
            const tds = tr.getElementsByTagName("td");
            for (let j = 0; j < tds.length; j++) {{
                rowText += (tds[j].textContent || tds[j].innerText) + " ";
            }}
            rowText = rowText.toUpperCase();
            
            let isMatch = true;
            for (let k = 0; k < keywords.length; k++) {{
                if (rowText.indexOf(keywords[k]) === -1) {{
                    isMatch = false;
                    break;
                }}
            }}
            if (isMatch || keywords.length === 0) {{
                tr.style.display = "";
                visibleCount++;
            }} else {{
                tr.style.display = "none";
            }}
        }}
        
        document.getElementById('countDisplay').innerText = '表示: ' + visibleCount + ' 件 / 全 ' + (trs.length - 1) + ' 件';
    }}

    document.getElementById("searchInput").addEventListener("keyup", function(event) {{
        if (event.key === "Enter") filterTable();
    }});

    function resetFilter() {{
        document.getElementById("searchInput").value = "";
        filterTable();
    }}

    function sortTable(n) {{
        const table = document.getElementById("setlistTable");
        const tbody = table.querySelector('tbody');
        const rows = Array.from(tbody.rows);
        const th = table.querySelectorAll('th')[n];
        
        let dir = th.getAttribute('data-dir') === 'asc' ? 'desc' : 'asc';
        
        table.querySelectorAll('th').forEach(h => h.setAttribute('data-dir', ''));
        th.setAttribute('data-dir', dir);

        rows.sort((a, b) => {{
            const cellA = a.cells[n].innerText.trim();
            const cellB = b.cells[n].innerText.trim();

            if (!isNaN(cellA) && !isNaN(cellB) && cellA !== '' && cellB !== '') {{
                const numA = parseFloat(cellA);
                const numB = parseFloat(cellB);
                return dir === 'asc' ? numA - numB : numB - numA;
            }}

            return dir === 'asc' 
                ? cellA.localeCompare(cellB, 'ja') 
                : cellB.localeCompare(cellA, 'ja');
        }});

        rows.forEach(row => tbody.appendChild(row));
    }}
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

