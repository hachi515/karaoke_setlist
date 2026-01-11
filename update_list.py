import pandas as pd
import requests
import datetime
import re

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
    11060: "れん部屋",
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
            
            all_data_frames.append(df)
            print(f"Port {port} ({room_map[port]}): OK")
            
    except Exception as e:
        print(f"Port {port}: Error - {e}")

# HTML生成
html_content = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Karaoke setlist all</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        /* レイアウトの基礎設定 */
        html, body {
            height: 100%;
            margin: 0;
            padding: 0;
            overflow: hidden; /* 画面全体のスクロールを禁止 */
            font-family: -apple-system, BlinkMacSystemFont, "Helvetica Neue", "Hiragino Sans", "Hiragino Kaku Gothic ProN", Arial, sans-serif;
            background-color: #fcfcfc;
            color: #333;
            display: flex;
            flex-direction: column; /* 縦並びのレイアウト */
        }

        /* ヘッダーエリア（固定部分） */
        header {
            flex-shrink: 0; /* 縮まないようにする */
            padding: 15px 20px 0 20px;
            background-color: #fcfcfc;
        }

        h1 { 
            margin: 0 0 5px 0; 
            font-size: 1.5rem; 
            text-align: left; /* 左寄せ */
        }
        
        .update-time { 
            color: #666; 
            font-size: 0.85em; 
            text-align: left; /* 左寄せ */
            margin-bottom: 15px; 
        }

        /* 検索エリアのデザイン */
        .search-container {
            display: flex;
            gap: 10px;
            align-items: center;
            margin-bottom: 10px;
            flex-wrap: wrap;
        }

        .search-box {
            width: 300px; /* PCでの基本幅 */
            padding: 6px 10px;
            font-size: 14px;
            border: 1px solid #ccc;
            border-radius: 4px;
            height: 36px; /* 高さを固定してスマートに */
            box-sizing: border-box;
        }

        .btn-group {
            display: flex;
            gap: 5px;
        }

        .btn {
            padding: 0 15px;
            height: 36px; /* 検索ボックスと同じ高さ */
            line-height: 36px; /* 文字を垂直中央に */
            font-size: 14px;
            cursor: pointer;
            background-color: #007bff; 
            color: white; 
            border: none; 
            border-radius: 4px;
            white-space: nowrap;
        }
        .btn:hover { background-color: #0056b3; }
        .btn-reset { background-color: #6c757d; }
        .btn-reset:hover { background-color: #545b62; }

        .count-display {
            text-align: right;
            font-size: 0.85em;
            color: #555;
            margin-bottom: 5px;
            padding-right: 20px;
        }

        /* テーブルエリア（スクロール部分） */
        .table-wrapper {
            flex-grow: 1; /* 余ったスペースを全部使う */
            overflow: auto; /* ここだけスクロールさせる */
            margin: 0 20px 20px 20px; /* 余白 */
            border: 1px solid #ddd;
            background-color: #fff;
            position: relative;
            -webkit-overflow-scrolling: touch;
        }

        table { 
            border-collapse: separate; 
            border-spacing: 0; 
            width: 100%; 
            font-size: 13px; 
            min-width: 800px; /* スマホで横スクロールさせるための最小幅 */
        }
        
        th, td { 
            padding: 8px 10px; 
            text-align: left; 
            border-right: 1px solid #eee;
            border-bottom: 1px solid #eee;
            vertical-align: middle;
        }

        /* ヘッダー固定 */
        th { 
            background-color: #f1f3f5; 
            position: sticky; 
            top: 0; 
            z-index: 10; 
            border-bottom: 1px solid #ccc;
            white-space: nowrap;
        }

        /* 列幅の調整 */
        th:nth-child(1), td:nth-child(1) { min-width: 90px; } /* 部屋主 */
        th:nth-child(2), td:nth-child(2) { min-width: 40px; text-align: center; } /* 順番 */
        th:nth-child(3), td:nth-child(3) { min-width: 200px; } /* 曲名 */

        th:last-child, td:last-child { border-right: none; }
        tr:last-child td { border-bottom: none; }
        tr:nth-child(even) { background-color: #f9f9f9; }

        /* スマホ用レスポンシブ設定 */
        @media (max-width: 600px) {
            header { padding: 10px 10px 0 10px; }
            
            h1 { font-size: 1.2rem; }
            
            .search-container {
                flex-direction: column; /* 縦並び */
                align-items: stretch; /* 横幅いっぱい */
                gap: 8px;
            }
            .search-box { width: 100%; }
            .btn-group { width: 100%; }
            .btn { flex: 1; text-align: center; } /* ボタンを均等幅に */
            
            .table-wrapper { margin: 0 10px 10px 10px; }
            .count-display { padding-right: 10px; }
        }
    </style>
</head>
<body>
    <header>
        <h1>Karaoke setlist all</h1>
"""

html_content += f'<div class="update-time">最終集計: {current_datetime_str}</div>'

html_content += f'''
        <div class="search-container">
            <input type="text" id="searchInput" class="search-box" placeholder="キーワード・日付で検索...">
            <div class="btn-group">
                <button onclick="filterTable()" class="btn"><i class="fas fa-search"></i> 検索</button>
                <button onclick="resetFilter()" class="btn btn-reset">リセット</button>
            </div>
        </div>
        <div class="count-display"></div> </header>
'''

if all_data_frames:
    final_df = pd.concat(all_data_frames, ignore_index=True)

    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
    
    # ソート: 取得日(降順) -> 順番(降順)
    final_df = final_df.sort_values(by=['取得日', '順番'], ascending=[False, False])

    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    html_content += '<div class="table-wrapper"><table id="setlistTable">'
    
    html_content += '<thead><tr>'
    for col in final_df.columns:
        html_content += f'<th onclick="sortTable({list(final_df.columns).index(col)})">{col} <i class="fas fa-sort"></i></th>'
    html_content += '</tr></thead>'
    
    html_content += '<tbody>'
    for _, row in final_df.iterrows():
        html_content += '<tr>'
        for val in row:
            html_content += f'<td>{val}</td>'
        html_content += '</tr>'
    html_content += '</tbody></table></div>'

else:
    html_content += "<p style='padding:20px;'>データの取得に失敗しました。</p>"

# JavaScript
html_content += """
<script>
    // 初期表示時に件数をカウント
    document.addEventListener('DOMContentLoaded', function() {
        updateCount();
    });

    function updateCount() {
        const table = document.getElementById("setlistTable");
        const trs = table.getElementsByTagName("tr");
        let visibleCount = 0;
        // ヘッダー行(0)を除く
        for (let i = 1; i < trs.length; i++) {
            if (trs[i].style.display !== "none") {
                visibleCount++;
            }
        }
        const total = trs.length - 1;
        const display = document.querySelector('.count-display');
        if(display) {
            display.innerText = '表示: ' + visibleCount + ' 件 / 全 ' + total + ' 件';
        }
    }

    function filterTable() {
        const input = document.getElementById("searchInput");
        const filter = input.value.toUpperCase();
        const keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        
        const table = document.getElementById("setlistTable");
        const trs = table.getElementsByTagName("tr");

        for (let i = 1; i < trs.length; i++) {
            const tr = trs[i];
            let rowText = "";
            const tds = tr.getElementsByTagName("td");
            for (let j = 0; j < tds.length; j++) {
                rowText += (tds[j].textContent || tds[j].innerText) + " ";
            }
            rowText = rowText.toUpperCase();
            
            let isMatch = true;
            for (let k = 0; k < keywords.length; k++) {
                if (rowText.indexOf(keywords[k]) === -1) {
                    isMatch = false;
                    break;
                }
            }
            tr.style.display = (isMatch || keywords.length === 0) ? "" : "none";
        }
        updateCount();
    }

    document.getElementById("searchInput").addEventListener("keyup", function(event) {
        if (event.key === "Enter") filterTable();
    });

    function resetFilter() {
        document.getElementById("searchInput").value = "";
        filterTable();
    }

    function sortTable(n) {
        const table = document.getElementById("setlistTable");
        const tbody = table.querySelector('tbody');
        const rows = Array.from(tbody.rows);
        const th = table.querySelectorAll('th')[n];
        
        let dir = th.getAttribute('data-dir') === 'asc' ? 'desc' : 'asc';
        
        table.querySelectorAll('th').forEach(h => h.setAttribute('data-dir', ''));
        th.setAttribute('data-dir', dir);

        rows.sort((a, b) => {
            const cellA = a.cells[n].innerText.trim();
            const cellB = b.cells[n].innerText.trim();

            if (!isNaN(cellA) && !isNaN(cellB) && cellA !== '' && cellB !== '') {
                const numA = parseFloat(cellA);
                const numB = parseFloat(cellB);
                return dir === 'asc' ? numA - numB : numB - numA;
            }

            return dir === 'asc' 
                ? cellA.localeCompare(cellB, 'ja') 
                : cellB.localeCompare(cellA, 'ja');
        });

        rows.forEach(row => tbody.appendChild(row));
    }
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

