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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Karaoke setlist all</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body { font-family: "Helvetica Neue", Arial, sans-serif; padding: 20px; color: #333; }
        h1 { margin-bottom: 10px; }
        
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

        /* スクロールコンテナ */
        .table-wrapper {
            max-height: 80vh;
            overflow-y: auto;
            border: 1px solid #ddd; /* 外枠 */
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }

        /* テーブル設定（隙間対策のため separate に変更） */
        table { 
            border-collapse: separate; 
            border-spacing: 0; 
            width: 100%; 
            font-size: 14px; 
        }
        
        /* セルの枠線設定（separateにしたので個別に線を引く） */
        th, td { 
            padding: 10px; 
            text-align: left; 
            border-right: 1px solid #ddd;
            border-bottom: 1px solid #ddd;
        }
        
        /* ヘッダー固有設定 */
        th { 
            background-color: #f2f2f2; 
            position: sticky; 
            top: 0; 
            z-index: 10; 
            cursor: pointer;
            border-top: none; /* 上の線はwrapperの線に任せる */
            border-bottom: 1px solid #ccc; /* ヘッダーの下線は少し濃く */
            box-shadow: 0 2px 2px -1px rgba(0, 0, 0, 0.1);
            user-select: none;
        }

        /* 右端の線はwrapperの線と重複するので消す */
        th:last-child, td:last-child {
            border-right: none;
        }
        
        /* 最後の行の下線も消す */
        tr:last-child td {
            border-bottom: none;
        }

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

    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
        final_df = final_df.sort_values(by=['順番'], ascending=True)

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
    html_content += "<p>データの取得に失敗しました。</p>"

# JavaScript (高速ソート版)
html_content += """
<script>
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
