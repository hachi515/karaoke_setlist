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
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Karaoke setlist all</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        /* 全体レイアウト: 画面の高さ100%を使い切る設定 */
        html, body {
            height: 100%;
            margin: 0;
            padding: 0;
            overflow: hidden; /* body自体のスクロールを禁止 */
            font-family: -apple-system, BlinkMacSystemFont, "Helvetica Neue", "Hiragino Sans", "Hiragino Kaku Gothic ProN", Arial, sans-serif;
            background-color: #fcfcfc;
            color: #333;
            display: flex;
            flex-direction: column; /* 縦並びのフレックスボックス */
        }

        /* ヘッダーエリア（タイトル・検索） */
        .header-area {
            flex: 0 0 auto; /* 高さは中身に合わせて固定 */
            padding: 15px 20px 10px 20px;
            background-color: #fff;
            border-bottom: 1px solid #ddd;
            box-shadow: 0 2px 4px rgba(0,0,0,0.03);
            z-index: 20;
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

        /* 検索ボックスエリア: シンプル・左寄せ・コンパクト */
        .search-container {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: center;
            justify-content: flex-start; /* 左寄せ */
        }

        .search-box {
            width: 300px; /* PCでは適度な幅に固定 */
            padding: 8px 12px;
            font-size: 14px;
            border: 1px solid #ccc;
            border-radius: 4px;
            box-sizing: border-box;
            background-color: #f9f9f9;
        }
        .search-box:focus {
            background-color: #fff;
            outline: 2px solid #007bff;
        }

        .btn-group {
            display: flex;
            gap: 5px;
        }

        .btn {
            padding: 8px 16px;
            font-size: 14px;
            cursor: pointer;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            white-space: nowrap;
            transition: background-color 0.2s;
        }
        .btn:hover { background-color: #0056b3; }
        .btn-reset { background-color: #6c757d; }
        .btn-reset:hover { background-color: #545b62; }

        /* 件数表示 */
        .count-display {
            text-align: left; /* 左寄せ */
            font-size: 0.85em;
            color: #666;
            margin-top: 5px;
            font-weight: bold;
        }

        /* テーブルラッパー: 残りの高さを全て使い、この中だけでスクロールさせる */
        .table-wrapper {
            flex: 1 1 auto; /* 残りのスペースを埋める */
            overflow: auto; /* 縦横スクロール */
            position: relative;
            background-color: #fff;
            -webkit-overflow-scrolling: touch; /* iOS慣性スクロール */
        }

        /* テーブル設定 */
        table { 
            border-collapse: separate; 
            border-spacing: 0; 
            width: 100%; 
            font-size: 13px; 
            min-width: 800px; /* スマホでも横につぶれないよう最小幅を確保 */
        }
        
        th, td { 
            padding: 10px 12px;
            text-align: left; 
            border-right: 1px solid #eee;
            border-bottom: 1px solid #eee;
            vertical-align: middle;
            line-height: 1.5;
        }

        /* ヘッダー固定設定 */
        th { 
            background-color: #f1f3f5;
            color: #444;
            font-weight: bold;
            position: sticky; 
            top: 0; 
            z-index: 10; 
            cursor: pointer;
            border-bottom: 2px solid #ddd;
            white-space: nowrap;
        }
        th:hover { background-color: #e9ecef; }

        /* 列幅の調整 */
        th:nth-child(1), td:nth-child(1) { min-width: 90px; } /* 部屋主 */
        th:nth-child(2), td:nth-child(2) { min-width: 50px; text-align: center; } /* 順番 */
        
        td { word-break: break-all; } /* 長い文字を折り返す */
        th:last-child, td:last-child { border-right: none; }
        tr:nth-child(even) { background-color: #fafafa; }
        
        /* --- スマホ向けレスポンシブ調整 --- */
        @media (max-width: 600px) {
            /* ヘッダー周りの余白を詰める */
            .header-area { padding: 10px 12px; }
            h1 { font-size: 1.3rem; }
            .update-time { margin-bottom: 10px; }
            
            /* 検索ボックスをスマホ幅いっぱいにする */
            .search-container { 
                flex-direction: column; /* 縦並び */
                align-items: stretch; /* 幅いっぱい */
                gap: 8px;
            }
            .search-box { width: 100%; font-size: 16px; /* iOS拡大防止 */ }
            
            .btn-group { 
                display: flex; 
                gap: 8px; 
            }
            .btn { flex: 1; text-align: center; padding: 10px; } /* ボタン押しやすく */
            
            /* テーブル文字サイズ微調整 */
            th, td { padding: 8px; font-size: 12px; }
        }
    </style>
</head>
<body>

    <div class="header-area">
        <h1>Karaoke setlist all</h1>
        <div class="update-time">最終集計: {current_datetime_str}</div>
        
        <div class="search-container">
            <input type="text" id="searchInput" class="search-box" placeholder="キーワード・日付 (例: 2026/01/11)...">
            <div class="btn-group">
                <button onclick="filterTable()" class="btn"><i class="fas fa-search"></i> 検索</button>
                <button onclick="resetFilter()" class="btn btn-reset">リセット</button>
            </div>
        </div>
        """

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

    # 件数表示
    html_content += f'<div class="count-display">全 {len(final_df)} 件</div>'
    html_content += '</div>' # header-area 終了

    # スクロールエリア
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
    html_content += '<p style="padding:20px;">データの取得に失敗しました。</p></div>'

# JavaScript
html_content += """
<script>
    function filterTable() {
        const input = document.getElementById("searchInput");
        const filter = input.value.toUpperCase();
        const keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        
        const table = document.getElementById("setlistTable");
        const trs = table.getElementsByTagName("tr");
        let visibleCount = 0;

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
            if (isMatch || keywords.length === 0) {
                tr.style.display = "";
                visibleCount++;
            } else {
                tr.style.display = "none";
            }
        }
        
        const countDisplay = document.querySelector('.count-display');
        if(countDisplay) countDisplay.innerText = '表示: ' + visibleCount + ' 件 / 全 ' + (trs.length - 1) + ' 件';
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
