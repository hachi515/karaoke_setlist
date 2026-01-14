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
        history_df = history_df.fillna("")
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
            df['部屋主'] = room_map[port]
            df['取得日'] = current_date_str
            new_data_frames.append(df)
            print(f"Port {port} ({room_map[port]}): OK")
            
    except Exception as e:
        print(f"Port {port}: Error - {e}")

# 新しいデータがある場合のみ処理
if new_data_frames:
    new_df = pd.concat(new_data_frames, ignore_index=True)
    combined_df = pd.concat([history_df, new_df], ignore_index=True)
    
    # 重複排除
    subset_cols = ['取得日', '部屋主', '曲名（ファイル名）', '歌った人', '作品名', '歌手名']
    existing_cols = [c for c in subset_cols if c in combined_df.columns]
    final_df = combined_df.drop_duplicates(subset=existing_cols, keep='last')
    final_df = final_df.fillna("")

    # ソート処理
    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
        
    final_df['temp_date'] = pd.to_datetime(final_df['取得日'], errors='coerce')
    final_df = final_df.sort_values(by=['temp_date', '順番'], ascending=[False, False])
    final_df = final_df.drop(columns=['temp_date'])
    
    # 列整理
    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    # 保存
    final_df.to_csv(history_file, index=False, encoding='utf-8-sig')
    print("履歴ファイルを更新しました。")

else:
    final_df = history_df
    print("新しいデータなし。過去データを使用。")


# ==========================================
# ★修正: クール集計表(CSV)の読み込み処理
# ==========================================
analysis_html_rows = ""
# ★ GitHub運用に合わせてCSVファイル名を指定 (日本語ファイル名でもOKですが、英数字推奨)
cool_file = "クール集計表.csv" 
cool_data_exists = False

# ファイルが存在しない場合、日本語ファイル名も試行するフェイルセーフ
if not os.path.exists(cool_file) and os.path.exists("クール集計表.csv"):
    cool_file = "クール集計表.csv"

if os.path.exists(cool_file):
    try:
        # ★ CSVとして読み込み
        # header=1: 2行目(index 1)をヘッダーとして読み込む設定（1行目は「2026年冬アニメ」などのタイトルのため）
        cool_df = pd.read_csv(cool_file, header=1, encoding='utf-8-sig')
        cool_df = cool_df.fillna("")
        
        # 集計対象期間の設定 (2026/1/1 - 2026/3/31)
        start_date = pd.to_datetime("2026/01/01")
        end_date = pd.to_datetime("2026/03/31")
        
        # Historyを集計期間でフィルタリング
        analysis_source_df = final_df.copy()
        analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
        
        target_history = analysis_source_df[
            (analysis_source_df['dt_obj'] >= start_date) & 
            (analysis_source_df['dt_obj'] <= end_date)
        ]
        
        # マッチング処理
        counts = []
        singers_list = [] 
        
        for idx, row in cool_df.iterrows():
            target_song = str(row.get('曲名', '')).strip()
            target_anime = str(row.get('作品名', '')).strip()
            
            mask = pd.Series([False] * len(target_history))
            
            if target_song:
                mask = mask | target_history['曲名（ファイル名）'].str.contains(pd.escape(target_song), case=False, na=False)
            
            if target_anime:
                mask = mask | target_history['曲名（ファイル名）'].str.contains(pd.escape(target_anime), case=False, na=False)
            
            hit_rows = target_history[mask]
            count = len(hit_rows)
            
            singers = sorted(list(set(hit_rows['歌った人'].astype(str).tolist())))
            singers_str = ", ".join(singers)
            
            counts.append(count)
            singers_list.append(singers_str)
            
        cool_df['歌唱数'] = counts
        cool_df['歌った人'] = singers_list
        
        # HTML行の生成
        for idx, row in cool_df.iterrows():
            count = row['歌唱数']
            row_class = "zero-count" if count == 0 else "has-count"
            
            analysis_html_rows += f'<tr class="{row_class}">'
            analysis_html_rows += f'<td>{row.get("作品名","")}</td>'
            analysis_html_rows += f'<td>{row.get("OP/ED","")}</td>'
            analysis_html_rows += f'<td>{row.get("歌手","")}</td>'
            analysis_html_rows += f'<td>{row.get("曲名","")}</td>'
            analysis_html_rows += f'<td class="count-cell">{count}</td>'
            analysis_html_rows += f'<td class="singer-cell">{row.get("歌った人","")}</td>'
            analysis_html_rows += '</tr>'
            
        cool_data_exists = True
        print(f"クール集計表({cool_file})の処理が完了しました。")

    except Exception as e:
        print(f"クール集計表の処理エラー: {e}")
else:
    print(f"クール集計表ファイル(cool_analysis.csv または クール集計表.csv) が見つかりません。")


# ==========================================
# HTML生成 (タブ切り替え対応)
# ==========================================

# セットリスト用データ
columns_to_hide = ['コメント'] 
if not final_df.empty:
    html_df = final_df.drop(columns=columns_to_hide, errors='ignore')
else:
    html_df = pd.DataFrame()

# セットリストHTML行生成
setlist_rows = ""
for _, row in html_df.iterrows():
    setlist_rows += '<tr>'
    for val in row:
        setlist_rows += f'<td>{val}</td>'
    setlist_rows += '</tr>'

setlist_headers = ""
for col in html_df.columns:
    setlist_headers += f'<th onclick="sortTable({list(html_df.columns).index(col)})">{col} <i class="fas fa-sort"></i></th>'


html_content = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Karaoke Dashboard</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        /* --- Base & Reset --- */
        :root {{
            --primary-color: #2c3e50;
            --accent-color: #3498db;
            --bg-color: #f4f7f6;
            --text-color: #333;
            --header-bg: #fff;
            --border-color: #e0e0e0;
        }}
        html, body {{
            height: 100%; margin: 0; padding: 0;
            overflow: hidden;
            font-family: "Helvetica Neue", Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
        }}

        /* --- Header Area --- */
        .header-area {{
            flex: 0 0 auto;
            background-color: var(--header-bg);
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            z-index: 100;
        }}
        .header-inner {{
            padding: 10px 20px;
            display: flex; justify-content: space-between; align-items: center;
        }}
        h1 {{ margin: 0; font-size: 1.2rem; color: var(--primary-color); }}
        .update-time {{ font-size: 0.8rem; color: #7f8c8d; }}

        /* --- Tabs Navigation --- */
        .tabs {{
            display: flex;
            padding: 0 20px;
            background: #fff;
            border-bottom: 1px solid var(--border-color);
        }}
        .tab-btn {{
            padding: 12px 24px;
            cursor: pointer;
            border: none;
            background: none;
            font-weight: bold;
            color: #7f8c8d;
            border-bottom: 3px solid transparent;
            transition: all 0.3s;
        }}
        .tab-btn:hover {{ color: var(--accent-color); }}
        .tab-btn.active {{
            color: var(--accent-color);
            border-bottom-color: var(--accent-color);
        }}

        /* --- Controls (Search) --- */
        .controls-row {{
            padding: 10px 20px;
            display: flex; gap: 10px; align-items: center;
            background-color: #fff;
            border-bottom: 1px solid var(--border-color);
        }}
        .search-box {{
            padding: 8px 12px; border: 1px solid #ccc; border-radius: 20px;
            width: 250px; outline: none; transition: 0.3s;
        }}
        .search-box:focus {{ border-color: var(--accent-color); box-shadow: 0 0 5px rgba(52,152,219,0.3); }}
        .btn {{
            padding: 6px 15px; border-radius: 20px; border: none; cursor: pointer;
            color: #fff; background-color: var(--accent-color); font-size: 0.9rem;
        }}
        .btn:hover {{ opacity: 0.9; }}
        .count-display {{ margin-left: auto; font-weight: bold; font-size: 0.9rem; }}

        /* --- Main Content Area --- */
        .content-area {{
            flex: 1; overflow: hidden; position: relative;
        }}
        
        .tab-content {{
            display: none;
            height: 100%;
            overflow: auto;
            -webkit-overflow-scrolling: touch;
            padding: 20px;
            box-sizing: border-box;
        }}
        .tab-content.active {{ display: block; }}

        /* --- Table Styling --- */
        table {{
            width: 100%; border-collapse: separate; border-spacing: 0;
            background: #fff; border-radius: 8px; overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            margin-bottom: 40px;
        }}
        th, td {{
            padding: 12px 15px; text-align: left;
            border-bottom: 1px solid #eee;
        }}
        th {{
            background-color: var(--primary-color);
            color: #fff;
            position: sticky; top: 0; z-index: 10;
            font-weight: 500;
        }}
        tr:last-child td {{ border-bottom: none; }}
        tr:nth-child(even) {{ background-color: #f8f9fa; }}
        tr:hover {{ background-color: #eaf2f8; }}

        /* --- Specific Styles for Analysis Table --- */
        .analysis-header {{
            margin-bottom: 15px; padding: 10px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; border-radius: 8px;
            display: flex; justify-content: space-between; align-items: center;
        }}
        .analysis-title {{ font-weight: bold; font-size: 1.1rem; }}
        .period-badge {{
            background: rgba(255,255,255,0.2); padding: 4px 10px;
            border-radius: 4px; font-size: 0.85rem;
        }}
        
        tr.zero-count {{ color: #aaa; }}
        tr.has-count {{ background-color: #e3f2fd; font-weight: 600; color: #1565c0; }}
        tr.has-count:nth-child(even) {{ background-color: #bbdefb; }}
        
        .count-cell {{ font-size: 1.2rem; text-align: center; color: #d35400; }}
        .singer-cell {{ font-size: 0.9rem; color: #555; }}

        @media (max-width: 600px) {{
            .search-box {{ width: 150px; }}
            .tabs {{ padding: 0 5px; }}
            .tab-btn {{ padding: 10px 15px; font-size: 0.9rem; }}
            th, td {{ padding: 8px 10px; font-size: 12px; }}
        }}
    </style>
</head>
<body>

    <div class="header-area">
        <div class="header-inner">
            <h1>Karaoke Dashboard</h1>
            <div class="update-time">{current_datetime_str} 更新</div>
        </div>
        <div class="tabs">
            <button class="tab-btn active" onclick="openTab('setlist')">セットリスト</button>
            <button class="tab-btn" onclick="openTab('analysis')">クール集計 (2025秋-2026冬)</button>
        </div>
    </div>

    <div id="setlist-controls" class="controls-row">
        <input type="text" id="searchInput" class="search-box" placeholder="キーワード検索...">
        <button onclick="filterTable()" class="btn"><i class="fas fa-search"></i></button>
        <button onclick="resetFilter()" class="btn" style="background:#95a5a6"><i class="fas fa-undo"></i></button>
        <div class="count-display" id="countDisplay">読み込み中...</div>
    </div>

    <div class="content-area">
        <div id="setlist" class="tab-content active">
            <table id="setlistTable">
                <thead><tr>{setlist_headers}</tr></thead>
                <tbody>{setlist_rows}</tbody>
            </table>
            {"" if setlist_rows else '<div style="padding:20px;text-align:center">データがありません</div>'}
        </div>

        <div id="analysis" class="tab-content">
            <div class="analysis-header">
                <div class="analysis-title"><i class="fas fa-chart-pie"></i> アニメソング集計状況</div>
                <div class="period-badge">対象期間: 2026/01/01 - 2026/03/31</div>
            </div>
            
            <table id="analysisTable">
                <thead>
                    <tr>
                        <th width="20%">作品名</th>
                        <th width="10%">OP/ED</th>
                        <th width="15%">歌手</th>
                        <th width="25%">曲名</th>
                        <th width="10%">歌唱数</th>
                        <th width="20%">歌った人</th>
                    </tr>
                </thead>
                <tbody>
                    {analysis_html_rows}
                </tbody>
            </table>
            {"" if cool_data_exists else '<div style="padding:20px;text-align:center">集計データが見つかりません。cool_analysis.csv (または クール集計表.csv) を配置してください。</div>'}
        </div>
    </div>

<script>
    function openTab(tabName) {{
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        
        document.getElementById(tabName).classList.add('active');
        
        const btnIndex = tabName === 'setlist' ? 0 : 1;
        document.querySelectorAll('.tab-btn')[btnIndex].classList.add('active');
        
        document.getElementById('setlist-controls').style.display = tabName === 'setlist' ? 'flex' : 'none';
    }}

    const initialCount = document.getElementById("setlistTable").tBodies[0].rows.length;
    document.getElementById('countDisplay').innerText = '全 ' + initialCount + ' 件';

    function filterTable() {{
        const input = document.getElementById("searchInput");
        const filter = input.value.toUpperCase();
        const keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        
        const table = document.getElementById("setlistTable");
        const trs = table.tBodies[0].rows;
        let visibleCount = 0;

        for (let i = 0; i < trs.length; i++) {{
            const tr = trs[i];
            let rowText = "";
            for (let j = 0; j < tr.cells.length; j++) {{
                rowText += tr.cells[j].innerText + " ";
            }}
            rowText = rowText.toUpperCase();
            
            let isMatch = true;
            for (let k = 0; k < keywords.length; k++) {{
                if (rowText.indexOf(keywords[k]) === -1) {{
                    isMatch = false; break;
                }}
            }}
            if (isMatch || keywords.length === 0) {{
                tr.style.display = ""; visibleCount++;
            }} else {{
                tr.style.display = "none";
            }}
        }}
        document.getElementById('countDisplay').innerText = '表示: ' + visibleCount + ' / ' + trs.length;
    }}

    document.getElementById("searchInput").addEventListener("keyup", function(e) {{
        if (e.key === "Enter") filterTable();
    }});
    function resetFilter() {{
        document.getElementById("searchInput").value = "";
        filterTable();
    }}

    function sortTable(n) {{
        const table = document.getElementById("setlistTable");
        const tbody = table.tBodies[0];
        const rows = Array.from(tbody.rows);
        const th = table.querySelectorAll('th')[n];
        let dir = th.getAttribute('data-dir') === 'asc' ? 'desc' : 'asc';
        
        table.querySelectorAll('th').forEach(h => h.setAttribute('data-dir', ''));
        th.setAttribute('data-dir', dir);

        rows.sort((a, b) => {{
            const valA = a.cells[n].innerText.trim();
            const valB = b.cells[n].innerText.trim();
            
            if (!isNaN(valA) && !isNaN(valB) && valA!=='' && valB!=='') {{
                return dir === 'asc' ? valA - valB : valB - valA;
            }}
            return dir === 'asc' ? valA.localeCompare(valB,'ja') : valB.localeCompare(valA,'ja');
        }});
        rows.forEach(row => tbody.appendChild(row));
    }}
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
    print("HTML生成完了: index.html")
