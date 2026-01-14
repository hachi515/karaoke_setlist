import pandas as pd
import requests
import datetime
import os
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

print("データを取得中...")
for port in target_ports:
    url = f"http://Ykr.moe:{port}/simplelist.php"
    try:
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0]
            df = df.fillna("") 
            df['部屋主'] = room_map[port]
            df['取得日'] = current_date_str
            new_data_frames.append(df)
            
    except Exception as e:
        pass # エラーは無視

# 新しいデータがある場合のみ処理
if new_data_frames:
    new_df = pd.concat(new_data_frames, ignore_index=True)
    combined_df = pd.concat([history_df, new_df], ignore_index=True)

    # --- クリーニング (ヘッダー行の混入除去) ---
    clean_check_cols = ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']
    for col in clean_check_cols:
        if col in combined_df.columns:
            combined_df = combined_df[combined_df[col] != col]

    # --- 重複排除 (keep='first' で過去の日付を維持) ---
    # これにより、日付が変わってもリストに残り続けている曲は、
    # 新しい日付で上書きされず、最初に取得した日付のまま残ります。
    subset_cols = ['部屋主', '順番', '曲名（ファイル名）', '歌った人']
    existing_cols = [c for c in subset_cols if c in combined_df.columns]
    
    final_df = combined_df.drop_duplicates(subset=existing_cols, keep='first')
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

    final_df.to_csv(history_file, index=False, encoding='utf-8-sig')
    print("履歴ファイルを更新しました。")

else:
    final_df = history_df
    print("新しいデータなし。過去データを使用。")


# ==========================================
# ★修正: クール集計表の処理 (CSV構造に対応)
# ==========================================
analysis_html_content = "" 
cool_data_exists = False
cool_file = "cool_analysis.csv" # 指定されたファイル名

# ファイル存在チェック（なければ似た名前を探す）
if not os.path.exists(cool_file):
    print(f"{cool_file} が見つかりません。代替ファイルを探します...")
    possible_files = [f for f in os.listdir('.') if f.endswith('.csv') and 'history' not in f]
    if possible_files:
        cool_file = possible_files[0]
        print(f"-> {cool_file} を使用します。")
    else:
        cool_file = None

if cool_file:
    try:
        # 文字コード対応読み込み
        raw_df = None
        for enc in ['utf-8-sig', 'cp932', 'shift_jis']:
            try:
                # header=Noneで読み込み、行ごとに解析する
                raw_df = pd.read_csv(cool_file, header=None, encoding=enc)
                print(f"集計表をエンコーディング {enc} で読み込みました。")
                break
            except UnicodeDecodeError:
                continue
        
        if raw_df is not None:
            raw_df = raw_df.fillna("")
            
            # 集計用の履歴データ準備 (2026/1/1 - 2026/3/31)
            start_date = pd.to_datetime("2026/01/01")
            end_date = pd.to_datetime("2026/03/31")
            
            analysis_source_df = final_df.copy()
            analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
            target_history = analysis_source_df[
                (analysis_source_df['dt_obj'] >= start_date) & 
                (analysis_source_df['dt_obj'] <= end_date)
            ]

            # --- CSV解析 & HTML生成ループ ---
            categorized_data = {}
            current_category = "未分類"
            
            # 行ごとの解析
            for idx, row in raw_df.iterrows():
                # 空行スキップ
                if not any(str(x).strip() for x in row):
                    continue

                col0 = str(row[0]).strip()
                
                # カテゴリ判定 (「20xx年」や「アニメ」を含み、かつ「作品名」ではない)
                # ファイル構造に合わせて柔軟に判定
                if ("年" in col0 or "アニメ" in col0) and "作品名" not in col0:
                    current_category = col0
                    if current_category not in categorized_data:
                        categorized_data[current_category] = []
                    continue
                
                # ヘッダー行スキップ
                if "作品名" in col0:
                    continue
                    
                # データ行 (列インデックス: 0=作品名, 1=OP/ED, 2=歌手, 3=曲名)
                # エラー回避のため長さをチェック
                anime = str(row[0]).strip() if len(row) > 0 else ""
                type_ = str(row[1]).strip() if len(row) > 1 else ""
                artist = str(row[2]).strip() if len(row) > 2 else ""
                song = str(row[3]).strip() if len(row) > 3 else ""
                
                # データが空ならスキップ
                if not anime and not song:
                    continue

                if current_category not in categorized_data:
                    categorized_data[current_category] = []
                
                categorized_data[current_category].append({
                    "anime": anime, "type": type_, "artist": artist, "song": song
                })

            # HTML組み立て
            for category, items in categorized_data.items():
                analysis_html_content += f"""
                <div class="category-header">{category}</div>
                <table class="analysisTable">
                    <thead>
                        <tr>
                            <th width="25%">作品名</th>
                            <th width="10%">OP/ED</th>
                            <th width="20%">歌手</th>
                            <th width="25%">曲名</th>
                            <th width="20%">歌唱数</th>
                        </tr>
                    </thead>
                    <tbody>
                """
                
                for item in items:
                    # マッチング処理
                    mask = pd.Series([False] * len(target_history))
                    
                    # 曲名で検索
                    if item["song"]:
                        # 特殊文字のエスケープ処理をして部分一致検索
                        safe_song = re.escape(item["song"])
                        mask = mask | target_history['曲名（ファイル名）'].str.contains(safe_song, case=False, na=False)
                    
                    # 作品名で検索
                    if item["anime"]:
                        safe_anime = re.escape(item["anime"])
                        mask = mask | target_history['曲名（ファイル名）'].str.contains(safe_anime, case=False, na=False)
                    
                    count = len(target_history[mask])
                    
                    # 0回でも表示
                    row_class = "zero-count" if count == 0 else "has-count"
                    
                    # グラフバー
                    bar_width = min(count * 20, 150)
                    bar_html = ""
                    if count > 0:
                        bar_html = f'<div class="bar-chart" style="width:{bar_width}px;"></div>'
                    
                    analysis_html_content += f'<tr class="{row_class}">'
                    analysis_html_content += f'<td>{item["anime"]}</td>'
                    analysis_html_content += f'<td align="center">{item["type"]}</td>'
                    analysis_html_content += f'<td>{item["artist"]}</td>'
                    analysis_html_content += f'<td>{item["song"]}</td>'
                    analysis_html_content += f'<td class="count-cell">'
                    analysis_html_content += f'  <div class="count-wrapper"><span class="count-num">{count}</span>{bar_html}</div>'
                    analysis_html_content += f'</td>'
                    analysis_html_content += '</tr>'
                
                analysis_html_content += "</tbody></table>"

            cool_data_exists = True
            print("クール集計処理が完了しました。")
        else:
            print("CSVファイルの読み込みに失敗しました。")

    except Exception as e:
        print(f"クール集計表処理エラー: {e}")
        import traceback
        traceback.print_exc()


# ==========================================
# HTML生成 (UI調整: 文字サイズ大、高速検索)
# ==========================================

columns_to_hide = ['コメント'] 
if not final_df.empty:
    html_df = final_df.drop(columns=columns_to_hide, errors='ignore')
else:
    html_df = pd.DataFrame()

# セットリスト行生成
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
            display: flex; flex-direction: column;
        }}

        /* Header */
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
        h1 {{ margin: 0; font-size: 1.4rem; color: var(--primary-color); }}
        .update-time {{ font-size: 0.9rem; color: #7f8c8d; }}

        /* Tabs */
        .tabs {{
            display: flex; padding: 0 20px;
            background: #fff; border-bottom: 1px solid var(--border-color);
        }}
        .tab-btn {{
            padding: 12px 25px; cursor: pointer; border: none; background: none;
            font-weight: bold; color: #7f8c8d;
            border-bottom: 3px solid transparent; transition: all 0.3s;
            font-size: 16px;
        }}
        .tab-btn:hover {{ color: var(--accent-color); }}
        .tab-btn.active {{
            color: var(--accent-color); border-bottom-color: var(--accent-color);
        }}

        /* Controls */
        .controls-row {{
            padding: 12px 20px; display: flex; gap: 10px; align-items: center;
            background-color: #fff; border-bottom: 1px solid var(--border-color);
            flex: 0 0 auto;
        }}
        .search-box {{
            padding: 8px 15px; border: 1px solid #ccc; border-radius: 20px;
            width: 300px; outline: none; transition: 0.3s; font-size: 16px;
        }}
        .search-box:focus {{ border-color: var(--accent-color); box-shadow: 0 0 5px rgba(52,152,219,0.3); }}
        .btn {{
            padding: 8px 20px; border-radius: 20px; border: none; cursor: pointer;
            color: #fff; background-color: var(--accent-color); font-size: 1rem;
        }}
        .count-display {{ margin-left: auto; font-weight: bold; font-size: 1rem; }}

        /* Content */
        .content-area {{
            flex: 1; display: flex; flex-direction: column; min-height: 0; position: relative;
        }}
        .tab-content {{
            display: none; flex: 1; overflow: auto;
            -webkit-overflow-scrolling: touch; padding: 15px 20px 50px 20px; box-sizing: border-box;
        }}
        .tab-content.active {{ display: block; }}

        /* Table (Setlist) */
        table {{
            width: 100%; border-collapse: separate; border-spacing: 0;
            background: #fff; border-radius: 8px; overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 20px;
        }}
        th, td {{
            padding: 12px 15px; text-align: left; border-bottom: 1px solid #eee;
            font-size: 16px; /* 文字サイズ大 */
            vertical-align: middle; line-height: 1.5;
        }}
        th {{
            background-color: var(--primary-color); color: #fff;
            position: sticky; top: 0; z-index: 10; font-weight: 600;
        }}
        tr:nth-child(even) {{ background-color: #f8f9fa; }}
        tr:hover {{ background-color: #eaf2f8; }}
        
        /* 高速検索用非表示クラス */
        tr.hidden {{ display: none !important; }}

        /* Analysis Table Styles */
        .category-header {{
            margin-top: 30px; margin-bottom: 10px; padding: 10px 15px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; border-radius: 8px;
            font-weight: bold; font-size: 1.3rem;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        }}
        tr.zero-count {{ color: #ccc; }}
        tr.has-count {{ background-color: #fff; font-weight: 600; color: #333; }}
        tr.has-count:nth-child(even) {{ background-color: #f0f8ff; }}
        
        .count-cell {{ width: 200px; }}
        .count-wrapper {{ display: flex; align-items: center; gap: 10px; }}
        .count-num {{ width: 30px; text-align: right; display:inline-block; font-size:1.2rem; }}
        .bar-chart {{
            height: 12px; background: linear-gradient(90deg, #3498db, #2980b9);
            border-radius: 6px;
        }}
        
        @media (max-width: 600px) {{
            .search-box {{ width: 150px; }}
            .tab-btn {{ padding: 10px 10px; font-size: 0.9rem; }}
            th, td {{ padding: 8px 8px; font-size: 13px; }}
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
            <button class="tab-btn" onclick="openTab('analysis')">クール集計</button>
        </div>
    </div>

    <div id="setlist-controls" class="controls-row">
        <input type="text" id="searchInput" class="search-box" placeholder="キーワード検索...">
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
                <div style="font-size:0.9rem; color:#7f8c8d; margin-bottom:10px; text-align:right;">集計対象: 2026/01/01 - 2026/03/31</div>
            </div>
            {analysis_html_content if cool_data_exists else '<div style="padding:20px;text-align:center;color:#e74c3c;">集計データがありません。<br>cool_analysis.csvを確認してください。</div>'}
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

    // --- 高速検索ロジック (Cache + Debounce) ---
    const searchInput = document.getElementById("searchInput");
    const table = document.getElementById("setlistTable");
    const countDisplay = document.getElementById('countDisplay');
    
    let tableData = [];
    let tbodyRows = [];
    let debounceTimer;

    window.addEventListener('DOMContentLoaded', () => {{
        const tbody = table.tBodies[0];
        if (tbody) {{
            tbodyRows = Array.from(tbody.rows);
            // 事前にテキストをキャッシュ (大文字変換済み)
            tableData = tbodyRows.map(row => row.innerText.toUpperCase());
            countDisplay.innerText = '全 ' + tbodyRows.length + ' 件';
        }}
    }});

    searchInput.addEventListener("keyup", function() {{
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(performSearch, 200); // 200ms遅延
    }});

    function performSearch() {{
        const filter = searchInput.value.toUpperCase();
        const keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        let visibleCount = 0;
        
        // DOMアクセスを最小限にして高速化
        const total = tableData.length;
        for (let i = 0; i < total; i++) {{
            let isMatch = true;
            const rowText = tableData[i];
            for (let k = 0; k < keywords.length; k++) {{
                if (rowText.indexOf(keywords[k]) === -1) {{
                    isMatch = false; break;
                }}
            }}
            
            if (isMatch || keywords.length === 0) {{
                tbodyRows[i].classList.remove('hidden');
                visibleCount++;
            }} else {{
                tbodyRows[i].classList.add('hidden');
            }}
        }}
        countDisplay.innerText = '表示: ' + visibleCount + ' / ' + total;
    }}

    function resetFilter() {{
        searchInput.value = "";
        performSearch();
    }}

    function sortTable(n) {{
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
        
        // ソート後にキャッシュ再構築
        tbodyRows = rows;
        tableData = tbodyRows.map(row => row.innerText.toUpperCase());
    }}
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
    print("HTML生成完了: index.html")
