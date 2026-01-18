import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
import glob
from itertools import groupby

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
    11010: "加古部屋",
    11011: "加古部屋",
    11012: "加古部屋",
    11013: "加古部屋",
    11014: "加古部屋",
    11015: "加古部屋",
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
    11106: "冨塚部屋",
    11107: "ブルーベリー部屋"
}

# --- ファイル名設定 ---
HISTORY_FILE = "history.csv"

# --- 関数定義 ---
def normalize_text(text):
    if not isinstance(text, str): return str(text)
    text = unicodedata.normalize('NFKC', text)
    text = re.sub(r'\.[a-zA-Z0-9]{3,4}$', '', text)
    text = re.sub(r'[\[\(\{【].*?[\]\)\}】]', ' ', text)
    text = re.sub(r'(key|KEY)?\s*[\+\-]\s*[0-9]+', ' ', text)
    text = re.sub(r'原キー', ' ', text)
    text = re.sub(r'(キー)?変更[:：]?', ' ', text)
    text = re.sub(r'[~〜～\-_=,.]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text.upper()

def normalize_offline_text(text):
    if not isinstance(text, str): return str(text)
    text = unicodedata.normalize('NFKC', text)
    text = re.sub(r'\.[a-zA-Z0-9]{3,4}$', '', text)
    text = re.sub(r'(key|KEY)?\s*[\+\-]\s*[0-9]+', ' ', text)
    text = re.sub(r'原キー', ' ', text)
    text = re.sub(r'(キー)?変更[:：]?', ' ', text)
    text = re.sub(r'[~〜～\-_=,.]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text.upper()

# ★追加関数: キー変更情報のみを削除する (作品名の抽出・移動はしない)
def remove_key_info(text):
    if not isinstance(text, str): return str(text)
    song = text
    # (Key+2), [原キー], Key: -1, キー変更-2 などのパターンを削除
    key_patterns = [
        r'[\(（\[【]\s*(?:key|KEY|キー)?\s*(?:[\+\-±]\d+|原キー|変更なし)\s*[\)）\]】]',
        r'(?:key|KEY|キー)[:：]?\s*[\+\-±]\d+',
        r'キー変更[:：]?\s*[\+\-±]?\d+'
    ]
    for pat in key_patterns:
        song = re.sub(pat, '', song, flags=re.IGNORECASE)
    
    return song.strip()

# --- 1. 既存ファイル読み込み ---
if os.path.exists(HISTORY_FILE):
    try:
        print(f"既存の {HISTORY_FILE} を読み込んでいます...")
        history_df = pd.read_csv(HISTORY_FILE, encoding='utf-8-sig')
        # カラム名にゴミが入っている場合のクリーニング
        clean_check_cols = ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']
        for col in clean_check_cols:
            if col in history_df.columns:
                history_df = history_df[history_df[col] != col]
    except Exception as e:
        print(f"既存ファイル読み込み警告: {e}")
        history_df = pd.DataFrame()
else:
    print("既存の履歴ファイルが見つかりません。新規作成します。")
    history_df = pd.DataFrame()

# --- 2. 新しいデータ取得 ---
target_ports = list(room_map.keys())
new_data_frames = []

print("最新データを取得中...")

for port in target_ports:
    url = f"http://Ykr.moe:{port}/simplelist.php"
    try:
        response = requests.get(url, timeout=20)
        response.raise_for_status()
        response.encoding = response.apparent_encoding
        
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0]
            if not df.empty:
                # カラム名をリセットして番号順でアクセス
                df.columns = range(df.shape[1])
                
                temp_df = pd.DataFrame()
                
                if df.shape[1] >= 1: temp_df['順番'] = df[0]
                
                # ★修正: 曲名からキー情報のみ削除
                if df.shape[1] >= 2:
                    temp_df['曲名（ファイル名）'] = df[1].apply(remove_key_info)
                else:
                    temp_df['曲名（ファイル名）'] = ""

                # 作品名カラムを用意 (抽出や移動はせず、空欄で初期化)
                temp_df['作品名'] = ""

                if df.shape[1] >= 3: temp_df['歌手名'] = df[2]
                if df.shape[1] >= 4: temp_df['歌った人'] = df[3]
                if df.shape[1] >= 5: temp_df['キー'] = df[4]
                if df.shape[1] >= 6: temp_df['コメント'] = df[5]
                
                if 'コメント' not in temp_df.columns:
                    temp_df['コメント'] = ""

                temp_df = temp_df.fillna("") 
                temp_df['部屋主'] = room_map.get(port, f"Port {port}")
                temp_df['取得日'] = current_date_str
                
                new_data_frames.append(temp_df)
    except Exception as e:
        pass

# --- 3. 結合・重複排除・保存 ---
if new_data_frames:
    print("データを結合して整理中...")
    new_df = pd.concat(new_data_frames, ignore_index=True)
    
    # 既存データと新データを結合
    combined_df = pd.concat([history_df, new_df], ignore_index=True)
    
    # ★修正: 重複判定のカラム設定（作品名、歌手名を追加して厳密に判定）
    subset_cols = ['部屋主', '順番', '曲名（ファイル名）', '作品名', '歌手名', '歌った人']
    
    # 実際に存在するカラムだけでチェック
    existing_cols = [c for c in subset_cols if c in combined_df.columns]
    
    # 重複排除 (古いデータを残す設定 keep='first')
    final_df = combined_df.drop_duplicates(subset=existing_cols, keep='first')
    
    # 欠損値埋め
    final_df = final_df.fillna("")

    # 順番を整数化
    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce').fillna(0).astype(int)

    # ソート処理
    if '取得日' in final_df.columns:
        final_df['temp_date'] = pd.to_datetime(final_df['取得日'], errors='coerce')
        final_df = final_df.sort_values(by=['temp_date', '順番'], ascending=[False, False])
        final_df = final_df.drop(columns=['temp_date'])

    # 列の並び替え（部屋主を先頭に）
    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    # 保存
    try:
        final_df.to_csv(HISTORY_FILE, index=False, encoding='utf-8-sig')
        print(f" -> '{HISTORY_FILE}' を更新しました。(全 {len(final_df)} 件)")
    except PermissionError:
        print(f"【エラー】'{HISTORY_FILE}' が開けません。閉じてから再実行してください。")
else:
    print("新しいデータが取得できませんでした。履歴は更新されません。")
    final_df = history_df
    if '順番' in final_df.columns:
         final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce').fillna(0).astype(int)


# ==========================================
# ★集計・HTML生成処理 (以下は変更なし)
# ==========================================
print("HTML生成用のデータを準備中...")

analysis_html_content = "" 
ranking_html_content = "" 
cool_data_exists = False
ranking_data_list = [] 

cool_file = "cool_analysis.csv" 

# オフラインリスト
offline_files = [
    "offline_list_2026_1st.csv",
    "offline_list_2025_1st.csv",
    "offline_list_2025_2nd.csv"
]
offline_targets = []

for file_path in offline_files:
    if os.path.exists(file_path):
        try:
            offline_df = pd.read_csv(file_path)
            offline_df = offline_df.fillna("")
            if '曲名' in offline_df.columns:
                targets = [normalize_offline_text(str(x)) for x in offline_df['曲名'].tolist()]
                offline_targets.extend(targets)
                print(f"オフラインリスト({file_path})を読み込みました。")
            else:
                pass
        except Exception as e:
            print(f"オフラインリスト({file_path})読み込みエラー: {e}")

if not os.path.exists(cool_file):
    possible_files = [f for f in os.listdir('.') if f.endswith('.csv') and 'history' not in f and 'offline' not in f]
    if possible_files:
        cool_file = possible_files[0]

if cool_file and os.path.exists(cool_file):
    try:
        raw_df = None
        for enc in ['utf-8-sig', 'cp932', 'shift_jis']:
            try:
                raw_df = pd.read_csv(cool_file, header=None, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        
        if raw_df is not None:
            raw_df = raw_df.fillna("")
            raw_df = raw_df.drop_duplicates(keep='last')
            
            start_date = pd.to_datetime("2026/01/01")
            end_date = pd.to_datetime("2026/03/31")
            
            analysis_source_df = final_df.copy()
            if not analysis_source_df.empty and '取得日' in analysis_source_df.columns:
                analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
                
                # 曲名カラムの特定
                song_col = '曲名（ファイル名）' if '曲名（ファイル名）' in analysis_source_df.columns else ''
                if not song_col and '曲名' in analysis_source_df.columns: song_col = '曲名'
                
                if song_col:
                    analysis_source_df['norm_filename'] = analysis_source_df[song_col].apply(normalize_text)
                    
                    def get_rescued_workname(row):
                        raw_work = str(row['作品名']) if '作品名' in row and pd.notna(row['作品名']) else ""
                        raw_song = str(row[song_col]) if pd.notna(row[song_col]) else ""
                        if raw_work.strip() in ["-", "−", "", "nan"]:
                            match = re.search(r'【(.*?)】', raw_song)
                            if match:
                                return normalize_text(match.group(1))
                        return normalize_text(raw_work)

                    if '作品名' in analysis_source_df.columns:
                        analysis_source_df['norm_workname'] = analysis_source_df.apply(get_rescued_workname, axis=1)
                    else:
                        analysis_source_df['norm_workname'] = ""

                    exclude_keywords = ['test', 'テスト', 'システム', 'admin', 'System']
                    
                    # 歌った人カラムの特定
                    singer_col = '歌った人' if '歌った人' in analysis_source_df.columns else ''
                    
                    if singer_col:
                        target_history = analysis_source_df[
                            (analysis_source_df['dt_obj'] >= start_date) & 
                            (analysis_source_df['dt_obj'] <= end_date) &
                            (~analysis_source_df[singer_col].astype(str).apply(lambda x: any(k in x for k in exclude_keywords)))
                        ]

                        categorized_data = {}
                        ALLOWED_CATEGORIES = ["2026年冬アニメ", "2025年秋アニメ"]
                        current_category = None
                        
                        for idx, row in raw_df.iterrows():
                            if not any(str(x).strip() for x in row): continue
                            col0 = str(row[0]).strip()

                            is_category_line = any(cat in col0 for cat in ALLOWED_CATEGORIES) and "作品名" not in col0
                            if is_category_line:
                                current_category = col0
                                if current_category not in categorized_data:
                                    categorized_data[current_category] = []
                                continue
                            if "作品名" in col0: continue
                            if current_category is None: continue

                            anime = str(row[0]).strip() if len(row) > 0 else ""
                            type_ = str(row[1]).strip() if len(row) > 1 else ""
                            artist = str(row[2]).strip() if len(row) > 2 else ""
                            song = str(row[3]).strip() if len(row) > 3 else ""
                            if not anime and not song: continue

                            categorized_data[current_category].append({
                                "anime": anime, "type": type_, "artist": artist, "song": song
                            })

                        def check_match(target_text, source_series):
                            if not target_text:
                                return pd.Series([False] * len(source_series))
                            safe_target = re.escape(target_text)
                            if re.match(r'^[A-Z0-9\s]+$', target_text):
                                pattern = r'(?:^|[^A-Z0-9])' + safe_target + r'(?:[^A-Z0-9]|$)'
                                return source_series.str.contains(pattern, regex=True, case=False, na=False)
                            else:
                                return source_series.str.contains(safe_target, case=False, na=False)

                        for category, items in categorized_data.items():
                            analysis_html_content += f"""
                            <div class="category-block">
                                <div class="category-header" onclick="toggleCategory(this)">
                                    {category} <i class="fas fa-chevron-down" style="float:right;"></i>
                                </div>
                                <div class="category-content">
                                <table class="analysisTable">
                                    <thead>
                                        <tr>
                                            <th style="width:25%; min-width:180px;">作品名</th>
                                            <th style="width:5%; min-width:40px;">作成</th> <th style="width:10%; min-width:60px;">OP/ED</th>
                                            <th style="width:20%; min-width:150px;">歌手</th>
                                            <th style="width:25%; min-width:180px;">曲名</th>
                                            <th style="width:15%; min-width:60px;">歌唱数</th>
                                        </tr>
                                    </thead>
                            """
                            
                            items.sort(key=lambda x: x['anime'])
                            def get_anime_key(x): return x['anime']
                            
                            for anime_name, group_iter in groupby(items, key=get_anime_key):
                                group_items = list(group_iter)
                                rowspan = len(group_items)
                                analysis_html_content += '<tbody class="anime-group">'
                                
                                for i, item in enumerate(group_items):
                                    target_song_norm = normalize_text(item["song"])
                                    target_anime_norm = normalize_text(item["anime"])
                                    
                                    song_match_mask = check_match(target_song_norm, target_history['norm_filename'])
                                    anime_match_mask = (
                                        target_history['norm_filename'].str.contains(re.escape(target_anime_norm), case=False, na=False) |
                                        target_history['norm_workname'].str.contains(re.escape(target_anime_norm), case=False, na=False)
                                    )
                                    
                                    if target_song_norm and target_anime_norm:
                                        final_mask = song_match_mask & anime_match_mask
                                    elif target_song_norm:
                                        final_mask = song_match_mask
                                    elif target_anime_norm:
                                        final_mask = anime_match_mask
                                    else:
                                        final_mask = pd.Series([False] * len(target_history))

                                    count = len(target_history[final_mask])
                                    
                                    creation_count = 0
                                    if target_song_norm:
                                        for offline_str in offline_targets:
                                            if target_song_norm in offline_str:
                                                if target_anime_norm:
                                                    if target_anime_norm in offline_str:
                                                        creation_count += 1
                                                else:
                                                    creation_count += 1

                                    ranking_data_list.append({
                                        "category": category,
                                        "anime": item["anime"],
                                        "song": item["song"],
                                        "artist": item["artist"],
                                        "type": item["type"],
                                        "count": count
                                    })

                                    row_class = ""
                                    if creation_count == 0:
                                        row_class = "gray-text"
                                    elif count == 0:
                                        row_class = "zero-count"
                                    else:
                                        row_class = "has-count"
                                    
                                    bar_width = min(count * 20, 150)
                                    bar_html = f'<div class="bar-chart" style="width:{bar_width}px;"></div>' if count > 0 else ""
                                    
                                    clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                                    search_word = f"{clean_anime} {item['song']}"
                                    link_tag_start = f'<a href="#host/search.php?searchword={search_word}" class="export-link" target="_blank">'
                                    
                                    analysis_html_content += f'<tr class="{row_class}">'
                                    if i == 0:
                                        analysis_html_content += f'<td rowspan="{rowspan}">{item["anime"]}</td>'
                                    
                                    analysis_html_content += f'<td align="center">{creation_count}</td>'
                                    analysis_html_content += f'<td align="center">{link_tag_start}{item["type"]}</a></td>'
                                    analysis_html_content += f'<td>{link_tag_start}{item["artist"]}</a></td>'
                                    analysis_html_content += f'<td>{link_tag_start}{item["song"]}</a></td>'
                                    analysis_html_content += f'<td class="count-cell"><div class="count-wrapper"><span class="count-num">{count}</span>{bar_html}</div></td>'
                                    analysis_html_content += '</tr>'
                                analysis_html_content += '</tbody>'
                            analysis_html_content += "</table></div></div>"
                            
                        cool_data_exists = True
                        
                        # --- ランキング生成 ---
                        for target_cat in ALLOWED_CATEGORIES:
                            if target_cat not in categorized_data: continue
                            
                            cat_items = [d for d in ranking_data_list if d["category"] == target_cat and d["count"] > 0]
                            cat_items.sort(key=lambda x: x["count"], reverse=True)
                            
                            ranking_html_content += f"""
                            <div class="category-block">
                                <div class="category-header" onclick="toggleCategory(this)">
                                    {target_cat} ランキング (TOP 20) <i class="fas fa-chevron-down" style="float:right;"></i>
                                </div>
                                <div class="category-content">
                                <table class="rankingTable">
                                    <thead>
                                        <tr>
                                            <th style="width:10%; min-width:60px;">順位</th>
                                            <th style="width:25%; min-width:180px;">作品名</th>
                                            <th style="width:25%; min-width:180px;">曲名</th>
                                            <th style="width:20%; min-width:150px;">歌手</th>
                                            <th style="width:20%; min-width:60px;">歌唱数</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                            """
                            
                            if not cat_items:
                                ranking_html_content += '<tr><td colspan="5" style="text-align:center; padding:20px;">歌唱データがありません</td></tr>'
                            else:
                                previous_count = None
                                current_rank = 0
                                for i, item in enumerate(cat_items):
                                    if item["count"] != previous_count:
                                        current_rank = i + 1
                                    
                                    if current_rank > 20:
                                        break
                                        
                                    previous_count = item["count"]
                                    
                                    rank_class = f"rank-{current_rank}" if current_rank <= 3 else "rank-normal"
                                    rank_display = f'<span class="rank-badge {rank_class}">{current_rank}</span>'
                                    if current_rank == 1: rank_display += ' <i class="fas fa-crown" style="color:#FFD700;"></i>'
                                    elif current_rank == 2: rank_display += ' <i class="fas fa-medal" style="color:#C0C0C0;"></i>'
                                    elif current_rank == 3: rank_display += ' <i class="fas fa-medal" style="color:#CD7F32;"></i>'
                                    
                                    bar_width = min(item["count"] * 20, 150)
                                    bar_html = f'<div class="bar-chart" style="width:{bar_width}px;"></div>'
                                    
                                    clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                                    search_word = f"{clean_anime} {item['song']}"
                                    
                                    ranking_html_content += f"""
                                    <tr class="has-count ranking-row" data-href="#host/search.php?searchword={search_word}" onclick="onRankingClick(this)">
                                        <td align="center" style="font-weight:bold; font-size:1.1rem;">{rank_display}</td>
                                        <td>{item["anime"]} <span style="font-size:0.8em; color:#777;">({item["type"]})</span></td>
                                        <td>{item["song"]}</td>
                                        <td>{item["artist"]}</td>
                                        <td class="count-cell"><div class="count-wrapper"><span class="count-num">{item["count"]}</span>{bar_html}</div></td>
                                    </tr>
                                    """
                            ranking_html_content += "</tbody></table></div></div>"

    except Exception as e:
        print(f"集計エラー: {e}")

# ==========================================
# HTML生成
# ==========================================
columns_to_hide = ['コメント']
if not final_df.empty:
    html_df = final_df.drop(columns=columns_to_hide, errors='ignore')
else:
    html_df = pd.DataFrame()

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
            --card-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            margin: 0;
            padding: 20px;
            font-size: 14px;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
        }}
        h1 {{
            text-align: center;
            color: var(--primary-color);
            margin-bottom: 30px;
            font-size: 1.8rem;
        }}
        
        /* Tabs */
        .tabs {{
            display: flex;
            justify-content: center;
            margin-bottom: 20px;
            background: var(--header-bg);
            padding: 10px;
            border-radius: 8px;
            box-shadow: var(--card-shadow);
        }}
        .tab-btn {{
            padding: 10px 20px;
            border: none;
            background: none;
            cursor: pointer;
            font-size: 1rem;
            font-weight: bold;
            color: #777;
            border-bottom: 3px solid transparent;
            transition: all 0.3s;
        }}
        .tab-btn.active {{
            color: var(--accent-color);
            border-bottom-color: var(--accent-color);
        }}
        .tab-content {{
            display: none;
            animation: fadeIn 0.5s;
        }}
        .tab-content.active {{
            display: block;
        }}
        
        /* Controls */
        .controls {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: var(--header-bg);
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: var(--card-shadow);
            flex-wrap: wrap;
            gap: 10px;
        }}
        .search-box {{
            flex: 1;
            min-width: 200px;
            position: relative;
        }}
        .search-box input {{
            width: 100%;
            padding: 10px 10px 10px 35px;
            border: 1px solid var(--border-color);
            border-radius: 20px;
            outline: none;
            box-sizing: border-box;
        }}
        .search-box i {{
            position: absolute;
            left: 12px;
            top: 50%;
            transform: translateY(-50%);
            color: #aaa;
        }}
        .count-display {{
            font-weight: bold;
            color: var(--primary-color);
        }}

        /* Table */
        .table-container {{
            background: var(--header-bg);
            border-radius: 8px;
            box-shadow: var(--card-shadow);
            overflow-x: auto;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            white-space: nowrap;
        }}
        th, td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }}
        th {{
            background-color: #f8f9fa;
            color: var(--primary-color);
            cursor: pointer;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        th:hover {{
            background-color: #e9ecef;
        }}
        tr:hover {{
            background-color: #f1f1f1;
        }}
        .hidden {{
            display: none;
        }}

        /* Analysis & Ranking Styles */
        .category-block {{
            background: #fff;
            margin-bottom: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            overflow: hidden;
        }}
        .category-header {{
            background: #eef2f5;
            padding: 12px 20px;
            font-weight: bold;
            color: var(--primary-color);
            cursor: pointer;
            border-left: 5px solid var(--accent-color);
        }}
        .category-content {{
            padding: 0;
            display: block; 
        }}
        .analysisTable, .rankingTable {{
            width: 100%;
            font-size: 0.95em;
        }}
        .analysisTable th, .rankingTable th {{
            background: #fff;
            border-bottom: 2px solid #eee;
            color: #555;
            font-weight: 600;
        }}
        .analysisTable td, .rankingTable td {{
            padding: 8px 12px;
            vertical-align: middle;
        }}
        /* Analysis Specific */
        .gray-text {{ color: #ccc; }}
        .zero-count {{ background-color: #fff0f0; }}
        .has-count {{ background-color: #f0fff4; }}
        .anime-group {{ border-bottom: 1px solid #eee; }}
        .anime-group:last-child {{ border-bottom: none; }}
        
        /* Ranking Specific */
        .rank-badge {{
            display: inline-block;
            width: 24px;
            height: 24px;
            line-height: 24px;
            border-radius: 50%;
            text-align: center;
            font-size: 0.9em;
            color: #fff;
        }}
        .rank-1 {{ background-color: #FFD700; text-shadow: 0 1px 1px rgba(0,0,0,0.3); }}
        .rank-2 {{ background-color: #C0C0C0; text-shadow: 0 1px 1px rgba(0,0,0,0.3); }}
        .rank-3 {{ background-color: #CD7F32; text-shadow: 0 1px 1px rgba(0,0,0,0.3); }}
        .rank-normal {{ background-color: #e0e0e0; color: #555; }}
        .ranking-row {{ cursor: pointer; transition: transform 0.1s; }}
        .ranking-row:hover {{ transform: translateX(5px); background-color: #e8f5e9; }}

        /* Bar Chart */
        .count-cell {{ position: relative; }}
        .count-wrapper {{ display: flex; align-items: center; gap: 10px; }}
        .bar-chart {{
            height: 8px;
            background: linear-gradient(90deg, var(--accent-color), #2ecc71);
            border-radius: 4px;
            opacity: 0.7;
        }}
        
        .export-link {{
            text-decoration: none;
            color: inherit;
        }}
        .export-link:hover {{
            text-decoration: underline;
            color: var(--accent-color);
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        @media (max-width: 768px) {{
            .controls {{ flex-direction: column; align-items: stretch; }}
            .search-box {{ width: 100%; }}
            th, td {{ padding: 10px; font-size: 12px; }}
        }}
    </style>
</head>
<body>

<div class="container">
    <h1><i class="fas fa-microphone-alt"></i> Karaoke Dashboard</h1>
    
    <div class="tabs">
        <button class="tab-btn active" onclick="openTab(event, 'setlist')">Setlist History</button>
        <button class="tab-btn" onclick="openTab(event, 'analysis')">Season Analysis</button>
        <button class="tab-btn" onclick="openTab(event, 'ranking')">Ranking</button>
    </div>

    <div id="setlist" class="tab-content active">
        <div class="controls">
            <div class="search-box">
                <i class="fas fa-search"></i>
                <input type="text" id="searchInput" onkeyup="performSearch()" placeholder="Search artist, song, etc...">
            </div>
            <div class="count-display" id="countDisplay"></div>
            <button onclick="resetFilter()" style="padding:8px 15px; border:1px solid #ccc; border-radius:5px; background:#fff; cursor:pointer;">Reset</button>
        </div>
        <div class="table-container">
            <table id="historyTable">
                <thead>
                    <tr>{setlist_headers}</tr>
                </thead>
                <tbody>
                    {setlist_rows}
                </tbody>
            </table>
        </div>
    </div>

    <div id="analysis" class="tab-content">
        <div style="text-align:center; margin-bottom:20px; color:#666;">
            <p>※ オフラインリスト登録済み(作成) かつ 歌唱履歴なし(0回) はグレーアウト表示</p>
        </div>
        {analysis_html_content if cool_data_exists else "<p style='text-align:center;'>分析用データ(cool_analysis.csv)が見つかりません。</p>"}
    </div>

    <div id="ranking" class="tab-content">
        <div style="text-align:center; margin-bottom:20px; color:#666;">
            <p>※ 今期の歌唱回数ランキング (TOP 20)</p>
        </div>
        {ranking_html_content if cool_data_exists else "<p style='text-align:center;'>ランキングデータがありません。</p>"}
    </div>

</div>

<script>
    function openTab(evt, tabName) {{
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tab-content");
        for (i = 0; i < tabcontent.length; i++) {{
            tabcontent[i].style.display = "none";
            tabcontent[i].classList.remove("active");
        }}
        tablinks = document.getElementsByClassName("tab-btn");
        for (i = 0; i < tablinks.length; i++) {{
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }}
        document.getElementById(tabName).style.display = "block";
        document.getElementById(tabName).classList.add("active");
        evt.currentTarget.className += " active";
    }}

    // 初期表示カウント
    const table = document.getElementById("historyTable");
    const searchInput = document.getElementById("searchInput");
    const countDisplay = document.getElementById("countDisplay");
    
    function updateCount() {{
        if(!table) return;
        const total = table.tBodies[0].rows.length;
        const visible = Array.from(table.tBodies[0].rows).filter(r => !r.classList.contains('hidden')).length;
        countDisplay.innerText = '表示: ' + visible + ' / ' + total;
    }}
    
    // 初期ロード時にカウント更新
    updateCount();

    function performSearch() {{
        const filter = searchInput.value.toUpperCase();
        const tr = table.tBodies[0].getElementsByTagName("tr");
        let visibleCount = 0;
        const total = tr.length;

        for (let i = 0; i < total; i++) {{
            let textValue = tr[i].textContent || tr[i].innerText;
            if (textValue.toUpperCase().indexOf(filter) > -1) {{
                tr[i].classList.remove('hidden');
                visibleCount++;
            }} else {{
                tr[i].classList.add('hidden');
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
    }}

    function toggleCategory(header) {{
        const content = header.nextElementSibling;
        const icon = header.querySelector('i');
        if (content.style.display === "none") {{
            content.style.display = "block";
            icon.classList.remove('fa-chevron-right');
            icon.classList.add('fa-chevron-down');
        }} else {{
            content.style.display = "none";
            icon.classList.remove('fa-chevron-down');
            icon.classList.add('fa-chevron-right');
        }}
    }}
    
    function onRankingClick(row) {{
        const searchWord = row.getAttribute('data-href').split('searchword=')[1];
        if(searchWord) {{
            const decoded = decodeURIComponent(searchWord);
            document.querySelector('.tab-btn[onclick*="setlist"]').click();
            const searchBox = document.getElementById('searchInput');
            searchBox.value = decoded;
            performSearch();
        }}
    }}
</script>

</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"HTMLファイルを生成しました: index.html")
