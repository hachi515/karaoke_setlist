import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
import glob
import shutil
import sys  # 強制終了用
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

def get_ordinal_str(n):
    if 11 <= (n % 100) <= 13: suffix = 'th'
    else: suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

# --- 1. 過去データ読み込み ---
history_files = glob.glob("history*.csv")
history_dfs = []
initial_row_count = 0

print(f"過去ログファイルを読み込み中... ({len(history_files)}ファイル)")

if history_files:
    for f in history_files:
        try:
            # エクセルで開いているとここでエラーになる
            df = pd.read_csv(f, encoding='utf-8-sig')
            df = df.fillna("")
            history_dfs.append(df)
        except PermissionError:
            print(f"【致命的エラー】ファイル '{f}' が開けません。Excel等で開いていませんか？")
            print("データを保護するため、処理を中断します。")
            input("Enterキーを押して終了してください...")
            sys.exit() # 強制終了
        except Exception as e:
            print(f"【エラー】ファイル '{f}' の読み込みに失敗: {e}")
            print("安全のため処理を中断します。")
            sys.exit() # 強制終了
else:
    print("過去ログファイルが見つかりません。新規作成モードで動作します。")

if history_dfs:
    history_df = pd.concat(history_dfs, ignore_index=True)
    initial_row_count = len(history_df)
    print(f" -> 既存データ: {initial_row_count}件 読み込み完了")
else:
    history_df = pd.DataFrame()


# --- 2. 新しいデータ取得 (★修正: エラー時は即停止) ---
target_ports = list(room_map.keys())
new_data_frames = []
connection_error_occurred = False # エラーフラグ

print("各部屋のデータを取得中...")
for port in target_ports:
    url = f"http://Ykr.moe:{port}/simplelist.php"
    try:
        # タイムアウト20秒
        response = requests.get(url, timeout=20)
        response.raise_for_status()
        response.encoding = response.apparent_encoding
        
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0]
            if not df.empty:
                df = df.fillna("") 
                df['部屋主'] = room_map[port]
                df['取得日'] = current_date_str
                new_data_frames.append(df)
            else:
                # テーブルはあるが空の場合（正常な通信だが曲がない）
                pass
    except Exception as e:
        # ★重要: 通信エラー発生時はフラグを立てて、後で保存をブロックする
        print(f"【通信エラー】ポート {port} に接続できませんでした: {e}")
        connection_error_occurred = True

# ★安全装置：通信エラーがあった場合、保存せずに終了
if connection_error_occurred:
    print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print("【警告】通信エラーが発生したため、処理を中断します。")
    print("既存の履歴ファイルを守るため、保存は行いません。")
    print("（ここで保存すると、データが不完全な状態で上書きされる危険があります）")
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    input("Enterキーを押して終了してください...")
    sys.exit() # プログラム終了

if not new_data_frames:
    print("\n【警告】新しいデータが1件も取得できませんでした。")
    print("全ポートがオフラインの可能性があります。保存せずに終了します。")
    input("Enterキーを押して終了してください...")
    sys.exit()


# --- 3. データの結合・整理 ---
print("データを結合しています...")
final_df = history_df.copy()

if new_data_frames:
    new_df = pd.concat(new_data_frames, ignore_index=True)
    final_df = pd.concat([final_df, new_df], ignore_index=True)

# 処理対象データがある場合のみ実行
if not final_df.empty:
    # 不要ヘッダー行の削除
    clean_check_cols = ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']
    for col in clean_check_cols:
        if col in final_df.columns:
            final_df = final_df[final_df[col] != col]

    # ★重複削除（最低限の設定）
    # 日付・部屋・順番・曲名・歌手すべてが一致する場合のみ削除
    subset_cols = ['取得日', '部屋主', '順番', '曲名（ファイル名）', '歌った人']
    existing_cols = [c for c in subset_cols if c in final_df.columns]
    
    final_df = final_df.drop_duplicates(subset=existing_cols, keep='last')
    final_df = final_df.fillna("")
    
    # 順番の数値化とソート
    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
    
    final_df['temp_date'] = pd.to_datetime(final_df['取得日'], errors='coerce')
    final_df = final_df.sort_values(by=['temp_date', '順番'], ascending=[False, False])
    final_df = final_df.drop(columns=['temp_date'])
    
    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    # --- 4. 保存処理 (バックアップ + 一時ファイル保存) ---
    final_row_count = len(final_df)
    
    print("履歴ファイルを保存処理中...")
    
    # 1. バックアップ作成
    backup_dir = "backup"
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    for f in glob.glob("history_*.csv"):
        try:
            shutil.copy(f, os.path.join(backup_dir, f"{f}.{timestamp}.bak"))
        except: pass

    # 2. 一時ファイルとして保存 (temp_history_*.csv)
    # いきなり本番ファイルを消さず、まずTempを作る
    save_df = final_df.copy()
    save_df['temp_date_sort'] = pd.to_datetime(save_df['取得日'], errors='coerce')
    save_df = save_df.sort_values(by=['temp_date_sort', '順番'], ascending=[True, True])
    save_df = save_df.drop(columns=['temp_date_sort'])
    
    chunk_size = 4000
    total_rows = len(save_df)
    temp_files = []
    
    try:
        if total_rows == 0:
            print("保存するデータがありません。")
        else:
            for i in range(0, total_rows, chunk_size):
                chunk = save_df.iloc[i : i + chunk_size]
                file_num = (i // chunk_size) + 1
                suffix = get_ordinal_str(file_num)
                temp_filename = f"temp_history_{suffix}.csv" # 一時ファイル名
                
                chunk.to_csv(temp_filename, index=False, encoding='utf-8-sig')
                temp_files.append(temp_filename)
                
            # 3. 保存成功を確認してから、本番ファイルを置き換える
            # 古いhistoryファイルを削除
            for f in glob.glob("history_*.csv"):
                os.remove(f)
            
            # Tempファイルを本番名にリネーム
            for temp_f in temp_files:
                final_name = temp_f.replace("temp_", "")
                os.rename(temp_f, final_name)
                print(f" -> {final_name} を保存しました")
                
        print("履歴ファイルの更新完了。")
        
    except Exception as e:
        print(f"【保存エラー】ファイルの書き込み中にエラーが発生しました: {e}")
        print("一時ファイルが残っている可能性があります。元のhistoryファイルは保護されているか確認してください。")

else:
    print("データが存在しません。更新をスキップします。")


# ==========================================
# ★集計処理 (変更なし)
# ==========================================

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
    possible_files = [f for f in os.listdir('.') if f.endswith('.csv') and 'history' not in f and 'offline' not in f and 'backup' not in f and 'temp' not in f]
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
            analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
            
            analysis_source_df['norm_filename'] = analysis_source_df['曲名（ファイル名）'].apply(normalize_text)
            
            def get_rescued_workname(row):
                raw_work = str(row['作品名']) if pd.notna(row['作品名']) else ""
                raw_song = str(row['曲名（ファイル名）']) if pd.notna(row['曲名（ファイル名）']) else ""
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
            
            target_history = analysis_source_df[
                (analysis_source_df['dt_obj'] >= start_date) & 
                (analysis_source_df['dt_obj'] <= end_date) &
                (~analysis_source_df['歌った人'].astype(str).apply(lambda x: any(k in x for k in exclude_keywords)))
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
            
            for target_cat in ALLOWED_CATEGORIES:
                if target_cat not in categorized_data:
                    continue
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
                        if current_rank == 1:
                            rank_display += ' <i class="fas fa-crown" style="color:#FFD700;"></i>'
                        elif current_rank == 2:
                            rank_display += ' <i class="fas fa-medal" style="color:#C0C0C0;"></i>'
                        elif current_rank == 3:
                            rank_display += ' <i class="fas fa-medal" style="color:#CD7F32;"></i>'
                            
                        bar_width = min(item["count"] * 20, 150)
                        bar_html = f'<div class="bar-chart" style="width:{bar_width}px;"></div>'
                        clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                        search_word = f"{clean_anime} {item['song']}"
                        
                        ranking_html_content += f"""
                        <tr class="has-count ranking-row" data-href="#host/search.php?searchword={search_word}" onclick="onRankingClick(this)">
                            <td align="center" style="font-weight:bold; font-size:1.1rem;">{rank_display}</td>
                            <td>{item["anime"]} <span style="font-size:0.8em; color:#777;">({item["type"]})</span></td>
                            <td>{item["song"]}</td> <td>{item["artist"]}</td>
                            <td class="count-cell"><div class="count-wrapper"><span class="count-num">{item["count"]}</span>{bar_html}</div></td>
                        </tr>
                        """
                ranking_html_content += "</tbody></table></div></div>"
        else:
            print("CSV読み込み失敗")
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
        }}
        html, body {{
            height: 100%; margin: 0; padding: 0;
            overflow: hidden; 
            font-family: "Helvetica Neue", Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            font-size: 13px; 
            display: flex; flex-direction: column;
        }}

        a.export-link {{
            color: inherit;
            text-decoration: none;
            pointer-events: none;
            cursor: default;
        }}

        tr.ranking-row {{
            cursor: default; 
        }}
        
        th, td {{
            padding: 5px 8px; text-align: left; border-bottom: 1px solid #eee;
            font-size: 13px; vertical-align: middle; line-height: 1.3;
        }}
        th {{
            background-color: var(--primary-color); color: #fff;
            position: sticky; top: 0; z-index: 10; font-weight: bold;
        }}

        .top-section {{
            flex: 0 0 auto;
            background-color: var(--header-bg);
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            z-index: 100;
        }}
        .header-inner {{
            padding: 8px 15px; display: flex; justify-content: space-between; align-items: center;
        }}
        h1 {{ margin: 0; font-size: 1.2rem; color: var(--primary-color); }}
        .update-time {{ font-size: 0.8rem; color: #7f8c8d; }}

        .tabs {{
            display: flex; padding: 0 15px; border-bottom: 1px solid var(--border-color);
        }}
        .tab-btn {{
            padding: 10px 20px; cursor: pointer; border: none; background: none;
            font-weight: bold; color: #7f8c8d; border-bottom: 3px solid transparent;
            font-size: 14px;
        }}
        .tab-btn.active {{ color: var(--accent-color); border-bottom-color: var(--accent-color); }}

        .controls-row {{
            padding: 8px 15px; display: flex; gap: 8px; align-items: center;
            background-color: #fff; border-bottom: 1px solid var(--border-color);
            height: 40px; 
        }}
        .search-box {{
            padding: 6px 12px; border: 1px solid #ccc; border-radius: 4px;
            width: 250px; font-size: 13px; outline: none;
        }}
        .btn {{
            padding: 6px 12px; border-radius: 4px; border: none; cursor: pointer;
            color: #fff; background-color: var(--accent-color); font-size: 13px;
            font-weight: bold; white-space: nowrap;
        }}
        .btn:hover {{ opacity: 0.9; }}
        .btn-dl {{ background-color: #2ecc71; }}
        .count-display {{ margin-left: auto; font-weight: bold; font-size: 13px; }}

        .ctrl-setlist {{ display: flex; width: 100%; align-items: center; gap:8px; }}
        .ctrl-analysis {{ display: none; width: 100%; align-items: center; justify-content: flex-end; }}
        .ctrl-ranking {{ display: none; width: 100%; align-items: center; justify-content: flex-end; }}

        .content-area {{
            flex: 1; position: relative; overflow: hidden; 
        }}
        .tab-content {{
            display: none; position: absolute; 
            top: 0; left: 0; right: 0; bottom: 0;
            overflow-y: auto; 
            -webkit-overflow-scrolling: touch;
            padding: 0 15px 40px 15px;
        }}
        .tab-content.active {{ display: block; }}

        table {{
            width: 100%; border-collapse: separate; border-spacing: 0;
            background: #fff; border-radius: 4px; margin-top: 10px; margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}
        
        tr:nth-child(even) {{ background-color: #fafafa; }}
        tr:hover {{ background-color: #f1f8ff; }}
        tr.hidden {{ display: none !important; }}

        .category-header {{
            margin-top: 20px; padding: 10px 15px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; border-radius: 6px;
            font-weight: bold; font-size: 1.1rem; cursor: pointer;
            user-select: none;
        }}
        .category-content {{ display: block; transition: all 0.3s; }}
        .category-content.collapsed {{ display: none; }}
        
        tr.zero-count {{ color: #ccc; }}
        .gray-text {{ color: gray !important; }}
        tr.has-count {{ background-color: #fff; color: #333; }}
        
        .count-wrapper {{ display: flex; align-items: center; gap: 8px; }}
        .count-num {{ width: 25px; text-align: right; font-size:1.1rem; }}
        .bar-chart {{
            height: 10px; background: linear-gradient(90deg, #3498db, #2980b9);
            border-radius: 5px;
        }}
        td[rowspan] {{
            background-color: #fff;
            border-right: 1px solid #eee;
            vertical-align: middle;
            font-weight: normal; color: inherit;      
        }}

        .rank-badge {{
            display: inline-block; width: 24px; height: 24px; line-height: 24px;
            border-radius: 50%; text-align: center; color: #fff; font-weight: bold; font-size: 12px;
            background-color: #95a5a6; 
        }}
        .rank-1 {{ background-color: #f1c40f; box-shadow: 0 0 5px #f39c12; font-size: 14px; width: 28px; height: 28px; line-height: 28px; }} 
        .rank-2 {{ background-color: #bdc3c7; box-shadow: 0 0 5px #7f8c8d; }} 
        .rank-3 {{ background-color: #d35400; opacity: 0.8; }} 
        
        .rankingTable tr:nth-child(1) td {{ background-color: #fffae6; }}
        .rankingTable tr:nth-child(2) td {{ background-color: #f8f9fa; }}

        @media print {{
            * {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
                color-adjust: exact !important;
            }}
            body {{
                overflow: visible !important;
                height: auto !important;
                display: block !important;
            }}
            .top-section {{ display: none !important; }}
            .content-area {{ overflow: visible !important; position: static !important; }}
            .tab-content {{ 
                position: static !important; 
                display: block !important; 
                overflow: visible !important; 
                padding: 0 !important;
            }}
            .category-content {{ display: block !important; }}
            
            tbody.anime-group {{
                break-inside: avoid;
                page-break-inside: avoid;
            }}
            .category-header {{ page-break-after: avoid; }}
            thead {{ display: table-header-group; }}
        }}
    </style>
</head>
<body>
    <div class="top-section">
        <div class="header-inner">
            <h1>Karaoke Dashboard</h1>
            <div class="update-time">{current_datetime_str} 更新</div>
        </div>
        <div class="tabs">
            <button class="tab-btn active" onclick="openTab('setlist')">セットリスト</button>
            <button class="tab-btn" onclick="openTab('analysis')">クール集計</button>
            <button class="tab-btn" onclick="openTab('ranking')">ランキング</button>
        </div>
        <div class="controls-row">
            <div id="ctrl-setlist" class="ctrl-setlist">
                <input type="text" id="searchInput" class="search-box" placeholder="キーワード (例: 曲名 歌手)...">
                <button onclick="performSearch()" class="btn"><i class="fas fa-search"></i> 検索</button>
                <button onclick="resetFilter()" class="btn" style="background:#95a5a6"><i class="fas fa-undo"></i></button>
                <div class="count-display" id="countDisplay">読み込み中...</div>
            </div>
            <div id="ctrl-analysis" class="ctrl-analysis">
                <button onclick="downloadHTML()" class="btn btn-dl"><i class="fas fa-file-code"></i> HTML保存</button>
            </div>
            <div id="ctrl-ranking" class="ctrl-ranking">
                <button onclick="downloadRanking()" class="btn btn-dl"><i class="fas fa-trophy"></i> ランキング保存</button>
            </div>
        </div>
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
            <div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/03/31</div>
            <div id="print-target">
                {analysis_html_content if cool_data_exists else '<div style="padding:20px;text-align:center;color:#e74c3c;">集計データがありません</div>'}
            </div>
        </div>

        <div id="ranking" class="tab-content">
            <div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/03/31</div>
            <div id="ranking-print-target">
                {ranking_html_content if ranking_html_content else '<div style="padding:20px;text-align:center;color:#e74c3c;">ランキング対象データがありません</div>'}
            </div>
        </div>
    </div>

<script>
    function onRankingClick(row) {{
    }}

    function openTab(tabName) {{
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.getElementById(tabName).classList.add('active');
        
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        
        let btnIndex = 0;
        if (tabName === 'analysis') btnIndex = 1;
        if (tabName === 'ranking') btnIndex = 2;
        
        document.querySelectorAll('.tab-btn')[btnIndex].classList.add('active');
        
        document.getElementById('ctrl-setlist').style.display = 'none';
        document.getElementById('ctrl-analysis').style.display = 'none';
        document.getElementById('ctrl-ranking').style.display = 'none';

        if(tabName === 'setlist') {{
            document.getElementById('ctrl-setlist').style.display = 'flex';
        }} else if(tabName === 'analysis') {{
            document.getElementById('ctrl-analysis').style.display = 'flex';
        }} else if(tabName === 'ranking') {{
            document.getElementById('ctrl-ranking').style.display = 'flex';
        }}
    }}

    function toggleCategory(header) {{
        const content = header.nextElementSibling;
        content.classList.toggle('collapsed');
        const icon = header.querySelector('i');
        icon.className = content.classList.contains('collapsed') ? 'fas fa-chevron-right' : 'fas fa-chevron-down';
        icon.style.float = 'right';
    }}

    function downloadHTML() {{
        const element = document.getElementById('print-target');
        const htmlContent = element.innerHTML;
        generateDownload(htmlContent, 'karaoke_analysis.html', 'クール集計結果');
    }}

    function downloadRanking() {{
        const element = document.getElementById('ranking-print-target');
        const htmlContent = element.innerHTML;
        generateDownload(htmlContent, 'karaoke_ranking.html', 'カラオケ歌唱ランキング');
    }}

    function generateDownload(content, filename, title) {{
        const fullHtml = `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>${{title}}</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {{ font-family: "Helvetica Neue", Arial, sans-serif; font-size: 13px; color: #333; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
        th, td {{ border: 1px solid #ccc; padding: 5px 8px; text-align: left; vertical-align: middle; }}
        th {{ background-color: #2c3e50; color: #fff; }}
        td[rowspan] {{ background-color: #fff; }}
        
        .category-header {{ 
            background: #667eea; color: white; padding: 10px; margin-top: 20px; 
            font-weight: bold; border-radius: 4px; cursor: pointer; user-select: none;
        }}
        .category-content {{ display: block; }}
        .category-content.collapsed {{ display: none; }}
        
        .gray-text {{ color: gray !important; }}
        tr.zero-count {{ color: #ccc; }}
        
        a.export-link {{
            display: block; 
            margin: -5px -8px; 
            padding: 5px 8px;  
            color: #333; 
            text-decoration: none; 
            box-sizing: border-box;
            cursor: pointer;
        }}
        a.export-link:hover {{ background-color: #eef2f7; color: #3498db; }}
        
        tr.ranking-row {{ cursor: pointer; }}
        tr.ranking-row:hover {{ background-color: #dbeafe; }}
        
        .count-wrapper {{ display: flex; align-items: center; gap: 8px; }}
        .count-num {{ width: 25px; text-align: right; }}
        .bar-chart {{ height: 10px; background: #3498db; border-radius: 5px; }}
        
        .rank-badge {{
            display: inline-block; width: 24px; height: 24px; line-height: 24px;
            border-radius: 50%; text-align: center; color: #fff; font-weight: bold; font-size: 12px;
            background-color: #95a5a6;
        }}
        .rank-1 {{ background-color: #f1c40f; width: 28px; height: 28px; line-height: 28px; }}
        .rank-2 {{ background-color: #bdc3c7; }}
        .rank-3 {{ background-color: #d35400; }}
        .rankingTable tr:nth-child(1) td {{ background-color: #fffae6; }}
        .rankingTable tr:nth-child(2) td {{ background-color: #f8f9fa; }}

        @media print {{
            * {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
                color-adjust: exact !important;
            }}
            .category-content {{ display: block !important; }}
            tbody.anime-group {{ break-inside: avoid; page-break-inside: avoid; }}
            .category-header {{ page-break-after: avoid; }}
            thead {{ display: table-header-group; }}
        }}
    </style>
</head>
<body>
    <h1>${{title}}</h1>
    <div style="text-align:right; font-size:0.9rem; color:#777;">出力日: {current_date_str}</div>
    ${{content}}

    <script>
        const host = 'http://ykr.moe:11059';

        function onRankingClick(row) {{
            if (window.getSelection().toString().length > 0) return;
            const rawHref = row.getAttribute('data-href');
            if (rawHref && rawHref.startsWith('#host')) {{
                const url = rawHref.replace('#host', host);
                window.open(url, '_blank');
            }}
        }}

        document.addEventListener('DOMContentLoaded', () => {{
            document.querySelectorAll('a.export-link').forEach(link => {{
                const rawHref = link.getAttribute('href');
                if (rawHref && rawHref.startsWith('#host')) {{
                    link.href = rawHref.replace('#host', host);
                }}
            }});
        }});

        function toggleCategory(header) {{
            const content = header.nextElementSibling;
            content.classList.toggle('collapsed');
            const icon = header.querySelector('i');
            if(icon) {{
                icon.className = content.classList.contains('collapsed') ? 'fas fa-chevron-right' : 'fas fa-chevron-down';
                icon.style.float = 'right';
            }}
        }}
    <\/script>
</body>
</html>`;

        const blob = new Blob([fullHtml], {{type: 'text/html'}});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
    }}

    const searchInput = document.getElementById("searchInput");
    const table = document.getElementById("setlistTable");
    const countDisplay = document.getElementById('countDisplay');
    let tableData = [];
    let tbodyRows = [];

    window.addEventListener('DOMContentLoaded', () => {{
        const tbody = table.tBodies[0];
        if (tbody) {{
            tbodyRows = Array.from(tbody.rows);
            tableData = tbodyRows.map(row => row.innerText.toUpperCase());
            countDisplay.innerText = '全 ' + tbodyRows.length + ' 件';
        }}
    }});

    searchInput.addEventListener("keyup", function(event) {{
        if (event.key === "Enter") performSearch();
    }});

    function performSearch() {{
        const filter = searchInput.value.toUpperCase();
        const keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        let visibleCount = 0;
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
