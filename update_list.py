import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
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
    11104: "MrN部屋",
    11105: "ヤマテル部屋",
    11106: "冨塚部屋",
    11107: "ブルーベリー部屋",
    11108: "コタ部屋",
    11109: "姫部屋"
}

# --- 関数: テキスト正規化 (検索キー用・履歴データ用) ---
def normalize_text(text):
    if not isinstance(text, str):
        return str(text)
    
    # 1. NFKC正規化
    text = unicodedata.normalize('NFKC', text)
    
    # 2. 拡張子の削除
    text = re.sub(r'\.[a-zA-Z0-9]{3,4}$', '', text)
    
    # 3. ★修正: 括弧の中身は削除せず、括弧記号のみをスペースに置換する
    # これにより【アンデッドアンラック】などが「アンデッドアンラック」として残り、マッチング可能になります。
    text = re.sub(r'[\[\(\{【\]\)\}】]', ' ', text)
    
    # 4. ★修正: キー変更情報を削除（ただし0から始まる数字はトラック番号とみなして残す）
    # (?!0) は「次が0ではない」場合のみマッチする -> -01, +02 などは消えずに残ります。
    text = re.sub(r'(key|KEY)?\s*[\+\-]\s*(?!0)[0-9]+', ' ', text)
    
    text = re.sub(r'原キー', ' ', text)
    text = re.sub(r'(キー)?変更[:：]?', ' ', text)
    
    # 5. 記号をスペースに置換
    text = re.sub(r'[~〜～\-_=,.]', ' ', text)
    
    # 6. スペース正規化
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text.upper()

# --- 関数: オフラインリスト用正規化 (括弧の中身を保持する) ---
def normalize_offline_text(text):
    if not isinstance(text, str):
        return str(text)
    
    # 1. NFKC正規化
    text = unicodedata.normalize('NFKC', text)
    
    # 2. 拡張子の削除
    text = re.sub(r'\.[a-zA-Z0-9]{3,4}$', '', text)
    
    # 3. ★修正: キー変更情報を削除（0から始まる数字は残す）
    text = re.sub(r'(key|KEY)?\s*[\+\-]\s*(?!0)[0-9]+', ' ', text)
    
    text = re.sub(r'原キー', ' ', text)
    text = re.sub(r'(キー)?変更[:：]?', ' ', text)
    
    # 4. 記号をスペースに置換 (ただし括弧類は削除しない！)
    text = re.sub(r'[~〜～\-_=,.]', ' ', text)
    
    # 5. スペース正規化
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text.upper()


# --- 1. 過去データ読み込み ---
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

# --- 2. 新しいデータ取得 ---
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
        pass 

if new_data_frames:
    new_df = pd.concat(new_data_frames, ignore_index=True)
    combined_df = pd.concat([history_df, new_df], ignore_index=True)

    clean_check_cols = ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']
    for col in clean_check_cols:
        if col in combined_df.columns:
            combined_df = combined_df[combined_df[col] != col]

    subset_cols = ['部屋主', '順番', '曲名（ファイル名）', '歌った人']
    existing_cols = [c for c in subset_cols if c in combined_df.columns]
    final_df = combined_df.drop_duplicates(subset=existing_cols, keep='first')
    final_df = final_df.fillna("")

    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
        
    final_df['temp_date'] = pd.to_datetime(final_df['取得日'], errors='coerce')
    final_df = final_df.sort_values(by=['temp_date', '順番'], ascending=[False, False])
    final_df = final_df.drop(columns=['temp_date'])
    
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
# ★集計処理
# ==========================================
analysis_html_content = "" 
ranking_html_content = "" 
cool_data_exists = False
ranking_data_list = [] 

created_lists_html = ""
uncreated_lists_html = ""

cool_file = "cool_analysis.csv" 

# --- オフラインリスト読み込み ---
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
                print(f"オフラインリスト({file_path})を読み込みました。追加件数: {len(targets)}")
            else:
                print(f"オフラインリスト({file_path})に'曲名'カラムが見つかりません。")
                
        except Exception as e:
            print(f"オフラインリスト({file_path})読み込みエラー: {e}")
    else:
        print(f"オフラインリスト({file_path})が見つかりません。")

print(f"オフラインリスト合計件数: {len(offline_targets)}")


# --- ★関数: カテゴリ別リストHTML生成 (メイン集計と同じレイアウト) ---
def generate_category_html_block(category_name, item_list):
    if not item_list:
        return ""
    
    # アニメ名でソート
    item_list.sort(key=lambda x: x['anime'])
    
    html = f"""
    <div class="category-block">
        <div class="category-header" onclick="toggleCategory(this)">
            {category_name} <i class="fas fa-chevron-down" style="float:right;"></i>
        </div>
        <div class="category-content">
        <table class="analysisTable">
            <thead>
                <tr>
                    <th style="width:30%; min-width:180px;">作品名</th>
                    <th style="width:10%; min-width:60px;">OP/ED</th>
                    <th style="width:25%; min-width:150px;">歌手</th>
                    <th style="width:35%; min-width:180px;">曲名</th>
                </tr>
            </thead>
    """
    
    def get_anime_key(x): return x['anime']
    
    for anime_name, group_iter in groupby(item_list, key=get_anime_key):
        group_items = list(group_iter)
        rowspan = len(group_items)
        
        html += '<tbody class="anime-group">'
        
        for i, item in enumerate(group_items):
            clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
            search_word = f"{clean_anime} {item['song']}"
            # 変更: target="_blank" を削除 (同じタブで開く)
            link_tag_start = f'<a href="#host/search.php?searchword={search_word}" class="export-link">'
            
            html += '<tr>'
            if i == 0:
                html += f'<td rowspan="{rowspan}">{item["anime"]}</td>'
            
            html += f'<td align="center">{link_tag_start}{item["type"]}</a></td>'
            html += f'<td>{link_tag_start}{item["artist"]}</a></td>'
            html += f'<td>{link_tag_start}{item["song"]}</a></td>'
            html += '</tr>'
        
        html += '</tbody>'
    
    html += "</table></div></div>"
    return html


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
                print(f"集計表({cool_file})をエンコーディング {enc} で読み込みました。")
                break
            except UnicodeDecodeError:
                continue
        
        if raw_df is not None:
            raw_df = raw_df.fillna("")
            
            print("CSV内の重複行を削除中...")
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

            # --- クール集計HTML生成 & リスト生成 ---
            for category, items in categorized_data.items():
                
                cat_created_items = []
                cat_uncreated_items = []

                # メイン集計用HTMLヘッダー
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
                        
                        # --- 歌唱数集計 ---
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
                        
                        # --- 作成数集計 ---
                        creation_count = 0
                        if target_song_norm:
                            for offline_str in offline_targets:
                                if target_song_norm in offline_str:
                                    if target_anime_norm:
                                        if target_anime_norm in offline_str:
                                            creation_count += 1
                                    else:
                                        creation_count += 1

                        # --- リストへの振り分け ---
                        if creation_count >= 1:
                            cat_created_items.append(item)
                        else:
                            cat_uncreated_items.append(item)

                        # ランキング用データ追加
                        ranking_data_list.append({
                            "category": category,
                            "anime": item["anime"],
                            "song": item["song"],
                            "artist": item["artist"],
                            "type": item["type"],
                            "count": count
                        })

                        # 行スタイル判定 (全て黒字)
                        row_class = "has-count"
                        
                        bar_width = min(count * 20, 150)
                        bar_html = f'<div class="bar-chart" style="width:{bar_width}px;"></div>' if count > 0 else ""
                        
                        clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                        search_word = f"{clean_anime} {item['song']}"
                        
                        # 変更: target="_blank" を削除 (同じタブ)
                        link_tag_start = f'<a href="#host/search.php?searchword={search_word}" class="export-link">'
                        
                        analysis_html_content += f'<tr class="{row_class}">'
                        if i == 0:
                            analysis_html_content += f'<td rowspan="{rowspan}">{item["anime"]}</td>'
                        
                        # 作成数カラム
                        analysis_html_content += f'<td align="center">{creation_count}</td>'

                        analysis_html_content += f'<td align="center">{link_tag_start}{item["type"]}</a></td>'
                        analysis_html_content += f'<td>{link_tag_start}{item["artist"]}</a></td>'
                        analysis_html_content += f'<td>{link_tag_start}{item["song"]}</a></td>'
                        
                        analysis_html_content += f'<td class="count-cell"><div class="count-wrapper"><span class="count-num">{count}</span>{bar_html}</div></td>'
                        analysis_html_content += '</tr>'
                    
                    analysis_html_content += '</tbody>'
                
                analysis_html_content += "</table></div></div>"

                # --- カテゴリごとのリストHTMLを生成して蓄積 ---
                created_lists_html += generate_category_html_block(category, cat_created_items)
                uncreated_lists_html += generate_category_html_block(category, cat_uncreated_items)

            cool_data_exists = True
            print("クール集計処理完了。")
            
            # ==========================================
            # ★ランキングHTML生成処理
            # ==========================================
            print("ランキング生成処理開始...")
            
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
            print("ランキング生成完了。")

        else:
            print("CSV読み込み失敗")

    except Exception as e:
        print(f"集計エラー: {e}")
        import traceback
        traceback.print_exc()


# ==========================================
# HTML生成 (HTML出力・印刷設定)
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
            flex-wrap: nowrap;
            overflow-x: auto;
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
        .btn-list {{ background-color: #9b59b6; font-size: 12px; }}
        .count-display {{ margin-left: auto; font-weight: bold; font-size: 13px; }}

        .ctrl-setlist {{ display: flex; width: 100%; align-items: center; gap:8px; }}
        .ctrl-analysis {{ display: none; width: 100%; align-items: center; justify-content: flex-end; gap:5px; }}
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
        
        /* 変更: グレーアウト用CSSを削除し、全て基本色で表示 */
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
                <button onclick="downloadList('list-created-content', 'created_list.html', '作成済みリスト')" class="btn btn-list">作成リスト保存</button>
                <button onclick="downloadList('list-uncreated-content', 'uncreated_list.html', '未作成リスト')" class="btn btn-list" style="background-color:#e74c3c;">未作成リスト保存</button>
                <button onclick="downloadHTML()" class="btn btn-dl" style="margin-left:10px;"><i class="fas fa-file-code"></i> HTML保存</button>
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

    <div id="list-created-content" style="display:none;">{created_lists_html}</div>
    <div id="list-uncreated-content" style="display:none;">{uncreated_lists_html}</div>

<script>
    function onRankingClick(row) {{
        if (window.getSelection().toString().length > 0) return;
        const rawHref = row.getAttribute('data-href');
        if (rawHref && rawHref.startsWith('#host')) {{
            const url = rawHref.replace('#host', host);
            // 変更: 同じタブで遷移するように変更
            window.location.href = url;
        }}
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

    function downloadList(elementId, filename, title) {{
        const element = document.getElementById(elementId);
        if(element) {{
            const htmlContent = element.innerHTML;
            generateDownload(htmlContent, filename, title);
        }}
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
        
        /* 変更: 保存ファイル内でもグレーアウト廃止 */
        tr.has-count {{ background-color: #fff; color: #333; }}
        
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
                // 変更: 同じタブで遷移
                window.location.href = url;
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
