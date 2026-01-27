import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
import json
from itertools import groupby

# --- 時刻設定 ---
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
current_date_str = now.strftime("%Y/%m/%d")
current_datetime_str = now.strftime("%Y/%m/%d %H:%M")

# --- 設定: ポート番号と部屋主の名前の対応表 ---
room_map = {
    11000: "ゆーふうりん部屋", 11001: "ゆーふうりん部屋", 11002: "ゆーふうりん部屋", 11003: "ゆーふうりん部屋",
    11004: "ゆーふうりん部屋", 11005: "ゆーふうりん部屋", 11006: "ゆーふうりん部屋", 11007: "ゆーふうりん部屋",
    11008: "ゆーふうりん部屋", 11009: "ゆーふうりん部屋", 11021: "成田部屋", 11022: "成田部屋",
    11028: "タマ部屋", 11058: "すみた部屋", 11059: "つぼはち部屋", 11063: "なぎ部屋", 11064: "naoo部屋",
    11066: "芝ちゃん部屋", 11067: "crom部屋", 11068: "けんしん部屋", 11069: "けんちぃ部屋",
    11070: "黒河部屋", 11071: "黒河部屋", 11074: "tukinowa部屋", 11077: "v3部屋", 11078: "のんでるん部屋",
    11079: "まどか部屋", 11084: "タカヒロ部屋", 11085: "タカヒロ部屋", 11086: "タカヒロ部屋",
    11087: "MiO部屋", 11088: "ほっしー部屋", 11092: "ヒロ部屋", 11101: "えみち部屋", 11102: "るえ部屋",
    11103: "ながし部屋", 11104: "MrN部屋", 11105: "ヤマテル部屋", 11106: "冨塚部屋", 11107: "ブルーベリー部屋",
    11108: "コタ部屋", 11109: "姫部屋"
}

# --- 関数: テキスト正規化 ---
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

def check_match(target_text, source_series):
    if not target_text: return pd.Series([False] * len(source_series))
    safe_target = re.escape(target_text)
    if re.match(r'^[A-Z0-9\s]+$', target_text):
        pattern = r'(?:^|[^A-Z0-9])' + safe_target + r'(?:[^A-Z0-9]|$)'
        return source_series.str.contains(pattern, regex=True, case=False, na=False)
    else:
        return source_series.str.contains(safe_target, case=False, na=False)

# --- 1. 過去データ読み込み ---
history_file = "history.csv"
if os.path.exists(history_file):
    try:
        history_df = pd.read_csv(history_file, encoding='utf-8-sig').fillna("")
    except Exception: history_df = pd.DataFrame()
else: history_df = pd.DataFrame()

# --- 2. 新しいデータ取得 ---
target_ports = list(room_map.keys())
new_data_frames = []
print("データを取得中...")
for port in target_ports:
    try:
        response = requests.get(f"http://Ykr.moe:{port}/simplelist.php", timeout=30)
        response.raise_for_status()
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0].fillna("") 
            df['部屋主'] = room_map[port]
            df['取得日'] = current_date_str
            new_data_frames.append(df)
    except Exception: pass 

if new_data_frames:
    new_df = pd.concat(new_data_frames, ignore_index=True)
    combined_df = pd.concat([history_df, new_df], ignore_index=True)
    for col in ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']:
        if col in combined_df.columns: combined_df = combined_df[combined_df[col] != col]
    subset_cols = ['部屋主', '順番', '曲名（ファイル名）', '歌った人']
    existing_cols = [c for c in subset_cols if c in combined_df.columns]
    final_df = combined_df.drop_duplicates(subset=existing_cols, keep='first').fillna("")
    if '順番' in final_df.columns: final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
    final_df['temp_date'] = pd.to_datetime(final_df['取得日'], errors='coerce')
    final_df = final_df.sort_values(by=['temp_date', '順番'], ascending=[False, False]).drop(columns=['temp_date'])
    cols = list(final_df.columns)
    if '部屋主' in cols: cols.insert(0, cols.pop(cols.index('部屋主')))
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
ranking_count_html_content = ""
ranking_user_html_content = ""
cool_data_exists = False
ranking_data_list = [] 
graph_series_data_count = {} 
graph_series_data_user = {}
created_lists_html = ""
uncreated_lists_html = ""
cool_file = "cool_analysis.csv" 

# --- オフラインリスト読み込み ---
offline_files = ["offline_list_2026_1st.csv", "offline_list_2025_1st.csv", "offline_list_2025_2nd.csv"]
offline_targets = []
for file_path in offline_files:
    if os.path.exists(file_path):
        try:
            offline_df = pd.read_csv(file_path).fillna("")
            if '曲名' in offline_df.columns:
                targets = [normalize_offline_text(str(x)) for x in offline_df['曲名'].tolist()]
                offline_targets.extend(targets)
                print(f"オフラインリスト({file_path})を読み込みました。追加件数: {len(targets)}")
        except Exception: pass

# --- カテゴリ別リストHTML生成関数 ---
def generate_category_html_block(category_name, item_list):
    if not item_list: return ""
    item_list.sort(key=lambda x: x['anime'])
    html = f"""<div class="category-block"><div class="category-header" onclick="toggleCategory(this)">{category_name} <i class="fas fa-chevron-down" style="float:right;"></i></div>
    <div class="category-content"><table class="analysisTable"><thead><tr>
    <th style="width:30%; min-width:180px;">作品名</th><th style="width:10%; min-width:60px;">OP/ED</th><th style="width:25%; min-width:150px;">歌手</th><th style="width:35%; min-width:180px;">曲名</th>
    </tr></thead><tbody class="anime-group">"""
    for anime_name, group_iter in groupby(item_list, key=lambda x: x['anime']):
        group_items = list(group_iter)
        rowspan = len(group_items)
        for i, item in enumerate(group_items):
            clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
            search_word = f"{clean_anime} {item['song']}"
            link = f'<a href="#host/search.php?searchword={search_word}" class="export-link">'
            html += '<tr>'
            if i == 0: html += f'<td rowspan="{rowspan}">{item["anime"]}</td>'
            html += f'<td align="center">{link}{item["type"]}</a></td><td>{link}{item["artist"]}</a></td><td>{link}{item["song"]}</a></td></tr>'
    html += '</tbody></table></div></div>'
    return html

if not os.path.exists(cool_file):
    possible = [f for f in os.listdir('.') if f.endswith('.csv') and 'history' not in f and 'offline' not in f]
    if possible: cool_file = possible[0]

if cool_file and os.path.exists(cool_file):
    try:
        raw_df = None
        for enc in ['utf-8-sig', 'cp932', 'shift_jis']:
            try: raw_df = pd.read_csv(cool_file, header=None, encoding=enc); break
            except UnicodeDecodeError: continue
        
        if raw_df is not None:
            raw_df = raw_df.fillna("").drop_duplicates(keep='last')
            
            # --- グラフ用データ準備 ---
            analysis_source_df = final_df.copy()
            analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
            analysis_source_df = analysis_source_df.dropna(subset=['dt_obj'])
            analysis_source_df['norm_filename'] = analysis_source_df['曲名（ファイル名）'].apply(normalize_text)
            
            def get_rescued_workname(row):
                raw_work = str(row['作品名']) if pd.notna(row['作品名']) else ""
                raw_song = str(row['曲名（ファイル名）']) if pd.notna(row['曲名（ファイル名）']) else ""
                if raw_work.strip() in ["-", "−", "", "nan"]:
                    match = re.search(r'【(.*?)】', raw_song)
                    if match: return normalize_text(match.group(1))
                return normalize_text(raw_work)
            
            if '作品名' in analysis_source_df.columns:
                analysis_source_df['norm_workname'] = analysis_source_df.apply(get_rescued_workname, axis=1)
            else: analysis_source_df['norm_workname'] = ""

            exclude_keywords = ['test', 'テスト', 'システム', 'admin', 'System']
            full_history = analysis_source_df[
                (~analysis_source_df['歌った人'].astype(str).apply(lambda x: any(k in x for k in exclude_keywords)))
            ].sort_values('dt_obj')
            
            target_history_display = full_history[(full_history['dt_obj'] >= pd.to_datetime("2026/01/01")) & (full_history['dt_obj'] <= pd.to_datetime("2026/03/31"))]

            categorized_data = {}
            ALLOWED_CATEGORIES = ["2026年冬アニメ", "2025年秋アニメ"]
            current_category = None
            for idx, row in raw_df.iterrows():
                if not any(str(x).strip() for x in row): continue
                col0 = str(row[0]).strip()
                if any(cat in col0 for cat in ALLOWED_CATEGORIES) and "作品名" not in col0:
                    current_category = col0
                    if current_category not in categorized_data: categorized_data[current_category] = []
                    continue
                if "作品名" in col0 or current_category is None: continue
                anime = str(row[0]).strip() if len(row) > 0 else ""
                type_ = str(row[1]).strip() if len(row) > 1 else ""
                artist = str(row[2]).strip() if len(row) > 2 else ""
                song = str(row[3]).strip() if len(row) > 3 else ""
                if anime or song:
                    categorized_data[current_category].append({"anime": anime, "type": type_, "artist": artist, "song": song})

            # --- グラフ計算 ---
            print("グラフデータ計算中...")
            graph_target_cat = "2026年冬アニメ"
            if graph_target_cat in categorized_data:
                winter_items = categorized_data[graph_target_cat]
                items_with_norm = [{"meta": i, "song_norm": normalize_text(i["song"]), "anime_norm": normalize_text(i["anime"]), "name": f"{i['anime']} - {i['song']}"} for i in winter_items]
                matched_records = []
                for idx, item in enumerate(items_with_norm):
                    song_pat, anime_pat = item["song_norm"], item["anime_norm"]
                    if not song_pat and not anime_pat: continue
                    song_match = check_match(song_pat, full_history['norm_filename'])
                    mask = None
                    if song_pat and anime_pat:
                        anime_match = (full_history['norm_filename'].str.contains(re.escape(anime_pat), case=False, na=False) | full_history['norm_workname'].str.contains(re.escape(anime_pat), case=False, na=False))
                        mask = song_match & anime_match
                    elif song_pat: mask = song_match
                    elif anime_pat:
                        mask = (full_history['norm_filename'].str.contains(re.escape(anime_pat), case=False, na=False) | full_history['norm_workname'].str.contains(re.escape(anime_pat), case=False, na=False))
                    if mask is not None:
                        for _, row in full_history[mask].iterrows(): matched_records.append({"date": row['dt_obj'], "item_idx": idx, "user": row['歌った人']})
                
                matched_records.sort(key=lambda x: x['date'])
                if matched_records:
                    unique_dates = sorted(list(set(r['date'] for r in matched_records)))
                    current_counts, current_users = {}, {}
                    rec_ptr, total_recs = 0, len(matched_records)
                    for current_dt in unique_dates:
                        dt_str = current_dt.strftime("%Y-%m-%d")
                        while rec_ptr < total_recs and matched_records[rec_ptr]['date'] <= current_dt:
                            rec = matched_records[rec_ptr]
                            idx, user = rec['item_idx'], rec['user']
                            current_counts[idx] = current_counts.get(idx, 0) + 1
                            if idx not in current_users: current_users[idx] = set()
                            current_users[idx].add(user)
                            rec_ptr += 1
                        # Count
                        ranking_src = [{"name": items_with_norm[k]["name"], "val": v} for k, v in current_counts.items()]
                        ranking_src.sort(key=lambda x: x['val'], reverse=True)
                        rank, prev = 1, -1
                        for i, d in enumerate(ranking_src):
                            if i > 0 and d['val'] < prev: rank = i + 1
                            prev = d['val']
                            if rank <= 20:
                                if d['name'] not in graph_series_data_count: graph_series_data_count[d['name']] = []
                                graph_series_data_count[d['name']].append({"x": dt_str, "y": rank})
                        # User
                        ranking_src = [{"name": items_with_norm[k]["name"], "val": len(v)} for k, v in current_users.items() if len(v)>0]
                        ranking_src.sort(key=lambda x: x['val'], reverse=True)
                        rank, prev = 1, -1
                        for i, d in enumerate(ranking_src):
                            if i > 0 and d['val'] < prev: rank = i + 1
                            prev = d['val']
                            if rank <= 20:
                                if d['name'] not in graph_series_data_user: graph_series_data_user[d['name']] = []
                                graph_series_data_user[d['name']].append({"x": dt_str, "y": rank})

            # --- 集計表生成 ---
            for category, items in categorized_data.items():
                cat_created_items, cat_uncreated_items = [], []
                analysis_html_content += f"""<div class="category-block"><div class="category-header" onclick="toggleCategory(this)">{category} <i class="fas fa-chevron-down" style="float:right;"></i></div>
                <div class="category-content"><table class="analysisTable"><thead><tr>
                <th style="width:25%; min-width:180px;">作品名</th><th style="width:5%; min-width:40px;">作成</th><th style="width:10%; min-width:60px;">OP/ED</th><th style="width:20%; min-width:150px;">歌手</th><th style="width:25%; min-width:180px;">曲名</th><th style="width:8%; min-width:60px;">人数</th><th style="width:15%; min-width:60px;">歌唱数</th>
                </tr></thead><tbody class="anime-group">"""
                
                for anime_name, group_iter in groupby(sorted(items, key=lambda x:x['anime']), key=lambda x:x['anime']):
                    group_items = list(group_iter)
                    rowspan = len(group_items)
                    for i, item in enumerate(group_items):
                        t_song, t_anime = normalize_text(item["song"]), normalize_text(item["anime"])
                        song_m = check_match(t_song, target_history_display['norm_filename'])
                        anime_m = (target_history_display['norm_filename'].str.contains(re.escape(t_anime), case=False, na=False) | target_history_display['norm_workname'].str.contains(re.escape(t_anime), case=False, na=False))
                        mask = (song_m & anime_m) if t_song and t_anime else (song_m if t_song else anime_m)
                        matched = target_history_display[mask]
                        cnt, u_cnt = len(matched), matched['歌った人'].nunique() if len(matched)>0 else 0
                        cre_cnt = 0
                        raw_song = normalize_offline_text(item["song"])
                        if t_song:
                            for off in offline_targets:
                                if (t_song in off or raw_song in off):
                                    if not t_anime or (t_anime in off): cre_cnt += 1
                        if cre_cnt >= 1: cat_created_items.append(item)
                        else: cat_uncreated_items.append(item)
                        
                        ranking_data_list.append({"category": category, "anime": item["anime"], "song": item["song"], "artist": item["artist"], "type": item["type"], "count": cnt, "user_count": u_cnt})
                        
                        clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                        search_word = f"{clean_anime} {item['song']}"
                        link = f'<a href="#host/search.php?searchword={search_word}" class="export-link">'
                        bar = f'<div class="bar-chart" style="width:{min(cnt*20,150)}px;"></div>' if cnt>0 else ""
                        u_bar = f'<div class="bar-chart-user" style="width:{min(u_cnt*20,100)}px;"></div>' if u_cnt>0 else ""
                        analysis_html_content += f'<tr class="has-count">'
                        if i == 0: analysis_html_content += f'<td rowspan="{rowspan}">{item["anime"]}</td>'
                        analysis_html_content += f'<td align="center">{cre_cnt}</td><td align="center">{link}{item["type"]}</a></td><td>{link}{item["artist"]}</a></td><td>{link}{item["song"]}</a></td><td class="count-cell"><div class="count-wrapper"><span class="count-num">{u_cnt}</span>{u_bar}</div></td><td class="count-cell"><div class="count-wrapper"><span class="count-num">{cnt}</span>{bar}</div></td></tr>'
                analysis_html_content += '</tbody></table></div></div>'
                created_lists_html += generate_category_html_block(category, cat_created_items)
                uncreated_lists_html += generate_category_html_block(category, cat_uncreated_items)
            cool_data_exists = True

            # --- ランキング表生成 ---
            def generate_ranking_html(mode="count"):
                html_out = ""
                for target_cat in ALLOWED_CATEGORIES:
                    if target_cat not in categorized_data: continue
                    cat_items = [d for d in ranking_data_list if d["category"] == target_cat and d["count"] > 0]
                    key = "count" if mode == "count" else "user_count"
                    cat_items.sort(key=lambda x: (x[key], x["count"] if mode=="user" else x["user_count"]), reverse=True)
                    title = f"{target_cat} 歌唱{'数' if mode=='count' else '人数'}ランキング (TOP 20)"
                    html_out += f'<div class="category-block"><div class="category-header" onclick="toggleCategory(this)">{title} <i class="fas fa-chevron-down" style="float:right;"></i></div><div class="category-content"><table class="rankingTable"><thead><tr><th style="width:10%;">順位</th><th style="width:25%;">作品名</th><th style="width:25%;">曲名</th><th style="width:15%;">歌手</th><th style="width:10%;">人数</th><th style="width:15%;">歌唱数</th></tr></thead><tbody>'
                    if not cat_items: html_out += '<tr><td colspan="6" style="text-align:center;padding:20px;">データなし</td></tr>'
                    else:
                        prev, rank = None, 0
                        for i, item in enumerate(cat_items):
                            curr = item[key]
                            if curr != prev: rank = i + 1
                            if rank > 20: break
                            prev = curr
                            rank_cls = f"rank-{rank}" if rank<=3 else "rank-normal"
                            rank_row = f"rank-row-{rank}" if rank<=3 else ""
                            badge = f'<span class="rank-badge {rank_cls}">{rank}</span>'
                            if rank==1: badge+=' <i class="fas fa-crown" style="color:#FFD700;"></i>'
                            clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                            link_w = f"{clean_anime} {item['song']}"
                            bar = f'<div class="bar-chart" style="width:{min(item["count"]*20,150)}px;"></div>'
                            u_bar = f'<div class="bar-chart-user" style="width:{min(item["user_count"]*20,100)}px;"></div>' if item["user_count"]>0 else ""
                            html_out += f'<tr class="has-count ranking-row {rank_row}" data-href="#host/search.php?searchword={link_w}" onclick="onRankingClick(this)"><td align="center" style="font-weight:bold;font-size:1.1rem;">{badge}</td><td>{item["anime"]}</td><td>{item["song"]}</td><td>{item["artist"]}</td><td class="count-cell"><div class="count-wrapper"><span class="count-num">{item["user_count"]}</span>{u_bar}</div></td><td class="count-cell"><div class="count-wrapper"><span class="count-num">{item["count"]}</span>{bar}</div></td></tr>'
                    html_out += '</tbody></table></div></div>'
                return html_out
            ranking_count_html_content = generate_ranking_html("count")
            ranking_user_html_content = generate_ranking_html("user")
            print("ランキング生成完了。")
    except Exception as e: print(f"集計エラー: {e}")

# HTML生成
html_df = final_df.drop(columns=['コメント'], errors='ignore') if not final_df.empty else pd.DataFrame()
setlist_rows = "".join(['<tr>' + "".join([f'<td>{val}</td>' for val in row]) + '</tr>' for _, row in html_df.iterrows()])
setlist_headers = "".join([f'<th onclick="sortTable({i})">{col} <i class="fas fa-sort"></i></th>' for i, col in enumerate(html_df.columns)])
graph_json_count = json.dumps(graph_series_data_count, ensure_ascii=False)
graph_json_user = json.dumps(graph_series_data_user, ensure_ascii=False)

html_content = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Karaoke Dashboard</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns/dist/chartjs-adapter-date-fns.bundle.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <style>
        :root {{ --primary-color: #2c3e50; --accent-color: #3498db; --bg-color: #f4f7f6; --text-color: #333; --header-bg: #fff; --border-color: #e0e0e0; }}
        html, body {{ height: 100%; margin: 0; padding: 0; overflow: hidden; font-family: "Helvetica Neue", Arial, sans-serif; background-color: var(--bg-color); color: var(--text-color); font-size: 13px; display: flex; flex-direction: column; }}
        a.export-link {{ color: inherit; text-decoration: none; pointer-events: none; cursor: default; }}
        tr.ranking-row {{ cursor: default; }}
        th, td {{ padding: 5px 8px; text-align: left; border-bottom: 1px solid #eee; font-size: 13px; vertical-align: middle; line-height: 1.3; }}
        th {{ background-color: var(--primary-color); color: #fff; position: sticky; top: 0; z-index: 10; font-weight: bold; cursor: pointer; }}
        .top-section {{ flex: 0 0 auto; background-color: var(--header-bg); box-shadow: 0 2px 5px rgba(0,0,0,0.1); z-index: 100; }}
        .header-inner {{ padding: 8px 15px; display: flex; justify-content: space-between; align-items: center; }}
        h1 {{ margin: 0; font-size: 1.2rem; color: var(--primary-color); }}
        .update-time {{ font-size: 0.8rem; color: #7f8c8d; }}
        .tabs {{ display: flex; padding: 0 15px; border-bottom: 1px solid var(--border-color); overflow-x: auto; }}
        .tab-btn {{ padding: 10px 15px; cursor: pointer; border: none; background: none; font-weight: bold; color: #7f8c8d; border-bottom: 3px solid transparent; font-size: 14px; white-space: nowrap; }}
        .tab-btn.active {{ color: var(--accent-color); border-bottom-color: var(--accent-color); }}
        .controls-row {{ padding: 8px 15px; display: flex; gap: 8px; align-items: center; background-color: #fff; border-bottom: 1px solid var(--border-color); height: 40px; flex-wrap: nowrap; overflow-x: auto; }}
        .search-box {{ padding: 6px 12px; border: 1px solid #ccc; border-radius: 4px; width: 250px; font-size: 13px; outline: none; }}
        .btn {{ padding: 6px 12px; border-radius: 4px; border: none; cursor: pointer; color: #fff; background-color: var(--accent-color); font-size: 13px; font-weight: bold; white-space: nowrap; }}
        .btn:hover {{ opacity: 0.9; }}
        .btn-dl {{ background-color: #2ecc71; }}
        .btn-list {{ background-color: #9b59b6; font-size: 12px; }}
        .count-display {{ margin-left: auto; font-weight: bold; font-size: 13px; }}
        .ctrl-group {{ display: none; width: 100%; align-items: center; gap:8px; }}
        .ctrl-group.active {{ display: flex; }}
        .ctrl-right {{ margin-left: auto; display: flex; gap: 8px; }}
        .content-area {{ flex: 1; position: relative; overflow: hidden; }}
        .tab-content {{ display: none; position: absolute; top: 0; left: 0; right: 0; bottom: 0; overflow-y: auto; -webkit-overflow-scrolling: touch; padding: 0 15px 40px 15px; }}
        .tab-content.active {{ display: block; }}
        table {{ width: 100%; border-collapse: separate; border-spacing: 0; background: #fff; border-radius: 4px; margin-top: 10px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
        tr:nth-child(even) {{ background-color: #fafafa; }}
        tr:hover {{ background-color: #f1f8ff; }}
        tr.hidden {{ display: none !important; }}
        .category-header {{ margin-top: 20px; padding: 10px 15px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 6px; font-weight: bold; font-size: 1.1rem; cursor: pointer; user-select: none; }}
        .category-content {{ display: block; transition: all 0.3s; }}
        .category-content.collapsed {{ display: none; }}
        tr.has-count {{ background-color: #fff; color: #333; }}
        .count-wrapper {{ display: flex; align-items: center; gap: 8px; }}
        .count-num {{ width: 25px; text-align: right; font-size:1.1rem; }}
        .bar-chart {{ height: 10px; background: linear-gradient(90deg, #3498db, #2980b9); border-radius: 5px; }}
        .bar-chart-user {{ height: 10px; background: linear-gradient(90deg, #2ecc71, #27ae60); border-radius: 5px; }}
        td[rowspan] {{ background-color: #fff; border-right: 1px solid #eee; vertical-align: middle; font-weight: normal; color: inherit; }}
        .rank-badge {{ display: inline-block; width: 24px; height: 24px; line-height: 24px; border-radius: 50%; text-align: center; color: #fff; font-weight: bold; font-size: 12px; background-color: #95a5a6; }}
        .rank-1 {{ background-color: #f1c40f; width: 28px; height: 28px; line-height: 28px; }}
        .rank-2 {{ background-color: #bdc3c7; }}
        .rank-3 {{ background-color: #d35400; }}
        tr.rank-row-1 td {{ background-color: #fff8e1 !important; }}
        tr.rank-row-2 td {{ background-color: #f5f5f5 !important; }}
        tr.rank-row-3 td {{ background-color: #fff0e6 !important; }}
        .rankingTable tr:nth-child(1) th {{ background-color: var(--primary-color) !important; color: #fff !important; }}
        
        /* グラフ用コンテナ */
        .chart-wrapper {{ background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-top: 15px; height: 75vh; display: flex; flex-direction: column; }}
        .chart-info {{ height: 35px; line-height: 35px; text-align: center; font-weight: bold; color: #2c3e50; background: #f1f3f5; border: 1px solid #e0e0e0; margin-bottom: 10px; border-radius: 4px; font-size: 14px; }}
        .canvas-container {{ flex: 1; position: relative; min-height: 0; }}
        @media print {{
             body {{ overflow: visible !important; height: auto !important; display: block !important; }}
             .top-section {{ display: none !important; }}
             .content-area {{ overflow: visible !important; position: static !important; }}
             .tab-content {{ position: static !important; display: block !important; overflow: visible !important; padding: 0 !important; }}
             .category-content {{ display: block !important; }}
             tbody.anime-group {{ break-inside: avoid; page-break-inside: avoid; }}
             .category-header {{ page-break-after: avoid; }}
             thead {{ display: table-header-group; }}
             .chart-wrapper {{ height: auto; }}
        }}
    </style>
</head>
<body>
    <div class="top-section">
        <div class="header-inner"><h1>Karaoke Dashboard</h1><div class="update-time">{current_datetime_str} 更新</div></div>
        <div class="tabs">
            <button class="tab-btn active" onclick="openTab('setlist', this)">セットリスト</button>
            <button class="tab-btn" onclick="openTab('analysis', this)">クール集計</button>
            <button class="tab-btn" onclick="openTab('ranking_count', this)">歌唱数ランキング</button>
            <button class="tab-btn" onclick="openTab('ranking_user', this)">歌唱人数ランキング</button>
            <button class="tab-btn" onclick="openTab('graph_view_count', this)">推移(数)</button>
            <button class="tab-btn" onclick="openTab('graph_view_user', this)">推移(人)</button>
        </div>
        <div class="controls-row">
            <div id="ctrl-setlist" class="ctrl-group active"><input type="text" id="searchInput" class="search-box" placeholder="キーワード..."><button onclick="performSearch()" class="btn"><i class="fas fa-search"></i> 検索</button><button onclick="resetFilter()" class="btn" style="background:#95a5a6"><i class="fas fa-undo"></i></button><div class="count-display" id="countDisplay">読み込み中...</div></div>
            <div id="ctrl-analysis" class="ctrl-group"><div class="ctrl-right"><button onclick="downloadList('list-created-content', 'created_list.html', '作成済みリスト')" class="btn btn-list">作成リスト保存</button><button onclick="downloadList('list-uncreated-content', 'uncreated_list.html', '未作成リスト')" class="btn btn-list" style="background-color:#e74c3c;">未作成リスト保存</button><button onclick="downloadHTML('print-target', 'karaoke_analysis.html', 'クール集計結果')" class="btn btn-dl" style="margin-left:10px;"><i class="fas fa-file-code"></i> HTML保存</button></div></div>
            <div id="ctrl-ranking-count" class="ctrl-group"><div class="ctrl-right"><button onclick="downloadHTML('ranking-count-print-target', 'karaoke_ranking_count.html', '歌唱数ランキング')" class="btn btn-dl"><i class="fas fa-trophy"></i> 歌唱数ランキング保存</button></div></div>
            <div id="ctrl-ranking-user" class="ctrl-group"><div class="ctrl-right"><button onclick="downloadHTML('ranking-user-print-target', 'karaoke_ranking_user.html', '歌唱人数ランキング')" class="btn btn-dl"><i class="fas fa-users"></i> 歌唱人数ランキング保存</button></div></div>
            <div id="ctrl-graph" class="ctrl-group"><div class="ctrl-right"><button onclick="downloadGraphPDF()" class="btn btn-dl" style="background-color:#e67e22;"><i class="fas fa-file-pdf"></i> PDF保存</button></div></div>
        </div>
    </div>
    <div class="content-area">
        <div id="setlist" class="tab-content active"><table id="setlistTable"><thead><tr>{setlist_headers}</tr></thead><tbody>{setlist_rows}</tbody></table>{"" if setlist_rows else '<div style="padding:20px;text-align:center">データがありません</div>'}</div>
        <div id="analysis" class="tab-content"><div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/03/31</div><div id="print-target">{analysis_html_content if cool_data_exists else '<div style="padding:20px;text-align:center;color:#e74c3c;">集計データがありません</div>'}</div></div>
        <div id="ranking_count" class="tab-content"><div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/03/31</div><div id="ranking-count-print-target">{ranking_count_html_content if ranking_count_html_content else '<div style="padding:20px;text-align:center;color:#e74c3c;">データがありません</div>'}</div></div>
        <div id="ranking_user" class="tab-content"><div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/03/31</div><div id="ranking-user-print-target">{ranking_user_html_content if ranking_user_html_content else '<div style="padding:20px;text-align:center;color:#e74c3c;">データがありません</div>'}</div></div>
        
        <div id="graph_view_count" class="tab-content">
            <div class="category-header">2026年冬アニメ 歌唱数ランキング推移 (Top 20)</div>
            <div class="chart-wrapper">
                <div id="chart-info-count" class="chart-info">グラフの点をタップ・ホバーで詳細を表示</div>
                <div class="canvas-container"><canvas id="rankingChartCount"></canvas></div>
            </div>
        </div>
        <div id="graph_view_user" class="tab-content">
            <div class="category-header">2026年冬アニメ 歌唱人数ランキング推移 (Top 20)</div>
            <div class="chart-wrapper">
                <div id="chart-info-user" class="chart-info">グラフの点をタップ・ホバーで詳細を表示</div>
                <div class="canvas-container"><canvas id="rankingChartUser"></canvas></div>
            </div>
        </div>
    </div>
    <div id="list-created-content" style="display:none;">{created_lists_html}</div>
    <div id="list-uncreated-content" style="display:none;">{uncreated_lists_html}</div>

<script>
    const host = 'http://ykr.moe:11059';
    const dataCount = {graph_json_count};
    const dataUser = {graph_json_user};
    let charts = {{ count: null, user: null }};
    const colors = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#46f0f0', '#f032e6', '#bcf60c', '#fabebe', '#008080', '#e6beff', '#9a6324', '#fffac8', '#800000', '#aaffc3', '#808000', '#ffd8b1', '#000075', '#808080'];

    function initChart(type, dataObj, canvasId) {{
        if(charts[type]) return;
        const ctx = document.getElementById(canvasId).getContext('2d');
        const infoDivId = type === 'count' ? 'chart-info-count' : 'chart-info-user';
        
        // 最新ランクでTop5判定
        const allKeys = Object.keys(dataObj);
        const latestRank = [];
        allKeys.forEach(key => {{
            const arr = dataObj[key];
            if(arr.length > 0) latestRank.push({{ key: key, rank: arr[arr.length - 1].y }});
        }});
        latestRank.sort((a,b) => a.rank - b.rank);
        const top5 = latestRank.slice(0, 5).map(x => x.key);

        const datasets = allKeys.map((key, i) => {{
            const color = colors[i % colors.length];
            const isTop5 = top5.includes(key);
            return {{
                label: key, data: dataObj[key],
                borderColor: color, backgroundColor: color,
                pointRadius: 4, pointHoverRadius: 8, tension: 0.1, fill: false, borderWidth: 2,
                hidden: !isTop5 // 初期表示はTop5のみ、他は非表示(取り消し線)
            }};
        }});

        charts[type] = new Chart(ctx, {{
            type: 'line', data: {{ datasets }},
            options: {{
                responsive: true, maintainAspectRatio: false,
                interaction: {{ mode: 'nearest', axis: 'x', intersect: true }},
                plugins: {{
                    tooltip: {{
                        enabled: false,
                        external: function(context) {{
                            const tooltip = context.tooltip;
                            const infoDiv = document.getElementById(infoDivId);
                            if (tooltip.opacity === 0) return;
                            if (tooltip.body) {{
                                const dataPoint = tooltip.dataPoints[0];
                                infoDiv.innerHTML = `<span style="color:${{dataPoint.dataset.borderColor}}">●</span> ${{dataPoint.dataset.label}} <span style="margin-left:10px; color:#555;">(${{dataPoint.label}}: ${{dataPoint.parsed.y}}位)</span>`;
                            }}
                        }}
                    }},
                    legend: {{ position: 'bottom', labels: {{ boxWidth: 10, padding: 15 }} }}
                }},
                scales: {{
                    y: {{ reverse: true, min: 0.5, max: 20.5, ticks: {{ stepSize: 1, callback: v => (v<1||v>20)?'':v }}, title: {{ display: true, text: '順位' }} }},
                    x: {{ type: 'time', time: {{ unit: 'day', displayFormats: {{ day: 'M/d' }} }}, title: {{ display: true, text: '日付' }} }}
                }}
            }}
        }});
    }}

    function downloadGraphPDF() {{
        const isCount = document.getElementById('graph_view_count').classList.contains('active');
        const id = isCount ? 'rankingChartCount' : 'rankingChartUser';
        const title = isCount ? "2026年冬アニメ 歌唱数ランキング推移" : "2026年冬アニメ 歌唱人数ランキング推移";
        const canvas = document.getElementById(id);
        const w = canvas.width, h = canvas.height;
        const newC = document.createElement('canvas'); newC.width=w; newC.height=h;
        const ctx = newC.getContext('2d'); ctx.fillStyle="#fff"; ctx.fillRect(0,0,w,h); ctx.drawImage(canvas,0,0);
        const pdf = new jsPDF({{ orientation: 'landscape' }});
        const ratio = Math.min(pdf.internal.pageSize.getWidth()/w, pdf.internal.pageSize.getHeight()/h)*0.85;
        pdf.text(title, 10, 10);
        pdf.addImage(newC.toDataURL('image/jpeg',1.0), 'JPEG', 10, 15, w*ratio, h*ratio);
        pdf.save("ranking_graph.pdf");
    }}

    function openTab(name, btn) {{
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.getElementById(name).classList.add('active');
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        if(btn) btn.classList.add('active'); else if(name==='setlist') document.querySelectorAll('.tab-btn')[0].classList.add('active');
        document.querySelectorAll('.ctrl-group').forEach(c => c.classList.remove('active'));
        if(name==='setlist') document.getElementById('ctrl-setlist').classList.add('active');
        else if(name.startsWith('analysis')) document.getElementById('ctrl-analysis').classList.add('active');
        else if(name==='ranking_count') document.getElementById('ctrl-ranking-count').classList.add('active');
        else if(name==='ranking_user') document.getElementById('ctrl-ranking-user').classList.add('active');
        else if(name==='graph_view_count') {{
            document.getElementById('ctrl-graph').classList.add('active');
            initChart('count', dataCount, 'rankingChartCount');
        }}
        else if(name==='graph_view_user') {{
            document.getElementById('ctrl-graph').classList.add('active');
            initChart('user', dataUser, 'rankingChartUser');
        }}
    }}
    function onRankingClick(row) {{ if(window.getSelection().toString().length>0)return; const u=row.getAttribute('data-href'); if(u)window.location.href=u.replace('#host',host); }}
    document.addEventListener('DOMContentLoaded',()=>{{ document.querySelectorAll('a.export-link').forEach(l=>l.href=l.getAttribute('href').replace('#host',host)); if(!document.querySelector('.tab-btn.active'))openTab('setlist',document.querySelectorAll('.tab-btn')[0]); }});
    function toggleCategory(h) {{ h.nextElementSibling.classList.toggle('collapsed'); const i=h.querySelector('i'); i.className=h.nextElementSibling.classList.contains('collapsed')?'fas fa-chevron-right':'fas fa-chevron-down'; i.style.float='right'; }}
    function downloadHTML(id, fn, t) {{ const c=document.getElementById(id).innerHTML; const b=new Blob([`<html><head><title>${{t}}</title><style>table{{width:100%;border-collapse:collapse}}th,td{{border:1px solid #ccc;padding:5px}}th{{background:#2c3e50;color:#fff}}.rank-badge{{display:inline-block;width:20px;background:#999;color:#fff;border-radius:50%;text-align:center}}.rank-1{{background:#f1c40f}}.category-header{{background:#667eea;color:#fff;padding:5px;margin-top:20px}}</style></head><body><h1>${{t}}</h1>${{c}}</body></html>`], {{type:'text/html'}}); const l=document.createElement('a'); l.href=URL.createObjectURL(b); l.download=fn; l.click(); }}
    function downloadList(id, fn, t) {{ if(document.getElementById(id)) downloadHTML(id, fn, t); }}
    const searchInput = document.getElementById("searchInput");
    const table = document.getElementById("setlistTable");
    const countDisplay = document.getElementById('countDisplay');
    let tbodyRows = [];
    window.onload = () => {{ if(table.tBodies[0]) {{ tbodyRows = Array.from(table.tBodies[0].rows); countDisplay.innerText = '全 ' + tbodyRows.length + ' 件'; }} }};
    function performSearch() {{ const k=searchInput.value.toUpperCase().replace(/　/g," ").split(" ").filter(s=>s); let c=0; tbodyRows.forEach(r=>{{ const m=k.every(w=>r.innerText.toUpperCase().includes(w)); r.classList.toggle('hidden',!m); if(m)c++; }}); countDisplay.innerText='表示: '+c+' / '+tbodyRows.length; }}
    function resetFilter() {{ searchInput.value=""; performSearch(); }}
    function sortTable(n) {{ const tb=table.tBodies[0], r=Array.from(tb.rows), th=table.querySelectorAll('th')[n], d=th.getAttribute('data-d')==='a'?'d':'a'; table.querySelectorAll('th').forEach(h=>h.setAttribute('data-d','')); th.setAttribute('data-d',d); r.sort((a,b)=>{{ const x=a.cells[n].innerText.trim(), y=b.cells[n].innerText.trim(); return d==='a'?(isNaN(x)?x.localeCompare(y,'ja'):x-y):(isNaN(x)?y.localeCompare(x,'ja'):y-x); }}); r.forEach(x=>tb.appendChild(x)); }}
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
    print("HTML生成完了: index.html")
