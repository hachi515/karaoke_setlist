import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
import json
import io
from concurrent.futures import ThreadPoolExecutor, as_completed

# ==========================================
# ★ 設定: GitHubリポジトリ情報
# ==========================================
GITHUB_USER = "hachi515"
GITHUB_REPO = "karaoke_setlist"
GITHUB_BRANCH = "main"

OFFLINE_FILES = [
    "offline_list_2026_1st.csv",
    "offline_list_2025_2nd.csv",
    "offline_list_2025_1st.csv"
]

# ==========================================
# ★ GAS連携・GitHub読み込み用設定
# ==========================================
GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyzKEPfj0bYcRyEdizwQXcduIOQFt2_njtFQSyGP9jBjrhR8pyVKwDol6VN7bLPrktq/exec"
CSV_EMPTY_PREFIX_BYTES = b'\xef\xbb\xbf\r\n\t '
EXPECTED_HISTORY_COLUMNS = [
    '取得日', '部屋主', '順番', '曲名（ファイル名）', '作品名', '歌手名', '歌った人'
]


def load_df_from_github(filename, **kwargs):
    url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{filename}"
    print(f"[GitHub] Loading {filename} from {url}...")
    try:
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            content_bytes = response.content
            for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                try:
                    df = pd.read_csv(io.BytesIO(content_bytes), encoding=enc, engine='python', **kwargs)
                    if not df.empty:
                        df.columns = df.columns.astype(str).str.replace('\ufeff', '').str.strip()
                    print(f"[GitHub] Success: Loaded {filename} (Encoding: {enc}). Rows: {len(df)}")
                    return df
                except Exception:
                    continue
            print(f"[GitHub] Failed to decode {filename}.")
            return pd.DataFrame()
        elif response.status_code == 404:
            print(f"[GitHub] File not found: {filename} (404).")
            return pd.DataFrame()
        else:
            print(f"[GitHub] Error fetching {filename}: Status {response.status_code}")
            return pd.DataFrame()
    except Exception as e:
        print(f"[GitHub] Connection error for {filename}: {e}")
        return pd.DataFrame()


def load_df_from_gas_with_status(filename, **kwargs):
    print(f"[GAS] Loading {filename}...")
    try:
        response = requests.get(GAS_WEB_APP_URL, params={'filename': filename}, timeout=60)
    except Exception as e:
        print(f"[GAS] Connection error: {e}")
        return pd.DataFrame(), "error"
    if response.status_code == 404:
        return pd.DataFrame(), "not_found"
    if response.status_code != 200:
        return pd.DataFrame(), "error"
    content_bytes = response.content
    response_text = response.text if isinstance(response.text, str) else ""
    if "Exception: Service error: Drive" in response_text:
        return pd.DataFrame(), "error"
    if not content_bytes or content_bytes.lstrip(CSV_EMPTY_PREFIX_BYTES) == b'':
        return pd.DataFrame(), "empty"
    for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
        try:
            df = pd.read_csv(io.BytesIO(content_bytes), encoding=enc, engine='python', **kwargs)
            if len(df.columns) > 0:
                df.columns = df.columns.astype(str).str.replace('\ufeff', '').str.strip()
            print(f"[GAS] Success: Loaded {filename} ({enc}). Rows: {len(df)}")
            return df, "ok"
        except Exception:
            continue
    return pd.DataFrame(), "error"


def load_df_from_gas(filename, **kwargs):
    df, _ = load_df_from_gas_with_status(filename, **kwargs)
    return df


def save_df_to_gas(filename, df):
    print(f"[GAS] Uploading {filename}...")
    try:
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        payload = {'filename': filename, 'content': csv_buffer.getvalue()}
        response = requests.post(GAS_WEB_APP_URL, json=payload, timeout=60)
        if response.status_code == 200:
            print(f"[GAS] Upload success: {response.text}")
            return True
        else:
            print(f"[GAS] Upload failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"[GAS] Upload error: {e}")
        return False


# ==========================================
# メイン処理
# ==========================================
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
current_date_str = now.strftime("%Y/%m/%d")
current_datetime_str = now.strftime("%Y/%m/%d %H:%M")

room_map = {
    11000: "ゆーふうりん部屋", 11001: "ゆーふうりん部屋", 11002: "ゆーふうりん部屋",
    11003: "ゆーふうりん部屋", 11004: "ゆーふうりん部屋", 11005: "ゆーふうりん部屋",
    11006: "ゆーふうりん部屋", 11007: "ゆーふうりん部屋", 11008: "ゆーふうりん部屋",
    11009: "ゆーふうりん部屋",
    11012: "加古部屋",
    11021: "成田部屋", 11022: "成田部屋",
    11028: "タマ部屋",
    11058: "すみた部屋", 11059: "つぼはち部屋", 11060: "れん部屋",
    11063: "なぎ部屋", 11064: "naoo部屋", 11066: "芝ちゃん部屋",
    11067: "crom部屋", 11068: "けんしん部屋", 11069: "けんちぃ部屋",
    11070: "黒河部屋", 11071: "黒河部屋", 11074: "tukinowa部屋",
    11077: "v3部屋", 11078: "のんでるん部屋", 11079: "まどか部屋",
    11087: "MiO部屋", 11088: "ほっしー部屋", 11091: "千秋部屋",
    11092: "ヒロ部屋", 11101: "えみち部屋", 11102: "るえ部屋",
    11103: "ながし部屋", 11104: "MrN部屋", 11105: "ヤマテル部屋",
    11106: "冨塚部屋", 11107: "ブルーベリー部屋", 11110: "加古部屋",
    11111: "ヒロ部屋"
}


def normalize_text(text):
    if not isinstance(text, str):
        return str(text)
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
    if not isinstance(text, str):
        return str(text)
    text = unicodedata.normalize('NFKC', text)
    text = re.sub(r'\.[a-zA-Z0-9]{3,4}$', '', text)
    text = re.sub(r'(key|KEY)?\s*[\+\-]\s*[0-9]+', ' ', text)
    text = re.sub(r'原キー', ' ', text)
    text = re.sub(r'(キー)?変更[:：]?', ' ', text)
    text = re.sub(r'[~〜～\-_=,.]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text.upper()


def check_match(target_text, source_series):
    if not target_text:
        return pd.Series([False] * len(source_series))
    safe_target = re.escape(target_text)
    if re.match(r'^[A-Z0-9\s]+$', target_text):
        pattern = r'(?:^|[^A-Z0-9])' + safe_target + r'(?:[^A-Z0-9]|$)'
        return source_series.str.contains(pattern, regex=True, case=False, na=False)
    else:
        return source_series.str.contains(safe_target, case=False, na=False)


# --- 履歴ローテーション設定 ---
HISTORY_MAX_ROWS = 9500
HISTORY_ARCHIVE_MISS_LIMIT = 3
IGNORE_FETCH_FAILURES = True
ROOM_FETCH_TIMEOUT = 6
ROOM_FETCH_WORKERS = 16
HISTORY_DEDUP_COLS = ['取得日', '部屋主', '順番', '曲名（ファイル名）', '歌った人']


def get_history_filename(num):
    return "history.csv" if num == 1 else f"history_{num}.csv"


def sort_history_df(df):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.fillna("")
    df = df.copy().fillna("")
    if '順番' in df.columns:
        df['順番'] = pd.to_numeric(df['順番'], errors='coerce')
    if '取得日' in df.columns:
        df['_temp_date'] = pd.to_datetime(df['取得日'], errors='coerce')
        sort_cols = ['_temp_date']
        ascending = [False]
        if '順番' in df.columns:
            sort_cols.append('順番')
            ascending.append(False)
        df = df.sort_values(by=sort_cols, ascending=ascending, kind='mergesort')
        df = df.drop(columns=['_temp_date'])
    cols = list(df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        df = df[cols]
    return format_history_order_column(df.fillna(""))


def format_history_order_column(df):
    if df is None or df.empty or '順番' not in df.columns:
        return pd.DataFrame() if df is None else df

    def _format_order(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        if s == "":
            return ""
        n = pd.to_numeric(pd.Series([s]), errors='coerce').iloc[0]
        if pd.isna(n):
            return s
        if float(n).is_integer():
            return str(int(n))
        return f"{n:g}"

    df = df.copy()
    df['順番'] = df['順番'].apply(_format_order)
    return df


def cleanup_history_df(df):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.fillna("")
    df = df.copy().fillna("")
    clean_check_cols = ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']
    for col in clean_check_cols:
        if col in df.columns:
            df = df[df[col].astype(str) != col]
    bad_col_patterns = [r'^Error:', r'Exception:\s*Service error:\s*Drive']
    bad_cols = []
    for c in df.columns:
        col = str(c).strip()
        for pat in bad_col_patterns:
            if re.search(pat, col, flags=re.IGNORECASE):
                bad_cols.append(c)
                break
    if bad_cols:
        df = df.drop(columns=bad_cols, errors='ignore')
    keep_cols = [c for c in EXPECTED_HISTORY_COLUMNS if c in df.columns]
    other_cols = [c for c in df.columns if c not in keep_cols]
    if keep_cols:
        df = df[keep_cols + other_cols]
    return sort_history_df(df)


def make_dedup_key(df, for_history_compare=False):
    if df is None or df.empty:
        return pd.Series([], dtype=str)
    work = df.copy().fillna("")
    dedup_cols = HISTORY_DEDUP_COLS.copy()
    if for_history_compare and '取得日' in dedup_cols:
        dedup_cols.remove('取得日')
    for col in dedup_cols:
        if col not in work.columns:
            work[col] = ""
    return work[dedup_cols].astype(str).agg("\u241f".join, axis=1)


def save_df_to_gas_checked(filename, df, min_existing_rows=0, allow_empty=False):
    if df is None:
        return False
    df = cleanup_history_df(df)
    if df.empty and not allow_empty:
        return False
    if len(df) < min_existing_rows:
        print(f"[Guard] {filename}: 行数減少のため拒否 ({min_existing_rows} -> {len(df)})")
        return False
    if not save_df_to_gas(filename, df):
        return False
    verify_df, verify_status = load_df_from_gas_with_status(filename)
    if verify_status != "ok":
        return False
    verify_df = cleanup_history_df(verify_df)
    if len(verify_df) < len(df):
        return False
    return True


def load_all_history_files():
    histories = []
    loaded_files = []
    missing_count = 0
    num = 1
    while missing_count < HISTORY_ARCHIVE_MISS_LIMIT:
        filename = get_history_filename(num)
        df, status = load_df_from_gas_with_status(filename)
        if status == "ok":
            df = cleanup_history_df(df)
            histories.append({"num": num, "filename": filename, "df": df})
            loaded_files.append(filename)
            missing_count = 0
        elif status in ("not_found", "empty"):
            missing_count += 1
        else:
            print(f"[STOP] {filename} 読込失敗。履歴更新を禁止。")
            return histories, loaded_files, False
        num += 1
    return histories, loaded_files, True


def fetch_room_df(port):
    url = f"http://ykr.moe:{port}/simplelist.php"
    response = requests.get(url, timeout=ROOM_FETCH_TIMEOUT)
    response.raise_for_status()
    dfs = pd.read_html(io.BytesIO(response.content))
    if not dfs:
        raise ValueError("テーブルなし")
    df = dfs[0].fillna("")
    if df.empty:
        raise ValueError("空テーブル")
    required_cols = ['順番', '曲名（ファイル名）']
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        raise ValueError(f"必要カラム不足: {missing_cols}")
    df = df.replace(r'\s*詳細を見る ▼', '', regex=True)
    df['部屋主'] = room_map[port]
    df['取得日'] = current_date_str
    return df


# --- 既存履歴の読み込み ---
history_records, loaded_history_files, history_load_ok = load_all_history_files()
if loaded_history_files:
    print(f"���歴ファイル読込: {', '.join(loaded_history_files)}")
else:
    print("履歴ファイルなし。")

history_dfs = [h['df'] for h in history_records]
full_history_before_update = cleanup_history_df(pd.concat(history_dfs, ignore_index=True)) if history_dfs else pd.DataFrame()

# --- 新データ取得 ---
target_ports = list(room_map.keys())
new_data_frames = []
failed_ports = []
fetched_ports = []

if not history_load_ok:
    print("[Guard] 履歴読込不完全のため、保存禁止モード。")

print("データ取得中...")
max_workers = min(ROOM_FETCH_WORKERS, max(1, len(target_ports)))
with ThreadPoolExecutor(max_workers=max_workers) as executor:
    future_to_port = {executor.submit(fetch_room_df, port): port for port in target_ports}
    for future in as_completed(future_to_port):
        port = future_to_port[future]
        room_name = room_map.get(port, "不明")
        try:
            df = future.result()
            new_data_frames.append(df)
            fetched_ports.append(port)
            print(f"[Fetch] OK {port} ({room_name}) rows={len(df)}")
        except Exception as e:
            failed_ports.append((port, str(e)))
            print(f"[Fetch] SKIP {port} ({room_name}) {e}")

print(f"[Fetch] 成功: {len(fetched_ports)}/{len(target_ports)}")

if not new_data_frames:
    print("[Fetch] 取得成功データなし。")
    full_df = full_history_before_update
else:
    new_df = cleanup_history_df(pd.concat(new_data_frames, ignore_index=True))
    dedup_cols = [c for c in HISTORY_DEDUP_COLS if c in new_df.columns]
    if dedup_cols:
        before_rows = len(new_df)
        new_df = new_df.drop_duplicates(subset=dedup_cols, keep='last')
        new_df = cleanup_history_df(new_df)
        print(f"[Dedup] {before_rows} -> {len(new_df)}")

    if not history_load_ok:
        if full_history_before_update.empty:
            full_df = new_df
        else:
            full_df = cleanup_history_df(pd.concat([full_history_before_update, new_df], ignore_index=True))
    else:
        existing_keys = set(make_dedup_key(full_history_before_update, for_history_compare=True).tolist()) if not full_history_before_update.empty else set()
        new_keys = make_dedup_key(new_df, for_history_compare=True)
        new_unique_df = new_df[~new_keys.isin(existing_keys)].copy()
        new_unique_df = cleanup_history_df(new_unique_df)

        if new_unique_df.empty:
            print("追加なし")
            full_df = full_history_before_update
        else:
            print(f"追加: {len(new_unique_df)} 行")
            if history_records:
                active_num = history_records[-1]['num']
                active_df = history_records[-1]['df']
                if len(active_df) >= HISTORY_MAX_ROWS:
                    active_num += 1
                    active_df = pd.DataFrame()
            else:
                active_num = 1
                active_df = pd.DataFrame()

            remaining_df = new_unique_df.reset_index(drop=True)
            saved_parts = {}
            while not remaining_df.empty:
                active_filename = get_history_filename(active_num)
                current_df = cleanup_history_df(active_df)
                remaining_capacity = HISTORY_MAX_ROWS - len(current_df)
                if remaining_capacity <= 0:
                    active_num += 1
                    active_df = pd.DataFrame()
                    continue
                append_part = remaining_df.iloc[:remaining_capacity].copy()
                next_df = cleanup_history_df(pd.concat([current_df, append_part], ignore_index=True))
                if save_df_to_gas_checked(active_filename, next_df, min_existing_rows=len(current_df)):
                    print(f"[History] {active_filename} +{len(append_part)} 行")
                    saved_parts[active_num] = next_df
                    remaining_df = remaining_df.iloc[remaining_capacity:].reset_index(drop=True)
                    if not remaining_df.empty:
                        active_num += 1
                        active_df = pd.DataFrame()
                    else:
                        active_df = next_df
                else:
                    print(f"[STOP] {active_filename} 保存失敗")
                    break

            merged_history_by_num = {h['num']: h['df'] for h in history_records}
            merged_history_by_num.update(saved_parts)
            full_df = cleanup_history_df(pd.concat([merged_history_by_num[n] for n in sorted(merged_history_by_num)], ignore_index=True)) if merged_history_by_num else pd.DataFrame()

if full_df is None or full_df.empty:
    full_df = pd.DataFrame()
else:
    full_df = cleanup_history_df(full_df)

print(f"全履歴データ合計: {len(full_df)} 行")

# ==========================================
# ★ オフラインリスト読込
# ==========================================
ALLOWED_CATEGORIES = ["2026年春アニメ", "2026年冬アニメ", "2025年秋アニメ"]
cool_file = "cool_analysis.csv"

print("オフラインリスト読込中...")
offline_targets = []
for filename in OFFLINE_FILES:
    odf = load_df_from_github(filename)
    if not odf.empty and '曲名' in odf.columns:
        targets = [normalize_offline_text(str(x)) for x in odf.fillna("")['曲名'].tolist()]
        offline_targets.extend(targets)
        print(f"  -> {filename}: {len(targets)} 件")
print(f"オフラインリスト合計: {len(offline_targets)}")


def is_created(song, anime):
    """曲名(+作品名)がオフラインリストに含まれていれば作成済み。"""
    sn = normalize_text(song)
    rn = normalize_offline_text(song)
    an = normalize_text(anime)
    if not sn and not rn:
        return False
    for s in offline_targets:
        if (sn and sn in s) or (rn and rn in s):
            if (not an) or (an in s):
                return True
    return False


# ==========================================
# ★ 集計処理
# ==========================================
raw_df = load_df_from_gas(cool_file, header=None)

embedded_categories = {cat: [] for cat in ALLOWED_CATEGORIES}
created_lists = {cat: [] for cat in ALLOWED_CATEGORIES}
uncreated_lists = {cat: [] for cat in ALLOWED_CATEGORIES}
ranking_base = []
trending_items = []

if not raw_df.empty:
    try:
        raw_df = raw_df.fillna("").drop_duplicates(keep='last')
        analysis_source_df = full_df.copy()
        analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
        analysis_source_df = analysis_source_df.dropna(subset=['dt_obj'])
        analysis_source_df['norm_filename'] = analysis_source_df['曲名（ファイル名）'].apply(normalize_text)

        def get_rescued_workname(row):
            raw_work = str(row['作品名']) if pd.notna(row['作品名']) else ""
            raw_song = str(row['曲名（ファイル名）']) if pd.notna(row['曲名（ファイル名）']) else ""
            if raw_work.strip() in ["-", "−", "", "nan"]:
                m = re.search(r'【(.*?)】', raw_song)
                if m:
                    return normalize_text(m.group(1))
            return normalize_text(raw_work)

        if '作品名' in analysis_source_df.columns:
            analysis_source_df['norm_workname'] = analysis_source_df.apply(get_rescued_workname, axis=1)
        else:
            analysis_source_df['norm_workname'] = ""

        exclude_keywords = ['test', 'テスト', 'システム', 'admin', 'System']
        full_history = analysis_source_df[
            ~analysis_source_df['歌った人'].astype(str).apply(lambda x: any(k in x for k in exclude_keywords))
        ].sort_values('dt_obj')

        start_date = pd.to_datetime("2026/01/01")
        end_date = pd.to_datetime("2026/06/30")
        target_history = full_history[
            (full_history['dt_obj'] >= start_date) & (full_history['dt_obj'] <= end_date)
        ]

        # cool_analysis.csv パース
        categorized_data = {}
        current_category = None
        for _, row in raw_df.iterrows():
            if not any(str(x).strip() for x in row):
                continue
            col0 = str(row[0]).strip()
            is_cat_line = any(cat in col0 for cat in ALLOWED_CATEGORIES) and "作品名" not in col0
            if is_cat_line:
                current_category = col0
                categorized_data.setdefault(current_category, [])
                continue
            if "作品名" in col0 or current_category is None:
                continue
            anime = str(row[0]).strip() if len(row) > 0 else ""
            type_ = str(row[1]).strip() if len(row) > 1 else ""
            artist = str(row[2]).strip() if len(row) > 2 else ""
            song = str(row[3]).strip() if len(row) > 3 else ""
            if anime or song:
                categorized_data[current_category].append({
                    "anime": anime, "type": type_, "artist": artist, "song": song
                })

        # 各曲の歌唱集計
        for category, items in categorized_data.items():
            if category not in embedded_categories:
                continue
            for item in sorted(items, key=lambda x: (x['anime'], x['song'])):
                target_song_norm = normalize_text(item['song'])
                target_anime_norm = normalize_text(item['anime'])
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

                matched_data = target_history[final_mask]
                count = int(len(matched_data))
                users = int(matched_data['歌った人'].nunique()) if count > 0 else 0
                created = is_created(item['song'], item['anime'])

                rec = {
                    "anime": item['anime'], "type": item['type'], "artist": item['artist'],
                    "song": item['song'], "count": count, "users": users,
                    "created": created, "category": category
                }
                embedded_categories[category].append(rec)
                ranking_base.append(rec)
                (created_lists[category] if created else uncreated_lists[category]).append(rec)

        # 急上昇スコア（直近7日 vs その前2週平均）
        if not target_history.empty:
            max_dt = target_history['dt_obj'].max()
            recent_start = max_dt - pd.Timedelta(days=6)
            base_start = recent_start - pd.Timedelta(days=14)
            base_end = recent_start - pd.Timedelta(days=1)
            for category, items in embedded_categories.items():
                for it in items:
                    song_norm = normalize_text(it['song'])
                    anime_norm = normalize_text(it['anime'])
                    song_match = check_match(song_norm, target_history['norm_filename'])
                    anime_match = (
                        target_history['norm_filename'].str.contains(re.escape(anime_norm), case=False, na=False) |
                        target_history['norm_workname'].str.contains(re.escape(anime_norm), case=False, na=False)
                    )
                    if song_norm and anime_norm:
                        mask = song_match & anime_match
                    elif song_norm:
                        mask = song_match
                    else:
                        mask = anime_match
                    matched = target_history[mask]
                    recent_df = matched[(matched['dt_obj'] >= recent_start) & (matched['dt_obj'] <= max_dt)]
                    base_df = matched[(matched['dt_obj'] >= base_start) & (matched['dt_obj'] <= base_end)]
                    recent = int(len(recent_df))
                    if recent < 3:
                        continue
                    baseline = float(len(base_df) / 2.0)
                    score = (recent - baseline) / max(baseline, 1.0)
                    if score <= 0:
                        continue
                    trending_items.append({
                        "anime": it['anime'], "song": it['song'], "artist": it['artist'],
                        "type": it['type'], "category": category,
                        "recent": recent, "baseline": round(baseline, 2),
                        "score": round(score, 4),
                        "users7d": int(recent_df['歌った人'].nunique()) if not recent_df.empty else 0,
                        "isNew": baseline == 0
                    })
        print("集計完了")
    except Exception as e:
        print(f"集計エラー: {e}")
        import traceback
        traceback.print_exc()
else:
    print("cool_analysis.csv 読み込み失敗。")

trending_items.sort(key=lambda x: (x['score'], x['recent'], x['users7d']), reverse=True)
trending_items = trending_items[:30]

rankings_count = {cat: [] for cat in ALLOWED_CATEGORIES}
rankings_users = {cat: [] for cat in ALLOWED_CATEGORIES}
for cat in ALLOWED_CATEGORIES:
    items = [x for x in ranking_base if x['category'] == cat and x['count'] > 0]
    rankings_count[cat] = sorted(items, key=lambda x: (x['count'], x['users']), reverse=True)
    rankings_users[cat] = sorted(items, key=lambda x: (x['users'], x['count']), reverse=True)

# ==========================================
# セットリスト用データ整形（取得日DESC -> 順番DESC）
# ==========================================
setlist_records = []
if not full_df.empty:
    sl_df = full_df.drop(columns=['コメント'], errors='ignore').fillna("").copy()
    sl_df['_dt'] = pd.to_datetime(sl_df['取得日'], errors='coerce')
    sl_df['_ord'] = pd.to_numeric(sl_df['順番'], errors='coerce')
    sl_df = sl_df.sort_values(by=['_dt', '_ord'], ascending=[False, False], kind='mergesort')

    for _, row in sl_df.iterrows():
        song = str(row.get('曲名（ファイル名）', ''))
        work = str(row.get('作品名', ''))
        rec = {
            "room": str(row.get('部屋主', '')),
            "song": song,
            "work": work,
            "artist": str(row.get('歌手名', '')),
            "singer": str(row.get('歌った人', '')),
            "order": str(row.get('順番', '')),
            "fetchedAt": str(row.get('取得日', '')),
            "created": is_created(song, work),
        }
        rec['search'] = normalize_text(" ".join([
            rec[k] for k in ["room", "song", "work", "artist", "singer", "order", "fetchedAt"]
        ]))
        setlist_records.append(rec)

embedded = {
    "updatedAt": current_datetime_str,
    "setlist": setlist_records,
    "categories": embedded_categories,
    "createdLists": created_lists,
    "uncreatedLists": uncreated_lists,
    "rankings": {"count": rankings_count, "users": rankings_users},
    "trending": trending_items,
    "config": {"categories": ALLOWED_CATEGORIES, "defaultPort": 11059, "breakpoint": 900}
}
app_json = json.dumps(embedded, ensure_ascii=False)

# ==========================================
# HTML出力
# ==========================================
html_content = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Karaoke Dashboard</title>
<style>
:root{
  --bg:#ffffff;--bg-soft:#f7f8fa;
  --line:#e3e6ec;--line-strong:#cfd4dc;
  --text:#1f2430;--text-sub:#5b6472;--text-mute:#8a93a1;
  --accent:#1e3a8a;--accent-soft:#eaf0ff;
  --warn:#b45309;--ok:#15803d;
  --gold:#c8a44b;--silver:#9aa1ab;--bronze:#a06b3e;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0;background:var(--bg);color:var(--text);
  font-family:-apple-system,"Segoe UI","Hiragino Sans","Yu Gothic UI",system-ui,sans-serif;
  font-size:13px;line-height:1.5}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}

/* ===== Top Bar ===== */
.top{position:sticky;top:0;z-index:20;background:#fff;border-bottom:1px solid var(--line)}
.top-row{display:flex;align-items:center;gap:12px;padding:8px 12px;flex-wrap:wrap}
.brand{font-weight:700;font-size:15px;letter-spacing:.02em}
.update{color:var(--text-mute);font-size:11.5px}
.spacer{flex:1}
.cfg{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--text-sub);flex-wrap:wrap}
.cfg input,.cfg select{height:28px;border:1px solid var(--line-strong);border-radius:4px;
  padding:0 6px;background:#fff;font-size:12px;color:var(--text)}
.cfg input{width:80px}
.tabs{display:flex;border-top:1px solid var(--line);overflow-x:auto}
.tab-btn{flex-shrink:0;border:0;background:transparent;padding:10px 14px;
  font-size:13px;font-weight:600;color:var(--text-sub);cursor:pointer;
  border-bottom:2px solid transparent;height:38px}
.tab-btn:hover{color:var(--text)}
.tab-btn.active{color:var(--accent);border-bottom-color:var(--accent)}

/* ===== Tab content ===== */
.tab{display:none}
.tab.active{display:block}

/* ===== Toolbar ===== */
.toolbar{display:flex;align-items:center;gap:8px;padding:6px 12px;
  background:var(--bg-soft);border-bottom:1px solid var(--line);
  flex-wrap:wrap}
.toolbar input[type="search"]{
  height:30px;border:1px solid var(--line-strong);border-radius:4px;
  padding:0 8px;font-size:13px;min-width:180px;flex:1;max-width:320px}
.toolbar select{height:30px;border:1px solid var(--line-strong);border-radius:4px;
  padding:0 6px;font-size:12px;background:#fff;color:var(--text)}
.toolbar .count{margin-left:auto;color:var(--text-mute);font-size:11.5px;white-space:nowrap}

.btn{height:30px;padding:0 10px;border:1px solid var(--line-strong);
  background:#fff;color:var(--text);font-size:12px;border-radius:4px;cursor:pointer}
.btn:hover{background:var(--bg-soft)}
.btn.primary{background:var(--accent);color:#fff;border-color:var(--accent)}
.btn.primary:hover{background:#16306b}

/* ===== Chips ===== */
.chips{display:flex;gap:4px;flex-wrap:wrap}
.chip{height:26px;padding:0 10px;border:1px solid var(--line-strong);background:#fff;
  color:var(--text-sub);font-size:12px;cursor:pointer;border-radius:2px;line-height:24px}
.chip:hover{background:var(--bg-soft)}
.chip.active{color:var(--accent);border-color:var(--accent);background:var(--accent-soft);font-weight:600}

/* ===== Tags（作成済み/未作成） ===== */
.tag{display:inline-block;font-size:11px;font-weight:700;letter-spacing:.02em;
  margin-right:4px;padding:0 2px;line-height:1.4}
.tag-ok{color:var(--ok)}
.tag-no{color:var(--warn)}

/* ===== Setlist (PC: 1行高密度) ===== */
.setlist-list{height:calc(100vh - 160px);overflow:auto;position:relative;background:#fff;
  contain:strict;will-change:transform}
.setlist-spacer{width:1px;opacity:0;pointer-events:none}
.setlist-items{position:absolute;top:0;left:0;right:0}
.sl-row{display:grid;
  grid-template-columns:44px 1fr 90px;
  grid-template-areas:"ord title date" "ord sub room";
  column-gap:10px;row-gap:1px;
  padding:4px 12px;border-bottom:1px solid var(--line);
  height:34px;align-items:center;background:#fff}
.sl-row:hover{background:var(--bg-soft)}
.sl-ord{grid-area:ord;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;
  font-weight:700;text-align:right;color:var(--text);font-size:13px}
.sl-title{grid-area:title;font-weight:700;font-size:15px;line-height:1.2;
  color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sl-title a{color:var(--text)}
.sl-title a:hover{color:var(--accent)}
.sl-sub{grid-area:sub;color:var(--text-sub);font-size:12px;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sl-sub .sep{color:var(--text-mute);margin:0 6px}
.sl-date{grid-area:date;text-align:right;font-size:11.5px;color:var(--text-mute);
  font-family:ui-monospace,monospace;line-height:1.2}
.sl-room{grid-area:room;text-align:right;font-size:11.5px;color:var(--text-sub);
  line-height:1.2;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
  padding-left:6px;border-left:3px solid var(--accent)}
.sl-card{display:none}

@media (max-width:899px){
  .sl-row{display:none}
  .sl-card{display:grid;
    grid-template-columns:36px 1fr;column-gap:8px;row-gap:1px;
    padding:8px 10px;margin:0 0 6px 0;
    border-bottom:1px solid var(--line);background:#fff;
    height:88px;overflow:hidden}
  .sl-card .ord{grid-row:1 / span 4;align-self:start;
    font-family:ui-monospace,monospace;font-weight:700;font-size:14px;
    text-align:right;color:var(--text-sub);padding-top:2px}
  .sl-card .title{font-weight:700;font-size:15px;line-height:1.2;color:var(--text);
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .sl-card .work{font-size:12.5px;color:var(--text-sub);line-height:1.3;
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .sl-card .singer{font-size:12px;color:var(--text-mute);line-height:1.3;
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .sl-card .meta{font-size:11px;color:var(--text-mute);text-align:right;
    border-top:1px dashed var(--line);padding-top:2px;margin-top:2px;
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .setlist-list{height:calc(100vh - 200px)}
}

/* ===== Generic table ===== */
.tbl{width:100%;border-collapse:collapse;background:#fff;font-size:13px}
.tbl thead th{position:sticky;top:0;background:var(--bg-soft);
  border-bottom:1px solid var(--line-strong);
  font-weight:600;color:var(--text-sub);font-size:12px;
  padding:6px 8px;text-align:left;white-space:nowrap}
.tbl tbody td{padding:5px 8px;border-bottom:1px solid var(--line);vertical-align:middle}
.tbl tbody tr:hover{background:var(--bg-soft)}
.tbl .num{text-align:right;font-family:ui-monospace,SFMono-Regular,Menlo,monospace}
.tbl .num.zero{color:var(--text-mute)}
.tbl .num.low{color:var(--text-sub)}
.tbl .num.mid{color:var(--text)}
.tbl .num.high{color:var(--accent);font-weight:700}
.tbl .ttl{font-weight:600}
.tbl .ttl a{color:var(--text)}
.tbl .ttl a:hover{color:var(--accent)}
.tbl .small{color:var(--text-sub);font-size:12px;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:240px}
.tbl .ord{font-family:ui-monospace,monospace;color:var(--text-sub)}
.tbl .clickable{cursor:pointer}

/* category section */
.cat-section{border-top:1px solid var(--line)}
.cat-section > summary{
  list-style:none;cursor:pointer;
  padding:8px 12px;background:var(--bg-soft);
  font-weight:700;font-size:13px;color:var(--text);
  border-bottom:1px solid var(--line);
  display:flex;align-items:center;gap:8px}
.cat-section > summary::-webkit-details-marker{display:none}
.cat-section > summary::before{content:"▸";color:var(--text-mute);font-size:11px;
  display:inline-block;transition:transform .1s;width:12px}
.cat-section[open] > summary::before{transform:rotate(90deg)}
.cat-section .meta-info{margin-left:auto;color:var(--text-mute);font-weight:400;font-size:11.5px}

/* Ranking row */
.rk-row td:first-child{padding-left:8px}
.rk-row.top1 td:first-child{border-left:3px solid var(--gold)}
.rk-row.top2 td:first-child{border-left:3px solid var(--silver)}
.rk-row.top3 td:first-child{border-left:3px solid var(--bronze)}

/* Trending */
.delta{color:var(--ok);font-weight:700;text-align:right;
  font-family:ui-monospace,monospace}
.tag-new{color:var(--ok);font-weight:700;font-size:11px;letter-spacing:.04em}

/* Mobile compact for tables */
@media (max-width:899px){
  .tbl thead .hide-sp{display:none}
  .tbl tbody td.hide-sp{display:none}
  .tbl tbody td{padding:5px 6px}
  .tbl thead th{padding:5px 6px;font-size:11.5px}
  .tbl .small{max-width:140px}
}

.empty{padding:40px 20px;text-align:center;color:var(--text-mute)}
.cat-body{padding:0}
.more-toggle{padding:6px 8px;color:var(--text-sub);font-size:12px;cursor:pointer;
  background:var(--bg-soft);border-top:1px solid var(--line)}
.more-toggle:hover{color:var(--accent)}
</style>
</head>
<body>

<div class="top">
  <div class="top-row">
    <span class="brand">Karaoke Dashboard</span>
    <span class="update">__UPDATED__ 更新</span>
    <span class="spacer"></span>
    <span class="cfg">
      <label>Port <input id="exportPort" type="number" value="11059"></label>
      <label>Link
        <select id="exportLinkType">
          <option value="eve">Everything</option>
          <option value="ykr">ゆかりすたー</option>
        </select>
      </label>
    </span>
  </div>
  <div class="tabs">
    <button class="tab-btn active" data-tab="setlist">セットリスト</button>
    <button class="tab-btn" data-tab="analysis">クール集計</button>
    <button class="tab-btn" data-tab="ranking_count">歌唱数ランキング</button>
    <button class="tab-btn" data-tab="ranking_users">歌唱人数ランキング</button>
    <button class="tab-btn" data-tab="trending">急上昇</button>
  </div>
</div>

<!-- Setlist -->
<section id="setlist" class="tab active">
  <div class="toolbar">
    <input id="searchInput" type="search" placeholder="検索 (例: 曲名 歌手)">
    <div class="chips" id="setlistSort"></div>
    <button id="saveSetlist" class="btn primary">HTML保存</button>
    <span class="count" id="setlistCount"></span>
  </div>
  <div class="toolbar" style="border-top:0">
    <div class="chips" id="roomFilters"></div>
  </div>
  <div id="setlistList" class="setlist-list">
    <div id="setlistSpacer" class="setlist-spacer"></div>
    <div id="setlistItems" class="setlist-items"></div>
  </div>
</section>

<!-- Analysis -->
<section id="analysis" class="tab">
  <div class="toolbar">
    <select id="analysisCategory"></select>
    <select id="analysisState">
      <option value="all">すべて</option>
      <option value="created">作成済み</option>
      <option value="uncreated">未作成</option>
      <option value="has">歌唱あり</option>
      <option value="none">未歌唱</option>
    </select>
    <select id="analysisSort">
      <option value="anime">作品名順</option>
      <option value="count">歌唱数↓</option>
      <option value="users">歌唱人数↓</option>
    </select>
    <button id="saveAnalysis" class="btn primary">HTML保存</button>
    <button id="saveCreated" class="btn">作成リスト保存</button>
    <button id="saveUncreated" class="btn">未作成リスト保存</button>
    <span class="count">2026/01/01-06/30</span>
  </div>
  <div id="analysisBody"></div>
</section>

<!-- Ranking: count -->
<section id="ranking_count" class="tab">
  <div class="toolbar">
    <select id="rankingCountCategory"></select>
    <button id="saveRankingCount" class="btn primary">HTML保存</button>
    <span class="count">2026/01/01-06/30</span>
  </div>
  <div id="rankingCountBody"></div>
</section>

<!-- Ranking: users -->
<section id="ranking_users" class="tab">
  <div class="toolbar">
    <select id="rankingUsersCategory"></select>
    <button id="saveRankingUsers" class="btn primary">HTML保存</button>
    <span class="count">2026/01/01-06/30</span>
  </div>
  <div id="rankingUsersBody"></div>
</section>

<!-- Trending -->
<section id="trending" class="tab">
  <div class="toolbar">
    <select id="trendingCategory"></select>
    <label style="font-size:12px;color:var(--text-sub)">
      <input id="trendingNewOnly" type="checkbox"> 初登場のみ
    </label>
    <button id="saveTrending" class="btn primary">HTML保存</button>
    <span class="count">直近7日 vs 前2週平均</span>
  </div>
  <div id="trendingBody"></div>
</section>

<script id="app-data" type="application/json">__APP_JSON__</script>
<script>
const APP=JSON.parse(document.getElementById('app-data').textContent);
const BP=APP.config.breakpoint||900;
const isMobile=()=>window.matchMedia('(max-width:'+(BP-1)+'px)').matches;

function esc(s){return (s==null?'':String(s)).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function host(){return 'http://ykr.moe:'+(document.getElementById('exportPort').value||APP.config.defaultPort);}
function searchPath(){return document.getElementById('exportLinkType').value==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword=';}
function ykr(word){return host()+'/'+searchPath()+encodeURIComponent(word);}
function numCls(v){if(!v)return 'zero';if(v<3)return 'low';if(v<10)return 'mid';return 'high';}
function tagHtml(c){return c?'<span class="tag tag-ok">済</span>':'<span class="tag tag-no">未</span>';}

/* ===== Tab switching ===== */
const tabBtns=[...document.querySelectorAll('.tab-btn')];
tabBtns.forEach(b=>b.onclick=()=>{
  tabBtns.forEach(x=>x.classList.remove('active'));
  b.classList.add('active');
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById(b.dataset.tab).classList.add('active');
  if(b.dataset.tab==='setlist')requestAnimationFrame(renderWindow);
});

/* ===== Setlist ===== */
const SRC=APP.setlist;
const sortOpts=[['date_desc','取得日↓'],['date_asc','取得日↑'],['order_desc','順番↓'],['song','曲名']];
let activeSort='date_desc';
const roomSet=new Set();
let filteredIdx=[];
let searchTimer=null;

const rooms=[...new Set(SRC.map(x=>x.room).filter(Boolean))].sort();

function renderSortChips(){
  document.getElementById('setlistSort').innerHTML=
    sortOpts.map(([k,l])=>'<button class="chip'+(k===activeSort?' active':'')+'" data-sort="'+k+'">'+l+'</button>').join('');
}
renderSortChips();
document.getElementById('setlistSort').onclick=e=>{
  const b=e.target.closest('[data-sort]');if(!b)return;
  activeSort=b.dataset.sort;renderSortChips();applySetlist();
};
document.getElementById('roomFilters').innerHTML=
  rooms.map(r=>'<button class="chip" data-room="'+esc(r)+'">'+esc(r)+'</button>').join('');
document.getElementById('roomFilters').onclick=e=>{
  const b=e.target.closest('[data-room]');if(!b)return;
  const r=b.dataset.room;
  if(roomSet.has(r)){roomSet.delete(r);b.classList.remove('active');}
  else{roomSet.add(r);b.classList.add('active');}
  applySetlist();
};

function applySetlist(){
  const kw=(document.getElementById('searchInput').value||'').trim().toUpperCase().split(/\s+/).filter(Boolean);
  const result=[];
  for(let i=0;i<SRC.length;i++){
    const it=SRC[i];
    if(roomSet.size && !roomSet.has(it.room))continue;
    let ok=true;
    for(let k=0;k<kw.length;k++){if(it.search.indexOf(kw[k])===-1){ok=false;break;}}
    if(ok)result.push(i);
  }
  result.sort((a,b)=>{
    const A=SRC[a],B=SRC[b];
    if(activeSort==='song')return (A.song||'').localeCompare(B.song||'','ja');
    if(activeSort==='date_asc'){
      const c=(A.fetchedAt||'').localeCompare(B.fetchedAt||'','ja');
      if(c!==0)return c;
      return (parseFloat(A.order)||0)-(parseFloat(B.order)||0);
    }
    if(activeSort==='order_desc'){return (parseFloat(B.order)||0)-(parseFloat(A.order)||0);}
    const c=(B.fetchedAt||'').localeCompare(A.fetchedAt||'','ja');
    if(c!==0)return c;
    return (parseFloat(B.order)||0)-(parseFloat(A.order)||0);
  });
  filteredIdx=result;
  document.getElementById('setlistCount').textContent='全'+SRC.length+'件 / 表示'+filteredIdx.length+'件';
  document.getElementById('setlistList').scrollTop=0;
  renderWindow();
}

const listEl=document.getElementById('setlistList');
listEl.addEventListener('scroll',()=>requestAnimationFrame(renderWindow));
window.addEventListener('resize',()=>requestAnimationFrame(renderWindow));

function rowH(){return isMobile()?94:34;}

function renderWindow(){
  const h=rowH();
  const top=listEl.scrollTop;
  const vh=listEl.clientHeight;
  const buf=isMobile()?4:8;
  const s=Math.max(0,Math.floor(top/h)-buf);
  const e=Math.min(filteredIdx.length,Math.ceil((top+vh)/h)+buf);
  document.getElementById('setlistSpacer').style.height=(filteredIdx.length*h)+'px';
  const wrap=document.getElementById('setlistItems');
  wrap.style.transform='translateY('+(s*h)+'px)';
  const html=[];
  const mobile=isMobile();
  for(let i=s;i<e;i++){
    const x=SRC[filteredIdx[i]];if(!x)continue;
    const q=(x.work?x.work+' ':'')+x.song;
    const t=tagHtml(x.created);
    if(mobile){
      html.push(
        '<article class="sl-card">'+
          '<div class="ord">'+esc(x.order||'-')+'</div>'+
          '<div class="title">'+t+'<a href="'+ykr(q)+'" target="_blank" rel="noopener">'+esc(x.song||'-')+'</a></div>'+
          '<div class="work">'+esc(x.work||'-')+(x.artist?' ／ '+esc(x.artist):'')+'</div>'+
          '<div class="singer">歌った人: '+esc(x.singer||'-')+'</div>'+
          '<div class="meta">'+esc(x.fetchedAt||'')+' ・ '+esc(x.room||'')+'</div>'+
        '</article>'
      );
    }else{
      html.push(
        '<div class="sl-row">'+
          '<div class="sl-ord">'+esc(x.order||'-')+'</div>'+
          '<div class="sl-title">'+t+'<a href="'+ykr(q)+'" target="_blank" rel="noopener">'+esc(x.song||'-')+'</a></div>'+
          '<div class="sl-date">'+esc(x.fetchedAt||'')+'</div>'+
          '<div class="sl-sub">'+esc(x.work||'-')+'<span class="sep">／</span>'+esc(x.artist||'-')+'<span class="sep">／</span>'+esc(x.singer||'-')+'</div>'+
          '<div class="sl-room">'+esc(x.room||'-')+'</div>'+
        '</div>'
      );
    }
  }
  wrap.innerHTML=html.join('');
}

document.getElementById('searchInput').addEventListener('input',()=>{
  clearTimeout(searchTimer);
  searchTimer=setTimeout(()=>requestAnimationFrame(applySetlist),150);
});
applySetlist();

/* ===== Category selectors ===== */
const cats=['ALL'].concat(APP.config.categories);
['analysisCategory','rankingCountCategory','rankingUsersCategory','trendingCategory'].forEach(id=>{
  document.getElementById(id).innerHTML=
    cats.map(c=>'<option value="'+esc(c)+'">'+(c==='ALL'?'すべて':esc(c))+'</option>').join('');
});

/* ===== Analysis ===== */
function renderAnalysis(){
  const cat=document.getElementById('analysisCategory').value;
  const st=document.getElementById('analysisState').value;
  const so=document.getElementById('analysisSort').value;
  let html='';
  for(const c of APP.config.categories){
    if(cat!=='ALL' && cat!==c)continue;
    let items=(APP.categories[c]||[]).slice();
    items=items.filter(x=>{
      if(st==='created')return x.created;
      if(st==='uncreated')return !x.created;
      if(st==='has')return x.count>0;
      if(st==='none')return x.count===0;
      return true;
    });
    items.sort((a,b)=>{
      if(so==='count')return b.count-a.count||b.users-a.users;
      if(so==='users')return b.users-a.users||b.count-a.count;
      return (a.anime||'').localeCompare(b.anime||'','ja')||(a.song||'').localeCompare(b.song||'','ja');
    });
    const total=items.length;
    const totalCnt=items.reduce((s,x)=>s+x.count,0);
    const totalUsr=items.reduce((s,x)=>s+x.users,0);
    html+='<details class="cat-section" open>'+
      '<summary>'+esc(c)+'<span class="meta-info">'+total+'曲 ／ 累計人数'+totalUsr+' ／ 累計歌唱'+totalCnt+'</span></summary>'+
      '<div class="cat-body"><table class="tbl"><thead><tr>'+
        '<th style="width:40px">作成</th>'+
        '<th>作品名</th>'+
        '<th class="hide-sp" style="width:60px">区分</th>'+
        '<th>曲名</th>'+
        '<th class="hide-sp">歌手</th>'+
        '<th class="num" style="width:50px">人数</th>'+
        '<th class="num" style="width:60px">歌唱</th>'+
      '</tr></thead><tbody>';
    if(items.length===0){
      html+='<tr><td colspan="7" class="empty">該当データなし</td></tr>';
    }else{
      for(const x of items){
        const q=(x.anime?x.anime+' ':'')+x.song;
        html+='<tr>'+
          '<td>'+tagHtml(x.created)+'</td>'+
          '<td class="small">'+esc(x.anime)+'</td>'+
          '<td class="hide-sp small">'+esc(x.type)+'</td>'+
          '<td class="ttl"><a href="'+ykr(q)+'" target="_blank" rel="noopener">'+esc(x.song)+'</a></td>'+
          '<td class="hide-sp small">'+esc(x.artist)+'</td>'+
          '<td class="num '+numCls(x.users)+'">'+x.users+'</td>'+
          '<td class="num '+numCls(x.count)+'">'+x.count+'</td>'+
        '</tr>';
      }
    }
    html+='</tbody></table></div></details>';
  }
  document.getElementById('analysisBody').innerHTML=html||'<div class="empty">データなし</div>';
}
['analysisCategory','analysisState','analysisSort'].forEach(id=>document.getElementById(id).onchange=renderAnalysis);
renderAnalysis();

/* ===== Ranking ===== */
function buildRankingRows(arr,mode,startIndex){
  let html='';
  let prev=null,rank=0;
  arr.forEach((x,i)=>{
    const v=mode==='count'?x.count:x.users;
    if(v!==prev)rank=startIndex+i+1;
    prev=v;
    const cls=rank===1?'top1':rank===2?'top2':rank===3?'top3':'';
    const q=(x.anime?x.anime+' ':'')+x.song;
    const usersHigh=mode==='users'?'high':numCls(x.users);
    const countHigh=mode==='count'?'high':numCls(x.count);
    html+='<tr class="rk-row '+cls+'">'+
      '<td class="ord">#'+rank+'</td>'+
      '<td>'+tagHtml(x.created)+'</td>'+
      '<td class="small">'+esc(x.anime)+'</td>'+
      '<td class="ttl"><a href="'+ykr(q)+'" target="_blank" rel="noopener">'+esc(x.song)+'</a></td>'+
      '<td class="hide-sp small">'+esc(x.artist)+'</td>'+
      '<td class="hide-sp small">'+esc(x.type)+'</td>'+
      '<td class="num '+usersHigh+'">'+x.users+'</td>'+
      '<td class="num '+countHigh+'">'+x.count+'</td>'+
    '</tr>';
  });
  return html;
}

function renderRanking(mode,bodyId,categoryId){
  const cat=document.getElementById(categoryId).value;
  let html='';
  for(const c of APP.config.categories){
    if(cat!=='ALL' && cat!==c)continue;
    const arr=(APP.rankings[mode]&&APP.rankings[mode][c])||[];
    html+='<details class="cat-section" open>'+
      '<summary>'+esc(c)+'<span class="meta-info">'+arr.length+'曲</span></summary>'+
      '<div class="cat-body"><table class="tbl"><thead><tr>'+
        '<th style="width:40px">順位</th>'+
        '<th style="width:36px">作成</th>'+
        '<th>作品名</th>'+
        '<th>曲名</th>'+
        '<th class="hide-sp">歌手</th>'+
        '<th class="hide-sp" style="width:50px">区分</th>'+
        '<th class="num" style="width:50px">人数</th>'+
        '<th class="num" style="width:60px">歌唱</th>'+
      '</tr></thead><tbody>';
    if(arr.length===0){
      html+='<tr><td colspan="8" class="empty">該当データなし</td></tr>';
    }else{
      const top=arr.slice(0,20);
      const rest=arr.slice(20);
      html+=buildRankingRows(top,mode,0);
      if(rest.length){
        html+='<tr><td colspan="8" style="padding:0">'+
          '<details><summary class="more-toggle">▸ もっと見る ('+rest.length+'件)</summary>'+
          '<table class="tbl" style="border:0"><tbody>'+buildRankingRows(rest,mode,top.length)+'</tbody></table>'+
          '</details></td></tr>';
      }
    }
    html+='</tbody></table></div></details>';
  }
  document.getElementById(bodyId).innerHTML=html||'<div class="empty">データなし</div>';
}
document.getElementById('rankingCountCategory').onchange=()=>renderRanking('count','rankingCountBody','rankingCountCategory');
document.getElementById('rankingUsersCategory').onchange=()=>renderRanking('users','rankingUsersBody','rankingUsersCategory');
renderRanking('count','rankingCountBody','rankingCountCategory');
renderRanking('users','rankingUsersBody','rankingUsersCategory');

/* ===== Trending ===== */
function renderTrending(){
  const cat=document.getElementById('trendingCategory').value;
  const onlyNew=document.getElementById('trendingNewOnly').checked;
  const arr=APP.trending.filter(x=>(cat==='ALL'||x.category===cat)&&(!onlyNew||x.isNew));
  let html='<table class="tbl"><thead><tr>'+
    '<th style="width:40px">順位</th>'+
    '<th>作品名</th>'+
    '<th>曲名</th>'+
    '<th class="hide-sp">歌手</th>'+
    '<th class="hide-sp" style="width:50px">区分</th>'+
    '<th class="num" style="width:60px">直近7日</th>'+
    '<th class="num hide-sp" style="width:50px">人</th>'+
    '<th class="num hide-sp" style="width:70px">前2週平均</th>'+
    '<th class="num" style="width:70px">増加率</th>'+
    '<th class="hide-sp" style="width:50px">NEW</th>'+
  '</tr></thead><tbody>';
  if(arr.length===0){
    html+='<tr><td colspan="10" class="empty">該当データなし</td></tr>';
  }else{
    arr.forEach((x,i)=>{
      const q=(x.anime?x.anime+' ':'')+x.song;
      html+='<tr>'+
        '<td class="ord">#'+(i+1)+'</td>'+
        '<td class="small">'+esc(x.anime)+'</td>'+
        '<td class="ttl"><a href="'+ykr(q)+'" target="_blank" rel="noopener">'+esc(x.song)+'</a></td>'+
        '<td class="hide-sp small">'+esc(x.artist)+'</td>'+
        '<td class="hide-sp small">'+esc(x.type)+'</td>'+
        '<td class="num mid">'+x.recent+'</td>'+
        '<td class="num hide-sp '+numCls(x.users7d)+'">'+x.users7d+'</td>'+
        '<td class="num hide-sp small">'+x.baseline+'</td>'+
        '<td class="delta">+'+Math.round(x.score*100)+'%</td>'+
        '<td class="hide-sp">'+(x.isNew?'<span class="tag-new">NEW</span>':'')+'</td>'+
      '</tr>';
    });
  }
  html+='</tbody></table>';
  document.getElementById('trendingBody').innerHTML=html;
}
['trendingCategory','trendingNewOnly'].forEach(id=>document.getElementById(id).onchange=renderTrending);
renderTrending();

/* ===== HTML Save ===== */
function pageStyle(){return document.querySelector('style').textContent;}
function makeHtml(title,content){
  const p=document.getElementById('exportPort').value||APP.config.defaultPort;
  const lt=document.getElementById('exportLinkType').value;
  const sp=lt==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword=';
  return '<!doctype html><html lang="ja"><head><meta charset="utf-8">'+
    '<meta name="viewport" content="width=device-width,initial-scale=1">'+
    '<title>'+esc(title)+'</title><style>'+pageStyle()+
    '\n.top,.tabs,.toolbar{display:none}body{padding:12px}'+
    '.setlist-list{height:auto;overflow:visible;position:static}'+
    '.setlist-spacer{display:none}.setlist-items{position:static;transform:none!important}'+
    '</style></head><body>'+
    '<h1 style="font-size:16px;margin:0 0 4px">'+esc(title)+'</h1>'+
    '<div style="color:var(--text-mute);font-size:11.5px;margin-bottom:10px">出力日: '+esc(APP.updatedAt)+'</div>'+
    content+
    '<script>(function(){var H="http://ykr.moe:'+p+'",S="'+sp+'";'+
    'document.querySelectorAll("a[data-q]").forEach(function(a){a.href=H+"/"+S+encodeURIComponent(a.dataset.q);a.target="_blank";a.rel="noopener";});'+
    '})();<\/script></body></html>';
}
function rewriteLinks(htmlStr){
  return htmlStr.replace(/<a href="http:\/\/ykr\.moe:\d+\/[^"]*?(?:searchword|anyword)=([^"]*)"([^>]*)>/g,
    function(_,q,rest){return '<a data-q="'+q+'"'+rest+'>';});
}
function saveFile(name,title,content){
  const html=makeHtml(title,rewriteLinks(content));
  const b=new Blob([html],{type:'text/html'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(b);a.download=name;a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}

/* セットリスト保存：表示中の全件（最大10000件まで）を静的に出力 */
document.getElementById('saveSetlist').onclick=()=>{
  const limit=Math.min(filteredIdx.length,10000);
  const rows=[];
  for(let i=0;i<limit;i++){
    const x=SRC[filteredIdx[i]];
    const q=(x.work?x.work+' ':'')+x.song;
    rows.push(
      '<div class="sl-row" style="display:grid;grid-template-columns:44px 1fr 90px;grid-template-areas:\'ord title date\' \'ord sub room\';column-gap:10px;padding:4px 12px;border-bottom:1px solid var(--line);height:auto">'+
        '<div class="sl-ord">'+esc(x.order||'-')+'</div>'+
        '<div class="sl-title">'+tagHtml(x.created)+'<a data-q="'+esc(q)+'">'+esc(x.song||'-')+'</a></div>'+
        '<div class="sl-date">'+esc(x.fetchedAt||'')+'</div>'+
        '<div class="sl-sub">'+esc(x.work||'-')+' ／ '+esc(x.artist||'-')+' ／ '+esc(x.singer||'-')+'</div>'+
        '<div class="sl-room">'+esc(x.room||'-')+'</div>'+
      '</div>'
    );
  }
  saveFile('setlist.html','セットリスト',rows.join(''));
};

document.getElementById('saveAnalysis').onclick=()=>saveFile('karaoke_analysis.html','クール集計',document.getElementById('analysisBody').innerHTML);
document.getElementById('saveRankingCount').onclick=()=>saveFile('karaoke_ranking_count.html','歌唱数ランキング',document.getElementById('rankingCountBody').innerHTML);
document.getElementById('saveRankingUsers').onclick=()=>saveFile('karaoke_ranking_users.html','歌唱人数ランキング',document.getElementById('rankingUsersBody').innerHTML);
document.getElementById('saveTrending').onclick=()=>saveFile('karaoke_trending.html','急上昇',document.getElementById('trendingBody').innerHTML);

document.getElementById('saveCreated').onclick=()=>{
  const cat=document.getElementById('analysisCategory').value;
  let html='';
  for(const [k,v] of Object.entries(APP.createdLists)){
    if(cat!=='ALL'&&cat!==k)continue;
    if(!v.length)continue;
    html+='<details class="cat-section" open><summary>'+esc(k)+'<span class="meta-info">'+v.length+'曲</span></summary>'+
      '<table class="tbl"><thead><tr><th>作品名</th><th class="hide-sp" style="width:60px">区分</th><th>曲名</th><th class="hide-sp">歌手</th></tr></thead><tbody>';
    for(const x of v){
      const q=(x.anime?x.anime+' ':'')+x.song;
      html+='<tr><td class="small">'+esc(x.anime)+'</td><td class="hide-sp small">'+esc(x.type)+'</td>'+
        '<td class="ttl"><a data-q="'+esc(q)+'">'+esc(x.song)+'</a></td><td class="hide-sp small">'+esc(x.artist)+'</td></tr>';
    }
    html+='</tbody></table></details>';
  }
  saveFile('created_list'+(cat==='ALL'?'':'_'+cat)+'.html','作成済みリスト',html||'<div class="empty">該当なし</div>');
};

document.getElementById('saveUncreated').onclick=()=>{
  const cat=document.getElementById('analysisCategory').value;
  let html='';
  for(const [k,v] of Object.entries(APP.uncreatedLists)){
    if(cat!=='ALL'&&cat!==k)continue;
    if(!v.length)continue;
    html+='<details class="cat-section" open><summary>'+esc(k)+'<span class="meta-info">'+v.length+'曲</span></summary>'+
      '<table class="tbl"><thead><tr><th>作品名</th><th class="hide-sp" style="width:60px">区分</th><th>曲名</th><th class="hide-sp">歌手</th></tr></thead><tbody>';
    for(const x of v){
      const q=(x.anime?x.anime+' ':'')+x.song;
      html+='<tr><td class="small">'+esc(x.anime)+'</td><td class="hide-sp small">'+esc(x.type)+'</td>'+
        '<td class="ttl"><a data-q="'+esc(q)+'">'+esc(x.song)+'</a></td><td class="hide-sp small">'+esc(x.artist)+'</td></tr>';
    }
    html+='</tbody></table></details>';
  }
  saveFile('uncreated_list'+(cat==='ALL'?'':'_'+cat)+'.html','未作成リスト',html||'<div class="empty">該当なし</div>');
};
</script>
</body>
</html>
"""

html_content = html_content.replace("__UPDATED__", current_datetime_str)
html_content = html_content.replace("__APP_JSON__", app_json)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
    print("HTML生成完了: index.html")
