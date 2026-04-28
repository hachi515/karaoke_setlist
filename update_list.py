import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
import json
import io
import base64
from itertools import groupby
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

# 推移タブ対象クール（特集）
TREND_TARGET_CATEGORY = "2026年春アニメ"
TREND_PERIOD_OPTIONS = [3, 7, 14, 30]

# ALLOWED CATEGORIES
ALLOWED_CATEGORIES = ["2026年春アニメ", "2026年冬アニメ", "2025年秋アニメ"]

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
            print(f"[GitHub] Error: Status {response.status_code}")
            return pd.DataFrame()
    except Exception as e:
        print(f"[GitHub] Connection error: {e}")
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


def load_json_from_gas(filename):
    """GASからJSONを読み込む(image_map.json用)"""
    try:
        response = requests.get(GAS_WEB_APP_URL, params={'filename': filename}, timeout=30)
        if response.status_code == 200:
            text = response.text
            if not text or "Exception" in text[:200]:
                return {}
            try:
                return json.loads(text)
            except Exception:
                return {}
        return {}
    except Exception as e:
        print(f"[GAS] JSON load error: {e}")
        return {}


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
current_date_dash = now.strftime("%Y-%m-%d")

room_map = {
    11000: "ゆーふうりん部屋", 11001: "ゆーふうりん部屋", 11002: "ゆーふうりん部屋",
    11003: "ゆーふうりん部屋", 11004: "ゆーふうりん部屋", 11005: "ゆーふうりん部屋",
    11006: "ゆーふうりん部屋", 11007: "ゆーふうりん部屋", 11008: "ゆーふうりん部屋",
    11009: "ゆーふうりん部屋", 11012: "加古部屋", 11021: "成田部屋", 11022: "成田部屋",
    11028: "タマ部屋", 11058: "すみた部屋", 11059: "つぼはち部屋", 11060: "れん部屋",
    11063: "なぎ部屋", 11064: "naoo部屋", 11066: "芝ちゃん部屋", 11067: "crom部屋",
    11068: "けんしん部屋", 11069: "けんちぃ部屋", 11070: "黒河部屋", 11071: "黒河部屋",
    11074: "tukinowa部屋", 11077: "v3部屋", 11078: "のんでるん部屋", 11079: "まどか部屋",
    11087: "MiO部屋", 11088: "ほっしー部屋", 11091: "千秋部屋", 11092: "ヒロ部屋",
    11101: "えみち部屋", 11102: "るえ部屋", 11103: "ながし部屋", 11104: "MrN部屋",
    11105: "ヤマテル部屋", 11106: "冨塚部屋", 11107: "ブルーベリー部屋",
    11110: "加古部屋", 11111: "ヒロ部屋"
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


# --- 1. 過去データ読み込み・安全な履歴ローテーション ---
HISTORY_MAX_ROWS = 9500
HISTORY_ARCHIVE_MISS_LIMIT = 3
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
    if df is None or df.empty or '順���' not in df.columns:
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
            print(f"[STOP] {filename} の読み込みに失敗。")
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
    print(f"履歴ファイルを読み込みました: {', '.join(loaded_history_files)}")
else:
    print("履歴ファイルなし。")

history_dfs = [h['df'] for h in history_records]
full_history_before_update = cleanup_history_df(pd.concat(history_dfs, ignore_index=True)) if history_dfs else pd.DataFrame()

# --- 2. 新しいデータ取得 ---
target_ports = list(room_map.keys())
new_data_frames = []
failed_ports = []
fetched_ports = []

print("データを取得中...")
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

print(f"[Fetch] 成功ポート: {len(fetched_ports)} / {len(target_ports)}")

if not new_data_frames:
    print("[Fetch] 取得成功データなし。")
    full_df = full_history_before_update
else:
    new_df = cleanup_history_df(pd.concat(new_data_frames, ignore_index=True))
    dedup_cols = [c for c in HISTORY_DEDUP_COLS if c in new_df.columns]
    if dedup_cols:
        new_df = new_df.drop_duplicates(subset=dedup_cols, keep='last')
        new_df = cleanup_history_df(new_df)
    print(f"[Fetch] 今回収集: {len(new_df)} 行")

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
            print("追加対象なし。")
            full_df = full_history_before_update
        else:
            print(f"追加対象: {len(new_unique_df)} 行")
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
            save_ok = True

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
                    print(f"[History] {active_filename} +{len(append_part)}行 → {len(next_df)}行")
                    saved_parts[active_num] = next_df
                    remaining_df = remaining_df.iloc[remaining_capacity:].reset_index(drop=True)
                    if not remaining_df.empty:
                        active_num += 1
                        active_df = pd.DataFrame()
                    else:
                        active_df = next_df
                else:
                    print(f"[STOP] {active_filename} 保存失敗。")
                    save_ok = False
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
# オフラインリスト読み込み
# ==========================================
offline_targets = []
print(f"GitHubからオフラインリストを読み込み...")
for filename in OFFLINE_FILES:
    offline_df = load_df_from_github(filename)
    if not offline_df.empty:
        offline_df = offline_df.fillna("")
        if '曲名' in offline_df.columns:
            targets = [normalize_offline_text(str(x)) for x in offline_df['曲名'].tolist()]
            offline_targets.extend(targets)
            print(f"  -> {filename}: {len(targets)}件")
print(f"オフライン合計: {len(offline_targets)}")


# ==========================================
# 集計データ作成
# ==========================================
analysis_source_df = full_df.copy()
if not analysis_source_df.empty:
    analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
    analysis_source_df = analysis_source_df.dropna(subset=['dt_obj'])
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
    full_history = analysis_source_df[
        (~analysis_source_df['歌った人'].astype(str).apply(lambda x: any(k in x for k in exclude_keywords)))
    ].sort_values('dt_obj').reset_index(drop=True)
else:
    full_history = pd.DataFrame()


# --- Cool Analysis 読み込み ---
cool_file = "cool_analysis.csv"
raw_df = load_df_from_gas(cool_file, header=None)

categorized_data = {}
if not raw_df.empty:
    try:
        raw_df = raw_df.fillna("").drop_duplicates(keep='last')
        current_category = None
        for idx, row in raw_df.iterrows():
            if not any(str(x).strip() for x in row):
                continue
            col0 = str(row[0]).strip()
            is_category_line = any(cat in col0 for cat in ALLOWED_CATEGORIES) and "作品名" not in col0
            if is_category_line:
                current_category = col0
                if current_category not in categorized_data:
                    categorized_data[current_category] = []
                continue
            if "作品名" in col0:
                continue
            if current_category is None:
                continue
            anime = str(row[0]).strip() if len(row) > 0 else ""
            type_ = str(row[1]).strip() if len(row) > 1 else ""
            artist = str(row[2]).strip() if len(row) > 2 else ""
            song = str(row[3]).strip() if len(row) > 3 else ""
            if not anime and not song:
                continue
            categorized_data[current_category].append({
                "anime": anime, "type": type_, "artist": artist, "song": song
            })
    except Exception as e:
        print(f"Cool解析エラー: {e}")


def is_song_created(item):
    target_song_norm = normalize_text(item["song"])
    target_song_raw_norm = normalize_offline_text(item["song"])
    target_anime_norm = normalize_text(item["anime"])
    creation_count = 0
    if target_song_norm:
        for offline_str in offline_targets:
            if (target_song_norm in offline_str) or (target_song_raw_norm in offline_str):
                if target_anime_norm:
                    if target_anime_norm in offline_str:
                        creation_count += 1
                else:
                    creation_count += 1
    return creation_count


def compute_match_for_item(item, history_df):
    """アイテム毎にhistory_df内のマッチ行のインデックスを返す"""
    if history_df.empty:
        return []
    target_song_norm = normalize_text(item["song"])
    target_anime_norm = normalize_text(item["anime"])
    song_match_mask = check_match(target_song_norm, history_df['norm_filename'])
    anime_match_mask = (
        history_df['norm_filename'].str.contains(re.escape(target_anime_norm), case=False, na=False) |
        history_df['norm_workname'].str.contains(re.escape(target_anime_norm), case=False, na=False)
    ) if target_anime_norm else pd.Series([False] * len(history_df))
    if target_song_norm and target_anime_norm:
        final_mask = song_match_mask & anime_match_mask
    elif target_song_norm:
        final_mask = song_match_mask
    elif target_anime_norm:
        final_mask = anime_match_mask
    else:
        return []
    return history_df.index[final_mask].tolist()


# --- 集計対象期間 ---
COOL_START = pd.to_datetime("2026/01/01")
COOL_END = pd.to_datetime("2026/06/30")
target_history = full_history[
    (full_history['dt_obj'] >= COOL_START) & (full_history['dt_obj'] <= COOL_END)
] if not full_history.empty else pd.DataFrame()


# --- カテゴリごとに作品単位でグルーピング ---
cool_data_for_js = {}

for category in ALLOWED_CATEGORIES:
    items = categorized_data.get(category, [])
    if not items:
        cool_data_for_js[category] = {"works": [], "max_count": 0, "max_user": 0}
        continue

    # アイテムごとに count, user_count, creation 計算
    items_enriched = []
    for item in items:
        match_idx = compute_match_for_item(item, target_history) if not target_history.empty else []
        if match_idx:
            matched = target_history.loc[match_idx]
            count = len(matched)
            user_count = matched['歌った人'].nunique()
        else:
            count = 0
            user_count = 0
        cc = is_song_created(item)
        items_enriched.append({
            **item,
            "count": count,
            "user_count": user_count,
            "creation_count": cc
        })

    # 作品名でグループ化
    items_enriched.sort(key=lambda x: x['anime'])
    works = []
    for anime_name, group_iter in groupby(items_enriched, key=lambda x: x['anime']):
        group = list(group_iter)
        op_n = sum(1 for g in group if 'OP' in g['type'].upper())
        ed_n = sum(1 for g in group if 'ED' in g['type'].upper())
        in_n = sum(1 for g in group if 'IN' in g['type'].upper())
        total_count = sum(g['count'] for g in group)
        total_user = sum(g['user_count'] for g in group)
        works.append({
            "anime": anime_name,
            "songs": group,
            "op_n": op_n, "ed_n": ed_n, "in_n": in_n,
            "total_count": total_count,
            "total_user": total_user
        })

    max_count = max([w['total_count'] for w in works], default=0)
    max_user = max([w['total_user'] for w in works], default=0)
    cool_data_for_js[category] = {
        "works": works,
        "max_count": max_count,
        "max_user": max_user
    }

print("クール集計データ生成完了。")


# ==========================================
# ランキング (歌唱数・歌唱人数)
# ==========================================
ranking_data_by_cat = {}
for category in ALLOWED_CATEGORIES:
    cool = cool_data_for_js.get(category, {"works": []})
    flat = []
    for w in cool['works']:
        for s in w['songs']:
            if s['count'] > 0:
                flat.append({
                    "anime": w['anime'],
                    "song": s['song'],
                    "artist": s['artist'],
                    "type": s['type'],
                    "count": s['count'],
                    "user_count": s['user_count']
                })
    ranking_data_by_cat[category] = flat


# ==========================================
# 推移（急上昇）計算: 2026年春アニメのみ
# ==========================================
trend_data_for_js = {}
target_cool_works = cool_data_for_js.get(TREND_TARGET_CATEGORY, {"works": []})['works']

# フラット化
trend_items = []
for w in target_cool_works:
    for s in w['songs']:
        trend_items.append({
            "anime": w['anime'],
            "song": s['song'],
            "artist": s['artist'],
            "type": s['type']
        })

now_dt = pd.Timestamp(now.replace(tzinfo=None).date()) + pd.Timedelta(days=1)  # exclusive end

for period_days in TREND_PERIOD_OPTIONS:
    cur_start = now_dt - pd.Timedelta(days=period_days)
    cur_end = now_dt
    prev_start = cur_start - pd.Timedelta(days=period_days)
    prev_end = cur_start

    if not full_history.empty:
        cur_history = full_history[(full_history['dt_obj'] >= cur_start) & (full_history['dt_obj'] < cur_end)]
        prev_history = full_history[(full_history['dt_obj'] >= prev_start) & (full_history['dt_obj'] < prev_end)]
        all_history_for_rank = full_history[(full_history['dt_obj'] >= COOL_START) & (full_history['dt_obj'] <= COOL_END)]
    else:
        cur_history = pd.DataFrame()
        prev_history = pd.DataFrame()
        all_history_for_rank = pd.DataFrame()

    item_stats = []
    for it in trend_items:
        cur_idx = compute_match_for_item(it, cur_history) if not cur_history.empty else []
        prev_idx = compute_match_for_item(it, prev_history) if not prev_history.empty else []
        all_idx = compute_match_for_item(it, all_history_for_rank) if not all_history_for_rank.empty else []
        cur_count = len(cur_idx)
        prev_count = len(prev_idx)
        cur_user = cur_history.loc[cur_idx]['歌った人'].nunique() if cur_idx else 0
        all_count = len(all_idx)
        all_user = all_history_for_rank.loc[all_idx]['歌った人'].nunique() if all_idx else 0
        delta = cur_count - prev_count
        is_new = (cur_count > 0 and prev_count == 0)
        item_stats.append({
            **it,
            "cur_count": cur_count,
            "cur_user": cur_user,
            "prev_count": prev_count,
            "delta": delta,
            "is_new": is_new,
            "all_count": all_count,
            "all_user": all_user
        })

    # 全体ランキング(歌唱数)で現在順位
    sorted_for_rank = sorted([x for x in item_stats if x['all_count'] > 0],
                             key=lambda x: (-x['all_count'], -x['all_user']))
    rank_map = {}
    prev_val = None
    cur_rank = 0
    for i, it in enumerate(sorted_for_rank):
        v = it['all_count']
        if v != prev_val:
            cur_rank = i + 1
            prev_val = v
        rank_map[(it['anime'], it['song'])] = cur_rank

    for it in item_stats:
        it['rank_now'] = rank_map.get((it['anime'], it['song']), 0)

    surge_count = sum(1 for x in item_stats if x['delta'] > 0)
    new_in = sum(1 for x in item_stats if x['is_new'])
    max_delta = max([x['delta'] for x in item_stats], default=0)

    # 急上昇順
    surge_sorted = sorted([x for x in item_stats if x['delta'] != 0 or x['cur_count'] > 0],
                          key=lambda x: (-x['delta'], -x['cur_count']))

    trend_data_for_js[str(period_days)] = {
        "kpi": {"surge_count": surge_count, "new_in": new_in, "max_delta": max_delta},
        "items": surge_sorted
    }


# ==========================================
# image_map.json 読み込み（Drive画像紐付け）
# ==========================================
image_map = load_json_from_gas("image_map.json")
if not isinstance(image_map, dict):
    image_map = {}
print(f"image_map: {sum(len(v) for v in image_map.values()) if isinstance(image_map, dict) else 0} entries")


# ==========================================
# 履歴データをJSへ埋め込む
# ==========================================
history_for_js = []
if not full_history.empty:
    h_subset = full_history[['取得日', '部屋主', '順番', '曲名（ファイル名）', '作品名', '歌手名', '歌った人', 'norm_filename', 'norm_workname']].copy()
    h_subset = h_subset.fillna("")
    h_subset['順番'] = h_subset['順番'].astype(str)
    for _, r in h_subset.iterrows():
        history_for_js.append({
            "d": str(r['取得日']),
            "rm": str(r['部屋主']),
            "o": str(r['順番']),
            "sg": str(r['曲名（ファイル名）']),
            "wk": str(r['作品名']),
            "ar": str(r['歌手名']),
            "u": str(r['歌った人']),
            "sn": str(r['norm_filename']),
            "wn": str(r['norm_workname'])
        })

# 部屋一覧
unique_rooms = sorted(list(set(room_map.values())))


# ==========================================
# JSON シリアライズ
# ==========================================
cool_json = json.dumps(cool_data_for_js, ensure_ascii=False)
ranking_json = json.dumps(ranking_data_by_cat, ensure_ascii=False)
trend_json = json.dumps(trend_data_for_js, ensure_ascii=False, default=str)
history_json = json.dumps(history_for_js, ensure_ascii=False)
image_map_json = json.dumps(image_map, ensure_ascii=False)
rooms_json = json.dumps(unique_rooms, ensure_ascii=False)
categories_json = json.dumps(ALLOWED_CATEGORIES, ensure_ascii=False)
trend_periods_json = json.dumps(TREND_PERIOD_OPTIONS, ensure_ascii=False)
trend_target_json = json.dumps(TREND_TARGET_CATEGORY, ensure_ascii=False)


# ==========================================
# HTML 生成
# ==========================================
html_content = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>Karaoke Dashboard</title>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Noto+Sans+JP:wght@400;500;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --primary:#1f2937; --accent:#6366f1; --accent-2:#7c3aed; --accent-soft:#a5b4fc;
  --bg:#f5f5fa; --panel:#fff; --text:#1f2937; --text-sub:#6b7280; --text-mute:#9ca3af;
  --border:#e5e7eb; --border-soft:#eef0f3;
  --green:#10b981; --green-bg:#ecfdf5; --red:#ef4444; --amber:#f59e0b;
  --gold:#f59e0b; --silver:#9ca3af; --bronze:#d97706;
  --radius:10px; --radius-lg:14px;
  --maxw:760px;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0;background:var(--bg);color:var(--text);
  font-family:"Inter","Noto Sans JP","Helvetica Neue",Arial,sans-serif;font-size:14px;line-height:1.45;
  -webkit-font-smoothing:antialiased;}
body{padding-bottom:20px}
a{color:inherit;text-decoration:none}
button{font-family:inherit}

/* ---------- Header ---------- */
.app-header{
  max-width:var(--maxw);margin:0 auto;padding:14px 16px 0 16px;
  display:flex;align-items:center;justify-content:space-between;
}
.brand{font-size:20px;font-weight:800;color:var(--primary);letter-spacing:0.01em}
.bell{
  width:34px;height:34px;border-radius:50%;background:#fff;border:1px solid var(--border);
  display:flex;align-items:center;justify-content:center;color:var(--text-sub);position:relative;cursor:default;
}
.bell::after{content:"";position:absolute;top:8px;right:9px;width:7px;height:7px;background:var(--accent);border-radius:50%}

/* Top settings cards */
.top-cards{
  max-width:var(--maxw);margin:10px auto 0;padding:0 16px;
  display:grid;grid-template-columns:1fr 1fr;gap:10px;
}
.top-card{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius);padding:10px 12px;
  display:flex;align-items:center;gap:10px;
}
.top-card .ico{
  width:34px;height:34px;border-radius:50%;background:#eef2ff;color:var(--accent);
  display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0;
}
.top-card .lbl{font-size:11.5px;color:var(--text-sub);margin-bottom:2px}
.top-card .val{display:flex;align-items:center;gap:6px}
.top-card input,.top-card select{
  border:none;background:transparent;font-weight:700;font-size:15px;color:var(--primary);outline:none;width:100%;
}
.top-card select{cursor:pointer;font-size:14px}
.top-card input[type=number]{width:90px}

/* Tabs */
.tabs-wrap{max-width:var(--maxw);margin:14px auto 0;padding:0 16px}
.tabs{
  display:flex;gap:0;background:transparent;border-bottom:1px solid var(--border);overflow-x:auto;
  scrollbar-width:none;
}
.tabs::-webkit-scrollbar{display:none}
.tab-btn{
  padding:10px 12px;border:none;background:none;color:var(--text-sub);font-weight:600;font-size:13px;
  white-space:nowrap;cursor:pointer;border-bottom:2px solid transparent;display:inline-flex;align-items:center;gap:5px;
  transition:color .15s, border-color .15s;
}
.tab-btn:hover{color:var(--primary)}
.tab-btn.active{color:var(--accent);border-bottom-color:var(--accent)}
.tab-btn i{font-size:13px}

/* Toolbar (per-tab) */
.tab-toolbar{max-width:var(--maxw);margin:12px auto 0;padding:0 16px;display:flex;flex-direction:column;gap:8px}
.toolbar-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.search-pill{
  flex:1;min-width:0;background:#fff;border:1px solid var(--border);border-radius:999px;padding:9px 14px;
  display:flex;align-items:center;gap:8px;
}
.search-pill input{flex:1;border:none;outline:none;font-size:13.5px;background:transparent;min-width:0}
.search-pill i{color:var(--text-mute)}
.icon-btn{
  width:38px;height:38px;border-radius:10px;background:#fff;border:1px solid var(--border);
  display:flex;align-items:center;justify-content:center;color:var(--text-sub);cursor:pointer;flex-shrink:0;
}
.icon-btn:hover{color:var(--accent);border-color:var(--accent-soft)}
.icon-btn.active{background:var(--accent);color:#fff;border-color:var(--accent)}

.pill-select{
  background:#fff;border:1px solid var(--border);border-radius:999px;padding:7px 14px;font-size:13px;
  display:inline-flex;align-items:center;gap:6px;cursor:pointer;font-weight:600;color:var(--primary);
  appearance:none;-webkit-appearance:none;background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='none' stroke='%236b7280' stroke-width='3'><polyline points='6 9 12 15 18 9'/></svg>");
  background-repeat:no-repeat;background-position:right 12px center;padding-right:30px;
}

.dl-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap;justify-content:flex-end}
.dl-btn{
  background:linear-gradient(180deg,var(--accent),var(--accent-2));color:#fff;border:none;border-radius:999px;
  padding:7px 14px;font-size:12.5px;font-weight:600;cursor:pointer;display:inline-flex;align-items:center;gap:6px;
}
.dl-btn.ghost{background:#fff;color:var(--accent);border:1px solid var(--accent-soft)}
.dl-btn:hover{filter:brightness(1.05)}

.update-line{
  max-width:var(--maxw);margin:6px auto 0;padding:0 16px;
  display:flex;justify-content:flex-end;font-size:11.5px;color:var(--text-mute);
}

.count-line{
  max-width:var(--maxw);margin:8px auto 0;padding:0 16px;
  font-size:13px;color:var(--text-sub);font-weight:600;display:flex;align-items:center;gap:6px;
}
.count-line i{color:var(--accent)}

/* ---------- Tab content container ---------- */
.tab-content{display:none;max-width:var(--maxw);margin:0 auto;padding:8px 16px 80px 16px}
.tab-content.active{display:block}

/* ---------- Card (common) ---------- */
.card{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius);
  margin-top:8px;overflow:hidden;
  transition:border-color .15s, box-shadow .15s;
}
.card.expanded{border-color:var(--accent-soft);box-shadow:0 4px 12px rgba(99,102,241,.08)}
.card-head{
  padding:12px;display:flex;align-items:flex-start;gap:10px;cursor:pointer;
}
.num-badge{
  width:28px;height:28px;border-radius:50%;background:linear-gradient(180deg,var(--accent),var(--accent-2));
  color:#fff;font-weight:800;font-size:13px;display:flex;align-items:center;justify-content:center;flex-shrink:0;
  box-shadow:0 2px 6px rgba(99,102,241,.3);
}
.num-badge.gold{background:linear-gradient(180deg,#fde68a,#f59e0b);color:#78350f;box-shadow:0 2px 6px rgba(245,158,11,.3)}
.num-badge.silver{background:linear-gradient(180deg,#e5e7eb,#9ca3af);color:#1f2937}
.num-badge.bronze{background:linear-gradient(180deg,#fed7aa,#d97706);color:#7c2d12}
.card-body{flex:1;min-width:0}
.card-chev{color:var(--text-mute);transition:transform .2s;font-size:12px;align-self:center}
.card.expanded .card-chev{transform:rotate(180deg)}
.card-detail{display:none;border-top:1px solid var(--border-soft);background:#fafbff}
.card.expanded .card-detail{display:block}

/* ---------- Setlist card ---------- */
.sl-top{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:2px}
.room-tag{
  display:inline-block;padding:2px 8px;font-size:11px;font-weight:600;color:var(--accent);
  background:#eef2ff;border:1px solid #e0e7ff;border-radius:5px;line-height:1.3;
}
.sl-date{font-size:12px;color:var(--text-mute);white-space:nowrap;font-variant-numeric:tabular-nums}
.sl-song{font-weight:700;font-size:15px;color:var(--primary);margin:4px 0 2px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sl-meta{font-size:12px;color:var(--text-sub);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sl-meta .sep{margin:0 6px;color:var(--text-mute)}

/* card-detail rows */
.detail-table{padding:8px 12px}
.detail-row{
  display:grid;grid-template-columns:90px 1fr;gap:8px;padding:6px 8px;
  background:#fff;border-radius:6px;margin-bottom:4px;align-items:start;
}
.detail-row .lbl{font-size:11.5px;color:var(--accent);font-weight:600;display:flex;align-items:center;gap:5px}
.detail-row .lbl i{font-size:11px}
.detail-row .val{font-size:13px;color:var(--primary);word-break:break-word;line-height:1.4}
.confirm-btn{
  display:flex;align-items:center;justify-content:center;gap:8px;
  margin:8px 12px 12px;padding:11px 16px;
  background:linear-gradient(180deg,var(--accent),var(--accent-2));color:#fff;
  border:none;border-radius:8px;font-weight:600;font-size:13.5px;cursor:pointer;width:calc(100% - 24px);
}
.confirm-btn i.fa-arrow-right,.confirm-btn i.fa-chevron-right{margin-left:auto}
.confirm-btn:hover{filter:brightness(1.05)}

/* ---------- Cool ---------- */
.cool-head-row{display:flex;align-items:flex-start;gap:10px}
.cool-anime-block{flex:1;min-width:0}
.cool-anime{font-weight:700;font-size:14.5px;color:var(--primary);line-height:1.35;
  display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.cool-types{margin-top:6px;display:flex;gap:5px;flex-wrap:wrap}
.type-pill{
  display:inline-flex;align-items:center;gap:4px;padding:2px 8px;font-size:10.5px;font-weight:700;
  border-radius:4px;color:var(--text-sub);background:#f3f4f6;border:1px solid var(--border);
}
.type-pill b{color:var(--primary)}
.type-pill.has{background:#eef2ff;border-color:#e0e7ff;color:var(--accent)}

.metric-block{display:flex;gap:10px;align-items:center}
.metric{display:flex;flex-direction:column;align-items:center;gap:3px;min-width:46px}
.metric .lbl{font-size:10.5px;color:var(--text-sub);font-weight:600}
.metric .val{font-size:13px;font-weight:800;color:var(--primary)}
.dot-grid{
  display:grid;grid-template-columns:repeat(3,6px);grid-template-rows:repeat(3,6px);gap:2px;
}
.dot{width:6px;height:6px;background:#e5e7eb;border-radius:1px}
.dot.fu{background:var(--green)}
.dot.fc{background:var(--accent)}

/* Cool detail (songs) */
.song-row{
  display:grid;grid-template-columns:36px 1fr auto;gap:10px;padding:8px 12px;
  border-bottom:1px solid var(--border-soft);align-items:center;
}
.song-row:last-child{border-bottom:none}
.song-type{
  display:inline-flex;align-items:center;justify-content:center;
  min-width:32px;padding:3px 6px;font-size:10.5px;font-weight:800;
  border-radius:4px;color:var(--accent);background:#eef2ff;border:1px solid #e0e7ff;
}
.song-info{min-width:0}
.song-name{font-weight:700;font-size:13.5px;color:var(--primary);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.song-artist{font-size:11.5px;color:var(--text-sub);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:2px}
.song-metrics{display:flex;gap:8px;align-items:center}
.song-confirm{
  text-align:center;padding:10px 12px 12px;
}
.show-all-link{
  display:flex;align-items:center;justify-content:center;gap:8px;padding:9px 12px;
  border-top:1px solid var(--border-soft);font-size:12.5px;color:var(--text-sub);font-weight:600;cursor:pointer;
  background:#fff;
}
.show-all-link:hover{color:var(--accent)}

/* ---------- Ranking ---------- */
.rank-card-top3{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius-lg);margin-top:10px;overflow:hidden;
}
.rank-card-top3.gold{background:linear-gradient(180deg,#fffbeb 0%,#fff 60%)}
.rank-card-top3.silver{background:linear-gradient(180deg,#f9fafb 0%,#fff 60%)}
.rank-card-top3.bronze{background:linear-gradient(180deg,#fff7ed 0%,#fff 60%)}
.rank-top3-head{padding:14px;display:flex;gap:12px;align-items:flex-start;cursor:pointer}
.rank-crown{font-size:18px;margin-bottom:2px}
.rank-crown.gold{color:var(--gold)}
.rank-crown.silver{color:var(--silver)}
.rank-crown.bronze{color:var(--bronze)}
.rank-top3-num{
  width:42px;height:42px;border-radius:50%;display:flex;align-items:center;justify-content:center;
  font-weight:800;font-size:18px;color:#fff;flex-shrink:0;
  border:3px solid #fff;box-shadow:0 0 0 2px var(--accent),0 4px 8px rgba(99,102,241,.3);
  background:linear-gradient(180deg,var(--accent),var(--accent-2));
}
.rank-top3-num.gold{box-shadow:0 0 0 2px var(--gold),0 4px 8px rgba(245,158,11,.3);background:linear-gradient(180deg,#fbbf24,#f59e0b)}
.rank-top3-num.silver{box-shadow:0 0 0 2px var(--silver),0 4px 8px rgba(156,163,175,.3);background:linear-gradient(180deg,#d1d5db,#9ca3af)}
.rank-top3-num.bronze{box-shadow:0 0 0 2px var(--bronze),0 4px 8px rgba(217,119,6,.3);background:linear-gradient(180deg,#f97316,#d97706)}
.rank-top3-info{flex:1;min-width:0}
.rank-top3-song{font-weight:800;font-size:18px;color:var(--primary);
  word-break:break-word;line-height:1.3}
.rank-top3-anime{font-size:12.5px;color:var(--text-sub);margin-top:2px;line-height:1.3}
.rank-top3-artist{font-size:12.5px;color:var(--text-sub);margin-top:1px}
.rank-top3-types{margin-top:8px;display:flex;gap:5px;flex-wrap:wrap}
.rank-top3-metrics{
  display:grid;grid-template-columns:1fr 1fr;gap:8px;padding:0 14px 14px;
}
.rank-top3-metric{
  background:#f9fafb;border:1px solid var(--border);border-radius:10px;padding:10px;text-align:center;
}
.rank-top3-metric .ico{
  width:32px;height:32px;border-radius:50%;background:#eef2ff;color:var(--accent);
  display:flex;align-items:center;justify-content:center;margin:0 auto 4px;
}
.rank-top3-metric .ico.user{background:var(--green-bg);color:var(--green)}
.rank-top3-metric .ico.song{background:#fff7ed;color:var(--amber)}
.rank-top3-metric .lbl{font-size:11px;color:var(--text-sub);font-weight:600;margin-bottom:2px}
.rank-top3-metric .val{font-size:22px;font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums}

/* Normal rank card (4位以下) */
.rank-card{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius);margin-top:8px;overflow:hidden;
}
.rank-card.expanded{border-color:var(--accent-soft);box-shadow:0 4px 12px rgba(99,102,241,.08)}
.rank-row{padding:10px 12px;display:grid;grid-template-columns:36px 1fr auto auto auto;gap:10px;align-items:center;cursor:pointer}
.rank-info{min-width:0}
.rank-anime{font-weight:700;font-size:14px;color:var(--primary);line-height:1.3;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-sub{font-size:11.5px;color:var(--text-sub);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:2px}
.rank-types-inline{margin-top:4px;display:flex;gap:4px;flex-wrap:wrap}
.rank-mini-metric{
  background:#f9fafb;border:1px solid var(--border);border-radius:8px;padding:5px 9px;text-align:center;min-width:54px;
}
.rank-mini-metric .lbl{font-size:9.5px;color:var(--text-sub);font-weight:600;display:flex;align-items:center;justify-content:center;gap:3px}
.rank-mini-metric .lbl i{font-size:9.5px;color:var(--accent)}
.rank-mini-metric .lbl i.user{color:var(--green)}
.rank-mini-metric .lbl i.song{color:var(--amber)}
.rank-mini-metric .val{font-size:14px;font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums}

/* ---------- Trend ---------- */
.trend-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:10px}
.trend-stat{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius);padding:10px;text-align:center;
}
.trend-stat .ico{
  width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;margin:0 auto 4px;
  font-size:13px;
}
.trend-stat .ico.up{background:#eef2ff;color:var(--accent)}
.trend-stat .ico.new{background:var(--green-bg);color:var(--green)}
.trend-stat .ico.max{background:#fff7ed;color:var(--amber)}
.trend-stat .lbl{font-size:11px;color:var(--text-sub);font-weight:600}
.trend-stat .val{font-size:18px;font-weight:800;color:var(--primary)}
.trend-stat .val small{font-size:11px;font-weight:600;color:var(--text-sub);margin-left:2px}

.trend-pickup{
  background:linear-gradient(135deg,#eef2ff 0%,#fff 50%,#fef3c7 100%);
  border:1px solid var(--accent-soft);border-radius:var(--radius-lg);
  margin-top:14px;padding:14px;position:relative;overflow:hidden;
}
.trend-pickup-icon{
  position:absolute;top:8px;right:10px;font-size:48px;color:rgba(99,102,241,.15);
}
.trend-pickup-head{display:flex;align-items:center;gap:6px;margin-bottom:8px;font-weight:700;font-size:13.5px;color:var(--primary)}
.trend-pickup-head i{color:var(--red)}
.trend-pickup-body{display:grid;grid-template-columns:80px 1fr auto;gap:12px;align-items:flex-start}
.thumb-square{
  width:80px;height:80px;border-radius:10px;background:#e5e7eb;overflow:hidden;flex-shrink:0;
  display:flex;align-items:center;justify-content:center;color:var(--text-mute);font-size:24px;position:relative;
}
.thumb-square img{width:100%;height:100%;object-fit:cover}
.thumb-square.large{width:96px;height:96px}
.thumb-tag{
  position:absolute;top:0;left:0;background:linear-gradient(180deg,#fbbf24,#f59e0b);color:#78350f;
  padding:2px 7px;font-size:10px;font-weight:800;border-radius:0 0 6px 0;letter-spacing:.05em;
}
.tp-info{min-width:0}
.tp-anime{font-weight:800;font-size:18px;color:var(--primary);line-height:1.25;word-break:break-word}
.tp-artist{font-size:12px;color:var(--text-sub);margin-top:3px}
.tp-type{margin-top:6px}
.tp-side{
  display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;text-align:center;
}
.tp-side .arrow{font-size:32px;color:var(--accent)}
.tp-side .pre-label{font-size:10px;color:var(--text-sub);font-weight:600}
.tp-side .delta{font-size:18px;font-weight:800;color:var(--accent)}

.tp-stats{
  display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-top:12px;
}
.tp-stat{background:#fff;border:1px solid var(--border);border-radius:10px;padding:8px;text-align:center}
.tp-stat .lbl{font-size:10.5px;color:var(--text-sub);font-weight:600;display:flex;align-items:center;justify-content:center;gap:3px}
.tp-stat .val{font-size:18px;font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums}
.tp-stat .lbl i{color:var(--accent)}

.notable-head{
  margin-top:18px;font-size:14px;font-weight:700;color:var(--primary);display:flex;align-items:center;gap:6px;
}
.notable-head i{color:var(--accent)}

.notable-card{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius);margin-top:8px;overflow:hidden;
}
.notable-card.expanded{border-color:var(--accent-soft);box-shadow:0 4px 12px rgba(99,102,241,.08)}
.notable-row{padding:10px 12px;display:grid;grid-template-columns:32px 56px 1fr auto auto auto auto;gap:8px;align-items:center;cursor:pointer}
.notable-num{width:28px;height:28px;border-radius:50%;background:linear-gradient(180deg,var(--accent),var(--accent-2));color:#fff;font-weight:800;font-size:12px;display:flex;align-items:center;justify-content:center}
.thumb-mini{width:48px;height:48px;border-radius:8px;background:#e5e7eb;overflow:hidden;flex-shrink:0;display:flex;align-items:center;justify-content:center;color:var(--text-mute)}
.thumb-mini img{width:100%;height:100%;object-fit:cover}
.notable-info{min-width:0}
.notable-anime{font-weight:700;font-size:13px;color:var(--primary);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.notable-artist{font-size:11px;color:var(--text-sub);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:2px}
.notable-delta{font-size:13px;font-weight:800;color:var(--accent);background:#eef2ff;padding:3px 8px;border-radius:6px;white-space:nowrap}
.notable-delta.new{color:var(--green);background:var(--green-bg)}
.notable-mini-metric{
  background:transparent;text-align:center;min-width:42px;
}
.notable-mini-metric .lbl{font-size:9.5px;color:var(--text-sub);display:flex;align-items:center;justify-content:center;gap:3px}
.notable-mini-metric .val{font-size:13px;font-weight:800;color:var(--primary)}
.notable-chev{color:var(--text-mute);font-size:11px}

/* ---------- Modal ---------- */
.modal-overlay{
  display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(15,23,42,.6);
  z-index:1000;align-items:flex-end;justify-content:center;
}
.modal-overlay.active{display:flex}
.modal{
  background:#fff;border-radius:14px 14px 0 0;width:100%;max-width:var(--maxw);max-height:88vh;
  display:flex;flex-direction:column;animation:slideUp .25s ease;
}
@keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:translateY(0);opacity:1}}
.modal-head{padding:14px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px}
.modal-head .ttl{flex:1;font-weight:700;font-size:14.5px;color:var(--primary)}
.modal-head .ttl small{display:block;font-size:11.5px;color:var(--text-sub);margin-top:2px;font-weight:500}
.modal-close{
  width:32px;height:32px;border-radius:50%;background:#f3f4f6;border:none;color:var(--text-sub);
  display:flex;align-items:center;justify-content:center;cursor:pointer;
}
.modal-body{flex:1;overflow-y:auto;padding:8px 12px 16px;-webkit-overflow-scrolling:touch}
.modal-summary{
  background:#eef2ff;border-radius:8px;padding:9px 12px;margin:6px 0 10px;font-size:12.5px;color:var(--accent);font-weight:600;
}
.modal-row{
  background:#fff;border:1px solid var(--border-soft);border-radius:8px;padding:9px 12px;margin-bottom:6px;
}
.modal-row .top{display:flex;justify-content:space-between;gap:8px;align-items:center;margin-bottom:3px}
.modal-row .user{font-weight:700;color:var(--primary);font-size:13.5px;display:flex;align-items:center;gap:5px}
.modal-row .user i{color:var(--accent);font-size:11px}
.modal-row .date{font-size:11.5px;color:var(--text-mute);font-variant-numeric:tabular-nums}
.modal-row .meta{font-size:11.5px;color:var(--text-sub);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}

/* ---------- Filter popover ---------- */
.popover{
  display:none;position:absolute;top:100%;right:0;margin-top:6px;background:#fff;
  border:1px solid var(--border);border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.1);
  padding:12px;min-width:280px;max-width:320px;z-index:50;
}
.popover.open{display:block}
.popover h4{margin:0 0 6px;font-size:12px;color:var(--text-sub);font-weight:600;text-transform:uppercase;letter-spacing:.04em}
.popover .opt{padding:7px 10px;border-radius:6px;cursor:pointer;font-size:13px;color:var(--primary)}
.popover .opt:hover,.popover .opt.selected{background:#eef2ff;color:var(--accent);font-weight:600}
.popover .opt.selected::before{content:"✓ ";color:var(--accent);font-weight:800}
.popover hr{border:none;border-top:1px solid var(--border-soft);margin:10px 0}
.room-chips{display:flex;flex-wrap:wrap;gap:5px;max-height:180px;overflow-y:auto}
.room-chip{
  display:inline-block;padding:4px 10px;font-size:11.5px;border-radius:999px;border:1px solid var(--border);
  background:#fff;color:var(--text-sub);cursor:pointer;
}
.room-chip.selected{background:var(--accent);color:#fff;border-color:var(--accent)}
.popover .actions{display:flex;justify-content:flex-end;gap:8px;margin-top:10px}
.popover .btn-clear{background:#f3f4f6;color:var(--text-sub);border:none;border-radius:6px;padding:6px 12px;font-size:12px;cursor:pointer}
.popover .btn-apply{background:var(--accent);color:#fff;border:none;border-radius:6px;padding:6px 12px;font-size:12px;cursor:pointer}

.toolbar-rel{position:relative}

/* ---------- Empty state ---------- */
.empty{padding:30px 16px;text-align:center;color:var(--text-mute);font-size:13px}

/* ---------- Env tab ---------- */
.env-section{margin-top:14px}
.env-section h3{font-size:14px;font-weight:700;color:var(--primary);margin:0 0 8px;padding-left:8px;border-left:3px solid var(--accent)}
.env-work-row{
  background:#fff;border:1px solid var(--border);border-radius:10px;padding:10px;margin-bottom:6px;
  display:flex;gap:10px;align-items:center;
}
.env-work-thumb{width:48px;height:48px;border-radius:8px;background:#e5e7eb;overflow:hidden;display:flex;align-items:center;justify-content:center;color:var(--text-mute);flex-shrink:0}
.env-work-thumb img{width:100%;height:100%;object-fit:cover}
.env-work-name{flex:1;font-weight:600;font-size:13px;color:var(--primary);min-width:0;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.env-upload-btn{
  background:var(--accent);color:#fff;border:none;border-radius:8px;padding:7px 12px;font-size:12px;cursor:pointer;
  display:inline-flex;align-items:center;gap:5px;
}
.env-upload-btn:hover{filter:brightness(1.05)}
.env-upload-btn input{display:none}
.env-status{margin-top:6px;padding:8px 12px;background:#eef2ff;border-radius:6px;color:var(--accent);font-size:12px;display:none}
.env-status.show{display:block}

/* ---------- Print ---------- */
@media print{
  *{-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important}
  .app-header,.top-cards,.tabs-wrap,.tab-toolbar,.update-line,.count-line,.confirm-btn,.dl-row{display:none !important}
  body{padding:0}
  .tab-content{display:block !important;max-width:none}
  .card,.rank-card,.notable-card{break-inside:avoid;page-break-inside:avoid}
  .card-detail{display:block !important}
  .card.expanded .card-chev{transform:none}
  .card-chev{display:none}
}

@media (max-width:420px){
  .top-cards{grid-template-columns:1fr 1fr}
  .top-card{padding:8px 10px}
  .tabs{padding:0}
}
</style>
</head>
<body>

<div class="app-header">
  <div class="brand">Karaoke Dashboard</div>
  <div class="bell" title="通知"><i class="far fa-bell"></i></div>
</div>

<div class="top-cards">
  <div class="top-card">
    <div class="ico"><i class="fas fa-network-wired"></i></div>
    <div style="flex:1;min-width:0">
      <div class="lbl">保存時ポート</div>
      <div class="val"><input type="number" id="exportPort" value="11059"></div>
    </div>
  </div>
  <div class="top-card">
    <div class="ico"><i class="fas fa-link"></i></div>
    <div style="flex:1;min-width:0">
      <div class="lbl">検索リンク</div>
      <div class="val">
        <select id="exportLinkType">
          <option value="eve">Everything</option>
          <option value="ykr">ゆかりすたー</option>
        </select>
      </div>
    </div>
  </div>
</div>

<div class="tabs-wrap">
  <div class="tabs">
    <button class="tab-btn active" data-tab="setlist"><i class="fas fa-music"></i> セットリスト</button>
    <button class="tab-btn" data-tab="cool"><i class="far fa-clock"></i> クール集計</button>
    <button class="tab-btn" data-tab="ranking"><i class="fas fa-trophy"></i> ランキング</button>
    <button class="tab-btn" data-tab="trend"><i class="fas fa-chart-line"></i> 推移</button>
    <button class="tab-btn" data-tab="env"><i class="fas fa-cog"></i> 環境設定</button>
  </div>
</div>

<div class="update-line"><i class="far fa-clock" style="margin-right:4px"></i><span id="updateLine">__UPDATE__ 更新</span></div>

<!-- ============ Setlist ============ -->
<div class="tab-content active" id="tab-setlist">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <div class="search-pill">
        <i class="fas fa-search"></i>
        <input type="text" id="slSearch" placeholder="曲名・作品名・歌手名・歌った人で検索">
      </div>
      <div class="toolbar-rel">
        <button class="icon-btn" id="slFilterBtn"><i class="fas fa-sliders-h"></i></button>
        <div class="popover" id="slPopover">
          <h4>並び替え</h4>
          <div class="opt selected" data-sort="date_desc">日付（新しい順） × 順番（大きい順）</div>
          <div class="opt" data-sort="date_asc">日付（古い順） × 順番（小さい順）</div>
          <div class="opt" data-sort="date_desc_order_asc">日付（新しい順） × 順番（小さい順）</div>
          <hr>
          <h4>部屋でフィルタ（複数選択可）</h4>
          <div class="room-chips" id="slRoomChips"></div>
          <div class="actions">
            <button class="btn-clear" id="slClearFilter">クリア</button>
            <button class="btn-apply" id="slApplyFilter">適用</button>
          </div>
        </div>
      </div>
    </div>
    <div class="dl-row">
      <button class="dl-btn ghost" onclick="downloadSetlistHTML()"><i class="fas fa-file-download"></i> セットリストHTML保存</button>
    </div>
  </div>
  <div class="count-line"><i class="fas fa-clipboard-list"></i><span id="slCount">0</span> 件</div>
  <div id="slList"></div>
</div>

<!-- ============ Cool ============ -->
<div class="tab-content" id="tab-cool">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <select class="pill-select" id="coolCat"></select>
      <select class="pill-select" id="coolSort">
        <option value="user">人気順 (人数)</option>
        <option value="count" selected>人気順 (歌唱数)</option>
        <option value="name">作品名順</option>
        <option value="created">作成数順</option>
      </select>
    </div>
    <div class="dl-row">
      <button class="dl-btn ghost" onclick="downloadCoolHTML('current')"><i class="fas fa-file-download"></i> このクールを保存</button>
      <button class="dl-btn" onclick="downloadCoolHTML('all')"><i class="fas fa-file-download"></i> 全クール保存</button>
    </div>
  </div>
  <div class="count-line"><i class="far fa-bookmark"></i><span id="coolCount">0</span> 作品</div>
  <div id="coolList"></div>
</div>

<!-- ============ Ranking ============ -->
<div class="tab-content" id="tab-ranking">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <select class="pill-select" id="rankCat"></select>
      <select class="pill-select" id="rankMode">
        <option value="count">歌唱数ランキング</option>
        <option value="user">歌唱人数ランキング</option>
      </select>
    </div>
    <div class="dl-row">
      <button class="dl-btn ghost" onclick="downloadRankingHTML('current')"><i class="fas fa-file-download"></i> このクールを保存</button>
      <button class="dl-btn" onclick="downloadRankingHTML('all')"><i class="fas fa-file-download"></i> 全クール保存</button>
    </div>
  </div>
  <div class="count-line"><i class="fas fa-list-ol"></i><span id="rankCount">0</span> 件</div>
  <div id="rankList"></div>
</div>

<!-- ============ Trend ============ -->
<div class="tab-content" id="tab-trend">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <select class="pill-select" id="trendCat" disabled></select>
    </div>
    <div class="toolbar-row">
      <select class="pill-select" id="trendPeriod"></select>
      <select class="pill-select" id="trendSort">
        <option value="surge">急上昇順</option>
        <option value="new">新規ランクイン順</option>
        <option value="max">最大伸び順</option>
      </select>
    </div>
    <div class="dl-row">
      <button class="dl-btn ghost" onclick="downloadTrendHTML()"><i class="fas fa-file-download"></i> 推移HTML保存</button>
    </div>
  </div>
  <div id="trendBody"></div>
</div>

<!-- ============ Env ============ -->
<div class="tab-content" id="tab-env">
  <div class="count-line" style="margin-top:12px"><i class="fas fa-cog"></i> 作品サムネイル管理（クール毎にDriveフォルダへアップロード）</div>
  <div class="env-status" id="envStatus"></div>
  <div id="envList"></div>
</div>

<!-- Modal -->
<div class="modal-overlay" id="modalOverlay">
  <div class="modal">
    <div class="modal-head">
      <div class="ttl" id="modalTitle">この曲を歌った人</div>
      <button class="modal-close" onclick="closeModal()"><i class="fas fa-times"></i></button>
    </div>
    <div class="modal-body" id="modalBody"></div>
  </div>
</div>

<script>
// ====== Inline data ======
const COOL_DATA = __COOL_JSON__;
const RANK_DATA = __RANK_JSON__;
const TREND_DATA = __TREND_JSON__;
const HISTORY = __HISTORY_JSON__;
const IMAGE_MAP = __IMAGE_MAP_JSON__;
const ROOMS = __ROOMS_JSON__;
const CATS = __CATS_JSON__;
const TREND_PERIODS = __TREND_PERIODS_JSON__;
const TREND_TARGET_CAT = __TREND_TARGET_JSON__;
const GAS_URL = "__GAS_URL__";
const UPDATE_TS = "__UPDATE_TS__";
const CURRENT_DATE = "__CURRENT_DATE__";

document.getElementById('updateLine').innerText = UPDATE_TS + ' 更新';

// ====== Utility ======
function jsNormalize(s){
  if(!s) return "";
  s = String(s).normalize('NFKC');
  s = s.replace(/\.[a-zA-Z0-9]{3,4}$/, '');
  s = s.replace(/[\[\(\{【].*?[\]\)\}】]/g, ' ');
  s = s.replace(/(key|KEY)?\s*[\+\-]\s*[0-9]+/g, ' ');
  s = s.replace(/原キー/g, ' ');
  s = s.replace(/(キー)?変更[:：]?/g, ' ');
  s = s.replace(/[~〜～\-_=,.]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s.toUpperCase();
}
function escHtml(s){
  return String(s||"").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function dotGrid(value, max, color){
  const filled = max>0 ? Math.min(9, Math.round((value/max)*9)) : 0;
  let h = '<span class="dot-grid">';
  for(let i=0;i<9;i++){
    h += `<span class="dot${i<filled?' '+color:''}"></span>`;
  }
  h += '</span>';
  return h;
}
function getThumbUrl(cat, work){
  const m = IMAGE_MAP[cat];
  if(!m) return null;
  const fid = m[work];
  if(!fid) return null;
  return `https://drive.google.com/thumbnail?id=${fid}&sz=w200`;
}
function buildSearchHref(word){
  const port = document.getElementById('exportPort').value || '11059';
  const linkType = document.getElementById('exportLinkType').value;
  const path = linkType === 'ykr' ? 'search_listerdb_filelist.php?anyword=' : 'search.php?searchword=';
  return `http://ykr.moe:${port}/${path}${encodeURIComponent(word)}`;
}
function findHistoryMatches(workName, songName){
  const sn = jsNormalize(songName);
  const wn = jsNormalize(workName);
  if(!sn && !wn) return [];
  return HISTORY.filter(h=>{
    let songOk = false, workOk = false;
    if(sn){
      // 同じ単語境界判定（簡易: アルファベットなら境界、それ以外は包含）
      if(/^[A-Z0-9 ]+$/.test(sn)){
        const re = new RegExp('(?:^|[^A-Z0-9])' + sn.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '(?:[^A-Z0-9]|$)','i');
        songOk = re.test(h.sn);
      } else {
        songOk = h.sn.indexOf(sn) >= 0;
      }
    }
    if(wn){
      workOk = (h.sn.indexOf(wn) >= 0) || (h.wn.indexOf(wn) >= 0);
    }
    if(sn && wn) return songOk && workOk;
    if(sn) return songOk;
    if(wn) return workOk;
    return false;
  });
}

// ====== Modal: 「この曲を歌った人」 ======
function openSingersModal(workName, songName){
  const matches = findHistoryMatches(workName, songName);
  const userCounts = {};
  matches.forEach(m=>{
    if(!userCounts[m.u]) userCounts[m.u] = {count:0, last:m.d, room:m.rm};
    userCounts[m.u].count++;
    if(m.d > userCounts[m.u].last) userCounts[m.u].last = m.d;
  });
  const users = Object.entries(userCounts).sort((a,b)=>b[1].count - a[1].count);
  const ttl = `${escHtml(songName)}<small>${escHtml(workName)} - ${matches.length}件 / ${users.length}人</small>`;
  document.getElementById('modalTitle').innerHTML = ttl;
  let body = `<div class="modal-summary"><i class="fas fa-users"></i> ${users.length}人がこの曲を歌っています（合計${matches.length}回）</div>`;
  if(users.length === 0){
    body += '<div class="empty">履歴が見つかりませんでした</div>';
  } else {
    users.forEach(([u, info])=>{
      body += `<div class="modal-row">
        <div class="top">
          <div class="user"><i class="fas fa-microphone"></i> ${escHtml(u)} <span style="color:var(--accent);font-size:11.5px;margin-left:4px">×${info.count}</span></div>
          <div class="date">${escHtml(info.last)}</div>
        </div>
        <div class="meta">最新: ${escHtml(info.room)}</div>
      </div>`;
    });
  }
  document.getElementById('modalBody').innerHTML = body;
  document.getElementById('modalOverlay').classList.add('active');
}
function closeModal(){
  document.getElementById('modalOverlay').classList.remove('active');
}
document.getElementById('modalOverlay').addEventListener('click',e=>{
  if(e.target.id === 'modalOverlay') closeModal();
});

// ====== Tabs ======
document.querySelectorAll('.tab-btn').forEach(b=>{
  b.addEventListener('click', ()=>{
    document.querySelectorAll('.tab-btn').forEach(x=>x.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(x=>x.classList.remove('active'));
    b.classList.add('active');
    document.getElementById('tab-'+b.dataset.tab).classList.add('active');
  });
});

// ====== Setlist ======
const SETLIST = HISTORY.map((h,i)=>({...h, idx:i, _orderNum: parseFloat(h.o) || -Infinity}));
let slState = {sort:'date_desc', rooms:new Set(), keyword:''};

function renderSetlist(){
  let arr = SETLIST.slice();
  // 部屋フィルタ
  if(slState.rooms.size > 0){
    arr = arr.filter(x => slState.rooms.has(x.rm));
  }
  // キーワード
  if(slState.keyword){
    const kws = slState.keyword.toUpperCase().replace(/　/g,' ').split(/\s+/).filter(Boolean);
    arr = arr.filter(x=>{
      const t = (x.d+' '+x.rm+' '+x.o+' '+x.sg+' '+x.wk+' '+x.ar+' '+x.u).toUpperCase();
      return kws.every(k=>t.indexOf(k)>=0);
    });
  }
  // ソート
  arr.sort((a,b)=>{
    if(slState.sort==='date_desc'){
      if(a.d!==b.d) return a.d<b.d?1:-1;
      return b._orderNum - a._orderNum;
    } else if(slState.sort==='date_asc'){
      if(a.d!==b.d) return a.d<b.d?-1:1;
      return a._orderNum - b._orderNum;
    } else if(slState.sort==='date_desc_order_asc'){
      if(a.d!==b.d) return a.d<b.d?1:-1;
      return a._orderNum - b._orderNum;
    }
    return 0;
  });

  document.getElementById('slCount').innerText = arr.length.toLocaleString();
  const list = document.getElementById('slList');
  if(arr.length === 0){
    list.innerHTML = '<div class="empty">該当データがありません</div>';
    return;
  }
  // 描画上限（パフォーマンス）
  const cap = 1000;
  const slice = arr.slice(0, cap);
  let html = '';
  slice.forEach((x,i)=>{
    const orderDisp = x.o || '';
    html += `<div class="card" data-cardidx="${x.idx}">
      <div class="card-head" onclick="toggleCard(this)">
        <div class="num-badge">${escHtml(orderDisp)}</div>
        <div class="card-body">
          <div class="sl-top">
            <span class="room-tag">${escHtml(x.rm)}</span>
            <span class="sl-date">${escHtml(x.d)}</span>
          </div>
          <div class="sl-song">${escHtml(x.sg)}</div>
          <div class="sl-meta">${escHtml(x.wk)}${x.wk&&x.ar?' <span class="sep">|</span> ':''}${escHtml(x.ar)}</div>
        </div>
        <i class="fas fa-chevron-down card-chev"></i>
      </div>
      <div class="card-detail">
        <div class="detail-table">
          <div class="detail-row"><div class="lbl"><i class="fas fa-music"></i>曲名</div><div class="val">${escHtml(x.sg)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-film"></i>作品名</div><div class="val">${escHtml(x.wk)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-user"></i>歌手名</div><div class="val">${escHtml(x.ar)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-microphone"></i>歌った人</div><div class="val">${escHtml(x.u)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-door-open"></i>部屋</div><div class="val">${escHtml(x.rm)}</div></div>
        </div>
        <button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(x.wk)}','${escAttr(x.sg)}')">
          <i class="fas fa-microphone"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i>
        </button>
      </div>
    </div>`;
  });
  if(arr.length > cap){
    html += `<div class="empty">（表示上限 ${cap} 件 / 全 ${arr.length} 件）絞り込みで全件確認できます</div>`;
  }
  list.innerHTML = html;
}
function escAttr(s){
  return String(s||"").replace(/'/g,"&#39;").replace(/"/g,'&quot;');
}
function toggleCard(headEl){
  headEl.parentElement.classList.toggle('expanded');
}

// Setlist filter popover
const slPopover = document.getElementById('slPopover');
const slFilterBtn = document.getElementById('slFilterBtn');
slFilterBtn.addEventListener('click',e=>{
  e.stopPropagation();
  slPopover.classList.toggle('open');
  slFilterBtn.classList.toggle('active');
});
document.addEventListener('click',e=>{
  if(!slPopover.contains(e.target) && e.target!==slFilterBtn){
    slPopover.classList.remove('open');
    slFilterBtn.classList.remove('active');
  }
});
slPopover.querySelectorAll('.opt').forEach(o=>{
  o.addEventListener('click',()=>{
    slPopover.querySelectorAll('.opt').forEach(x=>x.classList.remove('selected'));
    o.classList.add('selected');
    slState.sort = o.dataset.sort;
  });
});
// Build room chips
const slRoomChips = document.getElementById('slRoomChips');
ROOMS.forEach(r=>{
  const c = document.createElement('span');
  c.className = 'room-chip';
  c.innerText = r;
  c.addEventListener('click',()=>{
    if(slState.rooms.has(r)){slState.rooms.delete(r); c.classList.remove('selected');}
    else {slState.rooms.add(r); c.classList.add('selected');}
  });
  slRoomChips.appendChild(c);
});
document.getElementById('slClearFilter').addEventListener('click',()=>{
  slState.rooms.clear();
  slRoomChips.querySelectorAll('.room-chip').forEach(c=>c.classList.remove('selected'));
  slPopover.querySelectorAll('.opt').forEach(x=>x.classList.remove('selected'));
  slPopover.querySelector('[data-sort="date_desc"]').classList.add('selected');
  slState.sort = 'date_desc';
  renderSetlist();
});
document.getElementById('slApplyFilter').addEventListener('click',()=>{
  slPopover.classList.remove('open');
  slFilterBtn.classList.remove('active');
  renderSetlist();
});
document.getElementById('slSearch').addEventListener('input',e=>{
  slState.keyword = e.target.value.trim();
  renderSetlist();
});

// ====== Cool ======
const coolCatSel = document.getElementById('coolCat');
const coolSortSel = document.getElementById('coolSort');
CATS.forEach(c=>{
  const o = document.createElement('option');o.value=c;o.innerText=c;coolCatSel.appendChild(o);
});
coolCatSel.addEventListener('change', renderCool);
coolSortSel.addEventListener('change', renderCool);

function renderCool(){
  const cat = coolCatSel.value;
  const sort = coolSortSel.value;
  const data = COOL_DATA[cat] || {works:[], max_count:0, max_user:0};
  const works = data.works.slice();
  if(sort === 'count') works.sort((a,b)=>b.total_count-a.total_count || b.total_user-a.total_user);
  else if(sort === 'user') works.sort((a,b)=>b.total_user-a.total_user || b.total_count-a.total_count);
  else if(sort === 'name') works.sort((a,b)=>a.anime.localeCompare(b.anime,'ja'));
  else if(sort === 'created') works.sort((a,b)=>{
    const ac = a.songs.reduce((s,x)=>s+(x.creation_count||0),0);
    const bc = b.songs.reduce((s,x)=>s+(x.creation_count||0),0);
    return bc - ac;
  });
  const maxC = data.max_count, maxU = data.max_user;
  document.getElementById('coolCount').innerText = works.length;
  const list = document.getElementById('coolList');
  if(works.length === 0){
    list.innerHTML = '<div class="empty">データがありません</div>';
    return;
  }
  let html = '';
  works.forEach((w,i)=>{
    const rank = i+1;
    const opTag = `<span class="type-pill${w.op_n>0?' has':''}">OP <b>${w.op_n}</b></span>`;
    const edTag = `<span class="type-pill${w.ed_n>0?' has':''}">ED <b>${w.ed_n}</b></span>`;
    const inTag = `<span class="type-pill${w.in_n>0?' has':''}">IN <b>${w.in_n}</b></span>`;
    let songsHtml = '';
    w.songs.forEach(s=>{
      const sm = `<div class="metric"><div class="lbl">人数</div>${dotGrid(s.user_count,maxU,'fu')}<div class="val">${s.user_count}</div></div>
                  <div class="metric"><div class="lbl">歌唱</div>${dotGrid(s.count,maxC,'fc')}<div class="val">${s.count}</div></div>`;
      const tagShort = s.type ? s.type.toUpperCase().replace(/[^A-Z]/g,'').slice(0,2) || s.type : '';
      songsHtml += `<div class="song-row">
        <div class="song-type">${escHtml(tagShort||'-')}</div>
        <div class="song-info">
          <div class="song-name">${escHtml(s.song)}</div>
          <div class="song-artist">${escHtml(s.artist)}</div>
        </div>
        <div class="song-metrics">${sm}</div>
      </div>
      <div class="song-confirm">
        <button class="confirm-btn" style="margin:0;width:100%" onclick="event.stopPropagation();openSingersModal('${escAttr(w.anime)}','${escAttr(s.song)}')">
          <i class="fas fa-microphone"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i>
        </button>
      </div>`;
    });
    html += `<div class="card">
      <div class="card-head" onclick="toggleCard(this)">
        <div class="num-badge${rank===1?' gold':rank===2?' silver':rank===3?' bronze':''}">${rank}</div>
        <div class="card-body">
          <div class="cool-head-row">
            <div class="cool-anime-block">
              <div class="cool-anime">${escHtml(w.anime)}</div>
              <div class="cool-types">${opTag}${edTag}${inTag}</div>
            </div>
            <div class="metric-block">
              <div class="metric"><div class="lbl">人数</div>${dotGrid(w.total_user,maxU,'fu')}<div class="val">${w.total_user}</div></div>
              <div class="metric"><div class="lbl">歌唱数</div>${dotGrid(w.total_count,maxC,'fc')}<div class="val">${w.total_count}</div></div>
            </div>
          </div>
        </div>
        <i class="fas fa-chevron-down card-chev"></i>
      </div>
      <div class="card-detail">${songsHtml}</div>
    </div>`;
  });
  list.innerHTML = html;
}

// ====== Ranking ======
const rankCatSel = document.getElementById('rankCat');
const rankModeSel = document.getElementById('rankMode');
CATS.forEach(c=>{
  const o = document.createElement('option');o.value=c;o.innerText=c;rankCatSel.appendChild(o);
});
rankCatSel.addEventListener('change', renderRanking);
rankModeSel.addEventListener('change', renderRanking);

function renderRanking(){
  const cat = rankCatSel.value;
  const mode = rankModeSel.value;
  const all = (RANK_DATA[cat]||[]).slice();
  if(mode === 'count') all.sort((a,b)=>b.count-a.count || b.user_count-a.user_count);
  else all.sort((a,b)=>b.user_count-a.user_count || b.count-a.count);

  const top20 = [];
  let prevVal = null, curRank = 0;
  for(let i=0;i<all.length;i++){
    const v = mode==='count' ? all[i].count : all[i].user_count;
    if(v !== prevVal){ curRank = i+1; prevVal = v; }
    if(curRank > 20) break;
    top20.push({...all[i], rank:curRank});
  }
  document.getElementById('rankCount').innerText = top20.length;

  const list = document.getElementById('rankList');
  if(top20.length === 0){
    list.innerHTML = '<div class="empty">ランキング対象データがありません</div>';
    return;
  }
  let html = '';
  top20.forEach((r,idx)=>{
    const isTop3 = r.rank <= 3;
    const grade = r.rank===1?'gold':r.rank===2?'silver':r.rank===3?'bronze':'';
    if(isTop3){
      html += `<div class="rank-card-top3 ${grade}">
        <div class="rank-top3-head" onclick="toggleCard(this)">
          <div style="display:flex;flex-direction:column;align-items:center">
            <div class="rank-crown ${grade}"><i class="fas fa-crown"></i></div>
            <div class="rank-top3-num ${grade}">${r.rank}</div>
          </div>
          <div class="rank-top3-info">
            <div class="rank-top3-song">${escHtml(r.song)}</div>
            <div class="rank-top3-anime">${escHtml(r.anime)}</div>
            <div class="rank-top3-artist">${escHtml(r.artist)}</div>
            <div class="rank-top3-types"><span class="type-pill has">${escHtml(r.type||'-')}</span></div>
          </div>
          <i class="fas fa-chevron-down card-chev"></i>
        </div>
        <div class="rank-top3-metrics">
          <div class="rank-top3-metric"><div class="ico user"><i class="fas fa-users"></i></div><div class="lbl">人数</div><div class="val">${r.user_count}</div></div>
          <div class="rank-top3-metric"><div class="ico song"><i class="fas fa-microphone"></i></div><div class="lbl">歌唱数</div><div class="val">${r.count}</div></div>
        </div>
        <div class="card-detail">
          <button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(r.anime)}','${escAttr(r.song)}')">
            <i class="fas fa-users"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i>
          </button>
        </div>
      </div>`;
    } else {
      html += `<div class="rank-card">
        <div class="rank-row" onclick="toggleCard(this)">
          <div class="num-badge">${r.rank}</div>
          <div class="rank-info">
            <div class="rank-anime">${escHtml(r.anime)}</div>
            <div class="rank-sub">${escHtml(r.song)} / ${escHtml(r.artist)}</div>
          </div>
          <div class="rank-mini-metric"><div class="lbl"><i class="fas fa-users user"></i>人数</div><div class="val">${r.user_count}</div></div>
          <div class="rank-mini-metric"><div class="lbl"><i class="fas fa-microphone song"></i>歌唱数</div><div class="val">${r.count}</div></div>
          <i class="fas fa-chevron-down card-chev"></i>
        </div>
        <div class="card-detail">
          <div class="detail-table">
            <div class="detail-row"><div class="lbl"><i class="fas fa-music"></i>曲名</div><div class="val">${escHtml(r.song)}</div></div>
            <div class="detail-row"><div class="lbl"><i class="fas fa-film"></i>作品名</div><div class="val">${escHtml(r.anime)}</div></div>
            <div class="detail-row"><div class="lbl"><i class="fas fa-user"></i>歌手</div><div class="val">${escHtml(r.artist)}</div></div>
            <div class="detail-row"><div class="lbl"><i class="fas fa-tag"></i>種別</div><div class="val">${escHtml(r.type)}</div></div>
          </div>
          <button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(r.anime)}','${escAttr(r.song)}')">
            <i class="fas fa-users"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i>
          </button>
        </div>
      </div>`;
    }
  });
  list.innerHTML = html;
}

// ====== Trend ======
const trendCatSel = document.getElementById('trendCat');
const trendPeriodSel = document.getElementById('trendPeriod');
const trendSortSel = document.getElementById('trendSort');
{
  const o = document.createElement('option');o.value=TREND_TARGET_CAT;o.innerText=TREND_TARGET_CAT;trendCatSel.appendChild(o);
}
TREND_PERIODS.forEach(p=>{
  const o = document.createElement('option');o.value=String(p);o.innerText=`直近${p}日`;
  if(p===7) o.selected = true;
  trendPeriodSel.appendChild(o);
});
trendPeriodSel.addEventListener('change', renderTrend);
trendSortSel.addEventListener('change', renderTrend);

function renderTrend(){
  const period = trendPeriodSel.value;
  const sort = trendSortSel.value;
  const td = TREND_DATA[period] || {kpi:{surge_count:0,new_in:0,max_delta:0},items:[]};
  const items = td.items.slice();
  if(sort==='surge') items.sort((a,b)=>b.delta-a.delta || b.cur_count-a.cur_count);
  else if(sort==='new') items.sort((a,b)=>(b.is_new?1:0)-(a.is_new?1:0) || b.delta-a.delta);
  else items.sort((a,b)=>b.delta-a.delta);

  let html = '';
  // KPI cards
  html += `<div class="trend-stats">
    <div class="trend-stat"><div class="ico up"><i class="fas fa-arrow-trend-up"></i></div><div class="lbl">今週急上昇</div><div class="val">${td.kpi.surge_count}<small>曲</small></div></div>
    <div class="trend-stat"><div class="ico new"><i class="far fa-star"></i></div><div class="lbl">新規ランクイン</div><div class="val">${td.kpi.new_in}<small>曲</small></div></div>
    <div class="trend-stat"><div class="ico max"><i class="fas fa-arrow-up-right-dots"></i></div><div class="lbl">最大伸び</div><div class="val">+${td.kpi.max_delta}</div></div>
  </div>`;

  // Pickup
  if(items.length > 0){
    const p = items[0];
    const thumbUrl = getThumbUrl(TREND_TARGET_CAT, p.anime);
    html += `<div class="trend-pickup">
      <i class="fas fa-fire trend-pickup-icon"></i>
      <div class="trend-pickup-head"><i class="fas fa-fire"></i> 急上昇ピックアップ</div>
      <div class="trend-pickup-body">
        <div class="thumb-square large">
          ${thumbUrl ? `<img src="${thumbUrl}" alt="">` : '<i class="far fa-image"></i>'}
          <span class="thumb-tag">急上昇 No.1</span>
        </div>
        <div class="tp-info">
          <div class="tp-anime">${escHtml(p.song)}</div>
          <div class="tp-anime" style="font-size:13px;color:var(--text-sub);font-weight:600">${escHtml(p.anime)}</div>
          <div class="tp-artist">${escHtml(p.artist)}</div>
          <div class="tp-type"><span class="type-pill has">${escHtml(p.type||'-')}</span></div>
        </div>
        <div class="tp-side">
          <div class="arrow"><i class="fas fa-arrow-up"></i></div>
          <div class="pre-label">前週比</div>
          <div class="delta">${p.delta>=0?'+':''}${p.delta}</div>
        </div>
      </div>
      <div class="tp-stats">
        <div class="tp-stat"><div class="lbl"><i class="fas fa-microphone"></i>歌唱数</div><div class="val">${p.cur_count}</div></div>
        <div class="tp-stat"><div class="lbl"><i class="fas fa-users"></i>人数</div><div class="val">${p.cur_user}</div></div>
        <div class="tp-stat"><div class="lbl"><i class="fas fa-crown"></i>現在</div><div class="val">${p.rank_now||'-'}<small>位</small></div></div>
      </div>
      <button class="confirm-btn" style="margin-top:12px" onclick="openSingersModal('${escAttr(p.anime)}','${escAttr(p.song)}')">
        <i class="fas fa-users"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i>
      </button>
    </div>`;
  }

  // 注目の上昇曲 (2位以下)
  if(items.length > 1){
    html += `<div class="notable-head"><i class="fas fa-arrow-trend-up"></i> 注目の上昇曲</div>`;
    items.slice(1, 21).forEach((it,idx)=>{
      const rank = idx + 2;
      const thumbUrl = getThumbUrl(TREND_TARGET_CAT, it.anime);
      const deltaCls = it.is_new ? 'new' : '';
      const deltaTxt = it.is_new ? 'NEW' : (it.delta>=0?'↑+'+it.delta:''+it.delta);
      html += `<div class="notable-card">
        <div class="notable-row" onclick="toggleCard(this)">
          <div class="notable-num">${rank}</div>
          <div class="thumb-mini">${thumbUrl?`<img src="${thumbUrl}" alt="">`:'<i class="far fa-image"></i>'}</div>
          <div class="notable-info">
            <div class="notable-anime">${escHtml(it.song)}</div>
            <div class="notable-artist">${escHtml(it.anime)} / ${escHtml(it.artist)}</div>
            <div class="rank-types-inline"><span class="type-pill has">${escHtml(it.type||'-')}</span></div>
          </div>
          <div class="notable-delta ${deltaCls}">${deltaTxt}</div>
          <div class="notable-mini-metric"><div class="lbl"><i class="fas fa-users"></i>人数</div><div class="val">${it.cur_user}</div></div>
          <div class="notable-mini-metric"><div class="lbl"><i class="fas fa-microphone"></i>歌唱数</div><div class="val">${it.cur_count}</div></div>
          <i class="fas fa-chevron-right notable-chev"></i>
        </div>
        <div class="card-detail">
          <div class="detail-table">
            <div class="detail-row"><div class="lbl"><i class="fas fa-music"></i>曲名</div><div class="val">${escHtml(it.song)}</div></div>
            <div class="detail-row"><div class="lbl"><i class="fas fa-film"></i>作品名</div><div class="val">${escHtml(it.anime)}</div></div>
            <div class="detail-row"><div class="lbl"><i class="fas fa-user"></i>歌手</div><div class="val">${escHtml(it.artist)}</div></div>
          </div>
          <button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(it.anime)}','${escAttr(it.song)}')">
            <i class="fas fa-users"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i>
          </button>
        </div>
      </div>`;
    });
  }

  if(items.length === 0){
    html += '<div class="empty">急上昇データがありません</div>';
  }

  document.getElementById('trendBody').innerHTML = html;
}

// ====== Env (image upload) ======
function renderEnv(){
  const list = document.getElementById('envList');
  let html = '';
  CATS.forEach(cat=>{
    const works = (COOL_DATA[cat]||{works:[]}).works;
    if(!works.length) return;
    html += `<div class="env-section"><h3>${escHtml(cat)}</h3>`;
    works.forEach(w=>{
      const thumbUrl = getThumbUrl(cat, w.anime);
      const safeId = 'env_'+btoa(unescape(encodeURIComponent(cat+'|'+w.anime))).replace(/[^a-zA-Z0-9]/g,'');
      html += `<div class="env-work-row">
        <div class="env-work-thumb" id="${safeId}_thumb">${thumbUrl?`<img src="${thumbUrl}" alt="">`:'<i class="far fa-image"></i>'}</div>
        <div class="env-work-name">${escHtml(w.anime)}</div>
        <label class="env-upload-btn"><i class="fas fa-upload"></i> アップロード
          <input type="file" accept="image/*" onchange="handleImageUpload(this,'${escAttr(cat)}','${escAttr(w.anime)}','${safeId}_thumb')">
        </label>
      </div>`;
    });
    html += '</div>';
  });
  list.innerHTML = html || '<div class="empty">対象作品がありません</div>';
}
function showEnvStatus(msg, isError){
  const el = document.getElementById('envStatus');
  el.innerText = msg;
  el.style.background = isError ? '#fef2f2' : '#eef2ff';
  el.style.color = isError ? '#ef4444' : 'var(--accent)';
  el.classList.add('show');
  setTimeout(()=>el.classList.remove('show'), 5000);
}
async function handleImageUpload(input, cat, work, thumbId){
  const file = input.files[0];
  if(!file) return;
  showEnvStatus('画像を処理中... ('+file.name+')', false);
  try {
    const dataUrl = await cropToSquareDataUrl(file, 512);
    const base64 = dataUrl.split(',')[1];
    const payload = {
      action: 'upload_image',
      quarter: cat,
      work: work,
      filename: file.name,
      content_base64: base64,
      mime: 'image/jpeg'
    };
    showEnvStatus('Driveへアップロード中...', false);
    const res = await fetch(GAS_URL, {
      method: 'POST',
      headers: {'Content-Type':'text/plain;charset=utf-8'},
      body: JSON.stringify(payload)
    });
    const text = await res.text();
    let result;
    try { result = JSON.parse(text); } catch(e){ result = {ok:false, msg:text}; }
    if(result.ok && result.fileId){
      if(!IMAGE_MAP[cat]) IMAGE_MAP[cat] = {};
      IMAGE_MAP[cat][work] = result.fileId;
      const t = document.getElementById(thumbId);
      if(t) t.innerHTML = `<img src="${dataUrl}" alt="">`;
      showEnvStatus('アップロード完了: '+work, false);
    } else {
      showEnvStatus('失敗: '+(result.msg||'unknown'), true);
    }
  } catch(e){
    showEnvStatus('エラー: '+e.message, true);
  }
  input.value = '';
}
function cropToSquareDataUrl(file, size){
  return new Promise((resolve,reject)=>{
    const reader = new FileReader();
    reader.onload = e=>{
      const img = new Image();
      img.onload = ()=>{
        const s = Math.min(img.width, img.height);
        const sx = (img.width - s)/2;
        const sy = (img.height - s)/2;
        const canvas = document.createElement('canvas');
        canvas.width = size; canvas.height = size;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, sx, sy, s, s, 0, 0, size, size);
        resolve(canvas.toDataURL('image/jpeg', 0.85));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ====== HTML download ======
function buildExportHtml(title, contentHtml, embedScript){
  const port = document.getElementById('exportPort').value || '11059';
  const linkType = document.getElementById('exportLinkType').value;
  return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>${escHtml(title)}</title>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
<style>${document.querySelector('style').innerHTML}</style>
</head><body>
<div class="app-header"><div class="brand">${escHtml(title)}</div></div>
<div class="update-line">${escHtml(UPDATE_TS)} 出力</div>
<div class="tab-content active" style="display:block">${contentHtml}</div>
<div class="modal-overlay" id="modalOverlay"><div class="modal"><div class="modal-head"><div class="ttl" id="modalTitle"></div><button class="modal-close" onclick="closeModal()"><i class="fas fa-times"></i></button></div><div class="modal-body" id="modalBody"></div></div></div>
<script>
const PORT='${port}';const LINKTYPE='${linkType}';
const HISTORY=${JSON.stringify(HISTORY)};
const IMAGE_MAP=${JSON.stringify(IMAGE_MAP)};
${embedScript}
<\/script>
</body></html>`;
}
function downloadFile(filename, content){
  const blob = new Blob([content],{type:'text/html;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}
function commonExportScript(){
  return `
function escHtml(s){return String(s||"").replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function escAttr(s){return String(s||"").replace(/'/g,"&#39;").replace(/"/g,'&quot;');}
function jsNormalize(s){if(!s)return"";s=String(s).normalize('NFKC');s=s.replace(/\\.[a-zA-Z0-9]{3,4}$/,'');s=s.replace(/[\\[\\(\\{【].*?[\\]\\)\\}】]/g,' ');s=s.replace(/(key|KEY)?\\s*[\\+\\-]\\s*[0-9]+/g,' ');s=s.replace(/原キー/g,' ');s=s.replace(/(キー)?変更[:：]?/g,' ');s=s.replace(/[~〜～\\-_=,.]/g,' ');s=s.replace(/\\s+/g,' ').trim();return s.toUpperCase();}
function findHistoryMatches(workName,songName){const sn=jsNormalize(songName);const wn=jsNormalize(workName);if(!sn&&!wn)return[];return HISTORY.filter(h=>{let songOk=false,workOk=false;if(sn){if(/^[A-Z0-9 ]+$/.test(sn)){const re=new RegExp('(?:^|[^A-Z0-9])'+sn.replace(/[.*+?^\${}()|[\\]\\\\]/g,'\\\\$&')+'(?:[^A-Z0-9]|$)','i');songOk=re.test(h.sn);}else{songOk=h.sn.indexOf(sn)>=0;}}if(wn){workOk=(h.sn.indexOf(wn)>=0)||(h.wn.indexOf(wn)>=0);}if(sn&&wn)return songOk&&workOk;if(sn)return songOk;if(wn)return workOk;return false;});}
function openSingersModal(w,s){const m=findHistoryMatches(w,s);const u={};m.forEach(x=>{if(!u[x.u])u[x.u]={c:0,d:x.d,r:x.rm};u[x.u].c++;if(x.d>u[x.u].d){u[x.u].d=x.d;u[x.u].r=x.rm;}});const us=Object.entries(u).sort((a,b)=>b[1].c-a[1].c);document.getElementById('modalTitle').innerHTML=escHtml(s)+'<small>'+escHtml(w)+' - '+m.length+'件 / '+us.length+'人</small>';let b='<div class="modal-summary"><i class="fas fa-users"></i> '+us.length+'人がこの曲を歌っています（合計'+m.length+'回）</div>';if(us.length===0)b+='<div class="empty">履歴が見つかりませんでした</div>';else us.forEach(([uu,info])=>{b+='<div class="modal-row"><div class="top"><div class="user"><i class="fas fa-microphone"></i> '+escHtml(uu)+' <span style="color:var(--accent);font-size:11.5px;margin-left:4px">×'+info.c+'</span></div><div class="date">'+escHtml(info.d)+'</div></div><div class="meta">最新: '+escHtml(info.r)+'</div></div>';});document.getElementById('modalBody').innerHTML=b;document.getElementById('modalOverlay').classList.add('active');}
function closeModal(){document.getElementById('modalOverlay').classList.remove('active');}
document.getElementById('modalOverlay').addEventListener('click',e=>{if(e.target.id==='modalOverlay')closeModal();});
function toggleCard(h){h.parentElement.classList.toggle('expanded');}
function getThumbUrl(c,w){const m=IMAGE_MAP[c];if(!m)return null;const f=m[w];if(!f)return null;return 'https://drive.google.com/thumbnail?id='+f+'&sz=w200';}
document.querySelectorAll('a.export-link').forEach(l=>{const h=l.getAttribute('href');if(h&&h.indexOf('#search:')===0){const w=h.split('#search:')[1];const path=LINKTYPE==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword=';l.href='http://ykr.moe:'+PORT+'/'+path+encodeURIComponent(w);l.target='_blank';l.rel='noopener';}});
`;
}
function dotGridStr(value, max, color){
  const filled = max>0 ? Math.min(9, Math.round((value/max)*9)) : 0;
  let h = '<span class="dot-grid">';
  for(let i=0;i<9;i++){h += `<span class="dot${i<filled?' '+color:''}"></span>`;}
  h += '</span>';
  return h;
}
function buildCoolHtmlForCat(cat){
  const data = COOL_DATA[cat] || {works:[],max_count:0,max_user:0};
  const works = data.works.slice().sort((a,b)=>b.total_count-a.total_count);
  const maxC = data.max_count, maxU = data.max_user;
  let h = `<div style="font-size:14px;font-weight:700;color:var(--primary);padding:0 16px;border-left:3px solid var(--accent);margin:14px 0 8px">${escHtml(cat)} - ${works.length}作品</div>`;
  works.forEach((w,i)=>{
    const rank = i+1;
    const opTag = `<span class="type-pill${w.op_n>0?' has':''}">OP <b>${w.op_n}</b></span>`;
    const edTag = `<span class="type-pill${w.ed_n>0?' has':''}">ED <b>${w.ed_n}</b></span>`;
    const inTag = `<span class="type-pill${w.in_n>0?' has':''}">IN <b>${w.in_n}</b></span>`;
    let sh = '';
    w.songs.forEach(s=>{
      const sm = `<div class="metric"><div class="lbl">人数</div>${dotGridStr(s.user_count,maxU,'fu')}<div class="val">${s.user_count}</div></div><div class="metric"><div class="lbl">歌唱</div>${dotGridStr(s.count,maxC,'fc')}<div class="val">${s.count}</div></div>`;
      const tagShort = s.type ? s.type.toUpperCase().replace(/[^A-Z]/g,'').slice(0,2) || s.type : '';
      const linkSong = `<a class="export-link" href="#search:${escAttr(w.anime+' '+s.song)}">${escHtml(s.song)}</a>`;
      sh += `<div class="song-row"><div class="song-type">${escHtml(tagShort||'-')}</div><div class="song-info"><div class="song-name">${linkSong}</div><div class="song-artist">${escHtml(s.artist)}</div></div><div class="song-metrics">${sm}</div></div><div class="song-confirm"><button class="confirm-btn" style="margin:0;width:100%" onclick="event.stopPropagation();openSingersModal('${escAttr(w.anime)}','${escAttr(s.song)}')"><i class="fas fa-microphone"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i></button></div>`;
    });
    h += `<div class="card"><div class="card-head" onclick="toggleCard(this)"><div class="num-badge${rank===1?' gold':rank===2?' silver':rank===3?' bronze':''}">${rank}</div><div class="card-body"><div class="cool-head-row"><div class="cool-anime-block"><div class="cool-anime">${escHtml(w.anime)}</div><div class="cool-types">${opTag}${edTag}${inTag}</div></div><div class="metric-block"><div class="metric"><div class="lbl">人数</div>${dotGridStr(w.total_user,maxU,'fu')}<div class="val">${w.total_user}</div></div><div class="metric"><div class="lbl">歌唱数</div>${dotGridStr(w.total_count,maxC,'fc')}<div class="val">${w.total_count}</div></div></div></div></div><i class="fas fa-chevron-down card-chev"></i></div><div class="card-detail">${sh}</div></div>`;
  });
  return h;
}
function downloadCoolHTML(scope){
  const cat = coolCatSel.value;
  let body = '<div style="max-width:760px;margin:0 auto;padding:0 16px 60px">';
  if(scope === 'current'){
    body += buildCoolHtmlForCat(cat);
  } else {
    CATS.forEach(c=> body += buildCoolHtmlForCat(c));
  }
  body += '</div>';
  const ttl = scope==='current' ? `クール集計 - ${cat}` : 'クール集計（全クール）';
  const html = buildExportHtml(ttl, body, commonExportScript());
  const fname = scope==='current' ? `karaoke_cool_${cat}.html` : 'karaoke_cool_all.html';
  downloadFile(fname, html);
}

function buildRankHtmlForCat(cat, mode){
  const all = (RANK_DATA[cat]||[]).slice();
  if(mode==='count') all.sort((a,b)=>b.count-a.count || b.user_count-a.user_count);
  else all.sort((a,b)=>b.user_count-a.user_count || b.count-a.count);
  const top20 = []; let prevVal=null,curRank=0;
  for(let i=0;i<all.length;i++){
    const v = mode==='count' ? all[i].count : all[i].user_count;
    if(v!==prevVal){curRank=i+1;prevVal=v;}
    if(curRank>20) break;
    top20.push({...all[i],rank:curRank});
  }
  const modeTitle = mode==='count'?'歌唱数ランキング':'歌唱人数ランキング';
  let h = `<div style="font-size:14px;font-weight:700;color:var(--primary);padding:0 16px;border-left:3px solid var(--accent);margin:14px 0 8px">${escHtml(cat)} ${modeTitle}</div>`;
  top20.forEach(r=>{
    const isTop3 = r.rank<=3;
    const grade = r.rank===1?'gold':r.rank===2?'silver':r.rank===3?'bronze':'';
    const linkSong = `<a class="export-link" href="#search:${escAttr(r.anime+' '+r.song)}">${escHtml(r.song)}</a>`;
    if(isTop3){
      h += `<div class="rank-card-top3 ${grade}"><div class="rank-top3-head" onclick="toggleCard(this)"><div style="display:flex;flex-direction:column;align-items:center"><div class="rank-crown ${grade}"><i class="fas fa-crown"></i></div><div class="rank-top3-num ${grade}">${r.rank}</div></div><div class="rank-top3-info"><div class="rank-top3-song">${linkSong}</div><div class="rank-top3-anime">${escHtml(r.anime)}</div><div class="rank-top3-artist">${escHtml(r.artist)}</div><div class="rank-top3-types"><span class="type-pill has">${escHtml(r.type||'-')}</span></div></div><i class="fas fa-chevron-down card-chev"></i></div><div class="rank-top3-metrics"><div class="rank-top3-metric"><div class="ico user"><i class="fas fa-users"></i></div><div class="lbl">人数</div><div class="val">${r.user_count}</div></div><div class="rank-top3-metric"><div class="ico song"><i class="fas fa-microphone"></i></div><div class="lbl">歌唱数</div><div class="val">${r.count}</div></div></div><div class="card-detail"><button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(r.anime)}','${escAttr(r.song)}')"><i class="fas fa-users"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i></button></div></div>`;
    } else {
      h += `<div class="rank-card"><div class="rank-row" onclick="toggleCard(this)"><div class="num-badge">${r.rank}</div><div class="rank-info"><div class="rank-anime">${escHtml(r.anime)}</div><div class="rank-sub">${linkSong} / ${escHtml(r.artist)}</div></div><div class="rank-mini-metric"><div class="lbl"><i class="fas fa-users user"></i>人数</div><div class="val">${r.user_count}</div></div><div class="rank-mini-metric"><div class="lbl"><i class="fas fa-microphone song"></i>歌唱数</div><div class="val">${r.count}</div></div><i class="fas fa-chevron-down card-chev"></i></div><div class="card-detail"><button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(r.anime)}','${escAttr(r.song)}')"><i class="fas fa-users"></i> この曲を歌った人を確認する <i class="fas fa-chevron-right"></i></button></div></div>`;
    }
  });
  return h;
}
function downloadRankingHTML(scope){
  const cat = rankCatSel.value;
  const mode = rankModeSel.value;
  let body = '<div style="max-width:760px;margin:0 auto;padding:0 16px 60px">';
  if(scope==='current'){
    body += buildRankHtmlForCat(cat, mode);
  } else {
    CATS.forEach(c=>{
      body += buildRankHtmlForCat(c, mode);
    });
  }
  body += '</div>';
  const ttl = scope==='current' ? `ランキング - ${cat}` : 'ランキング（全クール）';
  const html = buildExportHtml(ttl, body, commonExportScript());
  const fname = scope==='current' ? `karaoke_rank_${mode}_${cat}.html` : `karaoke_rank_${mode}_all.html`;
  downloadFile(fname, html);
}

function downloadTrendHTML(){
  const period = trendPeriodSel.value;
  const sort = trendSortSel.value;
  // 直接DOM cloneでOK
  const node = document.getElementById('trendBody').cloneNode(true);
  const body = `<div style="max-width:760px;margin:0 auto;padding:0 16px 60px"><div style="font-size:14px;font-weight:700;color:var(--primary);padding:0 16px;border-left:3px solid var(--accent);margin:14px 0 8px">推移 - ${TREND_TARGET_CAT} (直近${period}日)</div>${node.outerHTML}</div>`;
  const html = buildExportHtml(`推移 - ${TREND_TARGET_CAT}`, body, commonExportScript());
  downloadFile(`karaoke_trend_${period}d.html`, html);
}

function downloadSetlistHTML(){
  // 表示中のリストをそのまま出力
  const node = document.getElementById('slList').cloneNode(true);
  const body = `<div style="max-width:760px;margin:0 auto;padding:0 16px 60px"><div style="font-size:14px;font-weight:700;color:var(--primary);padding:0 16px;border-left:3px solid var(--accent);margin:14px 0 8px">セットリスト</div>${node.outerHTML}</div>`;
  const html = buildExportHtml('セットリスト', body, commonExportScript());
  downloadFile('karaoke_setlist.html', html);
}

// ====== Init ======
renderSetlist();
renderCool();
renderRanking();
renderTrend();
renderEnv();
</script>
</body>
</html>
"""

# プレースホルダ置換
html_content = (html_content
    .replace("__COOL_JSON__", cool_json)
    .replace("__RANK_JSON__", ranking_json)
    .replace("__TREND_JSON__", trend_json)
    .replace("__HISTORY_JSON__", history_json)
    .replace("__IMAGE_MAP_JSON__", image_map_json)
    .replace("__ROOMS_JSON__", rooms_json)
    .replace("__CATS_JSON__", categories_json)
    .replace("__TREND_PERIODS_JSON__", trend_periods_json)
    .replace("__TREND_TARGET_JSON__", trend_target_json)
    .replace("__GAS_URL__", GAS_WEB_APP_URL)
    .replace("__UPDATE_TS__", current_datetime_str)
    .replace("__UPDATE__", current_datetime_str)
    .replace("__CURRENT_DATE__", current_date_str)
)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("HTML生成完了: index.html")
