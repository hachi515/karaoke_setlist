import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
import json
import io
from itertools import groupby
from concurrent.futures import ThreadPoolExecutor, as_completed

# ==========================================
# ★ 設定: GitHubリポジトリ情報
# ==========================================
# ここにファイルをアップロードしたGitHubのアカウント名とリポジトリ名を入力してください
GITHUB_USER = "hachi515"  # 例: "kenshin"
GITHUB_REPO = "karaoke_setlist"   # 例: "karaoke_list"
GITHUB_BRANCH = "main"                 # 通常は "main" または "master"

# 読み込むオフラインリストのファイル名（新しい順に並べるのを推奨）
OFFLINE_FILES = [
    "offline_list_2026_1st.csv",
    "offline_list_2025_2nd.csv",
    "offline_list_2025_1st.csv"
]

# ==========================================
# ★ GAS連携・GitHub読み込み用設定
# ==========================================
# デプロイしたGASのウェブアプリURL (履歴保存やCool解析の取得に使用)
GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyzKEPfj0bYcRyEdizwQXcduIOQFt2_njtFQSyGP9jBjrhR8pyVKwDol6VN7bLPrktq/exec"
CSV_EMPTY_PREFIX_BYTES = b'\xef\xbb\xbf\r\n\t '  # BOMと空白のみのCSVレスポンス判定に使う
EXPECTED_HISTORY_COLUMNS = [
    '取得日', '部屋主', '順番', '曲名（ファイル名）', '作品名', '歌手名', '歌った人'
]

def load_df_from_github(filename, **kwargs):
    """GitHubのRawデータからCSVを読み込む"""
    url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{filename}"
    print(f"[GitHub] Loading {filename} from {url}...")

    try:
        response = requests.get(url, timeout=30)

        if response.status_code == 200:
            content_bytes = response.content
            encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']

            for enc in encodings:
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
             print(f"[GitHub] File not found: {filename} (404). Check user/repo/filename.")
             return pd.DataFrame()
        else:
            print(f"[GitHub] Error fetching {filename}: Status {response.status_code}")
            return pd.DataFrame()

    except Exception as e:
        print(f"[GitHub] Connection error for {filename}: {e}")
        return pd.DataFrame()


def load_df_from_gas_with_status(filename, **kwargs):
    """GASからCSVデータをダウンロードしてDataFrameと読み込み状態を返す"""
    print(f"[GAS] Loading {filename}...")
    try:
        response = requests.get(GAS_WEB_APP_URL, params={'filename': filename}, timeout=60)
    except Exception as e:
        print(f"[GAS] Connection error: {e}")
        return pd.DataFrame(), "error"

    if response.status_code == 404:
        print(f"[GAS] File not found: {filename}")
        return pd.DataFrame(), "not_found"

    if response.status_code != 200:
        print(f"[GAS] Error: {response.status_code}")
        return pd.DataFrame(), "error"

    content_bytes = response.content
    response_text = response.text if isinstance(response.text, str) else ""
    if "Exception: Service error: Drive" in response_text:
        print(f"[GAS] Drive service error detected for {filename}.")
        return pd.DataFrame(), "error"

    if not content_bytes or content_bytes.lstrip(CSV_EMPTY_PREFIX_BYTES) == b'':
        print(f"[GAS] Empty file: {filename}")
        return pd.DataFrame(), "empty"

    encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']

    for enc in encodings:
        try:
            df = pd.read_csv(io.BytesIO(content_bytes), encoding=enc, engine='python', **kwargs)
            if len(df.columns) > 0:
                df.columns = df.columns.astype(str).str.replace('\ufeff', '').str.strip()
            print(f"[GAS] Success: Loaded {filename} ({enc}). Rows: {len(df)}")
            return df, "ok"
        except Exception:
            continue

    print(f"[GAS] Failed to decode {filename}.")
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
    11012: "加古部屋",
    11021: "成田部屋",
    11022: "成田部屋",
    11028: "タマ部屋",
    11058: "すみた部屋",
    11059: "つぼはち部屋",
    11060: "れん部屋",
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
    11087: "MiO部屋",
    11088: "ほっしー部屋",
    11091: "千秋部屋",
    11092: "ヒロ部屋",
    11101: "えみち部屋",
    11102: "るえ部屋",
    11103: "ながし部屋",
    11104: "MrN部屋",
    11105: "ヤマテル部屋",
    11106: "冨塚部屋",
    11107: "ブルーベリー部屋",
    11110: "加古部屋",
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


# --- 1. 過去データ読み込み・安全な履歴ローテーション ---
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

    bad_col_patterns = [
        r'^Error:',
        r'Exception:\s*Service error:\s*Drive'
    ]
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
        print(f"[Guard] {filename}: 保存対象が None のため中止します。")
        return False

    df = cleanup_history_df(df)

    if df.empty and not allow_empty:
        print(f"[Guard] {filename}: 空データでの上書きを拒否しました。")
        return False

    if len(df) < min_existing_rows:
        print(f"[Guard] {filename}: 行数が {min_existing_rows} -> {len(df)} に減るため保存を拒否しました。")
        return False

    if not save_df_to_gas(filename, df):
        return False

    verify_df, verify_status = load_df_from_gas_with_status(filename)
    if verify_status != "ok":
        print(f"[Guard] {filename}: 保存後の再読込に失敗しました。")
        return False

    verify_df = cleanup_history_df(verify_df)
    if len(verify_df) < len(df):
        print(f"[Guard] {filename}: 保存後の行数が不足しています。expected={len(df)}, actual={len(verify_df)}")
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
            print(f"[STOP] {filename} の読み込みに失敗したため、履歴更新を禁止します。")
            return histories, loaded_files, False

        num += 1

    return histories, loaded_files, True


def fetch_room_df(port):
    url = f"http://ykr.moe:{port}/simplelist.php"
    response = requests.get(url, timeout=ROOM_FETCH_TIMEOUT)
    response.raise_for_status()

    dfs = pd.read_html(io.BytesIO(response.content))
    if not dfs:
        raise ValueError("HTML内にテーブルが見つかりません")

    df = dfs[0].fillna("")
    if df.empty:
        raise ValueError("取得テーブルが空です")

    required_cols = ['順番', '曲名（ファイル名）']
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        raise ValueError(f"必要カラム不足: {missing_cols} / columns={list(df.columns)}")

    df = df.replace(r'\s*詳細を見る ▼', '', regex=True)
    df['部屋主'] = room_map[port]
    df['取得日'] = current_date_str
    return df


# --- 既存履歴の読み込み ---
history_records, loaded_history_files, history_load_ok = load_all_history_files()
if loaded_history_files:
    print(f"履歴ファイルを読み込みました: {', '.join(loaded_history_files)}")
else:
    print("履歴ファイルなし。取得成功時のみ history.csv を新規作成します。")

history_dfs = [h['df'] for h in history_records]
full_history_before_update = cleanup_history_df(pd.concat(history_dfs, ignore_index=True)) if history_dfs else pd.DataFrame()

final_df = history_records[-1]['df'] if history_records else pd.DataFrame()
archive_dfs = [h['df'] for h in history_records[:-1]] if len(history_records) > 1 else []

# --- 2. 新しいデータ取得 ---
target_ports = list(room_map.keys())
new_data_frames = []
failed_ports = []
fetched_ports = []
fetch_status = {}

if not history_load_ok:
    print("[Guard] 履歴の読み込みが不完全です。収集は実行しますが、履歴ファイルへの保存は禁止します。")

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
            fetch_status[port] = {"room": room_name, "status": "ok", "rows": len(df)}
            print(f"[Fetch] OK {port} ({room_name}) rows={len(df)}")
        except Exception as e:
            failed_ports.append((port, str(e)))
            fetch_status[port] = {"room": room_name, "status": "error", "error": str(e)}
            print(f"[Fetch] SKIP {port} ({room_name}) 取得失敗: {e}")

success_ports_count = len(fetched_ports)
print(f"[Fetch] 成功ポート: {success_ports_count} / {len(target_ports)}")

if failed_ports:
    print("[Fetch] 取得できなかったポートは無視します: " + ", ".join([f"{p}" for p, _ in failed_ports]))

if not new_data_frames:
    print("[Fetch] 取得成功データがありません。履歴ファイルは更新しません。")
    full_df = full_history_before_update

else:
    new_df = cleanup_history_df(pd.concat(new_data_frames, ignore_index=True))

    dedup_cols = [c for c in HISTORY_DEDUP_COLS if c in new_df.columns]
    if dedup_cols:
        before_rows = len(new_df)
        new_df = new_df.drop_duplicates(subset=dedup_cols, keep='last')
        new_df = cleanup_history_df(new_df)
        print(f"[Dedup] 今回取得分の重複除去: {before_rows} -> {len(new_df)} 行")

    print(f"[Fetch] 今回収集できた行数: {len(new_df)} 行")

    if not history_load_ok:
        print("[Guard] 履歴読み込み不完全のため、今回取得分は表示用のみ使用し、history 系CSVには保存しません。")
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
            print("追加対象の新規履歴はありません。履歴ファイルは更新しません。")
            full_df = full_history_before_update
        else:
            print(f"追加対象の新規履歴: {len(new_unique_df)} 行")

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
                    print(f"[History] {active_filename} に {len(append_part)} 行追記しました。現在 {len(next_df)} 行。")
                    saved_parts[active_num] = next_df
                    remaining_df = remaining_df.iloc[remaining_capacity:].reset_index(drop=True)

                    if not remaining_df.empty:
                        active_num += 1
                        active_df = pd.DataFrame()
                    else:
                        active_df = next_df
                else:
                    print(f"[STOP] {active_filename} の保存確認に失敗したため、以降の保存を停止します。既存履歴は削除していません。")
                    save_ok = False
                    break

            merged_history_by_num = {h['num']: h['df'] for h in history_records}
            merged_history_by_num.update(saved_parts)
            full_df = cleanup_history_df(pd.concat([merged_history_by_num[n] for n in sorted(merged_history_by_num)], ignore_index=True)) if merged_history_by_num else pd.DataFrame()

            if save_ok:
                print("履歴ファイルの安全更新が完了しました。")
            else:
                print("履歴ファイルの一部更新で停止しました。既存履歴は削除していません。")

if full_df is None or full_df.empty:
    full_df = pd.DataFrame()
else:
    full_df = cleanup_history_df(full_df)

print(f"全履歴データ合計: {len(full_df)} 行")

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

ALLOWED_CATEGORIES = ["2026年春アニメ", "2026年冬アニメ", "2025年秋アニメ"]

category_options = '<option value="ALL">すべて保存</option>\n'
for cat in ALLOWED_CATEGORIES:
    category_options += f'<option value="{cat}">{cat}</option>\n'

# --- オフラインリスト読み込み (GitHubから) ---
offline_targets = []

print(f"GitHubからオフラインリストを読み込みます... (User: {GITHUB_USER}, Repo: {GITHUB_REPO})")

for filename in OFFLINE_FILES:
    offline_df = load_df_from_github(filename)
    if not offline_df.empty:
        offline_df = offline_df.fillna("")
        if '曲名' in offline_df.columns:
            targets = [normalize_offline_text(str(x)) for x in offline_df['曲名'].tolist()]
            offline_targets.extend(targets)
            print(f"  -> {filename}: {len(targets)} 件追加")
        else:
            print(f"  -> {filename}: '曲名'カラムが見つかりません。")
    else:
        print(f"  -> {filename}: 読み込み失敗または空です。")

print(f"オフラインリスト合計件数: {len(offline_targets)}")


def is_song_created(song_name):
    if not song_name:
        return False
    norm = normalize_text(str(song_name))
    norm_raw = normalize_offline_text(str(song_name))
    if not norm and not norm_raw:
        return False
    for offline_str in offline_targets:
        if not offline_str:
            continue
        if (norm and norm in offline_str) or (norm_raw and norm_raw in offline_str):
            return True
    return False


# --- ★関数: 相対バー幅の計算（最大値基準） ---
def calc_bar_pct(value, max_value, min_pct=4):
    """
    相対スケーリング:
    - max_value を基準に 0〜100% で表現
    - 0 のときは 0%
    - 1以上だが極端に小さい時も視認できるよう min_pct を下限とする
    """
    try:
        v = float(value)
    except Exception:
        v = 0
    try:
        m = float(max_value)
    except Exception:
        m = 0
    if v <= 0 or m <= 0:
        return 0
    pct = (v / m) * 100.0
    if pct < min_pct:
        pct = min_pct
    if pct > 100:
        pct = 100
    return int(round(pct))


# --- ★関数: カテゴリ別リストHTML生成 (フラット 1曲1行) ---
def generate_category_html_block(category_name, item_list):
    if not item_list:
        return ""

    item_list.sort(key=lambda x: x['anime'])

    html = f"""
    <section class="cat-block">
        <header class="cat-header" onclick="toggleCategory(this)">
            <span class="cat-title">{category_name}</span>
            <span class="cat-count">{len(item_list)} 曲</span>
            <i class="fas fa-chevron-down cat-chev"></i>
        </header>
        <div class="cat-content">
        <table class="flatTable">
            <thead>
                <tr>
                    <th class="th-anime">作品名</th>
                    <th class="th-type">種別</th>
                    <th class="th-artist">歌手</th>
                    <th class="th-song">曲名</th>
                </tr>
            </thead>
            <tbody>
    """

    for item in item_list:
        clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
        search_word = f"{clean_anime} {item['song']}"
        link_start = f'<a href="#search_link/{search_word}" class="export-link">'

        html += f'''<tr>
            <td class="td-anime">{item["anime"]}</td>
            <td class="td-type">{link_start}<span class="type-tag">{item["type"]}</span></a></td>
            <td class="td-artist">{link_start}{item["artist"]}</a></td>
            <td class="td-song">{link_start}{item["song"]}</a></td>
        </tr>'''

    html += "</tbody></table></div></section>"
    return html


# --- Cool Analysis読み込み ---
raw_df = load_df_from_gas(cool_file, header=None)

if not raw_df.empty:
    try:
        raw_df = raw_df.fillna("")
        raw_df = raw_df.drop_duplicates(keep='last')

        analysis_source_df = full_df.copy()
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
        ].sort_values('dt_obj')

        start_date = pd.to_datetime("2026/01/01")
        end_date = pd.to_datetime("2026/06/30")
        target_history = full_history[
            (full_history['dt_obj'] >= start_date) &
            (full_history['dt_obj'] <= end_date)
        ]

        categorized_data = {}
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

        print("グラフデータ計算中...")
        graph_target_cat = "2026年春アニメ"

        if graph_target_cat in categorized_data:
            winter_items = categorized_data[graph_target_cat]

            items_with_norm = []
            for item in winter_items:
                items_with_norm.append({
                    "meta": item,
                    "song_norm": normalize_text(item["song"]),
                    "anime_norm": normalize_text(item["anime"]),
                    "name": f"{item['anime']} {item['song']}"
                })

            matched_records = []
            for idx, item in enumerate(items_with_norm):
                song_pat = item["song_norm"]
                anime_pat = item["anime_norm"]
                if not song_pat and not anime_pat: continue

                song_match = check_match(song_pat, full_history['norm_filename'])
                mask = None
                if song_pat and anime_pat:
                    anime_match = (
                        full_history['norm_filename'].str.contains(re.escape(anime_pat), case=False, na=False) |
                        full_history['norm_workname'].str.contains(re.escape(anime_pat), case=False, na=False)
                    )
                    mask = song_match & anime_match
                elif song_pat:
                    mask = song_match
                elif anime_pat:
                    mask = (
                        full_history['norm_filename'].str.contains(re.escape(anime_pat), case=False, na=False) |
                        full_history['norm_workname'].str.contains(re.escape(anime_pat), case=False, na=False)
                    )

                if mask is not None:
                    for _, row in full_history[mask].iterrows():
                        matched_records.append({
                            "date": row['dt_obj'],
                            "item_idx": idx,
                            "user": row['歌った人']
                        })

            matched_records.sort(key=lambda x: x['date'])

            if matched_records:
                unique_dates = sorted(list(set(r['date'] for r in matched_records)))
                current_counts = {}
                current_users = {}
                rec_ptr = 0
                total_recs = len(matched_records)

                for current_dt in unique_dates:
                    dt_str = current_dt.strftime("%Y-%m-%d")

                    while rec_ptr < total_recs and matched_records[rec_ptr]['date'] <= current_dt:
                        rec = matched_records[rec_ptr]
                        idx = rec['item_idx']
                        user = rec['user']

                        current_counts[idx] = current_counts.get(idx, 0) + 1
                        if idx not in current_users:
                            current_users[idx] = set()
                        current_users[idx].add(user)
                        rec_ptr += 1

                    ranking_src_count = []
                    for idx, cnt in current_counts.items():
                        ranking_src_count.append({"name": items_with_norm[idx]["name"], "val": cnt})
                    ranking_src_count.sort(key=lambda x: x['val'], reverse=True)

                    rank = 1
                    prev_val = -1
                    for i, d in enumerate(ranking_src_count):
                        if i > 0 and d['val'] < prev_val: rank = i + 1
                        prev_val = d['val']
                        if rank <= 20:
                            if d['name'] not in graph_series_data_count: graph_series_data_count[d['name']] = []
                            graph_series_data_count[d['name']].append({"x": dt_str, "y": rank})

                    ranking_src_user = []
                    for idx, u_set in current_users.items():
                        if len(u_set) > 0:
                            ranking_src_user.append({"name": items_with_norm[idx]["name"], "val": len(u_set)})
                    ranking_src_user.sort(key=lambda x: x['val'], reverse=True)

                    rank = 1
                    prev_val = -1
                    for i, d in enumerate(ranking_src_user):
                        if i > 0 and d['val'] < prev_val: rank = i + 1
                        prev_val = d['val']
                        if rank <= 20:
                            if d['name'] not in graph_series_data_user: graph_series_data_user[d['name']] = []
                            graph_series_data_user[d['name']].append({"x": dt_str, "y": rank})

        print("グラフデータ計算完了。")

        # =========================================================
        # クール集計テーブル生成
        # 2パス方式:
        #  Pass1: 各曲の歌唱数 / 歌唱人数 / 作成カウントを算出し、
        #         カテゴリ内 max を計算（バー相対スケーリング用）。
        #  Pass2: 作品名（anime）でグループ化し rowspan 集約。
        #         同時にスマホ用「作品アコーディオン」用の構造を出力。
        # =========================================================
        for category, items in categorized_data.items():

            cat_created_items = []
            cat_uncreated_items = []
            rendered_items = []  # Pass1 の結果（item に count/user_count/creation_count を付与）

            items.sort(key=lambda x: x['anime'])

            for item in items:
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

                matched_data = target_history[final_mask]
                count = len(matched_data)
                user_count = matched_data['歌った人'].nunique() if count > 0 else 0

                creation_count = 0
                target_song_raw_norm = normalize_offline_text(item["song"])

                if target_song_norm:
                    for offline_str in offline_targets:
                        if (target_song_norm in offline_str) or (target_song_raw_norm in offline_str):
                            if target_anime_norm:
                                if target_anime_norm in offline_str:
                                    creation_count += 1
                            else:
                                creation_count += 1

                if creation_count >= 1:
                    cat_created_items.append(item)
                else:
                    cat_uncreated_items.append(item)

                ranking_data_list.append({
                    "category": category,
                    "anime": item["anime"],
                    "song": item["song"],
                    "artist": item["artist"],
                    "type": item["type"],
                    "count": count,
                    "user_count": user_count
                })

                rendered_items.append({
                    **item,
                    "count": count,
                    "user_count": user_count,
                    "creation_count": creation_count
                })

            # 相対バー基準: 同一カテゴリ内の最大値で正規化
            max_count_in_cat = max([r['count'] for r in rendered_items], default=0)
            max_user_in_cat = max([r['user_count'] for r in rendered_items], default=0)

            analysis_html_content += f"""
            <section class="cat-block">
                <header class="cat-header" onclick="toggleCategory(this)">
                    <span class="cat-title">{category}</span>
                    <span class="cat-count">{len(rendered_items)} 曲</span>
                    <span class="cat-meta">最大歌唱数 {max_count_in_cat} / 最大人数 {max_user_in_cat}</span>
                    <i class="fas fa-chevron-down cat-chev"></i>
                </header>
                <div class="cat-content">
                <table class="analysisTable">
                    <thead>
                        <tr>
                            <th class="th-anime">作品名</th>
                            <th class="th-made">作成</th>
                            <th class="th-type">種別</th>
                            <th class="th-artist">歌手</th>
                            <th class="th-song">曲名</th>
                            <th class="th-num">人数</th>
                            <th class="th-num">歌唱数</th>
                        </tr>
                    </thead>
            """

            # 作品名でグループ化（rowspan 集約 + スマホアコーディオン）
            anime_groups = []
            for anime_name, group_iter in groupby(rendered_items, key=lambda x: x['anime']):
                anime_groups.append((anime_name, list(group_iter)))

            for g_idx, (anime_name, group_items) in enumerate(anime_groups):
                rowspan = len(group_items)
                group_total_count = sum(g['count'] for g in group_items)
                group_total_user = sum(g['user_count'] for g in group_items)
                group_total_made = sum(g['creation_count'] for g in group_items)

                # スマホ用アコーディオンヘッダ（PCではCSSで非表示）
                analysis_html_content += f'''<tbody class="anime-group" data-anime="{anime_name}">
                    <tr class="anime-head" onclick="toggleAnimeGroup(this)">
                        <td colspan="7">
                            <span class="anime-head-chev"><i class="fas fa-chevron-right"></i></span>
                            <span class="anime-head-title">{anime_name}</span>
                            <span class="anime-head-meta">
                                <span class="ahm">曲 {rowspan}</span>
                                <span class="ahm">人数 {group_total_user}</span>
                                <span class="ahm">歌唱 {group_total_count}</span>
                                <span class="ahm made-summary">作成 {group_total_made}</span>
                            </span>
                        </td>
                    </tr>
                '''

                for i, item in enumerate(group_items):
                    user_pct = calc_bar_pct(item['user_count'], max_user_in_cat)
                    count_pct = calc_bar_pct(item['count'], max_count_in_cat)

                    user_bar = f'<span class="ibar ibar-user" style="width:{user_pct}%"></span>' if item['user_count'] > 0 else ""
                    count_bar = f'<span class="ibar ibar-count" style="width:{count_pct}%"></span>' if item['count'] > 0 else ""

                    clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                    search_word = f"{clean_anime} {item['song']}"
                    link_start = f'<a href="#search_link/{search_word}" class="export-link">'

                    cc = item['creation_count']
                    if cc >= 1:
                        made_class = "made-yes"
                        made_label = str(cc)
                    else:
                        made_class = "made-no"
                        made_label = "—"

                    analysis_html_content += '<tr class="anime-song">'
                    if i == 0:
                        analysis_html_content += f'<td class="td-anime" rowspan="{rowspan}" data-label="作品名">{item["anime"]}</td>'
                    analysis_html_content += f'''<td class="td-made" data-label="作成"><span class="made-tag {made_class}">{made_label}</span></td>
                        <td class="td-type" data-label="種別">{link_start}<span class="type-tag">{item["type"]}</span></a></td>
                        <td class="td-artist" data-label="歌手">{link_start}{item["artist"]}</a></td>
                        <td class="td-song" data-label="曲名">{link_start}{item["song"]}</a></td>
                        <td class="td-num" data-label="人数"><span class="num-cell"><b>{item["user_count"]}</b><span class="ibar-wrap">{user_bar}</span></span></td>
                        <td class="td-num" data-label="歌唱数"><span class="num-cell"><b>{item["count"]}</b><span class="ibar-wrap">{count_bar}</span></span></td>
                    </tr>'''

                analysis_html_content += '</tbody>'

            analysis_html_content += "</table></div></section>"

            created_lists_html += generate_category_html_block(category, cat_created_items)
            uncreated_lists_html += generate_category_html_block(category, cat_uncreated_items)

        cool_data_exists = True
        print("クール集計処理完了。")

        print("ランキング生成処理開始...")

        def generate_ranking_html(mode="count"):
            html_out = ""
            for target_cat in ALLOWED_CATEGORIES:
                if target_cat not in categorized_data:
                    continue

                cat_items = [d for d in ranking_data_list if d["category"] == target_cat and d["count"] > 0]

                if mode == "count":
                    cat_items.sort(key=lambda x: (x["count"], x["user_count"]), reverse=True)
                    rank_title = f"{target_cat} 歌唱数ランキング (TOP 20)"
                    val_key = "count"
                else:
                    cat_items.sort(key=lambda x: (x["user_count"], x["count"]), reverse=True)
                    rank_title = f"{target_cat} 歌唱人数ランキング (TOP 20)"
                    val_key = "user_count"

                # ランキング内（TOP20想定）でも相対バーが効くよう、上位の最大値を基準にする
                top_slice = cat_items[:20]
                max_count_rank = max([d["count"] for d in top_slice], default=0)
                max_user_rank = max([d["user_count"] for d in top_slice], default=0)

                html_out += f"""
                <section class="cat-block">
                    <header class="cat-header" onclick="toggleCategory(this)">
                        <span class="cat-title">{rank_title}</span>
                        <i class="fas fa-chevron-down cat-chev"></i>
                    </header>
                    <div class="cat-content">
                    <table class="rankingTable">
                        <thead>
                            <tr>
                                <th class="th-rank">順位</th>
                                <th class="th-anime">作品名</th>
                                <th class="th-song">曲名</th>
                                <th class="th-artist">歌手</th>
                                <th class="th-num">人数</th>
                                <th class="th-num">歌唱数</th>
                            </tr>
                        </thead>
                        <tbody>
                """

                if not cat_items:
                    html_out += '<tr><td colspan="6" class="rank-empty">歌唱データがありません</td></tr>'
                else:
                    previous_val = None
                    current_rank = 0

                    for i, item in enumerate(cat_items):
                        current_val = item[val_key]

                        if current_val != previous_val:
                            current_rank = i + 1

                        if current_rank > 20:
                            break

                        previous_val = current_val

                        if current_rank == 1:
                            rank_tag_class = "rank-tag rank-gold"
                        elif current_rank == 2:
                            rank_tag_class = "rank-tag rank-silver"
                        elif current_rank == 3:
                            rank_tag_class = "rank-tag rank-bronze"
                        else:
                            rank_tag_class = "rank-tag rank-normal"

                        row_rank_class = f"rank-row-{current_rank}" if current_rank <= 3 else ""

                        rank_display = f'<span class="{rank_tag_class}">{current_rank}</span>'

                        user_pct = calc_bar_pct(item["user_count"], max_user_rank)
                        count_pct = calc_bar_pct(item["count"], max_count_rank)

                        user_bar = f'<span class="ibar ibar-user" style="width:{user_pct}%"></span>' if item["user_count"] > 0 else ""
                        count_bar = f'<span class="ibar ibar-count" style="width:{count_pct}%"></span>' if item["count"] > 0 else ""

                        clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                        search_word = f"{clean_anime} {item['song']}"
                        link_start = f'<a href="#search_link/{search_word}" class="export-link">'

                        html_out += f"""
                        <tr class="ranking-row {row_rank_class}" data-href="#search_link/{search_word}" onclick="toggleRankRow(this, event)">
                            <td class="td-rank" data-label="順位">{rank_display}</td>
                            <td class="td-anime" data-label="作品名">
                                <span class="anime-main">{item["anime"]}</span>
                                <span class="type-tag-inline">{item["type"]}</span>
                                <span class="rank-mob-chev"><i class="fas fa-chevron-right"></i></span>
                            </td>
                            <td class="td-song" data-label="曲名">{link_start}{item["song"]}</a></td>
                            <td class="td-artist" data-label="歌手">{item["artist"]}</td>
                            <td class="td-num" data-label="人数"><span class="num-cell"><b>{item["user_count"]}</b><span class="ibar-wrap">{user_bar}</span></span></td>
                            <td class="td-num" data-label="歌唱数"><span class="num-cell"><b>{item["count"]}</b><span class="ibar-wrap">{count_bar}</span></span></td>
                        </tr>
                        """

                html_out += "</tbody></table></div></section>"
            return html_out

        ranking_count_html_content = generate_ranking_html("count")
        ranking_user_html_content = generate_ranking_html("user")

        print("ランキング生成完了。")

    except Exception as e:
        print(f"集計エラー: {e}")
        import traceback
        traceback.print_exc()

else:
    print("CSV読み込み失敗: cool_analysis.csv がGASから取得できませんでした。")


# ==========================================
# HTML生成 (HTML出力・印刷設定)
# ==========================================

columns_to_hide = ['コメント']

# --- セットリスト用 DataFrame の整形：取得日(降順) → 順番(降順) ---
if not full_df.empty:
    html_df = full_df.drop(columns=columns_to_hide, errors='ignore').copy()

    # 並び順を「取得日(新しい順) → 順番(大きい順)」に統一
    html_df['_dt_sort'] = pd.to_datetime(html_df['取得日'], errors='coerce')
    html_df['_order_sort'] = pd.to_numeric(html_df['順番'], errors='coerce')
    html_df = html_df.sort_values(
        by=['_dt_sort', '_order_sort'],
        ascending=[False, False],   # 取得日: 新しい順 / 順番: 大きい順
        kind='mergesort',
        na_position='last'
    )
    html_df = html_df.drop(columns=['_dt_sort', '_order_sort'])
else:
    html_df = pd.DataFrame()

def _safe(v):
    if v is None:
        return ""
    s = str(v)
    return s if s and s.lower() != 'nan' else ""

setlist_rows = ""
total_setlist_rows = 0
if not html_df.empty:
    cols = list(html_df.columns)
    has_room = '部屋主' in cols
    has_order = '順番' in cols
    has_song = '曲名（ファイル名）' in cols
    has_work = '作品名' in cols
    has_artist = '歌手名' in cols
    has_singer = '歌った人' in cols
    has_date = '取得日' in cols

    for _, row in html_df.iterrows():
        room = _safe(row['部屋主']) if has_room else ""
        order = _safe(row['順番']) if has_order else ""
        song = _safe(row['曲名（ファイル名）']) if has_song else ""
        work = _safe(row['作品名']) if has_work else ""
        artist = _safe(row['歌手名']) if has_artist else ""
        singer = _safe(row['歌った人']) if has_singer else ""
        date = _safe(row['取得日']) if has_date else ""

        # 検索用テキスト（hidden）
        search_text = " ".join([room, order, song, work, artist, singer, date]).upper()

        # 作成有無タグはセットリストでは表示しない（仕様変更）
        setlist_rows += f'''<tr class="setlist-row" data-search="{search_text}">
            <td class="cell-main">
                <div class="row-top">
                    <span class="col-order">{order}</span>
                    <span class="col-song" title="{song}">{song}</span>
                    <span class="col-date">{date}</span>
                </div>
                <div class="row-sub">
                    <span class="col-room"><span class="room-tag">{room}</span></span>
                    <span class="col-work" title="{work}">{work}</span>
                    <span class="col-artist">{artist}</span>
                    <span class="col-singer"><i class="fas fa-microphone"></i> {singer}</span>
                </div>
            </td>
        </tr>'''
        total_setlist_rows += 1

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
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns/dist/chartjs-adapter-date-fns.bundle.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <style>
        :root {{
            --primary: #1f2937;
            --primary-soft: #374151;
            --accent: #2563eb;
            --accent-soft: #3b82f6;
            --bg: #f5f6f8;
            --panel: #ffffff;
            --text: #1f2937;
            --text-sub: #6b7280;
            --text-mute: #9ca3af;
            --border: #e5e7eb;
            --border-soft: #eef0f3;
            --green: #10b981;
            --green-bg: #ecfdf5;
            --red: #ef4444;
            --red-bg: #fef2f2;
            --amber: #f59e0b;
            --row-h: 24px;
            --radius: 6px;
            --radius-lg: 10px;
        }}

        * {{ box-sizing: border-box; }}
        html, body {{
            height: 100%; margin: 0; padding: 0;
            overflow: hidden;
            font-family: "Inter", "Noto Sans JP", "Helvetica Neue", Arial, sans-serif;
            background: var(--bg);
            color: var(--text);
            font-size: 13px;
            line-height: 1.4;
            display: flex; flex-direction: column;
            -webkit-font-smoothing: antialiased;
        }}

        a.export-link {{
            color: inherit;
            text-decoration: none;
            pointer-events: none;
            cursor: default;
        }}

        /* ============== Top header ============== */
        .top-section {{
            flex: 0 0 auto;
            background: var(--panel);
            border-bottom: 1px solid var(--border);
            box-shadow: 0 1px 0 rgba(0,0,0,0.02);
            z-index: 100;
        }}
        .header-inner {{
            padding: 8px 16px;
            display: flex; justify-content: space-between; align-items: center; gap: 12px;
            flex-wrap: wrap;
        }}
        h1 {{
            margin: 0; font-size: 1.05rem; font-weight: 700;
            color: var(--primary);
            letter-spacing: 0.02em;
            display: flex; align-items: center; gap: 8px;
        }}
        h1::before {{
            content: "";
            display: inline-block; width: 4px; height: 18px;
            background: linear-gradient(180deg, var(--accent), var(--accent-soft));
            border-radius: 2px;
        }}
        .update-time {{ font-size: 0.78rem; color: var(--text-sub); }}
        .port-input-wrapper {{
            display: inline-flex; align-items: center; gap: 6px;
            margin-left: 12px; font-size: 12px;
            color: var(--primary-soft);
        }}
        .port-input-wrapper input, .port-input-wrapper select {{
            padding: 4px 8px;
            border: 1px solid var(--border);
            border-radius: var(--radius);
            background: #fff;
            font-family: inherit;
            font-size: 12px;
            color: var(--text);
        }}
        .port-input-wrapper input {{ width: 80px; text-align: center; }}

        .tabs {{
            display: flex; padding: 0 12px; border-bottom: 1px solid var(--border);
            overflow-x: auto; background: var(--panel);
        }}
        .tab-btn {{
            padding: 9px 16px; cursor: pointer; border: none; background: none;
            font-weight: 600; color: var(--text-sub);
            border-bottom: 2px solid transparent;
            font-size: 13px; white-space: nowrap;
            transition: color 0.15s, border-color 0.15s;
        }}
        .tab-btn:hover {{ color: var(--primary); }}
        .tab-btn.active {{ color: var(--accent); border-bottom-color: var(--accent); }}

        .controls-row {{
            padding: 8px 16px;
            display: flex; gap: 8px; align-items: center;
            background: var(--panel); border-bottom: 1px solid var(--border);
            min-height: 44px;
            flex-wrap: nowrap;
            overflow-x: auto;
        }}
        .search-box {{
            padding: 6px 12px; border: 1px solid var(--border); border-radius: var(--radius);
            width: 280px; font-size: 13px; outline: none;
            background: #fff;
            transition: border-color 0.15s, box-shadow 0.15s;
        }}
        .search-box:focus {{
            border-color: var(--accent);
            box-shadow: 0 0 0 3px rgba(37,99,235,0.12);
        }}
        .btn {{
            padding: 6px 12px; border-radius: var(--radius); border: 1px solid var(--accent);
            cursor: pointer; color: #fff; background: var(--accent);
            font-size: 12.5px; font-weight: 600; white-space: nowrap;
            transition: background 0.15s, transform 0.05s;
        }}
        .btn:hover {{ background: #1d4ed8; }}
        .btn:active {{ transform: translateY(1px); }}
        .btn.btn-ghost {{ background: #fff; color: var(--text-sub); border-color: var(--border); }}
        .btn.btn-ghost:hover {{ background: #f3f4f6; color: var(--primary); }}
        .btn-dl {{ background: var(--green); border-color: var(--green); }}
        .btn-dl:hover {{ background: #059669; }}
        .btn-list {{ background: #6366f1; border-color: #6366f1; }}
        .btn-list:hover {{ background: #4f46e5; }}
        .btn-list.danger {{ background: var(--red); border-color: var(--red); }}
        .btn-list.danger:hover {{ background: #dc2626; }}
        .count-display {{
            margin-left: auto; font-weight: 600; font-size: 12.5px;
            color: var(--primary-soft);
            background: #f3f4f6;
            padding: 5px 12px;
            border-radius: 999px;
        }}

        .ctrl-setlist {{ display: flex; width: 100%; align-items: center; gap: 8px; }}
        .ctrl-analysis, .ctrl-ranking, .ctrl-graph {{ display: none; width: 100%; align-items: center; gap: 8px; }}

        .content-area {{
            flex: 1; position: relative; overflow: hidden;
        }}
        .tab-content {{
            display: none; position: absolute;
            top: 0; left: 0; right: 0; bottom: 0;
            overflow-y: auto;
            -webkit-overflow-scrolling: touch;
            padding: 0 16px 32px 16px;
        }}
        .tab-content.active {{ display: block; }}

        /* ============== Category block (共通) ============== */
        .cat-block {{
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            margin-top: 14px;
            overflow: hidden;
            box-shadow: 0 1px 2px rgba(0,0,0,0.03);
        }}
        .cat-header {{
            padding: 9px 14px;
            background: linear-gradient(180deg, #fafbfc 0%, #f3f4f6 100%);
            color: var(--primary);
            font-weight: 700; font-size: 0.95rem;
            cursor: pointer; user-select: none;
            display: flex; align-items: center; gap: 10px;
            border-bottom: 1px solid var(--border);
            border-left: 3px solid var(--accent);
        }}
        .cat-title {{ flex: 1; }}
        .cat-count, .cat-meta {{
            font-size: 0.72rem; font-weight: 600;
            color: var(--text-sub);
            background: #fff;
            padding: 2px 8px;
            border-radius: 4px;
            border: 1px solid var(--border);
        }}
        .cat-meta {{ color: var(--text-mute); }}
        .cat-chev {{ color: var(--text-mute); transition: transform 0.2s; font-size: 0.85rem; }}
        .cat-content.collapsed {{ display: none; }}

        /* ============== Setlist Table ============== */
        #setlistTable {{
            width: 100%;
            border-collapse: collapse;
            background: var(--panel);
            border-radius: var(--radius-lg);
            margin-top: 12px;
            overflow: hidden;
            border: 1px solid var(--border);
        }}
        #setlistTable thead {{ display: none; }}
        #setlistTable tr.setlist-row {{
            display: block;
            border-bottom: 1px solid var(--border-soft);
            padding: 0;
            background: #fff;
            transition: background 0.12s;
        }}
        #setlistTable tr.setlist-row:nth-child(even) {{ background: #fbfcfd; }}
        #setlistTable tr.setlist-row:hover {{ background: #eef4fb; }}
        #setlistTable td.cell-main {{
            display: block;
            padding: 4px 12px 5px 12px;
            border: none;
        }}
        .row-top {{
            display: grid;
            grid-template-columns: 50px 1fr 90px;
            gap: 12px;
            align-items: baseline;
        }}
        .row-sub {{
            display: grid;
            grid-template-columns: 130px 1.4fr 1fr 1fr;
            gap: 10px;
            align-items: start;
            margin-top: 1px;
            padding-left: 62px;
        }}
        .col-order {{
            font-family: "Inter", monospace;
            font-weight: 700;
            font-size: 13px;
            color: var(--text-sub);
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .col-song {{
            font-weight: 700;
            font-size: 14px;
            color: var(--primary);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .col-date {{
            font-size: 12px;
            color: var(--text-mute);
            text-align: right;
            font-variant-numeric: tabular-nums;
            white-space: nowrap;
        }}
        .row-sub > span {{
            font-size: 12.5px;
            color: var(--text-sub);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            line-height: 1.35;
        }}
        /* 作品名は2行まで折返し */
        .col-work {{
            color: #4b5563;
            white-space: normal;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
            line-height: 1.3;
        }}
        .room-tag {{
            display: inline-block;
            padding: 1px 8px;
            font-size: 11.5px;
            font-weight: 600;
            color: var(--accent);
            background: #eff6ff;
            border: 1px solid #dbeafe;
            border-radius: 4px;
            line-height: 1.5;
            white-space: nowrap;
            max-width: 100%;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .col-artist {{ color: #4b5563; font-style: normal; font-size: 12.5px; }}
        .col-singer {{ color: var(--text-sub); font-size: 12.5px; }}
        .col-singer i {{ font-size: 10px; color: var(--text-mute); margin-right: 2px; }}

        tr.hidden {{ display: none !important; }}

        .setlist-pseudo-header {{
            display: grid;
            grid-template-columns: 50px 1fr 90px;
            gap: 12px;
            padding: 8px 12px;
            background: linear-gradient(180deg, #1f2937 0%, #111827 100%);
            color: #fff;
            font-size: 11.5px;
            font-weight: 600;
            letter-spacing: 0.03em;
            position: sticky; top: 0;
            border-radius: var(--radius-lg) var(--radius-lg) 0 0;
            z-index: 5;
        }}
        .setlist-pseudo-header span:nth-child(1) {{ text-align: right; cursor: pointer; }}
        .setlist-pseudo-header span:nth-child(2) {{ cursor: pointer; }}
        .setlist-pseudo-header span:nth-child(3) {{ text-align: right; cursor: pointer; }}
        .setlist-pseudo-header span:hover {{ color: var(--accent-soft); }}

        /* ============== Mobile setlist (cards) ============== */
        @media (max-width: 720px) {{
            .setlist-pseudo-header {{ display: none; }}
            #setlistTable {{
                background: transparent;
                border: none;
                box-shadow: none;
            }}
            #setlistTable tr.setlist-row {{
                background: #fff !important;
                border: 1px solid var(--border);
                border-radius: var(--radius-lg);
                margin-bottom: 6px;
                box-shadow: 0 1px 2px rgba(0,0,0,0.03);
            }}
            #setlistTable td.cell-main {{ padding: 6px 10px; }}
            .row-top {{
                grid-template-columns: 32px 1fr 70px;
                gap: 8px;
            }}
            .col-song {{
                font-size: 13.5px;
                white-space: normal;
                display: -webkit-box;
                -webkit-line-clamp: 2;
                -webkit-box-orient: vertical;
                overflow: hidden;
                line-height: 1.3;
            }}
            .col-order {{ font-size: 12px; align-self: start; padding-top: 1px; }}
            .col-date {{ font-size: 11px; align-self: start; padding-top: 1px; }}
            .row-sub {{
                grid-template-columns: auto 1fr;
                grid-template-areas:
                    "room work"
                    "artist singer";
                padding-left: 0;
                gap: 2px 8px;
                margin-top: 4px;
                padding-top: 4px;
                border-top: 1px dashed var(--border);
            }}
            .col-room {{ grid-area: room; }}
            .col-work {{ grid-area: work; font-size: 12px; -webkit-line-clamp: 2; }}
            .col-artist {{ grid-area: artist; font-size: 11.5px; }}
            .col-singer {{ grid-area: singer; font-size: 11.5px; text-align: right; }}
        }}

        /* ============== 共通テーブル (analysis / ranking / list) ============== */
        .analysisTable, .rankingTable, .flatTable {{
            width: 100%;
            border-collapse: collapse;
            background: var(--panel);
            font-size: 12.5px;
        }}
        .analysisTable thead th,
        .rankingTable thead th,
        .flatTable thead th {{
            background: #f9fafb;
            color: var(--primary-soft);
            font-weight: 700;
            font-size: 11.5px;
            letter-spacing: 0.04em;
            text-align: left;
            padding: 7px 10px;
            border-bottom: 1px solid var(--border);
            white-space: nowrap;
        }}
        .analysisTable tbody td,
        .rankingTable tbody td,
        .flatTable tbody td {{
            padding: 4px 10px;
            border-bottom: 1px solid var(--border-soft);
            vertical-align: middle;
            line-height: 1.35;
        }}
        .analysisTable tbody tr:hover td,
        .rankingTable tbody tr:hover td,
        .flatTable tbody tr:hover td {{
            background: #f8fafc;
        }}
        .th-anime {{ width: 22%; }}
        .th-made {{ width: 50px; text-align: center !important; }}
        .th-type {{ width: 60px; text-align: center !important; }}
        .th-artist {{ width: 16%; }}
        .th-song {{ width: 26%; }}
        .th-num {{ width: 110px; }}
        .th-rank {{ width: 60px; text-align: center !important; }}
        .td-made {{ text-align: center; }}
        .td-type {{ text-align: center; }}
        .td-rank {{ text-align: center; }}
        .td-anime {{ font-weight: 600; color: var(--primary); vertical-align: middle; background: #fafbfc; }}
        .td-song {{ color: var(--primary-soft); }}
        .td-artist {{ color: var(--text-sub); }}
        .td-num {{ color: var(--text); white-space: nowrap; }}
        .num-cell {{ display: inline-flex; align-items: center; gap: 6px; }}
        .num-cell b {{
            display: inline-block;
            min-width: 22px; text-align: right;
            font-variant-numeric: tabular-nums;
            font-weight: 700;
        }}
        .ibar-wrap {{
            display: inline-block;
            width: 70px; height: 5px;
            background: #f1f5f9;
            border-radius: 2px;
            overflow: hidden;
        }}
        .ibar {{
            display: inline-block;
            height: 100%;
        }}
        .ibar-user {{ background: var(--green); }}
        .ibar-count {{ background: var(--accent); }}
        .type-tag {{
            display: inline-block;
            padding: 1px 7px;
            font-size: 10.5px;
            font-weight: 700;
            color: var(--text-sub);
            background: #f3f4f6;
            border: 1px solid var(--border);
            border-radius: 3px;
            line-height: 1.5;
        }}
        .type-tag-inline {{
            display: inline-block;
            margin-left: 6px;
            font-size: 10.5px;
            font-weight: 600;
            color: var(--text-mute);
        }}
        .made-tag {{
            display: inline-block;
            min-width: 28px;
            padding: 1px 6px;
            font-size: 11px; font-weight: 800;
            border-radius: 3px;
            line-height: 1.5;
        }}
        .made-tag.made-yes {{ color: var(--green); background: var(--green-bg); border: 1px solid #a7f3d0; }}
        .made-tag.made-no  {{ color: var(--text-mute); background: #f9fafb; border: 1px solid var(--border); }}

        /* ===== 作品アコーディオン（クール集計 スマホ用） ===== */
        .anime-head {{ display: none; }}  /* PCでは非表示 */
        .anime-head td {{
            background: linear-gradient(180deg,#f9fafb,#f3f4f6);
            border-top: 1px solid var(--border);
            padding: 8px 10px !important;
            cursor: pointer;
        }}
        .anime-head-chev {{
            display: inline-block; width: 16px; color: var(--text-mute);
            transition: transform 0.2s;
        }}
        .anime-head-title {{
            font-weight: 700; color: var(--primary); font-size: 13px; margin-right: 6px;
        }}
        .anime-head-meta {{
            display: inline-flex; gap: 6px; font-size: 11px; color: var(--text-sub);
        }}
        .ahm {{
            background: #fff; border: 1px solid var(--border); padding: 1px 6px;
            border-radius: 3px; font-weight: 600;
        }}
        .ahm.made-summary {{ color: var(--green); border-color: #a7f3d0; background: var(--green-bg); }}
        tbody.anime-group.expanded .anime-head-chev i {{ transform: rotate(90deg); display: inline-block; }}

        /* Ranking rectangle tag */
        .rank-tag {{
            display: inline-block;
            min-width: 30px; padding: 2px 7px;
            font-weight: 800; font-size: 12px;
            text-align: center;
            border-radius: 3px;
            font-variant-numeric: tabular-nums;
        }}
        .rank-normal {{ background: #f3f4f6; color: var(--text-sub); }}
        .rank-gold   {{ background: linear-gradient(180deg,#fde68a,#f59e0b); color: #78350f; box-shadow: inset 0 -1px 0 rgba(0,0,0,0.08); }}
        .rank-silver {{ background: linear-gradient(180deg,#e5e7eb,#9ca3af); color: #1f2937; }}
        .rank-bronze {{ background: linear-gradient(180deg,#fed7aa,#d97706); color: #7c2d12; }}
        tr.ranking-row {{ cursor: default; }}
        tr.rank-row-1 td {{ background: #fffbeb !important; }}
        tr.rank-row-2 td {{ background: #f9fafb !important; }}
        tr.rank-row-3 td {{ background: #fff7ed !important; }}
        .rank-empty {{ text-align: center; padding: 16px; color: var(--text-mute); }}
        .rank-mob-chev {{ display: none; color: var(--text-mute); margin-left: 6px; transition: transform 0.2s; }}

        /* ============== Mobile (analysis/ranking) ============== */
        @media (max-width: 720px) {{
            /* クール集計: アコーディオン化 */
            .analysisTable {{
                display: block;
                background: transparent;
            }}
            .analysisTable thead {{ display: none; }}
            .analysisTable tbody.anime-group {{
                display: block;
                background: #fff;
                border: 1px solid var(--border);
                border-radius: 8px;
                margin-bottom: 6px;
                overflow: hidden;
            }}
            .analysisTable tbody.anime-group tr {{ display: block; }}
            .analysisTable tbody.anime-group tr.anime-head {{ display: block; }}
            .analysisTable tbody.anime-group tr.anime-head td {{
                display: block; padding: 8px 12px !important;
                border: none;
                background: linear-gradient(180deg,#f9fafb,#f3f4f6);
            }}
            .analysisTable tbody.anime-group tr.anime-song {{
                display: none;
                padding: 6px 12px 4px 12px;
                border-top: 1px dashed var(--border);
            }}
            .analysisTable tbody.anime-group.expanded tr.anime-song {{ display: grid; }}
            .analysisTable tbody.anime-group tr.anime-song {{
                grid-template-columns: 60px 1fr 1fr;
                grid-template-areas:
                    "made   song    type"
                    "made   artist  artist"
                    "nums   nums    nums";
                gap: 2px 8px;
                background: #fff;
            }}
            .analysisTable tbody.anime-group tr.anime-song td {{
                display: block;
                border: none;
                padding: 1px 0;
                font-size: 12px;
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-anime {{
                display: none;  /* スマホではヘッダに集約 */
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-made {{
                grid-area: made; align-self: start;
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-song {{
                grid-area: song; font-weight: 700; color: var(--primary); font-size: 13px;
                white-space: normal;
                display: -webkit-box;
                -webkit-line-clamp: 2;
                -webkit-box-orient: vertical;
                overflow: hidden;
                line-height: 1.3;
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-type {{
                grid-area: type; text-align: right;
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-artist {{
                grid-area: artist; color: var(--text-sub); font-size: 11.5px;
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-num {{
                grid-area: nums; display: inline-flex !important; gap: 14px;
                margin-top: 2px; font-size: 11.5px;
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-num::before {{
                content: attr(data-label) ":";
                color: var(--text-mute); margin-right: 4px; font-weight: 600;
            }}
            .analysisTable tbody.anime-group tr.anime-song td.td-num:nth-of-type(7) {{
                grid-column: 2 / 4;
            }}

            /* ランキング: 行アコーディオン化 */
            .rankingTable {{ display: block; background: transparent; }}
            .rankingTable thead {{ display: none; }}
            .rankingTable tbody {{ display: block; }}
            .rankingTable tbody tr.ranking-row {{
                display: grid;
                grid-template-columns: 44px 1fr;
                grid-template-areas:
                    "rank anime"
                    "rank details";
                gap: 2px 10px;
                background: #fff !important;
                border: 1px solid var(--border);
                border-radius: 8px;
                margin-bottom: 6px;
                padding: 8px 10px;
                align-items: start;
            }}
            .rankingTable tbody tr.ranking-row td {{
                display: block;
                border: none;
                padding: 0;
                background: transparent !important;
                font-size: 12px;
            }}
            .rankingTable tbody tr.ranking-row td.td-rank {{
                grid-area: rank; align-self: center; text-align: left;
            }}
            .rankingTable tbody tr.ranking-row td.td-anime {{
                grid-area: anime;
                font-weight: 700; color: var(--primary); font-size: 13px;
                display: flex; align-items: center; gap: 4px;
            }}
            .rankingTable tbody tr.ranking-row td.td-anime .anime-main {{
                white-space: normal;
                display: -webkit-box;
                -webkit-line-clamp: 2;
                -webkit-box-orient: vertical;
                overflow: hidden;
                flex: 1;
            }}
            .rankingTable tbody tr.ranking-row td.td-anime .rank-mob-chev {{ display: inline-block; }}
            .rankingTable tbody tr.ranking-row.expanded td.td-anime .rank-mob-chev i {{
                transform: rotate(90deg); display: inline-block;
            }}
            .rankingTable tbody tr.ranking-row td.td-song,
            .rankingTable tbody tr.ranking-row td.td-artist,
            .rankingTable tbody tr.ranking-row td.td-num {{
                grid-area: details;
                display: none !important;
            }}
            .rankingTable tbody tr.ranking-row.expanded td.td-song,
            .rankingTable tbody tr.ranking-row.expanded td.td-artist,
            .rankingTable tbody tr.ranking-row.expanded td.td-num {{
                display: block !important;
                font-size: 11.5px;
                margin-top: 2px;
                color: var(--text-sub);
            }}
            .rankingTable tbody tr.ranking-row.expanded td.td-song {{
                color: var(--primary); font-weight: 600;
            }}
            .rankingTable tbody tr.ranking-row.expanded td.td-num {{
                display: inline-flex !important; gap: 14px;
            }}
            .rankingTable tbody tr.ranking-row.expanded td.td-num::before {{
                content: attr(data-label) ":";
                color: var(--text-mute); margin-right: 4px; font-weight: 600;
            }}
        }}

        /* ============== Graph ============== */
        .chart-wrapper {{
            background: var(--panel);
            padding: 10px 12px 12px 12px;
            border-radius: var(--radius-lg);
            border: 1px solid var(--border);
            margin-top: 12px;
            height: calc(100vh - 200px);
            display: flex;
            flex-direction: column;
        }}
        .chart-info {{
            min-height: 32px;
            padding: 6px 10px;
            text-align: center;
            font-weight: 600;
            font-size: 12.5px;
            color: var(--primary);
            background: #f9fafb;
            border: 1px solid var(--border);
            border-radius: var(--radius);
            margin-bottom: 8px;
        }}
        .canvas-container {{
            flex: 1;
            position: relative;
            min-height: 0;
        }}

        /* ============== Print ============== */
        @media print {{
            * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
            body {{ overflow: visible !important; height: auto !important; display: block !important; }}
            .top-section {{ display: none !important; }}
            .content-area {{ overflow: visible !important; position: static !important; }}
            .tab-content {{ position: static !important; display: block !important; overflow: visible !important; padding: 0 !important; }}
            .cat-content {{ display: block !important; }}
            .cat-header {{ page-break-after: avoid; }}
            tr {{ page-break-inside: avoid; }}
            thead {{ display: table-header-group; }}
            .chart-wrapper {{ height: auto; }}
        }}

        @media (max-width: 960px) {{
            .row-sub {{
                grid-template-columns: 110px 1fr 1fr;
                padding-left: 62px;
            }}
            .col-singer {{ grid-column: 3 / 4; text-align: right; }}
        }}

        @media (max-width: 420px) {{
            h1 {{ font-size: 0.95rem; }}
            .port-input-wrapper {{ display: none; }}
            .search-box {{ width: 180px; }}
        }}
    </style>
</head>
<body>
    <div class="top-section">
        <div class="header-inner">
            <div style="display:flex; align-items:center; flex-wrap:wrap;">
                <h1>Karaoke Dashboard</h1>
                <div class="port-input-wrapper">
                    <label for="exportPort"><i class="fas fa-network-wired"></i> 保存時ポート</label>
                    <input type="number" id="exportPort" value="11059" title="HTML保存時のURLポート番号を指定">
                </div>
                <div class="port-input-wrapper">
                    <label for="exportLinkType"><i class="fas fa-link"></i> 検索リンク</label>
                    <select id="exportLinkType">
                        <option value="eve">Everything</option>
                        <option value="ykr">ゆかりすたー</option>
                    </select>
                </div>
            </div>
            <div class="update-time"><i class="far fa-clock"></i> {current_datetime_str} 更新</div>
        </div>
        <div class="tabs">
            <button class="tab-btn active" onclick="openTab('setlist')">セットリスト</button>
            <button class="tab-btn" onclick="openTab('analysis')">クール集計</button>
            <button class="tab-btn" onclick="openTab('ranking_count')">歌唱数ランキング</button>
            <button class="tab-btn" onclick="openTab('ranking_user')">歌唱人数ランキング</button>
            <button class="tab-btn" onclick="openTab('graph_view_count')">推移(数)</button>
            <button class="tab-btn" onclick="openTab('graph_view_user')">推移(人)</button>
        </div>
        <div class="controls-row">
            <div id="ctrl-setlist" class="ctrl-setlist">
                <input type="text" id="searchInput" class="search-box" placeholder="キーワード検索 (例: 曲名 歌手 部屋)...">
                <button onclick="performSearch()" class="btn"><i class="fas fa-search"></i> 検索</button>
                <button onclick="resetFilter()" class="btn btn-ghost"><i class="fas fa-undo"></i></button>
                <div class="count-display" id="countDisplay">読み込み中...</div>
            </div>
            <div id="ctrl-analysis" class="ctrl-analysis">
                <select id="exportTargetCategory" style="padding:5px 8px; border-radius:6px; font-size:12.5px; border:1px solid var(--border); background:#fff;">
                    {category_options}
                </select>
                <button onclick="downloadListWithCategory('list-created-content', 'created_list.html', '作成済みリスト')" class="btn btn-list">作成リスト保存</button>
                <button onclick="downloadListWithCategory('list-uncreated-content', 'uncreated_list.html', '未作成リスト')" class="btn btn-list danger">未作成リスト保存</button>
                <button onclick="downloadHTMLWithCategory()" class="btn btn-dl"><i class="fas fa-file-code"></i> HTML保存</button>
            </div>
            <div id="ctrl-ranking-count" class="ctrl-ranking">
                <select id="exportTargetRankingCount" style="padding:5px 8px; border-radius:6px; font-size:12.5px; border:1px solid var(--border); background:#fff;">
                    {category_options}
                </select>
                <button onclick="downloadRankingWithCategory('count')" class="btn btn-dl"><i class="fas fa-trophy"></i> 歌唱数ランキング保存</button>
            </div>
            <div id="ctrl-ranking-user" class="ctrl-ranking">
                <select id="exportTargetRankingUser" style="padding:5px 8px; border-radius:6px; font-size:12.5px; border:1px solid var(--border); background:#fff;">
                    {category_options}
                </select>
                <button onclick="downloadRankingWithCategory('user')" class="btn btn-dl"><i class="fas fa-users"></i> 歌唱人数ランキング保存</button>
            </div>
            <div id="ctrl-graph" class="ctrl-graph" style="justify-content:flex-end;">
                <button onclick="downloadGraphHTML()" class="btn btn-dl"><i class="fas fa-file-code"></i> HTML保存</button>
            </div>
        </div>
    </div>

    <div class="content-area">
        <div id="setlist" class="tab-content active">
            <div class="setlist-pseudo-header">
                <span onclick="sortSetlist('order')">順番 <i class="fas fa-sort"></i></span>
                <span onclick="sortSetlist('song')">曲名 / 作品名・歌手・歌った人 <i class="fas fa-sort"></i></span>
                <span onclick="sortSetlist('date')">取得日 <i class="fas fa-sort"></i></span>
            </div>
            <table id="setlistTable">
                <tbody>{setlist_rows}</tbody>
            </table>
            {"" if setlist_rows else '<div style="padding:20px;text-align:center;color:var(--text-mute);">データがありません</div>'}
        </div>

        <div id="analysis" class="tab-content">
            <div style="margin-top:10px; font-size:0.78rem; color:var(--text-mute); text-align:right;">集計対象期間: 2026/01/01 - 2026/06/30　/　バーは各カテゴリ内最大値を100%とする相対表示</div>
            <div id="print-target">
                {analysis_html_content if cool_data_exists else '<div style="padding:20px;text-align:center;color:var(--red);">集計データがありません</div>'}
            </div>
        </div>

        <div id="ranking_count" class="tab-content">
            <div style="margin-top:10px; font-size:0.78rem; color:var(--text-mute); text-align:right;">集計対象期間: 2026/01/01 - 2026/06/30　/　バーはTOP20内最大値を100%とする相対表示</div>
            <div id="ranking-count-print-target">
                {ranking_count_html_content if ranking_count_html_content else '<div style="padding:20px;text-align:center;color:var(--red);">ランキング対象データがありません</div>'}
            </div>
        </div>

        <div id="ranking_user" class="tab-content">
            <div style="margin-top:10px; font-size:0.78rem; color:var(--text-mute); text-align:right;">集計対象期間: 2026/01/01 - 2026/06/30　/　バーはTOP20内最大値を100%とする相対表示</div>
            <div id="ranking-user-print-target">
                {ranking_user_html_content if ranking_user_html_content else '<div style="padding:20px;text-align:center;color:var(--red);">ランキング対象データがありません</div>'}
            </div>
        </div>

        <div id="graph_view_count" class="tab-content">
            <section class="cat-block" style="margin-top:14px;">
                <header class="cat-header"><span class="cat-title">2026年春アニメ 歌唱数ランキング推移 (Top 20)</span></header>
            </section>
            <div class="chart-wrapper">
                <div id="chart-info-count" class="chart-info">グラフの点をタップ・ホバーで詳細を表示</div>
                <div class="canvas-container"><canvas id="rankingChartCount"></canvas></div>
            </div>
        </div>
        <div id="graph_view_user" class="tab-content">
            <section class="cat-block" style="margin-top:14px;">
                <header class="cat-header"><span class="cat-title">2026年春アニメ 歌唱人数ランキング推移 (Top 20)</span></header>
            </section>
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

    const colors = [
        '#2563eb','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#ec4899','#84cc16',
        '#f97316','#14b8a6','#a855f7','#0ea5e9','#dc2626','#65a30d','#7c3aed','#0891b2',
        '#db2777','#ca8a04','#4f46e5','#059669'
    ];

    function initChart(type, dataObj, canvasId) {{
        if(charts[type]) return;
        const ctx = document.getElementById(canvasId).getContext('2d');
        const infoDivId = type === 'count' ? 'chart-info-count' : 'chart-info-user';

        const allKeys = Object.keys(dataObj);
        const latestRank = [];
        allKeys.forEach(key => {{
            const arr = dataObj[key];
            if(arr.length > 0) {{
                latestRank.push({{ key: key, rank: arr[arr.length - 1].y }});
            }}
        }});
        latestRank.sort((a,b) => a.rank - b.rank);
        const top5 = latestRank.slice(0, 5).map(x => x.key);

        const datasets = allKeys.map((key, i) => {{
            const color = colors[i % colors.length];
            const isTop5 = top5.includes(key);
            return {{
                label: key,
                data: dataObj[key],
                borderColor: color,
                backgroundColor: color,
                pointRadius: 3,
                pointHoverRadius: 6,
                tension: 0.15,
                fill: false,
                borderWidth: 2,
                hidden: !isTop5
            }};
        }});

        charts[type] = new Chart(ctx, {{
            type: 'line',
            data: {{ datasets }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                interaction: {{ mode: 'nearest', axis: 'x', intersect: true }},
                plugins: {{
                    tooltip: {{
                        enabled: false,
                        external: function(context) {{
                            const tooltip = context.tooltip;
                            const infoDiv = document.getElementById(infoDivId);
                            if (tooltip.opacity === 0) return;
                            if (tooltip.body) {{
                                const dp = tooltip.dataPoints[0];
                                const dateObj = new Date(dp.label);
                                const dateStr = (dateObj.getMonth() + 1) + '/' + dateObj.getDate();
                                infoDiv.innerHTML = `<span style="color:${{dp.dataset.borderColor}}">●</span> ${{dp.dataset.label}}　${{dateStr}}（${{dp.parsed.y}}位）`;
                            }}
                        }}
                    }},
                    legend: {{
                        position: 'bottom',
                        labels: {{ boxWidth: 10, padding: 10, font: {{ size: 11 }} }},
                        onClick: function(e, legendItem, legend) {{
                            const index = legendItem.datasetIndex;
                            const ci = legend.chart;
                            if (ci.isDatasetVisible(index)) {{ ci.hide(index); legendItem.hidden = true; }}
                            else {{ ci.show(index); legendItem.hidden = false; }}
                        }}
                    }}
                }},
                scales: {{
                    y: {{
                        reverse: true, min: 0.5, max: 20.5,
                        ticks: {{
                            stepSize: 1,
                            callback: function(val) {{ if (val % 1 === 0 && val >= 1 && val <= 20) return val; return ''; }}
                        }},
                        title: {{ display: true, text: '順位' }},
                        grid: {{ color: '#eef2f7' }}
                    }},
                    x: {{
                        type: 'time',
                        time: {{ unit: 'day', displayFormats: {{ day: 'M/d' }} }},
                        title: {{ display: true, text: '日付' }},
                        grid: {{ color: '#f3f4f6' }}
                    }}
                }}
            }}
        }});
    }}

    function downloadGraphHTML() {{
        const isCount = document.getElementById('graph_view_count').classList.contains('active');
        const canvasId = isCount ? 'rankingChartCount' : 'rankingChartUser';
        const title = isCount ? "推移(数)" : "推移(人)";
        const filename = 'graph.html';

        const canvas = document.getElementById(canvasId);
        const imgData = canvas.toDataURL('image/png');

        const headerText = isCount ?
            '2026年春アニメ 歌唱数ランキング推移 (Top 20)' :
            '2026年春アニメ 歌唱人数ランキング推移 (Top 20)';

        const content = `
            <section class="cat-block"><header class="cat-header"><span class="cat-title">${{headerText}}</span></header></section>
            <div class="chart-wrapper">
                <img src="${{imgData}}" style="width:100%; max-width:900px; border:1px solid #e5e7eb; display:block; margin:0 auto; border-radius:8px;">
            </div>
        `;
        generateDownload(content, filename, title);
    }}

    function openTab(tabName) {{
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.getElementById(tabName).classList.add('active');

        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

        let btns = document.querySelectorAll('.tab-btn');
        for(let i=0; i<btns.length; i++) {{
            if(btns[i].innerText.includes("セットリスト") && tabName === 'setlist') btns[i].classList.add('active');
            else if(btns[i].innerText.includes("クール集計") && tabName === 'analysis') btns[i].classList.add('active');
            else if(btns[i].innerText.includes("歌唱数ランキング") && tabName === 'ranking_count') btns[i].classList.add('active');
            else if(btns[i].innerText.includes("歌唱人数ランキング") && tabName === 'ranking_user') btns[i].classList.add('active');
            else if(btns[i].innerText.includes("推移(数)") && tabName === 'graph_view_count') btns[i].classList.add('active');
            else if(btns[i].innerText.includes("推移(人)") && tabName === 'graph_view_user') btns[i].classList.add('active');
        }}

        document.getElementById('ctrl-setlist').style.display = 'none';
        document.getElementById('ctrl-analysis').style.display = 'none';
        document.getElementById('ctrl-ranking-count').style.display = 'none';
        document.getElementById('ctrl-ranking-user').style.display = 'none';
        document.getElementById('ctrl-graph').style.display = 'none';

        if(tabName === 'setlist') document.getElementById('ctrl-setlist').style.display = 'flex';
        else if(tabName === 'analysis') document.getElementById('ctrl-analysis').style.display = 'flex';
        else if(tabName === 'ranking_count') document.getElementById('ctrl-ranking-count').style.display = 'flex';
        else if(tabName === 'ranking_user') document.getElementById('ctrl-ranking-user').style.display = 'flex';
        else if(tabName === 'graph_view_count') {{
            document.getElementById('ctrl-graph').style.display = 'flex';
            initChart('count', dataCount, 'rankingChartCount');
        }}
        else if(tabName === 'graph_view_user') {{
            document.getElementById('ctrl-graph').style.display = 'flex';
            initChart('user', dataUser, 'rankingChartUser');
        }}
    }}

    function toggleCategory(header) {{
        const content = header.nextElementSibling;
        if (!content) return;
        content.classList.toggle('collapsed');
        const icon = header.querySelector('i.cat-chev');
        if(icon) {{
            icon.style.transform = content.classList.contains('collapsed') ? 'rotate(-90deg)' : 'rotate(0deg)';
        }}
    }}

    // クール集計: スマホで作品アコーディオンを開閉
    function toggleAnimeGroup(headRow) {{
        const tbody = headRow.parentElement;
        if (!tbody) return;
        tbody.classList.toggle('expanded');
    }}

    // ランキング: スマホで行アコーディオンを開閉
    function toggleRankRow(row, ev) {{
        // モバイル幅でだけ動作
        if (!window.matchMedia('(max-width: 720px)').matches) return;
        if (window.getSelection && window.getSelection().toString().length > 0) return;
        // 曲名リンク等を踏んだ場合はトグルしない（PC版エクスポートでは別スクリプトでリンク化されるためここでは無視でOK）
        row.classList.toggle('expanded');
    }}

    function downloadHTML(elementId, filename, title) {{
        const element = document.getElementById(elementId);
        if(element) {{
            const htmlContent = element.innerHTML;
            generateDownload(htmlContent, filename, title);
        }}
    }}

    function extractCategoryHTML(containerId, targetCat) {{
        const container = document.getElementById(containerId);
        if (!container) return "";
        const blocks = container.querySelectorAll('.cat-block');
        let content = "";
        blocks.forEach(block => {{
            const header = block.querySelector('.cat-header .cat-title') || block.querySelector('.cat-header');
            if (header && header.innerText.includes(targetCat)) {{
                const clone = block.cloneNode(true);
                const catContent = clone.querySelector('.cat-content');
                if (catContent) catContent.classList.remove('collapsed');
                content += clone.outerHTML;
            }}
        }});
        return content;
    }}

    function downloadHTMLWithCategory() {{
        const targetCat = document.getElementById('exportTargetCategory').value;
        if (targetCat === "ALL") {{
            downloadHTML('print-target', 'karaoke_analysis.html', 'クール集計結果');
        }} else {{
            let content = extractCategoryHTML('print-target', targetCat);
            if (!content) content = `<div style="padding:20px;text-align:center;">${{targetCat}} のデータがありません</div>`;
            generateDownload(content, `karaoke_analysis_${{targetCat}}.html`, `${{targetCat}} 集計結果`);
        }}
    }}

    function downloadRankingWithCategory(mode) {{
        const targetCat = mode === 'count'
            ? document.getElementById('exportTargetRankingCount').value
            : document.getElementById('exportTargetRankingUser').value;
        const elementId = mode === 'count' ? 'ranking-count-print-target' : 'ranking-user-print-target';
        const baseTitle = mode === 'count' ? 'カラオケ歌唱数ランキング' : 'カラオケ歌唱人数ランキング';
        const baseFilename = mode === 'count' ? 'karaoke_ranking_count' : 'karaoke_ranking_user';

        if (targetCat === "ALL") {{
            downloadHTML(elementId, `${{baseFilename}}.html`, baseTitle);
        }} else {{
            let content = extractCategoryHTML(elementId, targetCat);
            if (!content) content = `<div style="padding:20px;text-align:center;">${{targetCat}} のデータがありません</div>`;
            generateDownload(content, `${{baseFilename}}_${{targetCat}}.html`, `${{targetCat}} ${{baseTitle}}`);
        }}
    }}

    function downloadListWithCategory(elementId, baseFilename, baseTitle) {{
        const targetCat = document.getElementById('exportTargetCategory').value;
        if (targetCat === "ALL") {{
            downloadHTML(elementId, baseFilename, baseTitle);
        }} else {{
            let content = extractCategoryHTML(elementId, targetCat);
            if (!content) content = `<div style="padding:20px;text-align:center;">${{targetCat}} のデータがありません</div>`;
            const ext = baseFilename.split('.').pop();
            const name = baseFilename.replace('.' + ext, '');
            generateDownload(content, `${{name}}_${{targetCat}}.${{ext}}`, `${{targetCat}} ${{baseTitle}}`);
        }}
    }}

    function generateDownload(content, filename, title) {{
        const portValue = document.getElementById('exportPort').value || '11059';
        const linkType = document.getElementById('exportLinkType').value;
        const searchPath = linkType === 'ykr' ? 'search_listerdb_filelist.php?anyword=' : 'search.php?searchword=';

        const fullHtml = `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${{title}}</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body {{ font-family: "Inter","Noto Sans JP","Helvetica Neue",Arial,sans-serif; font-size: 12.5px; color: #1f2937; background:#f5f6f8; padding:16px; }}
        h1 {{ font-size:1.2rem; margin: 0 0 6px 0; color: #1f2937; border-left: 4px solid #2563eb; padding-left: 10px; }}
        .meta {{ font-size: 0.8rem; color:#6b7280; margin-bottom:12px; }}
        table {{ width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow:hidden; border:1px solid #e5e7eb; margin-bottom: 16px; }}
        th, td {{ padding: 5px 10px; text-align: left; border-bottom: 1px solid #eef0f3; vertical-align: middle; font-size: 12.5px; }}
        th {{ background: #f9fafb; color: #374151; font-weight: 700; font-size: 11.5px; }}
        td.td-anime {{ background:#fafbfc; font-weight: 600; }}
        .cat-block {{ background:#fff; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden; margin-bottom:14px; }}
        .cat-header {{ padding: 9px 14px; background: linear-gradient(180deg,#fafbfc,#f3f4f6); border-left:3px solid #2563eb; font-weight:700; cursor:pointer; display:flex; gap:10px; align-items:center; }}
        .cat-title {{ flex:1; }}
        .cat-count, .cat-meta {{ font-size:0.75rem; color:#6b7280; background:#fff; padding:2px 8px; border-radius:4px; border:1px solid #e5e7eb; }}
        .cat-content.collapsed {{ display:none; }}
        a.export-link {{ display:inline-block; color: #2563eb; text-decoration:none; padding: 2px 0; }}
        a.export-link:hover {{ text-decoration: underline; color:#1d4ed8; }}
        .type-tag {{ display:inline-block; padding:1px 7px; font-size:10.5px; font-weight:700; color:#6b7280; background:#f3f4f6; border:1px solid #e5e7eb; border-radius:3px; }}
        .made-tag {{ display:inline-block; min-width:28px; padding:1px 6px; font-size:11px; font-weight:800; border-radius:3px; text-align:center; }}
        .made-tag.made-yes {{ color:#10b981; background:#ecfdf5; border:1px solid #a7f3d0; }}
        .made-tag.made-no  {{ color:#9ca3af; background:#f9fafb; border:1px solid #e5e7eb; }}
        .num-cell {{ display:inline-flex; align-items:center; gap:6px; }}
        .num-cell b {{ min-width:22px; text-align:right; font-weight:700; }}
        .ibar-wrap {{ display:inline-block; width:70px; height:5px; background:#f1f5f9; border-radius:2px; overflow:hidden; }}
        .ibar {{ display:inline-block; height:100%; }}
        .ibar-user {{ background:#10b981; }}
        .ibar-count {{ background:#2563eb; }}
        .rank-tag {{ display:inline-block; min-width:30px; padding:2px 7px; font-weight:800; font-size:12px; text-align:center; border-radius:3px; }}
        .rank-normal {{ background:#f3f4f6; color:#6b7280; }}
        .rank-gold {{ background:linear-gradient(180deg,#fde68a,#f59e0b); color:#78350f; }}
        .rank-silver {{ background:linear-gradient(180deg,#e5e7eb,#9ca3af); color:#1f2937; }}
        .rank-bronze {{ background:linear-gradient(180deg,#fed7aa,#d97706); color:#7c2d12; }}
        tr.rank-row-1 td {{ background:#fffbeb !important; }}
        tr.rank-row-2 td {{ background:#f9fafb !important; }}
        tr.rank-row-3 td {{ background:#fff7ed !important; }}
        .anime-head {{ display:none; }}
        .ahm {{ background:#fff; border:1px solid #e5e7eb; padding:1px 6px; border-radius:3px; font-weight:600; margin-right:4px; font-size:11px; color:#6b7280; }}
        .ahm.made-summary {{ color:#10b981; border-color:#a7f3d0; background:#ecfdf5; }}
        .anime-head-title {{ font-weight:700; color:#1f2937; }}
        .chart-wrapper {{ background:#fff; padding:10px; border-radius:8px; border:1px solid #e5e7eb; }}
        .rank-mob-chev {{ display:none; }}

        /* スマホレイアウト（出力HTMLでもアコーディオン） */
        @media (max-width: 720px) {{
            .anime-head {{ display:block; }}
            .anime-head td {{
                display:block; padding:8px 12px !important;
                background:linear-gradient(180deg,#f9fafb,#f3f4f6); cursor:pointer;
                border:none;
            }}
            table.analysisTable, table.analysisTable tbody {{ display:block; background:transparent; border:none; }}
            table.analysisTable thead {{ display:none; }}
            table.analysisTable tbody.anime-group {{
                display:block; background:#fff; border:1px solid #e5e7eb;
                border-radius:8px; margin-bottom:6px; overflow:hidden;
            }}
            table.analysisTable tbody.anime-group tr {{ display:block; }}
            table.analysisTable tbody.anime-group tr.anime-song {{
                display:none; padding:6px 12px 4px 12px;
                border-top: 1px dashed #e5e7eb;
                grid-template-columns: 60px 1fr 1fr;
                grid-template-areas: "made song type" "made artist artist" "nums nums nums";
                gap: 2px 8px;
            }}
            table.analysisTable tbody.anime-group.expanded tr.anime-song {{ display:grid; }}
            table.analysisTable tbody.anime-group tr.anime-song td {{ display:block; border:none; padding:1px 0; font-size:12px; }}
            table.analysisTable tbody.anime-group tr.anime-song td.td-anime {{ display:none; }}
            table.analysisTable tbody.anime-group tr.anime-song td.td-made {{ grid-area:made; align-self:start; }}
            table.analysisTable tbody.anime-group tr.anime-song td.td-song {{ grid-area:song; font-weight:700; color:#1f2937; font-size:13px; }}
            table.analysisTable tbody.anime-group tr.anime-song td.td-type {{ grid-area:type; text-align:right; }}
            table.analysisTable tbody.anime-group tr.anime-song td.td-artist {{ grid-area:artist; color:#6b7280; }}
            table.analysisTable tbody.anime-group tr.anime-song td.td-num {{
                grid-area:nums; display:inline-flex !important; gap:14px; margin-top:2px; font-size:11.5px;
            }}
            table.analysisTable tbody.anime-group tr.anime-song td.td-num::before {{
                content: attr(data-label) ":"; color:#9ca3af; margin-right:4px; font-weight:600;
            }}

            /* ランキング モバイル */
            table.rankingTable {{ display:block; background:transparent; border:none; }}
            table.rankingTable thead {{ display:none; }}
            table.rankingTable tbody {{ display:block; }}
            table.rankingTable tbody tr.ranking-row {{
                display:grid; grid-template-columns: 44px 1fr;
                grid-template-areas: "rank anime" "rank details";
                gap: 2px 10px;
                background:#fff !important; border:1px solid #e5e7eb;
                border-radius:8px; margin-bottom:6px; padding:8px 10px; align-items:start;
            }}
            table.rankingTable tbody tr.ranking-row td {{ display:block; border:none; padding:0; background:transparent !important; font-size:12px; }}
            table.rankingTable tbody tr.ranking-row td.td-rank {{ grid-area:rank; align-self:center; }}
            table.rankingTable tbody tr.ranking-row td.td-anime {{ grid-area:anime; font-weight:700; color:#1f2937; font-size:13px; display:flex; gap:4px; align-items:center;}}
            table.rankingTable tbody tr.ranking-row td.td-anime .rank-mob-chev {{ display:inline-block; color:#9ca3af; }}
            table.rankingTable tbody tr.ranking-row.expanded td.td-anime .rank-mob-chev i {{ transform:rotate(90deg); display:inline-block; }}
            table.rankingTable tbody tr.ranking-row td.td-song,
            table.rankingTable tbody tr.ranking-row td.td-artist,
            table.rankingTable tbody tr.ranking-row td.td-num {{ display:none !important; grid-area:details; }}
            table.rankingTable tbody tr.ranking-row.expanded td.td-song,
            table.rankingTable tbody tr.ranking-row.expanded td.td-artist {{ display:block !important; font-size:11.5px; margin-top:2px; }}
            table.rankingTable tbody tr.ranking-row.expanded td.td-num {{ display:inline-flex !important; gap:14px; }}
            table.rankingTable tbody tr.ranking-row.expanded td.td-num::before {{
                content: attr(data-label) ":"; color:#9ca3af; margin-right:4px; font-weight:600;
            }}
        }}

        @media print {{
            * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
            .cat-content {{ display:block !important; }}
            tr {{ page-break-inside: avoid; }}
            thead {{ display: table-header-group; }}
            .anime-head {{ display:none !important; }}
        }}
    </style>
</head>
<body>
    <h1>${{title}}</h1>
    <div class="meta">出力日: {current_date_str}</div>
    ${{content}}

    <script>
        const host = 'http://ykr.moe:${{portValue}}';
        const searchPath = '${{searchPath}}';

        document.addEventListener('DOMContentLoaded', () => {{
            // クール集計・リスト・ランキング 内のすべての export-link をクリック可能に
            document.querySelectorAll('a.export-link').forEach(link => {{
                const rawHref = link.getAttribute('href');
                if (rawHref && rawHref.startsWith('#search_link/')) {{
                    const word = rawHref.split('#search_link/')[1];
                    link.href = host + '/' + searchPath + word;
                    link.target = '_blank';
                    link.rel = 'noopener';
                }}
            }});

            // ランキング行クリック→ 検索ページ（PC）/ アコーディオン（スマホ）
            document.querySelectorAll('tr[data-href]').forEach(row => {{
                row.addEventListener('click', (e) => {{
                    if (window.getSelection().toString().length > 0) return;
                    if (e.target.closest('a')) return;  // 曲名リンクなどはそのまま遷移させる
                    if (window.matchMedia('(max-width: 720px)').matches) {{
                        row.classList.toggle('expanded');
                        return;
                    }}
                    const rawHref = row.getAttribute('data-href');
                    if (rawHref && rawHref.startsWith('#search_link/')) {{
                        const word = rawHref.split('#search_link/')[1];
                        window.open(host + '/' + searchPath + word, '_blank');
                    }}
                }});
            }});

            // 作品アコーディオン（スマホ）
            document.querySelectorAll('tr.anime-head').forEach(headRow => {{
                headRow.addEventListener('click', () => {{
                    const tbody = headRow.parentElement;
                    if (tbody) tbody.classList.toggle('expanded');
                }});
            }});
        }});

        function toggleCategory(header) {{
            const content = header.nextElementSibling;
            if (!content) return;
            content.classList.toggle('collapsed');
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

    // ===== Setlist 検索＆ソート =====
    const searchInput = document.getElementById("searchInput");
    const setlistTable = document.getElementById("setlistTable");
    const countDisplay = document.getElementById('countDisplay');
    let setlistRows = [];
    let setlistSearchText = [];
    let setlistSortDir = {{ order: 'desc', song: 'asc', date: 'desc' }};

    window.addEventListener('DOMContentLoaded', () => {{
        const tbody = setlistTable.tBodies[0];
        if (tbody) {{
            setlistRows = Array.from(tbody.rows);
            setlistSearchText = setlistRows.map(r => r.getAttribute('data-search') || r.innerText.toUpperCase());
            countDisplay.innerText = '全 ' + setlistRows.length + ' 件';
        }}
    }});

    searchInput.addEventListener("keyup", function(event) {{
        if (event.key === "Enter") performSearch();
    }});

    function performSearch() {{
        const filter = searchInput.value.toUpperCase();
        const keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        let visible = 0;
        for (let i = 0; i < setlistRows.length; i++) {{
            const text = setlistSearchText[i];
            let match = true;
            for (const kw of keywords) {{
                if (text.indexOf(kw) === -1) {{ match = false; break; }}
            }}
            if (match || keywords.length === 0) {{
                setlistRows[i].classList.remove('hidden');
                visible++;
            }} else {{
                setlistRows[i].classList.add('hidden');
            }}
        }}
        countDisplay.innerText = '表示: ' + visible + ' / ' + setlistRows.length;
    }}

    function resetFilter() {{
        searchInput.value = "";
        performSearch();
    }}

    function sortSetlist(key) {{
        const dir = setlistSortDir[key] === 'asc' ? 'desc' : 'asc';
        setlistSortDir[key] = dir;
        const tbody = setlistTable.tBodies[0];
        if (!tbody) return;
        const rows = Array.from(tbody.rows);

        const getVal = (row, k) => {{
            if (k === 'order') {{
                const t = row.querySelector('.col-order');
                const n = parseFloat(t ? t.innerText : '');
                return isNaN(n) ? Number.POSITIVE_INFINITY : n;
            }} else if (k === 'song') {{
                const t = row.querySelector('.col-song');
                return t ? t.innerText : '';
            }} else if (k === 'date') {{
                const t = row.querySelector('.col-date');
                return t ? t.innerText : '';
            }}
            return '';
        }};

        rows.sort((a, b) => {{
            const va = getVal(a, key);
            const vb = getVal(b, key);
            if (typeof va === 'number' && typeof vb === 'number') {{
                return dir === 'asc' ? va - vb : vb - va;
            }}
            return dir === 'asc' ? String(va).localeCompare(String(vb), 'ja') : String(vb).localeCompare(String(va), 'ja');
        }});

        rows.forEach(r => tbody.appendChild(r));
        setlistRows = rows;
        setlistSearchText = rows.map(r => r.getAttribute('data-search') || r.innerText.toUpperCase());
    }}
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
    print("HTML生成完了: index.html")
