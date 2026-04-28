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
    # URLを構築
    url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{filename}"
    print(f"[GitHub] Loading {filename} from {url}...")
    
    try:
        response = requests.get(url, timeout=30)
        
        if response.status_code == 200:
            content_bytes = response.content
            
            # 複数の文字コードで読み込みを試行
            encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
            
            for enc in encodings:
                try:
                    df = pd.read_csv(io.BytesIO(content_bytes), encoding=enc, engine='python', **kwargs)
                    
                    # カラム名のクリーニング
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
    """GASからCSVデータをダウンロードしてDataFrameとして返す（既存機能：履歴用）"""
    df, _ = load_df_from_gas_with_status(filename, **kwargs)
    return df

def save_df_to_gas(filename, df):
    """DataFrameをCSV文字列に変換してGASへアップロードする"""
    print(f"[GAS] Uploading {filename}...")
    try:
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        
        payload = {
            'filename': filename,
            'content': csv_buffer.getvalue()
        }
        
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

# --- 関数: テキスト正規化 ---
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
# 方針:
# 1) セトリ取得は履歴読み込み結果に関係なく必ず実行する
# 2) 取得できないポート・部屋は無視し、取得できた分だけ履歴へ追記する
# 3) 履歴読み込みに失敗した場合だけ、既存履歴を壊さないため history 系CSVへの保存を禁止する
# 4) 空データ保存・行数減少保存は禁止し、蓄積済みデータを消さない
# 5) 重複判定は内容一致を優先し、取得日が異なっても同一行の二重登録を防ぐ
# 6) 一定行数に達したら history_2.csv, history_3.csv ... へ移行し、以後は最新ファイルへ追記する

HISTORY_MAX_ROWS = 9500   # 1ファイルあたりの最大行数。この行数に達したら次の history_N.csv へ移行
HISTORY_ARCHIVE_MISS_LIMIT = 3  # 欠番があっても後続アーカイブを拾えるよう、3件連続未検出まで探索
IGNORE_FETCH_FAILURES = True  # True: 取得失敗ポートがあっても、取れた分だけ処理を続行する
ROOM_FETCH_TIMEOUT = 6        # 死んでいるポートで長時間止まらないよう短めにする
ROOM_FETCH_WORKERS = 16       # セトリ取得を並列化して、常に収集が進むようにする

# 基本は取得日込みで重複判定し、同日内の重複を確実に除去する
HISTORY_DEDUP_COLS = ['取得日', '部屋主', '順番', '曲名（ファイル名）', '歌った人']


def get_history_filename(num):
    """num=1 は history.csv、num>=2 は history_2.csv ..."""
    return "history.csv" if num == 1 else f"history_{num}.csv"


def sort_history_df(df):
    """履歴を表示・保存しやすい順に整える。元データは削らない。"""
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
    """順番の表示を整数優先に整え、小数点以下 .0 の表示を防ぐ。"""
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
    """CSVヘッダー行の混入などを除去。"""
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.fillna("")

    df = df.copy().fillna("")
    clean_check_cols = ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']
    for col in clean_check_cols:
        if col in df.columns:
            df = df[df[col].astype(str) != col]

    # GASエラー文言がカラム名として混入するケースを除去し、履歴列のみを残す
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
    """存在しない列は空文字として扱い、履歴比較用キーを作る。"""
    if df is None or df.empty:
        return pd.Series([], dtype=str)

    work = df.copy().fillna("")
    dedup_cols = HISTORY_DEDUP_COLS.copy()
    if for_history_compare and '取得日' in dedup_cols:
        # 収集日が異なる再取得（例: 25日のデータを28日に再収集）でも二重登録しない
        dedup_cols.remove('取得日')

    for col in dedup_cols:
        if col not in work.columns:
            work[col] = ""
    return work[dedup_cols].astype(str).agg("\u241f".join, axis=1)


def save_df_to_gas_checked(filename, df, min_existing_rows=0, allow_empty=False):
    """空保存・異常な行数減少を防ぎ、保存後に再読込で確認する。"""
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

    # 保存後チェック。ここで失敗しても既存ファイル削除はしていないため、履歴消失リスクは抑えられる。
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
    """history.csv, history_2.csv, history_3.csv ... を読み込む。読み込みエラー時は更新禁止。"""
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
    """1部屋分のHTMLテーブルを取得する。失敗は呼び出し側で無視して処理継続する。"""
    url = f"http://ykr.moe:{port}/simplelist.php"
    response = requests.get(url, timeout=ROOM_FETCH_TIMEOUT)
    response.raise_for_status()

    dfs = pd.read_html(io.BytesIO(response.content))
    if not dfs:
        raise ValueError("HTML内にテーブルが見つかりません")

    df = dfs[0].fillna("")
    if df.empty:
        raise ValueError("取得テーブルが空です")

    # エラーページ等をテーブルとして読んだ場合に備え、最低限のカラムを確認する
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

# 後続処理との互換用。final_df は最新の履歴ファイル相当として扱う。
final_df = history_records[-1]['df'] if history_records else pd.DataFrame()
archive_dfs = [h['df'] for h in history_records[:-1]] if len(history_records) > 1 else []

# --- 2. 新しいデータ取得 ---
target_ports = list(room_map.keys())
new_data_frames = []
failed_ports = []
fetched_ports = []
fetch_status = {}

# 重要:
# 履歴読み込みに失敗していても、セトリ収集自体は必ず動かす。
# ただし、履歴読み込みが不完全な時は既存履歴を壊す恐れがあるため history 系CSVへは保存しない。
if not history_load_ok:
    print("[Guard] 履歴の読み込みが不完全です。収集は実行しますが、履歴ファイルへの保存は禁止します。")

print("データを取得中...")

# 死んでいるポートで処理全体が止まらないよう、短いタイムアウト + 並列取得にする。
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
            fetch_status[port] = {
                "room": room_name,
                "status": "ok",
                "rows": len(df)
            }
            print(f"[Fetch] OK {port} ({room_name}) rows={len(df)}")

        except Exception as e:
            failed_ports.append((port, str(e)))
            fetch_status[port] = {
                "room": room_name,
                "status": "error",
                "error": str(e)
            }
            print(f"[Fetch] SKIP {port} ({room_name}) 取得失敗: {e}")

success_ports_count = len(fetched_ports)
print(f"[Fetch] 成功ポート: {success_ports_count} / {len(target_ports)}")

if failed_ports:
    print("[Fetch] 取得できなかったポートは無視します: " + ", ".join([f"{p}" for p, _ in failed_ports]))

# 取得できた分が1件もない場合は、既存履歴をそのまま使い、保存もしない。
if not new_data_frames:
    print("[Fetch] 取得成功データがありません。履歴ファイルは更新しません。")
    full_df = full_history_before_update

else:
    new_df = cleanup_history_df(pd.concat(new_data_frames, ignore_index=True))

    # 複数ポートが同じ部屋を指す場合、今回取得分の中で同じ曲が重複することがあるため除去する。
    dedup_cols = [c for c in HISTORY_DEDUP_COLS if c in new_df.columns]
    if dedup_cols:
        before_rows = len(new_df)
        new_df = new_df.drop_duplicates(subset=dedup_cols, keep='last')
        new_df = cleanup_history_df(new_df)
        print(f"[Dedup] 今回取得分の重複除去: {before_rows} -> {len(new_df)} 行")

    print(f"[Fetch] 今回収集できた行数: {len(new_df)} 行")

    # 履歴読み込みが不完全な時は、保存すると既存CSVを壊す可能性がある。
    # その場合でもダッシュボード表示・集計には今回取得分を使えるよう full_df には反映する。
    if not history_load_ok:
        print("[Guard] 履歴読み込み不完全のため、今回取得分は表示用のみ使用し、history 系CSVには保存しません。")
        if full_history_before_update.empty:
            full_df = new_df
        else:
            full_df = cleanup_history_df(pd.concat([full_history_before_update, new_df], ignore_index=True))

    else:
        # 既存履歴と同じ内容キーの行は保存対象から外す（取得日違いの再収集も重複として除外）。
        existing_keys = set(make_dedup_key(full_history_before_update, for_history_compare=True).tolist()) if not full_history_before_update.empty else set()
        new_keys = make_dedup_key(new_df, for_history_compare=True)
        new_unique_df = new_df[~new_keys.isin(existing_keys)].copy()
        new_unique_df = cleanup_history_df(new_unique_df)

        if new_unique_df.empty:
            print("追加対象の新規履歴はありません。履歴ファイルは更新しません。")
            full_df = full_history_before_update
        else:
            print(f"追加対象の新規履歴: {len(new_unique_df)} 行")

            # 最新の履歴ファイルへ追記。満杯なら history_2.csv, history_3.csv ... へ移行。
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

                # min_existing_rows を渡すので、既存ファイルより行数が減る保存は拒否される。
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

            # 表示・集計用に、保存済み分だけ反映する。失敗時も既存履歴は削らない。
            merged_history_by_num = {h['num']: h['df'] for h in history_records}
            merged_history_by_num.update(saved_parts)
            full_df = cleanup_history_df(pd.concat([merged_history_by_num[n] for n in sorted(merged_history_by_num)], ignore_index=True)) if merged_history_by_num else pd.DataFrame()

            if save_ok:
                print("履歴ファイルの安全更新が完了しました。")
            else:
                print("履歴ファイルの一部更新で停止しました。既存履歴は削除していません。")

# --- 全履歴データの結合（ローテーション済み history 系CSV 含む）---
if full_df is None or full_df.empty:
    full_df = pd.DataFrame()
else:
    full_df = cleanup_history_df(full_df)

print(f"全履歴データ合計: {len(full_df)} 行")

# ==========================================
# ==========================================
# ★集計処理
# ==========================================

cool_file = "cool_analysis.csv"
ALLOWED_CATEGORIES = ["2026年春アニメ", "2026年冬アニメ", "2025年秋アニメ"]

print(f"GitHubからオフラインリストを読み込みます... (User: {GITHUB_USER}, Repo: {GITHUB_REPO})")
offline_targets = []
for filename in OFFLINE_FILES:
    offline_df = load_df_from_github(filename)
    if not offline_df.empty and '曲名' in offline_df.columns:
        targets = [normalize_offline_text(str(x)) for x in offline_df.fillna("")['曲名'].tolist()]
        offline_targets.extend(targets)
        print(f"  -> {filename}: {len(targets)} 件追加")

raw_df = load_df_from_gas(cool_file, header=None)
embedded_categories = {cat: [] for cat in ALLOWED_CATEGORIES}
created_lists = {cat: [] for cat in ALLOWED_CATEGORIES}
uncreated_lists = {cat: [] for cat in ALLOWED_CATEGORIES}
ranking_base = []
trending_items = []

if not raw_df.empty:
    raw_df = raw_df.fillna("").drop_duplicates(keep='last')
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

    analysis_source_df['norm_workname'] = analysis_source_df.apply(get_rescued_workname, axis=1) if '作品名' in analysis_source_df.columns else ""
    exclude_keywords = ['test', 'テスト', 'システム', 'admin', 'System']
    full_history = analysis_source_df[(~analysis_source_df['歌った人'].astype(str).apply(lambda x: any(k in x for k in exclude_keywords)))].sort_values('dt_obj')

    start_date = pd.to_datetime("2026/01/01")
    end_date = pd.to_datetime("2026/06/30")
    target_history = full_history[(full_history['dt_obj'] >= start_date) & (full_history['dt_obj'] <= end_date)]

    categorized_data = {}
    current_category = None
    for _, row in raw_df.iterrows():
        if not any(str(x).strip() for x in row):
            continue
        col0 = str(row[0]).strip()
        is_category_line = any(cat in col0 for cat in ALLOWED_CATEGORIES) and "作品名" not in col0
        if is_category_line:
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
            categorized_data[current_category].append({"anime": anime, "type": type_, "artist": artist, "song": song})

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

            target_song_raw_norm = normalize_offline_text(item['song'])
            created = False
            if target_song_norm:
                for offline_str in offline_targets:
                    if (target_song_norm in offline_str) or (target_song_raw_norm in offline_str):
                        if (not target_anime_norm) or (target_anime_norm in offline_str):
                            created = True
                            break

            row = {"anime": item['anime'], "type": item['type'], "artist": item['artist'], "song": item['song'], "count": count, "users": users, "created": created, "category": category}
            embedded_categories[category].append(row)
            ranking_base.append(row)
            (created_lists[category] if created else uncreated_lists[category]).append(row)

    if not target_history.empty:
        max_dt = target_history['dt_obj'].max()
        recent_start = max_dt - pd.Timedelta(days=6)
        base_start = recent_start - pd.Timedelta(days=14)
        base_end = recent_start - pd.Timedelta(days=1)
        for category, items in embedded_categories.items():
            for item in items:
                song_norm = normalize_text(item['song'])
                anime_norm = normalize_text(item['anime'])
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
                    "anime": item['anime'], "song": item['song'], "artist": item['artist'], "type": item['type'],
                    "category": category, "recent": recent, "baseline": round(baseline, 2), "score": round(score, 4),
                    "users7d": int(recent_df['歌った人'].nunique()) if not recent_df.empty else 0, "isNew": baseline == 0
                })

trending_items.sort(key=lambda x: (x['score'], x['recent'], x['users7d']), reverse=True)
trending_items = trending_items[:30]

rankings_count = {cat: [] for cat in ALLOWED_CATEGORIES}
rankings_users = {cat: [] for cat in ALLOWED_CATEGORIES}
for cat in ALLOWED_CATEGORIES:
    items = [x for x in ranking_base if x['category'] == cat and x['count'] > 0]
    rankings_count[cat] = sorted(items, key=lambda x: (x['count'], x['users']), reverse=True)
    rankings_users[cat] = sorted(items, key=lambda x: (x['users'], x['count']), reverse=True)

setlist_records = []

def is_created(song, anime):
    sn = normalize_text(song)
    rn = normalize_offline_text(song)
    an = normalize_text(anime)
    if not sn:
        return False
    for offline_str in offline_targets:
        if (sn in offline_str) or (rn in offline_str):
            if (not an) or (an in offline_str):
                return True
    return False

if not full_df.empty:
    html_df = full_df.drop(columns=['コメント'], errors='ignore').fillna("").copy()
    html_df['_dt'] = pd.to_datetime(html_df.get('取得日', ''), errors='coerce')
    html_df['_ord'] = pd.to_numeric(html_df.get('順番', ''), errors='coerce').fillna(-1)
    html_df = html_df.sort_values(by=['_dt', '_ord'], ascending=[False, False], kind='mergesort')
    for _, row in html_df.iterrows():
        rec = {
            "room": str(row.get('部屋主', '')),
            "song": str(row.get('曲名（ファイル名）', '')),
            "work": str(row.get('作品名', '')),
            "artist": str(row.get('歌手名', '')),
            "singer": str(row.get('歌った人', '')),
            "order": str(row.get('順番', '')),
            "orderNum": float(row.get('_ord', -1)) if pd.notna(row.get('_ord', -1)) else -1,
            "fetchedAt": str(row.get('取得日', '')),
            "created": is_created(str(row.get('曲名（ファイル名）', '')), str(row.get('作品名', ''))),
        }
        rec['search'] = normalize_text(" ".join([str(rec[k]) for k in ["room", "song", "work", "artist", "singer", "order", "fetchedAt"]]))
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

html_content_template = """<!DOCTYPE html>
<html lang='ja'>
<head>
<meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Karaoke Dashboard</title>
<style>
:root{{--bg:#ffffff;--bg-soft:#f7f8fa;--line:#e3e6ec;--line-strong:#cfd4dc;--text:#1f2430;--text-sub:#5b6472;--text-mute:#8a93a1;--accent:#1e3a8a;--accent-soft:#eaf0ff;--warn:#b45309;--ok:#15803d;--row-h:34px;}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--text);font-family:-apple-system,'Segoe UI','Hiragino Sans','Yu Gothic UI',system-ui,sans-serif;font-size:13px;line-height:1.35}}
a{{color:var(--accent);text-decoration:none}}a:hover{{text-decoration:underline}}
.top{{position:sticky;top:0;background:#fff;border-bottom:1px solid var(--line);z-index:50}}
.head-row{{display:flex;gap:10px;align-items:center;padding:8px 12px;flex-wrap:wrap}}
.head-title{{font-size:15px;font-weight:700}}
.head-muted{{font-size:11.5px;color:var(--text-mute)}}
label.ctrl{{display:flex;gap:6px;align-items:center;font-size:11.5px;color:var(--text-sub)}}
input,select,button{{height:32px;border:1px solid var(--line-strong);background:#fff;border-radius:4px;padding:0 8px;font-size:13px}}
button{{cursor:pointer}}button.save{{color:#fff;background:var(--accent);border-color:var(--accent)}}
.tabs{{display:flex;overflow:auto;border-top:1px solid var(--line)}}
.tab-btn{{border:0;background:none;padding:9px 10px;font-weight:600;color:var(--text-sub);border-bottom:2px solid transparent;height:36px;white-space:nowrap}}
.tab-btn.active{{color:var(--accent);border-bottom-color:var(--accent)}}
.tab{{display:none;padding:8px 10px}}.tab.active{{display:block}}
.toolbar{{position:sticky;top:77px;background:#fff;border-bottom:1px solid var(--line);padding:6px 0;z-index:20;display:flex;gap:8px;align-items:center;flex-wrap:wrap}}
.seg{{display:flex;gap:6px;flex-wrap:wrap}}
.seg-btn{{height:28px;padding:0 8px;border:0;border-bottom:2px solid transparent;background:none;color:var(--text-sub)}}
.seg-btn.active{{color:var(--accent);border-bottom-color:var(--accent);font-weight:600}}
.count-txt{{margin-left:auto;font-size:11.5px;color:var(--text-mute)}}
.list-wrap{{height:72vh;overflow:auto;position:relative;border-top:1px solid var(--line)}}
.spacer{{width:1px;opacity:0}}.items{{position:absolute;left:0;right:0;top:0;will-change:transform}}
.set-row{{display:grid;grid-template-columns:44px 1fr 90px;grid-template-areas:'ord title date' 'ord sub room';column-gap:10px;row-gap:2px;padding:4px 10px;border-bottom:1px solid var(--line);min-height:34px}}
.set-row:hover{{background:var(--bg-soft)}}
.ord{{grid-area:ord;text-align:right;font-family:ui-monospace,monospace;font-weight:700;color:var(--text-sub)}}
.cr-tag{{display:block;font-size:11.5px;font-weight:700;line-height:1.1}}.cr-ok{{color:var(--ok)}}.cr-ng{{color:var(--warn)}}
.title{{grid-area:title;font-size:16px;font-weight:700;line-height:1.25;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.sub{{grid-area:sub;color:var(--text-sub);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.date{{grid-area:date;text-align:right;color:var(--text-mute);font-size:11.5px}}
.room{{grid-area:room;justify-self:end;max-width:140px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-left:3px solid var(--accent);padding-left:6px;color:var(--text-sub);font-size:11.5px}}
.section-note{{text-align:right;font-size:11.5px;color:var(--text-mute);padding:4px 0}}
.tbl{{border:1px solid var(--line);border-radius:6px;overflow:hidden}}
.tr{{display:grid;grid-template-columns:38px 1.1fr 64px 1.2fr 1fr 64px 64px;align-items:center;gap:8px;padding:6px 8px;border-bottom:1px solid var(--line)}}
.tr.hd{{position:sticky;top:113px;background:var(--bg-soft);border-bottom:1px solid var(--line-strong);font-size:11.5px;color:var(--text-sub);z-index:10}}
.tr:last-child{{border-bottom:none}}
.num{{text-align:right;font-family:ui-monospace,monospace}}
.num.zero{{color:var(--text-mute)}}.num.low{{color:var(--text-sub)}}.num.mid{{color:var(--text)}}.num.high{{color:var(--accent);font-weight:700}}
.stat-ok{{color:var(--ok);font-weight:700}}.stat-ng{{color:var(--warn);font-weight:700}}
.rank-row{{display:grid;grid-template-columns:38px 1fr 64px 64px;gap:8px;align-items:center;padding:6px 8px;border-bottom:1px solid var(--line);cursor:pointer}}
.rank-row.top1{{border-left:3px solid #c8a44b}}.rank-row.top2{{border-left:3px solid #9aa1ab}}.rank-row.top3{{border-left:3px solid #a06b3e}}
.rank-title{{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}} .rank-song{{font-weight:700}}
.trend-row{{display:grid;grid-template-columns:42px 1fr 1.2fr 1fr 56px 70px 70px 70px 70px 50px;gap:8px;align-items:center;padding:6px 8px;border-bottom:1px solid var(--line)}}
.trend-up{{color:var(--ok);font-weight:700}}.new-txt{{color:var(--ok);font-weight:700}}
@media (max-width:899px){{
  .toolbar{{top:111px}}
  .set-row{{grid-template-columns:36px 1fr;grid-template-areas:'ord title' 'ord sub' 'ord singer' 'ord meta';padding:8px 10px;min-height:88px;border:1px solid var(--line);border-radius:6px;margin:6px 0}}
  .date,.room{{display:none}}.singer{{grid-area:singer;color:var(--text-mute);font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
  .meta2{{grid-area:meta;text-align:right;color:var(--text-mute);font-size:11.5px}}
  .list-wrap{{height:70vh;border-top:none}}
  .tr,.tr.hd{{grid-template-columns:36px 1fr 56px 56px}}
  .tr.hd .hide-sm,.tr .hide-sm{{display:none}}
  .rank-row{{grid-template-columns:28px 1fr auto auto}}
  .trend-row{{grid-template-columns:28px 1fr auto}}
  .trend-row .hide-sm{{display:none}}
}}
</style>
</head>
<body>
<div class='top'>
  <div class='head-row'>
    <div class='head-title'>Karaoke Dashboard</div>
    <div class='head-muted'>__TIME__ 更新</div>
    <label class='ctrl'>Port <input id='exportPort' type='number' value='11059'></label>
    <label class='ctrl'>Link <select id='exportLinkType'><option value='eve'>Everything</option><option value='ykr'>ゆかりすたー</option></select></label>
  </div>
  <div class='tabs'>
    <button class='tab-btn active' data-tab='setlist'>セットリスト</button>
    <button class='tab-btn' data-tab='analysis'>クール集計</button>
    <button class='tab-btn' data-tab='ranking_count'>歌唱数ランキング</button>
    <button class='tab-btn' data-tab='ranking_user'>歌唱人数ランキング</button>
    <button class='tab-btn' data-tab='trending'>🔥急上昇</button>
  </div>
</div>

<div id='setlist' class='tab active'>
  <div class='toolbar'>
    <input id='searchInput' placeholder='検索'>
    <div class='seg' id='setSort'></div>
    <button id='saveSetlist' class='save'>HTML保存</button>
    <div class='count-txt' id='setCount'></div>
    <div style='width:100%'></div>
    <div class='seg' id='roomFilters'></div>
  </div>
  <div id='setWrap' class='list-wrap'><div id='setSpacer' class='spacer'></div><div id='setItems' class='items'></div></div>
</div>

<div id='analysis' class='tab'>
  <div class='toolbar'>
    <select id='anaCat'></select>
    <select id='anaState'><option value='all'>すべて</option><option value='created'>作成済み</option><option value='uncreated'>未作成</option><option value='has'>歌唱あり</option><option value='none'>未歌唱</option></select>
    <select id='anaSort'><option value='anime'>作品名</option><option value='count'>歌唱数↓</option><option value='users'>人数↓</option></select>
    <button id='saveCreated'>作成リスト保存</button><button id='saveUncreated'>未作成リスト保存</button><button id='saveAnalysis' class='save'>HTML保存</button>
  </div>
  <div class='section-note'>集計対象: 2026/01/01 - 2026/06/30</div>
  <div id='anaBody'></div>
</div>

<div id='ranking_count' class='tab'>
  <div class='toolbar'><button id='saveRankCount' class='save'>HTML保存</button></div>
  <div class='section-note'>集計対象: 2026/01/01 - 2026/06/30</div>
  <div id='rankCountBody'></div>
</div>

<div id='ranking_user' class='tab'>
  <div class='toolbar'><button id='saveRankUser' class='save'>HTML保存</button></div>
  <div class='section-note'>集計対象: 2026/01/01 - 2026/06/30</div>
  <div id='rankUserBody'></div>
</div>

<div id='trending' class='tab'>
  <div class='toolbar'><select id='trendCat'></select><select id='trendNew'><option value='all'>すべて</option><option value='new'>NEWのみ</option></select><button id='saveTrend' class='save'>HTML保存</button></div>
  <div id='trendBody'></div>
</div>

<script id='app-data' type='application/json'>__APP_JSON__</script>
<script>
const APP=JSON.parse(document.getElementById('app-data').textContent);const BP=APP.config.breakpoint||900;
const tabs=[...document.querySelectorAll('.tab-btn')];tabs.forEach(b=>b.onclick=()=>{{tabs.forEach(x=>x.classList.remove('active'));b.classList.add('active');document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));document.getElementById(b.dataset.tab).classList.add('active')}});
function host(){{const p=document.getElementById('exportPort').value||APP.config.defaultPort;return `http://ykr.moe:${{p}}`;}}function sp(){{return document.getElementById('exportLinkType').value==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword='}}function ykr(q){{return `${{host()}}/${{sp()}}${{encodeURIComponent(q)}}`;}}

const src=APP.setlist;const mql=window.matchMedia(`(max-width:${{BP-1}}px)`);let rowH=mql.matches?88:34;
const setSortDefs=[['date_desc','取得日↓'],['date_asc','取得日↑'],['order_desc','順番↓'],['song','曲名']];let setSort='date_desc';let roomSel=new Set();let setIdx=[];let tmr;
const rooms=[...new Set(src.map(x=>x.room).filter(Boolean))].sort();
function renderSetSort(){{document.getElementById('setSort').innerHTML=setSortDefs.map(([k,l])=>`<button class='seg-btn ${{k===setSort?'active':''}}' data-s='${{k}}'>${{l}}</button>`).join('')}}
renderSetSort();document.getElementById('setSort').onclick=e=>{{const b=e.target.closest('[data-s]');if(!b)return;setSort=b.dataset.s;renderSetSort();applySet();}};
document.getElementById('roomFilters').innerHTML=rooms.map(r=>`<button class='seg-btn' data-r='${{r}}'>${{r}}</button>`).join('');document.getElementById('roomFilters').onclick=e=>{{const b=e.target.closest('[data-r]');if(!b)return;const r=b.dataset.r;if(roomSel.has(r)){{roomSel.delete(r);b.classList.remove('active')}}else{{roomSel.add(r);b.classList.add('active')}}applySet();}};
function applySet(){{const kw=(document.getElementById('searchInput').value||'').trim().toUpperCase().split(/\s+/).filter(Boolean);setIdx=[];for(let i=0;i<src.length;i++){{const x=src[i];if(roomSel.size&&!roomSel.has(x.room))continue;let ok=true;for(const k of kw){{if(!x.search.includes(k)){{ok=false;break}}}}if(ok)setIdx.push(i)}}setIdx.sort((a,b)=>{{const A=src[a],B=src[b];if(setSort==='song')return A.song.localeCompare(B.song,'ja');if(setSort==='order_desc')return (B.orderNum||-1)-(A.orderNum||-1);if(setSort==='date_asc')return A.fetchedAt.localeCompare(B.fetchedAt,'ja')||((A.orderNum||-1)-(B.orderNum||-1));return B.fetchedAt.localeCompare(A.fetchedAt,'ja')||((B.orderNum||-1)-(A.orderNum||-1));}});document.getElementById('setCount').textContent=`全${{src.length}}件 / 表示${{setIdx.length}}件`;renderSetWindow();}}
const setWrap=document.getElementById('setWrap');setWrap.addEventListener('scroll',()=>requestAnimationFrame(renderSetWindow));
function renderSetWindow(){{const top=setWrap.scrollTop,vh=setWrap.clientHeight,buf=Math.max(6,Math.floor((vh/rowH)*0.5));const s=Math.max(0,Math.floor(top/rowH)-buf),e=Math.min(setIdx.length,Math.ceil((top+vh)/rowH)+buf);document.getElementById('setSpacer').style.height=`${{setIdx.length*rowH}}px`;const out=document.getElementById('setItems');out.style.transform=`translateY(${{s*rowH}}px)`;const mobile=mql.matches;out.innerHTML=setIdx.slice(s,e).map(i=>{{const x=src[i];const tag=x.created?`<span class='cr-tag cr-ok'>[済]</span>`:`<span class='cr-tag cr-ng'>[未]</span>`;if(mobile){{return `<div class='set-row'><div class='ord'>${{String(x.order||'').padStart(2,'0')}}${{tag}}</div><div class='title'><a target='_blank' href='${{ykr(`${{x.work}} ${{x.song}}`)}}'>${{x.song||'-'}}</a></div><div class='sub'>${{x.work||'-'}} ／ ${{x.artist||'-'}}</div><div class='singer'>${{x.singer||'-'}}</div><div class='meta2'>${{x.fetchedAt||'-'}}・${{x.room||'-'}}</div></div>`;}}return `<div class='set-row'><div class='ord'>${{String(x.order||'')}}${{tag}}</div><div class='title'><a target='_blank' href='${{ykr(`${{x.work}} ${{x.song}}`)}}'>${{x.song||'-'}}</a></div><div class='sub'>${{x.work||'-'}} ／ ${{x.artist||'-'}} ／ ${{x.singer||'-'}}</div><div class='date'>${{x.fetchedAt||'-'}}</div><div class='room'>${{x.room||'-'}}</div></div>`;}}).join('')}}
mql.addEventListener('change',()=>{{rowH=mql.matches?88:34;requestAnimationFrame(renderSetWindow);}});
document.getElementById('searchInput').addEventListener('input',()=>{{clearTimeout(tmr);tmr=setTimeout(()=>requestAnimationFrame(applySet),150)}});applySet();

const cats=['ALL',...APP.config.categories];document.getElementById('anaCat').innerHTML=cats.map(c=>`<option value='${{c}}'>${{c==='ALL'?'すべて':c}}</option>`).join('');document.getElementById('trendCat').innerHTML=cats.map(c=>`<option value='${{c}}'>${{c==='ALL'?'すべて':c}}</option>`).join('');
function numClass(v){{if(v===0)return 'zero';if(v<=2)return 'low';if(v<=9)return 'mid';return 'high';}}
function renderAnalysis(){{const cat=document.getElementById('anaCat').value,st=document.getElementById('anaState').value,so=document.getElementById('anaSort').value;let list=[];for(const [k,arr] of Object.entries(APP.categories)){{if(cat!=='ALL'&&k!==cat)continue;list=list.concat(arr.map(x=>({{...x,category:k}})))}}list=list.filter(x=>st==='all'||(st==='created'&&x.created)||(st==='uncreated'&&!x.created)||(st==='has'&&x.count>0)||(st==='none'&&x.count===0));list.sort((a,b)=>so==='anime'?a.anime.localeCompare(b.anime,'ja'):so==='count'?(b.count-a.count)||(b.users-a.users):((b.users-a.users)||(b.count-a.count)));let html='';for(const c of APP.config.categories){{const rows=list.filter(x=>x.category===c);if(!rows.length)continue;html+=`<details class='tbl' open><summary>${{c}}</summary><div class='tr hd'><div>作成</div><div>作品名</div><div class='hide-sm'>区分</div><div>曲名</div><div class='hide-sm'>歌手</div><div class='num'>人数</div><div class='num'>歌唱数</div></div>${{rows.map(r=>`<div class='tr'><div class='${{r.created?'stat-ok':'stat-ng'}}'>${{r.created?'済':'未'}}</div><div class='hide-sm'>${{r.anime}}</div><div class='hide-sm'>${{r.type}}</div><div><a target='_blank' href='${{ykr(`${{r.anime}} ${{r.song}}`)}}'>${{r.song}}</a></div><div class='hide-sm'>${{r.artist}}</div><div class='num ${{numClass(r.users)}}'>${{r.users}}</div><div class='num ${{numClass(r.count)}}'>${{r.count}}</div></div>`).join('')}}</details>`;}}document.getElementById('anaBody').innerHTML=html||'<div class="tbl" style="padding:10px">データなし</div>';}}
['anaCat','anaState','anaSort'].forEach(id=>document.getElementById(id).onchange=renderAnalysis);renderAnalysis();

function renderRanking(kind,target){{let html='';for(const c of APP.config.categories){{const arr=(APP.rankings[kind][c]||[]);let prev=null,rank=0;const rows=arr.map((x,i)=>{{const v=kind==='count'?x.count:x.users;if(v!==prev)rank=i+1;prev=v;return {{...x,rank}}});const top20=rows.filter(x=>x.rank<=20),rest=rows.filter(x=>x.rank>20);const draw=(r)=>r.map(x=>`<div class='rank-row ${{x.rank===1?'top1':x.rank===2?'top2':x.rank===3?'top3':''}}' data-q='${{x.anime}} ${{x.song}}'><div class='num'>#${{x.rank}}</div><div class='rank-title'>${{x.anime}} ／ <span class='rank-song'>${{x.song}}</span> ／ ${{x.artist}} (${{x.type}})</div><div class='num' style='${{kind==='users'?'font-weight:700;':''}}'>${{x.users}}</div><div class='num' style='${{kind==='count'?'font-weight:700;':''}}'>${{x.count}}</div></div>`).join('');html+=`<details class='tbl' open><summary>${{c}} TOP20</summary>${{draw(top20)}}${{rest.length?`<details><summary>もっと見る</summary>${{draw(rest)}}</details>`:''}}</details>`;}}document.getElementById(target).innerHTML=html;document.querySelectorAll('#'+target+' [data-q]').forEach(el=>el.onclick=()=>location.href=ykr(el.dataset.q));}}
renderRanking('count','rankCountBody');renderRanking('users','rankUserBody');

function renderTrend(){{const c=document.getElementById('trendCat').value,nm=document.getElementById('trendNew').value==='new';const arr=APP.trending.filter(x=>(c==='ALL'||x.category===c)&&(!nm||x.isNew));const hd=`<div class='trend-row' style='background:var(--bg-soft);font-size:11.5px;color:var(--text-sub);border-top:1px solid var(--line)'><div>順位</div><div>曲名</div><div class='hide-sm'>作品名</div><div class='hide-sm'>歌手</div><div class='hide-sm'>区分</div><div class='hide-sm num'>7日回</div><div class='hide-sm num'>7日人</div><div class='hide-sm num'>前2週</div><div>増加率</div><div>NEW</div></div>`;document.getElementById('trendBody').innerHTML=hd+arr.map((x,i)=>`<div class='trend-row'><div class='num'>#${{i+1}}</div><div><a target='_blank' href='${{ykr(`${{x.anime}} ${{x.song}}`)}}'>${{x.song}}</a></div><div class='hide-sm'>${{x.anime}}</div><div class='hide-sm'>${{x.artist}}</div><div class='hide-sm'>${{x.type}}</div><div class='hide-sm num'>${{x.recent}}</div><div class='hide-sm num'>${{x.users7d}}</div><div class='hide-sm num'>${{x.baseline}}</div><div class='trend-up'>+${{Math.round(x.score*100)}}%</div><div>${{x.isNew?'<span class="new-txt">[NEW]</span>':''}}</div></div>`).join('');}}
['trendCat','trendNew'].forEach(id=>document.getElementById(id).onchange=renderTrend);renderTrend();

function makeHtml(title,content){{const p=document.getElementById('exportPort').value||APP.config.defaultPort;const link=document.getElementById('exportLinkType').value==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword=';return `<!doctype html><html lang='ja'><head><meta charset='utf-8'><title>${{title}}</title><style>body{{font-family:-apple-system,'Segoe UI','Hiragino Sans','Yu Gothic UI',sans-serif;font-size:13px;color:#1f2430}}a{{color:#1e3a8a}}.tbl{{border:1px solid #e3e6ec;border-radius:6px;overflow:hidden}}.row{{display:grid;gap:8px;padding:6px 8px;border-bottom:1px solid #e3e6ec}}.num{{text-align:right;font-family:ui-monospace,monospace}}</style></head><body><h1>${{title}}</h1>${{content}}<script>const h='http://ykr.moe:${{p}}',s='${{link}}';document.querySelectorAll('[data-q]').forEach(x=>x.href=h+'/'+s+encodeURIComponent(x.dataset.q));<\/script></body></html>`;}}
function dl(name,title,content){{const b=new Blob([makeHtml(title,content)],{{type:'text/html'}});const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=name;a.click();}}
document.getElementById('saveSetlist').onclick=()=>{{const html=setIdx.map(i=>{{const x=src[i];return `<div class='row' style='grid-template-columns:44px 1fr 90px'><div class='num'>${{x.order}}</div><div><strong><a data-q='${{x.work}} ${{x.song}}'>${{x.song}}</a></strong><div>${{x.work}} ／ ${{x.artist}} ／ ${{x.singer}}</div><div>${{x.created?'済':'未'}}</div></div><div class='num'>${{x.fetchedAt}}</div></div>`}}).join('');dl('setlist.html','セットリスト',`<div class='tbl'>${{html}}</div>`);}};
document.getElementById('saveAnalysis').onclick=()=>dl('karaoke_analysis.html','クール集計',document.getElementById('anaBody').innerHTML);
document.getElementById('saveRankCount').onclick=()=>dl('karaoke_ranking_count.html','歌唱数ランキング',document.getElementById('rankCountBody').innerHTML);
document.getElementById('saveRankUser').onclick=()=>dl('karaoke_ranking_user.html','歌唱人数ランキング',document.getElementById('rankUserBody').innerHTML);
document.getElementById('saveTrend').onclick=()=>dl('karaoke_trending.html','急上昇',document.getElementById('trendBody').innerHTML);
document.getElementById('saveCreated').onclick=()=>{{const cat=document.getElementById('anaCat').value;let rows=[];for(const [k,v] of Object.entries(APP.createdLists)){{if(cat!=='ALL'&&cat!==k)continue;rows=rows.concat(v)}}dl('created_list.html','作成済みリスト',rows.map(x=>`<div><a data-q='${{x.anime}} ${{x.song}}'>${{x.song}}</a> ／ ${{x.anime}} ／ ${{x.artist}}</div>`).join(''));}};
document.getElementById('saveUncreated').onclick=()=>{{const cat=document.getElementById('anaCat').value;let rows=[];for(const [k,v] of Object.entries(APP.uncreatedLists)){{if(cat!=='ALL'&&cat!==k)continue;rows=rows.concat(v)}}dl('uncreated_list.html','未作成リスト',rows.map(x=>`<div><a data-q='${{x.anime}} ${{x.song}}'>${{x.song}}</a> ／ ${{x.anime}} ／ ${{x.artist}}</div>`).join(''));}};
</script>
</body></html>"""

html_content = html_content_template.replace("__APP_JSON__", app_json).replace("__TIME__", current_datetime_str)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
    print("HTML生成完了: index.html")
