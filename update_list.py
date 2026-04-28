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
if not full_df.empty:
    html_df = full_df.drop(columns=['コメント'], errors='ignore').fillna("")
    for _, row in html_df.iterrows():
        rec = {
            "room": str(row.get('部屋主', '')),
            "song": str(row.get('曲名（ファイル名）', '')),
            "work": str(row.get('作品名', '')),
            "artist": str(row.get('歌手名', '')),
            "singer": str(row.get('歌った人', '')),
            "order": str(row.get('順番', '')),
            "fetchedAt": str(row.get('取得日', '')),
        }
        rec['search'] = normalize_text(" ".join([rec[k] for k in ["room", "song", "work", "artist", "singer", "order", "fetchedAt"]]))
        setlist_records.append(rec)

embedded = {
    "updatedAt": current_datetime_str,
    "setlist": setlist_records,
    "categories": embedded_categories,
    "createdLists": created_lists,
    "uncreatedLists": uncreated_lists,
    "rankings": {"count": rankings_count, "users": rankings_users},
    "trending": trending_items,
    "config": {"categories": ALLOWED_CATEGORIES, "defaultPort": 11059}
}
app_json = json.dumps(embedded, ensure_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang='ja'>
<head>
<meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Karaoke Dashboard</title>
<style>
:root{{--accent:#2563eb;--line:#e5e7eb;--bg:#f8fafc;--text:#0f172a}}body{{margin:0;background:var(--bg);color:var(--text);font-family:-apple-system,'Segoe UI','Hiragino Sans','Yu Gothic UI',sans-serif;line-height:1.5}}.top{{position:sticky;top:0;background:#fff;border-bottom:1px solid var(--line);z-index:10}}.row{{padding:8px 12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}}.tabs{{display:flex;overflow:auto}}.tab-btn{{border:0;background:none;padding:10px 12px;font-weight:600;color:#64748b}}.tab-btn.active{{color:var(--accent);border-bottom:3px solid var(--accent)}}.tab{{display:none;padding:10px 12px}}.tab.active{{display:block}}.card{{background:#fff;border:1px solid var(--line);border-radius:8px;box-shadow:0 1px 2px rgba(0,0,0,.06);padding:10px;margin:8px 0}}.chips{{display:flex;gap:6px;flex-wrap:wrap}}.chip{{border:1px solid var(--line);border-radius:999px;padding:6px 10px;background:#fff;cursor:pointer;min-height:44px}}.chip.active{{border-color:var(--accent);color:var(--accent)}}input,select,button{{min-height:44px;border-radius:8px;border:1px solid #cbd5e1;padding:0 10px;background:#fff}}button.primary{{background:var(--accent);color:#fff;border-color:var(--accent)}}.small{{font-size:12px;color:#64748b}}.list-wrap{{height:65vh;overflow:auto;position:relative}}.spacer{{width:1px;opacity:0}}.items{{position:absolute;left:0;right:0;top:0}}.meta{{display:flex;justify-content:space-between;gap:8px;font-size:12px;color:#64748b}}.title{{font-weight:700;font-size:16px}}.muted{{color:#64748b;font-size:13px}}.detail{{display:none;margin-top:6px;border-top:1px dashed var(--line);padding-top:6px}}.card.open .detail{{display:block}}.badge{{display:inline-block;padding:2px 8px;border-radius:999px;background:#eef2ff;color:#1e3a8a;font-size:12px}}.metric{{font-family:ui-monospace,SFMono-Regular,Menlo,monospace}}.heat1{{background:#eff6ff}}.heat2{{background:#dbeafe}}.heat3{{background:#bfdbfe}}.heat4{{background:#93c5fd}}.heat5{{background:#60a5fa;color:#fff}}.rank-item{{border-left:4px solid transparent}}.rank-item.top1{{border-left-color:#d4af37}}.rank-item.top2{{border-left-color:#c0c0c0}}.rank-item.top3{{border-left-color:#cd7f32}}
</style>
</head>
<body>
<div class='top'><div class='row'><strong>Karaoke Dashboard</strong><span class='small'>{current_datetime_str} 更新</span><label>Port<input id='exportPort' type='number' value='11059'></label><label>Link<select id='exportLinkType'><option value='eve'>Everything</option><option value='ykr'>ゆかりすたー</option></select></label></div><div class='tabs'><button class='tab-btn active' data-tab='setlist'>セットリスト</button><button class='tab-btn' data-tab='analysis'>クール集計</button><button class='tab-btn' data-tab='ranking'>ランキング</button><button class='tab-btn' data-tab='trending'>🔥急上昇</button></div></div>
<div id='setlist' class='tab active'><div class='row'><input id='searchInput' placeholder='検索'><button id='saveSetlist' class='primary'>HTML保存</button></div><div class='row chips' id='setlistSort'></div><div class='row chips' id='roomFilters'></div><div class='small' id='setlistCount'></div><div id='setlistList' class='list-wrap'><div id='setlistSpacer' class='spacer'></div><div id='setlistItems' class='items'></div></div></div>
<div id='analysis' class='tab'><div class='row'><select id='analysisCategory'></select><select id='analysisState'><option value='all'>すべて</option><option value='created'>作成済みのみ</option><option value='uncreated'>未作成のみ</option><option value='has'>歌唱ありのみ</option><option value='none'>未歌唱のみ</option></select><select id='analysisSort'><option value='anime'>作品名</option><option value='count'>歌唱数↓</option><option value='users'>歌唱人数↓</option></select><button id='saveCreated'>作成リスト保存</button><button id='saveUncreated'>未作成リスト保存</button><button id='saveAnalysis' class='primary'>HTML保存</button></div><div id='analysisBody'></div></div>
<div id='ranking' class='tab'><div class='row'><select id='rankingMode'><option value='count'>歌唱数↓</option><option value='users'>歌唱人数↓</option></select><button id='saveRanking' class='primary'>HTML保存</button></div><div id='rankingBody'></div></div>
<div id='trending' class='tab'><div class='row'><select id='trendingCategory'></select><label><input id='trendingNewOnly' type='checkbox'>初登場のみ</label><button id='saveTrending' class='primary'>HTML保存</button></div><div id='trendingBody'></div></div>
<script id='app-data' type='application/json'>{app_json}</script>
<script>
const APP = JSON.parse(document.getElementById('app-data').textContent);
const tabs=[...document.querySelectorAll('.tab-btn')];tabs.forEach(b=>b.onclick=()=>{{tabs.forEach(x=>x.classList.remove('active'));b.classList.add('active');document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));document.getElementById(b.dataset.tab).classList.add('active')}});
function host(){{const p=document.getElementById('exportPort').value||APP.config.defaultPort;return `http://ykr.moe:${{p}}`;}} function searchPath(){{return document.getElementById('exportLinkType').value==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword='}} function ykr(word){{return `${{host()}}/${{searchPath()}}${{encodeURIComponent(word)}}`;}}
const sortOpts=[['date_desc','取得日↓'],['date_asc','取得日↑'],['room','部屋主'],['song','曲名']];let activeSort='date_desc';let roomSelected=new Set();let filteredIdx=[];let searchTimer=null;const rowH=152,buf=8;const src=APP.setlist;const rooms=[...new Set(src.map(x=>x.room).filter(Boolean))].sort();
function renderSortChips(){{document.getElementById('setlistSort').innerHTML=sortOpts.map(([k,l])=>`<button class='chip ${{k===activeSort?'active':''}}' data-sort='${{k}}'>${{l}}</button>`).join('')}}
renderSortChips();document.getElementById('setlistSort').onclick=(e)=>{{const b=e.target.closest('[data-sort]');if(!b)return;activeSort=b.dataset.sort;renderSortChips();applySetlist();}};
document.getElementById('roomFilters').innerHTML=rooms.map(r=>`<button class='chip' data-room='${{r}}'>${{r}}</button>`).join('');document.getElementById('roomFilters').onclick=(e)=>{{const b=e.target.closest('[data-room]');if(!b)return;const r=b.dataset.room;if(roomSelected.has(r)){{roomSelected.delete(r);b.classList.remove('active')}}else{{roomSelected.add(r);b.classList.add('active')}}applySetlist();}};
function applySetlist(){{const kw=(document.getElementById('searchInput').value||'').trim().toUpperCase().split(/\s+/).filter(Boolean);filteredIdx=[];for(let i=0;i<src.length;i++){{const it=src[i];if(roomSelected.size&& !roomSelected.has(it.room))continue;let ok=true;for(const k of kw){{if(!it.search.includes(k)){{ok=false;break}}}}if(ok)filteredIdx.push(i)}}filteredIdx.sort((a,b)=>{{const A=src[a],B=src[b];if(activeSort==='room')return A.room.localeCompare(B.room,'ja');if(activeSort==='song')return A.song.localeCompare(B.song,'ja');if(activeSort==='date_asc')return A.fetchedAt.localeCompare(B.fetchedAt,'ja');return B.fetchedAt.localeCompare(A.fetchedAt,'ja')}});document.getElementById('setlistCount').textContent=`全${{src.length}}件 / 表示${{filteredIdx.length}}件`;renderWindow();}}
const list=document.getElementById('setlistList');list.addEventListener('scroll',()=>requestAnimationFrame(renderWindow));
function renderWindow(){{const top=list.scrollTop;const vh=list.clientHeight;const s=Math.max(0,Math.floor(top/rowH)-buf);const e=Math.min(filteredIdx.length,Math.ceil((top+vh)/rowH)+buf);document.getElementById('setlistSpacer').style.height=`${{filteredIdx.length*rowH}}px`;const wrap=document.getElementById('setlistItems');wrap.style.transform=`translateY(${{s*rowH}}px)`;wrap.innerHTML=filteredIdx.slice(s,e).map(idx=>{{const x=src[idx];const q=`${{x.work}} ${{x.song}}`;return `<article class='card' onclick='this.classList.toggle("open")'><div class='meta'><span class='badge'>${{x.room||'-'}}</span><span>${{x.fetchedAt||''}}</span></div><div class='title'><a href='${{ykr(q)}}' target='_blank'>${{x.song||'-'}}</a></div><div class='muted'>${{x.work||'-'}} / ${{x.artist||'-'}}</div><div class='muted'>歌った人: ${{x.singer||'-'}} ／ 順番: ${{x.order||'-'}} ▼</div><div class='detail small'>取得日: ${{x.fetchedAt||'-'}}</div></article>`}}).join('')}}
document.getElementById('searchInput').addEventListener('input',()=>{{clearTimeout(searchTimer);searchTimer=setTimeout(()=>requestAnimationFrame(applySetlist),150)}});applySetlist();
function heat(v,max){{if(max<=0)return 'heat1';const r=v/max;return r>.8?'heat5':r>.6?'heat4':r>.4?'heat3':r>.2?'heat2':'heat1'}}
const cats=['ALL',...APP.config.categories];['analysisCategory','trendingCategory'].forEach(id=>document.getElementById(id).innerHTML=cats.map(c=>`<option value='${{c}}'>${{c==='ALL'?'すべて':c}}</option>`).join(''));
function renderAnalysis(){{const cat=document.getElementById('analysisCategory').value;const st=document.getElementById('analysisState').value;const so=document.getElementById('analysisSort').value;let works=[];for(const [k,arr] of Object.entries(APP.categories)){{if(cat!=='ALL'&&cat!==k)continue;const map={{}};for(const it of arr){{const key=`${{k}}||${{it.anime}}`;if(!map[key])map[key]={{category:k,anime:it.anime,count:0,users:0,songs:[]}};map[key].count+=it.count;map[key].users+=it.users;map[key].songs.push(it)}}works=works.concat(Object.values(map))}}works=works.filter(w=>st==='all'||(st==='created'&&w.songs.every(s=>s.created))||(st==='uncreated'&&w.songs.some(s=>!s.created))||(st==='has'&&w.count>0)||(st==='none'&&w.count===0));works.sort((a,b)=>so==='anime'?a.anime.localeCompare(b.anime,'ja'):so==='count'?b.count-a.count:b.users-a.users);const maxC=Math.max(...works.map(w=>w.count),1),maxU=Math.max(...works.map(w=>w.users),1);document.getElementById('analysisBody').innerHTML=works.map(w=>`<section class='card'><div class='meta'><strong>${{w.anime}}</strong><span class='badge'>${{w.category}}</span></div><div><span class='badge ${{heat(w.users,maxU)}}'>👥 ${{w.users}}</span> <span class='badge ${{heat(w.count,maxC)}}'>🎤 ${{w.count}}</span> <span class='badge'>楽曲 ${{w.songs.length}}</span></div><details><summary>詳細</summary>${{w.songs.map(s=>`<div class='small'><a target='_blank' href='${{ykr(`${{s.anime}} ${{s.song}}`)}}'>${{s.type}} ${{s.song}}</a> / ${{s.artist}} ${{s.created?'[作成済✓]':'[未作成]'}} 👥${{s.users}} 🎤${{s.count}}</div>`).join('')}}</details></section>`).join('')||'<div class="card">データなし</div>'}}
['analysisCategory','analysisState','analysisSort'].forEach(id=>document.getElementById(id).onchange=renderAnalysis);renderAnalysis();
function rankRows(mode,arr){{let out='';let prev=null,rank=0;arr.forEach((x,i)=>{{const v=mode==='count'?x.count:x.users;if(v!==prev)rank=i+1;prev=v;const top=rank===1?'top1':rank===2?'top2':rank===3?'top3':'';out+=`<div class='card rank-item ${{top}}'><div class='meta'><strong>#${{rank}} ${{x.anime}} — <a target='_blank' href='${{ykr(`${{x.anime}} ${{x.song}}`)}}'>${{x.song}}</a></strong><span class='metric'>👥 ${{x.users}}  🎤 ${{x.count}}</span></div><div class='small'>${{x.artist}} (${{x.type}})</div></div>`}});return out}}
function renderRanking(){{const mode=document.getElementById('rankingMode').value;let html='';for(const cat of APP.config.categories){{const arr=APP.rankings[mode][cat]||[];const top20=arr.slice(0,20);html+=`<section class='card'><details><summary><strong>${{cat}}</strong> TOP20</summary>${{rankRows(mode,top20)}}${{arr.length>20?`<details><summary>もっと見る</summary>${{rankRows(mode,arr.slice(20))}}</details>`:''}}</details></section>`}}document.getElementById('rankingBody').innerHTML=html}}
document.getElementById('rankingMode').onchange=renderRanking;renderRanking();
function renderTrending(){{const cat=document.getElementById('trendingCategory').value;const onlyNew=document.getElementById('trendingNewOnly').checked;const arr=APP.trending.filter(x=>(cat==='ALL'||x.category===cat)&&(!onlyNew||x.isNew));document.getElementById('trendingBody').innerHTML=arr.map((x,i)=>`<article class='card'><div class='meta'><strong>#${{i+1}} ${{x.anime}} — <a target='_blank' href='${{ykr(`${{x.anime}} ${{x.song}}`)}}'>${{x.song}}</a></strong><span class='badge'>${{x.category}}</span></div><div class='small'>${{x.artist}} (${{x.type}})</div><div class='metric'>直近7日: 🎤${{x.recent}} 👥${{x.users7d}} ／ 前2週平均: 🎤${{x.baseline}}</div><div>${{x.isNew?'<span class="badge">初登場</span>':''}} <span class='badge' style='background:#dcfce7;color:#166534'>+${{Math.round(x.score*100)}}%</span></div></article>`).join('')||'<div class="card">データなし</div>'}}
['trendingCategory','trendingNewOnly'].forEach(id=>document.getElementById(id).onchange=renderTrending);renderTrending();
function makeHtml(title,content){{const p=document.getElementById('exportPort').value||APP.config.defaultPort;const lt=document.getElementById('exportLinkType').value;const sp=lt==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword=';return `<!doctype html><html lang='ja'><head><meta charset='utf-8'><title>${{title}}</title><style>body{{font-family:-apple-system,Segoe UI,sans-serif;background:#fff;padding:12px}}.card{{border:1px solid #e5e7eb;border-radius:8px;padding:10px;margin:8px 0}}</style></head><body><h1>${{title}}</h1>${{content}}<script>const host='http://ykr.moe:${{p}}',sp='${{sp}}';document.querySelectorAll('a[data-q]').forEach(a=>a.href=host+'/'+sp+encodeURIComponent(a.dataset.q));<\/script></body></html>`;}}
function saveFile(name,title,content){{const b=new Blob([makeHtml(title,content)],{{type:'text/html'}});const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=name;a.click();}}
document.getElementById('saveSetlist').onclick=()=>{{const cards=filteredIdx.map(i=>{{const x=src[i];const q=`${{x.work}} ${{x.song}}`;return `<div class='card'><div>${{x.room}} / ${{x.fetchedAt}}</div><strong><a data-q='${{q}}'>${{x.song}}</a></strong><div>${{x.work}} / ${{x.artist}}</div><div>${{x.singer}} / ${{x.order}}</div></div>`}}).join('');saveFile('setlist.html','セットリスト',cards);}};
document.getElementById('saveAnalysis').onclick=()=>saveFile('karaoke_analysis.html','クール集計',document.getElementById('analysisBody').innerHTML);
document.getElementById('saveRanking').onclick=()=>saveFile('karaoke_ranking.html','ランキング',document.getElementById('rankingBody').innerHTML);
document.getElementById('saveTrending').onclick=()=>saveFile('karaoke_trending.html','急上昇',document.getElementById('trendingBody').innerHTML);
document.getElementById('saveCreated').onclick=()=>{{const cat=document.getElementById('analysisCategory').value;let rows=[];for(const [k,v] of Object.entries(APP.createdLists)){{if(cat!=='ALL'&&cat!==k)continue;rows=rows.concat(v)}}saveFile('created_list.html','作成済みリスト',rows.map(x=>`<div class='card'><a data-q='${{x.anime}} ${{x.song}}'>${{x.song}}</a> / ${{x.anime}} / ${{x.artist}}</div>`).join(''));}};
document.getElementById('saveUncreated').onclick=()=>{{const cat=document.getElementById('analysisCategory').value;let rows=[];for(const [k,v] of Object.entries(APP.uncreatedLists)){{if(cat!=='ALL'&&cat!==k)continue;rows=rows.concat(v)}}saveFile('uncreated_list.html','未作成リスト',rows.map(x=>`<div class='card'><a data-q='${{x.anime}} ${{x.song}}'>${{x.song}}</a> / ${{x.anime}} / ${{x.artist}}</div>`).join(''));}};
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
    print("HTML生成完了: index.html")
