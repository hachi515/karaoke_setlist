import pandas as pd
import requests
import datetime
import os
import re
import unicodedata
import json
import io
import hashlib
from itertools import groupby

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


def normalize_df_for_compare(df):
    """保存後検証用にDataFrameを文字列化して比較しやすくする"""
    if df is None:
        return pd.DataFrame()
    check_df = df.copy().fillna("")
    check_df.columns = check_df.columns.astype(str).str.replace('\ufeff', '').str.strip()
    check_df = check_df.astype(str)

    # pandasのCSV再読込で "1.0" と "1" のように揺れやすい値を軽く正規化
    def clean_value(v):
        v = v.strip()
        if re.fullmatch(r'-?\d+\.0', v):
            return v[:-2]
        return v

    for col in check_df.columns:
        check_df[col] = check_df[col].map(clean_value)
    return check_df


def df_content_hash(df):
    """DataFrameの行順・列順を含めた内容ハッシュを作る"""
    check_df = normalize_df_for_compare(df)
    csv_text = check_df.to_csv(index=False, lineterminator="\n")
    return hashlib.sha256(csv_text.encode("utf-8")).hexdigest()


def save_df_to_gas_verified(filename, df, allow_empty=False):
    """
    GASへ保存したあと、同じファイルを読み戻して行数と内容を検証する。
    検証NGの場合は False を返し、呼び出し側で history.csv の切り詰め・上書きを止める。
    """
    df_to_save = df.copy().fillna("") if df is not None else pd.DataFrame()

    if df_to_save.empty and not allow_empty:
        print(f"[Verify] {filename} は空DataFrameのため保存を中止しました。")
        return False

    if not save_df_to_gas(filename, df_to_save):
        return False

    read_df, status = load_df_from_gas_with_status(filename)
    if status != "ok":
        print(f"[Verify] {filename} の読み戻しに失敗しました: {status}")
        return False

    read_df = read_df.fillna("")

    if len(read_df) != len(df_to_save):
        print(f"[Verify] {filename} 行数不一致: save={len(df_to_save)}, read={len(read_df)}")
        return False

    save_norm = normalize_df_for_compare(df_to_save)
    read_norm = normalize_df_for_compare(read_df)

    if list(read_norm.columns) != list(save_norm.columns):
        print(f"[Verify] {filename} カラム不一致")
        print(f"  save columns: {list(save_norm.columns)}")
        print(f"  read columns: {list(read_norm.columns)}")
        return False

    if df_content_hash(read_norm) != df_content_hash(save_norm):
        print(f"[Verify] {filename} 内容ハッシュ不一致")
        return False

    print(f"[Verify] {filename} 保存検証OK: {len(df_to_save)} rows")
    return True


def make_timestamp_filename(prefix, ext="csv"):
    """バックアップ・退避用のタイムスタンプ付きファイル名を作る"""
    ts = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y%m%d_%H%M%S")
    return f"{prefix}_{ts}.{ext}"


def dataframe_has_required_columns(df, required_cols):
    """必要カラムが存在するか確認する"""
    return all(col in df.columns for col in required_cols)


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


# --- 1. 過去データ読み込み (history.csvはGASから) ---
# 安全方針:
# - history.csv / history_*.csv が「空」で読めた場合は、保存中・通信異常・Drive反映遅延の可能性があるため更新しない
# - アーカイブ保存後は必ず読み戻し検証して、検証OKのときだけ history.csv を切り詰める
# - history_backup_日時.csv / history_inbox_日時.csv のような一時退避ファイルは作成しない
# - 収集先サーバーが切断・空応答になった部屋は、履歴から前回値を表示用に補完し、削除されたように見せない
# - 補完した前回値は表示用だけに使い、history.csv へ「今日の取得分」として再保存しない

HISTORY_MAX_ROWS = 9500   # この行数以上になったらアーカイブを作成する
HISTORY_KEEP_ROWS = 8000  # アーカイブ後にhistory.csvに残す新しい行数
HISTORY_ARCHIVE_MISS_LIMIT = 3  # 欠番があっても後続アーカイブを拾えるように3件連続で未検出になるまで探索する

# 初回作成を許可する場合だけ True にする。通常運用では False 推奨。
ALLOW_INITIAL_HISTORY_CREATE = False

# 一時ファイルを作成しない設定。
# history_2.csv, history_3.csv... は一時ファイルではなく、行数超過時の正式アーカイブとして残す。
HISTORY_BACKUP_ENABLED = False
HISTORY_INBOX_BACKUP_ENABLED = False
REQUIRE_BACKUP_BEFORE_HISTORY_OVERWRITE = False

# 表示保護:
# 取得失敗した部屋がある場合、その部屋の最新履歴をセットリスト表示へ補完する。
# 補完データは history.csv には保存しない。
KEEP_PREVIOUS_LIST_ON_FETCH_FAILURE = True

history_file = "history.csv"
history_update_allowed = False
setlist_display_df = pd.DataFrame()

history_df, history_status = load_df_from_gas_with_status(history_file)

if history_status == "ok":
    history_df = history_df.fillna("")
    history_update_allowed = True

elif history_status == "not_found":
    if ALLOW_INITIAL_HISTORY_CREATE:
        print("history.csv が存在しないため、初回作成モードで新規作成します。")
        history_df = pd.DataFrame()
        history_update_allowed = True
    else:
        print("history.csv が存在しません。誤初期化防止のため更新を停止します。")
        history_df = pd.DataFrame()
        history_update_allowed = False

elif history_status == "empty":
    print("history.csv が空で読まれました。保存中または異常の可能性があるため、更新を停止します。")
    history_df = pd.DataFrame()
    history_update_allowed = False

else:
    print("history.csv の読み込みに失敗したため、更新を停止します。")
    history_df = pd.DataFrame()
    history_update_allowed = False


def get_latest_rows_for_rooms(source_df, room_names):
    """
    取得失敗時の表示補完用。
    指定された部屋主の最新取得日の行だけを history から取り出す。
    これは表示用であり、history.csv に再保存しない。
    """
    if source_df is None or source_df.empty or not room_names:
        return pd.DataFrame()

    required_cols = ['部屋主', '取得日']
    if not dataframe_has_required_columns(source_df, required_cols):
        print("[Fallback] 履歴に必要カラムがないため、前回値補完をスキップします。")
        return pd.DataFrame()

    fallback_df = source_df[source_df['部屋主'].isin(room_names)].copy()
    if fallback_df.empty:
        return pd.DataFrame()

    fallback_df['temp_date'] = pd.to_datetime(fallback_df['取得日'], errors='coerce')
    fallback_df = fallback_df.dropna(subset=['temp_date'])
    if fallback_df.empty:
        return pd.DataFrame()

    latest_dates = fallback_df.groupby('部屋主')['temp_date'].transform('max')
    fallback_df = fallback_df[fallback_df['temp_date'].eq(latest_dates)].copy()
    fallback_df = fallback_df.drop(columns=['temp_date'], errors='ignore')
    fallback_df['取得状態'] = '前回値（取得失敗のため保持）'
    return fallback_df.fillna("")


def sort_setlist_display(df):
    """セットリスト表示用に、部屋主・順番で見やすく並べる。"""
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy().fillna("")
    sort_cols = []
    ascending = []

    if '部屋主' in out.columns:
        sort_cols.append('部屋主')
        ascending.append(True)

    if '順番' in out.columns:
        out['順番'] = pd.to_numeric(out['順番'], errors='coerce')
        sort_cols.append('順番')
        ascending.append(True)

    if sort_cols:
        out = out.sort_values(by=sort_cols, ascending=ascending, na_position='last')

    return out.fillna("")


# --- アーカイブファイルの読み込み (history_2.csv, history_3.csv ...) ---
archive_dfs = []
archive_num = 2
loaded_archive_files = []
missing_archive_count = 0
max_archive_num = 1

while missing_archive_count < HISTORY_ARCHIVE_MISS_LIMIT:
    archive_file = f"history_{archive_num}.csv"
    archive_df, archive_status = load_df_from_gas_with_status(archive_file)

    if archive_status == "not_found":
        missing_archive_count += 1
        archive_num += 1
        continue

    # empty は「存在しない」扱いにしない。既存アーカイブを上書きして壊す危険があるため停止。
    if archive_status == "empty":
        print(f"{archive_file} が空で読まれました。既存アーカイブ保護のため更新を停止します。")
        history_update_allowed = False
        break

    if archive_status != "ok":
        print(f"{archive_file} の読み込みに失敗しました。既存アーカイブ保護のため更新を停止します。")
        history_update_allowed = False
        break

    archive_df = archive_df.fillna("")
    archive_dfs.append(archive_df)
    loaded_archive_files.append(archive_file)
    max_archive_num = archive_num
    missing_archive_count = 0
    archive_num += 1

next_archive_num = max_archive_num + 1

if loaded_archive_files:
    print(f"アーカイブファイルを読み込みました: {', '.join(loaded_archive_files)}")
else:
    print("アーカイブファイルなし。")


# --- 2. 新しいデータ取得 ---
target_ports = list(room_map.keys())
new_data_frames = []
fetch_status = {}

print("データを取得中...")
for port in target_ports:
    room_name = room_map[port]
    fetch_status[port] = {"room": room_name, "status": "pending", "rows": 0}
    url = f"http://Ykr.moe:{port}/simplelist.php"

    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()

        dfs = pd.read_html(response.content)

        if not dfs:
            fetch_status[port]["status"] = "empty"
            print(f"[Fetch] port {port} ({room_name}) 空応答")
            continue

        df = dfs[0].fillna("")
        df = df.replace(r'\s*詳細を見る ▼', '', regex=True)

        # ヘッダーだけ、または完全空テーブルは取得失敗扱いにして、前回値補完の対象にする。
        if df.empty:
            fetch_status[port]["status"] = "empty"
            print(f"[Fetch] port {port} ({room_name}) テーブル空")
            continue

        df['部屋主'] = room_name
        df['取得日'] = current_date_str
        new_data_frames.append(df)

        fetch_status[port]["status"] = "ok"
        fetch_status[port]["rows"] = len(df)

    except Exception as e:
        fetch_status[port]["status"] = "error"
        fetch_status[port]["error"] = str(e)
        print(f"[Fetch] port {port} ({room_name}) 取得失敗: {e}")


success_ports = [p for p, info in fetch_status.items() if info["status"] == "ok"]
failed_ports = [p for p, info in fetch_status.items() if info["status"] != "ok"]

success_rooms = {fetch_status[p]["room"] for p in success_ports}
failed_rooms = {fetch_status[p]["room"] for p in failed_ports}

# 同じ部屋主で複数ポートがある場合、1つでも成功していればその部屋は成功扱いにする。
fallback_rooms = sorted(failed_rooms - success_rooms)

print(f"[Fetch] 成功ポート: {len(success_ports)} / {len(target_ports)}")
if failed_ports:
    failed_summary = ", ".join([f"{p}:{fetch_status[p]['room']}({fetch_status[p]['status']})" for p in failed_ports])
    print(f"[Fetch] 取得失敗/空応答ポート: {failed_summary}")


if new_data_frames:
    new_df = pd.concat(new_data_frames, ignore_index=True).fillna("")
    print(f"新規取得データ: {len(new_df)} 行")

    # セットリスト表示は、今回取得できたデータを基本にする。
    # 取得失敗した部屋がある場合は、履歴から前回値を表示用だけに補完する。
    display_parts = [new_df.copy()]
    if KEEP_PREVIOUS_LIST_ON_FETCH_FAILURE and fallback_rooms:
        fallback_display_df = get_latest_rows_for_rooms(history_df, fallback_rooms)
        if not fallback_display_df.empty:
            display_parts.append(fallback_display_df)
            print(f"[Fallback] 取得失敗部屋の前回値を表示用に補完: {', '.join(fallback_rooms)} / {len(fallback_display_df)} 行")
        else:
            print(f"[Fallback] 取得失敗部屋の前回値が履歴から見つかりませんでした: {', '.join(fallback_rooms)}")

    setlist_display_df = sort_setlist_display(pd.concat(display_parts, ignore_index=True).fillna(""))

    # ここでは一時退避ファイル history_inbox_*.csv は作成しない。
    # 保存安全性は save_df_to_gas_verified() の読み戻し検証で担保する。

    combined_df = pd.concat([history_df, new_df], ignore_index=True)
    combined_df = combined_df.fillna("")

    clean_check_cols = ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']
    for col in clean_check_cols:
        if col in combined_df.columns:
            combined_df = combined_df[combined_df[col] != col]

    # 別日に同じ部屋・同じ順番・同じ曲・同じ人が歌った履歴を消さないため、取得日を重複キーに含める。
    subset_cols = ['取得日', '部屋主', '順番', '曲名（ファイル名）', '歌った人']

    if dataframe_has_required_columns(combined_df, subset_cols):
        before_dedup = len(combined_df)
        final_df = combined_df.drop_duplicates(subset=subset_cols, keep='first')
        print(f"重複除去: {before_dedup} -> {len(final_df)} 行")
    else:
        missing_cols = [c for c in subset_cols if c not in combined_df.columns]
        print(f"[Dedup] 必要カラム不足のため、キー指定重複除去を行いません: {missing_cols}")
        final_df = combined_df.drop_duplicates(keep='first')

    final_df = final_df.fillna("")

    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')

    # 余計な保存列を作らず、取得日と順番だけで並べる。
    # 以前の版で作られた history.csv に「取得日時」が残っている場合も、ここで削除する。
    # 表示補完用の「取得状態」も history.csv には保存しない。
    final_df = final_df.drop(columns=['取得日時', '取得状態'], errors='ignore')

    final_df['temp_date'] = pd.to_datetime(final_df['取得日'], errors='coerce') if '取得日' in final_df.columns else pd.NaT

    sort_cols = ['temp_date']
    ascending = [False]
    if '順番' in final_df.columns:
        sort_cols.append('順番')
        ascending.append(False)

    final_df = final_df.sort_values(by=sort_cols, ascending=ascending, na_position='last')
    final_df = final_df.drop(columns=['temp_date'], errors='ignore')

    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    # ここでは一時退避ファイル history_backup_*.csv は作成しない。
    # history.csv の上書き前バックアップは使わず、保存後の読み戻し検証のみ行う。

    if history_update_allowed:
        if len(final_df) >= HISTORY_MAX_ROWS:
            rows_to_keep = final_df.iloc[:HISTORY_KEEP_ROWS].copy()
            rows_to_archive = final_df.iloc[HISTORY_KEEP_ROWS:].copy()
            archive_filename = f"history_{next_archive_num}.csv"

            # 誤上書き防止: 保存予定のアーカイブ名が本当に未使用か確認。
            _, existing_archive_status = load_df_from_gas_with_status(archive_filename)

            if existing_archive_status != "not_found":
                print(f"[Archive] {archive_filename} は未使用ではありません。状態={existing_archive_status}。上書き防止のため更新を停止します。")
                history_update_allowed = False

            elif save_df_to_gas_verified(archive_filename, rows_to_archive):
                print(f"[Archive] {len(rows_to_archive)} 行を {archive_filename} に検証付きでアーカイブしました。")
                archive_dfs.append(rows_to_archive.fillna(""))
                loaded_archive_files.append(archive_filename)
                next_archive_num += 1
                final_df = rows_to_keep

            else:
                print(f"[Archive] {archive_filename} への退避または検証に失敗しました。history.csv は切り詰めません。")
                history_update_allowed = False

    if history_update_allowed:
        if save_df_to_gas_verified(history_file, final_df):
            print("history.csv を検証付きで更新しました。")
        else:
            print("history.csv の更新または検証に失敗しました。")
    else:
        print("履歴の保全を優先し、history.csv の更新をスキップしました。")

else:
    final_df = history_df

    # 全ポート失敗・空応答時は history.csv を一切更新しない。
    # セットリスト表示には履歴から各部屋の前回値を出す。
    all_rooms = sorted(set(room_map.values()))
    if KEEP_PREVIOUS_LIST_ON_FETCH_FAILURE:
        setlist_display_df = get_latest_rows_for_rooms(history_df, all_rooms)
        if not setlist_display_df.empty:
            print(f"[Fallback] 全ポート取得失敗のため、履歴から前回値を表示します: {len(setlist_display_df)} 行")
        else:
            setlist_display_df = history_df.copy()
            print("[Fallback] 前回値を抽出できなかったため、history.csv 全体を表示に使用します。")
    else:
        setlist_display_df = history_df.copy()

    setlist_display_df = sort_setlist_display(setlist_display_df)
    print("新しいデータなし。過去データを使用します。history.csv は更新しません。")


# --- 全履歴データの結合（アーカイブ含む）---
if archive_dfs:
    full_df = pd.concat([final_df] + archive_dfs, ignore_index=True)
    full_df = full_df.fillna("")
    print(f"全履歴データ合計: {len(full_df)} 行（history.csv + アーカイブ {len(archive_dfs)} ファイル）")
else:
    full_df = final_df

# 念のため、表示用データが空の場合は全履歴を表示に使う。
# 集計処理は引き続き full_df（history.csv + アーカイブ）を使う。
if setlist_display_df.empty:
    setlist_display_df = full_df.copy().fillna("")
else:
    setlist_display_df = setlist_display_df.fillna("")


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

# 集計対象カテゴリの定義
ALLOWED_CATEGORIES = ["2026年春アニメ", "2026年冬アニメ", "2025年秋アニメ"]

# --- HTMLコントロール用 プルダウン生成 ---
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


# --- ★関数: カテゴリ別リストHTML生成 ---
def generate_category_html_block(category_name, item_list):
    if not item_list:
        return ""
    
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
            
            link_tag_start = f'<a href="#search_link/{search_word}" class="export-link">'
            
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

        for category, items in categorized_data.items():
            
            cat_created_items = []
            cat_uncreated_items = []

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
                            <th style="width:8%; min-width:60px;">人数</th>
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

                    row_class = "has-count"
                    
                    bar_width = min(count * 20, 150)
                    bar_html = f'<div class="bar-chart" style="width:{bar_width}px;"></div>' if count > 0 else ""
                    
                    user_bar_width = min(user_count * 20, 100)
                    user_bar_html = f'<div class="bar-chart-user" style="width:{user_bar_width}px;"></div>' if user_count > 0 else ""

                    clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                    search_word = f"{clean_anime} {item['song']}"
                    
                    link_tag_start = f'<a href="#search_link/{search_word}" class="export-link">'
                    
                    analysis_html_content += f'<tr class="{row_class}">'
                    if i == 0:
                        analysis_html_content += f'<td rowspan="{rowspan}">{item["anime"]}</td>'
                    
                    analysis_html_content += f'<td align="center">{creation_count}</td>'
                    analysis_html_content += f'<td align="center">{link_tag_start}{item["type"]}</a></td>'
                    analysis_html_content += f'<td>{link_tag_start}{item["artist"]}</a></td>'
                    analysis_html_content += f'<td>{link_tag_start}{item["song"]}</a></td>'
                    analysis_html_content += f'<td class="count-cell"><div class="count-wrapper"><span class="count-num">{user_count}</span>{user_bar_html}</div></td>'
                    analysis_html_content += f'<td class="count-cell"><div class="count-wrapper"><span class="count-num">{count}</span>{bar_html}</div></td>'
                    analysis_html_content += '</tr>'
                
                analysis_html_content += '</tbody>'
            
            analysis_html_content += "</table></div></div>"

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

                html_out += f"""
                <div class="category-block">
                    <div class="category-header" onclick="toggleCategory(this)">
                        {rank_title} <i class="fas fa-chevron-down" style="float:right;"></i>
                    </div>
                    <div class="category-content">
                    <table class="rankingTable">
                        <thead>
                            <tr>
                                <th style="width:10%; min-width:60px;">順位</th>
                                <th style="width:25%; min-width:180px;">作品名</th>
                                <th style="width:25%; min-width:180px;">曲名</th>
                                <th style="width:15%; min-width:150px;">歌手</th>
                                <th style="width:10%; min-width:60px;">人数</th>
                                <th style="width:15%; min-width:60px;">歌唱数</th>
                            </tr>
                        </thead>
                        <tbody>
                """
                
                if not cat_items:
                    html_out += '<tr><td colspan="6" style="text-align:center; padding:20px;">歌唱データがありません</td></tr>'
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
                        
                        rank_class = f"rank-{current_rank}" if current_rank <= 3 else "rank-normal"
                        row_rank_class = f"rank-row-{current_rank}" if current_rank <= 3 else ""

                        rank_display = f'<span class="rank-badge {rank_class}">{current_rank}</span>'
                        
                        if current_rank == 1:
                            rank_display += ' <i class="fas fa-crown" style="color:#FFD700;"></i>'
                        elif current_rank == 2:
                            rank_display += ' <i class="fas fa-medal" style="color:#C0C0C0;"></i>'
                        elif current_rank == 3:
                            rank_display += ' <i class="fas fa-medal" style="color:#CD7F32;"></i>'
                            
                        bar_width = min(item["count"] * 20, 150)
                        bar_html = f'<div class="bar-chart" style="width:{bar_width}px;"></div>'

                        user_bar_width = min(item["user_count"] * 20, 100)
                        user_bar_html = f'<div class="bar-chart-user" style="width:{user_bar_width}px;"></div>' if item["user_count"] > 0 else ""

                        clean_anime = re.sub(r'[（\(].*?[）\)]', '', item['anime']).strip()
                        search_word = f"{clean_anime} {item['song']}"
                        
                        html_out += f"""
                        <tr class="has-count ranking-row {row_rank_class}" data-href="#search_link/{search_word}">
                            <td align="center" style="font-weight:bold; font-size:1.1rem;">{rank_display}</td>
                            <td>{item["anime"]} <span style="font-size:0.8em; color:#777;">({item["type"]})</span></td>
                            <td>{item["song"]}</td> <td>{item["artist"]}</td>
                            <td class="count-cell"><div class="count-wrapper"><span class="count-num">{item["user_count"]}</span>{user_bar_html}</div></td>
                            <td class="count-cell"><div class="count-wrapper"><span class="count-num">{item["count"]}</span>{bar_html}</div></td>
                        </tr>
                        """
                        
                html_out += "</tbody></table></div></div>"
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
if not setlist_display_df.empty:
    html_df = setlist_display_df.drop(columns=columns_to_hide, errors='ignore')
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

        .port-input-wrapper {{
            display: inline-flex;
            align-items: center;
            gap: 5px;
            margin-left: 15px;
            font-size: 13px;
            font-weight: bold;
            color: var(--primary-color);
        }}
        .port-input-wrapper input, .port-input-wrapper select {{
            padding: 3px 5px;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-family: monospace;
        }}
        .port-input-wrapper input {{
            width: 70px;
            text-align: center;
        }}

        .tabs {{
            display: flex; padding: 0 15px; border-bottom: 1px solid var(--border-color); overflow-x: auto;
        }}
        .tab-btn {{
            padding: 10px 20px; cursor: pointer; border: none; background: none;
            font-weight: bold; color: #7f8c8d; border-bottom: 3px solid transparent;
            font-size: 14px; white-space: nowrap;
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
        .ctrl-ranking {{ display: none; width: 100%; align-items: center; justify-content: flex-end; gap:5px; }}
        .ctrl-graph {{ display: none; width: 100%; align-items: center; justify-content: flex-end; }}

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
        
        tr.has-count {{ background-color: #fff; color: #333; }}
        
        .count-wrapper {{ display: flex; align-items: center; gap: 8px; }}
        .count-num {{ width: 25px; text-align: right; font-size:1.1rem; }}
        .bar-chart {{
            height: 10px; background: linear-gradient(90deg, #3498db, #2980b9);
            border-radius: 5px;
        }}

        .bar-chart-user {{
            height: 10px; background: linear-gradient(90deg, #2ecc71, #27ae60);
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
        
        tr.rank-row-1 td {{ background-color: #fff8e1 !important; }}
        tr.rank-row-2 td {{ background-color: #f5f5f5 !important; }}
        tr.rank-row-3 td {{ background-color: #fff0e6 !important; }}

        .rankingTable tr:nth-child(1) th {{ background-color: var(--primary-color) !important; color: #fff !important; }}

        .chart-wrapper {{
            background: #fff;
            padding: 10px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            margin-top: 15px;
            height: 75vh;
            display: flex;
            flex-direction: column;
        }}
        .chart-info {{
            min-height: 35px;
            height: auto;
            line-height: 1.4;
            padding: 5px;
            text-align: center;
            font-weight: bold;
            color: #2c3e50;
            background: #f8f9fa;
            border: 1px solid #e0e0e0;
            margin-bottom: 5px;
            border-radius: 4px;
            font-size: 14px;
            white-space: normal;
            overflow: visible;
            word-break: break-all;
        }}
        .canvas-container {{
            flex: 1;
            position: relative;
            min-height: 0;
        }}

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
            .chart-wrapper {{ height: auto; }}
        }}
    </style>
</head>
<body>
    <div class="top-section">
        <div class="header-inner">
            <div style="display:flex; align-items:center;">
                <h1>Karaoke Dashboard</h1>
                <div class="port-input-wrapper">
                    <label for="exportPort"><i class="fas fa-network-wired"></i> 保存時ポート:</label>
                    <input type="number" id="exportPort" value="11059" title="HTML保存時のURLポート番号を指定">
                </div>
                <div class="port-input-wrapper" style="margin-left:20px;">
                    <label for="exportLinkType"><i class="fas fa-link"></i> 検索リンク:</label>
                    <select id="exportLinkType">
                        <option value="eve">Everything</option>
                        <option value="ykr">ゆかりすたー</option>
                    </select>
                </div>
            </div>
            <div class="update-time">{current_datetime_str} 更新</div>
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
                <input type="text" id="searchInput" class="search-box" placeholder="キーワード (例: 曲名 歌手)...">
                <button onclick="performSearch()" class="btn"><i class="fas fa-search"></i> 検索</button>
                <button onclick="resetFilter()" class="btn" style="background:#95a5a6"><i class="fas fa-undo"></i></button>
                <div class="count-display" id="countDisplay">読み込み中...</div>
            </div>
            <div id="ctrl-analysis" class="ctrl-analysis">
                <select id="exportTargetCategory" style="margin-right:10px; padding:4px; border-radius:4px; font-size:13px; border: 1px solid #ccc;">
                    {category_options}
                </select>
                <button onclick="downloadListWithCategory('list-created-content', 'created_list.html', '作成済みリスト')" class="btn btn-list">作成リスト保存</button>
                <button onclick="downloadListWithCategory('list-uncreated-content', 'uncreated_list.html', '未作成リスト')" class="btn btn-list" style="background-color:#e74c3c;">未作成リスト保存</button>
                <button onclick="downloadHTMLWithCategory()" class="btn btn-dl" style="margin-left:10px;"><i class="fas fa-file-code"></i> HTML保存</button>
            </div>
            <div id="ctrl-ranking-count" class="ctrl-ranking">
                <select id="exportTargetRankingCount" style="margin-right:10px; padding:4px; border-radius:4px; font-size:13px; border: 1px solid #ccc;">
                    {category_options}
                </select>
                <button onclick="downloadRankingWithCategory('count')" class="btn btn-dl"><i class="fas fa-trophy"></i> 歌唱数ランキング保存</button>
            </div>
            <div id="ctrl-ranking-user" class="ctrl-ranking">
                <select id="exportTargetRankingUser" style="margin-right:10px; padding:4px; border-radius:4px; font-size:13px; border: 1px solid #ccc;">
                    {category_options}
                </select>
                <button onclick="downloadRankingWithCategory('user')" class="btn btn-dl"><i class="fas fa-users"></i> 歌唱人数ランキング保存</button>
            </div>
            <div id="ctrl-graph" class="ctrl-graph">
                <button onclick="downloadGraphHTML()" class="btn btn-dl" style="background-color:#e67e22;"><i class="fas fa-file-code"></i> HTML保存</button>
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
            <div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/06/30</div>
            <div id="print-target">
                {analysis_html_content if cool_data_exists else '<div style="padding:20px;text-align:center;color:#e74c3c;">集計データがありません</div>'}
            </div>
        </div>

        <div id="ranking_count" class="tab-content">
            <div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/06/30</div>
            <div id="ranking-count-print-target">
                {ranking_count_html_content if ranking_count_html_content else '<div style="padding:20px;text-align:center;color:#e74c3c;">ランキング対象データがありません</div>'}
            </div>
        </div>
        
        <div id="ranking_user" class="tab-content">
            <div style="margin-top:15px; font-size:0.9rem; color:#7f8c8d; text-align:right;">集計対象: 2026/01/01 - 2026/06/30</div>
            <div id="ranking-user-print-target">
                {ranking_user_html_content if ranking_user_html_content else '<div style="padding:20px;text-align:center;color:#e74c3c;">ランキング対象データがありません</div>'}
            </div>
        </div>

        <div id="graph_view_count" class="tab-content">
            <div class="category-header">2026年春アニメ 歌唱数ランキング推移 (Top 20)</div>
            <div class="chart-wrapper">
                <div id="chart-info-count" class="chart-info">グラフの点をタップ・ホバーで詳細を表示</div>
                <div class="canvas-container"><canvas id="rankingChartCount"></canvas></div>
            </div>
        </div>
        <div id="graph_view_user" class="tab-content">
            <div class="category-header">2026年春アニメ 歌唱人数ランキング推移 (Top 20)</div>
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
    
    // --- グラフ用データ ---
    const dataCount = {graph_json_count};
    const dataUser = {graph_json_user};
    let charts = {{ count: null, user: null }};
    
    // 標準カラーパレット
    const colors = [
        '#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#46f0f0', '#f032e6', 
        '#bcf60c', '#fabebe', '#008080', '#e6beff', '#9a6324', '#fffac8', '#800000', '#aaffc3', 
        '#808000', '#ffd8b1', '#000075', '#808080'
    ];

    function initChart(type, dataObj, canvasId) {{
        if(charts[type]) return;
        const ctx = document.getElementById(canvasId).getContext('2d');
        const infoDivId = type === 'count' ? 'chart-info-count' : 'chart-info-user';
        
        // 最新の順位でTOP5を判定
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
                pointRadius: 4, 
                pointHoverRadius: 8, 
                tension: 0.1, 
                fill: false, 
                borderWidth: 2,
                hidden: !isTop5 // TOP5以外は初期非表示
            }};
        }});

        charts[type] = new Chart(ctx, {{
            type: 'line', 
            data: {{ datasets }},
            options: {{
                responsive: true, 
                maintainAspectRatio: false,
                interaction: {{
                    mode: 'nearest',
                    axis: 'x',
                    intersect: true
                }},
                plugins: {{
                    tooltip: {{
                        enabled: false,
                        external: function(context) {{
                            const tooltip = context.tooltip;
                            const infoDiv = document.getElementById(infoDivId);
                            if (tooltip.opacity === 0) return;
                            
                            if (tooltip.body) {{
                                const dataPoint = tooltip.dataPoints[0];
                                const dateObj = new Date(dataPoint.label);
                                const dateStr = (dateObj.getMonth() + 1) + '/' + dateObj.getDate();
                                infoDiv.innerHTML = `<span style="color:${{dataPoint.dataset.borderColor}}">●</span> ${{dataPoint.dataset.label}}　${{dateStr}}（${{dataPoint.parsed.y}}位）`;
                            }}
                        }}
                    }},
                    legend: {{ 
                        position: 'bottom',
                        labels: {{ boxWidth: 10, padding: 15 }},
                        onClick: function(e, legendItem, legend) {{
                            const index = legendItem.datasetIndex;
                            const ci = legend.chart;
                            if (ci.isDatasetVisible(index)) {{
                                ci.hide(index);
                                legendItem.hidden = true;
                            }} else {{
                                ci.show(index);
                                legendItem.hidden = false;
                            }}
                        }}
                    }}
                }},
                scales: {{
                    y: {{ 
                        reverse: true, 
                        min: 0.5, 
                        max: 20.5, 
                        ticks: {{ 
                            stepSize: 1, 
                            callback: function(val) {{ 
                                if (val % 1 === 0 && val >= 1 && val <= 20) return val;
                                return ''; 
                            }} 
                        }},
                        title: {{ display: true, text: '順位' }}
                    }},
                    x: {{ 
                        type: 'time', 
                        time: {{ unit: 'day', displayFormats: {{ day: 'M/d' }} }},
                        title: {{ display: true, text: '日付' }}
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
            <div class="category-header">${{headerText}}</div>
            <div class="chart-wrapper">
                <img src="${{imgData}}" style="width:100%; max-width:800px; border:1px solid #ccc; display:block; margin:0 auto;">
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
        document.querySelector('.ctrl-graph').style.display = 'none';

        if(tabName === 'setlist') document.getElementById('ctrl-setlist').style.display = 'flex';
        else if(tabName === 'analysis') document.getElementById('ctrl-analysis').style.display = 'flex';
        else if(tabName === 'ranking_count') document.getElementById('ctrl-ranking-count').style.display = 'flex';
        else if(tabName === 'ranking_user') document.getElementById('ctrl-ranking-user').style.display = 'flex';
        else if(tabName === 'graph_view_count') {{
            document.querySelector('.ctrl-graph').style.display = 'flex';
            initChart('count', dataCount, 'rankingChartCount');
        }}
        else if(tabName === 'graph_view_user') {{
            document.querySelector('.ctrl-graph').style.display = 'flex';
            initChart('user', dataUser, 'rankingChartUser');
        }}
    }}

    function toggleCategory(header) {{
        const content = header.nextElementSibling;
        content.classList.toggle('collapsed');
        const icon = header.querySelector('i');
        if(icon) {{
            icon.className = content.classList.contains('collapsed') ? 'fas fa-chevron-right' : 'fas fa-chevron-down';
            icon.style.float = 'right';
        }}
    }}

    // --- HTML出力用関数群 ---
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
        const blocks = container.querySelectorAll('.category-block');
        let content = "";
        blocks.forEach(block => {{
            const header = block.querySelector('.category-header');
            if (header && header.innerText.includes(targetCat)) {{
                // 出力時にcollapsedを解除して表示させる
                const clone = block.cloneNode(true);
                const catContent = clone.querySelector('.category-content');
                if (catContent) catContent.classList.remove('collapsed');
                
                const icon = clone.querySelector('i.fa-chevron-right');
                if (icon) icon.className = 'fas fa-chevron-down';

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

    // HTML生成時に選択したリンク・ポート情報を埋め込む
    function generateDownload(content, filename, title) {{
        const portValue = document.getElementById('exportPort').value || '11059';
        const linkType = document.getElementById('exportLinkType').value;
        const searchPath = linkType === 'ykr' ? 'search_listerdb_filelist.php?anyword=' : 'search.php?searchword=';
        
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
        
        /* プレーンなリンクスタイル */
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
        .bar-chart-user {{ height: 10px; background: #2ecc71; border-radius: 5px; }}
        
        .rank-badge {{
            display: inline-block; width: 24px; height: 24px; line-height: 24px;
            border-radius: 50%; text-align: center; color: #fff; font-weight: bold; font-size: 12px;
            background-color: #95a5a6;
        }}
        .rank-1 {{ background-color: #f1c40f; width: 28px; height: 28px; line-height: 28px; }}
        .rank-2 {{ background-color: #bdc3c7; }}
        .rank-3 {{ background-color: #d35400; }}
        
        tr.rank-row-1 td {{ background-color: #fff8e1 !important; }}
        tr.rank-row-2 td {{ background-color: #f5f5f5 !important; }}
        tr.rank-row-3 td {{ background-color: #fff0e6 !important; }}

        .chart-wrapper {{
            background: #fff; padding: 10px; border-radius: 8px; border: 1px solid #ccc;
        }}

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
        const host = 'http://ykr.moe:${{portValue}}';
        const searchPath = '${{searchPath}}';

        document.addEventListener('DOMContentLoaded', () => {{
            // クール集計・リスト類のリンク設定
            document.querySelectorAll('a.export-link').forEach(link => {{
                const rawHref = link.getAttribute('href');
                if (rawHref && rawHref.startsWith('#search_link/')) {{
                    const word = rawHref.split('#search_link/')[1];
                    link.href = host + '/' + searchPath + word;
                }}
            }});
            
            // ランキング等の行クリック設定
            document.querySelectorAll('tr[data-href]').forEach(row => {{
                row.addEventListener('click', () => {{
                    if (window.getSelection().toString().length > 0) return;
                    const rawHref = row.getAttribute('data-href');
                    if (rawHref && rawHref.startsWith('#search_link/')) {{
                        const word = rawHref.split('#search_link/')[1];
                        window.location.href = host + '/' + searchPath + word;
                    }}
                }});
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
