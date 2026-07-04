# -*- coding: utf-8 -*-
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
GITHUB_USER = "hachi515"
GITHUB_REPO = "karaoke_setlist"
GITHUB_BRANCH = "main"

OFFLINE_FILES = [
    "offline_list_2026_1st.csv",
    "offline_list_2025_2nd.csv",
    "offline_list_2025_1st.csv"
]

GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyzKEPfj0bYcRyEdizwQXcduIOQFt2_njtFQSyGP9jBjrhR8pyVKwDol6VN7bLPrktq/exec"
CSV_EMPTY_PREFIX_BYTES = b'\xef\xbb\xbf\r\n\t '
EXPECTED_HISTORY_COLUMNS = ['取得日', '部屋主', '順番', '曲名（ファイル名）', '作品名', '歌手名', '歌った人']

ALLOWED_CATEGORIES = ["2026年夏アニメ", "2026年春アニメ", "2026年冬アニメ", "2025年秋アニメ"]
TREND_TARGET_CATEGORY = "2026年夏アニメ"
TREND_PERIOD_OPTIONS = [3, 7, 14, 30]
TREND_PICKUP_TOTAL = 10

ASS_TAGS_FILENAME = "ass_tag_extensions.txt"
ASS_TAG_DEFAULT_TEXT = r"""【しあわせフレーバー（くるくる1回転イン、ワイプ後に文字大きくする）】
行：{\org([%W/2],[%y100])\fsp0\fs[%f-5]\alpha&HFF&}{\fr0.05\t([%n3]0,[%n3+20]1,\fr0\fs[%f]\alpha&H00&)}
文字：{\org([%x+%w/2],[%y])\fr10\fry360\alpha&HFF&\move(%x,[%y+40],%x,%y,[%n6]0,[%n6+80]0,23)\t([%n6]0,[%n*6+50]0,0.5,\fr0\fry0\alpha&H00&)}{\t(%s,%s,\fs[%f+15])\t(%s,%e,\fs%f)}

【春夏秋冬（文字単位螺旋イン、ワイプ後に文字縮小して↑方向にアウト）】
文字：{\org([%x+%w/2],[%y])\frx0\fry0\frz0\fscx100\fscy100\alpha&HFF&\move(%x,[%y+30],%x,%y,[%n25],[%n25+500],7)\t([%n25],[%n25+260],0.45,\alpha&H00&)}
{\t([%V+420+%n25],[%V+560+%n25],23,\alpha&H1A&\fsp-7)}
{\move(%x,%y,%x,[%y-70],[%V+560+%n25],[%V+980+%n25],7)}
{\t([%V+560+%n25],[%V+980+%n25],7,\fscx78\fscy78\3a&HFF&\blura&H00&)}
{\t([%V+820+%n25],[%V+980+%n25],1,\alpha&HFF&)}

【氷の城壁（文字単位螺旋イン、ワイプ後にくるくる回転アウト）】
文字：{\org([%x+%w/2],[%y])\frx0\fry0\frz0\fscx100\fscy100\alpha&HFF&\move(%x,[%y+40],%x,%y,[%n6]0,[%n6+80]0,23)\t([%n6]0,[%n6+40]0,0.45,\alpha&H00&)}{\t([%e+470],[%e+1170],1,\frx0\fry-270\frz0\fscx78\fscy78\alpha&HFF&\3a&HFF&\4a&HFF&)}

【青空ジャンピングハート】
行：{\org([%W/2],[%y100])\fsp0\fs[%f-5]\alpha&HFF&}{\fr0.05\t([%n3]0,[%n*3+20]1,\fr0\fsp0\fs[%f]\alpha&H00&)}
文字：{\org([%w/2+%x],[%y])}{\fr10\fry990\alpha&HFF&\t([%n3]0,[%n3+80]1,0.5,\fr0\fry0\alpha&H00&)}{\t(%s,%s,\fs[%f+15])\t(%s,%e,\fs%f)}{\move([%x-0],%y,[%x-50],[%y+50],[%e+200],[%e+1000])}{\t([%e+200],[%e+1000],\bord3\frz180\fry270\fs10\1a&Hff&\3a&Hff&\4a&Hff&)}

【北斗の拳】
{\org([%x+%w/2],[%y])\fad(200,200)
\fscx100\fscy100

\t([%s],[%s+200],0.5,\fscx90\fscy90)
\t([%e-200],[%e],0.5,\fscx100\fscy100)

\t([%V+80],[%V+160],1,\fscx80\fscy80\1c&H888888&)
\t([%V+160],[%V+280],0.5,\fscx150\fscy150\1c&HFFFFFF&\3c&HFFFFFF&\4c&HFFFFFF&\glow180\gblur6)

\move(%x,%y,
[%x + ((%n37)%17-8)85 + ((%n13)%5-2)30],
[%y - 100 + ((%n29)%11-5)40],
[%V+200],[%V+1000],23)

\t([%V+160],[%V+1000],1,
\fscx[60+((%n17)%40)]\fscy[60+((%n23)%40)]
\frz[((%n*41)%360)]
\1a&HFF&\3a&HFF&\4a&HFF&\blura&HFF&)

【SI-VIS（中央に集まりながらフェードアウト）】
行：{\org([%W/2],[%y100])\fsp0\fs[%f-5]\alpha&HFF&}{\fr0.05\t([%n3]0,[%n*3+20]1,\fr0\fsp0\fs[%f]\alpha&H00&)}
文字：{\move(%x,%y,%x,[%y-80],[%e+200],[%e+800],23)}{\t([%e+200],[%e+800],28,\alpha&HFF&\fscx70\fs30\3a&HFF&\4a&HFF&\glow0)} 

【フリーレン（１文字ずつフェードアウト）】
行：{\org([%W/2],[%y100])\fsp0\fs[%f-5]\alpha&HFF&}{\fr0.05\t([%n3]0,[%n3+20]1,\fr0\fsp0\fs[%f]\alpha&H00&)}
文字：{\org([%x+%w/2],[%y])\alpha&HFF&\fscx100\fscy100
\t([%n3]0,[%n3+60]0,0.5,\alpha&H00&)
\t([%n3]0,[%n3+50]0,23,\fscx110\fscy110)
\t([%n3+50]0,[%n3+100]0,23,\fscx100\fscy100)
}
{\t([%e+200],[%e+1000],1,\bord3\fs10\1a&Hff&\3a&Hff&\4a&Hff&)}

【クラニカ（ふわっと文字が大きくなって登場、ふわっと↓方向にアウト）】
文字：{\org([%x+%w/2],[%y])\frz4\fscx90\fscy90\alpha&HFF&\move(%x,[%y+20],%x,%y,[%n24],[%n24+500],7)\t([%n24],[%n24+500],7,\frz0\fscx100\fscy100\alpha&H00&)} {\t([%V+320+%n18],[%V+440+%n18],7,\alpha&H18&\fsp-3)} {\move(%x,%y,%x,[%y-16],[%V+200+%n18],[%V+1300+%n18])} {\t([%V+440+%n18],[%V+960+%n18],7,\fscx94\fscy94\frz4\alpha&HFF&\3a&HFF&\4a&HFF&\blura&HFF&\fs40)}
{\t([%V],[%V+300],7,\1a&H30&\3a&H60&\blura&H100&)}"""
ass_tag_default_json = json.dumps(ASS_TAG_DEFAULT_TEXT, ensure_ascii=False)

def load_df_from_github(filename, **kwargs):
    url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{filename}"
    try:
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                try:
                    df = pd.read_csv(io.BytesIO(response.content), encoding=enc, engine='python', **kwargs)
                    if not df.empty:
                        df.columns = df.columns.astype(str).str.replace('\ufeff', '').str.strip()
                    print(f"[GitHub] OK {filename} rows={len(df)}")
                    return df
                except Exception:
                    continue
        return pd.DataFrame()
    except Exception as e:
        print(f"[GitHub] err {e}")
        return pd.DataFrame()


def load_df_from_gas_with_status(filename, **kwargs):
    try:
        response = requests.get(GAS_WEB_APP_URL, params={'filename': filename}, timeout=60)
    except Exception:
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
            return df, "ok"
        except Exception:
            continue
    return pd.DataFrame(), "error"


def load_df_from_gas(filename, **kwargs):
    df, _ = load_df_from_gas_with_status(filename, **kwargs)
    return df


def load_json_from_gas(filename):
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
    except Exception:
        return {}


def save_df_to_gas(filename, df):
    try:
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        payload = {'filename': filename, 'content': csv_buffer.getvalue()}
        response = requests.post(GAS_WEB_APP_URL, json=payload, timeout=60)
        return response.status_code == 200
    except Exception:
        return False


# ==========================================
# ローカル(リポジトリ)を履歴CSVの正本(プライマリ)として扱うためのヘルパー
# ------------------------------------------
# GAS のみに依存していた構成では、GAS が一時的に not_found / empty / error を
# 返すと履歴データが丸ごと消失する事故が発生した。
# リポジトリにコミットされたCSVを正本にすることで、GAS の状態に左右されずに
# 履歴を恒久保持できるようにする。
# ==========================================
def load_df_from_local(filename, **kwargs):
    if not filename or not os.path.exists(filename):
        return pd.DataFrame()
    try:
        with open(filename, 'rb') as f:
            content_bytes = f.read()
        if not content_bytes or content_bytes.lstrip(CSV_EMPTY_PREFIX_BYTES) == b'':
            return pd.DataFrame()
        for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
            try:
                df = pd.read_csv(io.BytesIO(content_bytes), encoding=enc, engine='python', **kwargs)
                if len(df.columns) > 0:
                    df.columns = df.columns.astype(str).str.replace('\ufeff', '').str.strip()
                print(f"[Local] OK {filename} rows={len(df)}")
                return df
            except Exception:
                continue
        return pd.DataFrame()
    except Exception as e:
        print(f"[Local] err {filename}: {e}")
        return pd.DataFrame()


def save_df_to_local(filename, df):
    try:
        if df is None or df.empty:
            return False
        df.to_csv(filename, index=False, encoding='utf-8-sig')
        return True
    except Exception as e:
        print(f"[Local Save] err {filename}: {e}")
        return False


now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
current_date_str = now.strftime("%Y/%m/%d")
current_datetime_str = now.strftime("%Y/%m/%d %H:%M")

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
    11105: "ヤマテル部屋", 11106: "冨塚部屋", 11122: "みにゅ部屋", 11107: "ブルーベリー部屋",
    11110: "加古部屋", 11111: "ヒロ部屋", 11056: "愁部屋"
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
    return source_series.str.contains(safe_target, case=False, na=False)


HISTORY_MAX_ROWS = 9500
HISTORY_ARCHIVE_MISS_LIMIT = 3
ROOM_FETCH_TIMEOUT = 6
ROOM_FETCH_WORKERS = 16
HISTORY_DEDUP_COLS = ['取得日', '部屋主', '順番', '曲名（ファイル名）', '歌った人']


def get_history_filename_candidates(num):
    if num == 1:
        # 既存環境差分吸収:
        # - history.csv
        # - history_1.csv
        # - history1.csv
        # のどちらでも読めるようにする
        return ["history.csv", "history_1.csv", "history1.csv"]
    return [f"history_{num}.csv", f"history{num}.csv"]


def get_history_filename(num, filename_by_num=None):
    if filename_by_num and num in filename_by_num:
        return filename_by_num[num]
    return get_history_filename_candidates(num)[0]


def load_history_df_with_fallback(filename):
    """
    履歴CSVは「ローカル(リポジトリ) → GitHub Raw → GAS」の順にフォールバックして読み込む。
    リポジトリにコミットされたファイルを正本扱いすることで、GAS が一時的に
    not_found / empty / error を返しても履歴データが消失しないようにする。
    """
    local_df = load_df_from_local(filename)
    if not local_df.empty:
        return local_df, "ok"

    gh_df = load_df_from_github(filename)
    if not gh_df.empty:
        return gh_df, "ok"

    df, st = load_df_from_gas_with_status(filename)
    return df, st


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

    def _fmt(v):
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
    df['順番'] = df['順番'].apply(_fmt)
    return df


def cleanup_history_df(df):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.fillna("")
    df = df.copy().fillna("")
    for col in ['部屋主', '曲名（ファイル名）', '作品名', '歌手名']:
        if col in df.columns:
            df = df[df[col].astype(str) != col]
    bad_cols = []
    for c in df.columns:
        col = str(c).strip()
        for pat in [r'^Error:', r'Exception:\s*Service error:\s*Drive']:
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


def save_df_to_gas_checked(filename, df, min_existing_rows=0):
    if df is None:
        return False
    df = cleanup_history_df(df)
    if df.empty or len(df) < min_existing_rows:
        return False
    if not save_df_to_gas(filename, df):
        return False
    verify_df, st = load_df_from_gas_with_status(filename)
    if st != "ok":
        return False
    if len(cleanup_history_df(verify_df)) < len(df):
        return False
    return True


def save_history_persistent(filename, df, min_existing_rows=0):
    """
    履歴CSVを「ローカル(リポジトリ)を正本」として保存する。

    1) ローカルファイルへ必ず保存 (リポジトリにコミットされ恒久保存される)
    2) GAS への保存は best-effort (失敗しても履歴自体は失われない)

    既存行数 (min_existing_rows) より少ない行数を書き出すことは
    「履歴の縮退」を意味するため、保存をブロックする。
    """
    if df is None:
        return False
    df = cleanup_history_df(df)
    if df.empty or len(df) < min_existing_rows:
        print(f"[STOP] {filename} would shrink: new={len(df)} prev={min_existing_rows}")
        return False

    if not save_df_to_local(filename, df):
        return False

    # GAS への保存はベストエフォート (バックアップ用途)。
    # 失敗してもローカル(リポジトリ)に保存済みなので履歴は保全される。
    try:
        if save_df_to_gas(filename, df):
            print(f"[GAS Save] OK {filename}")
        else:
            print(f"[GAS Save] WARN {filename}: GAS save failed (local primary OK)")
    except Exception as e:
        print(f"[GAS Save] WARN {filename}: {e} (local primary OK)")

    return True


def load_all_history_files():
    histories = []
    loaded = []
    filename_by_num = {}
    miss = 0
    num = 1
    while miss < HISTORY_ARCHIVE_MISS_LIMIT:
        found = False
        states = []
        for fn in get_history_filename_candidates(num):
            df, st = load_history_df_with_fallback(fn)
            states.append(st)
            if st == "ok":
                df = cleanup_history_df(df)
                histories.append({"num": num, "filename": fn, "df": df})
                loaded.append(fn)
                filename_by_num[num] = fn
                miss = 0
                found = True
                break
        if found:
            num += 1
            continue
        if any(st == "error" for st in states):
            return histories, loaded, filename_by_num, False
        miss += 1
        num += 1
    return histories, loaded, filename_by_num, True


def fetch_room_df(port):
    url = f"http://ykr.moe:{port}/simplelist.php"
    response = requests.get(url, timeout=ROOM_FETCH_TIMEOUT)
    response.raise_for_status()
    dfs = pd.read_html(io.BytesIO(response.content))
    if not dfs:
        raise ValueError("テーブルなし")
    df = dfs[0].fillna("")
    if df.empty or '順番' not in df.columns or '曲名（ファイル名）' not in df.columns:
        raise ValueError("カラム不足")
    df = df.replace(r'\s*詳細を見る ▼', '', regex=True)
    df['部屋主'] = room_map[port]
    df['取得日'] = current_date_str
    return df


# --- 既存履歴読み込み ---
history_records, loaded_history_files, history_filename_by_num, history_load_ok = load_all_history_files()
print(f"履歴ファイル: {loaded_history_files}")
history_dfs = [h['df'] for h in history_records]
full_history_before_update = cleanup_history_df(pd.concat(history_dfs, ignore_index=True)) if history_dfs else pd.DataFrame()

# --- 新データ取得 ---
print("データ取得中...")
new_data_frames = []
fetched_ports = []
target_ports = list(room_map.keys())
max_workers = min(ROOM_FETCH_WORKERS, max(1, len(target_ports)))
with ThreadPoolExecutor(max_workers=max_workers) as ex:
    futures = {ex.submit(fetch_room_df, p): p for p in target_ports}
    for f in as_completed(futures):
        port = futures[f]
        try:
            df = f.result()
            new_data_frames.append(df)
            fetched_ports.append(port)
        except Exception as e:
            print(f"[Fetch SKIP] {port}: {e}")

print(f"成功: {len(fetched_ports)}/{len(target_ports)}")

if not new_data_frames:
    full_df = full_history_before_update
else:
    new_df = cleanup_history_df(pd.concat(new_data_frames, ignore_index=True))
    dedup_cols = [c for c in HISTORY_DEDUP_COLS if c in new_df.columns]
    if dedup_cols:
        new_df = new_df.drop_duplicates(subset=dedup_cols, keep='last')
        new_df = cleanup_history_df(new_df)

    if not history_load_ok:
        full_df = cleanup_history_df(pd.concat([full_history_before_update, new_df], ignore_index=True)) if not full_history_before_update.empty else new_df
    else:
        existing_keys = set(make_dedup_key(full_history_before_update, for_history_compare=True).tolist()) if not full_history_before_update.empty else set()
        new_keys = make_dedup_key(new_df, for_history_compare=True)
        new_unique_df = new_df[~new_keys.isin(existing_keys)].copy()
        new_unique_df = cleanup_history_df(new_unique_df)

        if new_unique_df.empty:
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
                fn = get_history_filename(active_num, filename_by_num=history_filename_by_num)
                cur = cleanup_history_df(active_df)
                cap = HISTORY_MAX_ROWS - len(cur)
                if cap <= 0:
                    active_num += 1
                    active_df = pd.DataFrame()
                    continue
                part = remaining_df.iloc[:cap].copy()
                nxt = cleanup_history_df(pd.concat([cur, part], ignore_index=True))
                if save_history_persistent(fn, nxt, min_existing_rows=len(cur)):
                    saved_parts[active_num] = nxt
                    remaining_df = remaining_df.iloc[cap:].reset_index(drop=True)
                    if not remaining_df.empty:
                        active_num += 1
                        active_df = pd.DataFrame()
                    else:
                        active_df = nxt
                else:
                    print(f"[STOP] {fn}")
                    break

            merged = {h['num']: h['df'] for h in history_records}
            merged.update(saved_parts)
            full_df = cleanup_history_df(pd.concat([merged[n] for n in sorted(merged)], ignore_index=True)) if merged else pd.DataFrame()

if full_df is None or full_df.empty:
    full_df = pd.DataFrame()
else:
    full_df = cleanup_history_df(full_df)

print(f"全履歴: {len(full_df)} 行")

# --- オフライン読込 ---
offline_targets = []
for fn in OFFLINE_FILES:
    odf = load_df_from_github(fn)
    if not odf.empty and '曲名' in odf.columns:
        offline_targets.extend([normalize_offline_text(str(x)) for x in odf['曲名'].fillna("").tolist()])

# --- 集計用ソース ---
analysis_source_df = full_df.copy()
if not analysis_source_df.empty:
    analysis_source_df['dt_obj'] = pd.to_datetime(analysis_source_df['取得日'], errors='coerce')
    analysis_source_df = analysis_source_df.dropna(subset=['dt_obj'])
    analysis_source_df['norm_filename'] = analysis_source_df['曲名（ファイル名）'].apply(normalize_text)

    def _resc(row):
        rw = str(row['作品名']) if pd.notna(row['作品名']) else ""
        rs = str(row['曲名（ファイル名）']) if pd.notna(row['曲名（ファイル名）']) else ""
        if rw.strip() in ["-", "−", "", "nan"]:
            m = re.search(r'【(.*?)】', rs)
            if m:
                return normalize_text(m.group(1))
        return normalize_text(rw)

    analysis_source_df['norm_workname'] = analysis_source_df.apply(_resc, axis=1) if '作品名' in analysis_source_df.columns else ""
    excl = ['test', 'テスト', 'システム', 'admin', 'System']
    full_history = analysis_source_df[
        ~analysis_source_df['歌った人'].astype(str).apply(lambda x: any(k in x for k in excl))
    ].sort_values('dt_obj').reset_index(drop=True)
else:
    full_history = pd.DataFrame()


raw_df = load_df_from_gas("cool_analysis.csv", header=None)
categorized_data = {}
if not raw_df.empty:
    raw_df = raw_df.fillna("").drop_duplicates(keep='last')
    cur_cat = None
    for idx, row in raw_df.iterrows():
        if not any(str(x).strip() for x in row):
            continue
        col0 = str(row[0]).strip()
        is_cat = any(c in col0 for c in ALLOWED_CATEGORIES) and "作品名" not in col0
        if is_cat:
            cur_cat = col0
            categorized_data.setdefault(cur_cat, [])
            continue
        if "作品名" in col0 or cur_cat is None:
            continue
        anime = str(row[0]).strip() if len(row) > 0 else ""
        type_ = str(row[1]).strip() if len(row) > 1 else ""
        artist = str(row[2]).strip() if len(row) > 2 else ""
        song = str(row[3]).strip() if len(row) > 3 else ""
        if not anime and not song:
            continue
        categorized_data[cur_cat].append({"anime": anime, "type": type_, "artist": artist, "song": song})


def is_song_created(item):
    sn = normalize_text(item["song"])
    sr = normalize_offline_text(item["song"])
    wn = normalize_text(item["anime"])
    cnt = 0
    if sn:
        for s in offline_targets:
            if (sn in s) or (sr in s):
                if wn:
                    if wn in s:
                        cnt += 1
                else:
                    cnt += 1
    return cnt


def compute_match_for_item(item, hdf):
    if hdf.empty:
        return []
    sn = normalize_text(item["song"])
    wn = normalize_text(item["anime"])
    sm = check_match(sn, hdf['norm_filename'])
    am = (
        hdf['norm_filename'].str.contains(re.escape(wn), case=False, na=False) |
        hdf['norm_workname'].str.contains(re.escape(wn), case=False, na=False)
    ) if wn else pd.Series([False] * len(hdf))
    if sn and wn:
        mask = sm & am
    elif sn:
        mask = sm
    elif wn:
        mask = am
    else:
        return []
    return hdf.index[mask].tolist()


def assign_matches_with_artist_disambiguation(items, hdf):
    """
    items のそれぞれを hdf にマッチさせ、同一 (作品, 曲名) で複数ヒットする
    バージョン違いを歌手名で振り分けた、item ごとの履歴 index リストを返す。
    """
    if hdf is None or hdf.empty or not items:
        return [[] for _ in items]

    raw = [compute_match_for_item(it, hdf) for it in items]

    # group items by normalized (anime, song)
    buckets = {}
    for i, it in enumerate(items):
        key = (normalize_text(it.get("anime", "")), normalize_text(it.get("song", "")))
        buckets.setdefault(key, []).append(i)

    assigned = [list(idx) for idx in raw]
    if '歌手名' not in hdf.columns:
        return assigned

    history_artists_norm = hdf['歌手名'].astype(str).apply(normalize_text)
    for member_indices in buckets.values():
        if len(member_indices) <= 1:
            continue
        union = set()
        for mi in member_indices:
            union.update(raw[mi])
        if not union:
            continue
        # reset and reassign each history row to exactly one item by 歌手名
        for mi in member_indices:
            assigned[mi] = []
        member_artists = [
            (mi, normalize_text(items[mi].get("artist", "")))
            for mi in member_indices
        ]
        for hidx in union:
            h_art = history_artists_norm.loc[hidx] if hidx in history_artists_norm.index else ""
            best_mi = None
            best_score = -1
            for mi, an in member_artists:
                if not an:
                    continue
                if h_art and (an in h_art or h_art in an):
                    score = len(an)  # longer matched artist wins
                    if score > best_score:
                        best_score = score
                        best_mi = mi
            if best_mi is None:
                # fallback: keep the row on the first item to avoid losing it
                best_mi = member_indices[0]
            assigned[best_mi].append(hidx)
    return assigned


COOL_START = pd.to_datetime("2026/01/01")
COOL_END = pd.to_datetime("2026/09/30")
target_history = full_history[
    (full_history['dt_obj'] >= COOL_START) & (full_history['dt_obj'] <= COOL_END)
] if not full_history.empty else pd.DataFrame()

cool_data_for_js = {}
for cat in ALLOWED_CATEGORIES:
    items = categorized_data.get(cat, [])
    if not items:
        cool_data_for_js[cat] = {"works": [], "max_count": 0, "max_user": 0}
        continue

    # 「上伊那ぼたん、酔へる姿は百合の花」「感情グラス」のように
    # 同じ曲名で複数のボーカルバージョンがある場合に二重計上しないため、
    # 履歴行を歌手名と最も一致する 1 件にだけ割り当てる。
    assigned_idx_per_item = assign_matches_with_artist_disambiguation(items, target_history)

    enriched = []
    for i, it in enumerate(items):
        idx = assigned_idx_per_item[i]
        if idx:
            mt = target_history.loc[idx]
            cnt = len(mt)
            uc = mt['歌った人'].nunique()
        else:
            cnt = 0
            uc = 0
        enriched.append({**it, "count": cnt, "user_count": uc, "creation_count": is_song_created(it)})
    enriched.sort(key=lambda x: x['anime'])
    works = []
    for anime_name, gi in groupby(enriched, key=lambda x: x['anime']):
        g = list(gi)
        works.append({
            "anime": anime_name,
            "songs": g,
            "op_n": sum(1 for x in g if 'OP' in x['type'].upper()),
            "ed_n": sum(1 for x in g if 'ED' in x['type'].upper()),
            "in_n": sum(1 for x in g if 'IN' in x['type'].upper()),
            # 作成数: offline_list を参照して算出した creation_count を、
            # 種別ごと（OP/ED/IN）と作品全体で合計する
            "op_created": sum(min(1, x['creation_count']) for x in g if 'OP' in x['type'].upper()),
            "ed_created": sum(min(1, x['creation_count']) for x in g if 'ED' in x['type'].upper()),
            "in_created": sum(min(1, x['creation_count']) for x in g if 'IN' in x['type'].upper()),
            "total_created": sum(min(1, x['creation_count']) for x in g),
            "total_count": sum(x['count'] for x in g),
            "total_user": sum(x['user_count'] for x in g)
        })
    cool_data_for_js[cat] = {
        "works": works,
        "max_count": max([w['total_count'] for w in works], default=0),
        "max_user": max([w['total_user'] for w in works], default=0)
    }

ranking_data_by_cat = {}
for cat in ALLOWED_CATEGORIES:
    flat = []
    for w in cool_data_for_js.get(cat, {"works": []})['works']:
        for s in w['songs']:
            if s['count'] > 0:
                flat.append({
                    "anime": w['anime'], "song": s['song'], "artist": s['artist'],
                    "type": s['type'], "count": s['count'], "user_count": s['user_count']
                })
    ranking_data_by_cat[cat] = flat


trend_data_for_js = {}
target_works = cool_data_for_js.get(TREND_TARGET_CATEGORY, {"works": []})['works']
trend_items = []
for w in target_works:
    for s in w['songs']:
        trend_items.append({"anime": w['anime'], "song": s['song'], "artist": s['artist'], "type": s['type']})

now_dt = pd.Timestamp(now.replace(tzinfo=None).date()) + pd.Timedelta(days=1)

for pd_days in TREND_PERIOD_OPTIONS:
    cs = now_dt - pd.Timedelta(days=pd_days)
    ce = now_dt
    ps = cs - pd.Timedelta(days=pd_days)
    pe = cs
    if not full_history.empty:
        cur_h = full_history[(full_history['dt_obj'] >= cs) & (full_history['dt_obj'] < ce)]
        prv_h = full_history[(full_history['dt_obj'] >= ps) & (full_history['dt_obj'] < pe)]
        all_h = full_history[(full_history['dt_obj'] >= COOL_START) & (full_history['dt_obj'] <= COOL_END)]
    else:
        cur_h = pd.DataFrame()
        prv_h = pd.DataFrame()
        all_h = pd.DataFrame()

    stats = []
    # 同一作品+曲名で複数バージョンがある場合は歌手名で振り分けて二重計上を防止
    cur_assigned = assign_matches_with_artist_disambiguation(trend_items, cur_h)
    prv_assigned = assign_matches_with_artist_disambiguation(trend_items, prv_h)
    all_assigned = assign_matches_with_artist_disambiguation(trend_items, all_h)
    for i, it in enumerate(trend_items):
        ci = cur_assigned[i]
        pi = prv_assigned[i]
        ai = all_assigned[i]
        cc = len(ci)
        pc = len(pi)
        cu = cur_h.loc[ci]['歌った人'].nunique() if ci else 0
        pu = prv_h.loc[pi]['歌った人'].nunique() if pi else 0
        ac = len(ai)
        au = all_h.loc[ai]['歌った人'].nunique() if ai else 0
        stats.append({
            **it, "cur_count": cc, "cur_user": cu,
            "prev_count": pc, "prev_user": pu, "delta": cc - pc,
            "is_new": (cc > 0 and pc == 0),
            "all_count": ac, "all_user": au
        })

    surge_count = sum(1 for x in stats if x['delta'] > 0)
    new_in = sum(1 for x in stats if x['is_new'])
    max_delta = max([x['delta'] for x in stats], default=0)
    sorted_items = sorted([x for x in stats if x['delta'] != 0 or x['cur_count'] > 0],
                          key=lambda x: (-x['delta'], -x['cur_count']))
    trend_data_for_js[str(pd_days)] = {
        "kpi": {"surge_count": surge_count, "new_in": new_in, "max_delta": max_delta},
        "items": sorted_items[:TREND_PICKUP_TOTAL]
    }


image_map = load_json_from_gas("image_map.json")
if not isinstance(image_map, dict):
    image_map = {}


# --- ★修正: history_for_js は列名を直接参照せず .get() で安全に構築 ---
def _rescue_workname(wk_raw, sg_raw):
    wks = str(wk_raw).strip() if wk_raw is not None else ""
    if wks in ['-', '−', '', 'nan']:
        m = re.search(r'【(.*?)】', str(sg_raw))
        if m:
            return m.group(1)
    return wk_raw

history_for_js = []
if not full_df.empty:
    try:
        records = full_df.fillna("").to_dict('records')
    except Exception as e:
        print(f"[Warn] to_dict失敗: {e}")
        records = []

    col_aliases = {
        'd':  ['取得日'],
        'rm': ['部屋主'],
        'o':  ['順番'],
        'sg': ['曲名（ファイル名）', '曲名(ファイ   名)', '曲名'],
        'wk': ['作品名'],
        'ar': ['歌手名', '歌手'],
        'u':  ['歌った人', '歌唱者'],
    }

    def pick(rec, keys):
        for k in keys:
            if k in rec:
                v = rec[k]
                if v is None:
                    return ""
                return str(v)
        return ""

    for r in records:
        d  = pick(r, col_aliases['d'])
        rm = pick(r, col_aliases['rm'])
        o  = pick(r, col_aliases['o'])
        sg = pick(r, col_aliases['sg'])
        wk = pick(r, col_aliases['wk'])
        ar = pick(r, col_aliases['ar'])
        u  = pick(r, col_aliases['u'])

        # 取得日を必ず "YYYY/MM/DD" (ゼロ埋め) に正規化する。
        # 文字列ソートで日付降順を担保するため (例: "2026/4/29" と "2026/04/29" の混在を防ぐ)。
        if d:
            _dt = pd.to_datetime(d, errors='coerce')
            if pd.notna(_dt):
                d = _dt.strftime('%Y/%m/%d')

        sn = normalize_text(sg)
        wn_src = _rescue_workname(wk, sg)
        wn = normalize_text(wn_src)

        history_for_js.append({
            "d": d, "rm": rm, "o": o,
            "sg": sg, "wk": wk, "ar": ar, "u": u,
            "sn": sn, "wn": wn
        })

print(f"history_for_js件数: {len(history_for_js)}")

unique_rooms = sorted(list(set(room_map.values())))

cool_json = json.dumps(cool_data_for_js, ensure_ascii=False)
ranking_json = json.dumps(ranking_data_by_cat, ensure_ascii=False)
trend_json = json.dumps(trend_data_for_js, ensure_ascii=False, default=str)
history_json = json.dumps(history_for_js, ensure_ascii=False)
image_map_json = json.dumps(image_map, ensure_ascii=False)
rooms_json = json.dumps(unique_rooms, ensure_ascii=False)
categories_json = json.dumps(ALLOWED_CATEGORIES, ensure_ascii=False)
trend_periods_json = json.dumps(TREND_PERIOD_OPTIONS, ensure_ascii=False)
trend_target_json = json.dumps(TREND_TARGET_CATEGORY, ensure_ascii=False)


html_content = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Karaoke Dashboard</title>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Noto+Sans+JP:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --primary:#1f2937; --accent:#6366f1; --accent-2:#7c3aed; --accent-soft:#a5b4fc; --accent-bg:#eef2ff;
  --bg:#f7f7fb; --panel:#fff; --text:#1f2937; --text-sub:#6b7280; --text-mute:#9ca3af;
  --border:#e8e8ef; --border-soft:#eef0f3;
  --green:#10b981; --green-bg:#ecfdf5; --green-bd:#a7f3d0;
  --orange:#f97316; --orange-bg:#fff7ed; --orange-bd:#fed7aa;
  --red:#ef4444; --amber:#f59e0b;
  --gold:#f59e0b; --silver:#9ca3af; --bronze:#d97706;
  --radius:10px; --radius-lg:14px;
  --maxw:600px;
}
@media (min-width:768px){
  :root{ --maxw:880px; }
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
html,body{margin:0;padding:0;background:var(--bg);color:var(--text);
  font-family:"Inter","Noto Sans JP","Helvetica Neue",Arial,sans-serif;
  font-size:15px;line-height:1.5;
  -webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;
  text-rendering:optimizeLegibility;
  font-feature-settings:"palt" 1;
}
body{padding-bottom:24px}
img{image-rendering:auto;-ms-interpolation-mode:bicubic}
button{font-family:inherit;font-size:inherit}
input,select{font-family:inherit;font-size:inherit}
a{color:inherit;text-decoration:none}

/* Header */
.app-header{max-width:var(--maxw);margin:0 auto;padding:18px 16px 0 16px;display:flex;justify-content:space-between;align-items:center}
.brand{font-size:24px;font-weight:800;color:var(--primary);letter-spacing:-0.01em}
.theme-toggle{
  width:32px;height:32px;border-radius:8px;background:#fff;border:1px solid var(--border);
  color:var(--text-sub);cursor:pointer;display:inline-flex;align-items:center;justify-content:center;
  font-size:13px;flex-shrink:0;
}
.theme-toggle:hover{color:var(--accent);border-color:var(--accent-soft)}

/* Dark mode */
body.dark{
  --primary:#f1f5f9; --accent:#818cf8; --accent-2:#a78bfa; --accent-soft:#4338ca; --accent-bg:#1e1b4b;
  --bg:#0f172a; --panel:#1e293b; --text:#f1f5f9; --text-sub:#cbd5e1; --text-mute:#94a3b8;
  --border:#334155; --border-soft:#1e293b;
  --green-bg:#064e3b; --green-bd:#047857;
  --orange-bg:#7c2d12; --orange-bd:#9a3412;
}
body.dark .top-card,body.dark .card,body.dark .icon-btn,body.dark .pill-select,body.dark .pill-text,
body.dark .search-pill,body.dark .rank-row-flat,body.dark .notable-row,body.dark .trend-stat,
body.dark .modal,body.dark .modal-row,body.dark .env-csv-text,body.dark .env-work-row,
body.dark .env-lock-card,body.dark .popover,body.dark .pager button,body.dark .theme-toggle,
body.dark .room-chip,body.dark .env-lock-input,body.dark .ass-card,body.dark .ass-input,
body.dark .ass-text,body.dark .ass-toolbar-note,body.dark .ass-raw-panel{
  background:var(--panel);color:var(--text);border-color:var(--border);
}
body.dark .top-card input,body.dark .top-card select,body.dark .search-pill input,
body.dark .env-csv-text,body.dark .env-lock-input,body.dark .room-search,
body.dark .ass-input,body.dark .ass-text{
  color:var(--text);background:transparent;
}
body.dark .room-search{border-color:var(--border)}
body.dark .num-badge{background:#312e81;color:#c7d2fe}
body.dark .modal-summary{background:#1e1b4b;color:#a5b4fc}
body.dark .detail-row{background:var(--panel);border-color:var(--border-soft)}
body.dark .card-detail{background:#0b1220}
body.dark .song-row{border-color:var(--border-soft)}
body.dark .growth-badge{background:var(--panel);border-color:var(--border)}
body.dark .growth-badge.flat{background:#374151;color:var(--text-sub)}
body.dark .modal-close{background:#374151;color:var(--text-sub)}
body.dark .confirm-btn{background:#374151;color:var(--text-sub);border-color:var(--border)}
body.dark .env-status{background:#1e1b4b;color:#a5b4fc}
body.dark .btn-clear{background:#374151;color:var(--text-sub)}
body.dark .dl-btn.ghost{background:var(--panel);color:var(--accent)}
body.dark .type-pill.tp-none,body.dark .type-chip.tp-none{background:#374151;color:var(--text-sub);border-color:var(--border)}

/* Dark mode: 急上昇ピックアップ */
body.dark .trend-pickup{
  background:linear-gradient(135deg,#3f1d1d 0%,#2a1414 60%,#4a1f1f 100%);
  border-color:#7f1d1d;
}
body.dark .trend-pickup-head{color:#fecaca}
body.dark .trend-pickup-head i{color:#f87171}
body.dark .trend-pickup-row{background:var(--panel);border-color:#7f1d1d;color:var(--text)}
body.dark .trend-pickup-row .pickup-anime{color:var(--text)}
body.dark .trend-pickup-row .pickup-sub{color:var(--text-sub)}
body.dark .trend-pickup-row .num-badge{background:#7f1d1d;color:#fecaca}

/* Top settings */
.top-cards{
  max-width:var(--maxw);margin:14px auto 0;padding:0 16px;
  display:grid;grid-template-columns:1fr 1fr;gap:8px;
}
.top-card{
  background:#fff;border:1px solid var(--border);border-radius:10px;padding:7px 12px;
  display:flex;align-items:center;gap:9px;
}
.top-card .ico{
  width:28px;height:28px;border-radius:7px;background:var(--accent-bg);color:var(--accent);
  display:flex;align-items:center;justify-content:center;font-size:12px;flex-shrink:0;
}
.top-card .lbl{font-size:10.5px;color:var(--text-sub);margin-bottom:1px;font-weight:500;line-height:1.2}
.top-card input,.top-card select{
  border:none;background:transparent;font-weight:700;font-size:14px;color:var(--primary);
  outline:none;width:100%;cursor:pointer;padding:0;line-height:1.2;
}
.top-card input[type=number]{cursor:text}

/* Saved HTML export-link inherits surrounding text styling */
.export-link{color:inherit;text-decoration:none}
.export-link:hover{text-decoration:underline}

/* Tabs */
.tabs-wrap{max-width:var(--maxw);margin:16px auto 0;padding:0 16px}
.tabs{display:flex;gap:0;border-bottom:1px solid var(--border);overflow-x:auto;scrollbar-width:none}
.tabs::-webkit-scrollbar{display:none}
.tab-btn{
  padding:11px 12px;border:none;background:none;color:var(--text-sub);
  font-weight:600;font-size:13.5px;white-space:nowrap;cursor:pointer;
  border-bottom:2px solid transparent;display:inline-flex;align-items:center;gap:5px;
  transition:color .15s, border-color .15s;
}
.tab-btn:hover{color:var(--primary)}
.tab-btn.active{color:var(--accent);border-bottom-color:var(--accent)}

/* Toolbar */
.tab-toolbar{max-width:var(--maxw);margin:14px auto 0;padding:0 16px;display:flex;flex-direction:column;gap:8px}
.toolbar-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}

.search-pill{
  flex:1;min-width:0;background:#fff;border:1px solid var(--border);border-radius:12px;padding:11px 16px;
  display:flex;align-items:center;gap:10px;
}
.search-pill input{flex:1;border:none;outline:none;font-size:14px;background:transparent;min-width:0}
.search-pill i{color:var(--text-mute);font-size:14px}

.icon-btn{
  width:44px;height:44px;border-radius:12px;background:#fff;border:1px solid var(--border);
  display:flex;align-items:center;justify-content:center;color:var(--text-sub);cursor:pointer;flex-shrink:0;
  font-size:14px;
}
.icon-btn:hover{color:var(--accent);border-color:var(--accent-soft)}
.icon-btn.active{background:var(--accent);color:#fff;border-color:var(--accent)}

.pill-select{
  background:#fff;border:1px solid var(--border);border-radius:10px;padding:8px 10px;font-size:13px;
  display:inline-flex;align-items:center;gap:6px;cursor:pointer;font-weight:600;color:var(--primary);
  appearance:none;-webkit-appearance:none;
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%236b7280' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'><polyline points='6 9 12 15 18 9'/></svg>");
  background-repeat:no-repeat;background-position:right 8px center;padding-right:26px;
  min-width:0;flex:1 1 auto;max-width:100%;
}
.pill-text{
  background:#fff;border:1px solid var(--border);border-radius:10px;padding:8px 10px;font-size:12.5px;
  font-weight:600;color:var(--primary);display:inline-flex;align-items:center;gap:6px;
}

.dl-row{display:flex;gap:6px;align-items:center;flex-wrap:wrap;justify-content:flex-end}
.dl-btn{
  background:linear-gradient(180deg,var(--accent),var(--accent-2));color:#fff;border:none;border-radius:8px;
  padding:7px 10px;font-size:11.5px;font-weight:600;cursor:pointer;display:inline-flex;align-items:center;gap:5px;
  white-space:nowrap;flex-shrink:0;
}
.dl-btn.ghost{background:#fff;color:var(--accent);border:1px solid var(--accent-soft)}
.dl-btn:hover{filter:brightness(1.05)}
.dl-btn.compact{padding:6px 8px;font-size:11px}
.dl-btn .dl-lbl{display:inline}
@media (max-width:520px){
  .dl-btn .dl-lbl{display:none}
  .dl-btn{padding:7px 9px}
}

.update-line{
  max-width:var(--maxw);margin:10px auto 0;padding:0 16px;
  display:flex;justify-content:flex-end;font-size:11.5px;color:var(--text-mute);
}
.count-line{
  max-width:var(--maxw);margin:14px auto 0;padding:0 16px;
  font-size:14px;color:var(--text-sub);font-weight:600;display:flex;align-items:center;gap:6px;
}
.count-line i{color:var(--accent)}

.tab-content{display:none;max-width:var(--maxw);margin:0 auto;padding:6px 16px 80px}
.tab-content.active{display:block}

/* ===== Type pills ===== */
.type-pill{
  display:inline-flex;align-items:center;gap:4px;padding:3px 10px;font-size:11px;font-weight:700;
  border-radius:6px;letter-spacing:.02em;
}
.type-pill b{font-weight:800}
.type-pill.tp-op{background:var(--accent-bg);border:1px solid #e0e7ff;color:var(--accent)}
.type-pill.tp-ed{background:var(--orange-bg);border:1px solid var(--orange-bd);color:var(--orange)}
.type-pill.tp-in{background:var(--green-bg);border:1px solid var(--green-bd);color:var(--green)}
.type-pill.tp-none{background:#f3f4f6;border:1px solid var(--border);color:var(--text-sub)}
.type-pill.tp-created{background:#fef2f2;border:1px solid #fecaca;color:#dc2626}
.type-pill.tp-created b{color:#dc2626}
body.dark .type-pill.tp-created{background:#3f1d1d;border-color:#7f1d1d;color:#fca5a5}
body.dark .type-pill.tp-created b{color:#fca5a5}
.song-created-mark{display:inline-flex;align-items:center;gap:3px;margin-left:8px;
  padding:1px 6px;font-size:10.5px;font-weight:700;
  background:#fef2f2;border:1px solid #fecaca;color:#dc2626;border-radius:4px;letter-spacing:.02em}
.song-created-mark i{font-size:9.5px}
body.dark .song-created-mark{background:#3f1d1d;border-color:#7f1d1d;color:#fca5a5}

/* Type chip (square, for cool song-row) */
.type-chip{
  width:38px;height:38px;border-radius:8px;display:flex;align-items:center;justify-content:center;
  font-size:11px;font-weight:800;letter-spacing:.04em;flex-shrink:0;
}
.type-chip.tp-op{background:var(--accent-bg);color:var(--accent);border:1px solid #e0e7ff}
.type-chip.tp-ed{background:var(--orange-bg);color:var(--orange);border:1px solid var(--orange-bd)}
.type-chip.tp-in{background:var(--green-bg);color:var(--green);border:1px solid var(--green-bd)}
.type-chip.tp-none{background:#f3f4f6;color:var(--text-sub);border:1px solid var(--border)}

/* Card common */
.card{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius-lg);
  margin-top:10px;overflow:hidden;transition:border-color .15s, box-shadow .15s;
  box-shadow:0 1px 2px rgba(15,23,42,.03);
}
.card.expanded{border-color:var(--accent-soft);box-shadow:0 4px 14px rgba(99,102,241,.10)}

.num-badge{
  width:30px;height:30px;border-radius:50%;
  background:#e0e7ff;
  color:#4f46e5;font-weight:700;font-size:13px;
  display:flex;align-items:center;justify-content:center;
  flex-shrink:0;
  font-variant-numeric:tabular-nums;
}
.num-badge.gold{background:#fef3c7;color:#92400e}
.num-badge.silver{background:#e5e7eb;color:#374151}
.num-badge.bronze{background:#fed7aa;color:#9a3412}

.card-chev{color:var(--text-mute);transition:transform .2s;font-size:13px;align-self:center}
.card.expanded .card-chev{transform:rotate(180deg)}
.card-detail{display:none;border-top:1px solid var(--border-soft);background:#fbfbfd}
.card.expanded .card-detail{display:block}

/* ===== Setlist ===== */
.sl-card-head{
  padding:12px 14px;display:grid;grid-template-columns:30px 1fr 14px;gap:12px;
  align-items:center;cursor:pointer;
}
.sl-body{min-width:0}
.sl-top{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:6px}
.room-tag{
  display:inline-block;padding:3px 10px;font-size:11.5px;font-weight:600;
  border-radius:6px;line-height:1.4;border:1px solid;
}
.sl-date{font-size:12.5px;color:var(--text-mute);white-space:nowrap;font-variant-numeric:tabular-nums}
.sl-song{
  font-weight:700;font-size:17px;color:var(--primary);margin:0;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.3;
  display:flex;align-items:center;gap:6px;
}
.sl-song i.song-mic{color:var(--accent);font-size:14px;flex-shrink:0}
.sl-meta{font-size:13px;color:var(--text-sub);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:3px;line-height:1.4}
.sl-meta .sep{margin:0 8px;color:var(--text-mute)}

.detail-table{padding:10px 14px 12px}
.detail-row{
  display:grid;grid-template-columns:104px 1fr;gap:10px;padding:10px 14px;
  background:#fff;align-items:start;
  border:1px solid var(--border-soft);border-top:none;
}
.detail-row:first-child{border-top:1px solid var(--border-soft);border-top-left-radius:8px;border-top-right-radius:8px}
.detail-row:last-of-type{border-bottom-left-radius:8px;border-bottom-right-radius:8px}
.detail-row .lbl{font-size:12.5px;color:var(--accent);font-weight:600;display:flex;align-items:center;gap:6px}
.detail-row .lbl i{color:var(--accent);font-size:12px}
.detail-row .val{font-size:14px;color:var(--primary);word-break:break-word;line-height:1.5}

/* 控えめなボタン */
.confirm-btn{
  display:flex;align-items:center;justify-content:center;gap:8px;
  margin:10px 14px 14px;padding:10px 14px;
  background:#f3f4f6;color:var(--text-sub);
  border:1px solid var(--border);border-radius:10px;
  font-weight:600;font-size:13px;cursor:pointer;width:calc(100% - 28px);
}
.confirm-btn:hover{background:#eef2ff;color:var(--accent);border-color:var(--accent-soft)}
.confirm-btn i.fa-chevron-right{margin-left:auto;font-size:11px}

/* Pagination */
.pager{
  display:flex;justify-content:center;align-items:center;gap:6px;margin-top:18px;flex-wrap:wrap;
}
.pager button{
  background:#fff;border:1px solid var(--border);border-radius:8px;padding:6px 12px;
  font-size:13px;font-weight:600;color:var(--text-sub);cursor:pointer;min-width:36px;
}
.pager button.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.pager button:disabled{opacity:.4;cursor:not-allowed}
.pager .info{font-size:12px;color:var(--text-mute);margin:0 8px}

/* ===== Cool ===== */
.cool-head{
  padding:12px 14px;display:grid;grid-template-columns:30px 1fr auto 14px;gap:12px;align-items:center;cursor:pointer;
}
.cool-info{min-width:0}
.cool-anime{font-weight:700;font-size:14.5px;color:var(--primary);line-height:1.3;
  display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}
.cool-types{margin-top:6px;display:flex;gap:6px;flex-wrap:wrap}
.cool-metrics{display:flex;gap:14px;align-items:center}

/* flat metric (used in cool, ranking-row, notable, trend-pickup) */
.flat-metric{display:flex;flex-direction:column;align-items:center;gap:2px;min-width:46px}
.flat-metric .lbl{font-size:10.5px;color:var(--text-sub);font-weight:600;letter-spacing:.02em}
.flat-metric .val-row{display:flex;align-items:center;gap:6px}
.flat-metric .val-row b{font-size:18px;font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums;line-height:1}
.flat-metric .icon-circle{
  width:24px;height:24px;border-radius:50%;display:flex;align-items:center;justify-content:center;
  font-size:11px;flex-shrink:0;
}
.flat-metric .icon-circle.user{background:var(--accent-bg);color:var(--accent)}
.flat-metric .icon-circle.song{background:var(--orange-bg);color:var(--orange)}

/* Cool detail rows */
.song-row{
  display:grid;grid-template-columns:38px 1fr auto;gap:12px;padding:12px 16px;
  border-bottom:1px solid var(--border-soft);align-items:center;
}
.song-row:last-child{border-bottom:none}
.song-info-wrap{min-width:0}
.song-name{font-weight:700;font-size:14px;color:var(--primary);line-height:1.3;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.song-artist{font-size:12px;color:var(--text-sub);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:3px}
.song-metrics{display:flex;gap:12px;align-items:center}
.song-metrics .flat-metric{min-width:40px}
.song-metrics .flat-metric .val-row b{font-size:15px}
.song-metrics .flat-metric .icon-circle{width:20px;height:20px;font-size:9.5px}

/* ===== Ranking ===== */
/* Unified flat layout for all ranks. Top3 differentiated by background + badge. */
.rank-row-flat{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius-lg);margin-top:8px;
  padding:12px 16px;display:grid;grid-template-columns:30px 1fr auto auto;gap:14px;align-items:center;
  box-shadow:0 1px 2px rgba(15,23,42,.03);
}
.rank-row-flat.gold{background:linear-gradient(180deg,#fffbeb 0%,#fff 80%);border-color:#fde68a}
.rank-row-flat.silver{background:linear-gradient(180deg,#f9fafb 0%,#fff 80%);border-color:#d1d5db}
.rank-row-flat.bronze{background:linear-gradient(180deg,#fff7ed 0%,#fff 80%);border-color:#fed7aa}
.rank-row-flat .rank-info{min-width:0}
.rank-row-flat .rank-anime{font-weight:700;font-size:14.5px;color:var(--primary);line-height:1.3;
  display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden}
.rank-row-flat .rank-sub{font-size:12px;color:var(--text-sub);margin-top:3px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-row-flat .rank-types-inline{margin-top:5px;display:flex;gap:5px;flex-wrap:wrap}
.rank-row-flat .flat-metric{min-width:54px}
.rank-row-flat .flat-metric .val-row b{font-size:18px}
.rank-row-flat .flat-metric .icon-circle{width:26px;height:26px;font-size:12px}

/* ===== Trend ===== */
.trend-stats{display:flex;flex-direction:row;gap:6px;margin-top:14px}
.trend-stat{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius);
  padding:8px 10px;display:flex;align-items:center;gap:8px;
  flex:1 1 0;min-width:0;
}
.trend-stat .ico{
  width:28px;height:28px;border-radius:8px;display:flex;align-items:center;justify-content:center;
  font-size:12px;flex-shrink:0;
}
.trend-stat .ico.up{background:var(--accent-bg);color:var(--accent)}
.trend-stat .ico.new{background:var(--green-bg);color:var(--green)}
.trend-stat .ico.max{background:var(--orange-bg);color:var(--orange)}
.trend-stat .lbl{flex:1;min-width:0;font-size:11.5px;color:var(--text-sub);font-weight:600;line-height:1.25}
.trend-stat .val{font-size:16px;font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums;white-space:nowrap}
.trend-stat .val small{font-size:10.5px;font-weight:600;color:var(--text-sub);margin-left:2px}

.trend-pickup{
  background:linear-gradient(135deg,#fee2e2 0%,#fff5f5 60%,#fecaca 100%);
  border:1px solid #fecaca;border-radius:var(--radius-lg);
  margin-top:14px;padding:12px 14px;
}
.trend-pickup-head{display:flex;align-items:center;gap:6px;margin-bottom:8px;font-weight:700;font-size:13.5px;color:var(--primary)}
.trend-pickup-head i{color:var(--red)}
/* Compact pickup row: same layout as notable-row, differentiated by parent gradient and badge */
.trend-pickup-row{
  background:#fff;border:1px solid #fecaca;border-radius:var(--radius-lg);
  padding:11px 14px;display:grid;grid-template-columns:30px 1fr auto auto;gap:12px;align-items:center;
  box-shadow:0 1px 2px rgba(239,68,68,.08);position:relative;
}
.trend-pickup-row .num-badge{background:#fee2e2;color:#b91c1c}
.trend-pickup-row .pickup-info{min-width:0}
.trend-pickup-row .pickup-anime{font-weight:700;font-size:14.5px;color:var(--primary);line-height:1.3;
  display:-webkit-box;-webkit-line-clamp:1;-webkit-box-orient:vertical;overflow:hidden}
.trend-pickup-row .pickup-sub{font-size:12px;color:var(--text-sub);margin-top:3px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.trend-pickup-row .pickup-types-inline{margin-top:5px;display:flex;gap:5px;flex-wrap:wrap}
.trend-pickup-row .flat-metric{min-width:54px}
.trend-pickup-row .flat-metric .val-row b{font-size:18px}
.trend-pickup-row .flat-metric .icon-circle{width:26px;height:26px;font-size:12px}
.thumb-square{
  width:104px;height:104px;border-radius:12px;background:#e5e7eb;overflow:hidden;flex-shrink:0;
  display:flex;align-items:center;justify-content:center;color:var(--text-mute);font-size:24px;position:relative;
}
.thumb-square img{width:100%;height:100%;object-fit:cover}

.notable-head{
  margin-top:18px;font-size:14.5px;font-weight:700;color:var(--primary);display:flex;align-items:center;gap:6px;
}
.notable-head i{color:var(--accent)}

.notable-row{
  background:#fff;border:1px solid var(--border);border-radius:var(--radius-lg);margin-top:8px;
  padding:11px 14px;display:grid;grid-template-columns:30px 1fr auto auto;gap:12px;align-items:center;
  box-shadow:0 1px 2px rgba(15,23,42,.03);position:relative;
}
.notable-info{min-width:0}
.notable-anime{font-weight:700;font-size:14px;color:var(--primary);line-height:1.3;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.notable-artist{font-size:11.5px;color:var(--text-sub);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:3px}
.notable-types{margin-top:5px;display:flex;gap:5px;flex-wrap:wrap}
.notable-row .flat-metric{min-width:48px}
.notable-row .flat-metric .val-row b{font-size:16px}
.notable-row .flat-metric .icon-circle{width:24px;height:24px;font-size:11px}

/* Growth rate badge on top-right */
.growth-badge{
  position:absolute;top:-6px;right:8px;
  background:#fff;border:1px solid var(--border);border-radius:10px;
  padding:1px 7px;font-size:10.5px;font-weight:700;
  font-variant-numeric:tabular-nums;line-height:1.5;
  box-shadow:0 1px 2px rgba(15,23,42,.06);
}
.growth-badge.up{background:#fef2f2;border-color:#fecaca;color:#dc2626}
.growth-badge.down{background:#eff6ff;border-color:#bfdbfe;color:#2563eb}
.growth-badge.flat{background:#f3f4f6;border-color:var(--border);color:var(--text-sub)}
.growth-badge.new{background:var(--green-bg);border-color:var(--green-bd);color:var(--green)}

/* Modal */
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
.modal-head .ttl{flex:1;font-weight:700;font-size:15px;color:var(--primary)}
.modal-head .ttl small{display:block;font-size:11.5px;color:var(--text-sub);margin-top:2px;font-weight:500}
.modal-close{
  width:32px;height:32px;border-radius:50%;background:#f3f4f6;border:none;color:var(--text-sub);
  display:flex;align-items:center;justify-content:center;cursor:pointer;
}
.modal-body{flex:1;overflow-y:auto;padding:8px 12px 16px;-webkit-overflow-scrolling:touch}
.modal-summary{
  background:var(--accent-bg);border-radius:8px;padding:9px 12px;margin:6px 0 10px;font-size:12.5px;color:var(--accent);font-weight:600;
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

/* Filter popover */
.popover{
  display:none;position:absolute;top:100%;right:0;margin-top:6px;background:#fff;
  border:1px solid var(--border);border-radius:12px;box-shadow:0 8px 24px rgba(0,0,0,.1);
  padding:14px;width:320px;z-index:50;
}
.popover.open{display:block}
.popover h4{margin:0 0 8px;font-size:11.5px;color:var(--text-sub);font-weight:700;text-transform:uppercase;letter-spacing:.05em}
.popover .room-search{
  width:100%;border:1px solid var(--border);border-radius:8px;padding:8px 10px;font-size:13px;outline:none;margin-bottom:8px;
}
.room-grid{
  display:grid;grid-template-columns:1fr 1fr;gap:6px;max-height:240px;overflow-y:auto;
}
.room-chip{
  display:flex;align-items:center;justify-content:center;padding:8px 10px;font-size:12px;
  border-radius:8px;border:1px solid var(--border);background:#fff;color:var(--text-sub);cursor:pointer;
  text-align:center;font-weight:500;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
}
.room-chip:hover{border-color:var(--accent-soft)}
.room-chip.selected{background:var(--accent);color:#fff;border-color:var(--accent);font-weight:600}
.popover .actions{display:flex;justify-content:flex-end;gap:8px;margin-top:12px}
.btn-clear{background:#f3f4f6;color:var(--text-sub);border:none;border-radius:8px;padding:7px 14px;font-size:12.5px;cursor:pointer;font-weight:600}
.btn-apply{background:var(--accent);color:#fff;border:none;border-radius:8px;padding:7px 14px;font-size:12.5px;cursor:pointer;font-weight:600}
.toolbar-rel{position:relative}

.empty{padding:40px 16px;text-align:center;color:var(--text-mute);font-size:14px}

/* Env */
.env-section{margin-top:14px}
.env-work-row{
  background:#fff;border:1px solid var(--border);border-radius:12px;padding:11px;margin-bottom:6px;
  display:flex;gap:11px;align-items:center;
}
.env-work-thumb{width:54px;height:54px;border-radius:10px;background:#e5e7eb;overflow:hidden;display:flex;align-items:center;justify-content:center;color:var(--text-mute);flex-shrink:0;font-size:18px}
.env-work-thumb img{width:100%;height:100%;object-fit:cover}
.env-work-name{flex:1;font-weight:600;font-size:14px;color:var(--primary);min-width:0;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.env-upload-btn{
  background:var(--accent);color:#fff;border:none;border-radius:8px;padding:9px 13px;font-size:12.5px;cursor:pointer;
  display:inline-flex;align-items:center;gap:5px;
}
.env-upload-btn:hover{filter:brightness(1.05)}
.env-upload-btn input{display:none}
.env-status{margin:8px 0;padding:10px 12px;background:var(--accent-bg);border-radius:8px;color:var(--accent);font-size:13px;display:none;font-weight:600}
.env-status.show{display:block}
.env-lock{padding:30px 16px;display:flex;justify-content:center}
.env-lock-card{
  background:#fff;border:1px solid var(--border);border-radius:14px;padding:24px 20px;
  width:100%;max-width:340px;text-align:center;box-shadow:0 1px 4px rgba(15,23,42,.04);
}
.env-lock-ico{
  width:48px;height:48px;border-radius:50%;background:var(--accent-bg);color:var(--accent);
  display:flex;align-items:center;justify-content:center;font-size:18px;margin:0 auto 10px;
}
.env-lock-title{font-size:13.5px;font-weight:600;color:var(--text-sub);margin-bottom:14px}
.env-lock-input{
  width:100%;padding:10px 12px;border:1px solid var(--border);border-radius:10px;
  font-size:14px;outline:none;margin-bottom:10px;text-align:center;letter-spacing:.2em;
}
.env-lock-input:focus{border-color:var(--accent)}
.env-lock-btn{
  width:100%;padding:10px;background:linear-gradient(180deg,var(--accent),var(--accent-2));
  color:#fff;border:none;border-radius:10px;font-weight:700;font-size:14px;cursor:pointer;
}
.env-lock-btn:hover{filter:brightness(1.05)}
.env-lock-msg{margin-top:10px;font-size:12px;color:var(--red);min-height:16px}
/* Site-wide password gate (shown until unlocked) */
.site-lock{
  position:fixed;inset:0;z-index:9999;background:var(--bg);
  display:flex;align-items:center;justify-content:center;padding:30px 16px;overflow:auto;
}
body.site-locked{overflow:hidden}
body.site-locked > *:not(#siteLockScreen){display:none !important}
body.dark .site-lock{background:var(--bg)}
.env-csv-wrap{margin-top:8px}
.env-csv-actions{display:flex;gap:6px;justify-content:flex-end;margin-bottom:6px}
.env-csv-text{
  width:100%;min-height:240px;max-height:50vh;
  border:1px solid var(--border);border-radius:10px;padding:10px 12px;
  font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;
  font-size:12px;line-height:1.5;outline:none;background:#fff;color:var(--primary);
  white-space:pre;overflow:auto;resize:vertical;box-sizing:border-box;
}
.env-csv-text:focus{border-color:var(--accent)}
.env-csv-msg{margin-top:6px;font-size:12px;color:var(--text-sub);min-height:16px}
.env-csv-msg.error{color:var(--red)}
.env-csv-msg.ok{color:var(--green)}

/* ===== ASS tag extension editor ===== */
.ass-intro{font-size:12.5px;color:var(--text-sub);margin:6px 0 0;line-height:1.5}
.ass-msg{margin:8px 0 0;font-size:12px;color:var(--text-sub);min-height:16px;font-weight:600}
.ass-msg.error{color:var(--red)}
.ass-msg.ok{color:var(--green)}
.ass-actions{display:flex;gap:6px;justify-content:flex-end;align-items:center;flex-wrap:wrap;margin:8px 0 10px}
.ass-toolbar-note{background:#fff;border:1px solid var(--border);border-radius:10px;padding:9px 12px;font-size:12px;color:var(--text-sub);line-height:1.5;margin-top:8px}
.ass-list{display:flex;flex-direction:column;gap:10px;margin-top:10px}
.ass-card{background:#fff;border:1px solid var(--border);border-radius:var(--radius-lg);overflow:hidden;box-shadow:0 1px 2px rgba(15,23,42,.03)}
.ass-card-head{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid var(--border-soft)}
.ass-card-no{width:30px;height:30px;border-radius:50%;background:var(--accent-bg);color:var(--accent);display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;flex-shrink:0;font-variant-numeric:tabular-nums}
.ass-input{flex:1;min-width:0;border:1px solid var(--border);border-radius:8px;background:#fff;color:var(--primary);padding:8px 10px;font-size:13px;font-weight:700;outline:none}
.ass-input:focus,.ass-text:focus{border-color:var(--accent)}
.ass-delete{width:34px;height:34px;border-radius:8px;border:1px solid var(--border);background:#fff;color:var(--text-sub);display:flex;align-items:center;justify-content:center;cursor:pointer;flex-shrink:0}
.ass-delete:hover{color:var(--red);border-color:#fecaca;background:#fef2f2}
.ass-fields{padding:12px;display:grid;grid-template-columns:1fr;gap:10px}
.ass-field-label{display:flex;align-items:center;gap:6px;font-size:12px;font-weight:800;color:var(--accent);margin-bottom:5px;letter-spacing:.02em}
.ass-text{width:100%;min-height:70px;border:1px solid var(--border);border-radius:10px;background:#fff;color:var(--text);padding:10px 12px;font-size:12.5px;line-height:1.55;font-family:"Noto Sans JP",ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;resize:vertical;outline:none;white-space:pre;overflow:auto}
.ass-text.char{min-height:112px}
.ass-raw-panel{display:none;background:#fff;border:1px solid var(--border);border-radius:var(--radius-lg);padding:10px;margin-top:10px}
.ass-raw-panel.show{display:block}
.ass-raw-panel .env-csv-text{min-height:460px}
.ass-hidden{display:none !important}

/* Print */
@media print{
  *{-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important}
  .app-header,.top-cards,.tabs-wrap,.tab-toolbar,.update-line,.count-line,.confirm-btn,.dl-row,.pager{display:none !important}
  body{padding:0}
  .tab-content{display:block !important;max-width:none}
  .card{break-inside:avoid;page-break-inside:avoid}
  .card-detail{display:block !important}
  .card-chev{display:none}
}

@media (max-width:520px){
  .app-header,.top-cards,.tabs-wrap,.tab-toolbar,.update-line,.count-line,.tab-content{padding-left:8px;padding-right:8px}
  .tab-content{padding-bottom:80px}
}
@media (max-width:420px){
  body{font-size:14.5px}
  .top-cards{grid-template-columns:1fr 1fr;gap:8px}
  .notable-row{grid-template-columns:30px 1fr auto auto;gap:10px;padding:10px 12px}
  .rank-row-flat{padding:11px 12px;gap:12px}
}
</style>
</head>
<body class="site-locked">

<!-- Site-wide password gate -->
<div id="siteLockScreen" class="site-lock">
  <div class="env-lock-card">
    <div class="env-lock-ico"><i class="fas fa-lock"></i></div>
    <div class="env-lock-title">Karaoke Dashboard はパスワードで保護されています</div>
    <input type="password" id="siteLockPwInput" class="env-lock-input" placeholder="パスワードを入力" autocomplete="off">
    <button class="env-lock-btn" id="siteLockPwBtn" type="button">解除</button>
    <div class="env-lock-msg" id="siteLockPwMsg"></div>
  </div>
</div>
<script>
(function(){
  var SITE_PASSWORD='0515';
  function unlock(){
    sessionStorage.setItem('siteUnlocked','1');
    document.body.classList.remove('site-locked');
    var s=document.getElementById('siteLockScreen'); if(s) s.style.display='none';
  }
  if(sessionStorage.getItem('siteUnlocked')==='1'){
    unlock();
    return;
  }
  function tryUnlock(){
    var v=document.getElementById('siteLockPwInput').value;
    var msg=document.getElementById('siteLockPwMsg');
    if(v===SITE_PASSWORD){ msg.innerText=''; unlock(); }
    else{ msg.innerText='パスワードが正しくありません'; document.getElementById('siteLockPwInput').value=''; }
  }
  document.addEventListener('DOMContentLoaded', function(){
    var btn=document.getElementById('siteLockPwBtn');
    var inp=document.getElementById('siteLockPwInput');
    if(btn) btn.addEventListener('click', tryUnlock);
    if(inp){ inp.addEventListener('keydown', function(e){ if(e.key==='Enter') tryUnlock(); }); inp.focus(); }
  });
})();
</script>

<div class="app-header">
  <div class="brand">Karaoke Dashboard</div>
  <div style="display:flex;align-items:center;gap:8px">
    <a class="theme-toggle" id="viewerLink" href="viewer.html" title="Karaoke Viewer を開く" aria-label="Karaoke Viewer を開く" style="text-decoration:none;width:auto;padding:0 10px;font-size:11px;font-weight:700;color:var(--accent);border-color:var(--accent-soft)"><i class="fas fa-external-link-alt" style="margin-right:5px"></i>Viewer</a>
    <button class="theme-toggle" id="themeToggle" type="button" title="ダークモード切替" aria-label="ダークモード切替"><i class="fas fa-moon"></i></button>
  </div>
</div>

<div class="top-cards">
  <div class="top-card">
    <div class="ico"><i class="fas fa-network-wired"></i></div>
    <div style="flex:1;min-width:0">
      <div class="lbl">保存時ポート</div>
      <input type="number" id="exportPort" value="11059">
    </div>
  </div>
  <div class="top-card">
    <div class="ico"><i class="fas fa-link"></i></div>
    <div style="flex:1;min-width:0">
      <div class="lbl">検索リンク</div>
      <select id="exportLinkType">
        <option value="eve">Everything</option>
        <option value="ykr">ゆかりすたー</option>
      </select>
    </div>
  </div>
</div>

<div class="tabs-wrap">
  <div class="tabs">
    <button class="tab-btn active" data-tab="setlist"><i class="fas fa-music"></i> セットリスト</button>
    <button class="tab-btn" data-tab="cool"><i class="far fa-clock"></i> クール集計</button>
    <button class="tab-btn" data-tab="ranking"><i class="fas fa-trophy"></i> ランキング</button>
    <button class="tab-btn" data-tab="trend"><i class="fas fa-chart-line"></i> 推移</button>
    <button class="tab-btn" data-tab="ass"><i class="fas fa-tags"></i> ASSタグ拡張</button>
    <button class="tab-btn" data-tab="env"><i class="fas fa-cog"></i> 環境設定</button>
  </div>
</div>

<div class="update-line"><i class="far fa-clock" style="margin-right:4px"></i><span id="updateLine">__UPDATE__ 更新</span></div>

<!-- Setlist -->
<div class="tab-content active" id="tab-setlist">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <div class="search-pill">
        <i class="fas fa-search"></i>
        <input type="text" id="slSearch" placeholder="曲名・作品名・歌手名・歌った人で検索">
      </div>
      <div class="toolbar-rel">
        <button class="icon-btn" id="slFilterBtn"><i class="fas fa-sliders"></i></button>
        <div class="popover" id="slPopover">
          <h4>部屋でフィルタ（複数選択可）</h4>
          <input type="text" id="roomSearch" class="room-search" placeholder="部屋名で絞り込み">
          <div class="room-grid" id="slRoomChips"></div>
          <div class="actions">
            <button class="btn-clear" id="slClearFilter">クリア</button>
            <button class="btn-apply" id="slApplyFilter">適用</button>
          </div>
        </div>
      </div>
    </div>
    <div class="dl-row">
      <select class="pill-select" id="slDlYear" title="HTML保存期間"></select>
      <button class="dl-btn ghost" onclick="downloadSetlistHTML()"><i class="fas fa-file-download"></i> HTML保存</button>
    </div>
  </div>
  <div class="count-line"><i class="fas fa-clipboard-list"></i><span id="slCount">0</span> 件</div>
  <div id="slList"></div>
  <div class="pager" id="slPager"></div>
</div>

<!-- Cool -->
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
      <button class="dl-btn ghost compact" onclick="downloadCoolHTML('current')" title="このクール保存"><i class="fas fa-file-download"></i><span class="dl-lbl">このクール</span></button>
      <button class="dl-btn compact" onclick="downloadCoolHTML('all')" title="全クール保存"><i class="fas fa-file-download"></i><span class="dl-lbl">全クール</span></button>
    </div>
  </div>
  <div class="count-line"><i class="far fa-bookmark"></i><span id="coolCount">0</span> 作品</div>
  <div id="coolList"></div>
</div>

<!-- Ranking -->
<div class="tab-content" id="tab-ranking">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <select class="pill-select" id="rankCat"></select>
      <select class="pill-select" id="rankMode">
        <option value="count">歌唱数ランキング</option>
        <option value="user">歌唱人数ランキング</option>
      </select>
      <button class="dl-btn ghost compact" onclick="downloadRankingHTML('current')" title="このクール保存"><i class="fas fa-file-download"></i><span class="dl-lbl">このクール</span></button>
      <button class="dl-btn compact" onclick="downloadRankingHTML('all')" title="全クール保存"><i class="fas fa-file-download"></i><span class="dl-lbl">全クール</span></button>
    </div>
  </div>
  <div class="count-line"><i class="fas fa-list-ol"></i><span id="rankCount">0</span> 件</div>
  <div id="rankList"></div>
</div>

<!-- Trend -->
<div class="tab-content" id="tab-trend">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <select class="pill-select" id="trendPeriod"></select>
      <select class="pill-select" id="trendSort">
        <option value="surge">急上昇順</option>
        <option value="new">新規ランクイン順</option>
        <option value="max">最大伸び順</option>
      </select>
      <select class="pill-select" id="trendMode">
        <option value="count">歌唱数</option>
        <option value="user">歌唱人数</option>
      </select>
      <button class="dl-btn ghost compact" onclick="downloadTrendHTML()" title="HTML保存"><i class="fas fa-file-download"></i><span class="dl-lbl">HTML保存</span></button>
    </div>
  </div>
  <div id="trendBody"></div>
</div>

<!-- ASS Tags -->
<div class="tab-content" id="tab-ass">
  <div class="tab-toolbar">
    <div class="toolbar-row">
      <button class="dl-btn ghost compact" id="assLoad" type="button"><i class="fas fa-cloud-download-alt"></i><span class="dl-lbl">読み込み</span></button>
      <button class="dl-btn compact" id="assSave" type="button"><i class="fas fa-save"></i><span class="dl-lbl">保存</span></button>
      <button class="dl-btn ghost compact" id="assAdd" type="button"><i class="fas fa-plus"></i><span class="dl-lbl">項目追加</span></button>
      <button class="dl-btn ghost compact" id="assRawToggle" type="button"><i class="fas fa-code"></i><span class="dl-lbl">テキスト編集</span></button>
    </div>
    <div class="ass-toolbar-note"><i class="fas fa-info-circle"></i> 【タイトル】を1項目として扱い、各項目の「行」「文字」を別ブロックで編集します。未保存の初期値は下のプリセットです。</div>
  </div>
  <div class="count-line"><i class="fas fa-tags"></i><span id="assCount">0</span> 項目</div>
  <div class="ass-msg" id="assMsg"></div>
  <div id="assList" class="ass-list"></div>
  <div id="assRawPanel" class="ass-raw-panel">
    <div class="ass-actions">
      <button class="dl-btn ghost compact" id="assApplyRaw" type="button"><i class="fas fa-check"></i><span class="dl-lbl">テキストを項目へ反映</span></button>
    </div>
    <textarea id="assRawText" class="env-csv-text" placeholder="【タイトル】&#10;行：...&#10;文字：..." spellcheck="false"></textarea>
  </div>
</div>

<!-- Env -->
<div class="tab-content" id="tab-env">
  <div id="envContent">
    <div class="count-line" style="margin-top:18px"><i class="fas fa-file-csv"></i> cool_analysis.csv 編集</div>
    <div class="env-csv-wrap">
      <div class="env-csv-actions">
        <button class="dl-btn ghost compact" id="envCsvLoad" type="button"><i class="fas fa-cloud-download-alt"></i><span class="dl-lbl">読み込み</span></button>
        <button class="dl-btn compact" id="envCsvSave" type="button"><i class="fas fa-save"></i><span class="dl-lbl">保存</span></button>
      </div>
      <textarea id="envCsvText" class="env-csv-text" placeholder="「読み込み」を押すと cool_analysis.csv の内容を表示します。編集後「保存」を押すとGAS経由で書き込みます。" spellcheck="false"></textarea>
      <div class="env-csv-msg" id="envCsvMsg"></div>
    </div>
  </div>
</div>

<div class="modal-overlay" id="modalOverlay">
  <div class="modal">
    <div class="modal-head">
      <div class="ttl" id="modalTitle"></div>
      <button class="modal-close" onclick="closeModal()"><i class="fas fa-times"></i></button>
    </div>
    <div class="modal-body" id="modalBody"></div>
  </div>
</div>

<script>
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
const ASS_TAGS_FILENAME = "__ASS_TAGS_FILENAME__";
const ASS_TAG_DEFAULT_TEXT = __ASS_TAG_DEFAULT_JSON__;
const PAGE_SIZE = 200;

document.getElementById('updateLine').innerText = UPDATE_TS + ' 更新';

// Dark mode toggle
(function(){
  const tg = document.getElementById('themeToggle');
  if(!tg) return;
  function apply(d){
    document.body.classList.toggle('dark', d);
    const ic = tg.querySelector('i');
    if(ic) ic.className = d ? 'fas fa-sun' : 'fas fa-moon';
  }
  let dark = false;
  try { dark = localStorage.getItem('darkMode') === '1'; } catch(e){}
  apply(dark);
  tg.addEventListener('click', ()=>{
    dark = !dark;
    try { localStorage.setItem('darkMode', dark ? '1' : '0'); } catch(e){}
    apply(dark);
  });
})();

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
function escHtml(s){return String(s||"").replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function escAttr(s){return String(s||"").replace(/'/g,"&#39;").replace(/"/g,'&quot;');}

function roomColor(name){
  let h = 0;
  for(let i=0;i<name.length;i++) h = (h*31 + name.charCodeAt(i)) & 0xffffffff;
  const hue = Math.abs(h) % 360;
  return {bg:`hsl(${hue},65%,94%)`, fg:`hsl(${hue},45%,38%)`, bd:`hsl(${hue},55%,84%)`};
}
function roomTagHtml(name){
  const c = roomColor(name);
  return `<span class="room-tag" style="background:${c.bg};color:${c.fg};border-color:${c.bd}">${escHtml(name)}</span>`;
}

function typeClass(t){
  if(!t) return 'tp-none';
  const u = String(t).toUpperCase();
  if(u.indexOf('OP')>=0) return 'tp-op';
  if(u.indexOf('ED')>=0) return 'tp-ed';
  if(u.indexOf('IN')>=0) return 'tp-in';
  return 'tp-none';
}
function typePillHtml(t, count){
  if(typeof count === 'number'){
    const cls = typeClass(t);
    const dispVal = count > 0 ? count : '-';
    return `<span class="type-pill ${cls}">${escHtml(t)} <b>${dispVal}</b></span>`;
  }
  const lbl = t || '-';
  return `<span class="type-pill ${typeClass(t)}">${escHtml(lbl)}</span>`;
}
function typeChipHtml(t){
  return `<div class="type-chip ${typeClass(t)}">${escHtml(t||'-')}</div>`;
}

function getThumbUrl(cat, work){
  return null;
}

function findHistoryMatches(workName,songName){
  const sn=jsNormalize(songName);const wn=jsNormalize(workName);
  if(!sn&&!wn) return [];
  return HISTORY.filter(h=>{
    let songOk=false,workOk=false;
    if(sn){
      if(/^[A-Z0-9 ]+$/.test(sn)){
        const re=new RegExp('(?:^|[^A-Z0-9])'+sn.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'(?:[^A-Z0-9]|$)','i');
        songOk=re.test(h.sn);
      } else songOk=h.sn.indexOf(sn)>=0;
    }
    if(wn) workOk=(h.sn.indexOf(wn)>=0)||(h.wn.indexOf(wn)>=0);
    if(sn&&wn) return songOk&&workOk;
    if(sn) return songOk;
    if(wn) return workOk;
    return false;
  });
}

function openSingersModal(workName, songName){
  const matches = findHistoryMatches(workName, songName);
  const userCounts = {};
  matches.forEach(m=>{
    if(!userCounts[m.u]) userCounts[m.u]={count:0,last:m.d,room:m.rm};
    userCounts[m.u].count++;
    if(m.d>userCounts[m.u].last){userCounts[m.u].last=m.d;userCounts[m.u].room=m.rm;}
  });
  const users = Object.entries(userCounts).sort((a,b)=>b[1].count-a[1].count);
  document.getElementById('modalTitle').innerHTML = escHtml(songName)+'<small>'+escHtml(workName)+' - '+matches.length+'件 / '+users.length+'人</small>';
  let body = `<div class="modal-summary"><i class="fas fa-users"></i> ${users.length}人がこの曲を歌っています（合計${matches.length}回）</div>`;
  if(users.length===0) body += '<div class="empty">履歴が見つかりませんでした</div>';
  else users.forEach(([u,info])=>{
    body += `<div class="modal-row"><div class="top"><div class="user"><i class="fas fa-microphone"></i> ${escHtml(u)} <span style="color:var(--accent);font-size:11.5px;margin-left:4px">×${info.count}</span></div><div class="date">${escHtml(info.last)}</div></div><div class="meta">最新: ${escHtml(info.room)}</div></div>`;
  });
  document.getElementById('modalBody').innerHTML = body;
  document.getElementById('modalOverlay').classList.add('active');
}
function closeModal(){document.getElementById('modalOverlay').classList.remove('active');}
document.getElementById('modalOverlay').addEventListener('click',e=>{if(e.target.id==='modalOverlay') closeModal();});

document.querySelectorAll('.tab-btn').forEach(b=>{
  b.addEventListener('click',()=>{
    document.querySelectorAll('.tab-btn').forEach(x=>x.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(x=>x.classList.remove('active'));
    b.classList.add('active');
    document.getElementById('tab-'+b.dataset.tab).classList.add('active');
  });
});

function toggleCard(headEl){headEl.parentElement.classList.toggle('expanded');}

// ===== Setlist =====
const SETLIST = HISTORY.map((h,i)=>{const n=parseFloat(h.o);return {...h, idx:i, _orderNum: Number.isFinite(n) ? n : -Infinity};});

// ============================================================
// 【固定ソートルール / DO NOT BREAK】
//   セットリストの並び順は以下の2段階比較に固定する:
//     第一順位: 取得日 (x.d, "YYYY/MM/DD" 文字列) を降順 — 新しい日付が上
//     第二順位: 順番   (x._orderNum, 数値)         を降順 — 大きい番号が上
//   この関数 compareSetlistEntries が比較ロジックの唯一の定義。
//   並び替えを行う箇所は arr.sort(compareSetlistEntries) を呼ぶこと。
//   ソートロジックを変更する必要が出た場合は、必ずこの関数 1 箇所のみを修正し、
//   呼び出し側にインラインで比較関数を書き直さないこと（過去に各所で
//   ばらばらに書き換わって規約が破綻した経緯があるため）。
// ============================================================
function compareSetlistEntries(a, b){
  if(a.d !== b.d) return a.d < b.d ? 1 : -1;
  return b._orderNum - a._orderNum;
}

// 起動時セルフチェック: 規約に違反する変更が入った場合に気付けるようにする。
(function _verifySetlistSortContract(){
  try {
    const sample = [
      {d:'2026/04/29', _orderNum: 5},
      {d:'2026/04/29', _orderNum: 10},
      {d:'2026/04/30', _orderNum: 1},
    ].sort(compareSetlistEntries);
    const ok = sample[0].d === '2026/04/30'
            && sample[1].d === '2026/04/29' && sample[1]._orderNum === 10
            && sample[2].d === '2026/04/29' && sample[2]._orderNum === 5;
    if(!ok) console.warn('[setlist-sort] 固定ソートルールが壊れています。compareSetlistEntries を確認してください。', sample);
  } catch(e){ console.warn('[setlist-sort] self-check failed', e); }
})();

let slState = {rooms:new Set(), keyword:'', page:1, filtered:[]};

function applySlFilter(){
  let arr = SETLIST.slice();
  if(slState.rooms.size>0) arr = arr.filter(x=>slState.rooms.has(x.rm));
  if(slState.keyword){
    const kws = slState.keyword.toUpperCase().replace(/　/g,' ').split(/\s+/).filter(Boolean);
    arr = arr.filter(x=>{
      const t = (x.d+' '+x.rm+' '+x.o+' '+x.sg+' '+x.wk+' '+x.ar+' '+x.u).toUpperCase();
      return kws.every(k=>t.indexOf(k)>=0);
    });
  }
  // 並び順は compareSetlistEntries に集約 (第一=取得日 desc, 第二=順番 desc)
  arr.sort(compareSetlistEntries);
  slState.filtered = arr;
  slState.page = 1;
}

function renderSetlist(){
  applySlFilter();
  drawSetlistPage();
}
function drawSetlistPage(){
  const arr = slState.filtered;
  document.getElementById('slCount').innerText = arr.length.toLocaleString();
  const list = document.getElementById('slList');
  if(arr.length===0){
    list.innerHTML = '<div class="empty">該当データがありません</div>';
    document.getElementById('slPager').innerHTML = '';
    return;
  }
  const totalPages = Math.ceil(arr.length / PAGE_SIZE);
  if(slState.page > totalPages) slState.page = totalPages;
  const startIdx = (slState.page - 1) * PAGE_SIZE;
  const slice = arr.slice(startIdx, startIdx + PAGE_SIZE);

  let html = '';
  slice.forEach((x, idx)=>{
    // バッジには 順番 (x.o) を表示する。同一日付内では 順番 降順で並ぶため、
    // 大きい番号が上に来る並びとバッジ表示が一致する。順番が空/数値でない場合は
    // 行位置にフォールバック。
    const fallbackNum = (startIdx + idx + 1).toString();
    const dispNum = (x.o !== undefined && x.o !== null && String(x.o).trim() !== '')
      ? String(x.o).trim()
      : fallbackNum;
    const metaPieces = [];
    if(x.wk) metaPieces.push(escHtml(x.wk));
    if(x.ar) metaPieces.push(escHtml(x.ar));
    const metaLine = metaPieces.join('<span class="sep">|</span>');

    const songIcon = x.sg && /singing|歌|うた/i.test(x.sg) ? '<i class="fas fa-microphone song-mic"></i>' : '';

    html += `<div class="card">
      <div class="sl-card-head" onclick="toggleCard(this)">
        <div class="num-badge">${escHtml(dispNum)}</div>
        <div class="sl-body">
          <div class="sl-top">
            ${roomTagHtml(x.rm)}
            <span class="sl-date">${escHtml(x.d)}</span>
          </div>
          <div class="sl-song">${songIcon}<span style="overflow:hidden;text-overflow:ellipsis">${escHtml(x.sg)}</span></div>
          ${metaLine?`<div class="sl-meta">${metaLine}</div>`:''}
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
          <div class="detail-row"><div class="lbl"><i class="fas fa-hashtag"></i>順番</div><div class="val">${escHtml(x.o)}</div></div>
        </div>
        <button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(x.wk)}','${escAttr(x.sg)}')">
          <i class="fas fa-users"></i> この曲を歌った人を確認 <i class="fas fa-chevron-right"></i>
        </button>
      </div>
    </div>`;
  });
  list.innerHTML = html;
  drawPager(totalPages);
  window.scrollTo({top:0, behavior:'smooth'});
}
function drawPager(totalPages){
  const p = document.getElementById('slPager');
  if(totalPages<=1){p.innerHTML='';return;}
  const cur = slState.page;
  const win = 2;
  let h = '';
  h += `<button onclick="slGoPage(1)" ${cur===1?'disabled':''}>« 最初</button>`;
  h += `<button onclick="slGoPage(${cur-1})" ${cur===1?'disabled':''}>‹</button>`;
  let start = Math.max(1, cur-win);
  let end = Math.min(totalPages, cur+win);
  if(start>1) h += `<span class="info">…</span>`;
  for(let i=start;i<=end;i++){
    h += `<button onclick="slGoPage(${i})" class="${i===cur?'active':''}">${i}</button>`;
  }
  if(end<totalPages) h += `<span class="info">…</span>`;
  h += `<button onclick="slGoPage(${cur+1})" ${cur===totalPages?'disabled':''}>›</button>`;
  h += `<button onclick="slGoPage(${totalPages})" ${cur===totalPages?'disabled':''}>最後 »</button>`;
  h += `<span class="info">${cur}/${totalPages}</span>`;
  p.innerHTML = h;
}
function slGoPage(n){
  slState.page = n;
  drawSetlistPage();
}

const slPopover = document.getElementById('slPopover');
const slFilterBtn = document.getElementById('slFilterBtn');
slFilterBtn.addEventListener('click',e=>{
  e.stopPropagation();
  slPopover.classList.toggle('open');
  slFilterBtn.classList.toggle('active');
});
document.addEventListener('click',e=>{
  if(!slPopover.contains(e.target) && e.target!==slFilterBtn && !slFilterBtn.contains(e.target)){
    slPopover.classList.remove('open');
    slFilterBtn.classList.remove('active');
  }
});

const slRoomChips = document.getElementById('slRoomChips');
function buildRoomChips(filter){
  slRoomChips.innerHTML = '';
  ROOMS.filter(r=>!filter || r.indexOf(filter)>=0).forEach(r=>{
    const c = document.createElement('div');
    c.className = 'room-chip' + (slState.rooms.has(r)?' selected':'');
    c.innerText = r;
    c.addEventListener('click',()=>{
      if(slState.rooms.has(r)){slState.rooms.delete(r);c.classList.remove('selected');}
      else{slState.rooms.add(r);c.classList.add('selected');}
    });
    slRoomChips.appendChild(c);
  });
}
buildRoomChips('');
document.getElementById('roomSearch').addEventListener('input',e=>buildRoomChips(e.target.value.trim()));
document.getElementById('slClearFilter').addEventListener('click',()=>{
  slState.rooms.clear();
  buildRoomChips('');
  renderSetlist();
});
document.getElementById('slApplyFilter').addEventListener('click',()=>{
  slPopover.classList.remove('open');
  slFilterBtn.classList.remove('active');
  renderSetlist();
});
let slSearchTimer = null;
document.getElementById('slSearch').addEventListener('input',e=>{
  slState.keyword = e.target.value.trim();
  clearTimeout(slSearchTimer);
  slSearchTimer = setTimeout(renderSetlist, 200);
});

// ===== Cool =====
const coolCatSel = document.getElementById('coolCat');
const coolSortSel = document.getElementById('coolSort');
CATS.forEach(c=>{const o=document.createElement('option');o.value=c;o.innerText=c;coolCatSel.appendChild(o);});
coolCatSel.addEventListener('change', renderCool);
coolSortSel.addEventListener('change', renderCool);

function flatMetricHtml(label, value, kind){
  const ic = kind==='user' ? '<i class="fas fa-users"></i>' : '<i class="fas fa-microphone"></i>';
  return `<div class="flat-metric"><div class="lbl">${label}</div><div class="val-row"><b>${value}</b><div class="icon-circle ${kind}">${ic}</div></div></div>`;
}

function buildCoolCard(w, rank, cat){
  // OP/ED/IN ピルは「作成済みの曲数」を表示する。種別ごとの曲が無い場合は '-'。
  const opVal = (w.op_n||0)>0 ? (w.op_created||0) : '-';
  const edVal = (w.ed_n||0)>0 ? (w.ed_created||0) : '-';
  const inVal = (w.in_n||0)>0 ? (w.in_created||0) : '-';
  const opTag = `<span class="type-pill tp-op">OP <b>${opVal}</b></span>`;
  const edTag = `<span class="type-pill tp-ed">ED <b>${edVal}</b></span>`;
  const inTag = `<span class="type-pill tp-in">IN <b>${inVal}</b></span>`;
  // 作品単位の作成数合計（赤字）: X／Y 作成 形式（X=作成済み曲数, Y=総曲数）
  const createdTag = (w.total_created||0) > 0
    ? `<span class="type-pill tp-created" title="作成数の合計"><b>${w.total_created}</b>／${w.songs.length} 作成</span>`
    : '';
  let songsHtml = '';
  w.songs.forEach(s=>{
    const sm = `${flatMetricHtml('人数',s.user_count,'user')}${flatMetricHtml('歌唱数',s.count,'song')}`;
    const createdMark = (s.creation_count||0) > 0
      ? `<span class="song-created-mark" title="offline_list 内の作成数">作成${s.creation_count}</span>`
      : '';
    songsHtml += `<div class="song-row">
      ${typeChipHtml(s.type)}
      <div class="song-info-wrap">
        <div class="song-name">${escHtml(s.song)}${createdMark}</div>
        <div class="song-artist">${escHtml(s.artist)}</div>
      </div>
      <div class="song-metrics">${sm}</div>
    </div>`;
  });
  return `<div class="card">
    <div class="cool-head" onclick="toggleCard(this)">
      <div class="num-badge${rank===1?' gold':rank===2?' silver':rank===3?' bronze':''}">${rank}</div>
      <div class="cool-info">
        <div class="cool-anime">${escHtml(w.anime)}</div>
        <div class="cool-types">${opTag}${edTag}${inTag}${createdTag}</div>
      </div>
      <div class="cool-metrics">
        ${flatMetricHtml('人数',w.total_user,'user')}
        ${flatMetricHtml('歌唱数',w.total_count,'song')}
      </div>
      <i class="fas fa-chevron-down card-chev"></i>
    </div>
    <div class="card-detail">${songsHtml}</div>
  </div>`;
}

function renderCool(){
  const cat = coolCatSel.value;
  const sort = coolSortSel.value;
  const data = COOL_DATA[cat] || {works:[]};
  const works = data.works.slice();
  if(sort==='count') works.sort((a,b)=>b.total_count-a.total_count || b.total_user-a.total_user);
  else if(sort==='user') works.sort((a,b)=>b.total_user-a.total_user || b.total_count-a.total_count);
  else if(sort==='name') works.sort((a,b)=>a.anime.localeCompare(b.anime,'ja'));
  else if(sort==='created') works.sort((a,b)=>{
    const ac = a.songs.reduce((s,x)=>s+(x.creation_count||0),0);
    const bc = b.songs.reduce((s,x)=>s+(x.creation_count||0),0);
    return bc-ac;
  });
  document.getElementById('coolCount').innerText = works.length;
  const list = document.getElementById('coolList');
  if(works.length===0){list.innerHTML='<div class="empty">データがありません</div>';return;}
  let html = '';
  works.forEach((w,i)=>html += buildCoolCard(w, i+1, cat));
  list.innerHTML = html;
}

// ===== Ranking =====
const rankCatSel = document.getElementById('rankCat');
const rankModeSel = document.getElementById('rankMode');
CATS.forEach(c=>{const o=document.createElement('option');o.value=c;o.innerText=c;rankCatSel.appendChild(o);});
rankCatSel.addEventListener('change', renderRanking);
rankModeSel.addEventListener('change', renderRanking);

function buildRankingTop20(cat, mode){
  const all = (RANK_DATA[cat]||[]).slice();
  if(mode==='count') all.sort((a,b)=>b.count-a.count || b.user_count-a.user_count);
  else all.sort((a,b)=>b.user_count-a.user_count || b.count-a.count);
  const top20=[];let pv=null,cr=0;
  for(let i=0;i<all.length;i++){
    const v = mode==='count' ? all[i].count : all[i].user_count;
    if(v!==pv){cr=i+1;pv=v;}
    if(cr>20) break;
    top20.push({...all[i],rank:cr});
  }
  return top20;
}
function buildRankCardHtml(r, mode){
  const grade = r.rank===1?'gold':r.rank===2?'silver':r.rank===3?'bronze':'';
  return `<div class="rank-row-flat ${grade}">
    <div class="num-badge ${grade}">${r.rank}</div>
    <div class="rank-info">
      <div class="rank-anime">${escHtml(r.anime)}</div>
      <div class="rank-sub">${escHtml(r.song)} / ${escHtml(r.artist)}</div>
      <div class="rank-types-inline">${typePillHtml(r.type)}</div>
    </div>
    ${flatMetricHtml('人数',r.user_count,'user')}
    ${flatMetricHtml('歌唱数',r.count,'song')}
  </div>`;
}
function renderRanking(){
  const cat = rankCatSel.value;
  const mode = rankModeSel.value;
  const top20 = buildRankingTop20(cat, mode);
  document.getElementById('rankCount').innerText = top20.length;
  const list = document.getElementById('rankList');
  if(top20.length===0){list.innerHTML='<div class="empty">ランキング対象データがありません</div>';return;}
  let html='';
  top20.forEach(r=>html+=buildRankCardHtml(r, mode));
  list.innerHTML = html;
}

// ===== Trend =====
const trendPeriodSel = document.getElementById('trendPeriod');
const trendSortSel = document.getElementById('trendSort');
const trendModeSel = document.getElementById('trendMode');
TREND_PERIODS.forEach(p=>{
  const o=document.createElement('option');o.value=String(p);o.innerText=`直近${p}日`;
  if(p===7) o.selected=true;
  trendPeriodSel.appendChild(o);
});
trendPeriodSel.addEventListener('change', renderTrend);
trendSortSel.addEventListener('change', renderTrend);
if(trendModeSel) trendModeSel.addEventListener('change', renderTrend);

function trendCurValue(it, mode){
  return mode==='user' ? (it.cur_user||0) : (it.cur_count||0);
}
function trendPrevValue(it, mode){
  if(mode==='user') return (it.prev_user!=null ? it.prev_user : 0);
  return (it.prev_count!=null ? it.prev_count : 0);
}
function trendDelta(it, mode){
  return trendCurValue(it, mode) - trendPrevValue(it, mode);
}
function growthBadgeHtml(it, mode){
  const cur = trendCurValue(it, mode);
  const prev = trendPrevValue(it, mode);
  const d = cur - prev;
  const unit = mode==='user' ? '人' : '曲';
  if(prev<=0 && cur>0) return `<span class="growth-badge new">NEW +${cur}${unit}</span>`;
  if(prev<=0 && cur<=0) return '';
  if(d===0) return `<span class="growth-badge flat">±0${unit}</span>`;
  if(d>0) return `<span class="growth-badge up">+${d}${unit}</span>`;
  return `<span class="growth-badge down">${d}${unit}</span>`;
}

function buildTrendItems(){
  const sort = trendSortSel.value;
  const mode = trendModeSel ? trendModeSel.value : 'count';
  const period = trendPeriodSel.value;
  const td = TREND_DATA[period] || {kpi:{surge_count:0,new_in:0,max_delta:0},items:[]};
  // Recompute delta per current mode for sorting / display
  const items = td.items.map(it=>{
    const d = trendDelta(it, mode);
    return Object.assign({}, it, {_delta_m: d, _cur_m: trendCurValue(it, mode), _is_new_m: (trendPrevValue(it,mode)<=0 && trendCurValue(it,mode)>0)});
  });
  if(sort==='surge') items.sort((a,b)=>b._delta_m-a._delta_m || b._cur_m-a._cur_m);
  else if(sort==='new') items.sort((a,b)=>(b._is_new_m?1:0)-(a._is_new_m?1:0) || b._delta_m-a._delta_m);
  else items.sort((a,b)=>b._delta_m-a._delta_m);
  // Recompute kpi for current mode
  const surge_count = items.filter(x=>x._delta_m>0).length;
  const new_in = items.filter(x=>x._is_new_m).length;
  const max_delta = items.reduce((m,x)=>Math.max(m, x._delta_m), 0);
  return {kpi:{surge_count, new_in, max_delta}, items, mode};
}

function buildTrendHtml(){
  const {kpi, items, mode} = buildTrendItems();
  const unit = mode==='user' ? '人' : '曲';
  let html = '';
  html += `<div class="trend-stats">
    <div class="trend-stat"><div class="ico up"><i class="fas fa-arrow-trend-up"></i></div><div class="lbl">急上昇</div><div class="val">${kpi.surge_count}<small>${unit}</small></div></div>
    <div class="trend-stat"><div class="ico new"><i class="far fa-star"></i></div><div class="lbl">新規ランクイン</div><div class="val">${kpi.new_in}<small>${unit}</small></div></div>
    <div class="trend-stat"><div class="ico max"><i class="fas fa-arrow-up"></i></div><div class="lbl">最大伸び</div><div class="val">+${kpi.max_delta}</div></div>
  </div>`;

  if(items.length>0){
    const p = items[0];
    const cur_user = p.cur_user||0;
    const cur_count = p.cur_count||0;
    html += `<div class="trend-pickup">
      <div class="trend-pickup-head"><i class="fas fa-fire"></i> 急上昇ピックアップ</div>
      <div class="trend-pickup-row">
        <div class="num-badge">1</div>
        <div class="pickup-info">
          <div class="pickup-anime">${escHtml(p.anime)}</div>
          <div class="pickup-sub">${escHtml(p.song)} / ${escHtml(p.artist)}</div>
          <div class="pickup-types-inline">${typePillHtml(p.type)}</div>
        </div>
        ${flatMetricHtml('人数',cur_user,'user')}
        ${flatMetricHtml('歌唱数',cur_count,'song')}
        ${growthBadgeHtml(p, mode)}
      </div>
    </div>`;
  }

  if(items.length>1){
    html += `<div class="notable-head"><i class="fas fa-arrow-trend-up"></i> 注目の上昇曲</div>`;
    items.slice(1, 10).forEach((it,idx)=>{
      const rank = idx+2;
      html += `<div class="notable-row">
        <div class="num-badge">${rank}</div>
        <div class="notable-info">
          <div class="notable-anime">${escHtml(it.anime)}</div>
          <div class="notable-artist">${escHtml(it.song)} / ${escHtml(it.artist)}</div>
          <div class="notable-types">${typePillHtml(it.type)}</div>
        </div>
        ${flatMetricHtml('人数',it.cur_user||0,'user')}
        ${flatMetricHtml('歌唱数',it.cur_count||0,'song')}
        ${growthBadgeHtml(it, mode)}
      </div>`;
    });
  }

  if(items.length===0) html += '<div class="empty">急上昇データがありません</div>';
  return html;
}
function renderTrend(){
  document.getElementById('trendBody').innerHTML = buildTrendHtml();
}

// ===== ASS tag extension =====
let assEntries = [];
let assRawMode = false;

function showAssMsg(msg, kind){
  const el = document.getElementById('assMsg');
  if(!el) return;
  el.innerText = msg || '';
  el.className = 'ass-msg' + (kind ? ' '+kind : '');
}
function assNormalizeText(text){
  return String(text||'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').trim();
}
function parseAssTagText(text){
  text = assNormalizeText(text);
  if(!text) return [];
  const heads = [];
  const re = /【([^】]+)】/g;
  let m;
  while((m = re.exec(text)) !== null){
    heads.push({title:m[1].trim(), start:m.index, bodyStart:re.lastIndex});
  }
  if(heads.length === 0){
    return [{title:'新規ASSタグ', row:'', char:text}];
  }
  const entries = [];
  heads.forEach((h, i)=>{
    const nextStart = i + 1 < heads.length ? heads[i+1].start : text.length;
    const body = text.slice(h.bodyStart, nextStart).trim();
    const entry = {title:h.title || ('項目'+(i+1)), row:'', char:''};
    const labels = [];
    const labelRe = /(?:^|\n)(行|文字)：/g;
    let lm;
    while((lm = labelRe.exec(body)) !== null){
      labels.push({kind:lm[1], labelStart:lm.index, contentStart:labelRe.lastIndex});
    }
    if(labels.length === 0){
      entry.char = body;
    }else{
      const prefix = body.slice(0, labels[0].labelStart).trim();
      if(prefix) entry.char = prefix;
      labels.forEach((lb, li)=>{
        const end = li + 1 < labels.length ? labels[li+1].labelStart : body.length;
        const val = body.slice(lb.contentStart, end).trim();
        if(lb.kind === '行') entry.row = entry.row ? entry.row + '\n' + val : val;
        else entry.char = entry.char ? entry.char + '\n' + val : val;
      });
    }
    entries.push(entry);
  });
  return entries;
}
function serializeAssTagEntries(entries){
  return entries.map((e, idx)=>{
    const title = (e.title || '').trim() || ('項目'+(idx+1));
    const row = assNormalizeText(e.row);
    const chr = assNormalizeText(e.char);
    let out = '【' + title + '】';
    if(row) out += '\n行：' + row;
    if(chr) out += '\n文字：' + chr;
    return out;
  }).join('\n\n');
}
function readAssEntriesFromCards(){
  return Array.from(document.querySelectorAll('.ass-card')).map(card=>({
    title: card.querySelector('.ass-title')?.value || '',
    row: card.querySelector('.ass-row')?.value || '',
    char: card.querySelector('.ass-char')?.value || ''
  }));
}
function assAutosize(el){
  if(!el) return;
  el.style.height = 'auto';
  el.style.height = Math.max(el.scrollHeight + 2, el.classList.contains('char') ? 112 : 70) + 'px';
}
function renderAssEditor(entries){
  assEntries = (entries || []).map(e=>({title:e.title||'', row:e.row||'', char:e.char||''}));
  const list = document.getElementById('assList');
  const count = document.getElementById('assCount');
  if(count) count.innerText = assEntries.length;
  if(!list) return;
  if(assEntries.length === 0){
    list.innerHTML = '<div class="empty">ASSタグ項目がありません</div>';
    return;
  }
  list.innerHTML = assEntries.map((e, idx)=>`<div class="ass-card" data-index="${idx}">
    <div class="ass-card-head">
      <div class="ass-card-no">${idx+1}</div>
      <input class="ass-input ass-title" type="text" value="${escAttr(e.title)}" placeholder="【】内の項目名">
      <button class="ass-delete" type="button" title="この項目を削除" aria-label="この項目を削除"><i class="fas fa-trash-alt"></i></button>
    </div>
    <div class="ass-fields">
      <div class="ass-field">
        <div class="ass-field-label"><i class="fas fa-grip-lines"></i> 行</div>
        <textarea class="ass-text ass-row" placeholder="行：に入れるASSタグ">${escHtml(e.row)}</textarea>
      </div>
      <div class="ass-field">
        <div class="ass-field-label"><i class="fas fa-font"></i> 文字</div>
        <textarea class="ass-text ass-char char" placeholder="文字：に入れるASSタグ">${escHtml(e.char)}</textarea>
      </div>
    </div>
  </div>`).join('');
  list.querySelectorAll('textarea').forEach(t=>{
    assAutosize(t);
    t.addEventListener('input',()=>assAutosize(t));
  });
  list.querySelectorAll('.ass-delete').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const card = btn.closest('.ass-card');
      const idx = Number(card.dataset.index);
      const cur = readAssEntriesFromCards();
      cur.splice(idx, 1);
      renderAssEditor(cur);
      showAssMsg('項目を削除しました。保存するまでGASには反映されません。', '');
    });
  });
}
function assCurrentText(){
  if(assRawMode){
    return document.getElementById('assRawText').value;
  }
  return serializeAssTagEntries(readAssEntriesFromCards());
}
function assSetRawMode(on){
  const panel = document.getElementById('assRawPanel');
  const list = document.getElementById('assList');
  const btn = document.getElementById('assRawToggle');
  assRawMode = !!on;
  if(assRawMode){
    document.getElementById('assRawText').value = serializeAssTagEntries(readAssEntriesFromCards());
    panel.classList.add('show');
    list.classList.add('ass-hidden');
    if(btn) btn.innerHTML = '<i class="fas fa-table"></i><span class="dl-lbl">項目編集</span>';
  }else{
    panel.classList.remove('show');
    list.classList.remove('ass-hidden');
    if(btn) btn.innerHTML = '<i class="fas fa-code"></i><span class="dl-lbl">テキスト編集</span>';
  }
}
function assApplyRawToCards(){
  const raw = document.getElementById('assRawText').value;
  const entries = parseAssTagText(raw);
  renderAssEditor(entries);
  assSetRawMode(false);
  showAssMsg('テキストを項目へ反映しました。保存するまでGASには反映されません。', 'ok');
}
async function assLoad(){
  showAssMsg('読み込み中...', '');
  let text = '';
  try{
    const res = await fetch(GAS_URL+'?filename='+encodeURIComponent(ASS_TAGS_FILENAME), {method:'GET'});
    if(res.ok) text = await res.text();
    if(!res.ok || !assNormalizeText(text) || text.slice(0,200).indexOf('Exception') >= 0){
      text = ASS_TAG_DEFAULT_TEXT;
      showAssMsg('保存ファイルが未作成または空のため、初期プリセットを表示しました。', '');
    }else{
      showAssMsg('読み込み完了 ('+parseAssTagText(text).length+' 項目)', 'ok');
    }
  }catch(e){
    text = ASS_TAG_DEFAULT_TEXT;
    showAssMsg('読み込みエラーのため初期プリセットを表示しました: '+e.message, 'error');
  }
  renderAssEditor(parseAssTagText(text));
  assSetRawMode(false);
}
async function assSave(){
  const content = assCurrentText();
  if(!assNormalizeText(content)){ showAssMsg('内容が空です', 'error'); return; }
  if(!confirm(ASS_TAGS_FILENAME+' を保存します。よろしいですか？')) return;
  showAssMsg('保存中...', '');
  try{
    const res = await fetch(GAS_URL, {
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},
      body: JSON.stringify({filename:ASS_TAGS_FILENAME, content:content})
    });
    if(!res.ok){ showAssMsg('保存失敗: HTTP '+res.status, 'error'); return; }
    if(assRawMode) renderAssEditor(parseAssTagText(content));
    assSetRawMode(false);
    showAssMsg('保存完了。次回読み込み時もこの内容が表示されます。', 'ok');
  }catch(e){
    showAssMsg('保存エラー: '+e.message, 'error');
  }
}
function assAddEntry(){
  const cur = assRawMode ? parseAssTagText(document.getElementById('assRawText').value) : readAssEntriesFromCards();
  cur.push({title:'新規ASSタグ', row:'', char:''});
  renderAssEditor(cur);
  assSetRawMode(false);
  showAssMsg('項目を追加しました。保存するまでGASには反映されません。', '');
}
function initAssEditor(){
  const loadBtn = document.getElementById('assLoad');
  if(!loadBtn) return;
  renderAssEditor(parseAssTagText(ASS_TAG_DEFAULT_TEXT));
  showAssMsg('初期プリセットを表示中です。保存済みデータを使う場合は「読み込み」を押してください。', '');
  loadBtn.addEventListener('click', assLoad);
  document.getElementById('assSave').addEventListener('click', assSave);
  document.getElementById('assAdd').addEventListener('click', assAddEntry);
  document.getElementById('assRawToggle').addEventListener('click', ()=>assSetRawMode(!assRawMode));
  document.getElementById('assApplyRaw').addEventListener('click', assApplyRawToCards);
}

// ===== Env =====
// CSV editor for cool_analysis.csv
function showCsvMsg(msg, kind){
  const el = document.getElementById('envCsvMsg');
  el.innerText = msg;
  el.className = 'env-csv-msg' + (kind ? ' '+kind : '');
}
async function envCsvLoad(){
  showCsvMsg('読み込み中...', '');
  try{
    const res = await fetch(GAS_URL+'?filename=cool_analysis.csv', {method:'GET'});
    if(!res.ok){ showCsvMsg('読み込み失敗: HTTP '+res.status, 'error'); return; }
    const text = await res.text();
    document.getElementById('envCsvText').value = text;
    showCsvMsg('読み込み完了 ('+text.length+' 文字)', 'ok');
  }catch(e){
    showCsvMsg('読み込みエラー: '+e.message, 'error');
  }
}
async function envCsvSave(){
  const content = document.getElementById('envCsvText').value;
  if(!content){ showCsvMsg('内容が空です', 'error'); return; }
  if(!confirm('cool_analysis.csv を保存します。よろしいですか？')) return;
  showCsvMsg('保存中...', '');
  try{
    const res = await fetch(GAS_URL, {
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},
      body: JSON.stringify({filename:'cool_analysis.csv', content:content})
    });
    if(!res.ok){ showCsvMsg('保存失敗: HTTP '+res.status, 'error'); return; }
    showCsvMsg('保存完了。次回更新で反映されます。', 'ok');
  }catch(e){
    showCsvMsg('保存エラー: '+e.message, 'error');
  }
}
document.getElementById('envCsvLoad').addEventListener('click', envCsvLoad);
document.getElementById('envCsvSave').addEventListener('click', envCsvSave);

// ===== HTML downloads =====
function buildExportHtml(title, contentHtml){
  const port = document.getElementById('exportPort').value || '11059';
  const linkType = document.getElementById('exportLinkType').value;
  const baseStyle = document.querySelector('style').innerHTML;
  return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>${escHtml(title)}</title>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Noto+Sans+JP:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>${baseStyle}</style>
</head><body>
<div class="app-header"><div class="brand">${escHtml(title)}</div><button class="theme-toggle" id="themeToggle" type="button" title="ダークモード切替" aria-label="ダークモード切替"><i class="fas fa-moon"></i></button></div>
<div class="update-line">${escHtml(UPDATE_TS)} 出力</div>
<div class="tab-content active" style="display:block">${contentHtml}</div>
<div class="modal-overlay" id="modalOverlay"><div class="modal"><div class="modal-head"><div class="ttl" id="modalTitle"></div><button class="modal-close" onclick="closeModal()"><i class="fas fa-times"></i></button></div><div class="modal-body" id="modalBody"></div></div></div>
<script>
const PORT='${port}';const LINKTYPE='${linkType}';
const HISTORY=${JSON.stringify(HISTORY)};
const IMAGE_MAP=${JSON.stringify(IMAGE_MAP)};
${commonExportScript()}
<\/script>
</body></html>`;
}
function downloadFile(filename, content){
  const blob = new Blob([content],{type:'text/html;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = filename; a.click();
}
function commonExportScript(){
  return `
function escHtml(s){return String(s||"").replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function escAttr(s){return String(s||"").replace(/'/g,"&#39;").replace(/"/g,'&quot;');}
function jsNormalize(s){if(!s)return"";s=String(s).normalize('NFKC');s=s.replace(/\\.[a-zA-Z0-9]{3,4}$/,'');s=s.replace(/[\\[\\(\\{【].*?[\\]\\)\\}】]/g,' ');s=s.replace(/(key|KEY)?\\s*[\\+\\-]\\s*[0-9]+/g,' ');s=s.replace(/原キー/g,' ');s=s.replace(/(キー)?変更[:：]?/g,' ');s=s.replace(/[~〜～\\-_=,.]/g,' ');s=s.replace(/\\s+/g,' ').trim();return s.toUpperCase();}
function findHistoryMatches(w,s){const sn=jsNormalize(s);const wn=jsNormalize(w);if(!sn&&!wn)return[];return HISTORY.filter(h=>{let so=false,wo=false;if(sn){if(/^[A-Z0-9 ]+$/.test(sn)){const re=new RegExp('(?:^|[^A-Z0-9])'+sn.replace(/[.*+?^\${}()|[\\]\\\\]/g,'\\\\$&')+'(?:[^A-Z0-9]|$)','i');so=re.test(h.sn);}else{so=h.sn.indexOf(sn)>=0;}}if(wn){wo=(h.sn.indexOf(wn)>=0)||(h.wn.indexOf(wn)>=0);}if(sn&&wn)return so&&wo;if(sn)return so;if(wn)return wo;return false;});}
function openSingersModal(w,s){const m=findHistoryMatches(w,s);const u={};m.forEach(x=>{if(!u[x.u])u[x.u]={c:0,d:x.d,r:x.rm};u[x.u].c++;if(x.d>u[x.u].d){u[x.u].d=x.d;u[x.u].r=x.rm;}});const us=Object.entries(u).sort((a,b)=>b[1].c-a[1].c);document.getElementById('modalTitle').innerHTML=escHtml(s)+'<small>'+escHtml(w)+' - '+m.length+'件 / '+us.length+'人</small>';let b='<div class="modal-summary"><i class="fas fa-users"></i> '+us.length+'人がこの曲を歌っています（合計'+m.length+'回）</div>';if(us.length===0)b+='<div class="empty">履歴なし</div>';else us.forEach(([uu,info])=>{b+='<div class="modal-row"><div class="top"><div class="user"><i class="fas fa-microphone"></i> '+escHtml(uu)+' <span style=\"color:var(--accent);font-size:11.5px;margin-left:4px\">×'+info.c+'</span></div><div class="date">'+escHtml(info.d)+'</div></div><div class="meta">最新: '+escHtml(info.r)+'</div></div>';});document.getElementById('modalBody').innerHTML=b;document.getElementById('modalOverlay').classList.add('active');}
function closeModal(){document.getElementById('modalOverlay').classList.remove('active');}
document.getElementById('modalOverlay').addEventListener('click',e=>{if(e.target.id==='modalOverlay')closeModal();});
function toggleCard(h){h.parentElement.classList.toggle('expanded');}
// 検索キーワードから ()/（）/～/〜/／/~// および括弧内の付加情報を取り除く。
// 例: "ようこそ実力至上主義の教室へ(4th Season) MONSTER"
//   → "ようこそ実力至上主義の教室へ MONSTER"
function cleanSearchKeyword(s){
  if(!s) return '';
  s = String(s);
  s = s.replace(/\\([^)]*\\)/g,' ').replace(/（[^）]*）/g,' ');
  s = s.replace(/[～〜~／\\/]/g,' ');
  s = s.replace(/\\s+/g,' ').trim();
  return s;
}
document.querySelectorAll('a.export-link').forEach(l=>{const h=l.getAttribute('href');if(h&&h.indexOf('#search:')===0){const w=h.split('#search:')[1];const path=LINKTYPE==='ykr'?'search_listerdb_filelist.php?anyword=':'search.php?searchword=';l.href='http://ykr.moe:'+PORT+'/'+path+encodeURIComponent(cleanSearchKeyword(decodeURIComponent(w)));l.target='_blank';l.rel='noopener';}});
// ----- Row-level clickable area (full item navigates to its export-link) -----
(function(){
  var rowSelectors=['.rank-row-flat','.notable-row','.trend-pickup-row','.song-row','.detail-row'];
  rowSelectors.forEach(function(sel){
    document.querySelectorAll(sel).forEach(function(row){
      var link=row.querySelector('a.export-link');
      if(!link||!link.href) return;
      row.style.cursor='pointer';
      row.addEventListener('click',function(e){
        if(e.target.closest('a, button')) return;
        window.open(link.href, link.target||'_blank', 'noopener');
      });
    });
  });
  // Setlist exported card: clicking the song line in the head should navigate
  // (the rest of the head still toggles expansion).
  document.querySelectorAll('.sl-card-head .sl-song a.export-link').forEach(function(a){
    a.addEventListener('click',function(e){e.stopPropagation();});
  });
})();
// ----- Dark mode toggle (persisted in localStorage) -----
(function(){
  var tg=document.getElementById('themeToggle');
  if(!tg) return;
  function apply(d){
    document.body.classList.toggle('dark', d);
    var ic=tg.querySelector('i');
    if(ic) ic.className=d?'fas fa-sun':'fas fa-moon';
  }
  var dark=false;
  try{ dark=localStorage.getItem('darkMode')==='1'; }catch(e){}
  apply(dark);
  tg.addEventListener('click',function(){
    dark=!dark;
    try{ localStorage.setItem('darkMode', dark?'1':'0'); }catch(e){}
    apply(dark);
  });
})();
`;
}

function buildCoolHtmlForCatExport(cat){
  const data = COOL_DATA[cat]||{works:[]};
  const works = data.works.slice().sort((a,b)=>b.total_count-a.total_count);
  let h = `<div style="font-size:14px;font-weight:700;color:var(--primary);padding-left:10px;border-left:3px solid var(--accent);margin:14px 0 8px">${escHtml(cat)} - ${works.length}作品</div>`;
  works.forEach((w,i)=>{
    const card = buildCoolCard(w, i+1, cat);
    h += card.replace(/(<div class="song-name">)([^<]+)(<span class="song-created-mark"[\s\S]*?<\/span>)?(<\/div>)/g, (_,p1,n,mark,p4)=>{
      return p1 + `<a class="export-link" href="#search:${encodeURIComponent(w.anime+' '+n)}">${n}</a>` + (mark||'') + p4;
    });
  });
  return h;
}
function downloadCoolHTML(scope){
  const cat = coolCatSel.value;
  let body = '<div style="max-width:var(--maxw);margin:0 auto;padding:0 16px 60px">';
  if(scope==='current') body += buildCoolHtmlForCatExport(cat);
  else CATS.forEach(c=> body += buildCoolHtmlForCatExport(c));
  body += '</div>';
  const ttl = scope==='current' ? `クール集計 - ${cat}` : 'クール集計（全クール）';
  const fname = scope==='current' ? `karaoke_cool_${cat}.html` : 'karaoke_cool_all.html';
  downloadFile(fname, buildExportHtml(ttl, body));
}

function buildRankHtmlForCatExport(cat, mode){
  const top20 = buildRankingTop20(cat, mode);
  const modeT = mode==='count' ? '歌唱数ランキング' : '歌唱人数ランキング';
  let h = `<div style="font-size:14px;font-weight:700;color:var(--primary);padding-left:10px;border-left:3px solid var(--accent);margin:14px 0 8px">${escHtml(cat)} ${modeT}</div>`;
  top20.forEach(r=>{
    const card = buildRankCardHtml(r, mode);
    h += card
      .replace(/(<div class="rank-anime">)([^<]+)(<\/div>)/, (_,p1,n,p3)=>p1+`<a class="export-link" href="#search:${encodeURIComponent(r.anime+' '+r.song)}">${n}</a>`+p3)
      .replace(/(<div class="rank-sub">)([^<]+)(<\/div>)/, (_,p1,n,p3)=>p1+`<a class="export-link" href="#search:${encodeURIComponent(r.anime+' '+r.song)}">${n}</a>`+p3);
  });
  return h;
}
function downloadRankingHTML(scope){
  const cat = rankCatSel.value;
  const mode = rankModeSel.value;
  let body = '<div style="max-width:var(--maxw);margin:0 auto;padding:0 16px 60px">';
  if(scope==='current') body += buildRankHtmlForCatExport(cat, mode);
  else CATS.forEach(c=> body += buildRankHtmlForCatExport(c, mode));
  body += '</div>';
  const ttl = scope==='current' ? `ランキング - ${cat}` : 'ランキング（全クール）';
  const fname = scope==='current' ? `karaoke_rank_${mode}_${cat}.html` : `karaoke_rank_${mode}_all.html`;
  downloadFile(fname, buildExportHtml(ttl, body));
}

function downloadTrendHTML(){
  const period = trendPeriodSel.value;
  let inner = buildTrendHtml();
  // Wrap song titles in pickup-sub and notable-artist with export-links so the
  // saved-time port + 検索リンク mapping can rewrite them on click.
  inner = inner.replace(
    /(<div class="pickup-anime">)([^<]+)(<\/div>)\s*(<div class="pickup-sub">)([^<]+)(<\/div>)/g,
    (_, p1, anime, p3, p4, sub, p6)=>{
      const key = encodeURIComponent(anime + ' ' + sub.split(' / ')[0]);
      return p1 + anime + p3 + p4 + `<a class="export-link" href="#search:${key}">${sub}</a>` + p6;
    }
  );
  inner = inner.replace(
    /(<div class="notable-anime">)([^<]+)(<\/div>)\s*(<div class="notable-artist">)([^<]+)(<\/div>)/g,
    (_, p1, anime, p3, p4, sub, p6)=>{
      const key = encodeURIComponent(anime + ' ' + sub.split(' / ')[0]);
      return p1 + anime + p3 + p4 + `<a class="export-link" href="#search:${key}">${sub}</a>` + p6;
    }
  );
  const body = `<div style="max-width:var(--maxw);margin:0 auto;padding:0 16px 60px"><div style="font-size:14px;font-weight:700;color:var(--primary);padding-left:10px;border-left:3px solid var(--accent);margin:14px 0 8px">推移 - ${TREND_TARGET_CAT} (直近${period}日)</div>${inner}</div>`;
  downloadFile(`karaoke_trend_${period}d.html`, buildExportHtml(`推移 - ${TREND_TARGET_CAT}`, body));
}
function downloadSetlistHTML(){
  const yearSel = document.getElementById('slDlYear');
  const yearVal = yearSel ? yearSel.value : 'all';
  // Filter SETLIST by selected year (or 'all')
  let arr = SETLIST.slice();
  if(yearVal && yearVal !== 'all'){
    arr = arr.filter(x=>{
      const y = (x.d || '').slice(0,4);
      return y === yearVal;
    });
  }
  // 並び順は compareSetlistEntries に集約 (第一=取得日 desc, 第二=順番 desc)
  arr.sort(compareSetlistEntries);

  let cards = '';
  arr.forEach((x, idx)=>{
    // バッジには 順番 (x.o) を表示する。同一日付内では 順番 降順で並ぶため、
    // 大きい番号が上に来る並びとバッジ表示が一致する。順番が空/数値でない場合は
    // 行位置にフォールバック。
    const fallbackNum = (idx + 1).toString();
    const dispNum = (x.o !== undefined && x.o !== null && String(x.o).trim() !== '')
      ? String(x.o).trim()
      : fallbackNum;
    const metaPieces = [];
    if(x.wk) metaPieces.push(escHtml(x.wk));
    if(x.ar) metaPieces.push(escHtml(x.ar));
    const metaLine = metaPieces.join('<span class="sep">|</span>');
    const songIcon = x.sg && /singing|歌|うた/i.test(x.sg) ? '<i class="fas fa-microphone song-mic"></i>' : '';
    const searchKey = encodeURIComponent((x.wk ? x.wk + ' ' : '') + x.sg);
    const songLinkOpen = `<a class="export-link" href="#search:${searchKey}" onclick="event.stopPropagation()">`;
    const songLinkClose = `</a>`;
    cards += `<div class="card">
      <div class="sl-card-head" onclick="toggleCard(this)">
        <div class="num-badge">${escHtml(dispNum)}</div>
        <div class="sl-body">
          <div class="sl-top">
            ${roomTagHtml(x.rm)}
            <span class="sl-date">${escHtml(x.d)}</span>
          </div>
          <div class="sl-song">${songIcon}<span style="overflow:hidden;text-overflow:ellipsis">${songLinkOpen}${escHtml(x.sg)}${songLinkClose}</span></div>
          ${metaLine?`<div class="sl-meta">${metaLine}</div>`:''}
        </div>
        <i class="fas fa-chevron-down card-chev"></i>
      </div>
      <div class="card-detail">
        <div class="detail-table">
          <div class="detail-row"><div class="lbl"><i class="fas fa-music"></i>曲名</div><div class="val">${songLinkOpen}${escHtml(x.sg)}${songLinkClose}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-film"></i>作品名</div><div class="val">${escHtml(x.wk)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-user"></i>歌手名</div><div class="val">${escHtml(x.ar)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-microphone"></i>歌った人</div><div class="val">${escHtml(x.u)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-door-open"></i>部屋</div><div class="val">${escHtml(x.rm)}</div></div>
          <div class="detail-row"><div class="lbl"><i class="fas fa-hashtag"></i>順番</div><div class="val">${escHtml(x.o)}</div></div>
        </div>
        <button class="confirm-btn" onclick="event.stopPropagation();openSingersModal('${escAttr(x.wk)}','${escAttr(x.sg)}')">
          <i class="fas fa-users"></i> この曲を歌った人を確認 <i class="fas fa-chevron-right"></i>
        </button>
      </div>
    </div>`;
  });
  if(arr.length===0) cards = '<div class="empty">該当データがありません</div>';

  const periodLabel = (yearVal && yearVal !== 'all') ? (yearVal + '年') : '全件';
  const titleSuffix = (yearVal && yearVal !== 'all') ? ` (${yearVal}年)` : ' (全件)';
  const body = `<div style="max-width:var(--maxw);margin:0 auto;padding:0 16px 60px"><div style="font-size:14px;font-weight:700;color:var(--primary);padding-left:10px;border-left:3px solid var(--accent);margin:14px 0 8px">セットリスト${titleSuffix} - ${arr.length}件</div>${cards}</div>`;
  const fname = (yearVal && yearVal !== 'all') ? `karaoke_setlist_${yearVal}.html` : 'karaoke_setlist_all.html';
  downloadFile(fname, buildExportHtml('セットリスト' + titleSuffix, body));
}

// Populate setlist HTML download year selector
(function(){
  const sel = document.getElementById('slDlYear');
  if(!sel) return;
  const years = new Set();
  HISTORY.forEach(h=>{
    const y = (h.d || '').slice(0,4);
    if(/^\d{4}$/.test(y)) years.add(y);
  });
  const sorted = Array.from(years).sort((a,b)=>b.localeCompare(a));
  sel.innerHTML = '';
  const optAll = document.createElement('option');
  optAll.value = 'all'; optAll.innerText = '全件';
  sel.appendChild(optAll);
  sorted.forEach(y=>{
    const o = document.createElement('option');
    o.value = y; o.innerText = y + '年';
    sel.appendChild(o);
  });
  if(sorted.length>0) sel.value = sorted[0];
})();

renderSetlist();
renderCool();
renderRanking();
renderTrend();
initAssEditor();
</script>
</body>
</html>
"""

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
    .replace("__ASS_TAGS_FILENAME__", ASS_TAGS_FILENAME)
    .replace("__ASS_TAG_DEFAULT_JSON__", ass_tag_default_json)
    .replace("__UPDATE_TS__", current_datetime_str)
    .replace("__UPDATE__", current_datetime_str)
    .replace("__TREND_CAT__", TREND_TARGET_CATEGORY)
)

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
print("HTML生成完了: index.html")

# ==========================================
# 外部閲覧用ページ viewer.html を生成
# ==========================================
# index.html とは独立した、クール集計/ランキング/推移/マイリスト のみのページ。
# 機材係 (部屋) 選択モーダル、検索ボタン (各楽曲・別タブ無し)、
# マイリスト登録ボタン (歌唱数の上)、マイリストタブ (KaraokeDB機能を移行) を備える。
# テンプレートは viewer_template/ 配下に配置。
import re as _re_v

_viewer_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "viewer_template")
try:
    with open(os.path.join(_viewer_dir, "extra.css"), "r", encoding="utf-8") as f:
        _v_css_extra = f.read()
    with open(os.path.join(_viewer_dir, "body.html"), "r", encoding="utf-8") as f:
        _v_body = f.read()
    with open(os.path.join(_viewer_dir, "script.js"), "r", encoding="utf-8") as f:
        _v_js = f.read()

    # Base CSS (reuse the same <style> block we just embedded into index.html)
    _m = _re_v.search(r"<style>([\s\S]*?)</style>", html_content)
    _v_base_css = _m.group(1) if _m else ""

    # 仕様で指定された機材係(部屋)マップ。viewer 専用として固定で持つ。
    _viewer_rooms = {
        11000: "ゆーふうりん部屋", 11001: "ゆーふうりん部屋", 11002: "ゆーふうりん部屋",
        11003: "ゆーふうりん部屋", 11004: "ゆーふうりん部屋", 11005: "ゆーふうりん部屋",
        11006: "ゆーふうりん部屋", 11007: "ゆーふうりん部屋", 11008: "ゆーふうりん部屋",
        11009: "ゆーふうりん部屋",
        11021: "成田部屋", 11022: "成田部屋", 11028: "タマ部屋",
        11058: "すみた部屋", 11059: "つぼはち部屋", 11060: "れん部屋", 11061: "百合川部屋",
        11063: "なぎ部屋", 11064: "naoo部屋", 11066: "芝ちゃん部屋", 11067: "crom部屋",
        11068: "けんしん部屋", 11069: "けんちぃ部屋",
        11070: "黒河部屋", 11071: "黒河部屋", 11073: "コウ艦長部屋", 11074: "tukinowa部屋",
        11077: "v3部屋", 11078: "のんでるん部屋", 11079: "まどか部屋",
        11082: "ヤマテル部屋", 11084: "タカヒロ部屋", 11085: "タカヒロ部屋",
        11086: "タカヒロ部屋", 11087: "MiO部屋", 11088: "ほっしー部屋",
        11101: "えみち部屋", 11102: "るえ部屋", 11103: "ながし部屋", 11104: "MrN部屋",
        11105: "ヤマテル部屋", 11106: "遠見塚部屋", 11107: "ブルーベリー部屋",
        11110: "加古部屋",
    }
    _viewer_rooms_json = json.dumps({str(k): v for k, v in _viewer_rooms.items()}, ensure_ascii=False)

    _v_js_rendered = (_v_js
        .replace("__COOL_DATA_LINE__", "const COOL_DATA = " + cool_json + ";")
        .replace("__RANK_DATA_LINE__", "const RANK_DATA = " + ranking_json + ";")
        .replace("__TREND_DATA_LINE__", "const TREND_DATA = " + trend_json + ";")
        .replace("__HISTORY_DATA_LINE__", "const HISTORY = " + history_json + ";")
        .replace("__ROOMS_JSON__", _viewer_rooms_json)
        .replace("__CATS_JSON__", categories_json)
        .replace("__TREND_PERIODS_JSON__", trend_periods_json)
        .replace("__TREND_TARGET_JSON__", trend_target_json)
        .replace("__UPDATE_TS__", current_datetime_str)
    )
    _v_body_rendered = _v_body.replace("__UPDATE__", current_datetime_str)

    _viewer_html = (
        '<!DOCTYPE html>\n'
        '<html lang="ja">\n'
        '<head>\n'
        '<meta charset="UTF-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        '<title>Karaoke Viewer</title>\n'
        '<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">\n'
        '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Noto+Sans+JP:wght@400;500;600;700;800&display=swap" rel="stylesheet">\n'
        '<style>\n'
        + _v_base_css + '\n/* === viewer extra === */\n' + _v_css_extra +
        '\n</style>\n'
        '</head>\n'
        + _v_body_rendered + '\n'
        '<script>\n' + _v_js_rendered + '\n</script>\n'
        '</body>\n'
        '</html>\n'
    )
    with open("viewer.html", "w", encoding="utf-8") as f:
        f.write(_viewer_html)
    print("HTML生成完了: viewer.html")
except FileNotFoundError as e:
    print("[viewer] テンプレート未検出のためスキップ: {}".format(e))
except Exception as e:
    print("[viewer] 生成失敗: {}".format(e))
