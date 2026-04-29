"""GAS からマイリストを取得し、リポジトリ内に CSV / JSON としてバックアップするスクリプト。

定期 GitHub Actions ワークフローから実行する。コミットは呼び出し側で行う。
- 出力先: リポジトリルートの `mylist_backup.csv` (UTF-8 BOM 付き) と `mylist_backup.json`。
- GAS URL は viewer_template/script.js に書かれている公開エンドポイントを利用する。

意図しない上書き防止のため、応答が配列でない / 0 件の場合は既存ファイルを保持して終了する。
"""

from __future__ import annotations

import csv
import io
import json
import os
import re
import sys
from urllib.request import Request, urlopen

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT_JS = os.path.join(REPO_ROOT, "viewer_template", "script.js")
OUT_CSV = os.path.join(REPO_ROOT, "mylist_backup.csv")
OUT_JSON = os.path.join(REPO_ROOT, "mylist_backup.json")

CSV_HEADER = ["id", "歌唱者", "作品名", "歌手", "曲名", "パート分け", "練習中", "要復習", "非公開"]


def _extract_gas_url() -> str:
    with open(SCRIPT_JS, "r", encoding="utf-8") as fh:
        text = fh.read()
    m = re.search(r'const\s+MYLIST_GAS_URL\s*=\s*"([^"]+)"', text)
    if not m:
        raise RuntimeError("MYLIST_GAS_URL を viewer_template/script.js から検出できません。")
    return m.group(1)


def _fetch(url: str) -> list:
    req = Request(url, headers={"User-Agent": "karaoke-mylist-backup/1.0"})
    with urlopen(req, timeout=60) as resp:
        body = resp.read().decode("utf-8")
    data = json.loads(body)
    if not isinstance(data, list):
        raise RuntimeError("GAS の応答が配列ではありません。")
    return data


def _to_csv(rows: list) -> str:
    buf = io.StringIO()
    buf.write("\ufeff")  # BOM for Excel compatibility
    w = csv.writer(buf, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
    w.writerow(CSV_HEADER)
    for r in rows:
        if not isinstance(r, dict):
            continue
        # 非公開設定を含む完全バックアップ (バックアップなので非公開も保存)
        w.writerow([
            r.get("id", ""),
            r.get("singer", "") or "",
            r.get("work", "") or "",
            r.get("artist", "") or "",
            r.get("song", "") or "",
            "1" if r.get("isPartDivision") else "",
            "1" if r.get("isPracticing") else "",
            "1" if r.get("isReviewNeeded") else "",
            "1" if r.get("isPrivate") else "",
        ])
    return buf.getvalue()


def main() -> int:
    url = _extract_gas_url()
    print(f"[mylist-backup] GAS URL: {url[:60]}...")
    try:
        rows = _fetch(url)
    except Exception as exc:  # noqa: BLE001
        print(f"[mylist-backup] 取得失敗: {exc}", file=sys.stderr)
        return 1

    print(f"[mylist-backup] 取得件数: {len(rows)}")
    # Safety: 0 件の応答で既存ファイルを上書きしない (誤消去防止)
    if len(rows) == 0 and os.path.exists(OUT_JSON):
        print("[mylist-backup] 0件のため既存ファイルを保持して終了します。", file=sys.stderr)
        return 0

    # JSON
    with open(OUT_JSON, "w", encoding="utf-8") as fh:
        json.dump(rows, fh, ensure_ascii=False, indent=2)

    # CSV
    csv_text = _to_csv(rows)
    with open(OUT_CSV, "w", encoding="utf-8", newline="") as fh:
        fh.write(csv_text)

    print(f"[mylist-backup] 書き出し完了: {OUT_CSV}, {OUT_JSON}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
