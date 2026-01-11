import pandas as pd
import requests
import datetime
import re

# --- ① 設定: ポート番号と部屋主の名前の対応表 ---
room_map = {
    11021: "成田部屋",
    11028: "タマ部屋",
    11058: "すみた部屋",
    11059: "つぼはち部屋",
    11063: "なぎ部屋",
    11064: "naoo部屋",
    11066: "芝ちゃん部屋",
    11068: "けんしん部屋",
    11069: "けんちぃ部屋",
    11070: "黒河部屋",
    11074: "tukinowa部屋",
    11077: "v3部屋",
    11078: "のんでるん部屋",
    11079: "まどか部屋",
    11084: "タカヒロ部屋",
    11088: "ほっしー部屋",
    11101: "えみち部屋",
    11103: "ながし部屋",
    11106: "冨塚部屋"
}

# 取得対象のポート（対応表のキーをそのまま使う）
target_ports = list(room_map.keys())

all_data_frames = []

for port in target_ports:
    url = f"http://Ykr.moe:{port}/simplelist.php"
    try:
        # データ取得
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        # HTMLの表を読み込み
        dfs = pd.read_html(response.content)
        if dfs:
            df = dfs[0]
            
            # ③ コメントが無い場合はNaNではなく空欄にする
            df = df.fillna("")
            
            # ④ 項目名のsourceを「部屋主」に変更し、名前を入れる
            # まず列を追加
            df['部屋主'] = room_map[port]
            
            # DataFrameをリストに追加
            all_data_frames.append(df)
            print(f"Port {port} ({room_map[port]}): OK")
            
    except Exception as e:
        print(f"Port {port}: Error - {e}")

# HTML生成の準備
html_content = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Karaoke setlist all</title> <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body { font-family: "Helvetica Neue", Arial, "Hiragino Kaku Gothic ProN", "Hiragino Sans", sans-serif; padding: 20px; color: #333; }
        h1 { margin-bottom: 10px; }
        
        /* ⑦ 検索ボックスとボタンのデザイン */
        .search-container { margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
        .search-box {
            width: 100%; max-width: 400px; padding: 10px; font-size: 16px;
            border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;
        }
        .btn {
            padding: 10px 20px; font-size: 16px; cursor: pointer;
            background-color: #007bff; color: white; border: none; border-radius: 4px;
            text-decoration: none; display: inline-block;
        }
        .btn:hover { background-color: #0056b3; }
        .btn-update { background-color: #28a745; margin-left: 10px; font-size: 14px; }
        .btn-update:hover { background-color: #218838; }

        /* テーブルのデザイン */
        table { border-collapse: collapse; width: 100%; font-size: 14px; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background-color: #f2f2f2; position: sticky; top: 0; z-index: 1; cursor: pointer; }
        th:hover { background-color: #e2e2e2; }
        tr:nth-child(even) { background-color: #fff; }
        
        /* ⑤ 今日の日付の行の色分け */
        tr.today-row { background-color: #e3f2fd !important; } /* 薄い青 */
        tr.today-row:nth-child(even) { background-color: #bbdefb !important; } /* 少し濃い青 */

        .update-time { color: #666; font-size: 0.9em; margin-bottom: 5px; }
        .hidden { display: none; }
        
        /* レスポンシブ対応 */
        @media (max-width: 600px) {
            table { font-size: 12px; }
            td, th { padding: 6px; }
            .btn { width: 100%; margin-top: 5px; margin-left: 0; }
        }
    </style>
</head>
<body>
    <h1>Karaoke setlist all</h1>
"""

# 更新時刻の取得（日本時間）
now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
today_str = now.strftime("%Y/%m/%d") # 今日の日付文字列（例: 2026/01/11）
html_content += f'<div class="update-time">最終集計: {now.strftime("%Y/%m/%d %H:%M:%S")}</div>'

# ⑥ 手動更新ボタン（GitHub Actionsへのリンク）
# ※セキュリティ上、静的ページから直接スクリプトを動かすことはできないため、実行ページへのリンクとします
repo_url = "https://github.com/hachi515/karaoke_setlist/actions/workflows/update.yml" 
html_content += f'''
    <div class="search-container">
        <input type="text" id="searchInput" class="search-box" placeholder="キーワードを入力 (スペースで複数検索)...">
        <button onclick="filterTable()" class="btn"><i class="fas fa-search"></i> 検索</button>
        <button onclick="resetFilter()" class="btn" style="background-color: #6c757d;">リセット</button>
        <a href="{repo_url}" target="_blank" class="btn btn-update"><i class="fas fa-sync-alt"></i> 更新画面へ</a>
        <p style="font-size:0.8em; color:#666; margin-top:5px;">※「更新画面へ」→「Run workflow」で手動更新できます</p>
    </div>
'''

if all_data_frames:
    # データを結合
    final_df = pd.concat(all_data_frames, ignore_index=True)

    # ④ 順番で並び替え（古い順）
    # '順番'カラムを数値として扱うために変換（エラー回避）
    if '順番' in final_df.columns:
        final_df['順番'] = pd.to_numeric(final_df['順番'], errors='coerce')
        final_df = final_df.sort_values(by=['順番'], ascending=True)

    # ④ 部屋主カラムを左端に持ってくるためにカラム順序を入れ替え
    cols = list(final_df.columns)
    if '部屋主' in cols:
        cols.insert(0, cols.pop(cols.index('部屋主')))
        final_df = final_df[cols]

    # HTMLテーブルの手動構築（色分けクラスを付与するため）
    html_content += '<table id="setlistTable">'
    
    # ヘッダー
    html_content += '<thead><tr>'
    for col in final_df.columns:
        html_content += f'<th onclick="sortTable({list(final_df.columns).index(col)})">{col} <i class="fas fa-sort"></i></th>'
    html_content += '</tr></thead>'
    
    # ボディ
    html_content += '<tbody>'
    for _, row in final_df.iterrows():
        # ⑤ 日付判定（行のデータ全体を見て、今日の日付が含まれていればクラスを付与）
        row_str = " ".join(row.astype(str))
        row_class = "today-row" if today_str in row_str else ""
        
        html_content += f'<tr class="{row_class}">'
        for val in row:
            html_content += f'<td>{val}</td>'
        html_content += '</tr>'
    html_content += '</tbody></table>'

else:
    html_content += "<p>データの取得に失敗しました。</p>"

# ⑦ 検索・ソート用JavaScript
html_content += """
<script>
    // 検索フィルター機能（全角・半角スペース対応）
    function filterTable() {
        var input = document.getElementById("searchInput");
        var filter = input.value.toUpperCase();
        
        // 全角スペースを半角に変換し、スペースで区切ってキーワードリストを作成
        var keywords = filter.replace(/　/g, " ").split(" ").filter(k => k.length > 0);
        
        var table = document.getElementById("setlistTable");
        var tr = table.getElementsByTagName("tr");

        for (var i = 1; i < tr.length; i++) {
            var td = tr[i].getElementsByTagName("td");
            var rowText = "";
            for (var j = 0; j < td.length; j++) {
                rowText += td[j].textContent || td[j].innerText;
            }
            rowText = rowText.toUpperCase();
            
            // すべてのキーワードが含まれているかチェック (AND検索)
            var isMatch = true;
            for (var k = 0; k < keywords.length; k++) {
                if (rowText.indexOf(keywords[k]) === -1) {
                    isMatch = false;
                    break;
                }
            }
            
            if (isMatch || keywords.length === 0) {
                tr[i].style.display = "";
            } else {
                tr[i].style.display = "none";
            }
        }
    }

    // Enterキーで検索実行
    document.getElementById("searchInput").addEventListener("keyup", function(event) {
        if (event.key === "Enter") {
            filterTable();
        }
    });

    // リセット機能
    function resetFilter() {
        document.getElementById("searchInput").value = "";
        filterTable();
    }

    // 簡易ソート機能（ヘッダークリックで動作）
    function sortTable(n) {
        var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
        table = document.getElementById("setlistTable");
        switching = true;
        dir = "asc"; 
        while (switching) {
            switching = false;
            rows = table.rows;
            for (i = 1; i < (rows.length - 1); i++) {
                shouldSwitch = false;
                x = rows[i].getElementsByTagName("TD")[n];
                y = rows[i + 1].getElementsByTagName("TD")[n];
                
                var xContent = x.innerHTML.toLowerCase();
                var yContent = y.innerHTML.toLowerCase();
                
                // 数値の場合は数値として比較
                if (!isNaN(xContent) && !isNaN(yContent) && xContent !== "" && yContent !== "") {
                     if (dir == "asc") {
                        if (Number(xContent) > Number(yContent)) { shouldSwitch = true; break; }
                    } else if (dir == "desc") {
                        if (Number(xContent) < Number(yContent)) { shouldSwitch = true; break; }
                    }
                } else {
                    if (dir == "asc") {
                        if (xContent > yContent) { shouldSwitch = true; break; }
                    } else if (dir == "desc") {
                        if (xContent < yContent) { shouldSwitch = true; break; }
                    }
                }
            }
            if (shouldSwitch) {
                rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                switching = true;
                switchcount ++;      
            } else {
                if (switchcount == 0 && dir == "asc") {
                    dir = "desc";
                    switching = true;
                }
            }
        }
    }
</script>
</body>
</html>
"""

# index.html として保存
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)
