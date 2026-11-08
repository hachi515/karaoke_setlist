# HTMLの更新と定期実行

## 2026年秋アニメ

`cool_analysis.csv` は添付CSVのスナップショットです。2026年秋の150曲を含む5クールを保持しています。
`update_list.py` はGASから同名CSVを読み、取得失敗・空応答時にはこのローカルファイルへフォールバックします。
今後CSVを更新する場合は、GAS上のファイルとこのスナップショットを更新してください。

- クール集計・ランキングは秋を初期表示にし、夏・春・冬・2025年秋も選択できます。
- 歌唱履歴の集計期間は従来の開始日を保ち、2026/01/01から2026/12/31までへ延長しました。
  秋の曲を放送前に歌った履歴も、この期間内なら含みます。
- 推移の対象は2026年秋です。直近3・7・14・30日の比較を継続します。
- 複数クールで継続する同一曲は各クールへ残し、重複除去はクール内だけで行います。
- 履歴取得が途中で失敗した場合は生成を中止し、履歴・既存HTMLの欠落を防ぎます。

Python 3.12を使用します。外部サービスへ接続しない整合性テストは次のとおりです。

```sh
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
```

実運用の `python update_list.py` は履歴CSV・HTMLを書き換え、GASへ履歴をバックアップする処理も含みます。
開発中の見た目確認は `python -m http.server 8000 --bind 127.0.0.1` で既存のHTMLを配信できます。
クラウド環境の `/workspace/.karaoke-env/build_offline.py` は、現在のHTML内の履歴スナップショットと
ローカルCSVを使い、外部通信・GAS書き込みなしで別ディレクトリへHTMLを生成します。
オフライン生成物は実サービスからの最新取得を証明するものではありません。

## GitHub Actionsで改善した点

`Update List` は毎時17分・47分（UTC。日本時間も分は同じ）に起動を要求します。
混雑しやすい正時を避け、30分ごとの機会を用意することで、1〜2時間に1回の更新を目指します。
従来の10分ごとから頻度を下げるだけで遅延が直るわけではなく、**最大更新間隔を保証する設定ではありません**。

- 同じ更新ジョブの重複実行を防ぎ、実行中のジョブはキャンセルしません。
- Pythonと主要依存関係のバージョンを固定し、生成前に整合性テストを実行します。
- status・mylistの別ジョブがmainへ先にpushした場合は、rebaseしてpushを最大3回試します。
- 15分の上限でハングしたジョブを検出します。
- ソース・テンプレート・CSVをmainへpushした時も起動します。GASのCSVだけを変更した場合は
  次の定期実行、またはActionsの `Update List` → `Run workflow` で反映します。

GitHub公式にも、混雑時にはscheduleが遅延し、キュー中のジョブが破棄される場合があると記載されています。
また、公開リポジトリで60日間活動がないとscheduleは自動無効化されます。
[公式ドキュメント](https://docs.github.com/en/actions/reference/workflows-and-actions/events-that-trigger-workflows#schedule)

2026/11/08の再調査では、GitHubのWeb画面でUpdate Listが `active`、生成・Pagesデプロイの成功を確認しました。
一方、定期起動の実績には約4時間の空白がありました。起動しなかった予定時刻の内部理由はGitHub側にしかなく、
cronの変更だけで1〜2時間間隔へ復旧したとは判断できません。手動実行の成功と定期起動の復旧は区別します。

再発時は `Update List` の実行履歴で次を分けて確認してください。

1. 実行自体がない：scheduleの遅延・間引き、無効化、デフォルトブランチ上の設定を確認。
2. 実行が失敗：GAS/ゆかりの応答、履歴取得エラー、push競合、Pagesデプロイのログを確認。
3. 実行成功でも公開HTMLが古い：Pagesが `gh-pages / (root)` を配信しているか確認。
   このワークフローは `gh-pages` へデプロイします。

## GitHubの定期起動に依存しない毎時更新（GAS）

[毎時起動スクリプト](.github/gas/hourly_update.gs) を用意しました。
既存のCSV用GASやWebアプリを変更せず、定期起動専用のApps Scriptプロジェクトへ追加できます。
GitHubの `schedule` は予備として残し、GASの時間主導型トリガーから毎時 `workflow_dispatch` を呼びます。
GitHub側のcronが開始されなくても、この起動経路は独立して動きます。

### 初回のみ行う設定

1. GitHubの **Settings → Developer settings → Personal access tokens → Fine-grained tokens** で、
   `hachi515/karaoke_setlist` だけを対象とするトークンを作成します。
   **Repository permissions → Actions: Read and write** が必要です。
   トークンはソース・チャットへ貼らず、次のGAS設定へ直接保存します。有効期限後は更新が必要です。
2. [Google Apps Script](https://script.google.com/home) で定期起動専用の新しいプロジェクトを作り、
   `.github/gas/hourly_update.gs` の内容を貼り付けて保存します。Webアプリとしてのデプロイは不要です。
3. GASの **プロジェクトの設定 → スクリプト プロパティ** に、名前 `GITHUB_ACTIONS_TOKEN`、
   値に作成したトークンを登録します。
4. GASの実行関数で **installKaraokeHourlyTrigger** を選んで一度実行し、
   自分のアカウントによる外部通信・トリガー作成を承認します。
   接続を確認して毎時トリガーを作り、最初の更新要求をその場で送ります。
5. GitHubの **Actions → Update List** に **Update List · gas-hourly** が現れ、
   生成・Pages公開まで成功することを確認します。さらに次の時間主導型実行も成功すれば毎時起動を確認できます。

### 動作と確認方法

- GASからの要求は `inputs.trigger_source=gas-hourly` を指定します。
  GitHubの実行タイトルで `schedule`・`manual`・`gas-hourly` を区別できます。
  イベントの種別自体は `workflow_dispatch`（画面によっては「Manually run」）になります。
- スクリプトプロパティ `KARAOKE_LAST_DISPATCH_AT` は最後にGitHubが起動要求を受理した時刻（UTC）です。
  生成・公開の成功時刻ではないので、GASの **実行数** とGitHub Actionsの結果も確認してください。
- インストール関数を再実行しても、この処理のトリガーだけを置き換えます。既存のCSV処理などは削除しません。
  新規作成に失敗した場合は既存トリガーを保持します。
- 認証失敗はエラーにし、一時的な5xx・短い429制限だけを最大3回試します。
  トークンはログへ出力しません。複数の起動処理が重なった場合は重複送信を抑制します。
- 停止するときは **removeKaraokeHourlyTrigger** を実行します。
- コードのサービス境界・認証失敗・再試行・トリガー保全は、次のテストで検証できます。

```sh
node --test tests/test_gas_scheduler.cjs
```

GASの権限承認・トークン登録・時間主導型トリガーは、GitHubへのpushだけでは設定されません。
このクラウドからGoogleアカウントの編集権限やトークン値は取得できないため、上記の初回設定が必要です。
設定前の段階では「毎時自動更新が復旧した」とは扱いません。
GASの時間主導型トリガーやActionsのランナー待ちにも遅延はあり、厳密な時間保証ではありません。

既存の外部cronサーバーがある場合は、同じ起動口を次のコマンドから利用できます。

```sh
# 外部実行環境で対象リポジトリにActions: write権限の認証を安全に設定してから実行。
gh workflow run update.yml --repo hachi515/karaoke_setlist --ref main -f trigger_source=external
```
