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

今回のクラウド環境からGitHub APIへのアクセスは403で、過去のActions実行ログは取得できませんでした。
5時間空いた直接原因は未確定です。再発時は `Update List` の実行履歴で次を分けて確認してください。

1. 実行自体がない：scheduleの遅延・間引き、無効化、デフォルトブランチ上の設定を確認。
2. 実行が失敗：GAS/ゆかりの応答、履歴取得エラー、push競合、Pagesデプロイのログを確認。
3. 実行成功でも公開HTMLが古い：Pagesが `gh-pages / (root)` を配信しているか確認。
   このワークフローは `gh-pages` へデプロイします。

## 1〜2時間以内の更新を重視する場合

GitHub内のscheduleだけに依存せず、外部のスケジューラーから毎時23分などに既存の
`workflow_dispatch` を呼ぶ構成を推奨します。登録済みのcronサーバーやクラウドスケジューラーを使用します。

```sh
# 外部実行環境で、対象リポジトリにActions: write権限の認証を安全に設定してから実行。
gh workflow run update.yml --repo hachi515/karaoke_setlist --ref main
```

HTTPで呼ぶ場合は `POST https://api.github.com/repos/hachi515/karaoke_setlist/actions/workflows/update.yml/dispatches`、
本文は `{"ref":"main"}` です。認証は対象リポジトリに絞ったGitHub Appまたはfine-grained tokenを使い、
トークンをリポジトリやチャットへ貼らず、外部サービスのシークレット管理に保存します。

外部起動でもActionsのランナー待ち・外部データ障害は残ります。時間上限を厳密に満たす必要がある場合は、
外部実行環境で生成・配信まで行い、最終成功時刻が2時間以上古い場合の監視を併用してください。
この変更では外部サービスの契約・認証設定や本番ジョブの登録は行っていません。
