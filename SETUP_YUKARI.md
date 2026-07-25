# 導入手順

1. この変更を `main` ブランチへマージします。
2. GitHub の **Settings → Pages** で `Deploy from a branch`、`main / (root)` を選びます。
3. **Actions → Check Yukari ports → Run workflow** を1回実行します。
4. 公開URLの `/yukari_ports.html` を開きます。

## 仕組み

- GitHub Actions が `ykr.moe:11000-11130` へTCP接続し、`yukari-status.json` を更新します。
- GitHub Pages側は同一オリジンのHTTPS JSONだけを読むため、混在コンテンツで遮断されません。
- 一覧は「接続中」だけが初期表示され、30秒ごとにJSONを再取得します。
- 接続先リンクは明示的に `http://` のままです。接続先サーバー自体が停止中、携帯回線から遮断、またはHTTPをHTTPSへ誤変換した場合は、リンク先は開けません。

## 運用上の注意

- 定期実行は、このワークフローがデフォルトブランチへマージされた後に有効になります。
- リポジトリ設定でGitHub Actionsの書き込み権限が制限されている場合は、**Settings → Actions → General → Workflow permissions** を確認してください。
