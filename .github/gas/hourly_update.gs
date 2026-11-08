/**
 * GitHub scheduleとは独立した、Update Listの毎時起動。
 * スクリプトプロパティ GITHUB_ACTIONS_TOKEN に対象リポジトリだけの
 * fine-grained PAT (Actions: write) を設定する。値をソースへ記入しない。
 */
const KARAOKE_DISPATCH_HANDLER = 'dispatchKaraokeUpdate';
const KARAOKE_WORKFLOW_API = 'https://api.github.com/repos/hachi515/karaoke_setlist/actions/workflows/update.yml';

function karaokeToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_ACTIONS_TOKEN');
  if (!token || !token.trim()) {
    throw new Error('スクリプトプロパティ GITHUB_ACTIONS_TOKEN を設定してください。値をログやソースへ貼らないでください。');
  }
  return token.trim();
}

function karaokeGitHubRequest_(method, suffix, token, body) {
  const options = {
    method: method,
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    muteHttpExceptions: true
  };
  if (body !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }
  for (let attempt = 0; attempt < 3; attempt++) {
    const response = UrlFetchApp.fetch(KARAOKE_WORKFLOW_API + suffix, options);
    const status = response.getResponseCode();
    if (status >= 200 && status < 300) return response;
    if ((status === 429 || status >= 500) && attempt < 2) {
      const headers = response.getAllHeaders();
      const retryKey = Object.keys(headers).find(key => key.toLowerCase() === 'retry-after');
      const retrySeconds = retryKey ? Number(headers[retryKey]) : NaN;
      // 長いレート制限は連続リトライせず、既存トリガーを残して失敗を通知する。
      if (Number.isFinite(retrySeconds) && retrySeconds > 30) break;
      const waitMs = Number.isFinite(retrySeconds) && retrySeconds > 0
        ? retrySeconds * 1000 : (attempt + 1) * 5000;
      Utilities.sleep(waitMs);
      continue;
    }
    throw new Error('GitHub API HTTP ' + status + '。トークンのActions書き込み権限・有効期限とActions実行履歴を確認してください。');
  }
  throw new Error('GitHub APIが一時的に利用できません。GASの実行履歴を確認してください。');
}

/** GASの実行画面から一度実行する。接続確認後に毎時トリガーを作り、初回更新も起動する。 */
function installKaraokeHourlyTrigger() {
  const token = karaokeToken_();
  const response = karaokeGitHubRequest_('get', '', token);
  const workflow = JSON.parse(response.getContentText());
  if (workflow.state !== 'active' || workflow.path !== '.github/workflows/update.yml') {
    throw new Error('対象のUpdate Listが有効ではありません。GitHub Actionsの設定を確認してください。');
  }
  const previous = ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction() === KARAOKE_DISPATCH_HANDLER);
  // 作成に失敗した場合でも、既存のトリガーは保持する。
  ScriptApp.newTrigger(KARAOKE_DISPATCH_HANDLER).timeBased().everyHours(1).create();
  previous.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  console.log('Update Listの毎時トリガーを設定しました。');
  dispatchKaraokeUpdate();
}

/** 時間主導型トリガーのハンドラー。GASのCSV・既存処理は変更しない。 */
function dispatchKaraokeUpdate() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('別の起動処理が実行中のため、今回の起動要求を省略しました。');
    return;
  }
  try {
    karaokeGitHubRequest_('post', '/dispatches', karaokeToken_(), {ref: 'main', inputs: {trigger_source: 'gas-hourly'}});
    const now = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('KARAOKE_LAST_DISPATCH_AT', now);
    console.log('Update Listの起動要求を送信しました: ' + now + '（生成・公開の成否はGitHub Actionsで確認）');
  } finally {
    lock.releaseLock();
  }
}

/** このファイルが作った毎時トリガーだけを解除する。 */
function removeKaraokeHourlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === KARAOKE_DISPATCH_HANDLER)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}
