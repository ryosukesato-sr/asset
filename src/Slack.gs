/**
 * Slack連携（スクリプトプロパティで設定）
 * 設定は「アプリを開く」→ 設定 から行うか、GASエディタの「プロジェクトの設定」→「スクリプトプロパティ」で設定可能。
 */

const SLACK_PROP_WEBHOOK = 'SLACK_WEBHOOK_URL';
const SLACK_PROP_ENABLED = 'SLACK_ALERT_ENABLED';

/**
 * Slack設定を取得（UI用。URLはマスクして返す）
 */
function getSlackSettings() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty(SLACK_PROP_WEBHOOK) || '';
  const enabled = props.getProperty(SLACK_PROP_ENABLED);
  return {
    webhookUrlSet: url.length > 0,
    webhookUrlMasked: url ? url.substring(0, 50) + '...' : '',
    enabled: enabled !== '0'
  };
}

/**
 * Slack設定を保存（設定画面から呼び出し）
 */
function setSlackSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  if (settings.webhookUrl !== undefined) {
    const v = String(settings.webhookUrl || '').trim();
    if (v) props.setProperty(SLACK_PROP_WEBHOOK, v);
    else props.deleteProperty(SLACK_PROP_WEBHOOK);
  }
  if (settings.enabled !== undefined) {
    props.setProperty(SLACK_PROP_ENABLED, settings.enabled ? '1' : '0');
  }
  return { success: true };
}

/**
 * Slackにメッセージを送信（Webhook）
 */
function sendSlackMessage(text) {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty(SLACK_PROP_WEBHOOK);
  if (!url || props.getProperty(SLACK_PROP_ENABLED) === '0') return { sent: false };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
    return { sent: res.getResponseCode() === 200 };
  } catch (e) {
    console.error('Slack send error:', e);
    return { sent: false };
  }
}
