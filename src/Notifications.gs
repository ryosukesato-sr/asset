/**
 * 日次アラート（リース・保証期限 / 返却待ちリマインド）とステータス変更通知
 */

/**
 * 日次トリガーから呼ぶ: リース・保証期限の30日前・7日前にSlack通知、返却待ちは返却予定日+1日にリマインド
 */
function runDailyAlerts() {
  const assets = getAssets();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const messages = [];

  assets.forEach(a => {
    const lease = parseDateSafe(a.leaseEndDate);
    const warranty = parseDateSafe(a.warrantyEndDate);
    if (lease) {
      const days = Math.floor((lease - today) / (24 * 60 * 60 * 1000));
      if (days === 30 || days === 7) {
        messages.push(`【リース終了】${days}日後: ${a.id} ${a.name}（${a.userName || '未割当'}）`);
      } else if (days >= 0 && days < 7) {
        messages.push(`【リース終了】あと${days}日: ${a.id} ${a.name}（${a.userName || '未割当'}）`);
      }
    }
    if (warranty) {
      const days = Math.floor((warranty - today) / (24 * 60 * 60 * 1000));
      if (days === 30 || days === 7) {
        messages.push(`【保証期限】${days}日後: ${a.id} ${a.name}（${a.userName || '未割当'}）`);
      } else if (days >= 0 && days < 7) {
        messages.push(`【保証期限】あと${days}日: ${a.id} ${a.name}（${a.userName || '未割当'}）`);
      }
    }
    // 返却待ち: 返却予定日から1日経過でリマインド
    if (a.status === '返却待ち' && a.returnDueDate) {
      const due = parseDateSafe(a.returnDueDate);
      if (due) {
        const yesterday = new Date(today);
        yesterday.setDate(yesterday.getDate() - 1);
        if (due.getTime() <= yesterday.getTime()) {
          messages.push(`【返却リマインド】返却予定日を過ぎています: ${a.id} ${a.name}（${a.userName || ''} / ${a.userEmail || ''}）`);
        }
      }
    }
  });

  if (messages.length > 0) {
    const text = 'IT資産管理 期限・リマインド\n' + messages.join('\n');
    sendSlackMessage(text);
    // 返却リマインドのみメール送信（使用者に）
    messages.forEach(msg => {
      if (msg.indexOf('返却リマインド') !== -1) {
        const m = msg.match(/（([^/]+) \/ ([^）]+)）/);
        if (m && m[2]) {
          try {
            MailApp.sendEmail(m[2].trim(), '【IT資産】返却リマインド', msg.replace(/【返却リマインド】/, ''));
          } catch (e) {}
        }
      }
    });
  }
}

/**
 * ステータス変更時にSlack通知（と必要ならメール）
 */
function notifyStatusChange(assetId, assetName, oldStatus, newStatus, userName, userEmail) {
  const text = `【IT資産】ステータス変更: ${assetId} ${assetName || ''}\n${oldStatus} → ${newStatus}\n使用者: ${userName || '-'} ${userEmail || ''}`;
  sendSlackMessage(text);
}

/**
 * 日次アラート用トリガーを1本だけ設定（メニューから実行）
 */
function installDailyAlertTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runDailyAlerts') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(8) // 毎朝8時
    .create();
  SpreadsheetApp.getUi().alert('日次アラートのトリガーを設定しました。毎朝8時頃に実行されます。');
}
