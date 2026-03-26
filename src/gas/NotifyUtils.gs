/**
 * NotifyUtils.gs — 通知ユーティリティ
 * Gmail への通知を担当する
 */

const NotifyUtils = {
  /**
   * 処理完了をGmailへ通知する
   * @param {string} fileName
   */
  notifyComplete(count) {
    this._sendEmail(
      `【完了】アンケート処理完了`,
      `${count}枚処理しました。`
    );
  },

  /**
   * エラーをGmailへ通知する
   * @param {string} fileName
   * @param {string} errorMsg
   */
  notifyError(fileName, errorMsg) {
    this._sendEmail(
      `【エラー】アンケート処理失敗: ${fileName}`,
      `ファイル名: ${fileName}\nエラー内容: ${errorMsg}`
    );
  },

  /**
   * Gmailでメールを送信する
   * @param {string} subject
   * @param {string} body
   */
  _sendEmail(subject, body) {
    Logger.log(`[通知] ${subject}`);
    const token = Config.CHATWORK_API_TOKEN;
    const roomId = Config.CHATWORK_ROOM_ID;
    if (!token || !roomId) {
      Logger.log('[通知] CHATWORK_API_TOKEN または CHATWORK_ROOM_ID が未設定');
      return;
    }
    const message = `[info][title]${subject}[/title]${body}[/info]`;
    UrlFetchApp.fetch(`https://api.chatwork.com/v2/rooms/${roomId}/messages`, {
      method: 'post',
      headers: { 'X-ChatWorkToken': token },
      payload: { body: message }
    });
  }
};
