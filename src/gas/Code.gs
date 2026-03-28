/**
 * Code.gs — メインエントリ
 *
 * ボタン押下で startProcessing() を実行。
 * 10枚ずつ処理し、未処理が残っていれば即トリガーを登録して連続処理する。
 */

/**
 * C列に注文コードが手動入力されたときにE・F・G列を生成する
 * ※ onEditはシンプルトリガー（外部API不可）なので installableトリガー経由で呼ぶ
 */
function onEditTrigger(e) {
  const sheet = e.range.getSheet();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  // 3行目以降のみ対象（1行目=警告行、2行目=ヘッダー）
  if (row < 3) return;

  // 集計シート名パターン（例：【PDS_2026-03-17】集計）のみ対象
  if (!sheet.getName().match(/【.+_\d{4}-\d{2}-\d{2}】集計/)) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // C列（3列目）: 進捗が「完了」になったら行全体に色をつける
  if (col === 3) {
    const value = sheet.getRange(row, 3).getValue();
    const totalCols = sheet.getLastColumn();
    if (value === '完了') {
      sheet.getRange(row, 1, 1, totalCols).setBackground('#D9EAD3'); // 緑
    } else {
      sheet.getRange(row, 1, 1, totalCols).setBackground(null); // 色をリセット
    }
    return;
  }

  // F列（6列目）: 注文コード入力時にG・H・I・J列を生成
  if (col !== 6) return;

  // 数式をセットしてから評価・H・I・J列を生成
  SheetsUtils._setRowFormulas(sheet, row);
  SpreadsheetApp.flush();
  const orderCode = sheet.getRange(row, 6).getValue();
  SheetsUtils._setRowTitle(sheet, row, ss, orderCode);

  // G列: カテゴリ（AH=都道府県, AI=エリア, AK=今回のご利用用途）を空でないものをスペース区切りで結合
  const catVals = sheet.getRange(row, 34, 1, 4).getValues()[0]; // AH=34, AI=35, AJ=36, AK=37
  const category = [catVals[0], catVals[1], catVals[3]].filter(v => v).join(' '); // AH, AI, AK
  sheet.getRange(row, 7).setValue(category);

  // Z〜AF列からanswersを復元してI列を再生成
  // Y=25:年齢, Z=26:普段サイズ(空), AA=27:トップス, AB=28:ボトムス, AC=29:身長, AD=30:マタニティ, AE=31:骨格
  const vals = sheet.getRange(row, 26, 1, 7).getValues()[0];
  // AM〜AQ列から品質・着心地・サイズ感・ドレス丈を取得（AM=39, AN=40, AO=41, AP=42, AQ=43）
  const qVals = sheet.getRange(row, 39, 1, 5).getValues()[0];
  const answers = {
    age:        vals[0],
    topSize:    vals[2],
    bottomSize: vals[3],
    height:     vals[4],
    maternity:  vals[5],
    bodyType:   vals[6],
    q3Quality:  qVals[0],
    q3Comfort:  qVals[1],
    q3Size:     qVals[3],
    q3Length:   qVals[4],
  };
  SheetsUtils._setRowSizeHtml(ss, sheet, row, answers);
  SheetsUtils._setRowDressHtml(ss, sheet, row);

  // 注文コード入力時に商品画像を保存
  const fileName = sheet.getRange(row, 5).getValue(); // E列: ファイル名
  ImageUtils.saveImages({ orderCode, fileName, sheet, row });
}

/**
 * スプレッドシートを開いたときに onEditTrigger が未登録なら自動登録する
 * 月スプシをコピーした初回オープン時に自動でトリガーが設定される
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PDS管理')
    .addItem('月初セットアップ（初回のみ）', 'setupOnEditTrigger')
    .addToUi();
  ui.createMenu('⚠️ 必須処理')
    .addItem('値で確定（注文コード入力完了後に実行）', 'confirmValues')
    .addToUi();

  // セットアップ済みフラグを確認（承認不要）
  const isSetup = PropertiesService.getDocumentProperties().getProperty('TRIGGER_SETUP_DONE');
  if (!isSetup) {
    SpreadsheetApp.getUi().alert(
      '月初セットアップが必要です\n\nPDS管理メニューから実行してください。'
    );
  }
}

/**
 * installableトリガーを登録する（初回のみ手動実行）
 */
function setupOnEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 既存トリガーを削除してから再登録
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onEditTrigger')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // マスタのスクリプトプロパティをこのスクリプトにコピー
  _copyPropertiesFromMaster();

  // セットアップ済みフラグを保存
  PropertiesService.getDocumentProperties().setProperty('TRIGGER_SETUP_DONE', 'true');
  SpreadsheetApp.getUi().alert('セットアップ完了！\n注文コード入力が自動で動作するようになりました。');
}

/**
 * マスタスプシの【設定】シートからフォルダIDを読み取り、
 * このスクリプトのプロパティに自動登録する。
 */
function _copyPropertiesFromMaster() {
  const MASTER_SS_ID = '1SRx_la6Tl_Y-Lrm5tF10yzPyQfEmiwzcZgydsG4_iKM';
  const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
  const sheet = masterSs.getSheetByName('設定');
  if (!sheet || sheet.getLastRow() < 1) {
    Logger.log('_copyPropertiesFromMaster: 【設定】シートが見つかりません');
    return;
  }

  const data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  const localProps = PropertiesService.getScriptProperties();
  data.forEach(([key, value]) => {
    if (key && value) localProps.setProperty(String(key), String(value));
  });
  Logger.log('_copyPropertiesFromMaster: プロパティコピー完了（' + data.length + '件）');
}

/**
 * アクティブな日付シートの全データ行（3行目以降）の数式を値に変換する。
 * CSV貼付の内容が変わっても参照先が消えないようにするための確定処理。
 */
function confirmValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 日付シート（【brand_yyyy-MM-dd】集計）のみ対象
  if (!sheet.getName().match(/【.+_\d{4}-\d{2}-\d{2}】集計/)) {
    SpreadsheetApp.getUi().alert('このシートは対象外です。\n【ブランド_yyyy-MM-dd】集計シートで実行してください。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('データがありません。');
    return;
  }

  // 3行目以降の全データを値に変換
  const lastCol = sheet.getLastColumn();
  const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol);
  const values = dataRange.getValues();
  dataRange.setValues(values);

  SpreadsheetApp.getUi().alert('値で確定しました（' + (lastRow - 2) + '行）。');
}

/**
 * ボタンから呼び出すエントリポイント。
 * 既存の連続処理トリガーをリセットしてから checkNewFiles() を実行する。
 */
function startProcessing() {
  // マスタのCSV貼付シートにデータがあるか確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const csvSheet = ss.getSheetByName('CSV貼付');
  if (!csvSheet || csvSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('CSV貼付シートにデータがありません。\nCSVを貼り付けてから処理を開始してください。');
    return;
  }

  _deleteContinueTriggers();
  checkNewFiles();
}

/**
 * 未処理ファイルを最大10枚処理する。
 * 処理後に未処理が残っていれば即トリガーを登録して連続処理する。
 * 全件完了したら完了通知を送る。
 */
function checkNewFiles() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('別の処理が実行中のためスキップします');
    return;
  }

  try {
    _checkNewFilesCore();
  } finally {
    lock.releaseLock();
  }
}

function _checkNewFilesCore() {
  const BATCH_SIZE = 7;
  const processedNames = SheetsUtils.getProcessedFileNames();
  const allowed = ['application/pdf', 'image/jpeg', 'image/png'];

  const targets = [
    { folderId: Config.DRIVE_FOLDER_OK_ID, isPublishOk: true  },
    { folderId: Config.DRIVE_FOLDER_NG_ID, isPublishOk: false },
  ];

  // 未処理ファイルを全件収集してファイル名昇順にソート
  const pending = [];
  targets.forEach(({ folderId, isPublishOk }) => {
    if (!folderId) return;
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log('フォルダ取得エラー (folderId=' + folderId + '): ' + e.message);
      return;
    }
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (processedNames.includes(file.getName())) continue;
      if (!allowed.includes(file.getMimeType())) continue;
      pending.push({ file, isPublishOk });
    }
  });
  // OK優先（isPublishOk=true を先に）、同じフォルダ内はファイル名昇順
  pending.sort((a, b) => {
    if (a.isPublishOk !== b.isPublishOk) return a.isPublishOk ? -1 : 1;
    return a.file.getName().localeCompare(b.file.getName());
  });

  if (pending.length === 0) {
    _clearProgressCell();
    NotifyUtils.notifyComplete(0);
    return;
  }

  const okCount = pending.filter(p => p.isPublishOk).length;
  const ngCount = pending.filter(p => !p.isPublishOk).length;
  _setProgressCell(okCount, ngCount);

  // 最大10枚を処理
  const batch = pending.slice(0, BATCH_SIZE);
  let count = 0;
  batch.forEach(({ file, isPublishOk }) => {
    try {
      processAnketFile(file, isPublishOk);
      count++;
    } catch (e) {
      Logger.log('ERROR: ' + file.getName() + ' — ' + e.message);
      Logger.log('スタック: ' + e.stack);
      SheetsUtils.writeLog(file.getName(), 'ERROR', e.message);
    }
  });

  // 未処理が残っていれば即トリガーを登録して続きを処理
  const afterBatch = pending.slice(batch.length);
  const remaining = afterBatch.length;
  const remainingOk = afterBatch.filter(p => p.isPublishOk).length;
  const remainingNg = afterBatch.filter(p => !p.isPublishOk).length;
  Logger.log('処理枚数: ' + batch.length + ' / 残り: ' + remaining + ' / pending合計: ' + pending.length);
  if (remaining > 0) {
    _setProgressCell(remainingOk, remainingNg);
    Logger.log('続きのトリガーを登録します');
    _deleteContinueTriggers();
    _scheduleContinueTrigger();
  } else {
    Logger.log('全件完了 → 完了通知');
    _clearProgressCell();
    NotifyUtils.notifyComplete(count);
  }
}

/**
 * 開始用シートのC13（掲載OK）・C14（掲載NG）セルに残件数を表示する
 * @param {number} okCount
 * @param {number} ngCount
 */
function _setProgressCell(okCount, ngCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('開始用');
  if (!sheet) return;
  sheet.getRange('C11').setValue(`処理中... (残り${okCount}件)`);
  sheet.getRange('C12').setValue(`処理中... (残り${ngCount}件)`);
  SpreadsheetApp.flush();
}

/**
 * 開始用シートのC13・C14セルをクリアする
 */
function _clearProgressCell() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('開始用');
  if (!sheet) return;
  sheet.getRange('C11').clearContent();
  sheet.getRange('C12').clearContent();
}


/**
 * checkNewFiles を1分後に実行するトリガーを登録する。
 * GASの最短トリガー間隔は1分のため、これが最速の連続実行手段。
 */
function _scheduleContinueTrigger() {
  ScriptApp.newTrigger('checkNewFiles')
    .timeBased()
    .after(60 * 1000)
    .create();
}

/**
 * 連続処理用トリガー（checkNewFiles）を全削除する。
 */
function _deleteContinueTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkNewFiles')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function processAnketFile(file, isPublishOk) {
  SheetsUtils.writeLog(file.getName(), 'PROCESSING', '');

  const imageData = OcrUtils.getImageData(file);
  const answers = AiUtils.readAnswersFromImage(imageData);
  answers.fileName = file.getName();
  answers.isPublishOk = isPublishOk;

  // 感情分析
  const analyzed = AiUtils.analyzeSentiment(answers);
  analyzed.isPublishOk = isPublishOk;

  // freeCommentがある場合のみ返信文を生成
  if (analyzed.freeComment) {
    analyzed.aiReply = AiUtils.generateReply(analyzed);
  } else {
    analyzed.aiReply = '';
  }

  // 集計シートに書き込む（HTML・返信文生成後）
  SheetsUtils.writeToSummarySheet(analyzed);

  // 意見まとめシートに書き込む
  SheetsUtils.writeToOpinionSheet(analyzed);

  SheetsUtils.writeLog(file.getName(), 'DONE', '');

  // 処理済みファイルの先頭に「済_」を付けてリネーム（二重付与を防止）
  const currentName = file.getName();
  if (!currentName.startsWith('済_')) {
    file.setName('済_' + currentName);
  }
}

/**
 * 全シートを一括作成する（初回セットアップ時に実行）
 * 既存シートはスキップ、既存内容は変更しない
 */
function initSheets() {
  const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);

  const summaryHeaders = [
    '処理日時', 'ファイル名', '注文コード', '年齢', '普段の洋服のサイズ', 'トップスサイズ', 'ボトムスサイズ',
    '身長(cm)', 'マタニティ', '骨格タイプ', '利用エリア', '都道府県', 'エリア',
    'どこで知りましたか？', '今回のご利用用途', 'ご利用商品はいかがでしたか？',
    '品質', '着心地', '※あまり良くないと答えた方',
    'サイズ感', 'ドレス丈/パンツ丈', 'レンタルの流れはいかがでしたか？',
    '※不便と答えた人用', '着用頂いた感想・サイズ感・レンタルの流れについてのご意見'
  ];

  const opinionHeaders = [
    '処理日時', '自由意見', '品質・着心地NG理由', 'レンタルNG理由',
    '感情', 'AI返信文', 'HTMLスニペット'
  ];

  const logHeaders = ['処理日時', 'ファイル名', 'ステータス', 'エラー内容'];

  const toneHeaders = ['項目', '設定値'];
  const toneData = [
    ['ブランド名・署名', 'PARTY DRESS STYLE'],
    ['回答の口調・文体', '丁寧でフレンドリー'],
    ['共感フレーズ（固定）', 'この度はご利用いただきありがとうございます！'],
    ['禁止語・NGワード', ''],
    ['回答文字数目安', '100〜150文字'],
    ['ネガティブ時の追記', '大変申し訳ございませんでした。改善に努めてまいります。']
  ];

  const sheets = [
    { name: '【PDS】累計集計',  headers: summaryHeaders, color: '#D9E8FC' },
    { name: '【OTONA】累計集計', headers: summaryHeaders, color: '#FCE4D6' },
    { name: '【PDS】意見まとめ',  headers: opinionHeaders, color: '#D9E8FC' },
    { name: '【OTONA】意見まとめ', headers: opinionHeaders, color: '#FCE4D6' },
    { name: '【ログ】処理履歴',  headers: logHeaders,    color: '#F2F2F2' },
    { name: '【設定】トンマナ',  headers: toneHeaders,   color: '#FFF2CC', data: toneData },
  ];

  sheets.forEach(({ name, headers, color, data }) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(color);
      if (data) {
        data.forEach(row => sheet.appendRow(row));
      }
      Logger.log('作成: ' + name);
    } else {
      Logger.log('スキップ（既存）: ' + name);
    }
  });

  Logger.log('initSheets 完了');
}

function testSetRowTitle() {
  const monthlySs = SpreadsheetApp.openById('1O9f0Ka805NSYfnpOmxkJfy1nRYfJzkBlj3o8gCTPeOU');
  const sheet = monthlySs.getSheets().find(s => s.getName().match(/【.+_\d{4}-\d{2}-\d{2}】集計/));
  if (!sheet) { Logger.log('日別シートが見つかりません'); return; }
  const orderCode = sheet.getRange(2, 5).getValue();
  Logger.log('対象シート: ' + sheet.getName() + ' 注文コード: ' + orderCode);
  SheetsUtils._setRowTitle(sheet, 2, monthlySs, orderCode);
}

function testCreateMonthlySpreadsheet() {
  const ss = SheetsUtils._getOrCreateMonthlySpreadsheet('PDS', '2026-03');
  Logger.log('月スプシID: ' + ss.getId());
  Logger.log('シート一覧: ' + ss.getSheets().map(s => s.getName()).join(', '));
}

function testCreateMonthlyDebug() {
  Logger.log('DRIVE_FOLDER_MONTHLY_ID: ' + Config.DRIVE_FOLDER_MONTHLY_ID);
  Logger.log('DRIVE_FOLDER_OK_ID: ' + Config.DRIVE_FOLDER_OK_ID);
  Logger.log('SPREADSHEET_ID: ' + Config.SPREADSHEET_ID);

  // OK フォルダにアクセスできるか確認
  try {
    const okFolder = DriveApp.getFolderById(Config.DRIVE_FOLDER_OK_ID);
    Logger.log('OKフォルダ名: ' + okFolder.getName());
  } catch(e) {
    Logger.log('OKフォルダエラー: ' + e.message);
  }

  // Monthly フォルダにアクセスできるか確認
  try {
    const monthlyFolder = DriveApp.getFolderById(Config.DRIVE_FOLDER_MONTHLY_ID);
    Logger.log('Monthlyフォルダ名: ' + monthlyFolder.getName());
  } catch(e) {
    Logger.log('Monthlyフォルダエラー: ' + e.message);
  }

  // マスタファイルにアクセスできるか確認
  try {
    const masterFile = DriveApp.getFileById(Config.SPREADSHEET_ID);
    Logger.log('マスタファイル名: ' + masterFile.getName());
  } catch(e) {
    Logger.log('マスタファイルエラー: ' + e.message);
  }
}

/**
 * Apps Script APIのトリガー登録をテストする（単体デバッグ用）
 * 実行してログを確認し、エラー内容を把握する
 */
function testRegisterTrigger() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log('対象SpreadsheetId: ' + spreadsheetId);
  SheetsUtils._registerTriggerViaApi(spreadsheetId);
  Logger.log('完了');
}

/**
 * マスタスプシのセットアップ済みフラグを立てる（マスタでのポップアップを抑制）
 */
function markMasterAsSetup() {
  PropertiesService.getDocumentProperties().setProperty('TRIGGER_SETUP_DONE', 'true');
  Logger.log('マスタセットアップ済みフラグ設定完了');
}

function testOcr() {
  const folder = DriveApp.getFolderById(Config.DRIVE_FOLDER_ID);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const mime = file.getMimeType();
    if (mime !== 'image/jpeg' && mime !== 'image/png' && mime !== 'application/pdf') continue;
    Logger.log('テスト対象: ' + file.getName());
    const imageData = OcrUtils.getImageData(file);
    const answers = AiUtils.readAnswersFromImage(imageData);
    Logger.log('読み取り結果: ' + JSON.stringify(answers));
    Logger.log('HTMLスニペット: ' + AiUtils.generateHtml(answers));
    break;
  }
}
