/**
 * SheetsUtils.gs — Google Sheets操作
 * 集計シート・意見シート・ログシートへの書き込みを担当する
 * シート名はブランド別（PDS / OTONA）
 */

const SheetsUtils = {
  /**
   * 【ログ】処理履歴シートからDONE済みのファイル名一覧を返す
   * @returns {Array} 処理済みファイル名の配列
   */
  getProcessedFileNames() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('【ログ】処理履歴');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    return data
      .filter(row => row[2] === 'DONE')
      .map(row => row[1]);
  },

  /**
   * 累計集計シートと日別集計シートへ回答を書き込む
   * isPublishOk=true  → 月スプシ（日別シート）+ 累計シート
   * isPublishOk=false → 累計シートのみ
   * @param {Object} answers - 回答マップ（brand, isPublishOk含む）
   */
  writeToSummarySheet(answers) {
    const masterSs = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
    const brand = answers.brand || 'PDS';
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    const month = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    const row = this._buildSummaryRow(answers, now);

    // 掲載OKのみ月スプシ（日別シート）に書き込む
    if (answers.isPublishOk) {
      const monthlySs = this._getOrCreateMonthlySpreadsheet(brand, month);
      this._syncCsvSheet(masterSs, monthlySs);
      const dailySheet = this._getOrCreateDailySheet(monthlySs, brand, today);
      // D列（処理日時）で行数を判定（A列=備考は常に空のため）
      const lastDataRow = dailySheet.getRange('D:D').getValues().filter(r => r[0] !== '').length + 1;
      dailySheet.getRange(lastDataRow, 1, 1, row.length).setValues([row]);
      this._setProgressCell(dailySheet, lastDataRow);
      this._setRowFormulas(dailySheet, lastDataRow);
      this._setRowTitle(dailySheet, lastDataRow, monthlySs, answers.orderCode);
      this._setRowSizeHtml(masterSs, dailySheet, lastDataRow, answers);
      this._setRowDressHtml(masterSs, dailySheet, lastDataRow);
    }

    // OK/NG 両方を累計シートに書き込む
    const totalSheetName = `【${brand}】累計集計`;
    const totalSheet = masterSs.getSheetByName(totalSheetName);
    if (totalSheet) {
      const cumulativeRow = this._buildCumulativeRow(answers, now);
      // A列（処理日時）で行数を判定
      const totalLastRow = totalSheet.getRange('A:A').getValues().filter(r => r[0] !== '').length + 1;
      totalSheet.getRange(totalLastRow, 1, 1, cumulativeRow.length).setValues([cumulativeRow]);
    }
  },

  /**
   * 意見まとめシートへ感情分析・HTML・返信文を書き込む
   * @param {Object} answers - 感情分析・HTML・返信文済み回答マップ
   */
  writeToOpinionSheet(answers) {
    const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
    const brand = answers.brand || 'PDS';
    const sheetName = `【${brand}】意見まとめ`;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('シートが見つかりません: ' + sheetName);
      return;
    }
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    const row = [
      now,
      answers.freeComment || ''
    ];
    sheet.appendRow(row);
  },

  /**
   * 【ログ】処理履歴シートへ記録する
   * @param {string} fileName
   * @param {string} status - 'PROCESSING' | 'DONE' | 'ERROR'
   * @param {string} errorMsg
   */
  writeLog(fileName, status, errorMsg) {
    const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('【ログ】処理履歴');
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([now, fileName, status, errorMsg]);
  },

  /**
   * 集計行データを構築する
   *
   * A: 備考（手動入力・自由記載）
   * B: 担当者（手動入力）
   * C: 進捗（手動入力）
   * D: 処理日時
   * E: ファイル名
   * F: 注文コード（手動入力）
   * G: ご利用日（数式）
   * H: タイトル（onEdit生成）
   * I: 【お客様の普段の洋服サイズ】（_setRowSizeHtmlで生成）
   * J: ☆お客様がレンタルしたドレスはこちら☆（onEdit生成）
   * K: 着用頂いた感想
   * L: AI返信文
   * M〜P: 商品_1（コード・リンク・単価・商品詳細）数式
   * Q〜T: 商品_2（コード・リンク・単価・商品詳細）数式
   * U〜X: 商品_3（コード・リンク・単価・商品詳細）数式
   * Y: 年齢
   * Z: 普段の洋服のサイズ（ブランク）
   * AA: トップスサイズ
   * AB: ボトムスサイズ
   * AC: 身長(cm)
   * AD: マタニティ
   * AE: 骨格タイプ
   * AF: 利用エリア（ブランク）
   * AG: 都道府県
   * AH: エリア
   * AI: どこで知りましたか？
   * AJ: 今回のご利用用途
   * AK: ご利用商品はいかがでしたか？（ブランク）
   * AL: 品質
   * AM: 着心地
   * AN: ※あまり良くないと答えた方
   * AO: サイズ感
   * AP: ドレス丈/パンツ丈
   * AQ: レンタルの流れはいかがでしたか？
   * AR: ※不便と答えた人用
   */
  _buildSummaryRow(answers, now) {
    return [
      '',                         // A: 備考（手動入力・自由記載）
      '',                         // B: 担当者（手動入力）
      '',                         // C: 進捗（手動入力）
      now,                        // D: 処理日時
      answers.fileName || '',     // E: ファイル名
      answers.orderCode || '',    // F: 注文コード（手動入力）
      '',                         // G: ご利用日（数式で自動入力）
      '',                         // H: タイトル（onEdit生成）
      '',                         // I: 【お客様の普段の洋服サイズ】（_setRowSizeHtmlで生成）
      '',                         // J: ☆お客様がレンタルしたドレスはこちら☆（onEdit生成）
      answers.freeComment ? `【ご利用頂いた感想・ご意見をお願いします】\n${answers.freeComment}` : '',  // K: 着用頂いた感想
      answers.aiReply ? `【スタッフコメント】\n${answers.aiReply}` : '',  // L: AI返信文
      '', '', '', '',             // M〜P: 商品_1（数式で自動入力）
      '', '', '', '',             // Q〜T: 商品_2（数式で自動入力）
      '', '', '', '',             // U〜X: 商品_3（数式で自動入力）
      answers.age || '',          // Y: 年齢
      '',                         // Z: 普段の洋服のサイズ（ブランク）
      answers.topSize || '',      // AA: トップスサイズ
      answers.bottomSize || '',   // AB: ボトムスサイズ
      answers.height || '',       // AC: 身長(cm)
      answers.maternity || '',    // AD: マタニティ
      answers.bodyType || '',     // AE: 骨格タイプ
      '',                         // AF: 利用エリア（ブランク）
      answers.prefecture || '',   // AG: 都道府県
      answers.area || '',         // AH: エリア
      answers.q1 || '',           // AI: どこで知りましたか？
      answers.q2 || '',           // AJ: 今回のご利用用途
      '',                         // AK: ご利用商品はいかがでしたか？（ブランク）
      answers.q3Quality || '',    // AL: 品質
      answers.q3Comfort || '',    // AM: 着心地
      answers.q3NgReason || '',   // AN: ※あまり良くないと答えた方
      answers.q3Size || '',       // AO: サイズ感
      answers.q3Length || '',     // AP: ドレス丈/パンツ丈
      answers.q4Flow || '',       // AQ: レンタルの流れはいかがでしたか？
      answers.q4NgReason || '',   // AR: ※不便と答えた人用
    ];
  },

  /**
   * 累計集計シート用の行データを生成する（新構成: A=処理日時から始まる21列）
   * @param {Object} answers - 回答マップ
   * @param {string} now - 処理日時文字列
   * @returns {Array} 21要素の配列
   */
  _buildCumulativeRow(answers, now) {
    return [
      now,                          // A: 処理日時
      answers.age || '',            // B: 年齢
      '',                           // C: 普段の洋服のサイズ（ブランク）
      answers.topSize || '',        // D: トップスサイズ
      answers.bottomSize || '',     // E: ボトムスサイズ
      answers.height || '',         // F: 身長(cm)
      answers.maternity || '',      // G: マタニティ
      answers.bodyType || '',       // H: 骨格タイプ
      '',                           // I: 利用エリア（ブランク）
      answers.prefecture || '',     // J: 都道府県
      answers.area || '',           // K: エリア
      answers.q1 || '',             // L: どこで知りましたか？
      answers.q2 || '',             // M: 今回のご利用用途
      '',                           // N: ご利用商品はいかがでしたか？（ブランク）
      answers.q3Quality || '',      // O: 品質
      answers.q3Comfort || '',      // P: 着心地
      answers.q3NgReason || '',     // Q: ※あまり良くないと答えた方
      answers.q3Size || '',         // R: サイズ感
      answers.q3Length || '',       // S: ドレス丈/パンツ丈
      answers.q4Flow || '',         // T: レンタルの流れはいかがでしたか？
      answers.q4NgReason || '',     // U: ※不便と答えた人用
    ];
  },

  /**
   * C列（進捗）にドロップダウンとデフォルト値「未対応」をセットする
   */
  _setProgressCell(sheet, row) {
    const cell = sheet.getRange(row, 3);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['未対応', '対応中', '完了'], true)
      .setAllowInvalid(false)
      .build();
    cell.setDataValidation(rule);
    cell.setValue('未対応');
  },

  /**
   * マスタの CSV貼付 シートの全内容を月スプシの CSV貼付 に上書きコピーする
   */
  _syncCsvSheet(masterSs, monthlySs) {
    const srcSheet = masterSs.getSheetByName('CSV貼付');
    const dstSheet = monthlySs.getSheetByName('CSV貼付');
    if (!srcSheet || !dstSheet) return;

    const lastRow = srcSheet.getLastRow();
    const lastCol = srcSheet.getLastColumn();
    if (lastRow < 1) return;

    const srcData = srcSheet.getRange(1, 1, lastRow, lastCol).getValues();

    // 既存データのA+Bキーセットを作成
    const dstLastRow = dstSheet.getLastRow();
    const existingKeys = new Set();
    if (dstLastRow > 0) {
      dstSheet.getRange(1, 1, dstLastRow, 2).getValues().forEach(r => {
        existingKeys.add(`${r[0]}_${r[1]}`);
      });
    }

    // 重複しない行のみ追記
    const newRows = srcData.filter(r => !existingKeys.has(`${r[0]}_${r[1]}`));
    if (newRows.length === 0) return;
    dstSheet.getRange(dstLastRow + 1, 1, newRows.length, lastCol).setValues(newRows);
  },

  /**
   * ブランド・月ごとのスプレッドシートを取得または新規作成する
   * スプシ名: 【PDS_2026-03】集計
   * 既存スプシはドライブ内をタイトル検索して再利用する
   * @param {string} brand
   * @param {string} month - 'yyyy-MM'
   * @returns {Spreadsheet}
   */
  _getOrCreateMonthlySpreadsheet(brand, month) {
    const title = `【${brand}_${month}】集計`;
    // 月スプシ保存先フォルダ（未設定の場合はマイドライブ直下）
    const monthlyFolderId = Config.DRIVE_FOLDER_MONTHLY_ID;
    const parentFolder = monthlyFolderId
      ? DriveApp.getFolderById(monthlyFolderId)
      : DriveApp.getRootFolder();
    const files = parentFolder.getFilesByName(title);
    if (files.hasNext()) {
      return SpreadsheetApp.openById(files.next().getId());
    }
    // マスタをコピー（スクリプトも含まれる）
    const masterFile = DriveApp.getFileById(Config.SPREADSHEET_ID);
    const newFile = masterFile.makeCopy(title, parentFolder);
    const newSs = SpreadsheetApp.openById(newFile.getId());

    // 不要なシートを削除（削除できないシートはスキップ）
    // 残すシート: CSV貼付・DB・SET内容（数式参照に必要）
    const keepSheets = ['CSV貼付', 'DB', 'SET内容'];
    newSs.getSheets().forEach(sheet => {
      if (keepSheets.includes(sheet.getName())) return;
      try { newSs.deleteSheet(sheet); } catch(e) { /* 削除できない場合はスキップ */ }
    });

    // Apps Script APIで月スプシのスクリプトにonEditTriggerを登録
    this._registerTriggerViaApi(newSs.getId());

    Logger.log('月スプシ作成（マスタコピー）: ' + title);
    return newSs;
  },

  /**
   * Apps Script APIを使って月スプシのスクリプトにonEditTriggerを登録する
   * @param {string} spreadsheetId - 月スプシのID
   */
  _registerTriggerViaApi(spreadsheetId) {
    const scriptId = Config.SCRIPT_ID;
    if (!scriptId) {
      Logger.log('SCRIPT_ID が設定されていません');
      return;
    }
    const token = ScriptApp.getOAuthToken();
    const url = `https://script.googleapis.com/v1/projects/${scriptId}/triggers`;
    // Apps Script API v1 の正しいペイロード形式
    const payload = {
      triggerFunction: 'onEditTrigger',
      eventType: 'ON_EDIT',
      service: 'SHEETS',
      resourceId: spreadsheetId,
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(url, options);
    const respText = response.getContentText();
    Logger.log('トリガー登録ステータス: ' + response.getResponseCode());
    Logger.log('トリガー登録レスポンス: ' + respText);
    if (response.getResponseCode() !== 200) {
      Logger.log('トリガー登録失敗 - SCRIPT_ID: ' + scriptId + ' SpreadsheetId: ' + spreadsheetId);
    }
  },

  /**
   * ブランド・日付ごとのシートを取得または作成する
   */
  _getOrCreateDailySheet(ss, brand, date) {
    const sheetName = `【${brand}_${date}】集計`;
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      this._addSummaryHeader(sheet);
    }
    return sheet;
  },

  /**
   * 指定行のG・M〜X列に数式をセットする
   * データ書き込み行にのみ適用（全行一括ではなく1行ずつ）
   *
   * G:  ご利用日（日付型）
   * M〜P: 商品_1（コード・リンク・単価・商品詳細）
   * Q〜T: 商品_2
   * U〜X: 商品_3
   */
  _setRowFormulas(sheet, row) {
    const i = row;

    // G: ご利用日
    const gCell = sheet.getRange(i, 7);
    gCell.setFormula(`=IFERROR(INDEX('CSV貼付'!$E:$E,MATCH(F${i},'CSV貼付'!$A:$A,0)),"")`);
    gCell.setNumberFormat('yyyy/MM/dd');

    // M: 商品コード_1
    sheet.getRange(i, 13).setFormula(`=IFERROR(REGEXREPLACE(INDEX('CSV貼付'!$C:$C,MATCH(F${i},'CSV貼付'!$A:$A,0)),"-[^-]+$",""),"")`);
    // N: リンク_1
    sheet.getRange(i, 14).setFormula(`=IFERROR(XLOOKUP(IFERROR(INDEX('CSV貼付'!$C:$C,MATCH(F${i},'CSV貼付'!$A:$A,0)),""),DB!$A:$A,DB!$C:$C),"")`);
    // O: 単価_1
    sheet.getRange(i, 15).setFormula(`=IFERROR(INDEX('CSV貼付'!$D:$D,MATCH(F${i},'CSV貼付'!$A:$A,0)),"")`);
    // P: 商品詳細_1
    sheet.getRange(i, 16).setFormula(`=IF(O${i}="","",IFERROR(XLOOKUP(O${i},'SET内容'!$A:$A,'SET内容'!$B:$B),""))`);

    // Q: 商品コード_2
    sheet.getRange(i, 17).setFormula(`=IFERROR(REGEXREPLACE(INDEX(FILTER('CSV貼付'!$C:$C,'CSV貼付'!$A:$A=F${i}),2),"-[^-]+$",""),"")`);
    // R: リンク_2
    sheet.getRange(i, 18).setFormula(`=IFERROR(XLOOKUP(IFERROR(INDEX(FILTER('CSV貼付'!$C:$C,'CSV貼付'!$A:$A=F${i}),2),""),DB!$A:$A,DB!$C:$C),"")`);
    // S: 単価_2
    sheet.getRange(i, 19).setFormula(`=IFERROR(INDEX(FILTER('CSV貼付'!$D:$D,'CSV貼付'!$A:$A=F${i}),2),"")`);
    // T: 商品詳細_2
    sheet.getRange(i, 20).setFormula(`=IF(S${i}="","",IFERROR(XLOOKUP(S${i},'SET内容'!$A:$A,'SET内容'!$B:$B),""))`);

    // U: 商品コード_3
    sheet.getRange(i, 21).setFormula(`=IFERROR(REGEXREPLACE(INDEX(FILTER('CSV貼付'!$C:$C,'CSV貼付'!$A:$A=F${i}),3),"-[^-]+$",""),"")`);
    // V: リンク_3
    sheet.getRange(i, 22).setFormula(`=IFERROR(XLOOKUP(IFERROR(INDEX(FILTER('CSV貼付'!$C:$C,'CSV貼付'!$A:$A=F${i}),3),""),DB!$A:$A,DB!$C:$C),"")`);
    // W: 単価_3
    sheet.getRange(i, 23).setFormula(`=IFERROR(INDEX(FILTER('CSV貼付'!$D:$D,'CSV貼付'!$A:$A=F${i}),3),"")`);
    // X: 商品詳細_3
    sheet.getRange(i, 24).setFormula(`=IF(W${i}="","",IFERROR(XLOOKUP(W${i},'SET内容'!$A:$A,'SET内容'!$B:$B),""))`);
  },

  /**
   * G列タイトルを生成して書き込む
   * CSV貼付から直接値を取得するため別ファイルでも動作する
   * @param {Sheet} sheet
   * @param {number} row
   * @param {Spreadsheet} ss - CSV貼付・DB・SET内容を持つスプシ（月スプシまたはマスタ）
   * @param {string} orderCode - 注文コード
   */
  _setRowTitle(sheet, row, ss, orderCode) {
    const i = row;

    // orderCode が渡されない場合はE列から取得（onEditTrigger経由）
    const code = orderCode || sheet.getRange(i, 5).getValue();
    if (!code) return;

    // CSV貼付から注文コードに対応するデータを取得
    const csvSheet = ss ? ss.getSheetByName('CSV貼付') : sheet.getParent().getSheetByName('CSV貼付');
    if (!csvSheet || csvSheet.getLastRow() < 2) return;
    const csvData = csvSheet.getRange(2, 1, csvSheet.getLastRow() - 1, 5).getValues();

    // 注文コードに一致する行を全取得
    const matchRows = csvData.filter(r => r[0] == code);
    if (matchRows.length === 0) return;

    // ご利用日（E列=index4）
    const dateVal = matchRows[0][4];
    if (!dateVal) return;
    const d = new Date(dateVal);
    const dateStr = `${d.getMonth() + 1}月${d.getDate()}日`;

    // 商品コード（C列=index2）から末尾サイズを除去
    const dbSheet = ss ? ss.getSheetByName('DB') : sheet.getParent().getSheetByName('DB');
    const setSheet = ss ? ss.getSheetByName('SET内容') : sheet.getParent().getSheetByName('SET内容');

    const products = matchRows.slice(0, 3).map(r => {
      const fullCode = r[2] || '';
      const shortCode = fullCode.replace(/-[^-]+$/, '');
      const price = r[3] || '';
      // SET内容から商品詳細を取得
      let detail = '';
      if (setSheet && price) {
        const setData = setSheet.getRange(1, 1, setSheet.getLastRow(), 2).getValues();
        const setRow = setData.find(sr => sr[0] == price);
        if (setRow) detail = setRow[1];
      }
      return { code: shortCode, detail };
    }).filter(p => p.code);

    if (products.length === 0) return;

    const codes = products.map(p => p.code);
    const codeStr = codes.join('・');
    const label = this._buildProductLabel(products.map(p => p.detail));

    // AH・AJ列を別途取得（AG=33→AH=34, AI=35→AJ=36、1列ずれ）
    const agAi = sheet.getRange(i, 33, 1, 4).getValues()[0];
    const prefecture = agAi[0]; // AG列
    const areaSub    = agAi[1]; // AH列
    const area  = [prefecture, areaSub].filter(v => v).join('・');
    const usage = agAi[3]; // AJ列

    const title = `${dateStr}　${usage}ご利用　${area}エリア｜${codeStr}（${label}）`;
    sheet.getRange(i, 8).setValue(title); // H列
  },

  /**
   * I列（【お客様の普段の洋服サイズ】）のHTMLを組み立てて書き込む
   * 商品サイズはM列の商品コード末尾（例：CR1-323BU-M → M）から取得
   * マタニティがある場合は p.ub 内に挿入
   */
  _setRowSizeHtml(ss, sheet, row, answers) {
    const i = row;

    // ドレスの商品コードを探す（P/T/X列の商品詳細で「ドレス」を含む行を優先）
    const productVals = sheet.getRange(i, 13, 1, 12).getValues()[0];
    // [M,N,O,P, Q,R,S,T, U,V,W,X] = [0,1,2,3, 4,5,6,7, 8,9,10,11]
    const sets = [
      { code: productVals[0], detail: productVals[3]  }, // 商品_1
      { code: productVals[4], detail: productVals[7]  }, // 商品_2
      { code: productVals[8], detail: productVals[11] }, // 商品_3
    ].filter(p => p.code !== '');

    // ドレスを含む商品コードを優先、なければ最初のコード
    const dressProduct = sets.find(p => p.detail && p.detail.includes('ドレス')) || sets[0];
    const productCodeShort = dressProduct ? dressProduct.code : '';

    // CSV貼付からサイズ付き元コードを取得してサイズ（末尾）を抽出
    let productSize = '';
    if (productCodeShort) {
      const csvSheet = ss.getSheetByName('CSV貼付');
      if (!csvSheet || csvSheet.getLastRow() < 2) return;
      const csvCodes = csvSheet.getRange(2, 3, csvSheet.getLastRow() - 1, 1).getValues().flat();
      const fullCode = csvCodes.find(c => c && c.toString().startsWith(productCodeShort));
      if (fullCode) productSize = fullCode.toString().split('-').pop();
    }

    // CSV貼付けB列の商品名から括弧内のカラーを抽出
    let color = '';
    if (productCodeShort) {
      const csvSheet = ss.getSheetByName('CSV貼付');
      const csvData = csvSheet.getRange(2, 2, csvSheet.getLastRow() - 1, 2).getValues();
      for (const [name, code] of csvData) {
        if (code && code.toString().startsWith(productCodeShort)) {
          const match = name.match(/\(([^)]+)\)/);
          if (match) color = match[1];
          break;
        }
      }
    }

    const age        = answers.age        || '';
    const height     = answers.height     ? `身長${answers.height}cm` : '';
    const maternity  = answers.maternity  ? `マタニティ${answers.maternity}ヶ月` : '';
    const topSize    = answers.topSize    || '';
    const bottomSize = answers.bottomSize || '';
    const bodyType   = answers.bodyType   || '';
    const q3Quality  = this._toStars(answers.q3Quality);
    const q3Comfort  = this._toStars(answers.q3Comfort);
    const q3Size     = answers.q3Size     || '';
    const q3Length   = answers.q3Length   || '';

    // p.ub の中身：年齢｜身長｜[マタニティ]<br>普段サイズ
    const baseItems = [age, height].filter(v => v).join('｜');
    const ubLine = maternity
      ? `${baseItems} ｜${maternity}<br class="sp-only">普段の洋服サイズ：(トップス ) ${topSize}　(ボトムス) ${bottomSize}`
      : `${baseItems} ｜<br class="sp-only">普段の洋服サイズ：(トップス ) ${topSize}　(ボトムス) ${bottomSize}`;

    const sizePart  = productSize ? `商品サイズ：${productSize}` : '';
    const colorPart = color       ? `カラー：${color}`           : '';
    const bodyPart  = bodyType    ? `骨格：${bodyType}`          : '';
    const p2Items   = [sizePart, colorPart, bodyPart].filter(v => v).join('\u3000\u3000｜');

    const p2Line = p2Items ? `${p2Items}<br>` : '';
    const html = `【お客様の普段の洋服サイズ】\n<div class="size-info"><p class="ub">${ubLine}</p><p>${p2Line}品質：${q3Quality}　　　　着心地：${q3Comfort}　　　<br class="sp-only">サイズ感：${q3Size}　　ドレス丈：${q3Length}</p></div>`;

    sheet.getRange(i, 9).setValue(html); // I列
  },

  /**
   * J列（☆お客様がレンタルしたドレスはこちら☆）のHTMLを組み立てて書き込む
   * 商品名はCSV貼付けB列から商品コードで引く
   * 複数商品は改行で結合
   * ブランドによりリンクの文字色が異なる（PDS: #ec407c / OTONA: #222222）
   */
  _setRowDressHtml(ss, sheet, row) {
    const i = row;
    const vals = sheet.getRange(i, 13, 1, 12).getValues()[0];
    // M=0, N=1, O=2, P=3, Q=4, R=5, S=6, T=7, U=8, V=9, W=10, X=11
    const products = [
      { code: vals[0], link: vals[1], price: vals[2],  detail: vals[3]  }, // 商品_1
      { code: vals[4], link: vals[5], price: vals[6],  detail: vals[7]  }, // 商品_2
      { code: vals[8], link: vals[9], price: vals[10], detail: vals[11] }, // 商品_3
    ].filter(p => p.code !== '');

    if (products.length === 0) return;

    // シート名からブランドを判定して文字色を決定
    const isOtona = sheet.getName().includes('OTONA');
    const linkColor = isOtona ? '#222222' : '#ec407c';

    // CSV貼付けからコード→商品名のマップを作成
    const csvSheet = ss.getSheetByName('CSV貼付');
    if (!csvSheet || csvSheet.getLastRow() < 2) return;
    const csvData = csvSheet.getRange(2, 2, csvSheet.getLastRow() - 1, 2).getValues();
    const nameMap = {};
    csvData.forEach(([name, code]) => {
      if (code && !nameMap[code]) {
        // 商品名末尾の商品コード部分を除去
        nameMap[code] = name.replace(new RegExp(code + '$'), '').trim();
      }
    });

    const keywordMap = {
      'ドレス単品':           'ドレス',
      'バッグ単品':           'バッグ',
      'ジャケット単品':       'ジャケット',
      'カーディガン単品':     'カーディガン',
      'ショール単品':         'ショール',
      'ヘアアクセサリー単品': 'ヘアアクセサリー',
      'ベルト単品':           'ベルト',
      'アクセサリー単品':     'アクセサリー',
      'イヤリング単品':       'イヤリング',
      '袱紗単品':             '袱紗',
      'トレンチコート単品':   'トレンチコート',
    };

    const blocks = products.map(p => {
      // nameMapはサイズ付きコードがキーなので前方一致で検索
      const fullCode = Object.keys(nameMap).find(k => k.startsWith(p.code));
      const name = fullCode ? nameMap[fullCode] : p.code;
      const priceStr = Number(p.price).toLocaleString();
      const kind = keywordMap[p.detail] || 'ドレス';
      const title = `☆お客様がレンタルした${kind}はこちら☆`;
      return `${title}\n\n【商品品番】\n${p.code}\n<a href="${p.link}"><span style="color:${linkColor}; font-weight:bold; font-size: 14px; text-decoration: underline;">${name}</span></a>\n\n【レンタル内容】\n${p.detail}\n3泊4日レンタル価格${priceStr}円（税込）`;
    });

    sheet.getRange(i, 10).setValue(blocks.join('\n\n')); // J列
  },

  /**
   * 商品詳細リストから括弧内ラベルを組み立てる
   * ドレスを先頭に、重複排除して結合
   */
  _buildProductLabel(details) {
    const keywordMap = {
      'ドレス単品':           'ドレス',
      'バッグ単品':           'バッグ',
      'ジャケット単品':       'ジャケット',
      'カーディガン単品':     'カーディガン',
      'ショール単品':         'ショール',
      'ヘアアクセサリー単品': 'ヘアアクセサリー',
      'ベルト単品':           'ベルト',
      'アクセサリー単品':     'アクセサリー',
      'イヤリング単品':       'イヤリング',
      '袱紗単品':             '袱紗',
      'トレンチコート単品':   'トレンチコート',
    };

    const filled = details.filter(d => d !== '');

    // セット系（単品以外）はそのまま返す
    if (filled.length === 1 && !keywordMap[filled[0]]) return filled[0];

    // 単品が1つだけ → そのまま
    if (filled.length === 1) return filled[0];

    // 複数 → キーワード抽出、ドレスを先頭に
    const keywords = filled.map(d => keywordMap[d] || d);
    const unique = [...new Set(keywords)];
    const dressParts = unique.filter(k => k === 'ドレス');
    const otherParts = unique.filter(k => k !== 'ドレス');
    return [...dressParts, ...otherParts].join('・');
  },

  /**
   * 品質・着心地の評価を星に変換する
   */
  _toStars(value) {
    const map = {
      'とても良い':     '★★★★★',
      'まあまあ良い':   '★★★★☆',
      '普通':           '★★★☆☆',
      'あまり良くない': '★★☆☆☆',
      'とても良くない': '★☆☆☆☆'
    };
    return map[value] || '';
  },

  /**
   * 集計シートにヘッダー行を追加する
   * 列順: A〜AR
   */
  _addSummaryHeader(sheet) {
    const headers = [
      // A〜C: 手動入力
      '備考', '担当者', '進捗',
      // D〜L: 基本情報・表示用
      '処理日時', 'ファイル名', '注文コード',
      'ご利用日', 'タイトル', '【お客様の普段の洋服サイズ】',
      '☆お客様がレンタルしたドレスはこちら☆', '着用頂いた感想・サイズ感・レンタルの流れについてのご意見', 'AI返信文',
      // M〜P: 商品_1（数式）
      '商品コード_1', 'リンク_1', '単価_1', '商品詳細_1',
      // Q〜T: 商品_2（数式）
      '商品コード_2', 'リンク_2', '単価_2', '商品詳細_2',
      // U〜X: 商品_3（数式）
      '商品コード_3', 'リンク_3', '単価_3', '商品詳細_3',
      // Y〜AR: アンケート回答
      '年齢', '普段の洋服のサイズ', 'トップスサイズ', 'ボトムスサイズ',
      '身長(cm)', 'マタニティ', '骨格タイプ', '利用エリア', '都道府県', 'エリア',
      'どこで知りましたか？', '今回のご利用用途', 'ご利用商品はいかがでしたか？',
      '品質', '着心地', '※あまり良くないと答えた方',
      'サイズ感', 'ドレス丈/パンツ丈', 'レンタルの流れはいかがでしたか？', '※不便と答えた人用'
    ];
    sheet.appendRow(headers);

    // 全体: 太字
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    // A〜C: 備考・担当者・進捗 黄色
    sheet.getRange(1, 1,  1, 3).setBackground('#FFF2CC');
    // D〜L: 基本情報 青
    sheet.getRange(1, 4,  1, 9).setBackground('#D9E8FC');
    // M〜P: 商品_1 水色
    sheet.getRange(1, 13, 1, 4).setBackground('#C9E7F5');
    // Q〜T: 商品_2 緑
    sheet.getRange(1, 17, 1, 4).setBackground('#C9F5D6');
    // U〜X: 商品_3 オレンジ
    sheet.getRange(1, 21, 1, 4).setBackground('#FCE5CD');
    // Y〜AR: アンケート回答 青
    sheet.getRange(1, 25, 1, headers.length - 24).setBackground('#D9E8FC');
  }
};
