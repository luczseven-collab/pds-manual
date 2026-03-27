/**
 * ImageUtils.gs — 商品画像管理ユーティリティ
 *
 * 画像DB（マスター）とお客様別画像フォルダへの保存を担当する。
 *
 * 処理フロー：
 * 1. 画像DBに商品画像があればそれを使う
 * 2. なければ商品ページから取得して画像DBに保存
 * 3. 画像DBからお客様別画像フォルダにコピー（連番付き）
 */

const ImageUtils = {

  /**
   * アンケート処理後に呼び出すメインメソッド。
   * 画像DBへの保存とお客様別画像へのコピーを行う。
   * @param {Object} analyzed - 処理済みアンケートデータ（orderCode, fileName, productUrl を含む）
   */
  saveImages(analyzed) {
    const imageDbFolderId       = Config.DRIVE_FOLDER_IMAGE_DB_ID;
    const customerImageFolderId = Config.DRIVE_FOLDER_CUSTOMER_IMAGE_ID;

    if (!imageDbFolderId || !customerImageFolderId) {
      Logger.log('ImageUtils: フォルダIDが未設定のためスキップ');
      return;
    }

    const orderCode = analyzed.orderCode;
    if (!orderCode) {
      Logger.log('ImageUtils: orderCode がないためスキップ');
      return;
    }

    // 集計シートの横並び列から商品コード・リンクを取得（最大3商品）
    // M=13:商品コード_1, N=14:リンク_1, Q=17:商品コード_2, R=18:リンク_2, U=21:商品コード_3, V=22:リンク_3
    const { sheet, row } = analyzed;
    const products = [];
    [[14, 15], [18, 19], [22, 23]].forEach(([codeCol, urlCol]) => {
      const code = sheet ? String(sheet.getRange(row, codeCol).getValue()) : '';
      const url  = sheet ? String(sheet.getRange(row, urlCol).getValue())  : '';
      if (code && url && code !== 'undefined') products.push({ productCode: code, productUrl: url });
    });

    if (products.length === 0) {
      Logger.log('ImageUtils: 集計シートに商品情報が見つかりません → row ' + row);
      return;
    }

    const scanFileName = analyzed.fileName.replace(/\.[^.]+$/, ''); // 拡張子を除く

    products.forEach(({ productCode, productUrl }) => {
      try {
        const dbFile = this._getOrFetchImageDbFile(productCode, productUrl, imageDbFolderId);
        if (!dbFile) return;
        this._copyToCustomerFolder(dbFile, scanFileName, productCode, customerImageFolderId);
      } catch (e) {
        Logger.log('ImageUtils ERROR (' + productCode + '): ' + e.message);
      }
    });
  },

  /**
   * 画像DBに画像があればそれを返す。なければ商品ページから取得して保存して返す。
   */
  _getOrFetchImageDbFile(productCode, productUrl, imageDbFolderId) {
    const codeShort = this._stripSize(productCode);
    const fileName  = codeShort + '.jpg';
    const folder    = DriveApp.getFolderById(imageDbFolderId);

    // 既存チェック
    const existing = folder.getFilesByName(fileName);
    if (existing.hasNext()) {
      Logger.log('ImageUtils: 画像DB既存 → ' + fileName);
      return existing.next();
    }

    // 商品ページから画像URLを取得
    const imageUrl = this._fetchProductImageUrl(productUrl);
    if (!imageUrl) return null;

    // 画像を取得してDriveに保存
    const imgResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    if (imgResponse.getResponseCode() !== 200) {
      Logger.log('ImageUtils: 画像取得失敗 (' + imgResponse.getResponseCode() + ') → ' + imageUrl);
      return null;
    }

    const blob = imgResponse.getBlob().setName(fileName);
    const file = folder.createFile(blob);
    Logger.log('ImageUtils: 画像DB保存 → ' + fileName);
    return file;
  },

  /**
   * 商品ページHTMLからメイン画像URLを取得する。
   */
  _fetchProductImageUrl(productUrl) {
    const response = UrlFetchApp.fetch(productUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('ImageUtils: 商品ページ取得失敗 → ' + productUrl);
      return null;
    }
    const match = response.getContentText().match(/src="(\/upload\/save_image\/[^"]+)"/);
    if (!match) {
      Logger.log('ImageUtils: メイン画像URLが見つかりません → ' + productUrl);
      return null;
    }
    return 'https://partydressstyle.jp' + match[1];
  },

  /**
   * 画像DBのファイルをお客様別画像フォルダに連番付きでコピーする。
   * ファイル名例: MX-5150FN_20260326_0001_K1-233WR_01.jpg
   */
  _copyToCustomerFolder(sourceFile, scanFileName, codeShort, customerImageFolderId) {
    const folder   = DriveApp.getFolderById(customerImageFolderId);
    const baseName = scanFileName + '_' + codeShort;

    // 連番を決定（_01から開始）
    let index = 1;
    while (folder.getFilesByName(baseName + '_' + String(index).padStart(2, '0') + '.jpg').hasNext()) {
      index++;
    }

    const newFileName = baseName + '_' + String(index).padStart(2, '0') + '.jpg';
    sourceFile.makeCopy(newFileName, folder);
    Logger.log('ImageUtils: お客様別画像コピー → ' + newFileName);
  },

  /**
   * 受注番号からCSV貼付シートで商品コードを取得し、DBシートで商品URLを返す。
   * CSV貼付シート構成: A列=受注番号、C列=商品コード
   * DBシート構成: A列=商品コード、B列=セット内容、C列=リンク
   */
  _getProductUrl(orderCode) {
    const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);

    // 受注番号から商品コードを取得
    const csvSheet = ss.getSheetByName('CSV貼付');
    if (!csvSheet || csvSheet.getLastRow() < 2) return null;

    const csvData = csvSheet.getRange(2, 1, csvSheet.getLastRow() - 1, 3).getValues();
    const csvRow = csvData.find(r => String(r[0]) === String(orderCode));
    if (!csvRow) {
      Logger.log('ImageUtils: CSV貼付に受注番号が見つかりません → ' + orderCode);
      return null;
    }
    const productCode = String(csvRow[2]);

    // 商品コードからDBシートでURLを取得
    const dbSheet = ss.getSheetByName('DB');
    if (!dbSheet || dbSheet.getLastRow() < 2) return null;

    const codeShort = this._stripSize(productCode);
    const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 3).getValues();
    const dbRow = dbData.find(r => this._stripSize(String(r[0])) === codeShort);
    return dbRow ? dbRow[2] : null;
  },

  /**
   * 受注番号からCSV貼付シートで商品コードを返す。
   */
  _getProductCode(orderCode) {
    const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
    const csvSheet = ss.getSheetByName('CSV貼付');
    if (!csvSheet || csvSheet.getLastRow() < 2) return null;
    const csvData = csvSheet.getRange(2, 1, csvSheet.getLastRow() - 1, 3).getValues();
    const row = csvData.find(r => String(r[0]) === String(orderCode));
    return row ? String(row[2]) : null;
  },

  /**
   * 商品コードの末尾サイズを除去する。
   * 例: K1-233WR-S → K1-233WR、CL1-453C-GRY-S → CL1-453C-GRY
   */
  _stripSize(code) {
    return String(code).replace(/-(XS|S|M|L|XL|XXL|[0-9]+)$/i, '');
  },
};
