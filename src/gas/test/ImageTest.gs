/**
 * ImageTest.gs — 画像管理機能の検証スクリプト
 *
 * 本番と同じ動作で以下を検証する：
 * 1. 商品ページからメイン画像URLを取得
 * 2. サイズ除去（商品コード末尾の -S/-M/-L 等を除く）
 * 3. 画像DBフォルダへの保存（重複スキップ）
 * 4. お客様別画像フォルダへのコピー（連番付き）
 *
 * 実行方法：各テスト関数をGASエディタから個別に実行する
 * 検証完了後このファイルは削除する
 */

// ---- テスト設定 ----
const TEST_PRODUCT_URL  = 'https://partydressstyle.jp/products/detail.php?product_id=142';
const TEST_PRODUCT_CODE = 'K1-233WR-S';   // サイズ付き商品コード
const TEST_SCAN_FILE    = 'MX-5150FN_20260326_202503_0001'; // スキャンファイル名（拡張子なし）

// スクリプトプロパティから取得（事前に設定が必要）
// DRIVE_FOLDER_IMAGE_DB_ID     → 画像DBフォルダID
// DRIVE_FOLDER_CUSTOMER_IMAGE_ID → お客様別画像フォルダID

// ---- ユーティリティ関数（本番実装と同じロジック） ----

/**
 * 商品コードの末尾サイズを除去する
 * 例: K1-233WR-S → K1-233WR
 *     CR1-406GR-M → CR1-406GR
 *     AB3-438GRY-XL → AB3-438GRY
 */
function _stripSize(code) {
  return code.replace(/-[A-Z0-9]+$/i, '');
}

/**
 * 商品ページHTMLからメイン画像のURLを取得する
 * /upload/save_image/ を含む最初のsrc属性を抽出する
 */
function _fetchProductImageUrl(productUrl) {
  const response = UrlFetchApp.fetch(productUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('商品ページ取得失敗: ' + response.getResponseCode());
  }
  const html = response.getContentText();
  const match = html.match(/src="(\/upload\/save_image\/[^"]+)"/);
  if (!match) {
    throw new Error('メイン画像URLが見つかりません');
  }
  return 'https://partydressstyle.jp' + match[1];
}

/**
 * 画像URLからblobを取得してDriveフォルダに保存する
 * 同名ファイルが存在する場合はスキップしてそのファイルを返す
 */
function _saveImageToFolder(imageUrl, fileName, folderId) {
  const folder = DriveApp.getFolderById(folderId);

  // 既存ファイルチェック
  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    Logger.log('スキップ（既存）: ' + fileName);
    return existing.next();
  }

  // 画像取得
  const imgResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
  if (imgResponse.getResponseCode() !== 200) {
    throw new Error('画像取得失敗: ' + imageUrl);
  }

  // Drive保存
  const blob = imgResponse.getBlob().setName(fileName);
  const file = folder.createFile(blob);
  Logger.log('保存完了: ' + fileName + ' (' + file.getId() + ')');
  return file;
}

/**
 * お客様別画像フォルダに連番付きでコピーする
 * 例: scan001_K1-233WR_01.jpg, _02.jpg ...
 */
function _copyToCustomerFolder(sourceFile, scanFileName, productCodeShort, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const baseName = scanFileName + '_' + productCodeShort;

  // 既存の連番ファイル数を確認
  let index = 1;
  while (folder.getFilesByName(baseName + '_' + String(index).padStart(2, '0') + '.jpg').hasNext()) {
    index++;
  }

  const newFileName = baseName + '_' + String(index).padStart(2, '0') + '.jpg';
  const copied = sourceFile.makeCopy(newFileName, folder);
  Logger.log('コピー完了: ' + newFileName + ' (' + copied.getId() + ')');
  return copied;
}

// ---- テスト関数 ----

/**
 * テスト1: サイズ除去の動作確認
 */
function testStripSize() {
  const cases = [
    { input: 'K1-233WR-S',    expected: 'K1-233WR'    },
    { input: 'CR1-406GR-M',   expected: 'CR1-406GR'   },
    { input: 'AB3-438GRY-XL', expected: 'AB3-438GRY'  },
    { input: 'H1-461CHC-M',   expected: 'H1-461CHC'   },
    { input: 'CL1-453C-GRY-S',expected: 'CL1-453C-GRY'},
  ];

  let allPassed = true;
  cases.forEach(({ input, expected }) => {
    const result = _stripSize(input);
    const passed = result === expected;
    Logger.log((passed ? '✓' : '✗') + ' ' + input + ' → ' + result + (passed ? '' : ' (期待値: ' + expected + ')'));
    if (!passed) allPassed = false;
  });

  Logger.log(allPassed ? '全テスト通過' : '失敗あり');
}

/**
 * テスト2: 商品ページからメイン画像URLを取得
 */
function testFetchProductImageUrl() {
  Logger.log('商品ページ: ' + TEST_PRODUCT_URL);
  const imageUrl = _fetchProductImageUrl(TEST_PRODUCT_URL);
  Logger.log('取得した画像URL: ' + imageUrl);
}

/**
 * テスト3: 画像DBへの保存（実際にDriveに保存される）
 */
function testSaveToImageDb() {
  const imageDbFolderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_IMAGE_DB_ID');
  if (!imageDbFolderId) {
    Logger.log('ERROR: DRIVE_FOLDER_IMAGE_DB_ID が設定されていません');
    return;
  }

  const codeShort  = _stripSize(TEST_PRODUCT_CODE);
  const fileName   = codeShort + '.jpg';
  const imageUrl   = _fetchProductImageUrl(TEST_PRODUCT_URL);

  Logger.log('商品コード（サイズなし）: ' + codeShort);
  Logger.log('ファイル名: ' + fileName);
  Logger.log('画像URL: ' + imageUrl);

  _saveImageToFolder(imageUrl, fileName, imageDbFolderId);
  Logger.log('テスト3完了');
}

/**
 * テスト4: お客様別画像へのコピー（実際にDriveにコピーされる）
 * ※テスト3を先に実行して画像DBに画像が存在する状態で実行すること
 */
function testCopyToCustomerFolder() {
  const imageDbFolderId      = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_IMAGE_DB_ID');
  const customerImageFolderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_CUSTOMER_IMAGE_ID');

  if (!imageDbFolderId || !customerImageFolderId) {
    Logger.log('ERROR: フォルダIDが設定されていません');
    return;
  }

  const codeShort = _stripSize(TEST_PRODUCT_CODE);
  const fileName  = codeShort + '.jpg';

  // 画像DBからファイルを取得
  const folder   = DriveApp.getFolderById(imageDbFolderId);
  const files    = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    Logger.log('ERROR: 画像DBに ' + fileName + ' が存在しません。テスト3を先に実行してください。');
    return;
  }
  const sourceFile = files.next();

  _copyToCustomerFolder(sourceFile, TEST_SCAN_FILE, codeShort, customerImageFolderId);

  // 2枚目も試す（連番テスト）
  _copyToCustomerFolder(sourceFile, TEST_SCAN_FILE, codeShort, customerImageFolderId);

  Logger.log('テスト4完了 — お客様別画像フォルダを確認してください');
}

/**
 * テスト5: 全体フロー（1〜4を一括実行）
 */
function testFullFlow() {
  Logger.log('=== 全体フロー検証開始 ===');
  testStripSize();
  Logger.log('---');
  testFetchProductImageUrl();
  Logger.log('---');
  testSaveToImageDb();
  Logger.log('---');
  testCopyToCustomerFolder();
  Logger.log('=== 全体フロー検証完了 ===');
}
