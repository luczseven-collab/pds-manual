# 画像管理機能 実装計画書

作成日: 2026-03-26

---

## 概要

アンケート処理時に、レンタル商品の画像を自動でGoogle Driveに保存する機能。
商品画像のマスターDB（画像DB）と、お客様ごとの画像（お客様別画像）を管理する。

---

## Driveフォルダ構成

```
📁 画像DB             ← 商品画像マスター（重複保存しない）
📁 お客様別画像        ← お客様ごとの画像（スキャンファイル名＋商品コード）
📁 掲載OK
📁 掲載NG
```

---

## 仕様詳細

### 1. 画像DB（マスター保存）

| 項目 | 内容 |
|------|------|
| 取得元 | DBシート「リンク」列のURL（`detail.php?product_id=XXX`） |
| 取得画像 | 商品ページのメイン画像（1枚目）|
| HTML構造 | `<li><a href="#"><img src="/upload/save_image/...">` の最初のimgタグ |
| ファイル名 | 商品コードからサイズ部分を除去 → `K1-233WR.jpg` |
| サイズパターン | `-S`, `-M`, `-L`, `-XL`, `-LL`, `-F` 等（末尾の `-[A-Z]+` を除去） |
| 重複時 | 同名ファイルが存在する場合はスキップ |
| エラー時 | ログ記録してスキップ（処理を止めない） |

### 2. お客様別画像

| 項目 | 内容 |
|------|------|
| 取得元 | 画像DBからコピー |
| 保存先 | `お客様別画像` フォルダ（フラット構造） |
| ファイル名 | `{スキャンファイル名}_{商品コード(サイズなし)}_01.jpg` |
| 連番 | 1枚目から `_01`、2枚目 `_02`... |
| 保存タイミング | `processAnketFile()` 内で自動実行 |
| エラー時 | ログ記録してスキップ |

**ファイル名例：**
```
scan001_K1-233WR_01.jpg   ← 1枚目
scan001_K1-233WR_02.jpg   ← 2枚目（同じお客さんで同じ商品が複数の場合）
scan002_CR1-406GR_01.jpg  ← 別のお客さん
```

---

## 実装ステップ

### Phase 0: 事前準備（実装前）
- [ ] Google Driveに `画像DB` フォルダを作成してIDを控える
- [ ] Google Driveに `お客様別画像` フォルダを作成してIDを控える
- [ ] サイズ除去の正規表現パターンを確定する（`/-[A-Z]+$/i`）

### Phase 1: Config.gs の拡張
- [ ] `DRIVE_FOLDER_IMAGE_DB_ID` をスクリプトプロパティに追加
- [ ] `DRIVE_FOLDER_CUSTOMER_IMAGE_ID` をスクリプトプロパティに追加
- [ ] `Config` オブジェクトに getter を追加

### Phase 2: ImageUtils.gs の新規作成
- [ ] `_stripSize(code)` — 商品コードからサイズを除去する
- [ ] `_fetchProductImageUrl(productUrl)` — 商品ページからメイン画像URLを取得する
- [ ] `_getOrFetchImageDbFile(productCode, productUrl)` — 画像DBに保存（既存ならスキップ）
- [ ] `saveToImageDb(productCode, productUrl)` — 画像DB保存のエントリポイント
- [ ] `saveToCustomerFolder(fileName, productCode)` — お客様別画像へのコピー

### Phase 3: Code.gs への組み込み
- [ ] `processAnketFile()` 内の `writeLog('DONE')` 前に画像保存処理を追加
- [ ] `analyzed.orderCode` と `analyzed.fileName` を使って `ImageUtils.saveToCustomerFolder()` を呼び出す

### Phase 4: 一括同期ツール（初回セットアップ用）
- [ ] `syncAllProductImages()` — DBシート全商品の画像を一括で画像DBに取り込む
- [ ] 100件ごとのバッチ処理（6分タイムアウト対策）

### Phase 5: テスト・確認
- [ ] `testSaveToImageDb()` — 単体テスト（1商品の画像DB保存）
- [ ] `testSaveToCustomerFolder()` — 単体テスト（お客様別画像コピー）
- [ ] 連番の動作確認
- [ ] 既存ファイルスキップの動作確認
- [ ] エラー時スキップの動作確認

---

## 新規ファイル

| ファイル | 役割 |
|---------|------|
| `src/gas/ImageUtils.gs` | 画像取得・保存・コピーのユーティリティ |

## 変更ファイル

| ファイル | 変更内容 |
|---------|---------|
| `src/gas/Config.gs` | フォルダID2件を追加 |
| `src/gas/Code.gs` | `processAnketFile()` に画像保存を追加・テスト関数追加 |

---

## GAS制約・注意点

| 制約 | 対応 |
|------|------|
| 実行時間6分上限 | 一括同期は100件バッチに分割 |
| UrlFetchApp レスポンス上限50MB | 商品画像1枚は問題なし |
| 商品ページ画像URLが相対パス | `https://partydressstyle.jp` をプレフィックスとして付与 |
| onEditトリガー30秒上限 | 画像保存はprocessAnketFile内（タイマートリガー）で実行するため問題なし |
