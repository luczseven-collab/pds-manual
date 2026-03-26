# セットアップ手順書

## 1. Google Cloud プロジェクト設定

1. Google Cloud Consoleで新規プロジェクトを作成
   - URL: https://console.cloud.google.com/
2. Cloud Vision API を有効化
   - URL: https://console.cloud.google.com/apis/library/vision.googleapis.com
3. APIキーを作成（Cloud Vision API用）
   - URL: https://console.cloud.google.com/apis/credentials
   1. 上記URLを開き、画面上部の「+ 認証情報を作成」をクリック
   2. 「APIキー」を選択
   3. 作成されたAPIキーをコピーしておく（後でスクリプトプロパティに設定）
   4. 「キーを制限」をクリックしてセキュリティ設定を行う
      - 「APIの制限」→「キーを制限」を選択
      - 「Cloud Vision API」にチェックを入れて保存
   > APIキーを制限することで、万が一漏洩した際の被害を最小限に抑えられる

## 2. Google Apps Script プロジェクト作成

1. Google Sheetsを新規作成（これがメインのスプレッドシートになる）
   - URL: https://sheets.new
2. 拡張機能 → Apps Script でエディタを開く
   - または直接: https://script.google.com/
3. `src/gas/` 配下のファイルをすべてコピー
   - `appsscript.json`（マニフェスト）
   - `Config.gs`
   - `Code.gs`
   - `OcrUtils.gs`
   - `CoordMapper.gs`
   - `SheetsUtils.gs`
   - `AiUtils.gs`
   - `NotifyUtils.gs`

## 3. スクリプトプロパティの設定

GASエディタ → プロジェクトの設定 → スクリプトプロパティ に以下を追加：

> Claude APIキーの取得: https://console.anthropic.com/settings/keys
> Chatwork APIトークンの取得: https://www.chatwork.com/service/packages/chatwork/subpackages/api/token.php

| プロパティ名 | 値 |
|------------|---|
| `DRIVE_FOLDER_ID` | 監視するGoogle DriveフォルダのID |
| `SPREADSHEET_ID` | メインスプレッドシートのID |
| `CLAUDE_API_KEY` | Claude APIキー |
| `VISION_API_KEY` | Cloud Vision APIキー |
| `CHATWORK_API_TOKEN` | ChatworkのAPIトークン |
| `CHATWORK_ROOM_ID` | 通知先ChatworkルームID |
| `NOTIFY_EMAIL` | エラー通知先メールアドレス |

## 4. スプレッドシートの初期シート作成

以下のシートを手動で作成する：

- `【累計】集計`（ヘッダーはSheetsUtils._addSummaryHeaderで自動追加）
- `【意見まとめ】`
- `【意見分析】`
- `【設定】トンマナ`
- `【設定】座標マップ`
- `【ログ】処理履歴`

## 5. 【設定】座標マップシートの設定

アンケート実物をスキャンし、各設問の選択肢座標を計測して入力する。

| カラム | 内容 |
|-------|------|
| A列 | 設問キー（例: `age`, `q1`, `freeComment`） |
| B列 | 表示名 |
| C列 | タイプ（`select` or `text`） |
| D列 | 選択肢（`/` 区切り） |
| E列 | 座標定義（JSON形式） |

座標定義のJSON形式例（選択式）：
```json
[
  {"x1": 100, "y1": 200, "x2": 150, "y2": 230},
  {"x1": 200, "y1": 200, "x2": 250, "y2": 230}
]
```

## 6. 【設定】トンマナシートの設定

| A列（設定項目） | B列（値） |
|--------------|---------|
| ブランド名・署名 | PARTY DRESS STYLE スタッフ一同 |
| 回答の口調・文体 | 丁寧語・敬語（です・ます調） |
| 共感フレーズ（固定） | この度はご利用いただき誠にありがとうございます。 |
| 禁止語・NGワード | （競合他社名等） |
| 回答文字数目安 | 100〜150文字程度 |
| ネガティブ時の追記 | 担当者より改めてご連絡させていただきます。 |

## 7. タイマートリガーの設定

GASエディタ → トリガー → トリガーを追加：
- 実行する関数: `checkNewFiles`
- イベントのソース: 時間主導型
- 時間ベースのトリガーのタイプ: 分ベースのタイマー
- 時間の間隔: 10分おき

## 8. 動作確認

1. 監視フォルダにテスト用スキャン画像（JPEG）をアップロード
2. 10分以内に処理が実行されることを確認
3. 集計シートとログシートに書き込まれていることを確認
