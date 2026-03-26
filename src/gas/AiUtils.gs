/**
 * AiUtils.gs — Claude API連携
 * 画像からアンケート回答を読み取り、感情分析・AI回答生成・HTML生成を行う
 */

const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODEL = 'claude-opus-4-6';

const AiUtils = {
  /**
   * アンケート画像からすべての回答をJSONで読み取る（ブランド判定含む）
   * @param {Object} imageData - { base64: string, mimeType: string }
   * @returns {Object} 設問ごとの回答マップ（brand: 'PDS' | 'OTONA' を含む）
   */
  readAnswersFromImage(imageData) {
    const prompt = `あなたは紙アンケートの読み取り専門家です。画像のアンケート用紙を上から順に丁寧に読み取り、JSONで返してください。

回答方法は「選択肢の文字を手書きの○で囲む」形式です。○は小さく薄い場合があります。

【ブランド判定】
用紙上部のロゴを見てください。
- 「PARTY DRESS STYLE」というロゴなら brand: "PDS"
- 「OTONA DRESS」というロゴなら brand: "OTONA"

アンケートの構成（上から順）：

[年齢] 横一列: 20歳以下 / 21歳~25歳 / 26歳~30歳 / 31歳~35歳 / 36歳~40歳 / 41歳以上
→ key: "age"

[普段の洋服サイズ]
・トップス: S / M / L / XL~  → key: "topSize"
・ボトムス: S / M / L / XL~  → key: "bottomSize"
・身長(手書き数値)           → key: "height"
・マタニティ(手書き月数)      → key: "maternity"

[骨格タイプ] ストレート / ナチュラル / ウェーブ → key: "bodyType"

[ご利用エリア] 手書き都道府県 → key: "prefecture" / 手書きエリア → key: "area"

[質問1] 「PARTY DRESS STYLEを何で知りましたか?」
横一列: web検索 / 友人紹介 / WeddingWEB招待状 / Instagram / リピート
→ key: "q1"

[質問2] 「今回ご利用の用途は何でしたか?」
横一列: 結婚式 / 結婚式二次会 / 謝恩会 / 食事会 / その他
→ key: "q2"

[質問3] 「ご利用商品はいかがでしたか?」※4行それぞれ独立して読む
・●品質    行: [1]とても良い / [2]まあまあ良い / [3]普通 / [4]あまり良くない / [5]とても良くない → key: "q3Quality"
  ※左端「とても良い」の文字の周囲に薄い○がある場合も必ず検出してください
・●着心地  行: [1]とても良い / [2]まあまあ良い / [3]普通 / [4]あまり良くない / [5]とても良くない → key: "q3Comfort"
  ※左端「とても良い」の文字の周囲に薄い○がある場合も必ず検出してください
・手書き理由（あまり良くない・とても良くない時）                                    → key: "q3NgReason"
・●サイズ感行: 小さい / やや小さい / 丁度いい / やや大きい / 大きい               → key: "q3Size"
・●ドレス丈行: ひざ上 / ひざ下 / ふくらはぎ / くるぶし上 / くるぶし下             → key: "q3Length"

[質問4] 「レンタルの流れはいかがでしたか?」
横一列: わかりやすい / 便利 / 普通 / わかりづらい / 不便
→ key: "q4Flow"
・手書き理由（わかりづらい・不便の時） → key: "q4NgReason"

[自由意見欄] 手書きテキスト → key: "freeComment"

ルール：
- ○で囲まれた選択肢を返す。未選択・未記入は空文字("")
- 値は上記の選択肢から返す（手書き数値・テキスト除く）
- web検索は小文字で記載されている場合も"web検索"を返す

【返すJSONの形式】
{
  "brand": "PDS",
  "age": "",
  "topSize": "",
  "bottomSize": "",
  "height": "手書きの数値のみ（例: 158）",
  "maternity": "手書きの月数（数値のみ）",
  "bodyType": "",
  "prefecture": "手書きの都道府県名のみ（例: 東京都）",
  "area": "手書きのエリア名のみ（例: 表参道）",
  "q1": "",
  "q2": "",
  "q3Quality": "",
  "q3Comfort": "",
  "q3NgReason": "品質・着心地でNG回答の場合の手書き理由テキスト",
  "q3Size": "",
  "q3Length": "",
  "q4Flow": "",
  "q4NgReason": "レンタルの流れでNG回答の場合の手書き理由テキスト",
  "freeComment": "着用頂いた感想・サイズ感・レンタルの流れについての手書きテキスト"
}

JSONのみを返してください。説明文・コードブロックは不要です。`;

    const response = UrlFetchApp.fetch(CLAUDE_API_URL, {
      method: 'post',
      headers: {
        'x-api-key': Config.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify({
        model: CLAUDE_MODEL,
        max_tokens: 1024,
        messages: [{
          role: 'user',
          content: [
            {
              type: 'image',
              source: {
                type: 'base64',
                media_type: imageData.mimeType,
                data: imageData.base64
              }
            },
            { type: 'text', text: prompt }
          ]
        }]
      })
    });

    const result = JSON.parse(response.getContentText());
    Logger.log('stop_reason: ' + result.stop_reason);
    const text = result.content[0].text.trim();
    Logger.log('Claude読み取り結果: ' + text);

    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('Claude APIからJSONが返りませんでした: ' + text);
    return JSON.parse(jsonMatch[0]);
  },

  /**
   * 自由意見の感情分析を実施する
   * @param {Object} answers
   * @returns {Object} sentiment フィールドを追加した回答マップ
   */
  analyzeSentiment(answers) {
    const opinion = answers.freeComment || '';
    if (!opinion) return { ...answers, sentiment: '', aiReply: '' };

    const prompt = `以下はドレスレンタルサービスの顧客アンケートの自由意見です。
感情を「ポジティブ」「ニュートラル」「ネガティブ」の3つから1つだけ選んで返してください。

意見：${opinion}`;

    const sentiment = this._callClaude(prompt).trim();
    return { ...answers, sentiment };
  },

  /**
   * トンマナ設定に基づいてAI返信文を生成する
   * @param {Object} answers - sentiment含む回答マップ
   * @returns {string} AI返信文
   */
  generateReply(answers) {
    const opinion = answers.freeComment || '';
    if (!opinion) return '';

    const toneSettings = this._loadToneSettings(answers.brand);
    const prompt = `あなたは「${toneSettings.brandName}」のスタッフです。
以下の設定に従って、顧客の意見に対する返信文を生成してください。

【設定】
- 冒頭固定文: ${toneSettings.openingText}
  ※冒頭固定文をそのまま文頭に使用し、その後に続く文章では「この度は」「ありがとうございます」などの重複表現を使わないこと
- 口調: ${toneSettings.tone}
- NGワード: ${toneSettings.ngWords}
- 文字数目安: ${toneSettings.wordCount}（冒頭固定文を含む）
${answers.sentiment === 'ネガティブ' ? `- ネガティブ対応: 顧客の不満に対して共感・謝罪を自然な流れで組み込んでください。定型文をそのまま使わず、意見の内容に合わせた言葉で表現してください。参考フレーズ：「${toneSettings.negativeNote}」` : ''}

【顧客の意見】
${opinion}

返信文のみを返してください。`;

    return this._callClaude(prompt);
  },

  /**
   * アンケート回答からHTMLスニペットを生成する
   * @param {Object} answers - 回答マップ
   * @returns {string} HTMLスニペット
   */
  generateHtml(answers) {
    const age = answers.age || '';
    const height = answers.height ? `身長${answers.height}cm` : '';
    const maternity = answers.maternity ? `マタニティ${answers.maternity}ヶ月` : '';
    const topSize = answers.topSize || '';
    const bottomSize = answers.bottomSize || '';
    const bodyType = answers.bodyType || '';
    const q3Quality = this._toStars(answers.q3Quality);
    const q3Comfort = this._toStars(answers.q3Comfort);
    const q3Size = answers.q3Size || '';
    const q3Length = answers.q3Length || '';

    const sizeLineItems = [age, height, maternity].filter(v => v);
    const sizeLine = sizeLineItems.join('｜');
    const clothSizeLine = `普段の洋服サイズ：(トップス) ${topSize}　(ボトムス) ${bottomSize}`;

    const html =
`<div class="size-info">
<p class="ub">${sizeLine}｜<br class="sp-only">${clothSizeLine}</p>
<p>骨格：${bodyType}
品質：${q3Quality}　　着心地：${q3Comfort}　　<br class="sp-only">サイズ感：${q3Size}　　ドレス丈：${q3Length}</p>
</div>`;

    return html;
  },

  /**
   * 品質・着心地の評価を星に変換する
   * @param {string} value
   * @returns {string}
   */
  _toStars(value) {
    const map = {
      'とても良い':    '★★★★★',
      'まあまあ良い':  '★★★★☆',
      '普通':          '★★★☆☆',
      'あまり良くない':'★★☆☆☆',
      'とても良くない':'★☆☆☆☆'
    };
    return map[value] || '';
  },

  /**
   * 意見分析レポートを生成する
   * @param {Array} opinions
   * @returns {string}
   */
  generateAnalysisReport(opinions) {
    const text = opinions.join('\n');
    const prompt = `以下はドレスレンタルサービスのアンケート自由意見一覧です。
以下の4点を分析してください：
1. 頻出キーワード（上位10件）
2. テーマ分類（カテゴリごとの意見数）
3. 感情割合（ポジティブ/ニュートラル/ネガティブ）
4. 改善提案（具体的なアクション）

意見一覧：
${text}`;
    return this._callClaude(prompt);
  },

  /**
   * 【設定】トンマナシートから設定を読み込む
   * @param {string} brand - 'PDS' | 'OTONA'
   */
  _loadToneSettings(brand) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = brand === 'OTONA' ? '【設定】トンマナ_OTONA' : '【設定】トンマナ_PDS';
    const sheet = ss.getSheetByName(sheetName) || ss.getSheetByName('【設定】トンマナ_PDS') || ss.getSheetByName('【設定】トンマナ');
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (const row of data.slice(1)) map[row[0]] = row[1];
    return {
      brandName:    map['ブランド名・署名'] || '',
      openingText:  map['冒頭固定文'] || '',
      tone:         map['回答の口調・文体'] || '',
      ngWords:      map['禁止語・NGワード'] || '',
      wordCount:    map['回答文字数目安'] || '100〜150文字',
      negativeNote: map['ネガティブ時の追記'] || ''
    };
  },

  /**
   * Claude APIテキスト呼び出し
   */
  _callClaude(prompt) {
    const response = UrlFetchApp.fetch(CLAUDE_API_URL, {
      method: 'post',
      headers: {
        'x-api-key': Config.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify({
        model: CLAUDE_MODEL,
        max_tokens: 1024,
        messages: [{ role: 'user', content: prompt }]
      })
    });
    const result = JSON.parse(response.getContentText());
    return result.content[0].text;
  }
};
