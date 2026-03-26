/**
 * CoordMapper.gs — 座標ベース選択肢判定
 * Vision APIのboundingBox座標を使って選択された項目を判定する
 *
 * 判定方針:
 * - ○で囲まれた文字はboundingBoxの高さ(h)が通常より大きくなる
 * - 同じ行（Y座標が近い）の選択肢の中で最もhが大きいものを選択と判定
 */

const CoordMapper = {
  /**
   * OCR結果から全設問の回答をマッピングする
   * @param {Object} ocrResult - OcrUtils.runOcr() の戻り値
   * @returns {Object} 設問ごとの回答マップ
   */
  mapAnswers(ocrResult) {
    const words = this._extractWords(ocrResult);

    return {
      age:         this._detectByHeight(words, AGE_MAP),
      topSize:     this._detectSizeTop(words),
      bottomSize:  this._detectSizeBottom(words),
      height:      this._extractHeight(words),
      bodyType:    this._detectByHeight(words, BODY_TYPE_MAP),
      area:        this._extractArea(words),
      maternity:   this._extractMaternity(words),
      q1:          this._detectByHeight(words, Q1_MAP),
      repeatCount: this._extractRepeatCount(words),
      q2:          this._detectByHeight(words, Q2_MAP),
      q2Other:     this._extractQ2Other(words),
      q3Quality:   this._detectQ3Quality(words),
      q3Comfort:   this._detectQ3Comfort(words),
      q3Size:      this._detectByHeight(words, Q3_SIZE_MAP),
      q3Length:    this._detectByHeight(words, Q3_LENGTH_MAP),
      q3NgReason:  '',
      q4Flow:      this._detectByHeight(words, Q4_MAP),
      q4NgReason:  '',
      freeComment: '',
    };
  },

  /**
   * OCR結果から全ワードを {text, x, y, w, h} の配列として抽出する
   */
  _extractWords(ocrResult) {
    const words = [];
    for (const block of ocrResult.pages[0].blocks) {
      for (const para of block.paragraphs) {
        for (const word of para.words) {
          const text = word.symbols.map(s => s.text).join('');
          const v = word.boundingBox.vertices;
          const x = v[0].x, y = v[0].y;
          const w = (v[2].x || v[1].x) - x;
          const h = (v[2].y || v[3].y) - y;
          words.push({ text, x, y, w, h });
        }
      }
    }
    return words;
  },

  /**
   * 座標マップから最もhが大きいワードを選択と判定する
   * @param {Array} words - 全ワード配列
   * @param {Array} coordMap - [{label, x1, x2, y1, y2}] の座標マップ
   * @returns {string} 選択された選択肢のlabel
   */
  _detectByHeight(words, coordMap) {
    // 各選択肢エリアの最大hを収集
    const results = coordMap.map(def => {
      const matched = words.filter(w =>
        w.x >= def.x1 && w.x <= def.x2 &&
        w.y >= def.y1 && w.y <= def.y2
      );
      const maxH = matched.length ? Math.max(...matched.map(w => w.h)) : 0;
      return { label: def.label, maxH };
    });

    Logger.log('高さ判定: ' + JSON.stringify(results));

    // hが0の選択肢を除外
    const valid = results.filter(r => r.maxH > 0);
    if (!valid.length) return '';

    // 最大hの選択肢を返す（全て同じhなら未選択と判定）
    const maxH = Math.max(...valid.map(r => r.maxH));
    const minH = Math.min(...valid.map(r => r.maxH));

    // 最大hが最小hより2px以上大きければ選択あり
    if (maxH - minH < 2) return '';

    return valid.find(r => r.maxH === maxH).label;
  },

  /**
   * トップスサイズを判定する
   * トップスとボトムスで同じS/M/L/XLがあるためY座標で区別
   */
  _detectSizeTop(words) {
    return this._detectByHeight(words, TOP_SIZE_MAP);
  },

  /**
   * ボトムスサイズを判定する
   */
  _detectSizeBottom(words) {
    return this._detectByHeight(words, BOTTOM_SIZE_MAP);
  },

  /**
   * 身長を抽出する（数値のboundingBoxが大きい）
   */
  _extractHeight(words) {
    // 「身長」ラベルの近く(y=600-650, x=500-700)にある数値を取得
    const numWord = words.find(w =>
      w.y >= 590 && w.y <= 660 && w.x >= 500 && w.x <= 720 &&
      /^\d+$/.test(w.text)
    );
    return numWord ? numWord.text : '';
  },

  /**
   * マタニティ月数を抽出する
   */
  _extractMaternity(words) {
    const numWord = words.find(w =>
      w.y >= 600 && w.y <= 660 && w.x >= 1150 && w.x <= 1380 &&
      /^\d+$/.test(w.text)
    );
    return numWord ? numWord.text : '';
  },

  /**
   * 利用エリアを抽出する（郡馬・高崎等）
   */
  _extractArea(words) {
    // Y座標780-830の範囲でx=400-700のテキストを結合
    const areaWords = words
      .filter(w => w.y >= 780 && w.y <= 840 && w.x >= 400 && w.x <= 720)
      .sort((a, b) => a.x - b.x)
      .map(w => w.text);
    return areaWords.join('');
  },

  /**
   * リピート回数を抽出する
   */
  _extractRepeatCount(words) {
    // リピートの「回目」の左側(x=1200-1410, y=955-995)にある数値
    const numWord = words.find(w =>
      w.y >= 955 && w.y <= 995 && w.x >= 1200 && w.x <= 1410 &&
      /^\d+$/.test(w.text)
    );
    return numWord ? numWord.text : '';
  },

  /**
   * Q2その他の記述を抽出する
   */
  _extractQ2Other(words) {
    // その他の括弧内(x=1238-1500, y=1075-1110)のテキスト
    const otherWords = words
      .filter(w => w.y >= 1075 && w.y <= 1110 && w.x >= 1238 && w.x <= 1500)
      .sort((a, b) => a.x - b.x)
      .map(w => w.text);
    return otherWords.join('');
  },

  /**
   * Q3品質を判定する（品質行のY座標で絞り込み）
   */
  _detectQ3Quality(words) {
    return this._detectByHeight(words, Q3_QUALITY_MAP);
  },

  /**
   * Q3着心地を判定する（着心地行のY座標で絞り込み）
   */
  _detectQ3Comfort(words) {
    return this._detectByHeight(words, Q3_COMFORT_MAP);
  }
};

// ===== 座標マップ定義 =====
// 各選択肢のboundingBox範囲 {label, x1, x2, y1, y2}
// X座標: アンケート左端=90, 右端=1530
// Y座標はスキャン画像の実測値

const AGE_MAP = [
  { label: '20歳以下',   x1: 380, x2: 560,  y1: 440, y2: 485 },
  { label: '21歳~25歳', x1: 580, x2: 760,  y1: 440, y2: 485 },
  { label: '26歳~30歳', x1: 760, x2: 960,  y1: 440, y2: 485 },
  { label: '31歳~35歳', x1: 960, x2: 1160, y1: 440, y2: 485 },
  { label: '36歳~40歳', x1: 1160, x2: 1360, y1: 440, y2: 485 },
  { label: '41歳以上',   x1: 1360, x2: 1530, y1: 440, y2: 485 },
];

const TOP_SIZE_MAP = [
  { label: 'S',   x1: 548, x2: 600,  y1: 535, y2: 585 },
  { label: 'M',   x1: 618, x2: 665,  y1: 535, y2: 585 },
  { label: 'L',   x1: 700, x2: 740,  y1: 535, y2: 585 },
  { label: 'XL~', x1: 748, x2: 830,  y1: 535, y2: 585 },
];

const BOTTOM_SIZE_MAP = [
  { label: 'S',   x1: 1170, x2: 1215, y1: 535, y2: 585 },
  { label: 'M',   x1: 1242, x2: 1290, y1: 535, y2: 585 },
  { label: 'L',   x1: 1322, x2: 1365, y1: 535, y2: 585 },
  { label: 'XL~', x1: 1377, x2: 1455, y1: 535, y2: 585 },
];

const BODY_TYPE_MAP = [
  { label: 'ストレート', x1: 390,  x2: 680,  y1: 685, y2: 735 },
  { label: 'ナチュラル', x1: 820,  x2: 1060, y1: 685, y2: 735 },
  { label: 'ウェーブ',   x1: 1260, x2: 1450, y1: 685, y2: 735 },
];

const Q1_MAP = [
  { label: 'web検索',        x1: 90,   x2: 310,  y1: 955, y2: 1000 },
  { label: '友人紹介',        x1: 335,  x2: 540,  y1: 955, y2: 1000 },
  { label: 'WeddingWEB招待状', x1: 560,  x2: 820,  y1: 955, y2: 1000 },
  { label: 'Instagram',      x1: 840,  x2: 1080, y1: 955, y2: 1000 },
  { label: 'リピート',        x1: 1085, x2: 1530, y1: 955, y2: 1000 },
];

const Q2_MAP = [
  { label: '結婚式',      x1: 90,   x2: 360,  y1: 1065, y2: 1115 },
  { label: '結婚式二次会', x1: 360,  x2: 660,  y1: 1065, y2: 1115 },
  { label: '謝恩会',      x1: 660,  x2: 920,  y1: 1065, y2: 1115 },
  { label: '食事会',      x1: 920,  x2: 1140, y1: 1065, y2: 1115 },
  { label: 'その他',      x1: 1140, x2: 1530, y1: 1065, y2: 1115 },
];

const Q3_QUALITY_MAP = [
  { label: 'とても良い',     x1: 90,   x2: 360,  y1: 1260, y2: 1310 },
  { label: 'まあまあ良い',   x1: 360,  x2: 660,  y1: 1260, y2: 1310 },
  { label: '普通',          x1: 660,  x2: 920,  y1: 1260, y2: 1310 },
  { label: 'あまり良くない', x1: 920,  x2: 1220, y1: 1260, y2: 1310 },
  { label: 'とても良くない', x1: 1220, x2: 1530, y1: 1260, y2: 1310 },
];

const Q3_COMFORT_MAP = [
  { label: 'とても良い',     x1: 90,   x2: 360,  y1: 1360, y2: 1410 },
  { label: 'まあまあ良い',   x1: 360,  x2: 660,  y1: 1360, y2: 1410 },
  { label: '普通',          x1: 660,  x2: 920,  y1: 1360, y2: 1410 },
  { label: 'あまり良くない', x1: 920,  x2: 1220, y1: 1360, y2: 1410 },
  { label: 'とても良くない', x1: 1220, x2: 1530, y1: 1360, y2: 1410 },
];

const Q3_SIZE_MAP = [
  { label: '小さい',    x1: 90,   x2: 330,  y1: 1560, y2: 1615 },
  { label: 'やや小さい', x1: 330,  x2: 630,  y1: 1560, y2: 1615 },
  { label: '丁度いい',  x1: 630,  x2: 930,  y1: 1560, y2: 1615 },
  { label: 'やや大きい', x1: 930,  x2: 1230, y1: 1560, y2: 1615 },
  { label: '大きい',    x1: 1230, x2: 1530, y1: 1560, y2: 1615 },
];

const Q3_LENGTH_MAP = [
  { label: 'ひざ上',    x1: 90,   x2: 330,  y1: 1658, y2: 1710 },
  { label: 'ひざ下',    x1: 330,  x2: 630,  y1: 1658, y2: 1710 },
  { label: 'ふくらはぎ', x1: 630,  x2: 930,  y1: 1658, y2: 1710 },
  { label: 'くるぶし上', x1: 930,  x2: 1230, y1: 1658, y2: 1710 },
  { label: 'くるぶし下', x1: 1230, x2: 1530, y1: 1658, y2: 1710 },
];

const Q4_MAP = [
  { label: 'わかりやすい', x1: 90,   x2: 330,  y1: 1835, y2: 1885 },
  { label: '便利',        x1: 330,  x2: 630,  y1: 1835, y2: 1885 },
  { label: '普通',        x1: 630,  x2: 930,  y1: 1835, y2: 1885 },
  { label: 'わかりづらい', x1: 930,  x2: 1230, y1: 1835, y2: 1885 },
  { label: '不便',        x1: 1230, x2: 1530, y1: 1835, y2: 1885 },
];
