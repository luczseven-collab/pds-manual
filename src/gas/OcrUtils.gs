/**
 * OcrUtils.gs — 画像取得ユーティリティ
 * スキャン画像をBase64に変換してClaudeに渡す準備をする
 * PDFの場合はDrive OCRでJPEGに変換する
 */

const OcrUtils = {
  /**
   * ファイルをBase64画像データとして返す
   * @param {File} file - Google DriveのFileオブジェクト（PDF/JPEG/PNG）
   * @returns {Object} { base64: string, mimeType: string }
   */
  getImageData(file) {
    const mime = file.getMimeType();

    if (mime === 'application/pdf') {
      return this._convertPdfToImageData(file);
    }

    return {
      base64: Utilities.base64Encode(file.getBlob().getBytes()),
      mimeType: 'image/jpeg'
    };
  },

  /**
   * PDFをDrive OCRでJPEGに変換してBase64で返す
   * @param {File} pdfFile
   * @returns {Object} { base64: string, mimeType: string }
   */
  _convertPdfToImageData(pdfFile) {
    const converted = Drive.Files.insert(
      {
        title: `tmp_ocr_${pdfFile.getName()}`,
        mimeType: 'application/vnd.google-apps.document'
      },
      pdfFile.getBlob(),
      { convert: true, ocr: true }
    );

    const exportUrl = `https://docs.google.com/feeds/download/documents/export/Export?id=${converted.id}&exportFormat=jpeg`;
    const jpeg = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
    });

    Drive.Files.remove(converted.id);

    return {
      base64: Utilities.base64Encode(jpeg.getBlob().getBytes()),
      mimeType: 'image/jpeg'
    };
  }
};
