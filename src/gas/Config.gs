/**
 * Config.gs — スクリプトプロパティ管理
 * APIキーや設定値はすべてここから取得する（コードへのハードコード禁止）
 */

const Config = {
  get DRIVE_FOLDER_OK_ID()      { return PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_OK_ID'); },
  get DRIVE_FOLDER_NG_ID()      { return PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_NG_ID'); },
  get DRIVE_FOLDER_MONTHLY_ID() { return PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_MONTHLY_ID'); },
  get SPREADSHEET_ID()     { return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); },
  get SCRIPT_ID()          { return PropertiesService.getScriptProperties().getProperty('SCRIPT_ID'); },
  get CLAUDE_API_KEY()     { return PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY'); },
  get VISION_API_KEY()     { return PropertiesService.getScriptProperties().getProperty('VISION_API_KEY'); },
  get CHATWORK_API_TOKEN() { return PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN'); },
  get CHATWORK_ROOM_ID()   { return PropertiesService.getScriptProperties().getProperty('CHATWORK_ROOM_ID'); },
  get NOTIFY_EMAIL()       { return PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL'); },
  get DRIVE_FOLDER_IMAGE_DB_ID()       { return PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_IMAGE_DB_ID'); },
  get DRIVE_FOLDER_CUSTOMER_IMAGE_ID() { return PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_CUSTOMER_IMAGE_ID'); },
};
