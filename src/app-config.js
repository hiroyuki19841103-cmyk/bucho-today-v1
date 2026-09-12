/* Microsoft Entraの既存TASK CONTROLアプリと同じCLIENT_IDを使用します。 */
window.BUCHO_TODAY_CONFIG = Object.freeze({
  APP_VERSION: "2026.09.12-v1.0",
  CLIENT_ID: "421c060c-f769-4d5b-9bab-8c3b9a4cf3ec",
  AUTHORITY: "https://login.microsoftonline.com/common/",
  REDIRECT_URI: "https://witty-mushroom-0fcb9cb10.3.azurestaticapps.net/",
  GRAPH_SCOPES: ["Files.ReadWrite.AppFolder", "Calendars.ReadWrite"],
  TASK_FILE_NAME: "TASK_CONTROL_Master.xlsx",
  TODAY_FILE_NAME: "BUCHO_TODAY_Master.xlsx",
  AUTO_SYNC_DELAY_MS: 6000,
  AUTO_PULL_MINUTES: 3
});
