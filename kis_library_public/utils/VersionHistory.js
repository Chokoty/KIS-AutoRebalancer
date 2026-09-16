function showVersionHistory() {
  const url = 'https://github.com/Chokoty/KIS-AutoTrader/releases';
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}'); google.script.host.close();</script>`
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, '업데이트 내역 열기...');
}