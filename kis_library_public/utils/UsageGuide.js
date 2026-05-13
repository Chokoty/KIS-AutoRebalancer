/**
 * GitHub README 기본 사용법 섹션으로 이동
 */
function showUsageGuide() {
  const url = 'https://github.com/Chokoty/kis-auto-rebalance#-기본-사용법';
  const html = `<html><body style="font-family:sans-serif;text-align:center;padding:20px;">
    <p>GitHub 사용법 페이지로 이동합니다...</p>
    <a href="${url}" target="_blank" style="color:#1a73e8;">바로 가기 →</a>
    <script>window.open('${url}','_blank');google.script.host.close();</script>
  </body></html>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(300).setHeight(100),
    '기본 사용법 안내'
  );
}
