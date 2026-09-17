/**
 * 업데이트 내역 보기 — GitHub 릴리스 페이지 링크를 보여주고, 알림 배너를 지운다.
 * (예전엔 window.open()으로 자동으로 새 탭을 열려고 했는데, 브라우저 팝업 차단에
 *  걸려 실제로는 아무 탭도 안 열렸다 — 사용자가 직접 누르는 링크로 바꿈)
 */
function showVersionHistory() {
  // 링크를 열든 안 열든, "확인했다"는 의도이므로 알림 배너는 바로 지운다.
  syncTemplateVersion(true);

  const url = 'https://github.com/Chokoty/KIS-AutoRebalancer/releases';
  const html = `<!DOCTYPE html><html><body style="font-family:'Malgun Gothic',sans-serif;padding:20px;text-align:center;">
    <p style="font-size:14px;color:#3c4043;margin-bottom:16px;">업데이트 내역은 GitHub 릴리스 페이지에서 확인할 수 있습니다.</p>
    <a href="${url}" target="_blank" style="display:inline-block;background:#1a73e8;color:white;text-decoration:none;padding:10px 20px;border-radius:6px;font-weight:bold;">GitHub에서 업데이트 내역 보기 ↗</a>
    <div style="margin-top:16px;"><button onclick="google.script.host.close()" style="background:none;border:1px solid #dadce0;border-radius:4px;padding:6px 14px;cursor:pointer;">닫기</button></div>
  </body></html>`;

  const output = HtmlService.createHtmlOutput(html).setWidth(360).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(output, '📜 업데이트 내역');
}
