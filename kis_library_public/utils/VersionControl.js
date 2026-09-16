/**
 * KIS AutoTrader 버전 확인 및 관리 (배포 자동화용)
 */

const KIS_SYSTEM_VERSION = "v1.1.0"; // 개발자가 새 업데이트를 할 때마다 이 버전을 올려서 배포합니다.

/**
 * 대시보드에 뿌려줄 업데이트 알림 메시지를 확인합니다.
 */
function checkVersionUpdate() {
  const props = PropertiesService.getDocumentProperties();
  let userVersion = props.getProperty('TEMPLATE_VERSION');
  
  // 사용자의 DocumentProperties에 버전이 없으면(과거 유저 등) v1.0.0으로 간주
  if (!userVersion) {
    userVersion = "v1.0.0";
  }
  
  if (userVersion !== KIS_SYSTEM_VERSION) {
    return `🔔 업데이트 알림: 새 버전(${KIS_SYSTEM_VERSION}) 코드가 적용되었습니다! 메뉴에서 '업데이트 내역'을 확인하시면 알림이 사라집니다.`;
  }
  return null;
}

/**
 * 현재 사용자의 시트를 라이브러리 최신 버전에 맞춰 동기화합니다 (알림 제거용)
 * @param {boolean} silent - true인 경우 완료 토스트를 띄우지 않습니다.
 */
function syncTemplateVersion(silent = false) {
  PropertiesService.getDocumentProperties().setProperty('TEMPLATE_VERSION', KIS_SYSTEM_VERSION);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    if (!silent) {
      ss.toast(`시트 버전이 ${KIS_SYSTEM_VERSION} (으)로 동기화되었습니다.`, '✅ 버전 업데이트 완료');
    }
    
    // 대시보드 시트가 열려있다면 즉시 알림 배너 지우기
    const dashboardSheet = ss.getSheetByName('📊 대시보드');
    if (dashboardSheet) {
      dashboardSheet.getRange('A2:N2').breakApart().clearContent().setBackground('white');
    }
  }
}
