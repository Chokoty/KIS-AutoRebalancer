// 메뉴 추가
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 KIS AutoTrader')
    .addItem('💰 계좌 현황 새로고침', 'updateAccountSheet')
    .addItem('🔄 대시보드 새로고침', 'updateDashboard')
    .addItem('⚡ 리밸런싱 실행', 'executeRebalanceFromDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 수익 실현')
      .addItem('📋 수익 실현 창 열기', 'openWithdrawDialog')
      .addSeparator()
      .addItem('🔓 보호 예수금 해제', 'releaseProtectedCash'))
    .addSeparator()
    .addItem('🛣️ 주간 자동 리밸런싱 (차선 유지) 활성/비활성', 'toggleHighwayLaneKeeping')
    .addSeparator()
    .addItem('🤖 AI 시장 분석 및 비중 제안', 'runAIBriefing')
    .addSeparator()
    .addItem('📖 기본 사용법 안내', 'showUsageGuide')
    .addSubMenu(ui.createMenu('⚙️ 설정 및 관리')
      .addItem('⚙️ 초기 설정', 'setupSheets')
      .addItem('🛡️ API 키 보안 설정', 'openSecureConfigDialog')
      .addItem('🔑 토큰 초기화 (오류 발생 시)', 'forceRefreshToken'))
    .addSeparator()
    .addToUi();
}

/**
 * 고속도로 차선 유지 (주간 자동 리밸런싱) 토글
 */
function toggleHighwayLaneKeeping() {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('HIGHWAY_LANE_KEEPING') === 'TRUE';
  const next = !current;

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'scheduledBiWeeklyRebalance') ScriptApp.deleteTrigger(t);
  });

  if (next) {
    ScriptApp.newTrigger('scheduledBiWeeklyRebalance')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(10)
      .create();
    props.setProperty('HIGHWAY_LANE_KEEPING', 'TRUE');
    SpreadsheetApp.getUi().alert('🛣️ 차선 유지(정기 리밸런싱)가 [ON] 되었습니다.\n(매주 월요일 오전 10시 실행)');
  } else {
    props.setProperty('HIGHWAY_LANE_KEEPING', 'FALSE');
    SpreadsheetApp.getUi().alert('🛣️ 차선 유지(정기 리밸런싱)가 [OFF] 되었습니다.');
  }
  updateDashboard();
}

// 초기 시트 설정
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 설정 시트
  let sheet = ss.getSheetByName('⚙️ 설정');
  let lastAccountType = '일반';

  if (sheet) {
    const existingValues = sheet.getRange('B2:B7').getValues();
    // B7(index 5): 계좌 종류, 기존 TRUE/FALSE 마이그레이션
    const raw = String(existingValues[5][0] || '');
    if (raw === 'TRUE') lastAccountType = '모의';
    else if (['일반', 'ISA', '모의'].includes(raw)) lastAccountType = raw;
    else lastAccountType = '일반';
  } else {
    sheet = ss.insertSheet('⚙️ 설정');
  }

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.getRange('A1').setValue('KIS API 설정').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange('A2:B10').setValues([
    ['APP KEY',                      '🛡️ 보안 저장됨'],
    ['APP SECRET',                   '🛡️ 보안 저장됨'],
    ['계좌번호',                     '🛡️ 보안 저장됨'],
    ['Gemini API Key',               '🛡️ 보안 저장됨'],
    ['API 키 발급처',                'https://aistudio.google.com/app/apikey'],
    ['계좌 종류 (일반/ISA/모의)',    lastAccountType],
    ['리밸런싱 임계치 (%)',           2.0],
    ['수익실현 임계치 (%)',           40.0],
    ['연 목표 수익률 (%)',            10.0]
  ]);
  sheet.getRange('A2:A10').setHorizontalAlignment('center');
  sheet.getRange('B2:B10').setHorizontalAlignment('right');
  sheet.getRange('B6').setFontColor('#1a73e8').setFontLine('underline');

  // 계좌 종류 드롭다운
  const accountRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['일반', 'ISA', '모의'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B7').setDataValidation(accountRule);

  trimExtraColumns(sheet);

  // 2. 대시보드 시트
  sheet = ss.getSheetByName('📊 대시보드');
  if (!sheet) sheet = ss.insertSheet('📊 대시보드', 0);
  setupDashboardSheet(sheet);

  // 3. 계좌현황 시트
  sheet = ss.getSheetByName('🏦 계좌현황');
  if (!sheet) sheet = ss.insertSheet('🏦 계좌현황');
  sheet.clear();
  sheet.getRange('A1').setValue('💰 계좌 현황').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange('A2:B5').setValues([
    ['예수금', 0],
    ['주문가능금액', 0],
    ['업데이트 시간', ''],
    ['총 평가액', 0]
  ]);
  sheet.getRange('A2:A5').setHorizontalAlignment('center');
  sheet.getRange('B2:B5').setHorizontalAlignment('right');
  trimExtraColumns(sheet);
  sheet.getRange('A6').setValue('💰 전체 수익률').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A7:I7').setValues([[
    '종목코드', '종목명', '보유수량', '평균단가', '현재가', '평가금액', '손익', '수익률(%)', '목표여부'
  ]]).setFontWeight('bold').setBackground('#4285f4').setFontColor('white').setHorizontalAlignment('center');

  // 4. 포트폴리오설정 시트
  sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) sheet = ss.insertSheet('📋 포트폴리오설정');
  sheet.clear();
  sheet.getRange('A1').setValue('🎯 목표 포트폴리오').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange('A2:D2').setValues([['종목코드', '종목명', '목표비율(%)', '유형']])
    .setFontWeight('bold').setBackground('#34a853').setFontColor('white').setHorizontalAlignment('center');
  sheet.getRange('A3:D9').setValues([
    ['440650', 'ACE 미국달러단기채권액티브', 30, '채권'],
    ['319640', 'TIGER 골드선물(H)', 20, '금'],
    ['261240', 'KODEX 미국달러선물', 10, '달러'],
    ['161510', 'PLUS 고배당주', 15, '국내주식'],
    ['315960', 'RISE 대형고배당10TR', 10, '국내주식'],
    ['379800', 'KODEX 미국S&P500TR', 10, '해외주식'],
    ['', '현금', 5, '현금']
  ]);
  sheet.getRange('A3:D').setHorizontalAlignment('right');
  trimExtraColumns(sheet);

  // 5. 수익실현 시트
  sheet = ss.getSheetByName('💰 수익실현');
  if (!sheet) sheet = ss.insertSheet('💰 수익실현');
  setupWithdrawSheet(sheet);

  // 6. 거래내역 시트
  setupTradeHistorySheet();

  // 7. 수익실현기록 시트
  setupProfitHistorySheet();

  // 8. 자동 새로고침 트리거 (onOpen)
  const triggers = ScriptApp.getProjectTriggers();
  const hasRefreshTrigger = triggers.some(t => t.getHandlerFunction() === 'automatedRefreshRoutine');
  if (!hasRefreshTrigger) {
    ScriptApp.newTrigger('automatedRefreshRoutine')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
  }

  SpreadsheetApp.getUi().alert('✅ 초기 설정이 완료되었습니다!\n\n곧바로 API 키 보안 설정 창이 열립니다.\n금융 정보를 안전하게 입력해 주세요.');
  openSecureConfigDialog();
}
