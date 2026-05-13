// 메뉴 추가
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 KIS AutoTrader')
    .addItem('💰 계좌 현황 새로고침', 'updateAccountSheet')
    .addItem('🔄 대시보드 새로고침', 'updateDashboard')
    .addItem('⚡ 리밸런싱 실행', 'executeRebalanceFromDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('🤖 AI 분석')
      .addItem('📊 AI 비중 제안', 'runAIBriefing')
      .addItem('💬 AI 빠른 질문', 'openAIQuickQuestion'))
    .addItem('💡 추천 비중 반영', 'applyLatestRecommendation')
    .addItem('📋 포트폴리오 종목 추가/관리', 'openPortfolioManagerDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 수익 실현')
      .addItem('📋 수익 실현 창 열기', 'openWithdrawDialog')
      .addSeparator()
      .addItem('🔓 보호 예수금 해제', 'releaseProtectedCash'))
    .addSeparator()
    .addItem('🛣️ 차선유지 (격주 정기 리밸런싱) 켜기/끄기', 'toggleHighwayLaneKeeping')
    .addSeparator()
    .addItem('📜 업데이트 내역 보기', 'showVersionHistory')
    .addItem('📖 기본 사용법 안내', 'showUsageGuide')
    .addSubMenu(ui.createMenu('⚙️ 설정 및 관리')
      .addItem('⚙️ 초기 설정',                  'setupSheets')
      .addItem('🛡️ API 키 보안 설정',           'openSecureConfigDialog')
      .addItem('⚙️ AI 프롬프트 상세 설정',      'openAIPromptSettings')
      .addItem('🔑 토큰 초기화 (오류 발생 시)', 'forceRefreshToken'))
    .addToUi();
}

/**
 * 고속도로 차선 유지 (격주 자동 리밸런싱) 토글
 */
function toggleHighwayLaneKeeping() {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('HIGHWAY_LANE_KEEPING') === 'TRUE';
  const next = !current;

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledBiWeeklyRebalance') ScriptApp.deleteTrigger(t);
  });

  if (next) {
    ScriptApp.newTrigger('scheduledBiWeeklyRebalance')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(10)
      .create();
    props.setProperty('HIGHWAY_LANE_KEEPING', 'TRUE');
    SpreadsheetApp.getUi().alert('🛣️ 차선유지 [ON]\n매주 월요일 10시 발동, 단 마지막 실행 후 13일 미만이면 자동 스킵 (격주 효과).');
  } else {
    props.setProperty('HIGHWAY_LANE_KEEPING', 'FALSE');
    SpreadsheetApp.getUi().alert('🛣️ 차선 유지(정기 리밸런싱)가 [OFF] 되었습니다.');
  }
  updateDashboard();
}

/**
 * 포트폴리오설정 시트 컬럼 순서 마이그레이션.
 * 구 레이아웃(C=목표비율, D=유형, E=초기비율) → 신 레이아웃(C=기준비율, D=운용비율, E=유형)
 */
function addInitialRatiosColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('포트폴리오설정 시트를 찾을 수 없습니다.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const c2Header = String(sheet.getRange('C2').getValue()).trim();
  const d2Header = String(sheet.getRange('D2').getValue()).trim();

  // 이미 새 레이아웃이면 아무 것도 안 함
  if (c2Header === '기준비율(%)' && d2Header === '운용비율(%)') {
    SpreadsheetApp.getUi().alert('✅ 이미 새 컬럼 순서입니다.\n(C=기준비율, D=운용비율, E=유형)');
    return;
  }

  // 구 레이아웃 감지: C=목표비율, D=유형, E=초기비율
  const dataRows = lastRow - 2;
  if (dataRows <= 0) return;

  const oldData = sheet.getRange(3, 3, dataRows, 3).getValues(); // C, D, E 읽기
  // oldData[i] = [oldC(목표비율), oldD(유형), oldE(초기비율)]
  const newData = oldData.map(row => [
    (typeof row[2] === 'number' && row[2] > 0) ? row[2] : row[0], // 새C=기준비율
    row[0], // 새D=운용비율(구C)
    row[1]  // 새E=유형(구D)
  ]);

  // 헤더 업데이트
  sheet.getRange('C2').setValue('기준비율(%)');
  sheet.getRange('D2').setValue('운용비율(%)');
  sheet.getRange('E2').setValue('유형');

  // 데이터 업데이트
  sheet.getRange(3, 3, dataRows, 3).setValues(newData);
  sheet.getRange(3, 3, dataRows, 1).setHorizontalAlignment('right');
  sheet.getRange(3, 4, dataRows, 1).setHorizontalAlignment('right');
  sheet.getRange(3, 5, dataRows, 1).setHorizontalAlignment('left');

  SpreadsheetApp.getUi().alert(
    '✅ 컬럼 순서가 변경되었습니다.\n\n' +
    'C열: 기준비율(%) — 고정값 (처음 설정한 목표)\n' +
    'D열: 운용비율(%) — 수정 가능 (AI/사람이 조정)\n' +
    'E열: 유형\n\n' +
    '기준비율(C열)이 의도한 값과 다르면 직접 수정해 주세요.'
  );
}

// 초기 시트 설정
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 설정 시트
  let sheet = ss.getSheetByName('⚙️ 설정');
  let lastAccountType = '일반';

  if (sheet) {
    const existingValues = sheet.getRange('B2:B7').getValues();
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
  sheet.getRange('A2:B11').setValues([
    ['APP KEY',                      '🛡️ 보안 저장됨'],
    ['APP SECRET',                   '🛡️ 보안 저장됨'],
    ['계좌번호',                     '🛡️ 보안 저장됨'],
    ['Gemini API Key',               '🛡️ 보안 저장됨'],
    ['API 키 발급처',                'https://aistudio.google.com/app/apikey'],
    ['계좌 종류 (일반/ISA/모의)',    lastAccountType],
    ['리밸런싱 임계치 (%)',           2.0],
    ['수익실현 임계치 (%)',           40.0],
    ['연 목표 수익률 (%)',            10.0],
    ['사용법 안내',                  'https://github.com/Chokoty/kis-auto-rebalance#-기본-사용법']
  ]);
  sheet.getRange('A2:A11').setHorizontalAlignment('center');
  sheet.getRange('B2:B11').setHorizontalAlignment('right');
  sheet.getRange('B6').setFontColor('#1a73e8').setFontLine('underline');
  sheet.getRange('B11').setFontColor('#1a73e8').setFontLine('underline');

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
  // 컬럼 순서: 종목코드 | 종목명 | 기준비율(%) [고정] | 운용비율(%) [수정가능] | 유형
  sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) sheet = ss.insertSheet('📋 포트폴리오설정');

  // 기존 기준비율(C열) 보존
  const prevPortLastRow = sheet.getLastRow();
  const savedInitialRatios = {};
  if (prevPortLastRow >= 3) {
    const existingData = sheet.getRange(3, 1, prevPortLastRow - 2, 3).getValues();
    existingData.forEach(row => {
      const code = String(row[0]).trim();
      const initRatio = row[2]; // C열: 기준비율
      if (code && typeof initRatio === 'number' && initRatio > 0) {
        savedInitialRatios[code] = initRatio;
      }
    });
  }

  sheet.clear();
  sheet.getRange('A1').setValue('🎯 목표 포트폴리오').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange('A2:E2').setValues([['종목코드', '종목명', '기준비율(%)', '운용비율(%)', '유형']])
    .setFontWeight('bold').setBackground('#34a853').setFontColor('white').setHorizontalAlignment('center');

  const defaultPortfolio = [
    ['440650', 'ACE 미국달러단기채권액티브', 30, 30, '채권'],
    ['319640', 'TIGER 골드선물(H)', 20, 20, '금'],
    ['261240', 'KODEX 미국달러선물', 10, 10, '달러'],
    ['161510', 'PLUS 고배당주', 15, 15, '국내주식'],
    ['315960', 'RISE 대형고배당10TR', 10, 10, '국내주식'],
    ['379800', 'KODEX 미국S&P500TR', 10, 10, '해외주식'],
    ['', '현금', 5, 5, '현금']
  ];

  // 기존 기준비율이 있으면 C열에 복원
  const dataWithRestoredInitial = defaultPortfolio.map(row => {
    const code = row[0];
    const saved = code ? savedInitialRatios[code] : null;
    return [row[0], row[1], saved || row[2], row[3], row[4]];
  });

  sheet.getRange(3, 1, dataWithRestoredInitial.length, 5).setValues(dataWithRestoredInitial);
  sheet.getRange('A3:E').setHorizontalAlignment('right');
  sheet.getRange('B3:B9').setHorizontalAlignment('left');
  trimExtraColumns(sheet);

  // 5. 거래내역 시트
  setupTradeHistorySheet();

  // 6. 수익실현기록 시트
  setupProfitHistorySheet();

  // 7. 비중변경이력 시트
  setupAIHistorySheet();

  // 8. 자동 새로고침 트리거
  const triggers = ScriptApp.getProjectTriggers();
  const hasRefreshTrigger = triggers.some(t => t.getHandlerFunction() === 'automatedRefreshRoutine');
  if (!hasRefreshTrigger) {
    ScriptApp.newTrigger('automatedRefreshRoutine')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
  }

  syncTemplateVersion();

  SpreadsheetApp.getUi().alert('✅ 초기 설정이 완료되었습니다!\n\n곧바로 API 키 보안 설정 창이 열립니다.\n금융 정보를 안전하게 입력해 주세요.');
  openSecureConfigDialog();
}
