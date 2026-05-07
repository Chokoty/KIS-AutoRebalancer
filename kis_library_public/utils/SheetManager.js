// 계좌 현황 시트 업데이트 (목표외 종목 수익률 제외)
function updateAccountSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🏦 계좌현황');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    ss.toast('계좌 정보를 조회 중입니다...', '⏳ 처리 중', -1);
    
    // 1. 계좌 정보 조회
    const balance = getBalance();
    const holdings = getHoldings();
    const targetPortfolio = getTargetPortfolio();
    
    Logger.log('총평가액: ' + balance.totalEval);
    Logger.log('현금: ' + balance.cash);
    Logger.log('보유주식 개수: ' + holdings.length);
    
    const totalEval = balance.totalEval;
    
    // 2. 목표 종목만 필터링하여 전체 수익률 계산
    const targetHoldings = holdings.filter(h => targetPortfolio[h.code]);
    
    let totalInvested = 0;  // 총 투자금액 (목표 종목만)
    let totalProfit = 0;    // 총 손익 (목표 종목만)
    
    targetHoldings.forEach(h => {
      totalInvested += h.avgPrice * h.quantity;
      totalProfit += h.profit;
    });
    
    const totalReturn = totalInvested > 0 ? (totalProfit / totalInvested) * 100 : 0;
    
    Logger.log('목표 종목 수: ' + targetHoldings.length);
    Logger.log('목표 종목 투자금액: ' + totalInvested);
    Logger.log('목표 종목 손익: ' + totalProfit);
    Logger.log('목표 종목 수익률: ' + totalReturn.toFixed(2) + '%');
    
    // 3. 시트 상단 정보 업데이트
    sheet.getRange('B2').setValue(balance.cash).setNumberFormat('#,##0');
    sheet.getRange('B3').setValue(balance.buyPower).setNumberFormat('#,##0');
    sheet.getRange('B4').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('B5').setValue(totalEval).setNumberFormat('#,##0');
    
    // 전체 수익률 표시 (목표 종목만, 색상 적용)
    const returnCell = sheet.getRange('B6');
    returnCell.setValue(totalReturn.toFixed(2) + '%');
    if (totalReturn > 0) {
      returnCell.setFontColor('#cc0000').setFontWeight('bold'); // 빨간색 (수익)
    } else if (totalReturn < 0) {
      returnCell.setFontColor('#0000cc').setFontWeight('bold'); // 파란색 (손실)
    } else {
      returnCell.setFontColor('#000000').setFontWeight('normal'); // 검정색
    }
    
    // 4. 헤더 업데이트 (A7로 이동)
    sheet.getRange('A7:I7').setValues([[
      '종목코드', '종목명', '보유수량', '평균단가', '현재가', '평가금액', '손익', '수익률(%)', '목표여부'
        ]]).setFontWeight('bold').setBackground('#4285f4').setFontColor('white').setHorizontalAlignment('center');
    
    // 5. 데이터 입력 및 나머지 로직 (기존 view_file 기반으로 복구)
    const lastRow = sheet.getLastRow();
    if (lastRow >= 8) {
      sheet.getRange(8, 1, lastRow - 7, 9).clearContent();
    }
    
    if (holdings.length > 0) {
      const data = holdings.map(h => {
        const isTarget = targetPortfolio[h.code] ? '✅ 목표종목' : '⚠️ 목표외';
        return [h.code, h.name, h.quantity, h.avgPrice, h.currentPrice, h.evalAmount, h.profit, h.profitRate, isTarget];
      });
      sheet.getRange(8, 1, data.length, 9).setValues(data);
      sheet.getRange(8, 1, data.length, 9).setHorizontalAlignment('right');
      sheet.getRange(8, 3, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(8, 4, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(8, 5, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(8, 6, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(8, 7, data.length, 1).setNumberFormat('#,##0');
      sheet.getRange(8, 8, data.length, 1).setNumberFormat('0.00%');
      
      for (let i = 0; i < holdings.length; i++) {
        const row = 8 + i;
        const profit = holdings[i].profit;
        const profitColor = profit > 0 ? '#ff0000' : (profit < 0 ? '#0000ff' : '#000000');
        sheet.getRange(row, 7).setFontColor(profitColor);
        sheet.getRange(row, 8).setFontColor(profitColor);
        const targetCell = sheet.getRange(row, 9);
        if (data[i][8].includes('목표종목')) targetCell.setFontColor('#137333').setFontWeight('bold');
        else targetCell.setFontColor('#f9ab00').setFontWeight('bold');
      }
    }
    
    const nonTargetCount = holdings.length - targetHoldings.length;
    const formattedCash = Number(balance.cash).toLocaleString('ko-KR');
    const formattedTotalEval = Number(balance.totalEval).toLocaleString('ko-KR');
    const sign = totalReturn > 0 ? '+' : '';
    
    const toastMsg = `💰 현금: ${formattedCash}원 | 📈 수익률: ${sign}${totalReturn.toFixed(2)}% (목표 종목만)\n📊 ${holdings.length}개 종목 (목표외 ${nonTargetCount}개 제외) | 💼 총: ${formattedTotalEval}원`;
    
    ss.toast(toastMsg, '✅ 계좌 현황 업데이트 완료!', 8);

    
  } catch (e) {
    Logger.log('updateAccountSheet 오류: ' + e.toString());
    ss.toast(e.message, '❌ 오류 발생', 10);
  }
}

/**
 * 포트폴리오 비중 변경 이력 시트 초기화
 * 컬럼: 시간 | 종목명 | 유형 | 변경전(%) | 변경후(%) | 변경이유 | 활용모델 | 상태
 */
function setupAIHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet) {
    sheet = ss.insertSheet('📝 비중변경이력');
  }

  // 유형 컬럼이 포함된 신형 헤더인지 확인 (col 3 = 유형)
  const currentHeader = sheet.getRange('A1:H1').getValues()[0];
  if (currentHeader[2] === '🏷️ 유형' && currentHeader[7] === '📊 상태') {
    trimExtraColumns(sheet, 8);
    return;
  }

  sheet.getRange('A1:H1').setValues([[
    '🕒 시간', '📄 종목명', '🏷️ 유형', '📉 변경 전(%)', '📈 변경 후(%)', '💡 변경 이유', '🤖 활용 모델', '📊 상태'
  ]])
  .setFontWeight('bold')
  .setBackground('#673ab7')
  .setFontColor('white')
  .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(6, 350);
  sheet.setColumnWidth(8, 100);
  sheet.getRange('A:H').setVerticalAlignment('middle');
  trimExtraColumns(sheet, 8);
}

/**
 * 거래내역 시트 초기화 (NEW)
 */
function setupTradeHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('📝 거래내역');
  if (!sheet) {
    sheet = ss.insertSheet('📝 거래내역');
  }
  
  const headerRange = sheet.getRange('A1:I1');
  const currentHeader = headerRange.getValues()[0];
  
  // 이미 설정되어 있다면 배수 처리
  if (currentHeader[0].includes('시간')) {
    trimExtraColumns(sheet, 9);
    return;
  }
  
  sheet.clear(); // 기존 내용 삭제 (제목 행 제거 포함)
  sheet.getRange('A1:I1').setValues([[
    '🕒 거래시간', '📄 구분', '🔢 종목코드', '📄 종목명', '📦 수량', '💵 가격', '💰 금액', '✅ 상태', '💬 메시지'
  ]])
  .setFontWeight('bold')
  .setBackground('#ea4335')
  .setFontColor('white')
  .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(9, 300);
  sheet.getRange('A:I').setVerticalAlignment('middle');
  sheet.getRange('A2:I').setHorizontalAlignment('right');
  trimExtraColumns(sheet, 9);
}

/**
 * 수익실현기록 시트 초기화
 */
function setupProfitHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('📝 수익실현기록');
  if (!sheet) {
    sheet = ss.insertSheet('📝 수익실현기록');
  }
  
  const headerRange = sheet.getRange('A1:H1');
  const currentHeader = headerRange.getValues()[0];
  if (currentHeader[0].includes('날짜')) {
    trimExtraColumns(sheet, 8);
    return;
  }
  
  sheet.clear();
  sheet.getRange('A1:H1').setValues([[
    '🕒 날짜', '📄 종목명', '📄 구분', '📦 수량', '💵 가격', '💰 거래금액', '💵 확정수익', '💰 누적수익'
  ]])
  .setFontWeight('bold')
  .setBackground('#fbbc04')
  .setFontColor('black')
  .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 150);
  sheet.getRange('A:H').setVerticalAlignment('middle');
  sheet.getRange('A2:H').setHorizontalAlignment('right');
  trimExtraColumns(sheet, 8);
}

/**
 * 시트의 빈 열을 제거하여 데이터가 있는 열까지만 남깁니다.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 대상 시트
 * @param {number} minCols 최소 유지할 열의 수 (기본값: 데이터가 있는 마지막 열)
 */
function trimExtraColumns(sheet, minCols) {
  if (!sheet) return;
  
  const maxCols = sheet.getMaxColumns();
  const lastCol = sheet.getLastColumn();
  const targetCol = Math.max(lastCol, minCols || 0);
  
  if (maxCols > targetCol && targetCol > 0) {
    try {
      sheet.deleteColumns(targetCol + 1, maxCols - targetCol);
    } catch (e) {
      Logger.log(`trimExtraColumns 오류 (${sheet.getName()}): ` + e.toString());
    }
  }
}

/**
 * 템플릿 배포용 초기화
 * 사용자 개인정보(API 키, 계좌번호 등) 및 거래 기록을 모두 지우고
 * 시트를 타인에게 공유하기 좋은 빈 상태로 만듭니다.
 * (개발자 전용 도구 - 메뉴에 노출하지 않음)
 */
function prepareTemplateSheet() {
  Logger.log('⚠️ [템플릿 배포 준비 시작] 개인정보 및 데이터를 초기화합니다.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 설정 시트 초기화 (민감정보 삭제)
  const configSheet = ss.getSheetByName('⚙️ 설정');
  if (configSheet) {
    configSheet.getRange('B2').setValue('');
    configSheet.getRange('B3').setValue('');
    configSheet.getRange('B4').setValue('12345678-01');
    configSheet.getRange('B5').setValue('');   // Gemini API Key
    configSheet.getRange('B7').setValue('일반'); // 계좌 종류
  }

  // 2. 기록 시트 내용 삭제 (헤더는 남김)
  const historySheets = ['📝 거래내역', '📝 수익실현기록', '📝 비중변경이력'];
  historySheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        // 모든 히스토리 시트의 헤더가 1줄로 통일되었으므로 2행부터 데이터 삭제
        const lastCol = sheet.getLastColumn() || 10;
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
  });

  // 3. 대시보드 데이터 영역 클리어 (A9:O 범위)
  const dashboardSheet = ss.getSheetByName('📊 대시보드');
  if (dashboardSheet) {
    const lastRow = dashboardSheet.getLastRow();
    if (lastRow >= 9) {
      dashboardSheet.getRange(9, 1, lastRow - 8, 15).clearContent();
    }
    dashboardSheet.getRange('B3:B6').setValue(0);
    dashboardSheet.getRange('H3:H5').setValue(0);
    dashboardSheet.getRange('E3:E7').setValue(0);
  }

  // 4. 계좌현황 데이터 영역 클리어 (A8:I 범위)
  const accountSheet = ss.getSheetByName('🏦 계좌현황');
  if (accountSheet) {
    const lastRow = accountSheet.getLastRow();
    if (lastRow >= 8) {
      accountSheet.getRange(8, 1, lastRow - 7, 9).clearContent();
    }
    accountSheet.getRange('B2:B5').setValue(0);
    accountSheet.getRange('B6').setValue('0.00%');
  }
  
  // 5. (생활비 인출 시트 제거됨 — 수익 실현은 다이얼로그로 처리)

  // 6. 모든 백그라운드 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // 7. 스크립트 속성 및 유저 속성(보안 저장소) 초기화
  PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getUserProperties().deleteAllProperties();

  Logger.log('✅ [템플릿 준비 완료] 모든 개인정보 및 거래 기록이 초기화되었습니다.');
}