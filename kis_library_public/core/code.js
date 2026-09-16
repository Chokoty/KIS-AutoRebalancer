// 예수금 API 응답 전체 확인
function debugBalanceAPI() {
  const config = getConfig();
  const trId = config.isMock ? 'VTTC8434R' : 'TTTC8434R';
  
  const params = {
    'CANO': config.account.split('-')[0],
    'ACNT_PRDT_CD': config.account.split('-')[1],
    'AFHR_FLPR_YN': 'N',
    'OFL_YN': '',
    'INQR_DVSN': '01',
    'UNPR_DVSN': '01',
    'FUND_STTL_ICLD_YN': 'N',
    'FNCG_AMT_AUTO_RDPT_YN': 'N',
    'PRCS_DVSN': '01',
    'CTX_AREA_FK100': '',
    'CTX_AREA_NK100': ''
  };
  
  const data = callKISAPI('/uapi/domestic-stock/v1/trading/inquire-balance', trId, params);
  
  Logger.log('=== output2 전체 ===');
  if (data.output2 && data.output2.length > 0) {
    const output2 = data.output2[0];
    
    Object.keys(output2).forEach(key => {
      Logger.log(`${key}: ${output2[key]}`);
    });
    
    Logger.log('\n=== 핵심 필드 ===');
    Logger.log('dnca_tot_amt (예수금총액): ' + output2.dnca_tot_amt);
    Logger.log('nxdy_excc_amt (익일정산금액): ' + output2.nxdy_excc_amt);
    Logger.log('prvs_rcdl_excc_amt (전일예수금): ' + output2.prvs_rcdl_excc_amt);
    Logger.log('ord_psbl_cash (주문가능금액): ' + output2.ord_psbl_cash);
    Logger.log('pchs_amt_smtl_amt (매입금액합계): ' + output2.pchs_amt_smtl_amt);
    Logger.log('evlu_amt_smtl_amt (평가금액합계): ' + output2.evlu_amt_smtl_amt);
    Logger.log('tot_evlu_amt (총평가금액): ' + output2.tot_evlu_amt);
  }
  
  SpreadsheetApp.getUi().alert('로그를 확인하세요!');
}

// API 파라미터 테스트
function testBalanceAPI() {
  const config = getConfig();
  const token = getAccessToken();
  
  Logger.log('=== 설정 확인 ===');
  Logger.log('계좌번호: ' + config.account);
  Logger.log('모의투자: ' + config.isMock);
  Logger.log('토큰: ' + (token ? '있음' : '없음'));
  
  const trId = config.isMock ? 'VTTC8434R' : 'TTTC8434R';
  
  const params = {
    'CANO': config.account.split('-')[0],
    'ACNT_PRDT_CD': config.account.split('-')[1],
    'AFHR_FLPR_YN': 'N',
    'INQR_DVSN': '01',
    'UNPR_DVSN': '01',
    'FUND_STTL_ICLD_YN': 'N',
    'FNCG_AMT_AUTO_RDPT_YN': 'N',
    'PRCS_DVSN': '01',
    'CTX_AREA_FK100': ' ',
    'CTX_AREA_NK100': ' '
  };
  
  Logger.log('=== 파라미터 ===');
  Logger.log(JSON.stringify(params, null, 2));
  
  const url = `${config.baseUrl}/uapi/domestic-stock/v1/trading/inquire-balance`;
  
  const queryParams = Object.keys(params)
    .filter(key => params[key] !== '' && params[key] !== null && params[key] !== undefined)
    .map(key => `${key}=${encodeURIComponent(params[key])}`)
    .join('&');
  
  const fullUrl = `${url}?${queryParams}`;
  
  Logger.log('=== URL ===');
  Logger.log(fullUrl);
  
  const headers = {
    'Content-Type': 'application/json',
    'authorization': `Bearer ${token}`,
    'appkey': config.appKey,
    'appsecret': config.appSecret,
    'tr_id': trId
  };
  
  const options = {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(fullUrl, options);
  const text = response.getContentText();
  
  Logger.log('=== 응답 ===');
  Logger.log(text);
  
  SpreadsheetApp.getUi().alert('로그를 확인하세요!\n\n확장 프로그램 > Apps Script > 실행 로그');
}

// 보유 주식 상세 확인
function debugHoldings() {
  const config = getConfig();
  const trId = config.isMock ? 'VTTC8434R' : 'TTTC8434R';
  
  const params = {
    'CANO': config.account.split('-')[0],
    'ACNT_PRDT_CD': config.account.split('-')[1],
    'AFHR_FLPR_YN': 'N',
    'OFL_YN': '',
    'INQR_DVSN': '01',
    'UNPR_DVSN': '01',
    'FUND_STTL_ICLD_YN': 'N',
    'FNCG_AMT_AUTO_RDPT_YN': 'N',
    'PRCS_DVSN': '01',
    'CTX_AREA_FK100': '',
    'CTX_AREA_NK100': ''
  };
  
  const data = callKISAPI('/uapi/domestic-stock/v1/trading/inquire-balance', trId, params);
  
  Logger.log('=== 전체 보유 종목 (output1) ===');
  if (data.output1 && Array.isArray(data.output1)) {
    data.output1.forEach((item, index) => {
      Logger.log(`[${index}] 종목코드: ${item.pdno}, 종목명: ${item.prdt_name}, 수량: ${item.hldg_qty}, 평가액: ${item.evlu_amt}`);
    });
  }
  
  Logger.log('=== 예수금 정보 (output2) ===');
  if (data.output2 && Array.isArray(data.output2)) {
    const output2 = data.output2[0];
    Logger.log('dnca_tot_amt (예수금총액): ' + output2.dnca_tot_amt);
    Logger.log('prvs_rcdl_excc_amt (전일예수금): ' + output2.prvs_rcdl_excc_amt);
    Logger.log('nxdy_excc_amt (익일정산금액): ' + output2.nxdy_excc_amt);
  }
  
  SpreadsheetApp.getUi().alert('로그를 확인하세요!');
}

// 주문 API 테스트
function testOrderAPI() {
  const config = getConfig();
  const token = getAccessToken();
  
  const stockCode = '161510'; // PLUS 고배당주
  const quantity = 1;
  
  const trId = config.isMock ? 'VTTC0802U' : 'TTTC0802U';
  const url = `${config.baseUrl}/uapi/domestic-stock/v1/trading/order-cash`;
  
  // 시장가 매수
  const payload = {
    'CANO': config.account.split('-')[0],
    'ACNT_PRDT_CD': config.account.split('-')[1],
    'PDNO': stockCode,
    'ORD_DVSN': '01', // 시장가
    'ORD_QTY': quantity.toString(),
    'ORD_UNPR': '0' // 시장가는 "0"
  };
  
  Logger.log('=== 설정 확인 ===');
  Logger.log('계좌번호: ' + config.account);
  Logger.log('모의투자: ' + config.isMock);
  Logger.log('TR_ID: ' + trId);
  
  Logger.log('=== Payload ===');
  Logger.log(JSON.stringify(payload, null, 2));
  
  const headers = {
    'Content-Type': 'application/json',
    'authorization': `Bearer ${token}`,
    'appkey': config.appKey,
    'appsecret': config.appSecret,
    'tr_id': trId,
    'custtype': 'P'
  };
  
  const options = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    Logger.log('=== 응답 ===');
    Logger.log(responseText);
    
    const data = JSON.parse(responseText);
    Logger.log('=== 파싱된 응답 ===');
    Logger.log('rt_cd: ' + data.rt_cd);
    Logger.log('msg_cd: ' + data.msg_cd);
    Logger.log('msg1: ' + data.msg1);
    
    const ui = SpreadsheetApp.getUi();
    if (data.rt_cd === '0') {
      ui.alert('✅ 주문 성공!\n\n' + data.msg1);
    } else {
      ui.alert('❌ 주문 실패\n\n' + data.msg1);
    }
    
  } catch (e) {
    Logger.log('=== 오류 ===');
    Logger.log(e.toString());
    
    SpreadsheetApp.getUi().alert('❌ 오류\n\n' + e.message);
  }
}

// 포트폴리오 설정 가져오기 (공통 유틸리티)
function getTargetPortfolio() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📋 포트폴리오설정');

  if (!sheet) {
    throw new Error('포트폴리오설정 시트가 없습니다. 초기 설정을 먼저 실행하세요.');
  }

  migratePortfolioToAdjModeIfNeeded(); // D열이 구버전(절대비율) 상태면 조정값으로 자동 변환

  const lastRow = sheet.getLastRow();
  const portfolio = {};
  
  for (let i = 3; i <= lastRow; i++) {
    const code = sheet.getRange(i, 1).getValue().toString().trim();
    const name = sheet.getRange(i, 2).getValue();
    const baseRatio = parseFloat(sheet.getRange(i, 3).getValue()) || 0;
    const adjRatio  = parseFloat(sheet.getRange(i, 4).getValue()) || 0;
    const ratio = baseRatio + adjRatio; // 실제 운용비율 = 기준(C) + 조정(D, +/-)
    const type = sheet.getRange(i, 5).getValue();

    if (code && ratio > 0 && name !== '현금') {
      portfolio[code] = {
        name: name,
        ratio: ratio,
        type: type
      };
    }
  }
  
  Logger.log('목표 포트폴리오: ' + JSON.stringify(portfolio));
  return portfolio;
}

/**
 * 2주 주기 리밸런싱 실행 (자동 트리거용)
 * 13일 간격 체크 로직 포함
 */
function scheduledBiWeeklyRebalance() {
  // 동시 실행 방지: 수동 실행 또는 다른 트리거와 충돌하지 않도록 Lock 획득
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('[스킵] 다른 리밸런싱이 실행 중입니다. (Lock 획득 실패)');
    return;
  }

  try {
    const props = PropertiesService.getDocumentProperties();
    const lastRun = props.getProperty('LAST_REBALANCE_DATE');
    const now = new Date();

    if (lastRun) {
      const lastDate = new Date(parseInt(lastRun));
      const diffDays = (now - lastDate) / (1000 * 60 * 60 * 24);
      if (diffDays < 6) {
        Logger.log('이번 주는 실행 주기가 아닙니다. (마지막 실행: ' + lastDate.toLocaleDateString() + ')');
        return;
      }
    }

    Logger.log('=== 2주 주기 자동 리밸런싱 시작 ===');

    // 1. 대시보드 새로고침 (계산 결과 시트에 기록)
    updateDashboard();

    // 2. 5초 대기 (시트 반영 대기)
    Utilities.sleep(5000);

    // 3. 대시보드 계산 결과를 그대로 실행 (재계산 없이 시트 읽어서 주문)
    const ordersExecuted = executeRebalanceSilently();

    // 4. 실행 시간 기록
    // LAST_REBALANCE_DATE: 항상 갱신 (13일 쿨다운 — 주기 제어용)
    // LAST_AUTO_TRADE_AT: 실제 거래가 있을 때만 갱신 (FSD 8시간 쿨다운 방지)
    props.setProperty('LAST_REBALANCE_DATE', now.getTime().toString());
    if (ordersExecuted > 0) {
      props.setProperty('LAST_AUTO_TRADE_AT', now.toISOString());
      Logger.log('리밸런싱 완료: ' + ordersExecuted + '건 주문 실행');
    } else {
      Logger.log('리밸런싱 완료: 주문 없음 (임계치 미달 또는 자산 균형)');
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * 자동 리밸런싱 실행 (확인창 없이 매매 진행)
 */
function executeRebalanceAutomated() {
  Logger.log('=== 리밸런싱 로직 계산 시작 ===');
  
  try {
    // 1. 데이터 수집
    const balance = getBalance();
    const holdings = getHoldings();
    const targetPortfolio = getTargetPortfolio();
    
    // 2. 실질 운용 자산 계산 (목표외 종목 제외)
    let nonTargetTotal = 0;
    holdings.forEach(h => {
      if (!targetPortfolio[h.code]) nonTargetTotal += h.evalAmount;
    });
    const managedTotal = balance.totalEval - nonTargetTotal;
    
    // 3. 목표 현금 비율 조회
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portfolioSheet = ss.getSheetByName('📋 포트폴리오설정');
    let targetCashRatio = 5;
    if (portfolioSheet) {
      const lastRow = portfolioSheet.getLastRow();
      for (let i = 3; i <= lastRow; i++) {
        if (portfolioSheet.getRange(i, 2).getValue() === '현금') {
          const baseCash = parseFloat(portfolioSheet.getRange(i, 3).getValue()) || 0;
          const adjCash  = parseFloat(portfolioSheet.getRange(i, 4).getValue()) || 0;
          targetCashRatio = (baseCash + adjCash) || 5;
          break;
        }
      }
    }
    
    // 4. 리밸런싱 계획 계산
    const config = getConfig();
    const rebalanceResult = calculateRebalancePlan(managedTotal, balance, holdings, targetPortfolio, {
      targetCashRatio: targetCashRatio,
      tolerance: config.rebalanceTolerance
    });
    
    const { sellOrders, buyOrders } = rebalanceResult;
    
    Logger.log(`매도 대상: ${sellOrders.length}건, 매수 대상: ${buyOrders.length}건`);
    
    if (sellOrders.length === 0 && buyOrders.length === 0) {
      Logger.log('리밸런싱이 필요하지 않습니다. (임계치 미달 또는 자산 균형)');
      return;
    }
    
    // 5. 실행
    executeAutoOrders(sellOrders, buyOrders);
    
    // 6. 대시보드 업데이트 (선택 사항)
    updateDashboard();
    
  } catch (e) {
    Logger.log('scheduledBiWeeklyRebalance 오류: ' + e.toString());
  }
}

/**
 * 실제 주문 실행 및 로그 기록 (자동화용)
 */
function executeAutoOrders(sellList, buyList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('📝 거래내역');
  if (!logSheet) {
    Logger.log('[오류] "📝 거래내역" 시트를 찾을 수 없어 자동 주문을 중단합니다.');
    return;
  }

  // 매도 먼저
  for (const order of sellList) {
    Logger.log(`[자동매도] ${order.name} 실행 중...`);
    const result = placeOrder(order.code, 'sell', order.quantity, 0);
    try {
      logSheet.appendRow([
        new Date(), '매도(자동)', order.code, order.name,
        order.quantity, order.price, order.quantity * order.price,
        result.success ? '성공' : '실패', result.message
      ]);
    } catch (logErr) {
      Logger.log('[경고] 거래내역 기록 실패: ' + logErr.toString());
    }
    Utilities.sleep(500);
  }

  // 매수 실행
  for (const order of buyList) {
    Logger.log(`[자동매수] ${order.name} 실행 중...`);
    const result = placeOrder(order.code, 'buy', order.quantity, 0);
    try {
      logSheet.appendRow([
        new Date(), '매수(자동)', order.code, order.name,
        order.quantity, order.price, order.quantity * order.price,
        result.success ? '성공' : '실패', result.message
      ]);
    } catch (logErr) {
      Logger.log('[경고] 거래내역 기록 실패: ' + logErr.toString());
    }
    Utilities.sleep(500);
  }
}
