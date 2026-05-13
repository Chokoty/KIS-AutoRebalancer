// 포트폴리오 설정 가져오기 (공통 유틸리티)
function getTargetPortfolio() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📋 포트폴리오설정');

  if (!sheet) {
    throw new Error('포트폴리오설정 시트가 없습니다. 초기 설정을 먼저 실행하세요.');
  }

  const lastRow = sheet.getLastRow();
  const portfolio = {};

  for (let i = 3; i <= lastRow; i++) {
    const code = sheet.getRange(i, 1).getValue().toString().trim();
    const name = sheet.getRange(i, 2).getValue();
    const ratio = parseFloat(sheet.getRange(i, 4).getValue()); // D: 운용비율
    const type = sheet.getRange(i, 5).getValue(); // E: 유형
    const baseRatioVal = sheet.getRange(i, 3).getValue(); // C: 기준비율
    const initialRatio = (typeof baseRatioVal === 'number' && baseRatioVal > 0)
      ? baseRatioVal : ratio;

    if (code && ratio > 0 && name !== '현금') {
      portfolio[code] = {
        name: name,
        ratio: ratio,
        initialRatio: initialRatio,
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
      if (diffDays < 13) {
        Logger.log('이번 주는 실행 주기가 아닙니다. (마지막 실행: ' + lastDate.toLocaleDateString() + ')');
        return;
      }
    }

    Logger.log('=== 2주 주기 자동 리밸런싱 시작 ===');

    updateDashboard();
    Utilities.sleep(5000);
    executeRebalanceAutomated();

    props.setProperty('LAST_REBALANCE_DATE', now.getTime().toString());
    Logger.log('리밸런싱 완료 및 시간 기록');

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
    const balance = getBalance();
    const holdings = getHoldings();
    const targetPortfolio = getTargetPortfolio();

    let nonTargetTotal = 0;
    holdings.forEach(h => {
      if (!targetPortfolio[h.code]) nonTargetTotal += h.evalAmount;
    });
    const managedTotal = balance.totalEval - nonTargetTotal;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portfolioSheet = ss.getSheetByName('📋 포트폴리오설정');
    let targetCashRatio = 5;
    if (portfolioSheet) {
      const lastRow = portfolioSheet.getLastRow();
      for (let i = 3; i <= lastRow; i++) {
        if (portfolioSheet.getRange(i, 2).getValue() === '현금') {
          targetCashRatio = parseFloat(portfolioSheet.getRange(i, 4).getValue()) || 5; // D: 운용비율
          break;
        }
      }
    }

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

    executeAutoOrders(sellOrders, buyOrders);
    updateDashboard();

  } catch (e) {
    Logger.log('executeRebalanceAutomated 오류: ' + e.toString());
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
