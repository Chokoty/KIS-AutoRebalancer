/**
 * 증권사 어댑터 디스패처 (broker-interface)
 *
 * Dashboard/Withdraw/PortfolioManager/code.js 등 상위 로직은 전부 이 4개 함수만
 * 호출한다 — 실제 증권사가 KIS든 나무든 토스든 위 파일들은 바뀔 필요가 없다.
 * 실제 구현은 kis_library_public/brokers/<증권사>Adapter.js 에 둔다.
 *
 * 현재는 KIS만 구현돼 있다. 나무·토스는 브로커 코드는 인식하지만
 * 어댑터가 아직 없어 명시적으로 실패한다(조용히 KIS로 되돌아가지 않음).
 */
const SUPPORTED_BROKERS = ['KIS', '나무', '토스'];

function getBrokerCode() {
  const saved = PropertiesService.getScriptProperties().getProperty('BROKER');
  return SUPPORTED_BROKERS.includes(saved) ? saved : 'KIS';
}

function setBroker(code) {
  if (!SUPPORTED_BROKERS.includes(code)) {
    throw new Error('지원하지 않는 증권사입니다: ' + code + ' (지원: ' + SUPPORTED_BROKERS.join(', ') + ')');
  }
  PropertiesService.getScriptProperties().setProperty('BROKER', code);
}

function getBalance() {
  const broker = getBrokerCode();
  if (broker === 'KIS') return kisGetBalance();
  throw new Error(broker + ' 연동은 아직 준비 중입니다.');
}

function getHoldings() {
  const broker = getBrokerCode();
  if (broker === 'KIS') return kisGetHoldings();
  throw new Error(broker + ' 연동은 아직 준비 중입니다.');
}

function getCurrentPrice(stockCode) {
  const broker = getBrokerCode();
  if (broker === 'KIS') return kisGetCurrentPrice(stockCode);
  throw new Error(broker + ' 연동은 아직 준비 중입니다.');
}

function placeOrder(stockCode, orderType, quantity, price = 0) {
  const broker = getBrokerCode();
  if (broker === 'KIS') return kisPlaceOrder(stockCode, orderType, quantity, price);
  throw new Error(broker + ' 연동은 아직 준비 중입니다.');
}
