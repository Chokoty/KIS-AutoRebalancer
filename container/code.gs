/**
 * [KIS Auto-Rebalancer - Container Script]
 *
 * 이 파일을 구글 시트의 Apps Script 프로젝트에 붙여넣기 하세요.
 * 라이브러리 설정: 확장 프로그램 > Apps Script > 라이브러리(+) 에서
 * kis_library_public 의 Script ID를 추가하고 식별자를 'KIS' 로 설정하세요.
 *
 * Script ID: 1LXA06wO7XtQmqqZ4GdnFm6w4bwzl8nrG5dhcE2qc6h0WFcxtxj-OFoc6
 */

function onOpen() {
  KIS.onOpen();
}

// 시스템이 계산해서 채우는 화면 — 사람이 직접 값을 입력하면 안 되는 시트
const READONLY_OUTPUT_SHEETS = [
  '📊 대시보드', '🏦 계좌현황', '📝 거래내역',
  '📝 수익실현기록', '📝 비중변경이력', '📊 기술지표이력'
];

function onEdit(e) {
  if (!e || !e.range) return;
  const ui = SpreadsheetApp.getUi();
  const sheetName = e.range.getSheet().getName();
  const revert = () => e.range.setValue(e.oldValue !== undefined ? e.oldValue : '');

  // 1. 포트폴리오설정: 전체가 팝업 전용
  if (sheetName === '📋 포트폴리오설정') {
    revert();
    ui.alert('⚠️ 직접 편집 불가',
      '포트폴리오 설정은 팝업 창에서만 수정할 수 있습니다.\n\n메뉴 → KIS AutoTrader → 포트폴리오 종목 관리',
      ui.ButtonSet.OK);
    return;
  }

  // 2. ⚙️ 설정: 계좌종류·임계치(B7:B10)만 팝업 전용 — API 키(B2:B5)는 그대로 시트 직접 입력 가능
  if (sheetName === '⚙️ 설정' && e.range.getColumn() === 2 &&
      e.range.getRow() >= 7 && e.range.getRow() <= 10) {
    revert();
    ui.alert('⚠️ 직접 편집 불가',
      '계좌 종류·임계치는 팝업 창에서만 수정할 수 있습니다.\n\n메뉴 → KIS AutoTrader → 설정 및 관리 → 기본 설정',
      ui.ButtonSet.OK);
    return;
  }

  // 3. 시스템 출력 시트: 아예 편집 금지
  if (READONLY_OUTPUT_SHEETS.indexOf(sheetName) !== -1) {
    revert();
    ui.alert('⚠️ 직접 편집 불가',
      '이 시트는 시스템이 자동으로 채우는 화면입니다. 값을 직접 입력하지 마세요.',
      ui.ButtonSet.OK);
  }
}

// 대시보드
function updateDashboard()                  { KIS.updateDashboard(); }
function updateAccountSheet()               { KIS.updateAccountSheet(); }
function executeRebalanceFromDashboard()    { KIS.executeRebalanceFromDashboard(); }
function automatedRefreshRoutine()          { KIS.automatedRefreshRoutine(); }

// 자동화
function scheduledBiWeeklyRebalance()       { KIS.scheduledBiWeeklyRebalance(); }
function toggleHighwayLaneKeeping()         { KIS.toggleHighwayLaneKeeping(); }
function emergencyStopAutomation()          { KIS.emergencyStopAutomation(); }

// 수익 실현
function openWithdrawDialog()               { KIS.openWithdrawDialog(); }
function executeWithdrawPlan(planData)      { return KIS.executeWithdrawPlan(planData); }
function releaseProtectedCash()             { KIS.releaseProtectedCash(); }

// 긴급 대응
function openEmergencySellDialog()          { KIS.openEmergencySellDialog(); }
function executeEmergencySell(code, name, qty, price) { return KIS.executeEmergencySell(code, name, qty, price); }

// 설정
function setupSheets()                      { KIS.setupSheets(); }
function openSecureConfigDialog()           { KIS.openSecureConfigDialog(); }
function saveSecureConfig(data)             { KIS.saveSecureConfig(data); }
function openBasicSettingsDialog()          { KIS.openBasicSettingsDialog(); }
function saveBasicSettings(data)            { KIS.saveBasicSettings(data); }
function forceRefreshToken()                { KIS.forceRefreshToken(); }
function addInitialRatiosColumn()           { KIS.addInitialRatiosColumn(); }

// AI 분석
function openAIBriefingPreview()            { KIS.openAIBriefingPreview(); }
function runAIBriefing()                    { KIS.runAIBriefing(); }
function openAIQuickQuestion()              { KIS.openAIQuickQuestion(); }
function runAIQuickQuestion(q, inclData)              { return KIS.runAIQuickQuestion(q, inclData); }
function runAIQuickQuestionMultiTurn(msgs, inclData)  { return KIS.runAIQuickQuestionMultiTurn(msgs, inclData); }
function applyAIProposedRatios(jsonStr)     { return KIS.applyAIProposedRatios(jsonStr); }
function applyAIProposedRatiosManual(js)    { return KIS.applyAIProposedRatiosManual(js); }
function applyLatestRecommendation()        { KIS.applyLatestRecommendation(); }
function openAIPromptSettings()             { KIS.openAIPromptSettings(); }
function saveSystemPrompt(p)               { return KIS.saveSystemPrompt(p); }
function resetSystemPrompt()               { KIS.resetSystemPrompt(); }
function unlockRatioChange()               { KIS.unlockRatioChange(); }

// 포트폴리오 관리
function openPortfolioManagerDialog()      { KIS.openPortfolioManagerDialog(); }
function searchStockByCode(code)           { return KIS.searchStockByCode(code); }
function savePortfolioSettings(rows)       { return KIS.savePortfolioSettings(rows); }

// 시스템 상태 / 차선유지 설정
function showSystemStatus()                { KIS.showSystemStatus(); }
function showHighwaySettings()             { KIS.showHighwaySettings(); }
function applyHighwaySettings(day, h, off) { KIS.applyHighwaySettings(day, h, off); }

// 안내
function showUsageGuide()                  { KIS.showUsageGuide(); }
function showVersionHistory()              { KIS.showVersionHistory(); }
