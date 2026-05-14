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

// 대시보드
function updateDashboard()                  { KIS.updateDashboard(); }
function updateAccountSheet()               { KIS.updateAccountSheet(); }
function executeRebalanceFromDashboard()    { KIS.executeRebalanceFromDashboard(); }
function automatedRefreshRoutine()          { KIS.automatedRefreshRoutine(); }

// 자동화
function scheduledBiWeeklyRebalance()       { KIS.scheduledBiWeeklyRebalance(); }
function toggleHighwayLaneKeeping()         { KIS.toggleHighwayLaneKeeping(); }

// 수익 실현
function openWithdrawDialog()               { KIS.openWithdrawDialog(); }
function executeWithdrawPlan(planData)      { return KIS.executeWithdrawPlan(planData); }
function releaseProtectedCash()             { KIS.releaseProtectedCash(); }

// 설정
function setupSheets()                      { KIS.setupSheets(); }
function openSecureConfigDialog()           { KIS.openSecureConfigDialog(); }
function saveSecureConfig(data)             { KIS.saveSecureConfig(data); }
function forceRefreshToken()                { KIS.forceRefreshToken(); }
function addInitialRatiosColumn()           { KIS.addInitialRatiosColumn(); }

// AI 분석
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

// 안내
function showUsageGuide()                  { KIS.showUsageGuide(); }
function showVersionHistory()              { KIS.showVersionHistory(); }
