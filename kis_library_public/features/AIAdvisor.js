/**
 * 시스템 기본 프롬프트 반환 (커스텀 저장된 것이 있으면 그것을 반환)
 */
function getSystemPrompt() {
  const customPrompt = PropertiesService.getUserProperties().getProperty('CUSTOM_AI_PROMPT');
  if (customPrompt) return customPrompt;

  return `너는 10년차 프로 기술적 분석 트레이더이자 자산 배분 전략가야.
아래 제공되는 포트폴리오 데이터와 각 종목의 기술적 지표 수치를 분석하고, Google 검색으로 최신 뉴스·X 동향까지 종합하여 비중 조정을 제안해줘.

[분석 스타일]:
⚖️ 균형형 → 리스크와 수익의 균형을 유지하되, 시장 상황에 맞게 자유롭게 판단해.

[종목별 기술적 분석 지침 — 아래 순서로 판단]:
1. **RSI(힘)**: 50선 위치, 과매수(>70)/과매도(<30), 다이버전스 여부
2. **MACD(방향)**: 히스토그램 양/음, MACD선과 Signal선 크로스, 제로라인 위치
3. **볼린저밴드(범위)**: 현재가가 상단/중간/하단 중 어디 위치, 밴드 수축/확장
4. **스토캐스틱(단기)**: %K/%D 위치, 20·80 구간 진입 여부
5. **거래량(무게)**: 평균 대비 배율, 가격 방향과 거래량 일치 여부

[컨플루언스 판단]:
- 3개 이상 지표가 같은 방향이면 강한 신호로 비중 조정에 적극 반영
- 지표가 상충되면 보수적으로 판단

[포트폴리오 지침]:
- 비중 조정의 유일한 근거는 기술적 지표(RSI, MACD, 볼린저밴드, 스토캐스틱, 거래량)다. 개별 종목의 수익률이 높다는 이유만으로 매도 비중을 낮추지 말 것
- 기준비중이 목표 기준이다. 기술적 신호가 없으면 기준비중 유지가 기본값
- **신호 강도별 비중 조정 범위** (기준비중 대비):
    • 신호 없음 (지표 혼재/중립): 기준비중 ±3%p 이내 유지
    • 매도 우호 (과열 신호 1~2개): 기준비중 대비 -5%p까지 축소
    • 매도 강신호 (과열 신호 3개 이상 컨플루언스): 기준비중 대비 -10%p까지 축소
    • 매수 우호 (과매도 신호 1~2개): 기준비중 대비 +5%p까지 확대
    • 매수 강신호 (과매도 신호 3개 이상 컨플루언스): 기준비중 대비 +10%p까지 확대
- 단일 종목 최대 비중: 40% 이하 (분산 원칙 유지)
- 급격한 전량 매도(0%) 또는 극단 집중은 지양. 기존 비중에서 점진적으로 조정하는 방향 선호
- **기준비중 drift 제한 (최우선)**: 각 종목 제안 비중은 기준비중 ±15%p 이내로 제한. 단기 기술 과열만으로 기준비중에서 크게 이탈하는 제안 금지
- **현금 비중 규칙**:
    • 기본: 현금 기준비중의 3배 이내 (예: 기준 5% → 최대 15%)
    • 포트폴리오 내 **강한 매도 신호(score ≤ -0.6) 종목이 절반 이상**이면: 현금 기준비중의 4배까지 허용 (예: 기준 5% → 최대 20%)
    • 현금 확대 시 각 종목 비중을 신호 강도에 비례해 균등 축소. 한 종목만 집중 매도 금지
- 분석 결과에 반드시 포함: 각 종목의 현재 운용비중이 기준비중 대비 얼마나 drift(이탈)되었는지
- 조정 근거는 반드시 실제 지표 수치(RSI 값, MACD 방향 등)를 인용해서 납득 가능하게 작성
- **용어 구분 매우 중요**:
    • **현금**: 포트폴리오 비중 카테고리 (예: 현금 비중 5%) — JSON에 code “CASH”, 본문에 “현금”으로 표기
    • **예수금**: 실제 계좌에 있는 돈의 금액 (예: 632,034원) — 비중 조정 제안에는 사용 금지
    • 비중 변화를 말할 때는 반드시 “현금” 사용. “예수금: 12% → 25%” 같이 쓰지 말 것 (예수금은 비율이 아닌 절대 금액)

[출력 형식]:
- 시장 요약 2~3줄 (Google Search 기반 최신 뉴스 포함)
- 종목별: “현재 비중 → 조정 비중 | 핵심 근거 1줄 (지표 수치 인용)”
- 전체 합산 100% JSON 포함

**중요: 응답은 아래 양식을 정확히 지켜서 마크다운으로 깔끔하게 출력하고, 문서는 가급적 짧고 굵게 써!**
마지막에는 반드시 지정된 형식의 JSON 데이터를 포함해야 해 (전체 합산 100%).

[출력 양식 예시]:
### 🌍 시장 요약
- 시장 뷰 요약 1줄
- 시장 뷰 요약 1줄

### 🎯 비중 조정 제안
- **[종목명]**: 현재 비중% ➡️ **조정 비중%** | RSI 34.2 과매도 + MACD 골든크로스 → 매수 우호적 (근거 예시처럼 실제 수치 인용)

[ALLOCATION_START] {“ratios”: [{“code”: “코드(현금은 'CASH')”, “ratio”: 숫자, “rationale”: “간략한 이유”}], “summary”: “전체 요약”} [ALLOCATION_END]

“이것은 투자 조언이 아닙니다. 스스로 판단하세요.”`;
}

/**
 * 사용자 프롬프트 설정창 열기
 */
function openAIPromptSettings() {
  const currentPrompt = getSystemPrompt();
  const html = `
    <html>
      <head>
        <style>
          body { font-family: 'Malgun Gothic', sans-serif; padding: 20px; color: #3c4043; }
          textarea { width: 100%; height: 400px; font-family: monospace; font-size: 13px; padding: 12px; border: 1px solid #dadce0; border-radius: 4px; line-height: 1.5; }
          .btn-container { display: flex; gap: 10px; margin-top: 20px; }
          .btn { flex: 1; padding: 12px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
          .btn-save { background: #4285F4; color: white; }
          .btn-reset { background: white; border: 1px solid #dadce0; color: #5f6368; }
          .hint { font-size: 12px; color: #5f6368; margin-bottom: 12px; }
          code { background: #f1f3f4; padding: 2px 4px; border-radius: 4px; }
        </style>
      </head>
      <body>
        <h3>🤖 AI 시스템 프롬프트 설정</h3>
        <p class="hint">분석 및 비중 제안 시 AI에게 전달되는 지침입니다. <code>{{TARGET_YIELD}}</code> 등은 자동 치환됩니다.</p>
        <textarea id="promptArea">${currentPrompt}</textarea>
        <div class="btn-container">
          <button class="btn btn-reset" onclick="reset()">초기화</button>
          <button class="btn btn-save" onclick="save()">저장하기</button>
        </div>
        <script>
          function save() {
            const prompt = document.getElementById('promptArea').value;
            google.script.run
              .withSuccessHandler(() => { alert('저장되었습니다.'); google.script.host.close(); })
              .saveSystemPrompt(prompt);
          }
          function reset() {
            if (!confirm('기본 프롬프트로 초기화하시겠습니까?')) return;
            google.script.run
              .withSuccessHandler(() => { alert('초기화되었습니다.'); google.script.host.close(); })
              .resetSystemPrompt();
          }
        </script>
      </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(600).setHeight(620).setTitle('AI 프롬프트 설정');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}

/**
 * 프롬프트 저장 (사용자 속성)
 */
function saveSystemPrompt(prompt) {
  PropertiesService.getUserProperties().setProperty('CUSTOM_AI_PROMPT', prompt);
}

/**
 * 프롬프트 초기화
 */
function resetSystemPrompt() {
  PropertiesService.getUserProperties().deleteProperty('CUSTOM_AI_PROMPT');
}

/**
 * 포트폴리오설정 C열(기준비율)이 비어 있으면 D열(운용비율)로 초기화.
 * @returns {boolean} 새로 초기화했으면 true
 */
function ensureInitialRatios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return false;
  const data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
  // col index: 0=종목코드, 1=종목명, 2=기준비율(C/고정), 3=운용비율(D), 4=유형(E)
  const needsInit = data.some(row => !(typeof row[2] === 'number' && row[2] > 0));
  if (!needsInit) return false;
  const updates = data.map(row => [
    (typeof row[2] === 'number' && row[2] > 0) ? row[2] : row[3]
  ]);
  sheet.getRange(3, 3, updates.length, 1).setValues(updates);
  return true;
}

/**
 * 포트폴리오설정 시트에서 종목코드 → 기준비율 맵을 반환
 * col index: 0=종목코드, 1=종목명, 2=기준비율(C/고정), 3=운용비율(D), 4=유형(E)
 */
function getInitialRatioMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  const map = {};
  if (!sheet) return map;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return map;
  const data = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
  data.forEach(row => {
    const code = String(row[0]).trim();
    const name = String(row[1]).trim();
    const key = code || 'CASH';
    const val = (typeof row[2] === 'number' && row[2] > 0) ? row[2] : row[3];
    map[key] = val;
    if (name === '현금') map['CASH'] = val;
  });
  return map;
}

/**
 * 대시보드 현재 상태를 AI를 위한 텍스트로 요약 (수익률 갭 + 상세 기술지표 포함)
 */
/**
 * 미리보기용 경량 포트폴리오 현황 — KIS API 미호출, 시트 값만 읽음 (빠름)
 */
function getDashboardBasicState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName('📊 대시보드');
  if (!dash) return "대시보드 데이터를 찾을 수 없습니다.";

  const totalEval    = dash.getRange('B3').getValue();
  const cash         = dash.getRange('B4').getValue();
  const lastUpdate   = dash.getRange('B6').getDisplayValue();
  const currentReturn = parseFloat((dash.getRange('H5').getValue() || '0').toString().replace('%', ''));
  const targetYield  = parseFloat((dash.getRange('H6').getValue() || 0)) * 100;

  const lastRow = dash.getLastRow();
  if (lastRow < 9) return `총 평가액: ${totalEval}, 업데이트: ${lastUpdate}`;

  const data = dash.getRange(9, 1, lastRow - 8, 10).getValues();
  const initialRatioMap = getInitialRatioMap();
  const gap = targetYield - currentReturn;

  let report = `[포트폴리오 현황 (기준: ${lastUpdate})]\n`;
  report += `총 평가액: ${totalEval.toLocaleString()}원 | 예수금: ${cash.toLocaleString()}원\n`;
  report += `현재 수익률: ${currentReturn >= 0 ? '+' : ''}${currentReturn.toFixed(2)}% | 목표: ${targetYield.toFixed(1)}% | 갭: ${gap >= 0 ? '+' : ''}${gap.toFixed(2)}%p\n`;
  report += `\n[종목별 현황] (현재비중 → 운용비중 | 기준비중)\n`;

  data.forEach(row => {
    const code = String(row[0]).trim();
    const name = row[1];
    if (!code) return;
    const evalAmount   = row[5];
    const currentRatio = row[6];
    const targetRatio  = row[7];
    const diff         = row[9];
    const initialRatio = initialRatioMap[code];
    const initialStr   = initialRatio != null ? ` | 기준: ${initialRatio}%` : '';
    report += `▶ ${name} (${code}): 현재 ${currentRatio} → 목표 ${targetRatio}${initialStr} (차이 ${diff}) | ${evalAmount.toLocaleString()}원\n`;
  });

  report += `\n※ 기술지표(RSI, MACD 등)는 AI 분석 실행 시 자동 포함됩니다.`;
  return report;
}

function getDashboardStateForAI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName('📊 대시보드');
  if (!dash) return "대시보드 데이터를 찾을 수 없습니다.";

  const totalEval     = dash.getRange('B3').getValue();
  const cash          = dash.getRange('B4').getValue();
  const lastUpdate    = dash.getRange('B6').getDisplayValue();
  const currentReturn = parseFloat((dash.getRange('H5').getValue() || '0').toString().replace('%', ''));
  const targetYield   = parseFloat((dash.getRange('H6').getValue() || 0)) * 100;

  const lastRow = dash.getLastRow();
  if (lastRow < 9) return `현재 총자산: ${totalEval}, 업데이트: ${lastUpdate}`;

  const data = dash.getRange(9, 1, lastRow - 8, 10).getValues();
  const gap = targetYield - currentReturn;
  const initialRatioMap = getInitialRatioMap();

  // 포트폴리오설정에서 현금 목표비중 조회
  let targetCashRatio = 5;
  try {
    const pSheet = ss.getSheetByName('📋 포트폴리오설정');
    if (pSheet) {
      const pData = pSheet.getRange(3, 1, Math.max(1, pSheet.getLastRow() - 2), 4).getValues();
      pData.forEach(r => { if (String(r[1]).trim() === '현금') targetCashRatio = parseFloat(r[3]) || 5; });
    }
  } catch(e) {}

  let report = `[포트폴리오 현황 (기준: ${lastUpdate})]\n`;
  report += `총 평가액: ${totalEval.toLocaleString()}원\n`;
  report += `예수금: ${cash.toLocaleString()}원 (※ 매수 대기 자금 — 현재비중 계산에서 제외. 리밸런싱 후 현금 목표비중 ${targetCashRatio}%만 유지되도록 설계됨)\n`;
  report += `현재 수익률: ${currentReturn >= 0 ? '+' : ''}${currentReturn.toFixed(2)}% | 목표: ${targetYield.toFixed(1)}% | 갭: ${gap >= 0 ? '+' : ''}${gap.toFixed(2)}%p (${gap > 0 ? '목표 미달' : '목표 초과'})\n`;
  report += `현금(CASH): 목표비중 ${targetCashRatio}% — 리밸런싱 시스템이 자동 관리. 예수금은 매수 대기 자금이므로 현재비중 계산 불필요. 비중 제안 시 현금은 목표비중(${targetCashRatio}%) 기준으로만 판단할 것.\n`;
  report += `\n[종목별 현황 + 기술적 지표 상세] (현재비중 → 운용비중 | 기준비중)\n`;

  data.forEach(row => {
    const code = String(row[0]).trim();
    const name = String(row[1] || '').trim();
    // 현금 행 제외 — 예수금은 AI 비중 판단 대상이 아님
    if (!code) return;
    if (name === '현금' || name.includes('현금')) return;

    const evalAmount    = row[5];
    const currentRatio  = row[6];
    const targetRatio   = row[7];
    const diff          = row[9];
    const initialRatio  = initialRatioMap[code];
    const initialStr    = initialRatio != null ? ` | 기준: ${initialRatio}%` : '';

    report += `\n▶ ${name} (${code})\n`;
    report += `  비중: 현재 ${currentRatio} → 목표 ${targetRatio}${initialStr} (차이 ${diff}) | 평가액: ${evalAmount.toLocaleString()}원\n`;

    try {
      const ta = getConfluenceScore(code);
      const s  = ta.signals || {};

      report += `  [컨플루언스 Score] ${ta.score > 0 ? '+' : ''}${ta.score} → ${ta.summary}\n`;

      if (s.rsi)    report += `  RSI(14): ${s.rsi.value} (${s.rsi.value < 30 ? '과매도' : s.rsi.value > 70 ? '과매수' : '중립'})\n`;
      if (s.macd)   report += `  MACD: ${s.macd.macd} / Signal: ${s.macd.signalLine} / Histogram: ${s.macd.histogram} (${s.macd.histogram > 0 ? '양봉-상승' : '음봉-하락'})\n`;
      if (s.bb)     report += `  볼린저밴드: 상단 ${s.bb.upper} / 중간 ${s.bb.middle} / 하단 ${s.bb.lower} | 현재가 ${s.bb.price} (${s.bb.signal >= 0.5 ? '하단 근접-과매도' : s.bb.signal <= -0.5 ? '상단 근접-과매수' : '중간대'})\n`;
      if (s.stoch)  report += `  스토캐스틱: %K ${s.stoch.k} / %D ${s.stoch.d} (${s.stoch.k < 20 ? '과매도구간' : s.stoch.k > 80 ? '과매수구간' : '중립구간'})\n`;
      if (s.volume) report += `  거래량: 평균 대비 ${s.volume.ratio}배 (${s.volume.signal > 0 ? '매수세 우위' : s.volume.signal < 0 ? '매도세 우위' : '평범'})\n`;
    } catch (e) {
      report += `  [기술지표 조회 실패]\n`;
    }
  });

  return report;
}

/**
 * AI 분석 실행 시 기술지표값을 '📊 기술지표이력' 시트에 기록 (매 실행마다 덮어쓰기)
 */
function recordTAHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('📊 기술지표이력');
  const HEADERS = ['날짜', '종목명', '종목코드', 'Score', '📍 판정', 'RSI', 'MACD', 'MACD Signal', 'Histogram', 'BB상단', 'BB중간', 'BB하단', 'Stoch %K', 'Stoch %D', '거래량비율', '신호요약'];
  const NCOL = HEADERS.length;

  if (!sheet) {
    sheet = ss.insertSheet('📊 기술지표이력');
    sheet.getRange(1, 1, 1, NCOL).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, NCOL).setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  } else {
    // 헤더가 구버전이면 갱신 (15→16 컬럼 마이그레이션)
    const curHeader = sheet.getRange(1, 1, 1, NCOL).getValues()[0];
    if (curHeader[4] !== '📍 판정') {
      sheet.clear();
      sheet.getRange(1, 1, 1, NCOL).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, NCOL).setFontWeight('bold').setBackground('#e8f0fe');
      sheet.setFrozenRows(1);
    }
  }

  // 기존 데이터 삭제 (헤더 제외)
  const prevLast = sheet.getLastRow();
  if (prevLast > 1) sheet.getRange(2, 1, prevLast - 1, NCOL).clearContent();

  const portfolioSheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!portfolioSheet) return;

  const portLastRow = portfolioSheet.getLastRow();
  const now = new Date();
  const rows = [];

  for (let i = 3; i <= portLastRow; i++) {
    const code = String(portfolioSheet.getRange(i, 1).getValue()).trim();
    const name = portfolioSheet.getRange(i, 2).getValue();
    if (!code || name === '현금') continue;

    try {
      const ta = getConfluenceScore(code);
      const s  = ta.signals || {};
      rows.push([
        now, name, code, ta.score,
        getScoreVerdict(ta.score),
        s.rsi    ? s.rsi.value      : '',
        s.macd   ? s.macd.macd      : '',
        s.macd   ? s.macd.signalLine : '',
        s.macd   ? s.macd.histogram : '',
        s.bb     ? s.bb.upper       : '',
        s.bb     ? s.bb.middle      : '',
        s.bb     ? s.bb.lower       : '',
        s.stoch  ? s.stoch.k        : '',
        s.stoch  ? s.stoch.d        : '',
        s.volume ? s.volume.ratio   : '',
        ta.summary
      ]);
    } catch (e) {
      Logger.log('[TA 이력] ' + code + ' 기록 실패: ' + e.toString());
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, NCOL).setValues(rows);
    // 판정 컬럼 색상 강조
    const verdictRange = sheet.getRange(2, 5, rows.length, 1);
    verdictRange.setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setColumnWidth(5, 130);   // 판정
    sheet.setColumnWidth(NCOL, 380); // 신호요약
  }
}

/**
 * AI 비중 제안 실행 전 프롬프트/데이터 확인 팝업
 */
function openAIBriefingPreview() {
  const config = getConfig();
  if (!config.geminiApiKey) {
    SpreadsheetApp.getUi().alert('Gemini API Key가 필요한 기능입니다.');
    return;
  }
  const promptJsonStr = JSON.stringify(getSystemPrompt());
  const html = `<!DOCTYPE html><html><head><style>
*{box-sizing:border-box;}
body{font-family:'Malgun Gothic',sans-serif;padding:16px 20px;color:#3c4043;margin:0;font-size:13px;line-height:1.5;}
.title{font-size:16px;font-weight:bold;color:#4285F4;padding-bottom:10px;border-bottom:2px solid #e8eaed;margin-bottom:14px;}
.label{font-size:11px;font-weight:bold;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;}
.hint{font-size:11px;color:#9aa0a6;margin-top:5px;}
textarea{width:100%;border:1px solid #dadce0;border-radius:4px;padding:10px;font-family:monospace;font-size:11.5px;resize:vertical;line-height:1.4;}
.section{margin-bottom:14px;}
.toggle-btn{background:white;border:1px solid #dadce0;border-radius:4px;padding:5px 12px;cursor:pointer;font-size:12px;color:#5f6368;}
.toggle-btn:hover{background:#f1f3f4;}
.spinner{color:#5f6368;font-size:12px;padding:6px 0;}
.btn-row{display:flex;gap:10px;margin-top:16px;}
.btn{flex:1;padding:11px;border:none;border-radius:6px;cursor:pointer;font-weight:bold;font-size:13px;}
.btn-primary{background:#4285F4;color:white;}
.btn-primary:hover{background:#1967d2;}
.btn-secondary{background:white;border:1px solid #dadce0;color:#5f6368;}
</style></head><body>
<div class="title">🤖 AI 비중 제안 — 실행 전 확인</div>
<div class="section">
  <div class="label">📝 시스템 프롬프트 (수정 가능)</div>
  <textarea id="promptArea" rows="12"></textarea>
  <div class="hint">수정하면 이번 분석에만 적용됩니다. 영구 저장은 ⚙️설정 → AI 프롬프트 설정에서 하세요.</div>
</div>
<div class="section">
  <button class="toggle-btn" onclick="toggleData()" id="toggleBtn">📊 현재 포트폴리오 데이터 확인 ▼</button>
  <div id="dataSection" style="display:none;margin-top:6px;">
    <div class="spinner" id="dataSpinner">⏳ 데이터 불러오는 중...</div>
    <textarea id="dataArea" rows="9" readonly style="display:none;"></textarea>
  </div>
</div>
<div class="btn-row">
  <button class="btn btn-secondary" onclick="google.script.host.close()">취소</button>
  <button class="btn btn-primary" id="startBtn" onclick="startAnalysis()">🚀 분석 시작</button>
</div>
<script>
var dataLoaded=false,dataExpanded=false;
document.getElementById('promptArea').value=${promptJsonStr};
function toggleData(){
  dataExpanded=!dataExpanded;
  document.getElementById('dataSection').style.display=dataExpanded?'block':'none';
  document.getElementById('toggleBtn').textContent=dataExpanded?'📊 현재 포트폴리오 데이터 숨기기 ▲':'📊 현재 포트폴리오 데이터 확인 ▼';
  if(dataExpanded&&!dataLoaded){
    dataLoaded=true;
    google.script.run
      .withSuccessHandler(function(d){
        document.getElementById('dataSpinner').style.display='none';
        var ta=document.getElementById('dataArea');ta.value=d;ta.style.display='block';
      })
      .withFailureHandler(function(e){document.getElementById('dataSpinner').textContent='❌ 로드 실패: '+e.message;})
      .getDashboardBasicState();
  }
}
function startAnalysis(){
  var btn=document.getElementById('startBtn');
  btn.disabled=true;
  btn.textContent='⏳ 분석 요청 중...';
  document.getElementById('promptArea').disabled=true;
  google.script.run
    .withSuccessHandler(function(){google.script.host.close();})
    .withFailureHandler(function(){google.script.host.close();})
    .runAIBriefing(document.getElementById('promptArea').value);
}
<\/script></body></html>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(680).setHeight(580).setTitle('AI 비중 제안');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}

/**
 * AI 빠른 질문 팝업 — 짧은 질문에 즉시 답변
 */
function openAIQuickQuestion() {
  const config = getConfig();
  if (!config.geminiApiKey) {
    SpreadsheetApp.getUi().alert('Gemini API Key가 필요한 기능입니다.');
    return;
  }
  const html = `<!DOCTYPE html><html><head><style>
*{box-sizing:border-box;}
html,body{height:100%;margin:0;padding:0;}
body{font-family:'Malgun Gothic',sans-serif;color:#3c4043;font-size:13px;line-height:1.5;display:flex;flex-direction:column;}
.hdr{font-size:16px;font-weight:bold;color:#4285F4;padding:14px 20px 10px;border-bottom:2px solid #e8eaed;flex-shrink:0;}
#inputSection{padding:14px 20px 16px;flex:1;display:flex;flex-direction:column;gap:6px;}
.lbl{font-size:11px;font-weight:bold;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;}
.hint{font-size:11px;color:#9aa0a6;}
textarea{width:100%;border:1px solid #dadce0;border-radius:4px;padding:9px;font-family:inherit;font-size:13px;line-height:1.4;}
#q{flex:1;min-height:80px;resize:vertical;}
.cb-row{display:flex;align-items:center;gap:8px;font-size:13px;}
.btn-row{display:flex;gap:10px;margin-top:4px;}
.btn{flex:1;padding:10px;border:none;border-radius:6px;cursor:pointer;font-weight:bold;font-size:13px;}
.bp{background:#4285F4;color:white;}.bp:disabled{opacity:.5;cursor:not-allowed;}
.bs{background:white;border:1px solid #dadce0;color:#5f6368;}
#chatSection{display:none;flex:1;flex-direction:column;min-height:0;}
.chat-log{flex:1;overflow-y:auto;padding:10px 16px;background:#fafafa;min-height:0;}
.chat-user{margin:8px 0;padding:10px 12px;background:#e8f0fe;border-radius:8px;border-left:3px solid #4285F4;}
.chat-ai{margin:8px 0;padding:10px 12px;background:#f8f9fa;border-radius:8px;border-left:3px solid #34a853;}
.chat-lbl{font-size:11px;font-weight:bold;color:#5f6368;margin-bottom:4px;}
.chat-txt{white-space:pre-wrap;font-size:13px;line-height:1.6;}
.chat-foot{flex-shrink:0;border-top:1px solid #e8eaed;padding:10px 16px 14px;}
#followQ{width:100%;resize:none;margin-bottom:8px;}
</style></head><body>
<div class="hdr">💬 AI에게 질문하기</div>
<div id="inputSection">
  <div class="lbl">질문 내용</div>
  <textarea id="q" placeholder="예: 지금 S&P500 비중 늘려도 괜찮을까요? 요즘 금 시장 상황은?"></textarea>
  <div class="hint">간단한 시황 질문도 가능합니다.</div>
  <div class="cb-row"><input type="checkbox" id="inclData" checked><label for="inclData">📊 현재 포트폴리오 데이터 포함 (데이터 기반 질문 권장)</label></div>
  <div class="btn-row">
    <button class="btn bs" onclick="google.script.host.close()">닫기</button>
    <button class="btn bp" id="firstBtn" onclick="ask()">💬 질문하기</button>
  </div>
</div>
<div id="chatSection">
  <div class="chat-log" id="chatLog"></div>
  <div class="chat-foot">
    <textarea id="followQ" rows="2" placeholder="이어서 질문하기..."></textarea>
    <div class="btn-row">
      <button class="btn bs" onclick="google.script.host.close()">닫기</button>
      <button class="btn bp" id="sendBtn" onclick="send()">💬 전송</button>
    </div>
  </div>
</div>
<script>
var hist=[],inclFlag=false,MAX=10;
function ask(){
  var q=document.getElementById('q').value.trim();
  if(!q){alert('질문을 입력해주세요.');return;}
  inclFlag=document.getElementById('inclData').checked;
  hist=[{role:'user',text:q}];
  var btn=document.getElementById('firstBtn');
  btn.disabled=true;btn.textContent='⏳ 질문 중...';
  google.script.run
    .withSuccessHandler(function(r){hist.push({role:'model',text:r});showChat();})
    .withFailureHandler(function(e){btn.disabled=false;btn.textContent='💬 질문하기';alert('오류: '+e.message);})
    .runAIQuickQuestion(hist,inclFlag);
}
function showChat(){
  document.getElementById('inputSection').style.display='none';
  var cs=document.getElementById('chatSection');cs.style.display='flex';
  document.getElementById('chatLog').innerHTML='';
  bubble('user',hist[0].text);bubble('model',hist[1].text);scrollBot();
}
function send(){
  var q=document.getElementById('followQ').value.trim();
  if(!q)return;
  hist.push({role:'user',text:q});
  bubble('user',q);
  document.getElementById('followQ').value='';
  var btn=document.getElementById('sendBtn');
  btn.disabled=true;btn.textContent='⏳';
  var lid='ld_'+Date.now();
  bubble('model','⏳ 응답 중...',lid);scrollBot();
  var msgs=hist.slice(-MAX);
  while(msgs.length&&msgs[0].role!=='user')msgs=msgs.slice(1);
  google.script.run
    .withSuccessHandler(function(r){
      var el=document.getElementById(lid);if(el)el.remove();
      hist.push({role:'model',text:r});bubble('model',r);
      btn.disabled=false;btn.textContent='💬 전송';scrollBot();
    })
    .withFailureHandler(function(e){
      var el=document.getElementById(lid);if(el)el.remove();
      hist.pop();document.getElementById('followQ').value=q;btn.disabled=false;btn.textContent='💬 전송';alert('오류: '+e.message);
    })
    .runAIQuickQuestion(msgs,inclFlag);
}
function bubble(role,text,id){
  var log=document.getElementById('chatLog'),d=document.createElement('div');
  d.className=role==='user'?'chat-user':'chat-ai';if(id)d.id=id;
  d.innerHTML='<div class="chat-lbl">'+(role==='user'?'나':'🤖 AI')+'</div><div class="chat-txt">'+(role==='user'?esc(text):md(text))+'</div>';
  log.appendChild(d);
}
function esc(t){return t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\\n/g,'<br>');}
function md(t){
  t=t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  t=t.replace(/\\*\\*([^*]+)\\*\\*/g,'<strong>$1</strong>');
  var lines=t.split('\\n');
  for(var i=0;i<lines.length;i++){
    var l=lines[i];
    if(/^#{1,3}\\s+/.test(l)){lines[i]='<strong>'+l.replace(/^#+\\s+/,'')+'</strong>';}
    else if(/^\\*\\s+/.test(l)){lines[i]='• '+l.replace(/^\\*\\s+/,'');}
    else if(/^-\\s+/.test(l)){lines[i]='• '+l.replace(/^-\\s+/,'');}
  }
  return lines.join('<br>');}
function scrollBot(){var l=document.getElementById('chatLog');l.scrollTop=l.scrollHeight;}
<\/script></body></html>`;
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(620).setHeight(600),
    ' '
  );
}

/**
 * AI 빠른 질문 처리 (서버 사이드)
 * question이 배열인 경우 멀티턴으로 처리 (구버전 컨테이너 하위 호환)
 */
function runAIQuickQuestion(question, includeData) {
  if (Array.isArray(question)) {
    return runAIQuickQuestionMultiTurn(question, includeData !== false);
  }

  const config = getConfig();
  if (!config.geminiApiKey) throw new Error('Gemini API Key가 설정되지 않았습니다.');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.geminiModelId}:generateContent?key=${config.geminiApiKey}`;

  const systemContext = `너는 한국 주식·ETF 포트폴리오 투자 전문가야. 사용자는 한국투자증권 API로 자동 리밸런싱을 운용 중이며, 국내주식·해외주식·채권·금·달러·현금을 분산하는 자산배분 포트폴리오를 관리하고 있어. 질문에 한국어로 간결하고 실용적으로 답변해줘. 투자 조언이 아닌 참고용 분석임을 명심해.`;

  let userPrompt = question;
  if (includeData) {
    const dashState = getDashboardStateForAI();
    userPrompt = `${question}\n\n[현재 포트폴리오 현황]:\n${dashState}`;
  }

  const payload = {
    system_instruction: { parts: [{ text: systemContext }] },
    contents: [{ parts: [{ text: userPrompt }] }],
    tools: [{ googleSearch: {} }]
  };
  const options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  let response, resJson;
  for (var attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);
    resJson = JSON.parse(response.getContentText());
    var code = response.getResponseCode();
    if (code !== 503 && code !== 429) break;
    if (attempt < 3) Utilities.sleep(Math.pow(2, attempt) * 3000);
    else throw new Error('AI 서버 과부하 (' + code + '). 잠시 후 다시 시도해 주세요.');
  }

  if (resJson.candidates && resJson.candidates[0] && resJson.candidates[0].content) {
    return resJson.candidates[0].content.parts[0].text;
  }
  throw new Error('AI 응답이 올바르지 않습니다: ' + response.getContentText().substring(0, 200));
}

/**
 * AI 멀티턴 질문 처리 — 대화 히스토리를 Gemini contents 배열로 전달
 * @param {{role:'user'|'model', text:string}[]} messages - 최근 N개 메시지 배열 (클라이언트에서 트리밍)
 * @param {boolean} inclData - true면 첫 번째 user 메시지에 포트폴리오 현황 주입
 */
function runAIQuickQuestionMultiTurn(messages, inclData) {
  const config = getConfig();
  if (!config.geminiApiKey) throw new Error('Gemini API Key가 설정되지 않았습니다.');
  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    throw new Error('메시지 배열이 비어있거나 유효하지 않습니다.');
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.geminiModelId}:generateContent?key=${config.geminiApiKey}`;

  const systemContext = `너는 한국 주식·ETF 포트폴리오 투자 전문가야. 사용자는 한국투자증권 API로 자동 리밸런싱을 운용 중이며, 국내주식·해외주식·채권·금·달러·현금을 분산하는 자산배분 포트폴리오를 관리하고 있어. 질문에 한국어로 간결하고 실용적으로 답변해줘. 투자 조언이 아닌 참고용 분석임을 명심해.`;

  const contents = messages.map(function(msg, idx) {
    let text = msg.text;
    if (inclData && idx === 0 && msg.role === 'user') {
      if (text.indexOf('[현재 포트폴리오 현황]') === -1) {
        text = text + '\n\n[현재 포트폴리오 현황]:\n' + getDashboardStateForAI();
      }
    }
    return { role: msg.role, parts: [{ text: text }] };
  });
  if (!contents[0] || contents[0].role !== 'user') {
    throw new Error('첫 번째 메시지는 반드시 사용자(role: user) 메시지여야 합니다.');
  }

  const payload = {
    system_instruction: { parts: [{ text: systemContext }] },
    contents: contents,
    tools: [{ googleSearch: {} }]
  };
  const options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  let response, resJson;
  for (var attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);
    resJson = JSON.parse(response.getContentText());
    var code = response.getResponseCode();
    if (code !== 503 && code !== 429) break;
    if (attempt < 3) Utilities.sleep(Math.pow(2, attempt) * 3000);
    else throw new Error('AI 서버 과부하 (' + code + '). 잠시 후 다시 시도해 주세요.');
  }

  if (resJson.candidates && resJson.candidates[0] && resJson.candidates[0].content) {
    return resJson.candidates[0].content.parts[0].text;
  }
  throw new Error('AI 응답이 올바르지 않습니다: ' + response.getContentText().substring(0, 200));
}

/**
 * AI 시장 브리핑 및 비중 제안 실행
 * @param {string} [customPrompt] — 프롬프트 확인 팝업에서 수정된 경우 전달, 없으면 저장된 프롬프트 사용
 */
function runAIBriefing(customPrompt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  
  if (!config.geminiApiKey) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Gemini API Key가 필요한 기능입니다.');
    return;
  }
  
  // 기준비율(C열) 미설정 시 현재 운용비율로 자동 초기화
  const wasInitialized = ensureInitialRatios();
  if (wasInitialized) {
    ss.toast('📋 포트폴리오설정 C열(기준비율)이 비어 있어 현재 운용비율로 초기화했습니다.\n원래 의도한 비율이 다르면 C열(기준비율)을 직접 수정하세요.', '⚠️ 기준비율 초기화', 8);
    Utilities.sleep(2000);
  }

  // 당일 추천 이미 존재하면 교체 여부 확인
  if (hasTodayRecommendation()) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert('⚠️ 오늘 이미 AI 추천이 있습니다', '다시 실행하면 오늘 추천 내역을 교체합니다.\n계속하시겠습니까?', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
  }

  ss.toast('Gemini AI가 대시보드와 시장 뉴스를 분석 중입니다...', '🤖 AI 분석 시작', -1);

  try {
    const dashState = getDashboardStateForAI();

    // 기술지표 이력 시트에 기록
    try { recordTAHistory(); } catch (e) { Logger.log('TA 이력 기록 실패: ' + e.toString()); }

    const analysisResult = getGeminiAnalysis(config, dashState, customPrompt);

    // 1. 당일 기존 추천 삭제 후 새 추천 기록
    if (analysisResult.json && analysisResult.json.ratios) {
      try { deleteTodayRecommendations(); recordAIRatios(analysisResult.json, '추천'); } catch (e) { Logger.log('추천 기록 실패: ' + e.message); }
    }

    ss.toast('분석이 완료되었습니다.', '✅', 2);

    // 2. 브리핑 결과 저장 — 나중에 "마지막 AI 브리핑 보기"로 열람 가능
    try {
      const docProps = PropertiesService.getDocumentProperties();
      docProps.setProperty('LAST_AI_BRIEFING_TEXT', analysisResult.text || '');
      docProps.setProperty('LAST_AI_BRIEFING_JSON', JSON.stringify(analysisResult.json || {}));
      docProps.setProperty('LAST_AI_BRIEFING_AT', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'));
      docProps.setProperty('LAST_AI_BRIEFING_SOURCE', '수동');
    } catch(e) { Logger.log('브리핑 저장 실패: ' + e.message); }

    // 3. 분석 결과 표시 (다이얼로그)
    showAIBriefingOutput(analysisResult.text, analysisResult.json);

  } catch (e) {
    Logger.log('AI 분석 오류: ' + e.toString());
    ss.toast('AI 분석 중 오류가 발생했습니다: ' + e.message, '❌ 오류');
  }
}

/**
 * Gemini API 호출 및 분석 결과 반환 (Search Grounding 포함)
 * @param {object} config
 * @param {string} dashState
 * @param {string} [customPrompt] — 미리보기 팝업에서 편집된 프롬프트. 없으면 저장된 프롬프트 사용.
 */
function getGeminiAnalysis(config, dashState, customPrompt) {
  // Search Grounding은 v1beta 필수
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${config.geminiModelId}:generateContent?key=${config.geminiApiKey}`;

  const basePrompt = (customPrompt || getSystemPrompt()).replace('{{TARGET_YIELD}}', config.targetYield);
  const prompt = `
${basePrompt}

[현재 계좌 대시보드 상황]:
${dashState}
`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    tools: [{ googleSearch: {} }]  // 최신 뉴스, 시장 정보, X 게시글 등 자동 검색
  };
  const options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  // 503(과부하) / 429(쿼터 초과) 시 최대 3회 지수 백오프 재시도
  let response, resContent, resJson;
  for (var attempt = 1; attempt <= 3; attempt++) {
    response = UrlFetchApp.fetch(url, options);
    resContent = response.getContentText();
    resJson = JSON.parse(resContent);
    var code = response.getResponseCode();
    if (code !== 503 && code !== 429) break;
    if (attempt < 3) {
      Logger.log('[AI] ' + code + ' 재시도 ' + attempt + '/2');
      Utilities.sleep(Math.pow(2, attempt) * 3000);
    } else {
      throw new Error('AI 서버 과부하 (' + code + '). 잠시 후 다시 시도해 주세요.');
    }
  }

  if (resJson.candidates && resJson.candidates[0].content) {
    const fullText = resJson.candidates[0].content.parts[0].text;
    let proposedJson = null;
    let cleanText = fullText;

    // 1. 정상 포맷: [ALLOCATION_START]...[ALLOCATION_END]
    const tagged = fullText.match(/\[ALLOCATION_START\]\s*([\s\S]*?)\s*\[ALLOCATION_END\]/);
    if (tagged) {
      try {
        let raw = tagged[1].trim().replace(/```json\s*|```\s*/g, '').trim();
        proposedJson = JSON.parse(raw);
        cleanText = fullText.replace(/\*?\*?\[ALLOCATION_START\][\s\S]*?\[ALLOCATION_END\]\*?\*?/, '').trim();
      } catch (e) {
        Logger.log('[AI] tagged JSON 파싱 실패: ' + e.message + ' / raw: ' + tagged[1].substring(0, 200));
      }
    }

    // 2. fallback: ratios 배열을 포함한 JSON 객체 직접 검색
    if (!proposedJson) {
      const objMatch = fullText.match(/\{\s*"ratios"\s*:\s*\[[\s\S]*?\]\s*\}/);
      if (objMatch) {
        try {
          proposedJson = JSON.parse(objMatch[0]);
          cleanText = fullText.replace(objMatch[0], '').trim();
          Logger.log('[AI] fallback JSON 파싱 성공 (ratios 객체 직접 매칭)');
        } catch (e) {
          Logger.log('[AI] fallback JSON 파싱 실패: ' + e.message);
        }
      }
    }

    // 3. 마지막 fallback: ```json ... ``` 코드블록 내부
    if (!proposedJson) {
      const codeBlock = fullText.match(/```json\s*([\s\S]*?)```/);
      if (codeBlock) {
        try {
          const parsed = JSON.parse(codeBlock[1].trim());
          if (parsed && parsed.ratios) {
            proposedJson = parsed;
            cleanText = fullText.replace(codeBlock[0], '').trim();
            Logger.log('[AI] codeblock JSON 파싱 성공');
          }
        } catch (e) {
          Logger.log('[AI] codeblock JSON 파싱 실패: ' + e.message);
        }
      }
    }

    if (!proposedJson) {
      Logger.log('[AI] proposedJson 추출 실패 — 응답 첫 500자: ' + fullText.substring(0, 500));
    }

    return { text: cleanText, json: proposedJson };
  } else {
    throw new Error('AI 응답이 올바르지 않습니다: ' + resContent);
  }
}

/**
 * 분석 결과를 모달창으로 표시 (마크다운 렌더링 + 비중 비교 테이블)
 */
function showAIBriefingOutput(content, proposedJson, titleSuffix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 비중 비교 테이블 데이터 (try/catch로 보호) ---
  const ratioMap = {};
  try {
    // 1단계: 포트폴리오 설정에서 목표 비중 기준 셋업 (D열 운용비율 기준)
    const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
    if (settingSheet && settingSheet.getLastRow() >= 3) {
      const setRows = settingSheet.getRange(3, 1, settingSheet.getLastRow() - 2, 4).getDisplayValues();
      setRows.forEach(function(row) {
        const code = String(row[0]).trim() || 'CASH';
        const name = String(row[1] || code).trim();
        const tgt  = parseFloat(String(row[3]).replace(/[^0-9.\-]/g, '')) || parseFloat(String(row[2]).replace(/[^0-9.\-]/g, '')) || 0;
        ratioMap[code] = { name: name, cur: 0, tgt: tgt };
      });
      Logger.log('[BRIEFING] settings loaded: ' + Object.keys(ratioMap).length + ' codes');
    }

    // 2단계: 대시보드에서 현재 비중 (& 목표 비중 보강)
    const dash = ss.getSheetByName('📊 대시보드');
    if (dash && dash.getLastRow() >= 9) {
      const numRows = dash.getLastRow() - 8;
      const displays = dash.getRange(9, 1, numRows, 8).getDisplayValues();
      Logger.log('[BRIEFING] dashboard sample: ' + JSON.stringify(displays[0] || []));

      displays.forEach(function(row) {
        const code = String(row[0]).trim();
        if (!code) return;
        const curStr = String(row[6] || '').replace(/[^0-9.\-]/g, '');
        const tgtStr = String(row[7] || '').replace(/[^0-9.\-]/g, '');
        const cur = parseFloat(curStr) || 0;
        const tgt = parseFloat(tgtStr) || 0;

        if (ratioMap[code]) {
          ratioMap[code].cur = cur;
          if (tgt > 0) ratioMap[code].tgt = tgt;
        } else {
          ratioMap[code] = { name: String(row[1] || code).trim(), cur: cur, tgt: tgt };
        }
      });

      const totalEval = parseFloat(dash.getRange('B3').getValue()) || 0;
      const cashAmt   = parseFloat(dash.getRange('B4').getValue()) || 0;
      // 현금 현재비중 = 포트폴리오설정의 현금 목표비중만 표시
      // 예수금은 매수 대기 자금이므로 실제 현재비중으로 계산하지 않음
      if (totalEval > 0) {
        if (!ratioMap['CASH']) ratioMap['CASH'] = { name: '현금', cur: 0, tgt: 5 };
        // cur은 표시용으로만 목표비중과 동일하게 — AI 판단 혼선 방지
        ratioMap['CASH'].cur = ratioMap['CASH'].tgt;
      }
    }
  } catch (e) {
    Logger.log('[BRIEFING] ratioMap build error: ' + e.toString());
  }

  Logger.log('[BRIEFING] proposedJson present: ' + (!!proposedJson) + ' / ratios: ' + (proposedJson && proposedJson.ratios ? proposedJson.ratios.length : 'none'));

  // 표 렌더링 — proposedJson 없어도 현재/목표는 항상 표시
  let ratioRows = '';
  let ratioTableHtml = '';
  try {
    const toIntPct = function(v) { return (typeof v === 'number' && v > 0 && v <= 1) ? Math.round(v * 100) : (parseFloat(v) || 0); };
    const hasAI = !!(proposedJson && proposedJson.ratios && proposedJson.ratios.length);

    // AI 제안값 매핑 (있을 때만)
    const aiByCode = {};
    let aiTotal = 0;
    let cashAiPct = null;
    if (hasAI) {
      proposedJson.ratios.forEach(function(p) {
        const code = String(p.code || '').trim();
        const name = String(p.name || '').trim();
        // 현금 감지: 코드가 비어있거나 이름에 현금/cash/예수금/예수 포함
        const isCash = code === 'CASH'
          || (!code && (!name || /현금|cash|예수/i.test(name)))
          || /현금|cash|예수금|예수/i.test(name);
        const ai = toIntPct(p.ratio);
        aiTotal += ai;
        if (isCash) { cashAiPct = ai; return; }
        if (code) aiByCode[code] = ai;
      });
    }

    // 일반 종목 행 — ratioMap 기준 (CASH 제외)
    const stockCodes = Object.keys(ratioMap).filter(c => c !== 'CASH');
    stockCodes.forEach(function(code) {
      const r = ratioMap[code];
      const aiVal = aiByCode[code];
      const cells = [
        '<td class="nm">' + (r.name || code) + '</td>',
        '<td class="num">' + r.tgt.toFixed(1) + '%</td>'
      ];
      if (hasAI) {
        if (aiVal !== undefined) {
          const chg = aiVal - r.tgt;
          const cls = chg > 0.5 ? 'up' : chg < -0.5 ? 'dn' : 'nc';
          const arrow = chg > 0.5 ? '▲' : chg < -0.5 ? '▼' : '=';
          cells.push('<td class="' + cls + '">' + aiVal + '%</td>');
          cells.push('<td class="' + cls + '">' + arrow + (chg > 0 ? '+' : '') + chg.toFixed(0) + '%p</td>');
        } else {
          cells.push('<td class="num">—</td>');
          cells.push('<td class="num">—</td>');
        }
      }
      ratioRows += '<tr>' + cells.join('') + '</tr>';
    });

    // 현금 행
    const cashRow = ratioMap['CASH'] || { name: '현금', cur: 0, tgt: 5 };
    const cashCells = [
      '<td class="nm">💵 현금</td>',
      '<td class="num">' + cashRow.tgt.toFixed(1) + '%</td>'
    ];
    if (hasAI) {
      const cashAi = cashAiPct !== null ? cashAiPct : Math.max(0, 100 - aiTotal);
      if (cashAiPct === null) aiTotal += cashAi;
      const ccg = cashAi - cashRow.tgt;
      const ccls = ccg > 0.5 ? 'up' : ccg < -0.5 ? 'dn' : 'nc';
      const carrow = ccg > 0.5 ? '▲' : ccg < -0.5 ? '▼' : '=';
      cashCells.push('<td class="' + ccls + '">' + cashAi + '%</td>');
      cashCells.push('<td class="' + ccls + '">' + carrow + (ccg > 0 ? '+' : '') + ccg.toFixed(0) + '%p</td>');
    }
    ratioRows += '<tr style="background:#f8f9fa;font-style:italic;">' + cashCells.join('') + '</tr>';

    // 합계 행 (AI 제안 있을 때만)
    if (hasAI) {
      const aiTotalRounded = Math.round(aiTotal * 100) / 100;
      const totalCls = Math.abs(aiTotalRounded - 100) < 0.01 ? 'nc' : 'dn';
      const totalIcon = Math.abs(aiTotalRounded - 100) < 0.01 ? '✅' : '⚠️';
      ratioRows += '<tr style="background:#e8f0fe;font-weight:bold;">' +
        '<td class="nm">' + totalIcon + ' 합계</td>' +
        '<td class="num">100.0%</td>' +
        '<td class="' + totalCls + '">' + aiTotalRounded.toFixed(2) + '%</td>' +
        '<td class="num">—</td>' +
        '</tr>';
    }

    // 헤더 결정
    const headerCells = hasAI
      ? '<th>목표%</th><th>AI 제안%</th><th>목표→제안</th>'
      : '<th>목표%</th>';
    const title = hasAI ? '📊 목표비율 → AI 제안' : '📊 목표비율';

    if (ratioRows) {
      ratioTableHtml = '<div class="section"><div class="st">' + title + '</div>' +
        '<table class="rt"><thead><tr><th style="text-align:left">종목명</th>' +
        headerCells + '</tr></thead><tbody>' + ratioRows + '</tbody></table></div>';
    } else {
      // ratioMap이 비어있어 표를 못 만든 경우 — 디버그 정보 표시
      const debugInfo = 'ratioMap keys: [' + Object.keys(ratioMap).join(', ') + '] / hasAI: ' + hasAI;
      Logger.log('[BRIEFING] empty ratioRows — ' + debugInfo);
      ratioTableHtml = '<div class="section" style="background:#fff3cd;padding:10px;border-radius:4px;">' +
        '<div class="st">⚠️ 비중 비교 표 데이터 없음</div>' +
        '<div style="font-size:12px;color:#5f6368;">' + debugInfo + '</div>' +
        '<div style="font-size:12px;color:#5f6368;margin-top:6px;">대시보드를 새로고침한 뒤 다시 시도해보세요.</div>' +
        '</div>';
    }
  } catch (e) {
    Logger.log('[BRIEFING] table render error: ' + e.toString() + ' / stack: ' + (e.stack || ''));
    ratioTableHtml = '<div class="section" style="background:#fce8e6;padding:10px;border-radius:4px;">' +
      '<div class="st">⚠️ 비중 비교 표 생성 오류</div>' +
      '<div style="font-size:12px;color:#c5221f;">' + e.message + '</div></div>';
  }

  // --- 마크다운 렌더링 ---
  function inlineMd(t) {
    return t.replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>').replace(/\*([^*\n]+)\*/g, '<em>$1</em>');
  }
  const mdLines = content.split('\n');
  const mdParts = [];
  let inList = false;
  mdLines.forEach(function(line) {
    if (/^### /.test(line)) {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<div class="h3">' + inlineMd(line.replace(/^### /, '')) + '</div>');
    } else if (/^## /.test(line)) {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<div class="h2">' + inlineMd(line.replace(/^## /, '')) + '</div>');
    } else if (/^- /.test(line)) {
      if (!inList) { mdParts.push('<ul>'); inList = true; }
      mdParts.push('<li>' + inlineMd(line.replace(/^- /, '')) + '</li>');
    } else if (line.trim() === '') {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<br>');
    } else {
      if (inList) { mdParts.push('</ul>'); inList = false; }
      mdParts.push('<p>' + inlineMd(line) + '</p>');
    }
  });
  if (inList) mdParts.push('</ul>');
  const mdHtml = mdParts.join('');

  const applyBtn = proposedJson
    ? '<input id="applyBtn" type="button" class="apply-btn" value="✅ 지금 즉시 비중 적용하기">'
    : '';

  // textarea에 JSON을 raw로 넣기 — HtmlService가 script/attribute를 sanitize해도 textContent는 안전
  const payloadTextarea = proposedJson
    ? '<textarea id="payloadData" style="display:none">' + JSON.stringify(proposedJson) + '<\/textarea>'
    : '';

  const html = '<!DOCTYPE html><html><head><style>'
    + '*{box-sizing:border-box;}'
    + 'body{font-family:\'Malgun Gothic\',sans-serif;line-height:1.6;color:#3c4043;margin:0;padding:0;font-size:13px;}'
    + '.wrap{padding:16px 20px;}'
    + '.title{font-size:17px;font-weight:bold;color:#4285F4;padding-bottom:10px;border-bottom:2px solid #e8eaed;margin-bottom:12px;}'
    + '.section{margin-bottom:12px;}'
    + '.st{font-weight:bold;font-size:12px;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;}'
    + '.rt{width:100%;border-collapse:collapse;font-size:12px;}'
    + '.rt th{background:#e8f0fe;color:#1967d2;padding:5px 8px;text-align:center;font-weight:bold;}'
    + '.rt td{padding:4px 8px;border-bottom:1px solid #f1f3f4;}'
    + '.rt td.nm{font-weight:500;}'
    + '.rt td.num{text-align:right;color:#5f6368;}'
    + '.rt td.up{text-align:right;color:#137333;font-weight:bold;}'
    + '.rt td.dn{text-align:right;color:#c5221f;font-weight:bold;}'
    + '.rt td.nc{text-align:right;color:#9aa0a6;}'
    + '.md{max-height:360px;overflow-y:auto;padding-right:6px;}'
    + '.md .h3{font-size:13px;font-weight:bold;color:#1a73e8;margin:10px 0 4px;border-left:3px solid #4285F4;padding-left:7px;}'
    + '.md .h2{font-size:14px;font-weight:bold;color:#3c4043;margin:12px 0 4px;}'
    + '.md p{margin:2px 0;}'
    + '.md ul{margin:3px 0 3px 14px;padding:0;}'
    + '.md li{margin-bottom:4px;}'
    + '.md strong{color:#1a73e8;}'
    + '.md em{color:#5f6368;}'
    + '.apply-btn{display:block;width:100%;padding:11px;background:#4285F4;color:white;border:none;border-radius:6px;cursor:pointer;font-weight:bold;font-size:14px;margin-top:12px;}'
    + '.apply-btn:disabled{opacity:.5;cursor:not-allowed;}'
    + '.footer{margin-top:10px;text-align:right;border-top:1px solid #f1f3f4;padding-top:10px;}'
    + '.close-btn{padding:7px 14px;cursor:pointer;border:1px solid #dadce0;background:white;border-radius:4px;font-size:13px;}'
    + '</style></head><body>'
    + payloadTextarea
    + '<div class="wrap">'
    + '<div class="title">🤖 Gemini AI 시장 브리핑</div>'
    + ratioTableHtml
    + '<div class="section"><div class="st">📝 분석 내용</div><div class="md">' + mdHtml + '</div></div>'
    + applyBtn
    + '<div class="footer"><button id="closeBtn" class="close-btn">닫기</button></div>'
    + '</div>'
    + '<script>'
    + '(function(){'
    + 'function setup(){'
    + 'var btn=document.getElementById(\'applyBtn\');'
    + 'if(btn){'
    + 'btn.addEventListener(\'click\',function(){'
    + 'if(!confirm(\'AI가 제안한 비중으로 포트폴리오 설정을 변경하시겠습니까?\'))return;'
    + 'var el=document.getElementById(\'payloadData\');'
    + 'var j=el?el.value:\'\';'
    + 'if(!j){alert(\'비중 데이터를 찾을 수 없습니다. 다시 분석을 실행해주세요.\');return;}'
    + 'btn.value=\'적용 중...\';btn.disabled=true;'
    + 'google.script.run'
    + '.withSuccessHandler(function(){alert(\'✅ 운용비율이 업데이트되었습니다.\\n대시보드가 자동으로 새로고침됩니다.\');google.script.host.close();})'
    + '.withFailureHandler(function(e){alert(\'오류: \'+e);btn.value=\'✅ 지금 즉시 비중 적용하기\';btn.disabled=false;})'
    + '.applyAIProposedRatiosManual(j);'
    + '});'
    + '}'
    + 'var close=document.getElementById(\'closeBtn\');'
    + 'if(close)close.addEventListener(\'click\',function(){google.script.host.close();});'
    + '}'
    + 'if(document.readyState===\'loading\'){'
    + 'document.addEventListener(\'DOMContentLoaded\',setup);'
    + '}else{setup();}'
    + '})();'
    + '<\/script></body></html>';

  const ui = HtmlService.createHtmlOutput(html)
    .setWidth(700).setHeight(900)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  const dialogTitle = 'AI 시장 분석 및 비중 제안' + (titleSuffix || '');
  SpreadsheetApp.getUi().showModelessDialog(ui, dialogTitle);
}

/**
 * AI가 제안한 비중을 실제 시트에 적용하고 이력을 업데이트함
 */
// ─────────────────────────────────────────────────────────────
// Layer 1: AI 비중 변경 락 (14일 고정 주기)
// ─────────────────────────────────────────────────────────────
const RATIO_LOCK_DAYS = 14;

/**
 * AI 비중 변경 락 상태 조회
 * @returns {{ locked, daysRemaining, daysSinceUpdate, lockDays, lastUpdate }}
 */
function getRatioLockStatus() {
  const props = PropertiesService.getScriptProperties();
  const lastStr = props.getProperty('LAST_AI_RATIO_UPDATE');
  const lockDays = RATIO_LOCK_DAYS;

  if (!lastStr) {
    return { locked: false, daysRemaining: 0, daysSinceUpdate: Infinity, lockDays, lastUpdate: null };
  }

  const last = new Date(lastStr);
  const now = new Date();
  const daysSinceUpdate = Math.floor((now - last) / 86400000);
  const daysRemaining = Math.max(0, lockDays - daysSinceUpdate);
  return { locked: daysSinceUpdate < lockDays, daysRemaining, daysSinceUpdate, lockDays, lastUpdate: last };
}

/**
 * 락 해제 (수동) — 사용자가 즉시 비중 변경 원할 때
 */
function unlockRatioChange() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_AI_RATIO_UPDATE');
  SpreadsheetApp.getUi().alert('🔓 AI 비중 변경 락이 해제되었습니다.');
}

/**
 * AI가 제안한 비중을 포트폴리오설정 시트에 적용
 * @param {string|object} jsonStr — AI 제안 JSON
 * @param {object} [options={}] — { force: 락 무시 여부 }
 */
function applyAIProposedRatios(jsonStr, options) {
  options = options || {};
  const data = typeof jsonStr === 'string' ? JSON.parse(jsonStr) : jsonStr;
  if (!data || !data.ratios) throw new Error('유효하지 않은 데이터입니다.');

  // Layer 1 락 체크 (수동 적용은 force=true로 무시)
  if (!options.force) {
    const lock = getRatioLockStatus();
    if (lock.locked) {
      Logger.log('[L1 락] 비중 변경 스킵 — ' + lock.daysRemaining + '일 남음 (락 ' + lock.lockDays + '일)');
      return { applied: false, locked: true, daysRemaining: lock.daysRemaining };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) throw new Error('포트폴리오설정 시트를 찾을 수 없습니다.');

  const lastRow = sheet.getLastRow();
  // col: 0=종목코드(A), 1=종목명(B), 2=기준비율(C/고정), 3=조정값(D) ← AI는 D열만 수정
  // D열 = AI제안(절대값) - C열(기준비율) → 조정값(+/-)
  const range = sheet.getRange(3, 1, lastRow - 2, 4);
  const values = range.getValues();

  const updatedValues = values.map(row => {
    const code = String(row[0]).trim();
    const proposal = data.ratios.find(r =>
      String(r.code).trim() === code || (code === '' && (r.code === 'CASH' || r.name === '현금'))
    );
    if (proposal) {
      const proposed = (typeof proposal.ratio === 'number' && proposal.ratio > 0 && proposal.ratio <= 1)
        ? Math.round(proposal.ratio * 100) : proposal.ratio;
      const base = parseFloat(row[2]) || 0;
      row[3] = proposed - base; // 조정값 = 제안 - 기준
    }
    return row;
  });

  range.setValues(updatedValues);

  // 락 타이머 리셋
  PropertiesService.getScriptProperties().setProperty('LAST_AI_RATIO_UPDATE', new Date().toISOString());

  // 적용 전 이전 날짜 '추천' 행 삭제 (이력에 추천+적용한비중 2종만 유지)
  deleteOldRecommendations();

  // 오늘 날짜 '추천' 행 상태를 '적용한 비중'으로 변경 (새 행 추가 없음)
  updateLatestRecommendationStatus('적용한 비중');

  // updateDashboard()는 호출하지 않음 — 대화상자 응답 속도를 위해 분리.
  // 비중 적용 후 대시보드는 사용자가 직접 새로고침하거나 자동트리거에서 반영됨.
  return { applied: true, locked: false };
}

/**
 * 수동 적용용 래퍼 — 다이얼로그 "지금 적용" 버튼에서 호출 (락 무시)
 */
function applyAIProposedRatiosManual(jsonStr) {
  const result = applyAIProposedRatios(jsonStr, { force: true });
  if (result && result.applied) updateDashboard();
  return result;
}

/**
 * 가장 최근 '추천' 배치의 상태를 일괄 변경 (새 행 추가 없이 상태 컬럼만 업데이트)
 */
function updateLatestRecommendationStatus(newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('📝 비중변경이력');
  if (!historySheet) return;
  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) return;

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data = historySheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const rowsToUpdate = [];

  data.forEach((row, i) => {
    if (!row[0]) return;
    const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (rowDate === todayStr && String(row[7]).trim() === '추천') {
      rowsToUpdate.push(i + 2);
    }
  });
  rowsToUpdate.forEach(rowIdx => historySheet.getRange(rowIdx, 8).setValue(newStatus));
}

/**
 * AI 비중 제안 내역을 '📝 비중변경이력' 시트에 누적 기록
 * 컬럼: 시간 | 종목명 | 유형 | 변경전(%) | 변경후(%) | 변경이유 | 활용모델 | 상태
 */
function recordAIRatios(data, status = '추천') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupAIHistorySheet();
  const historySheet = ss.getSheetByName('📝 비중변경이력');

  const config = getConfig();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  const toIntPct = v => (typeof v === 'number' && v > 0 && v <= 1) ? Math.round(v * 100) : (parseFloat(v) || 0);

  const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
  const settingData = settingSheet.getRange(3, 1, Math.max(1, settingSheet.getLastRow() - 2), 5).getValues();
  const currentMap = {};
  settingData.forEach(row => {
    const code = String(row[0]).trim() || 'CASH';
    currentMap[code] = { name: row[1], ratio: parseFloat(row[3]) || 0, type: row[4] || '' };
  });

  // 오늘 날짜의 기존 '추천' 행 인덱스 맵 (종목명 → rowIdx)
  const existingRowMap = {};
  const lastRow = historySheet.getLastRow();
  if (lastRow >= 2) {
    const existing = historySheet.getRange(2, 1, lastRow - 1, 8).getValues();
    existing.forEach((row, i) => {
      if (!row[0]) return;
      const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
      if (rowDate === todayStr && String(row[7]).trim() === '추천') {
        existingRowMap[String(row[1]).trim()] = i + 2; // 1-indexed sheet row
      }
    });
  }

  data.ratios.forEach(proposal => {
    const code = String(proposal.code || '').trim();
    const cur  = currentMap[code] || {};
    const assetName  = cur.name || (code === 'CASH' ? '현금' : proposal.name || code);
    const assetType  = cur.type || '';
    const beforeRatio = cur.ratio != null ? cur.ratio : '';
    const afterRatio  = toIntPct(proposal.ratio);
    const rowData = [now, assetName, assetType, beforeRatio, afterRatio, proposal.rationale || data.summary || 'AI 비중 제안', config.geminiModelId, status];

    const existingIdx = existingRowMap[assetName.trim()];
    if (existingIdx) {
      // 오늘 같은 종목 추천이 있으면 덮어쓰기
      historySheet.getRange(existingIdx, 1, 1, 8).setValues([rowData]);
    } else {
      // 없으면 새 행 추가
      historySheet.appendRow(rowData);
    }
  });

  trimExtraColumns(historySheet, 8);
}

function hasTodayRecommendation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet || sheet.getLastRow() < 2) return false;
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return data.some(row => {
    if (!row[0]) return false;
    return Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd') === todayStr && String(row[7]).trim() === '추천';
  });
}

function deleteTodayRecommendations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet || sheet.getLastRow() < 2) return;
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const toDelete = [];
  data.forEach((row, idx) => {
    if (!row[0]) return;
    if (Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd') === todayStr && String(row[7]).trim() === '추천') {
      toDelete.push(idx + 2);
    }
  });
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
}

/**
 * 오늘 이전 날짜의 '추천' 행 삭제 — 적용 시 호출하여 이력을 추천+적용한비중 2종으로 정리
 */
function deleteOldRecommendations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet || sheet.getLastRow() < 2) return;
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const toDelete = [];
  data.forEach((row, idx) => {
    if (!row[0]) return;
    const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (rowDate !== todayStr && String(row[7]).trim() === '추천') {
      toDelete.push(idx + 2);
    }
  });
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
}

/**
 * 비중변경이력 시트에서 직전 N일간 '추천' 평균 비중을 계산
 *  - 마지막 적용(LAST_AI_RATIO_UPDATE) 이후의 추천만 대상
 *  - 같은 종목명 여러 일자의 변경후(%)를 평균
 *  - 합계가 100이 되도록 가장 큰 항목에서 보정
 *
 * @param {number} days — 평균 윈도우 (일)
 * @returns {{ratios, summary, runCount}|null}
 */
function getAveragedAIRatios(days) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📝 비중변경이력');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const props = PropertiesService.getScriptProperties();
  const lastUpdateStr = props.getProperty('LAST_AI_RATIO_UPDATE');
  let cutoffMs = Date.now() - days * 86400000;
  if (lastUpdateStr) {
    cutoffMs = Math.max(cutoffMs, new Date(lastUpdateStr).getTime());
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const grouped = {};
  const seenTimes = {};
  let runCount = 0;

  data.forEach(row => {
    const time = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (!time || isNaN(time.getTime()) || time.getTime() < cutoffMs) return;
    const name = String(row[1] || '').trim();
    const ratio = parseFloat(row[4]) || 0;
    const status = String(row[7] || '').trim();
    if (status !== '추천' || !name) return;

    if (!grouped[name]) grouped[name] = [];
    grouped[name].push(ratio);

    const tk = time.getTime();
    if (!seenTimes[tk]) { seenTimes[tk] = true; runCount++; }
  });

  const names = Object.keys(grouped);
  if (names.length === 0 || runCount < 1) return null;

  // 평균 계산
  const averages = {};
  names.forEach(n => {
    const arr = grouped[n];
    averages[n] = arr.reduce((s, v) => s + v, 0) / arr.length;
  });

  // 포트폴리오 설정에서 종목명 → 코드 매핑
  const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!settingSheet) return null;
  const settings = settingSheet.getRange(3, 1, Math.max(1, settingSheet.getLastRow() - 2), 2).getValues();
  const ratios = [];
  settings.forEach(row => {
    const code = String(row[0]).trim() || 'CASH';
    const name = String(row[1] || '').trim();
    if (averages[name] !== undefined) {
      ratios.push({ code: code, name: name, ratio: Math.round(averages[name]) });
    }
  });

  // 합계 100% 보정 — 가장 큰 항목에서 ±조정
  if (ratios.length > 0) {
    const total = ratios.reduce((s, r) => s + r.ratio, 0);
    if (total !== 100) {
      let maxIdx = 0;
      for (let i = 1; i < ratios.length; i++) {
        if (ratios[i].ratio > ratios[maxIdx].ratio) maxIdx = i;
      }
      ratios[maxIdx].ratio += (100 - total);
    }
  }

  return {
    ratios: ratios,
    summary: `직전 ${days}일 ${runCount}회 AI 추천의 평균치 (단발 변동 평탄화)`,
    runCount: runCount
  };
}

/**
 * 가장 최근의 '추천' 상태인 비중 제안을 찾아 실제 포트폴리오에 반영함
 */
function applyLatestRecommendation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName('📝 비중변경이력');
  if (!historySheet) {
    setupAIHistorySheet();
    historySheet = ss.getSheetByName('📝 비중변경이력');
  }
  
  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('반영할 추천 내역이 없습니다.');
    return;
  }
  
  const data = historySheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // 가장 최근의 연속된 '추천' 행들을 찾음 (같은 시간에 생성된 것들)
  let latestTime = null;
  const recommendedRatios = [];
  const rowsToUpdate = [];

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const time   = row[0].toString();
    const status = row[7]; // col 8: 상태

    if (status === '추천') {
      if (latestTime === null) latestTime = time;
      if (time === latestTime) {
        recommendedRatios.push({ name: row[1], ratio: parseFloat(row[4]) }); // col 2: 종목명, col 5: 변경후%
        rowsToUpdate.push(i + 2); // 1-indexed, starting from row 2
      } else {
        break;
      }
    }
  }
  
  if (recommendedRatios.length === 0) {
    SpreadsheetApp.getUi().alert('적용할 수 있는 최근 추천 내역이 없습니다.');
    return;
  }
  
  // 포트폴리오 설정 업데이트
  const settingSheet = ss.getSheetByName('📋 포트폴리오설정');
  const lastSettingRow = settingSheet.getLastRow();
  const range = settingSheet.getRange(3, 1, lastSettingRow - 2, 4);
  const values = range.getValues();

  const updatedValues = values.map(row => {
    const code = String(row[0]).trim();
    const name = row[1];
    const match = recommendedRatios.find(r => r.name === code || r.name === name);
    if (match) row[3] = match.ratio; // D열(운용비율) 업데이트
    return row;
  });

  range.setValues(updatedValues);

  // 상태 업데이트 (추천 -> 적용됨), col 8
  rowsToUpdate.forEach(rowIdx => {
    historySheet.getRange(rowIdx, 8).setValue('적용됨');
  });

  // Layer 1 락 타이머 리셋 (수동 적용 — force 동등)
  PropertiesService.getScriptProperties().setProperty('LAST_AI_RATIO_UPDATE', new Date().toISOString());

  ss.toast('추천 비중이 성공적으로 반영되었습니다.', '✅ 반영 완료');
  updateDashboard();
}

/**
 * 외부 AI용 프롬프트 생성 및 클립보드 복사 다이얼로그 표시
 */
function copyAIPromptToClipboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  
  // 1. 대시보드 데이터 수집
  const dash = ss.getSheetByName('📊 대시보드');
  let dashData = "대시보드 데이터를 찾을 수 없습니다.";
  if (dash) {
    const lastRow = dash.getLastRow();
    if (lastRow >= 9) {
      const headers = ["종목코드", "종목명", "유형", "보유수량", "현재가", "평가액", "현재비율", "목표비율", "예상비중", "차이"];
      const rows = dash.getRange(9, 1, lastRow - 8, 10).getDisplayValues();
      dashData = headers.join("\t") + "\n" + rows.map(r => r.join("\t")).join("\n");
    }
  }

  // 2. 포트폴리오 설정 데이터 수집
  const setting = ss.getSheetByName('📋 포트폴리오설정');
  let settingData = "설정 데이터를 찾을 수 없습니다.";
  if (setting) {
    const lastRow = setting.getLastRow();
    if (lastRow >= 3) {
      const headers = ["종목코드", "종목명", "목표비율(%)", "유형"];
      const rows = setting.getRange(3, 1, lastRow - 2, 4).getDisplayValues();
      settingData = headers.join("\t") + "\n" + rows.map(r => r.join("\t")).join("\n");
    }
  }

  // 3. 프롬프트 템플릿 조립
  const basePrompt = getSystemPrompt().replace('{{TARGET_YIELD}}', config.targetYield);
  const prompt = `
${basePrompt}

### 현재 대시보드 현황
${dashData}

### 현재 목표 포트폴리오 설정
${settingData}
`;

  // 4. HTML 다이얼로그 표시
  const html = `
    <html>
      <head>
        <style>
          body { font-family: 'Malgun Gothic', sans-serif; padding: 20px; color: #3c4043; }
          textarea { width: 100%; height: 300px; font-family: monospace; font-size: 12px; padding: 10px; border: 1px solid #dadce0; border-radius: 4px; }
          .btn-copy { background: #4285F4; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold; width: 100%; margin-top: 15px; }
          .btn-copy:active { background: #1967d2; }
          .hint { font-size: 12px; color: #5f6368; margin-bottom: 10px; }
        </style>
      </head>
      <body>
        <h3>📋 AI 프롬프트 복사</h3>
        <p class="hint">아래 내용을 복사하여 ChatGPT, Claude 등 다른 AI에 붙여넣으세요.</p>
        <textarea id="promptText">${prompt.trim()}</textarea>
        <button class="btn-copy" onclick="copyText()">클립보드에 복사하기</button>
        <script>
          function copyText() {
            const textarea = document.getElementById('promptText');
            textarea.select();
            document.execCommand('copy');
            const btn = event.target;
            btn.innerText = '✅ 복사 완료!';
            btn.style.background = '#34A853';
            setTimeout(() => {
              google.script.host.close();
            }, 1000);
          }
        </script>
      </body>
    </html>
  `;
  
  const ui = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(500).setTitle('AI 프롬프트 생성');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}
