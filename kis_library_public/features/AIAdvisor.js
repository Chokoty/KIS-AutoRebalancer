/**
 * AI 시장 분석 및 비중 제안 (Gemini API + Google Search Grounding)
 */

/**
 * 대시보드 현재 상태를 AI에게 전달할 텍스트로 요약
 */
function getDashboardStateForAI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName('📊 대시보드');
  if (!dash) return '대시보드 데이터를 찾을 수 없습니다.';

  const totalEval  = dash.getRange('B3').getValue();
  const cash       = dash.getRange('B4').getValue();
  const lastUpdate = dash.getRange('B6').getDisplayValue();

  const lastRow = dash.getLastRow();
  if (lastRow < 9) return `현재 총자산: ${totalEval.toLocaleString()}원, 업데이트: ${lastUpdate}`;

  const data = dash.getRange(9, 1, lastRow - 8, 10).getValues();

  let report = `[포트폴리오 현황 (기준: ${lastUpdate})]\n`;
  report += `총 평가액: ${totalEval.toLocaleString()}원 | 예수금: ${cash.toLocaleString()}원\n`;
  report += '\n[종목별 현황]\n';

  data.forEach(row => {
    const code = String(row[0]).trim();
    const name = row[1];
    if (!code) return;
    const evalAmount   = row[5];
    const currentRatio = row[6];
    const targetRatio  = row[7];
    const diff         = row[9];
    report += `▶ ${name} (${code}): 현재비중 ${currentRatio} → 목표비중 ${targetRatio} (차이 ${diff}) | 평가액 ${typeof evalAmount === 'number' ? evalAmount.toLocaleString() : evalAmount}원\n`;
  });

  return report;
}

/**
 * AI 분석용 시스템 프롬프트
 */
function getAISystemPrompt() {
  return `너는 자산 배분 전략가이자 글로벌 시장 분석가야.
아래 포트폴리오 데이터와 Google Search로 최신 뉴스·시장 동향을 종합해서 비중 조정을 제안해줘.

[분석 방법]:
- 현재 포트폴리오 비중과 목표 비중의 차이 확인
- Google Search로 각 자산군(채권, 금, 달러, 국내주식, 해외주식 등)의 최신 뉴스·시장 상황 파악
- 시장 상황을 고려해 목표 비중 조정 제안

[포트폴리오 원칙]:
- 단일 종목 최대 비중: 40% 이하
- 합산 반드시 100% 유지
- 점진적 조정 선호 (급격한 변경 지양)

[출력 양식]:
### 🌍 시장 요약
- 요약 1줄
- 요약 1줄

### 🎯 비중 조정 제안
- **[종목명]**: 현재 X% ➡️ **조정 Y%** | 근거 1줄

[ALLOCATION_START] {"ratios": [{"code": "종목코드", "name": "종목명", "ratio": 숫자, "rationale": "간략한 이유"}], "summary": "전체 요약"} [ALLOCATION_END]

"이것은 투자 조언이 아닙니다. 스스로 판단하세요."`;
}

/**
 * AI 시장 분석 및 비중 제안 실행 (메뉴 진입점)
 */
function runAIBriefing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getUserProperties();
  const ssId = ss.getId();
  const geminiApiKey  = props.getProperty(ssId + '_GEMINI_API_KEY') || '';
  const geminiModelId = props.getProperty(ssId + '_GEMINI_MODEL_ID') || 'gemini-2.0-flash';

  if (!geminiApiKey) {
    SpreadsheetApp.getUi().alert('Gemini API Key가 필요합니다.\n메뉴 > ⚙️ 설정 및 관리 > 🛡️ API 키 보안 설정에서 입력해 주세요.');
    return;
  }

  ss.toast('Gemini AI가 포트폴리오와 시장 뉴스를 분석 중입니다...', '🤖 AI 분석 시작', -1);

  try {
    const dashState = getDashboardStateForAI();
    const result = getGeminiAnalysis(geminiApiKey, geminiModelId, dashState);
    showAIBriefingOutput(result.text, result.json);
  } catch (e) {
    Logger.log('AI 분석 오류: ' + e.toString());
    ss.toast('AI 분석 중 오류가 발생했습니다: ' + e.message, '❌ 오류');
  }
}

/**
 * Gemini API 호출 (Google Search Grounding 포함)
 */
function getGeminiAnalysis(geminiApiKey, geminiModelId, dashState) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModelId}:generateContent?key=${geminiApiKey}`;
  const prompt = getAISystemPrompt() + '\n\n[현재 계좌 대시보드 상황]:\n' + dashState;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    tools: [{ googleSearch: {} }]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const resJson  = JSON.parse(response.getContentText());

  if (!resJson.candidates || !resJson.candidates[0] || !resJson.candidates[0].content) {
    throw new Error('AI 응답이 올바르지 않습니다: ' + response.getContentText().substring(0, 300));
  }

  const fullText  = resJson.candidates[0].content.parts[0].text;
  const jsonMatch = fullText.match(/\[ALLOCATION_START\]\s*([\s\S]*?)\s*\[ALLOCATION_END\]/s);
  let proposedJson = null;
  let cleanText    = fullText;

  if (jsonMatch) {
    try {
      const rawJson = jsonMatch[1].trim().replace(/```json\s*|```\s*/g, '').trim();
      proposedJson = JSON.parse(rawJson);
      cleanText    = fullText.replace(/\[ALLOCATION_START\][\s\S]*?\[ALLOCATION_END\]/s, '').trim();
    } catch (e) {
      Logger.log('AI JSON 파싱 실패: ' + e.message);
    }
  }

  return { text: cleanText, json: proposedJson };
}

/**
 * 분석 결과 모달 표시
 */
function showAIBriefingOutput(content, proposedJson) {
  const htmlContent = content
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/### (.*?)(\n|$)/g, '<h3>$1</h3>')
    .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>')
    .replace(/\n/g, '<br>');

  const jsonStr = proposedJson ? encodeURIComponent(JSON.stringify(proposedJson)) : '';

  const actionHtml = proposedJson ? `
    <div style="background:#fff8e1;padding:15px;border-radius:8px;margin-top:20px;border:1px solid #ffca28;">
      <h4 style="margin-top:0;color:#f57c00;">📈 AI 비중 추천</h4>
      <p style="font-size:12px;color:#5f6368;">AI가 제안한 비중을 포트폴리오 설정에 반영합니다.</p>
      <div style="text-align:center;">
        <input type="button" value="추천 비중 적용하기" onclick="applyRatios('${jsonStr}')"
               style="background-color:#4285F4;color:white;border:none;padding:10px 20px;border-radius:4px;cursor:pointer;font-weight:bold;">
      </div>
    </div>` : '';

  const html = `
    <html>
      <head>
        <style>
          body { font-family:'Malgun Gothic',sans-serif; line-height:1.6; color:#3c4043; margin:0; }
          .container { padding:20px; }
          h2 { color:#4285F4; border-bottom:2px solid #f1f3f4; padding-bottom:10px; }
          h3 { color:#34a853; margin-top:16px; }
          .content { font-size:14px; max-height:450px; overflow-y:auto; padding-right:10px; }
          .footer { margin-top:20px; text-align:right; border-top:1px solid #f1f3f4; padding-top:15px; }
        </style>
        <script>
          function applyRatios(jsonEncoded) {
            if (!confirm('AI가 제안한 비중으로 포트폴리오 설정을 변경하시겠습니까?')) return;
            const btn = event.target;
            btn.value = '적용 중...';
            btn.disabled = true;
            google.script.run
              .withSuccessHandler(function() {
                alert('비중 설정이 성공적으로 업데이트되었습니다.');
                google.script.host.close();
              })
              .withFailureHandler(function(err) {
                alert('오류 발생: ' + err);
                btn.value = '추천 비중 적용하기';
                btn.disabled = false;
              })
              .applyAIProposedRatios(decodeURIComponent(jsonEncoded));
          }
        </script>
      </head>
      <body>
        <div class="container">
          <h2>🤖 Gemini AI 시장 분석 및 비중 제안</h2>
          <div class="content">${htmlContent}</div>
          ${actionHtml}
          <div class="footer">
            <input type="button" value="닫기" onclick="google.script.host.close()"
                   style="padding:8px 16px;cursor:pointer;border:1px solid #dadce0;background:white;border-radius:4px;">
          </div>
        </div>
      </body>
    </html>`;

  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(650).setHeight(750),
    ' '
  );
}

/**
 * AI가 제안한 비중을 포트폴리오설정 시트에 적용 (다이얼로그에서 호출)
 */
function applyAIProposedRatios(jsonStr) {
  const data = typeof jsonStr === 'string' ? JSON.parse(jsonStr) : jsonStr;
  if (!data || !data.ratios) throw new Error('유효하지 않은 데이터입니다.');

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) throw new Error('포트폴리오설정 시트를 찾을 수 없습니다.');

  const lastRow = sheet.getLastRow();
  const range   = sheet.getRange(3, 1, lastRow - 2, 3);
  const values  = range.getValues();

  const updated = values.map(row => {
    const code = String(row[0]).trim();
    const name = String(row[1]).trim();
    const match = data.ratios.find(r =>
      String(r.code || '').trim() === code ||
      String(r.name || '').trim() === name ||
      (code === '' && (r.code === 'CASH' || r.name === '현금'))
    );
    if (match) row[2] = match.ratio;
    return row;
  });

  range.setValues(updated);
  updateDashboard();
  return true;
}
