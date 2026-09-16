/**
 * KIS AutoTrader 기본 사용법 가이드 모달
 */
function showUsageGuide() {
  const html = `
    <html>
      <head>
        <style>
          body {
            font-family: 'Malgun Gothic', sans-serif;
            padding: 20px;
            color: #3c4043;
            line-height: 1.6;
          }
          h3 {
            color: #1a73e8;
            border-bottom: 2px solid #e8f0fe;
            padding-bottom: 10px;
            margin-top: 0;
          }
          .section-title {
            font-weight: bold;
            font-size: 16px;
            color: #202124;
            margin-top: 20px;
            margin-bottom: 8px;
            background-color: #f8f9fa;
            padding: 6px 10px;
            border-radius: 4px;
            border-left: 4px solid #4285F4;
          }
          .section-title.emergency {
            border-left-color: #ea4335;
            background-color: #fce8e6;
            color: #c5221f;
          }
          ul {
            margin: 0;
            padding-left: 20px;
            color: #3c4043;
            font-size: 14px;
          }
          li {
            margin-bottom: 8px;
          }
          strong {
            color: #202124;
          }
        </style>
      </head>
      <body>
        <h3>📖 KIS AutoTrader 기본 사용법</h3>

        <div class="section-title">1. 시스템 초기 세팅</div>
        <ul>
          <li><strong>⚙️ 초기 설정:</strong> 가장 먼저 메뉴에서 초기 설정을 눌러 필수 시트들을 생성하세요.</li>
          <li><strong>🛡️ API 키 보안 설정:</strong> 메뉴 → 설정 및 관리에서 한국투자증권 접속 정보(APP KEY, SECRET, 계좌번호)와 Gemini API 키를 입력하세요. 시트 셀이 아니라 본인 계정에만 암호화 저장됩니다.</li>
          <li><strong>🔧 기본 설정:</strong> 계좌 종류(일반/ISA/모의)와 리밸런싱 임계치·수익실현 임계치·연 목표 수익률도 이 팝업에서 정합니다. <code>⚙️ 설정</code> 시트 칸을 직접 고치는 건 막혀 있습니다 — 값이 틀렸을 땐 시트가 아니라 이 팝업을 여세요.</li>
        </ul>

        <div class="section-title">2. 포트폴리오(비중) 관리</div>
        <ul>
          <li><strong>📋 포트폴리오 종목 추가/관리:</strong> 종목코드를 검색해서 추가하고, 각 종목의 기준비율(%)을 정합니다. <code>📋 포트폴리오설정</code> 시트도 직접 편집이 막혀 있습니다 — 이 팝업에서만 바꿉니다.</li>
          <li><strong>비중 자동 계산:</strong> 종목을 추가하거나 삭제하거나 기준비율을 바꾸면, 나머지 종목들의 비율이 자동으로 비례해서 조정되어 합계가 항상 100%에 가깝게 맞춰집니다. 직접 일일이 계산해서 맞출 필요가 없습니다.</li>
          <li><strong>🤖 AI 비중 제안:</strong> 메뉴에서 AI 분석 → AI 비중 제안을 누르면 현재 시장 상황에 맞춘 조정안을 보여줍니다. 사람이 확인하고 승인해야 실제로 반영됩니다.</li>
        </ul>

        <div class="section-title">3. 매매 및 자동화</div>
        <ul>
          <li><strong>⚡ 리밸런싱 실행:</strong> 목표 비중과 현재 비중의 차이가 임계치 이상인 종목만 매도 후 매수합니다.</li>
          <li><strong>🛣️ 차선유지 (정기 리밸런싱):</strong> 지정한 요일·시간에 자동으로 리밸런싱을 실행합니다. 마지막 실행 후 며칠 이내면 자동으로 건너뜁니다.</li>
        </ul>

        <div class="section-title emergency">4. 🚨 긴급 대응 — 계획에 없던 일이 생겼을 때</div>
        <ul>
          <li><strong>🔴 종목 즉시 전량매도:</strong> 목표 비중과 무관하게, 보유 종목 하나를 골라 지금 당장 시장가로 전부 정리합니다. (포트폴리오에서 종목을 삭제하거나 비중을 0%로 바꾸는 것과는 다릅니다 — 그건 "다음 리밸런싱에서 이 종목을 목표로 안 본다"는 뜻일 뿐, 실제로 팔아주지는 않습니다.)</li>
          <li><strong>🚨 자동매매 긴급 정지:</strong> 차선유지(정기 리밸런싱) 예약을 즉시 끕니다. 이미 꺼져 있으면 그렇다고 알려줄 뿐, 실수로 다시 켜지지 않습니다.</li>
        </ul>

        <div class="section-title">5. 💰 수익 실현 (현금화)</div>
        <ul>
          <li><strong>⚖️ 비중 유지 비례매도 (기본):</strong> 필요한 금액을 입력하면 보유 비중을 그대로 유지한 채 여러 종목에서 조금씩 나눠 팝니다.</li>
          <li><strong>🎯 종목 선택 매도:</strong> 특정 종목 한두 개만 체크해서, 그 종목들에서만 필요한 금액만큼 팝니다. 자잘하게 여러 종목을 건드리고 싶지 않을 때 씁니다.</li>
          <li>매도 후 인출 금액은 <strong>2주간 리밸런싱 매수 여력에서 제외</strong>됩니다 — 실제로 인출하기 전에 자동매매가 그 돈으로 재매수하는 걸 막아줍니다.</li>
        </ul>

        <div class="section-title">6. 📐 리밸런싱 계산 방식</div>
        <ul>
          <li><strong>임계치 기준:</strong> 현재 비중이 목표 비중과 <code>리밸런싱 임계치(%)</code> 이상 차이날 때만 주문이 발생합니다. (기본 2%p)</li>
          <li><strong>매도 우선:</strong> 초과 종목을 먼저 매도하여 현금을 확보한 뒤, 부족 종목을 매수합니다.</li>
          <li><strong>수수료 포함 계산:</strong> 모든 매매 계산에 KIS 온라인 기준 수수료가 자동 반영됩니다.
            <ul style="margin-top: 4px;">
              <li>매수: <strong>0.015%</strong> (온라인 거래 수수료)</li>
              <li>매도: <strong>0.215%</strong> (수수료 0.015% + 증권거래세·농특세 0.2%)</li>
            </ul>
          </li>
          <li><strong>수익실현 임계치:</strong> 특정 종목의 수익률이 설정값 이상일 때 AI 브리핑 시 해당 종목의 비중 축소를 우선 고려합니다.</li>
        </ul>

        <div class="section-title">7. 📊 대시보드 항목 설명</div>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          <thead>
            <tr style="background:#e8f0fe;">
              <th style="padding:6px 10px;text-align:left;border:1px solid #dadce0;width:38%">항목</th>
              <th style="padding:6px 10px;text-align:left;border:1px solid #dadce0;">설명</th>
            </tr>
          </thead>
          <tbody>
            <tr><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💰 총 평가액</td><td style="padding:5px 10px;border:1px solid #dadce0;">보유 종목 평가액 합계 + 예수금 (현재 순자산)</td></tr>
            <tr style="background:#f8f9fa"><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💵 예수금</td><td style="padding:5px 10px;border:1px solid #dadce0;">매수에 즉시 사용 가능한 계좌 내 현금</td></tr>
            <tr><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💵 현금 비율</td><td style="padding:5px 10px;border:1px solid #dadce0;">총 평가액 중 예수금이 차지하는 비율</td></tr>
            <tr style="background:#f8f9fa"><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💵 총 투자금</td><td style="padding:5px 10px;border:1px solid #dadce0;">지금까지 실제로 입금한 원금 누계 (⚙️설정에 직접 입력)</td></tr>
            <tr><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💰 평가손익</td><td style="padding:5px 10px;border:1px solid #dadce0;">총 평가액 − 총 투자금 (미실현 손익)</td></tr>
            <tr style="background:#f8f9fa"><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">📈 수익률</td><td style="padding:5px 10px;border:1px solid #dadce0;">평가손익 ÷ 총 투자금 × 100</td></tr>
            <tr><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">🔄 리밸런싱 대상</td><td style="padding:5px 10px;border:1px solid #dadce0;">현재비중이 목표비중과 임계치 이상 벌어진 종목 수</td></tr>
            <tr style="background:#f8f9fa"><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">🛒 총 매수 필요액</td><td style="padding:5px 10px;border:1px solid #dadce0;">이번 리밸런싱에서 매수해야 할 금액 합계</td></tr>
            <tr><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">📉 총 매도 예정액</td><td style="padding:5px 10px;border:1px solid #dadce0;">이번 리밸런싱에서 매도할 금액 합계</td></tr>
            <tr style="background:#f8f9fa"><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💸 예상 제비용</td><td style="padding:5px 10px;border:1px solid #dadce0;">매수·매도 수수료 및 증권거래세 합계 예상액</td></tr>
            <tr><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💳 순 자산 변동</td><td style="padding:5px 10px;border:1px solid #dadce0;">총 매도 예정액 − 총 매수 필요액. 리밸런싱 실행 시 현금이 얼마나 증감하는지 나타냄. 음수면 현금 감소, 양수면 현금 증가.</td></tr>
            <tr style="background:#f8f9fa"><td style="padding:5px 10px;border:1px solid #dadce0;font-weight:bold">💡 월 인출 추천</td><td style="padding:5px 10px;border:1px solid #dadce0;">목표 수익률(연) 기준으로 매달 인출 가능한 권장 금액</td></tr>
          </tbody>
        </table>

        <div style="text-align: center; margin-top: 30px;">
          <button onclick="google.script.host.close()" style="background-color: #1a73e8; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold;">확인</button>
        </div>
      </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(750).setTitle('기본 사용법 가이드');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}
