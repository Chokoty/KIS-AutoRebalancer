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
          ul {
            margin: 0;
            padding-left: 20px;
            color: #3c4043;
            font-size: 14px;
          }
          li {
            margin-bottom: 8px;
          }
           .fsd-box {
            background-color: #fff8e1;
            padding: 12px;
            border-radius: 6px;
            border: 1px solid #ffca28;
            margin-top: 15px;
            font-size: 13.5px;
          }
          .fsd-title {
            color: #f57c00;
            font-weight: bold;
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
          <li><strong>🔑 API 키 입력:</strong> <code>⚙️ 설정</code> 시트에 한국투자증권 접속 정보(APP KEY, SECRET, 계좌번호)와 Gemini AI API 키를 입력하세요.</li>
        </ul>

        <div class="section-title">2. 포트폴리오(비중) 관리</div>
        <ul>
          <li><strong>📋 포트폴리오설정:</strong> 목표로 하는 종목코드와 운용 비중 총합이 100%가 되도록 기입하세요. (예: 삼성전자 50%, 현금 50%)</li>
          <li><strong>🤖 AI 제안 받기:</strong> 메뉴에서 <code>AI 시장 분석 및 비중 제안</code>을 누르면 현재 시장 상황에 맞춰 최적의 포트폴리오 비율을 제안해 줍니다.</li>
        </ul>

        <div class="section-title">3. 매매 및 자동화</div>
        <ul>
          <li><strong>⚡ 리밸런싱 실행:</strong> 대시보드에서 <code>리밸런싱 실행</code>을 누르면 설정된 목표 비중에 맞춰 자동으로 초과분은 매도하고 부족분은 매수합니다.</li>
          <li><strong>🛣️ 고속도로 차선 유지 (정기 리밸런싱):</strong> 매주 월요일 오전 10시에 자동으로 리밸런싱을 실행합니다.</li>
          <li><strong>🤖 AI 자율 포트폴리오 관리:</strong> 매일 AI가 알아서 시장을 분석하고 비중을 조절하며 자동 매매를 수행합니다.</li>
        </ul>

        <div class="section-title">4. 💰 수익 실현</div>
        <ul>
          <li>메뉴의 <code>수익 실현 창 열기</code>를 통해 원하는 금액을 입력하면 보유 비중에 비례하여 자동으로 종목별 매도 수량을 계산해 줍니다.</li>
          <li>매도 후 인출 금액은 <strong>2주간 리밸런싱 예수금에서 보호</strong>됩니다 (실수령 전 재매수 방지).</li>
        </ul>

        <div class="section-title">5. 📐 리밸런싱 계산 방식</div>
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

        <div class="section-title">6. 📊 대시보드 항목 설명</div>
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

        <div class="fsd-box">
          <div class="fsd-title">🏎️ FSD 드라이빙 모드 (AI 성향 설정)란?</div>
          AI 브리핑 시 전략을 결정하는 위험 감수 성향입니다.
          <ul style="margin-top: 5px;">
            <li><strong>🍃 Chill:</strong> 원금 보존 최우선. 보수적, 안전 자산(금/달러) 위주</li>
            <li><strong>⚖️ Standard:</strong> (기본값) 수익과 위험 방어의 적절한 균형</li>
            <li><strong>🏃 Hurry:</strong> 시장 기회를 엿보며 적극적으로 주식 비중 확대</li>
            <li><strong>🔥 Assertive:</strong> 단기 손실을 감내하더라도 초과 수익 강하게 추구</li>
            <li><strong>💀 Mad Max:</strong> 가장 공격적인 투자로 리스크 무시, 수익 극대화</li>
          </ul>
        </div>

        <div style="text-align: center; margin-top: 30px;">
          <button onclick="google.script.host.close()" style="background-color: #1a73e8; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold;">확인</button>
        </div>
      </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(750).setTitle('기본 사용법 가이드');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}
