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

        <div class="section-title">4. 💰 생활비 인출 (수익 실현)</div>
        <ul>
          <li>메뉴의 <code>생활비 인출</code> &gt; <code>인출 창 열기</code>를 통해 매월 추천되는 수익 인출 금액을 확인하고, 터치 한 번으로 비율에 맞춰 자동 매도하여 현금화할 수 있습니다.</li>
        </ul>

        <div style="text-align: center; margin-top: 30px;">
          <button onclick="google.script.host.close()" style="background-color: #1a73e8; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold;">확인</button>
        </div>
      </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(600).setTitle('기본 사용법 가이드');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}
