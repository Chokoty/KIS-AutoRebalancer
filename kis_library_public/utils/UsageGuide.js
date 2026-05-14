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

        <div class="section-title">1. 초기 설정</div>
        <ul>
          <li><strong>⚙️ 초기 설정 실행:</strong> 메뉴 <code>⚙️ 설정 및 관리 &gt; 초기 설정</code>을 눌러 필수 시트(대시보드, 계좌현황, 포트폴리오설정, 거래내역 등)를 자동 생성합니다.</li>
          <li><strong>🛡️ API 키 입력:</strong> <code>API 키 보안 설정</code> 창에서 한국투자증권 APP KEY, APP SECRET, 계좌번호와 Gemini API 키를 입력하세요. 모두 암호화되어 안전하게 보관됩니다.</li>
          <li><strong>⚙️ 설정 시트:</strong> 계좌 종류(일반/ISA/모의), 리밸런싱 임계치(%), 수익실현 임계치(%) 등을 확인하고 필요 시 수정하세요.</li>
        </ul>

        <div class="section-title">2. 포트폴리오 비중 설정</div>
        <ul>
          <li><strong>📋 포트폴리오설정 시트 구조:</strong>
            <ul style="margin-top: 4px;">
              <li><strong>C열 기준비율(%):</strong> 처음 설정한 목표 비중. 고정값으로, 변경 시 직접 수정합니다.</li>
              <li><strong>D열 운용비율(%):</strong> 실제 리밸런싱에 사용되는 비중. AI 제안 또는 수동으로 조정합니다.</li>
            </ul>
          </li>
          <li><strong>📋 종목 추가/관리:</strong> 메뉴 <code>포트폴리오 종목 추가/관리</code>에서 종목코드로 이름을 검색하고 목표 비중을 입력하세요. 운용비율의 합계가 100%가 되어야 합니다.</li>
        </ul>

        <div class="section-title">3. 리밸런싱 실행</div>
        <ul>
          <li><strong>⚡ 수동 리밸런싱:</strong> 메뉴 <code>리밸런싱 실행</code>을 누르면 운용비율(D열) 기준으로 초과 종목은 매도, 부족 종목은 매수합니다.</li>
          <li><strong>🛣️ 차선 유지 (격주 자동 리밸런싱):</strong> <code>차선유지 켜기/끄기</code>를 ON으로 설정하면 매주 월요일 오전 10시에 자동 실행됩니다. 단, 마지막 실행 후 13일 이내이면 자동 스킵(격주 효과)됩니다.</li>
          <li><strong>임계치 기준:</strong> 현재 비중과 운용비율의 차이가 설정한 임계치(%p) 이상일 때만 주문이 발생합니다. (기본 2%p)</li>
          <li><strong>수수료 자동 반영:</strong> ETF 기준 매수/매도 수수료 0.015%가 계산에 포함됩니다.</li>
        </ul>

        <div class="section-title">4. AI 비중 제안 및 반영</div>
        <ul>
          <li><strong>📊 AI 비중 제안:</strong> 메뉴 <code>🤖 AI 분석 &gt; AI 비중 제안</code>을 누르면 현재 시장 상황을 분석하여 "현재 비중 ➡️ 조정 비중" 형태로 제안합니다. 결과는 비중변경이력 시트에 자동 저장됩니다.</li>
          <li><strong>💡 추천 비중 반영:</strong> <code>추천 비중 반영</code> 메뉴를 누르면 비중변경이력의 최신 AI 추천을 운용비율(D열)에 한 번에 적용합니다. 적용 후 14일간 재변경이 제한됩니다.</li>
          <li><strong>💬 AI 빠른 질문:</strong> 포트폴리오와 무관한 시장·경제 궁금증을 AI에게 바로 질문할 수 있습니다.</li>
          <li><strong>⚙️ AI 프롬프트 상세 설정:</strong> AI가 분석 시 참고하는 시스템 프롬프트를 직접 수정할 수 있습니다.</li>
        </ul>

        <div class="section-title">5. 수익 실현</div>
        <ul>
          <li><code>💰 수익 실현 &gt; 수익 실현 창 열기</code>에서 원하는 금액을 입력하면 보유 비중에 비례해 종목별 매도 수량을 계산·실행합니다.</li>
          <li>매도 후 인출 예정 금액은 <strong>2주간 예수금 보호</strong>로 처리되어 리밸런싱 시 재매수되지 않습니다.</li>
        </ul>

        <div style="text-align: center; margin-top: 30px;">
          <button onclick="google.script.host.close()" style="background-color: #1a73e8; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold;">확인</button>
        </div>
      </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(750).setTitle('기본 사용법 가이드');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}
