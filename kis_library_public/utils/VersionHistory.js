/**
 * 버전 업데이트 내역 모달 표시
 */
function showVersionHistory() {
  // 사용자가 업데이트 내역을 확인하면 시트 버전을 최신으로 동기화(알림 배너 제거)
  if (typeof syncTemplateVersion === 'function') {
    syncTemplateVersion(true); // true = 조용한 동기화(토스트 생략)
  }

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
          .version-block {
            margin-bottom: 25px;
          }
          .version-title {
            font-weight: bold;
            font-size: 16px;
            color: #202124;
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 8px;
          }
          .version-date {
            font-size: 12px;
            color: #5f6368;
            font-weight: normal;
          }
          ul {
            margin: 0;
            padding-left: 20px;
            color: #3c4043;
            font-size: 14px;
          }
          li {
            margin-bottom: 6px;
          }
          .badge {
            background: #e8f0fe;
            color: #1a73e8;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: bold;
          }
          .badge-new {
            background: #e6f4ea;
            color: #137333;
          }
          .badge-fix {
            background: #fce8e6;
            color: #c5221f;
          }
        </style>
      </head>
      <body>
        <h3>📜 시스템 업데이트 내역</h3>
        
        <div class="version-block">
          <div class="version-title">
            v1.1.0 <span class="badge badge-new">New</span>
            <span class="version-date">2026.05.13</span>
          </div>
          <ul>
            <li><strong>📊 기준비율 / 운용비율 분리:</strong> 포트폴리오설정 시트가 기준비율(C열, 고정)과 운용비율(D열, AI/수동 조정)로 나뉘어, 처음 설정한 목표 비중과 실제 운용 비중을 구분할 수 있습니다.</li>
            <li><strong>🤖 AI 비중 제안 개선:</strong> 시장 분석 요약과 "현재 비중 ➡️ 조정 비중" 포맷으로 AI 제안이 간결하고 직관적으로 표시됩니다. 결과는 비중변경이력 시트에 자동 기록됩니다.</li>
            <li><strong>💬 AI 빠른 질문:</strong> 포트폴리오와 무관한 시장·경제 관련 자유 질문을 AI에게 바로 물어볼 수 있습니다.</li>
            <li><strong>💡 추천 비중 반영:</strong> 비중변경이력에 기록된 최신 AI 추천을 한 번에 운용비율에 반영할 수 있습니다.</li>
            <li><strong>📋 포트폴리오 종목 추가/관리:</strong> 다이얼로그에서 종목코드로 이름을 검색하고 목표 비중을 직접 추가·수정할 수 있습니다.</li>
            <li><strong>📝 시트 최적화:</strong> 모든 시트에서 불필요한 빈 열이 자동으로 제거됩니다.</li>
            <li><strong>📜 업데이트 내역 보기:</strong> 현재 보고 계신 버전 릴리스 노트를 확인할 수 있는 메뉴가 신설되었습니다.</li>
          </ul>
        </div>

        <div class="version-block">
          <div class="version-title">
            v1.0.0 <span class="badge">Release</span>
            <span class="version-date">2026.03.05 이력 요약</span>
          </div>
          <ul>
            <li><strong>💰 계좌 & 대시보드 연동:</strong> 한국투자증권 API 연동 및 주요 평가손익 실시간 새로고침 지원.</li>
            <li><strong>🤖 Gemini AI 분석 연동:</strong> 보유 포트폴리오 비중을 AI가 제안해주고, 자동으로 비중변경이력에 추천/적용 상태로 기록.</li>
            <li><strong>🛣️ 차선 유지 (격주 리밸런싱):</strong> 매주 월요일 자동 리밸런싱 트리거, 단 마지막 실행 후 13일 미만이면 스킵(격주 효과).</li>
            <li><strong>💰 수익 실현:</strong> 원하는 금액을 입력하면 보유 비중에 비례해 종목별 매도 수량을 계산·실행.</li>
            <li><strong>🛡️ 보안:</strong> API Key·Secret·계좌번호 등 민감 정보는 사용자 속성(User Properties)에 암호화 보관.</li>
          </ul>
        </div>

        <div style="text-align: center; margin-top: 30px;">
          <button onclick="google.script.host.close()" style="background-color: #f1f3f4; border: 1px solid #dadce0; color: #3c4043; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-weight: bold;">닫기</button>
        </div>
      </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(600).setHeight(550).setTitle('버전 업데이트 내역');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}
