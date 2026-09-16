// 메뉴 추가 (라이브러리용)
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 시트 열릴 때 설정 안내 토스트
  // TEMPLATE_VERSION 미설정 = 초기설정 미실행 시트 (새 복사본 등)
  const docProps = PropertiesService.getDocumentProperties();
  const templateVersion = docProps.getProperty('TEMPLATE_VERSION');
  if (!templateVersion) {
    ss.toast('KIS AutoTrader 메뉴 → 설정 및 관리 → 초기 설정을 실행하세요.', '⚠️ 초기 설정 필요', 15);
  } else {
    const portSheet = ss.getSheetByName('📋 포트폴리오설정');
    if (portSheet && String(portSheet.getRange('C2').getValue()).trim() !== '기준비율(%)') {
      ss.toast('KIS AutoTrader 메뉴 → 설정 및 관리 → 포트폴리오 설정 컬럼 업데이트를 실행하세요.', '⚠️ 포트폴리오 설정 업데이트 필요', 15);
    }
  }
  ui.createMenu('📊 KIS AutoTrader')
    .addItem('💰 계좌 현황 새로고침', 'updateAccountSheet')
    .addItem('🔄 대시보드 새로고침', 'updateDashboard')
    .addItem('⚡ 리밸런싱 실행', 'executeRebalanceFromDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('🤖 AI 분석')
      .addItem('📊 AI 비중 제안 (프롬프트 확인 후 실행)', 'openAIBriefingPreview')
      .addItem('💬 AI 빠른 질문', 'openAIQuickQuestion'))
    .addItem('💡 추천 비중 반영', 'applyLatestRecommendation')
    .addItem('📋 포트폴리오 종목 추가/관리', 'openPortfolioManagerDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 수익 실현')
      .addItem('📋 수익 실현 창 열기', 'openWithdrawDialog')
      .addSeparator()
      .addItem('🔓 보호 예수금 해제', 'releaseProtectedCash'))
    .addSeparator()
    .addItem('🛣️ 차선유지 설정 (정기 리밸런싱)', 'showHighwaySettings')
    .addItem('📊 시스템 상태 보기', 'showSystemStatus')
    .addSeparator()
    .addItem('📜 업데이트 내역 보기', 'showVersionHistory')
    .addItem('📖 기본 사용법 안내', 'showUsageGuide')
    .addSubMenu(ui.createMenu('⚙️ 설정 및 관리')
      .addItem('⚙️ 초기 설정',                  'setupSheets')
      .addItem('🛡️ API 키 보안 설정',           'openSecureConfigDialog')
      .addItem('🔧 기본 설정 (계좌종류·임계치)', 'openBasicSettingsDialog')
      .addItem('📋 포트폴리오 설정 컬럼 업데이트', 'addInitialRatiosColumn')
      .addItem('⚙️ AI 프롬프트 상세 설정',      'openAIPromptSettings')
      .addItem('🔑 토큰 초기화 (오류 발생 시)', 'forceRefreshToken')
      .addSeparator()
      .addItem('🔓 비중 락 강제 해제',          'unlockRatioChange'))
    .addToUi();
}


/**
 * 시스템 상태 한눈에 보기 (HTML 모달 — 타임아웃 없음)
 */
function showSystemStatus() {
  // 활성 트리거 조회
  const triggerMap = {};
  ScriptApp.getProjectTriggers().forEach(t => {
    triggerMap[t.getHandlerFunction()] = true;
  });

  // Layer 1 락 상태
  let lockHtml = '';
  try {
    const lock = getRatioLockStatus();
    if (lock.lastUpdate) {
      const lastStr = Utilities.formatDate(lock.lastUpdate, 'Asia/Seoul', 'MM/dd HH:mm');
      const lockStatus = lock.locked
        ? `🔒 잠김 (${lock.daysRemaining}일 남음)`
        : '🔓 해제됨 (변경 가능)';
      lockHtml = `<div class="row"><b>Layer 1 비중 락</b><br>${lockStatus}<br><span class="dim">마지막 변경: ${lastStr} | 락 주기: ${lock.lockDays}일</span></div>`;
    } else {
      lockHtml = '<div class="row"><b>Layer 1 비중 락</b><br>변경 이력 없음</div>';
    }
  } catch(e) {}

  // 보호 예수금
  let protHtml = '<div class="row"><b>📦 보호 예수금</b><br>없음</div>';
  try {
    const prot = getProtectedCash();
    if (prot.amount > 0) {
      protHtml = `<div class="row"><b>📦 보호 예수금</b><br>${prot.amount.toLocaleString()}원 (${prot.daysLeft}일 남음)</div>`;
    }
  } catch(e) {}

  // 다음 적용 예상 비중 (누적 평균 기반 미리보기)
  let previewHtml = '';
  try {
    const lockInfo = getRatioLockStatus();
    const averaged = getAveragedAIRatios(lockInfo.lockDays);

    // 현재 목표 비중 (포트폴리오 설정)
    const portSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📋 포트폴리오설정');
    const currentTargets = {};
    if (portSheet && portSheet.getLastRow() >= 3) {
      portSheet.getRange(3, 1, portSheet.getLastRow() - 2, 4).getValues().forEach(r => {
        const code = String(r[0]).trim() || 'CASH';
        currentTargets[code] = { name: String(r[1] || code).trim(), tgt: parseFloat(r[3]) || 0 };
      });
    }

    if (!averaged || averaged.runCount === 0) {
      previewHtml = '<div class="row"><b>🔮 AI 추천 누적 현황</b><br><span class="dim">아직 누적된 AI 추천 없음 — AI 비중 제안을 실행하면 누적됩니다.</span></div>';
    } else {
      const whenStr = lockInfo.locked
        ? `${lockInfo.daysRemaining}일 후 변경 잠금 해제`
        : '비중 변경 가능';
      let rows = '';
      averaged.ratios.forEach(r => {
        const cur = currentTargets[r.code] || { name: r.name, tgt: 0 };
        const diff = r.ratio - cur.tgt;
        const arrow = diff > 0.5 ? '▲' : diff < -0.5 ? '▼' : '=';
        const cls = diff > 0.5 ? 'up' : diff < -0.5 ? 'dn' : 'eq';
        const sign = diff > 0 ? '+' : '';
        rows += `<tr><td>${cur.name}</td><td class="num">${cur.tgt.toFixed(0)}%</td>` +
                `<td class="num"><b>${r.ratio}%</b></td>` +
                `<td class="num ${cls}">${arrow}${sign}${diff.toFixed(0)}%p</td></tr>`;
      });
      previewHtml = `<div class="row" style="background:#fff8e1;border-left:3px solid #f9ab00;">
        <b>🔮 AI 추천 누적 현황</b>
        <div class="dim" style="margin:4px 0 8px 0;">${averaged.summary} | ${whenStr}</div>
        <table class="preview-tbl">
          <thead><tr><th style="text-align:left">종목</th><th>현재 목표</th><th>예상</th><th>변화</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
    }
  } catch (e) {
    previewHtml = '<div class="row"><b>🔮 다음 적용 예상 비중</b><br><span class="dim">로딩 실패: ' + e.message + '</span></div>';
  }

  const row = (icon, label, active) =>
    `<div class="auto-row ${active ? 'on' : 'off'}">${active ? '✅' : '⛔'} ${label}</div>`;

  const laneOn = !!triggerMap['scheduledBiWeeklyRebalance'];
  const laneBadge = laneOn
    ? '<span class="badge on">🛣️ 차선유지 ON</span>'
    : '<span class="badge off">🛣️ 차선유지 OFF</span>';

  const html = `<!DOCTYPE html><html><head><style>
    body { font-family: 'Malgun Gothic', sans-serif; padding: 18px; color: #3c4043; }
    h2 { margin: 0 0 14px 0; color: #1a73e8; font-size: 17px; }
    .badge { padding: 4px 10px; border-radius: 12px; font-size: 12px; }
    .badge.on  { background: #34a853; color: white; }
    .badge.off { background: #5f6368; color: white; }
    .section-title { font-weight: bold; margin: 12px 0 6px 0; color: #202124; }
    .auto-row { padding: 6px 10px; border-radius: 4px; margin-bottom: 4px; font-size: 13px; }
    .auto-row.on  { background: #e6f4ea; color: #137333; }
    .auto-row.off { background: #f1f3f4; color: #5f6368; }
    .row { background: #f8f9fa; padding: 10px 12px; border-radius: 6px; margin-bottom: 8px; font-size: 13px; }
    .dim { color: #5f6368; font-size: 12px; }
    .btn { background: #1a73e8; color: white; border: none; padding: 8px 18px; border-radius: 6px; cursor: pointer; font-weight: bold; }
    .footer { text-align: right; margin-top: 14px; }
    .preview-tbl { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 4px; }
    .preview-tbl th { background: #fce8b2; padding: 5px 7px; font-weight: bold; }
    .preview-tbl td { padding: 4px 7px; border-bottom: 1px solid #fef0c7; }
    .preview-tbl .num { text-align: right; }
    .preview-tbl .up { color: #34a853; font-weight: bold; }
    .preview-tbl .dn { color: #ea4335; font-weight: bold; }
    .preview-tbl .eq { color: #5f6368; }
  </style></head><body>
    <h2>📊 시스템 상태</h2>
    <div style="display:flex;gap:8px;margin-bottom:14px;">${laneBadge}</div>
    <div class="section-title">📋 활성 자동화</div>
    ${row('🛣️', '정기 리밸런싱', !!triggerMap['scheduledBiWeeklyRebalance'])}
    <div class="section-title">🔒 락 / 보호 상태</div>
    ${lockHtml}
    ${protHtml}
    <div class="section-title">🔮 AI 추천 현황</div>
    ${previewHtml}
    <div class="footer"><button class="btn" onclick="google.script.host.close()">확인</button></div>
  </body></html>`;

  const ui = HtmlService.createHtmlOutput(html).setWidth(480).setHeight(720).setTitle('📊 시스템 상태');
  SpreadsheetApp.getUi().showModalDialog(ui, ' ');
}

/**
 * 고속도로 차선 유지 (자동 새로고침) 토글
 */
function toggleHighwayLaneKeeping() {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('HIGHWAY_LANE_KEEPING') === 'TRUE';
  const next = !current;
  
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    // 차선유지는 이제 정기 리밸런싱(scheduledBiWeeklyRebalance)을 의미함
    if (t.getHandlerFunction() === 'scheduledBiWeeklyRebalance') ScriptApp.deleteTrigger(t);
  });
  
  if (next) {
    // 주간 정기 리밸런싱 트리거 생성 (월요일 오전 10시)
    ScriptApp.newTrigger('scheduledBiWeeklyRebalance')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(10)
      .create();
    props.setProperty('HIGHWAY_LANE_KEEPING', 'TRUE');
    SpreadsheetApp.getUi().alert('🛣️ 차선유지 [ON]\n매주 월요일 10시 발동, 단 마지막 실행 후 13일 미만이면 자동 스킵 (격주 효과).');
  } else {
    props.setProperty('HIGHWAY_LANE_KEEPING', 'FALSE');
    SpreadsheetApp.getUi().alert('🛣️ 차선 유지(정기 리밸런싱)가 [OFF] 되었습니다.');
  }
  updateDashboard();
}

/**
 * 차선유지 설정 HTML 다이얼로그 — 요일/시간 버튼 선택
 */
function showHighwaySettings() {
  const props = PropertiesService.getScriptProperties();
  const isOn = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'scheduledBiWeeklyRebalance');
  const savedDay  = props.getProperty('HIGHWAY_WEEKDAY') || 'MONDAY';
  const savedHour = props.getProperty('HIGHWAY_HOUR')    || '10';

  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; margin: 0; background: #fff; }
  h3 { margin: 0 0 4px; font-size: 16px; color: #1a73e8; }
  .status { font-size: 12px; color: #5f6368; margin-bottom: 16px; }
  .status b { color: ${isOn ? '#137333' : '#c5221f'}; }
  label { display: block; font-size: 13px; font-weight: 600; color: #3c4043; margin: 14px 0 6px; }
  .btn-group { display: flex; flex-wrap: wrap; gap: 6px; }
  .btn-group button {
    padding: 7px 14px; border: 1.5px solid #dadce0; border-radius: 20px;
    background: #fff; font-size: 13px; cursor: pointer; color: #3c4043;
    transition: all .15s;
  }
  .btn-group button.selected {
    background: #1a73e8; color: #fff; border-color: #1a73e8; font-weight: 600;
  }
  .actions { margin-top: 20px; display: flex; justify-content: space-between; align-items: center; }
  .off-btn { font-size: 12px; color: #c5221f; background: none; border: none; cursor: pointer; padding: 0; }
  .save-btn {
    background: #1a73e8; color: #fff; border: none; border-radius: 4px;
    padding: 9px 24px; font-size: 14px; cursor: pointer; font-weight: 600;
  }
  .save-btn:hover { background: #1558b0; }
</style>
</head>
<body>
  <h3>🛣️ 차선유지 설정</h3>
  <div class="status">현재 상태: <b>${isOn ? 'ON' : 'OFF'}</b></div>

  <label>실행 요일</label>
  <div class="btn-group" id="dayGroup">
    <button onclick="sel('day',this,'MONDAY')"   class="${savedDay==='MONDAY'   ?'selected':''}">월요일</button>
    <button onclick="sel('day',this,'TUESDAY')"  class="${savedDay==='TUESDAY'  ?'selected':''}">화요일</button>
    <button onclick="sel('day',this,'WEDNESDAY')"class="${savedDay==='WEDNESDAY'?'selected':''}">수요일</button>
    <button onclick="sel('day',this,'THURSDAY')" class="${savedDay==='THURSDAY' ?'selected':''}">목요일</button>
    <button onclick="sel('day',this,'FRIDAY')"   class="${savedDay==='FRIDAY'   ?'selected':''}">금요일</button>
  </div>

  <label>실행 시간</label>
  <div class="btn-group" id="hourGroup">
    <button onclick="sel('hour',this,'8')"  class="${savedHour==='8' ?'selected':''}">오전 8시</button>
    <button onclick="sel('hour',this,'9')"  class="${savedHour==='9' ?'selected':''}">오전 9시</button>
    <button onclick="sel('hour',this,'10')" class="${savedHour==='10'?'selected':''}">오전 10시</button>
    <button onclick="sel('hour',this,'11')" class="${savedHour==='11'?'selected':''}">오전 11시</button>
    <button onclick="sel('hour',this,'13')" class="${savedHour==='13'?'selected':''}">오후 1시</button>
  </div>

  <div class="actions">
    <button class="off-btn" onclick="turnOff()">차선유지 끄기</button>
    <button class="save-btn" onclick="save()">저장</button>
  </div>

<script>
  var day  = '${savedDay}';
  var hour = '${savedHour}';

  function sel(type, el, val) {
    var group = el.parentNode;
    group.querySelectorAll('button').forEach(function(b){ b.classList.remove('selected'); });
    el.classList.add('selected');
    if (type === 'day')  day  = val;
    if (type === 'hour') hour = val;
  }

  function save() {
    google.script.run.withSuccessHandler(function(){ google.script.host.close(); })
      .applyHighwaySettings(day, hour, false);
  }

  function turnOff() {
    google.script.run.withSuccessHandler(function(){ google.script.host.close(); })
      .applyHighwaySettings('', '', true);
  }
<\/script>
</body>
</html>`;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(360).setHeight(320).setTitle('차선유지 설정');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🛣️ 차선유지 설정');
}

/**
 * showHighwaySettings에서 호출 — 트리거 실제 등록/해제
 */
function applyHighwaySettings(dayKey, hourStr, turnOff) {
  const props = PropertiesService.getScriptProperties();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledBiWeeklyRebalance') ScriptApp.deleteTrigger(t);
  });

  if (turnOff) {
    props.setProperty('HIGHWAY_LANE_KEEPING', 'FALSE');
    updateDashboard();
    return;
  }

  const DAY_MAP = {
    MONDAY:    { day: ScriptApp.WeekDay.MONDAY,    label: '월요일' },
    TUESDAY:   { day: ScriptApp.WeekDay.TUESDAY,   label: '화요일' },
    WEDNESDAY: { day: ScriptApp.WeekDay.WEDNESDAY, label: '수요일' },
    THURSDAY:  { day: ScriptApp.WeekDay.THURSDAY,  label: '목요일' },
    FRIDAY:    { day: ScriptApp.WeekDay.FRIDAY,    label: '금요일' },
  };
  const selected = DAY_MAP[dayKey];
  if (!selected) return;

  const hour = parseInt(hourStr) || 10;
  ScriptApp.newTrigger('scheduledBiWeeklyRebalance')
    .timeBased()
    .onWeekDay(selected.day)
    .atHour(hour)
    .create();

  props.setProperty('HIGHWAY_LANE_KEEPING', 'TRUE');
  props.setProperty('HIGHWAY_WEEKDAY', dayKey);
  props.setProperty('HIGHWAY_HOUR', hourStr);

  updateDashboard();
}

/**
 * 🔧 기본 설정 (계좌종류·임계치) — ⚙️ 설정 시트 B7:B10은 직접 편집이 막혀 있으므로
 * (onEdit 가드, container/code.gs) 값 변경은 이 팝업을 거친다.
 */
function openBasicSettingsDialog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️ 설정');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚙️ 설정 시트를 찾을 수 없습니다. 초기 설정을 먼저 실행하세요.');
    return;
  }
  const accountType    = String(sheet.getRange('B7').getValue() || '일반').trim();
  const rebalanceTol   = sheet.getRange('B8').getValue()  || 2.0;
  const profitThreshold= sheet.getRange('B9').getValue()  || 40.0;
  const targetYield    = sheet.getRange('B10').getValue() || 10.0;

  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; margin: 0; background: #fff; }
  h3 { margin: 0 0 4px; font-size: 16px; color: #1a73e8; }
  .dim { font-size: 12px; color: #5f6368; margin-bottom: 16px; }
  label { display: block; font-size: 13px; font-weight: 600; color: #3c4043; margin: 14px 0 6px; }
  .btn-group { display: flex; gap: 6px; }
  .btn-group button {
    padding: 7px 14px; border: 1.5px solid #dadce0; border-radius: 20px;
    background: #fff; font-size: 13px; cursor: pointer; color: #3c4043;
  }
  .btn-group button.selected { background: #1a73e8; color: #fff; border-color: #1a73e8; font-weight: 600; }
  input[type=number] { width: 100%; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 6px; font-size: 14px; box-sizing: border-box; }
  .row-hint { font-size: 11px; color: #9aa0a6; margin-top: 4px; }
  .actions { margin-top: 22px; text-align: right; }
  .save-btn {
    background: #1a73e8; color: #fff; border: none; border-radius: 4px;
    padding: 9px 24px; font-size: 14px; cursor: pointer; font-weight: 600;
  }
  .save-btn:hover { background: #1558b0; }
</style>
</head>
<body>
  <h3>🔧 기본 설정</h3>
  <div class="dim">계좌 종류·임계치는 여기서만 바꿀 수 있습니다. (시트 직접 수정은 막혀 있어요)</div>

  <label>계좌 종류</label>
  <div class="btn-group" id="typeGroup">
    <button onclick="sel(this,'일반')" class="${accountType === '일반' ? 'selected' : ''}">일반</button>
    <button onclick="sel(this,'ISA')"  class="${accountType === 'ISA'  ? 'selected' : ''}">ISA</button>
    <button onclick="sel(this,'모의')" class="${accountType === '모의' ? 'selected' : ''}">모의</button>
  </div>

  <label>리밸런싱 임계치 (%)</label>
  <input type="number" id="tol" value="${rebalanceTol}" min="0" step="0.5">
  <div class="row-hint">현재 비중이 목표와 이 값 이상 벌어져야 매매가 발생합니다. 기본 2.0</div>

  <label>수익실현 임계치 (%)</label>
  <input type="number" id="pt" value="${profitThreshold}" min="0" step="1">
  <div class="row-hint">종목 수익률이 이 값 이상이면 AI 브리핑에서 비중 축소를 우선 고려합니다. 기본 40</div>

  <label>연 목표 수익률 (%)</label>
  <input type="number" id="yield" value="${targetYield}" min="0" step="0.5">
  <div class="row-hint">대시보드의 "월 인출 추천" 계산에 쓰입니다. 기본 10</div>

  <div class="actions">
    <button class="save-btn" id="saveBtn" onclick="save()">저장</button>
  </div>

<script>
  var accountType = '${accountType}';

  function sel(el, val) {
    el.parentNode.querySelectorAll('button').forEach(function(b){ b.classList.remove('selected'); });
    el.classList.add('selected');
    accountType = val;
  }

  function save() {
    var btn = document.getElementById('saveBtn');
    btn.disabled = true; btn.textContent = '저장 중...';
    var data = {
      accountType: accountType,
      rebalanceTolerance: parseFloat(document.getElementById('tol').value) || 2.0,
      profitTakingThreshold: parseFloat(document.getElementById('pt').value) || 40.0,
      targetYield: parseFloat(document.getElementById('yield').value) || 10.0
    };
    google.script.run
      .withSuccessHandler(function(){ google.script.host.close(); })
      .withFailureHandler(function(e){ btn.disabled = false; btn.textContent = '저장'; alert('오류: ' + e.message); })
      .saveBasicSettings(data);
  }
<\/script>
</body>
</html>`;

  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(360).setHeight(440).setTitle('기본 설정');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔧 기본 설정');
}

/**
 * openBasicSettingsDialog에서 호출 — ⚙️ 설정 시트 B7:B10에 실제로 저장
 */
function saveBasicSettings(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️ 설정');
  if (!sheet) throw new Error('⚙️ 설정 시트를 찾을 수 없습니다.');
  sheet.getRange('B7').setValue(data.accountType || '일반');
  sheet.getRange('B8').setValue(data.rebalanceTolerance || 2.0);
  sheet.getRange('B9').setValue(data.profitTakingThreshold || 40.0);
  sheet.getRange('B10').setValue(data.targetYield || 10.0);
  updateDashboard();
}

/**
 * 포트폴리오설정 시트 컬럼 순서 마이그레이션.
 * 구 레이아웃(C=목표비율, D=유형, E=초기비율) → 신 레이아웃(C=초기비율, D=목표비율, E=유형)
 */
function addInitialRatiosColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('포트폴리오설정 시트를 찾을 수 없습니다.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const c2Header = String(sheet.getRange('C2').getValue()).trim();
  const d2Header = String(sheet.getRange('D2').getValue()).trim();

  // 이미 새 레이아웃이면 아무 것도 안 함
  if (c2Header === '기준비율(%)' && d2Header === '운용비율(%)') {
    SpreadsheetApp.getUi().alert('✅ 이미 새 컬럼 순서입니다.\n(C=초기비율, D=목표비율, E=유형)');
    return;
  }

  // 구 레이아웃 감지: C=목표비율, D=유형, E=초기비율
  const dataRows = lastRow - 2;
  if (dataRows <= 0) return;

  const oldData = sheet.getRange(3, 3, dataRows, 3).getValues(); // C, D, E 읽기
  // oldData[i] = [oldC(목표비율), oldD(유형), oldE(초기비율)]
  const newData = oldData.map(row => [
    (typeof row[2] === 'number' && row[2] > 0) ? row[2] : row[0], // 새C=초기비율(E에 값 있으면 E, 없으면 C)
    row[0], // 새D=목표비율(구C)
    row[1]  // 새E=유형(구D)
  ]);

  // 헤더 업데이트
  sheet.getRange('C2').setValue('기준비율(%)');
  sheet.getRange('D2').setValue('운용비율(%)');
  sheet.getRange('E2').setValue('유형');

  // 데이터 업데이트
  sheet.getRange(3, 3, dataRows, 3).setValues(newData);
  sheet.getRange(3, 3, dataRows, 1).setHorizontalAlignment('right'); // C: 초기비율
  sheet.getRange(3, 4, dataRows, 1).setHorizontalAlignment('right'); // D: 목표비율
  sheet.getRange(3, 5, dataRows, 1).setHorizontalAlignment('left');  // E: 유형

  SpreadsheetApp.getUi().alert(
    '✅ 컬럼 순서가 변경되었습니다.\n\n' +
    'C열: 기준비율(%) — 고정값 (처음 설정한 목표)\n' +
    'D열: 운용비율(%) — 수정 가능 (AI/사람이 조정)\n' +
    'E열: 유형\n\n' +
    '초기비율(C열)이 의도한 값과 다르면 직접 수정해 주세요.'
  );
}

// 초기 시트 설정에 수익실현 시트 추가
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 설정 시트
  let sheet = ss.getSheetByName('⚙️ 설정');
  let lastAppKey = '', lastAppSecret = '', lastAccount = '12345678-01', lastAccountType = '일반';

  if (sheet) {
    const existingValues = sheet.getRange('B2:B7').getValues();
    lastAppKey    = existingValues[0][0] || '';
    lastAppSecret = existingValues[1][0] || '';
    lastAccount   = existingValues[2][0] || '12345678-01';
    // B7(index 5): 계좌 종류, 기존 TRUE/FALSE 마이그레이션
    const raw = String(existingValues[5][0] || '');
    if (raw === 'TRUE') lastAccountType = '모의';
    else if (['일반', 'ISA', '모의'].includes(raw)) lastAccountType = raw;
    else lastAccountType = '일반';
  } else {
    sheet = ss.insertSheet('⚙️ 설정');
  }

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.getRange('A1').setValue('KIS API 설정').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange('A2:B10').setValues([
    ['APP KEY',                      '🛡️ 보안 저장됨'],
    ['APP SECRET',                   '🛡️ 보안 저장됨'],
    ['계좌번호',                     '🛡️ 보안 저장됨'],
    ['Gemini API Key',               '🛡️ 보안 저장됨'],
    ['API 키 발급처',                'https://aistudio.google.com/app/apikey'],
    ['계좌 종류 (일반/ISA/모의)',    lastAccountType],
    ['리밸런싱 임계치 (%)',           2.0],
    ['수익실현 임계치 (%)',           40.0],
    ['연 목표 수익률 (%)',            10.0]
  ]);
  sheet.getRange('A2:A10').setHorizontalAlignment('center');
  sheet.getRange('B2:B10').setHorizontalAlignment('right');
  sheet.getRange('B6').setFontColor('#1a73e8').setFontLine('underline');

  // 계좌 종류 드롭다운
  const accountRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['일반', 'ISA', '모의'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B7').setDataValidation(accountRule);

  trimExtraColumns(sheet);
  
  // 2. 대시보드 시트
  sheet = ss.getSheetByName('📊 대시보드');
  if (!sheet) {
    sheet = ss.insertSheet('📊 대시보드', 0);
  }
  setupDashboardSheet(sheet);
  
  // 3. 계좌현황 시트
  sheet = ss.getSheetByName('🏦 계좌현황');
  if (!sheet) {
    sheet = ss.insertSheet('🏦 계좌현황');
  }
  sheet.clear();
  sheet.getRange('A1').setValue('💰 계좌 현황').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange('A2:B5').setValues([
    ['예수금', 0],
    ['주문가능금액', 0],
    ['업데이트 시간', ''],
    ['총 평가액', 0]
  ]);
  sheet.getRange('A2:A5').setHorizontalAlignment('center');
  sheet.getRange('B2:B5').setHorizontalAlignment('right');
  trimExtraColumns(sheet);

  sheet.getRange('A6').setValue('💰 전체 수익률').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A7:I7').setValues([[
    '종목코드', '종목명', '보유수량', '평균단가', '현재가', '평가금액', '손익', '수익률(%)', '목표여부'
  ]]).setFontWeight('bold').setBackground('#4285f4').setFontColor('white').setHorizontalAlignment('center');
  
  // 4. 포트폴리오설정 시트
  sheet = ss.getSheetByName('📋 포트폴리오설정');
  if (!sheet) {
    sheet = ss.insertSheet('📋 포트폴리오설정');
  }
  // 포트폴리오설정 시트: 기존 초기비율(C열) 보존 후 재설정
  const prevPortLastRow = sheet.getLastRow();
  const savedInitialRatios = {};
  if (prevPortLastRow >= 3) {
    const existingData = sheet.getRange(3, 1, prevPortLastRow - 2, 3).getValues();
    existingData.forEach(row => {
      const code = String(row[0]).trim();
      const initRatio = row[2]; // C열: 새 레이아웃에서는 초기비율, 구 레이아웃에서는 목표비율
      if (code && typeof initRatio === 'number' && initRatio > 0) {
        savedInitialRatios[code] = initRatio;
      }
    });
  }

  sheet.clear();
  sheet.getRange('A1').setValue('🎯 목표 포트폴리오').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  // 컬럼 순서: 종목코드 | 종목명 | 기준비율(%) [고정] | 운용비율(%) [수정가능] | 유형
  sheet.getRange('A2:E2').setValues([['종목코드', '종목명', '기준비율(%)', '운용비율(%)', '유형']])
    .setFontWeight('bold').setBackground('#34a853').setFontColor('white').setHorizontalAlignment('center');

  const defaultPortfolio = [
    ['440650', 'ACE 미국달러단기채권액티브', 30, 30, '채권'],
    ['319640', 'TIGER 골드선물(H)', 20, 20, '금'],
    ['261240', 'KODEX 미국달러선물', 10, 10, '달러'],
    ['161510', 'PLUS 고배당주', 15, 15, '국내주식'],
    ['315960', 'RISE 대형고배당10TR', 10, 10, '국내주식'],
    ['379800', 'KODEX 미국S&P500TR', 10, 10, '해외주식'],
    ['', '현금', 5, 5, '현금']
  ];

  // 기존 초기비율이 있으면 C열(초기비율)에 복원
  const dataWithRestoredInitial = defaultPortfolio.map(row => {
    const code = row[0];
    const saved = code ? savedInitialRatios[code] : null;
    return [row[0], row[1], saved || row[2], row[3], row[4]];
  });

  sheet.getRange(3, 1, dataWithRestoredInitial.length, 5).setValues(dataWithRestoredInitial);
  sheet.getRange('A3:E').setHorizontalAlignment('right');
  sheet.getRange('B3:B9').setHorizontalAlignment('left');
  
  // 5. 거래내역 시트
  setupTradeHistorySheet();

  // 6. 수익실현기록 시트
  setupProfitHistorySheet();

  // 7. 비중변경이력 시트
  setupAIHistorySheet();

  // 8. 기술지표이력 시트
  setupTAHistorySheet();
  
  // 9. 자동 새로고침 트리거 기본 생성 (onOpen)
  const triggers = ScriptApp.getProjectTriggers();
  const hasRefreshTrigger = triggers.some(t => t.getHandlerFunction() === 'automatedRefreshRoutine');
  if (!hasRefreshTrigger) {
    ScriptApp.newTrigger('automatedRefreshRoutine')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
  }
  
  // 초기 템플릿 세팅 시 버전을 동기화합니다.
  syncTemplateVersion();
  
  SpreadsheetApp.getUi().alert('✅ 초기 설정이 완료되었습니다!\n\n곧바로 API 키 보안 설정 창이 열립니다.\n금융 정보를 안전하게 입력해 주세요.');
  
  // UX 개선: 작업 끝나고 이어서 바로 API 키 입력 창 열기
  openSecureConfigDialog();
}