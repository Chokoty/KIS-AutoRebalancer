/**
 * 🚨 종목 즉시 전량매도 — 목표 비중과 무관하게, 지금 이 순간 특정 종목을 전부 시장가로 정리한다.
 * 목표 비중을 0%로 바꾸는 것과는 다르다 — 그건 "다음 리밸런싱에서 제외"일 뿐 실제 매도로 이어지지 않는다.
 */
function openEmergencySellDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('보유 종목 확인 중...', '⏳ 처리 중', -1);
  updateDashboard();

  const dash = ss.getSheetByName('📊 대시보드');
  const holdings = [];
  if (dash && dash.getLastRow() >= 9) {
    dash.getRange(9, 1, dash.getLastRow() - 8, 6).getValues().forEach(function(row) {
      const code = String(row[0]).trim();
      const qty  = parseInt(row[3]) || 0;
      const price = parseFloat(row[4]) || 0;
      const evalAmt = parseFloat(row[5]) || 0;
      if (code && qty > 0) {
        holdings.push({ code: code, name: String(row[1]), type: String(row[2]), qty: qty, price: price, evalAmt: evalAmt });
      }
    });
  }

  if (holdings.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ 보유 종목이 없습니다.\n(먼저 대시보드를 새로고침해보세요)');
    return;
  }

  const initJson = JSON.stringify(holdings);

  const html = '<!DOCTYPE html><html><head><style>' +
'*{box-sizing:border-box;}' +
'body{font-family:"Malgun Gothic",sans-serif;padding:16px;margin:0;font-size:13px;color:#3c4043;}' +
'.title{font-size:16px;font-weight:bold;color:#c5221f;margin-bottom:6px;}' +
'.warn{font-size:12px;color:#5f6368;margin-bottom:14px;line-height:1.5;}' +
'table{width:100%;border-collapse:collapse;font-size:12px;}' +
'th{background:#fce8e6;color:#c5221f;padding:6px 7px;text-align:center;}' +
'td{padding:6px 7px;border-bottom:1px solid #f1f3f4;}' +
'td.num{text-align:right;}' +
'.btn-sell{background:#ea4335;color:white;border:none;border-radius:4px;padding:5px 10px;font-size:11px;font-weight:bold;cursor:pointer;white-space:nowrap;}' +
'.btn-sell:disabled{opacity:.5;cursor:not-allowed;}' +
'.footer{margin-top:14px;text-align:right;}' +
'.btn-close{padding:8px 16px;border:1px solid #dadce0;border-radius:4px;background:white;cursor:pointer;}' +
'</style></head><body>' +
'<div class="title">🚨 종목 즉시 전량매도</div>' +
'<div class="warn">목표 비중과 무관하게 지금 즉시 시장가로 전량 매도합니다. 실행하면 되돌릴 수 없습니다.</div>' +
'<table><thead><tr><th style="text-align:left">종목명</th><th>보유수량</th><th>평가액</th><th></th></tr></thead>' +
'<tbody id="tbody"></tbody></table>' +
'<div class="footer"><button class="btn-close" onclick="google.script.host.close()">닫기</button></div>' +
'<script>' +
'var holdings=' + initJson + ';' +
'function fmt(n){return Math.round(n).toLocaleString("ko-KR");}' +
'function render(){' +
'  var tbody=document.getElementById("tbody");' +
'  tbody.innerHTML=holdings.map(function(h,i){' +
'    return "<tr><td>"+h.name+"</td><td class=num>"+h.qty+"주</td><td class=num>"+fmt(h.evalAmt)+"원</td>"+' +
'      "<td><button class=btn-sell id=btn"+i+" onclick=\\"sellAll("+i+")\\">🔴 즉시 매도</button></td></tr>";' +
'  }).join("");' +
'}' +
'function sellAll(i){' +
'  var h=holdings[i];' +
'  if(!confirm(h.name+" "+h.qty+"주를 지금 전부 시장가로 매도합니다.\\n실제 매매가 진행되며 되돌릴 수 없습니다.\\n계속하시겠습니까?"))return;' +
'  var btn=document.getElementById("btn"+i);' +
'  btn.disabled=true;btn.textContent="매도 중...";' +
'  google.script.run' +
'    .withSuccessHandler(function(r){' +
'      if(r.success){btn.textContent="✅ 완료";alert(h.name+" "+h.qty+"주 매도 주문이 완료되었습니다.");}' +
'      else{btn.disabled=false;btn.textContent="🔴 즉시 매도";alert("❌ 매도 실패: "+r.message);}' +
'    })' +
'    .withFailureHandler(function(e){btn.disabled=false;btn.textContent="🔴 즉시 매도";alert("오류: "+e.message);})' +
'    .executeEmergencySell(h.code,h.name,h.qty,h.price);' +
'}' +
'render();' +
'<\/script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(520),
    ' '
  );
}

/**
 * openEmergencySellDialog에서 호출 — 실제 시장가 매도 주문 실행 + 거래내역 기록
 */
function executeEmergencySell(code, name, qty, price) {
  const result = placeOrder(code, 'sell', qty, 0);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('📝 거래내역');
  if (logSheet) {
    logSheet.appendRow([
      new Date(), '매도(긴급)', code, name, qty, price, qty * price,
      result.success ? '성공' : '실패', result.message || ''
    ]);
    trimExtraColumns(logSheet, 9);
  }

  if (result.success) updateDashboard();
  return result;
}
