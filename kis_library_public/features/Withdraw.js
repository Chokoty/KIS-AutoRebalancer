/**
 * 수익 실현 창 열기 — 대시보드를 새로고침한 뒤 데이터 읽어 오픈
 */
function openWithdrawDialog() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('최신 데이터로 업데이트 중...', '⏳ 처리 중', -1);
  updateDashboard();
  const dash = ss.getSheetByName('📊 대시보드');
  var totalEval = 0, cash = 0, profit = 0, recommended = 0;
  var holdings = [];
  var sellFeeRate = 0.0015;
  if (dash) {
    totalEval   = dash.getRange('B3').getValue() || 0;
    cash        = dash.getRange('B4').getValue() || 0;
    profit      = dash.getRange('H4').getValue() || 0;
    recommended = dash.getRange('H7').getValue() || 0;
    // 보유종목 데이터 수집 (계산을 클라이언트에서 처리하기 위해)
    var lastRow = dash.getLastRow();
    if (lastRow >= 9) {
      var rows = dash.getRange(9, 1, lastRow - 8, 8).getValues();
      rows.forEach(function(row) {
        var code = String(row[0]).trim();
        var qty  = parseInt(row[3]) || 0;
        var price = parseFloat(row[4]) || 0;
        var evalAmt = parseFloat(row[5]) || 0;
        var ratioRaw = row[7];
        // Google Sheets auto-converts "20.00%" strings to 0.20 numbers; normalize to percentage form
        var ratio = typeof ratioRaw === 'number'
          ? (ratioRaw > 0 && ratioRaw <= 1 ? ratioRaw * 100 : ratioRaw)
          : parseFloat(String(ratioRaw).replace('%','')) || 0;
        if (code && qty > 0 && price > 0 && ratio > 0) {
          holdings.push({ code: code, name: row[1], qty: qty, price: price, evalAmt: evalAmt, ratio: ratio });
        }
      });
    }
  }
  try { sellFeeRate = getConfig().sellFeeRate || 0.0015; } catch(e) {}

  if (totalEval === 0) {
    SpreadsheetApp.getUi().alert('⚠️ 먼저 대시보드를 새로고침하세요.\n(KIS AutoTrader → 대시보드 새로고침)');
    return;
  }

  const initJson = JSON.stringify({
    totalEval: totalEval, cash: cash, profit: profit,
    recommended: Math.round(recommended),
    holdings: holdings, fee: sellFeeRate
  });

  const html = HtmlService.createHtmlOutput('<!DOCTYPE html><html><head><style>' +
'* { box-sizing: border-box; }' +
'body { font-family: Malgun Gothic, sans-serif; padding: 16px; color: #3c4043; font-size: 13px; margin: 0; }' +
'.cards { display: flex; gap: 8px; margin-bottom: 14px; }' +
'.card { flex: 1; background: #f8f9fa; border-radius: 8px; padding: 10px 12px; }' +
'.card-label { font-size: 11px; color: #5f6368; margin-bottom: 2px; }' +
'.card-value { font-size: 16px; font-weight: bold; color: #1a73e8; }' +
'.section-title { font-weight: bold; margin-bottom: 6px; }' +
'.input-row { display: flex; gap: 8px; margin-bottom: 10px; align-items: center; }' +
'input[type=number] { flex: 1; padding: 9px 12px; border: 2px solid #4285f4; border-radius: 6px; font-size: 15px; outline: none; }' +
'.btn { padding: 9px 16px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 13px; }' +
'.btn:disabled { opacity: 0.5; cursor: not-allowed; }' +
'.btn-blue { background: #4285f4; color: white; }' +
'.btn-green { background: #34a853; color: white; }' +
'.btn-red { background: #ea4335; color: white; width: 100%; margin-top: 10px; }' +
'.btn-outline { background: white; border: 1px solid #dadce0; color: #5f6368; }' +
'.calc-row { display: flex; gap: 8px; margin-bottom: 14px; }' +
'table { width: 100%; border-collapse: collapse; font-size: 12px; }' +
'th { background: #ea4335; color: white; padding: 5px 7px; text-align: center; }' +
'td { padding: 4px 7px; border-bottom: 1px solid #f1f3f4; }' +
'td.num { text-align: right; } td.sell { text-align: right; color: #ea4335; font-weight: bold; }' +
'td.amount { text-align: right; color: #1a73e8; font-weight: bold; }' +
'td.ratio-up { text-align: right; color: #34a853; font-weight: bold; }' +
'td.ratio-dn { text-align: right; color: #ea4335; font-weight: bold; }' +
'td.ratio-nc { text-align: right; color: #5f6368; }' +
'.tbl-wrap { max-height: 220px; overflow-y: auto; border: 1px solid #f1f3f4; border-radius: 4px; margin-bottom: 10px; }' +
'tr.cash-row td { color: #5f6368; background: #f8f9fa; font-style: italic; }' +
'.summary { background: #f8f9fa; border-radius: 6px; padding: 10px 12px; margin-top: 10px; }' +
'.summary-row { display: flex; justify-content: space-between; padding: 4px 0; border-bottom: 1px solid #ececec; font-size: 13px; }' +
'.summary-row:last-child { border: none; font-weight: bold; font-size: 14px; }' +
'.ok { color: #137333; } .warn { color: #f9ab00; } .info { color: #1a73e8; }' +
'#planSection { display: none; margin-top: 14px; } #executeBtn { display: none; }' +
'</style></head><body>' +
'<div class="cards">' +
'<div class="card"><div class="card-label">총 자산</div><div class="card-value" id="totalEval">-</div></div>' +
'<div class="card"><div class="card-label">예수금</div><div class="card-value" id="cashVal">-</div></div>' +
'<div class="card"><div class="card-label">총 손익</div><div class="card-value" id="profitVal">-</div></div>' +
'</div>' +
'<div class="section-title">&#128176; 실현 금액 입력</div>' +
'<div class="input-row"><input type="number" id="amountInput" placeholder="금액 입력 (원)" min="0" step="10000">' +
'<button class="btn btn-green" id="recBtn" onclick="useRecommended()">추천</button></div>' +
'<div class="calc-row"><button class="btn btn-blue" id="calcBtn" onclick="calculate()" style="flex:1">&#128290; 계산하기</button>' +
'<button class="btn btn-outline" onclick="google.script.host.close()">닫기</button></div>' +
'<div id="planSection"><div class="section-title">&#128203; 전체 비중 변화 (비율 유지)</div>' +
'<div class="tbl-wrap"><table><thead><tr><th style="text-align:left">종목명</th><th>현재</th><th>매도</th><th>잔여</th><th>현재%</th><th>실현후%</th></tr></thead>' +
'<tbody id="planTable"></tbody></table></div>' +
'<div class="summary">' +
'<div class="summary-row"><span>총 매도 예정액</span><span id="totalSell">-</span></div>' +
'<div class="summary-row"><span>실현 후 예수금</span><span id="cashAfter">-</span></div>' +
'<div class="summary-row"><span>상태</span><span id="statusLabel">-</span></div>' +
'</div>' +
'<button class="btn btn-red" id="executeBtn" onclick="execute()">&#10003; 수익 실현 &mdash; 실제 매도 주문 진행</button>' +
'</div>' +
'<script>' +
'var planData=null; var d=' + initJson + ';' +
'function fmt(n){return Math.round(n).toLocaleString("ko-KR")+"원";}' +
'document.getElementById("totalEval").textContent=fmt(d.totalEval);' +
'document.getElementById("cashVal").textContent=fmt(d.cash);' +
'document.getElementById("profitVal").textContent=(d.profit>=0?"+":"")+fmt(d.profit);' +
'document.getElementById("profitVal").style.color=d.profit>=0?"#ea4335":"#1a73e8";' +
'document.getElementById("recBtn").textContent="추천 "+fmt(d.recommended);' +
'if(d.recommended>0)document.getElementById("amountInput").value=d.recommended;' +
'function useRecommended(){document.getElementById("amountInput").value=d.recommended;}' +
'function calcPlan(amount){' +
'  var fee=d.fee; var targetTotal=d.totalEval-amount;' +
'  var orders=[]; var totalSell=0; var sellMap={};' +
'  d.holdings.forEach(function(h){' +
'    var targetAmt=targetTotal*(h.ratio/100);' +
'    var excess=h.evalAmt-targetAmt;' +
'    if(excess<=0)return;' +
'    var sellQty=Math.min(h.qty,Math.ceil(excess/(h.price*(1-fee))));' +
'    if(sellQty<=0)return;' +
'    var rawAmt=sellQty*h.price;' +
'    var proceeds=Math.floor(rawAmt*(1-fee));' +
'    totalSell+=proceeds;' +
'    sellMap[h.code]={sellQty:sellQty,rawAmt:rawAmt,proceeds:proceeds};' +
'    orders.push({name:h.name,code:h.code,currentQty:h.qty,currentPrice:h.price,sellQty:sellQty,sellAmount:rawAmt,actualProceeds:proceeds,remainingQty:h.qty-sellQty});' +
'  });' +
'  var available=totalSell+d.cash;' +
'  var cashAfter=available-amount;' +
'  var newTotal=cashAfter;' +
'  d.holdings.forEach(function(h){var sk=sellMap[h.code]; var remQty=sk?h.qty-sk.sellQty:h.qty; newTotal+=remQty*h.price;});' +
'  var allRows=d.holdings.map(function(h){' +
'    var sk=sellMap[h.code]; var sellQty=sk?sk.sellQty:0; var remQty=h.qty-sellQty;' +
'    return{name:h.name,code:h.code,currentQty:h.qty,sellQty:sellQty,remainingQty:remQty,' +
'      sellAmount:sk?sk.rawAmt:0,' +
'      beforeRatio:d.totalEval>0?h.evalAmt/d.totalEval*100:0,' +
'      afterRatio:newTotal>0?remQty*h.price/newTotal*100:0,' +
'      targetRatio:h.ratio};' +
'  });' +
'  var cashBeforeRatio=d.totalEval>0?d.cash/d.totalEval*100:0;' +
'  var cashAfterRatio=newTotal>0?cashAfter/newTotal*100:0;' +
'  var shortage=Math.max(0,amount-available);' +
'  var surplus=Math.max(0,available-amount);' +
'  return{orders:orders,allRows:allRows,totalSell:totalSell,cashAfter:cashAfter,' +
'    cashBeforeRatio:cashBeforeRatio,cashAfterRatio:cashAfterRatio,' +
'    shortage:shortage,surplus:surplus,statusClass:shortage>1000?"warn":"ok",withdrawAmount:amount};' +
'}' +
'function calculate(){' +
'  var amount=parseInt(document.getElementById("amountInput").value)||0;' +
'  if(amount<=0){alert("실현 금액을 입력하세요.");return;}' +
'  var plan=calcPlan(amount);' +
'  planData=plan;' +
'  document.getElementById("planSection").style.display="block";' +
'  var tbody=document.getElementById("planTable"); tbody.innerHTML="";' +
'  function ratioClass(diff){return diff<-0.15?"ratio-dn":diff>0.15?"ratio-up":"ratio-nc";}' +
'  plan.allRows.forEach(function(r){' +
'    var tr=document.createElement("tr");' +
'    var diff=r.afterRatio-r.beforeRatio;' +
'    var sellCell=r.sellQty>0?"<td class=sell>"+r.sellQty+"주</td>":"<td class=num style=color:#bbb>-</td>";' +
'    var arrow=diff<-0.15?"▼":diff>0.15?"▲":"";' +
'    tr.innerHTML="<td>"+r.name+"</td><td class=num>"+r.currentQty+"주</td>"+sellCell+"<td class=num>"+r.remainingQty+"주</td>"+' +
'      "<td class=num>"+r.beforeRatio.toFixed(1)+"%</td>"+' +
'      "<td class="+ratioClass(diff)+">"+r.afterRatio.toFixed(1)+"% "+arrow+"</td>";' +
'    tbody.appendChild(tr);' +
'  });' +
'  var cashTr=document.createElement("tr"); cashTr.className="cash-row";' +
'  var cd=plan.cashAfterRatio-plan.cashBeforeRatio;' +
'  var ca=cd<-0.15?"▼":cd>0.15?"▲":"";' +
'  cashTr.innerHTML="<td>예수금(현금)</td><td class=num colspan=3>"+Math.round(plan.cashAfter).toLocaleString()+"원</td>"+' +
'    "<td class=num>"+plan.cashBeforeRatio.toFixed(1)+"%</td>"+' +
'    "<td class="+ratioClass(cd)+">"+plan.cashAfterRatio.toFixed(1)+"% "+ca+"</td>";' +
'  tbody.appendChild(cashTr);' +
'  document.getElementById("totalSell").textContent=fmt(plan.totalSell);' +
'  document.getElementById("cashAfter").textContent=fmt(plan.cashAfter);' +
'  var st=document.getElementById("statusLabel");' +
'  if(plan.statusClass==="warn"){st.className="warn";st.textContent="⚠️ "+fmt(plan.shortage)+" 부족";}' +
'  else if(plan.surplus>1000){st.className="info";st.textContent="ℹ️ "+fmt(plan.surplus)+" 여유";}' +
'  else{st.className="ok";st.textContent="✅ 실현 가능";}' +
'  document.getElementById("executeBtn").style.display=plan.statusClass!=="warn"?"block":"none";' +
'}' +
'function execute(){' +
'  if(!planData)return;' +
'  var amount=parseInt(document.getElementById("amountInput").value)||0;' +
'  if(!confirm(amount.toLocaleString()+"원 수익 실현을 위한 매도 주문을 실행합니다.\\n실제 매매가 진행됩니다. 계속하시겠습니까?"))return;' +
'  var btn=document.getElementById("executeBtn");' +
'  btn.disabled=true; btn.textContent="주문 실행 중...";' +
'  google.script.run' +
'    .withSuccessHandler(function(r){alert("완료\\n성공: "+r.success+"건 / 실패: "+r.fail+"건");google.script.host.close();})' +
'    .withFailureHandler(function(e){btn.disabled=false;btn.textContent="✅ 수익 실현 — 실제 매도 주문 진행";alert("오류: "+e.message);})' +
'    .executeWithdrawPlan(planData);' +
'}' +
'<\/script></body></html>'
  ).setWidth(520).setHeight(660).setTitle('💰 수익 실현');

  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * 다이얼로그용: 매도 주문 실행 + 인출 금액 2주 보호 설정
 */
function executeWithdrawPlan(planData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet    = ss.getSheetByName('📝 거래내역');
  const profitSheet = ss.getSheetByName('📝 수익실현기록');
  let successCount = 0, failCount = 0;

  // 대시보드에서 종목별 평균단가 수집 (col 15 = avgPrice)
  const avgPriceMap = {};
  const dash = ss.getSheetByName('📊 대시보드');
  if (dash && dash.getLastRow() >= 9) {
    dash.getRange(9, 1, dash.getLastRow() - 8, 15).getValues().forEach(function(row) {
      const code = String(row[0]).trim();
      if (code) avgPriceMap[code] = parseFloat(row[14]) || 0;
    });
  }

  planData.orders.forEach(order => {
    const result = placeOrder(order.code, 'sell', order.sellQty, 0);
    if (result.success) {
      successCount++;
      // 수익실현기록 시트에 기록
      if (profitSheet) {
        const avgPrice = avgPriceMap[order.code] || 0;
        const realizedProfit = avgPrice > 0 ? (order.currentPrice - avgPrice) * order.sellQty : 0;
        const lastRow = profitSheet.getLastRow();
        const cumulative = realizedProfit + (lastRow >= 2 ? (parseFloat(profitSheet.getRange(lastRow, 8).getValue()) || 0) : 0);
        profitSheet.appendRow([
          new Date(), order.name, '매도(수익실현)',
          order.sellQty, order.currentPrice, order.sellQty * order.currentPrice,
          Math.round(realizedProfit), Math.round(cumulative)
        ]);
        trimExtraColumns(profitSheet, 8);
      }
    } else {
      failCount++;
    }
    if (logSheet) {
      logSheet.appendRow([
        new Date(), '매도(수익실현)', order.code, order.name,
        order.sellQty, order.currentPrice, order.sellQty * order.currentPrice,
        result.success ? '성공' : '실패', result.message || ''
      ]);
      trimExtraColumns(logSheet, 9);
    }
    Utilities.sleep(300);
  });

  // 매도 성공 시 인출 금액을 2주간 리밸런싱 예수금에서 보호
  if (successCount > 0 && planData.withdrawAmount > 0) {
    setProtectedCash(planData.withdrawAmount);
  }

  return { success: successCount, fail: failCount };
}

/**
 * 인출 금액을 2주간 리밸런싱 예수금에서 보호 설정
 */
function setProtectedCash(amount) {
  const props  = PropertiesService.getScriptProperties();
  const expiry = new Date();
  expiry.setDate(expiry.getDate() + 14);
  props.setProperty('PROTECTED_CASH_AMOUNT', String(Math.round(amount)));
  props.setProperty('PROTECTED_CASH_EXPIRY',  expiry.toISOString());
}

/**
 * 보호 예수금 조회 (만료 시 자동 해제)
 * @returns {{ amount: number, daysLeft: number }}
 */
function getProtectedCash() {
  const props     = PropertiesService.getScriptProperties();
  const amount    = parseFloat(props.getProperty('PROTECTED_CASH_AMOUNT') || '0');
  const expiryStr = props.getProperty('PROTECTED_CASH_EXPIRY');
  if (!amount || !expiryStr) return { amount: 0, daysLeft: 0 };

  const expiry = new Date(expiryStr);
  if (new Date() > expiry) {
    props.deleteProperty('PROTECTED_CASH_AMOUNT');
    props.deleteProperty('PROTECTED_CASH_EXPIRY');
    return { amount: 0, daysLeft: 0 };
  }

  const daysLeft = Math.max(1, Math.ceil((expiry - new Date()) / 86400000));
  return { amount, daysLeft };
}

/**
 * 보호 예수금 수동 해제
 */
function releaseProtectedCash() {
  const info = getProtectedCash();
  if (info.amount <= 0) {
    SpreadsheetApp.getUi().alert('현재 보호된 예수금이 없습니다.');
    return;
  }
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('PROTECTED_CASH_AMOUNT');
  props.deleteProperty('PROTECTED_CASH_EXPIRY');
  SpreadsheetApp.getUi().alert(
    `🔓 보호 해제 완료\n${info.amount.toLocaleString()}원이 이제 리밸런싱 예수금에 포함됩니다.`
  );
  updateDashboard();
}

// Dummy
