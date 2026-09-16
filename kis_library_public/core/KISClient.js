// KIS API 호출 공통 함수
function callKISAPI(endpoint, trId, params = {}) {
  const config = getConfig();
  const token = getAccessToken();
  
  if (!token) {
    throw new Error('액세스 토큰을 가져올 수 없습니다.');
  }
  
  const url = `${config.baseUrl}${endpoint}`;
  
  // null, undefined만 제외하고 빈 문자열은 포함
  const queryParams = Object.keys(params)
    .filter(key => {
      const value = params[key];
      return value !== null && value !== undefined;
    })
    .map(key => `${key}=${encodeURIComponent(params[key])}`)
    .join('&');
  
  const fullUrl = queryParams ? `${url}?${queryParams}` : url;
  
  Logger.log('API 호출 URL: ' + fullUrl);
  
  const headers = {
    'Content-Type': 'application/json',
    'authorization': `Bearer ${token}`,
    'appkey': config.appKey,
    'appsecret': config.appSecret,
    'tr_id': trId
  };
  
  const options = {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(fullUrl, options);
    const responseText = response.getContentText();
    const data = JSON.parse(responseText);
    
    Logger.log('API 응답 코드: ' + data.rt_cd + ', 메시지: ' + data.msg1);
    
    // 토큰 만료 에러 ('EGW00001' 등) 또는 만료 메시지 처리
    const isTokenExpired = data.msg_cd === 'EGW00001' || (data.msg1 && data.msg1.includes('만료'));
    
    if (isTokenExpired && !params._isRetry) {
      Logger.log('토큰 만료 감지, 재발급 후 재시도합니다: ' + (data.msg_cd || data.msg1));
      
      const props = PropertiesService.getUserProperties();
      const prefix = config.appKey.substring(0, 8);
      props.deleteProperty('KIS_TOKEN_' + prefix);
      props.deleteProperty('KIS_EXPIRY_' + prefix);
      
      params._isRetry = true;
      return callKISAPI(endpoint, trId, params);
    }
    
    if (data.rt_cd === '0') {
      return data;
    } else {
      throw new Error(data.msg1 || 'API 호출 실패');
    }
  } catch (e) {
    Logger.log('API 호출 실패: ' + e);
    throw e;
  }
}

// 예수금과 주문가능금액 조회 (단일 API 사용)
function getBalance() {
  const config = getConfig();
  const trId = config.isMock ? 'VTTC8434R' : 'TTTC8434R';
  
  const params = {
    'CANO': config.account.split('-')[0],
    'ACNT_PRDT_CD': config.account.split('-')[1] || '01',
    'AFHR_FLPR_YN': 'N',
    'OFL_YN': '',
    'INQR_DVSN': '01',
    'UNPR_DVSN': '01',
    'FUND_STTL_ICLD_YN': 'N',
    'FNCG_AMT_AUTO_RDPT_YN': 'N',
    'PRCS_DVSN': '01',
    'CTX_AREA_FK100': '',
    'CTX_AREA_NK100': ''
  };
  
  try {
    const data = callKISAPI('/uapi/domestic-stock/v1/trading/inquire-balance', trId, params);
    const output2 = data.output2 && data.output2.length > 0 ? data.output2[0] : {};
    
    // 총평가금액 (API에서 계산된 값)
    const totalEval = parseInt(output2.tot_evlu_amt || 0);

    // 보유주식 평가액
    const holdingsValue = parseInt(output2.scts_evlu_amt || 0);

    // 실제 예수금(dnca_tot_amt): D+2 미결제 포함 없는 실제 현금
    // tot_evlu_amt - scts_evlu_amt는 미결제 매도 수령금을 포함할 수 있어 과대계상됨
    const dnca = parseInt(output2.dnca_tot_amt || 0);
    const buyPowerCalc = totalEval - holdingsValue;
    // 두 값 중 작은 값 사용 (보수적 추정으로 과매수 방지)
    const buyPower = dnca > 0 ? Math.min(dnca, buyPowerCalc) : buyPowerCalc;

    Logger.log('총평가액(tot_evlu_amt): ' + totalEval);
    Logger.log('보유주식(scts_evlu_amt): ' + holdingsValue);
    Logger.log('예수금(dnca_tot_amt): ' + dnca);
    Logger.log('주문가능금액(보수적): ' + buyPower + ' (계산값: ' + buyPowerCalc + ')');

    return {
      cash: buyPower,        // 실제 현금 = 보수적 주문가능금액
      buyPower: buyPower,    // 주문가능금액
      totalEval: totalEval,  // 총평가액
      holdingsValue: holdingsValue  // 보유주식평가액
    };
    
  } catch (e) {
    Logger.log('예수금 조회 실패: ' + e.toString());
    throw e;
  }
}

// 보유 주식 조회 (수정 - 보유주식 평가액 반환)
function getHoldings() {
  const config = getConfig();
  const trId = config.isMock ? 'VTTC8434R' : 'TTTC8434R';
  
  const params = {
    'CANO': config.account.split('-')[0],
    'ACNT_PRDT_CD': config.account.split('-')[1] || '01',
    'AFHR_FLPR_YN': 'N',
    'OFL_YN': '',
    'INQR_DVSN': '01',
    'UNPR_DVSN': '01',
    'FUND_STTL_ICLD_YN': 'N',
    'FNCG_AMT_AUTO_RDPT_YN': 'N',
    'PRCS_DVSN': '01',
    'CTX_AREA_FK100': '',
    'CTX_AREA_NK100': ''
  };
  
  try {
    const data = callKISAPI('/uapi/domestic-stock/v1/trading/inquire-balance', trId, params);
    const holdings = [];
    
    if (data.output1 && Array.isArray(data.output1)) {
      data.output1.forEach(item => {
        const quantity = parseInt(item.hldg_qty || 0); // 총 보유수량
        const ordPsblQty = parseInt(item.ord_psbl_qty || quantity); // 주문가능수량 (없으면 전체보유)
        const evalAmount = parseInt(item.evlu_amt || 0);
        
        if (item.pdno && (quantity > 0 || evalAmount > 0)) {
          const avgPrice = parseFloat(item.pchs_avg_pric || 0);
          const currentPrice = parseFloat(item.prpr || 0);
          const profit = parseInt(item.evlu_pfls_amt || 0);
          
          let profitRate = parseFloat(item.evlu_pfls_rt || 0);
          profitRate = profitRate / 100;
          
          holdings.push({
            code: item.pdno,
            name: item.prdt_name,
            quantity: quantity,
            ordPsblQty: ordPsblQty, // 추가: 실제로 매도 가능한 수량
            avgPrice: avgPrice,
            currentPrice: currentPrice,
            evalAmount: evalAmount,
            profit: profit,
            profitRate: profitRate
          });
        }
      });
    }
    
    Logger.log('파싱된 보유주식: ' + holdings.length + '개');
    
    // 보유주식 총 평가액도 반환
    const totalHoldingsValue = holdings.reduce((sum, h) => sum + h.evalAmount, 0);
    Logger.log('보유주식 총 평가액: ' + totalHoldingsValue);
    
    return holdings;
    
  } catch (e) {
    Logger.log('보유주식 조회 실패: ' + e.toString());
    throw e;
  }
}

// 현재가 조회
function getCurrentPrice(stockCode) {
  const config = getConfig();
  const trId = 'FHKST01010100';
  
  const params = {
    'FID_COND_MRKT_DIV_CODE': 'J',
    'FID_INPUT_ISCD': stockCode
  };
  
  try {
    const data = callKISAPI('/uapi/domestic-stock/v1/quotations/inquire-price', trId, params);
    const output = data.output || {};
    const price = parseInt(output.stck_prpr || 0);
    
    Logger.log(`${stockCode} 현재가: ${price}`);
    
    // API 호출 제한 방지
    Utilities.sleep(100);
    
    return price;
  } catch (e) {
    Logger.log(`현재가 조회 실패 (${stockCode}): ` + e.toString());
    return 0;
  }
}

// 주식 주문 (매수/매도) - 최종 수정
function placeOrder(stockCode, orderType, quantity, price = 0) {
  const config = getConfig();
  const isBuy = orderType === 'buy';
  const trId = config.isMock 
    ? (isBuy ? 'VTTC0802U' : 'VTTC0801U')
    : (isBuy ? 'TTTC0802U' : 'TTTC0801U');
  
  Logger.log(`=== 주문 요청 ===`);
  Logger.log(`종목코드: ${stockCode}`);
  Logger.log(`매수/매도: ${orderType} (isBuy: ${isBuy})`);
  Logger.log(`수량: ${quantity}`);
  Logger.log(`가격: ${price}`);
  Logger.log(`TR_ID: ${trId}`);
  
  const url = `${config.baseUrl}/uapi/domestic-stock/v1/trading/order-cash`;
  const token = getAccessToken();
  
  // 시장가일 때는 "0", 지정가일 때는 실제 가격
  const orderPrice = price > 0 ? price.toString() : '0';
  
  const payload = {
    'CANO': config.account.split('-')[0],
    'ACNT_PRDT_CD': config.account.split('-')[1] || '01',
    'PDNO': stockCode,
    'ORD_DVSN': price > 0 ? '00' : '01', // 00: 지정가, 01: 시장가
    'ORD_QTY': quantity.toString(),
    'ORD_UNPR': orderPrice // 시장가: "0", 지정가: 실제가격
  };
  
  Logger.log('요청 payload: ' + JSON.stringify(payload, null, 2));
  
  const headers = {
    'Content-Type': 'application/json',
    'authorization': `Bearer ${token}`,
    'appkey': config.appKey,
    'appsecret': config.appSecret,
    'tr_id': trId,
    'custtype': 'P'
  };
  
  const options = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const data = JSON.parse(responseText);
    
    Logger.log('=== 주문 응답 ===');
    Logger.log('응답 코드: ' + data.rt_cd);
    Logger.log('메시지: ' + data.msg1);
    Logger.log('전체 응답: ' + responseText);
    
    if (data.rt_cd === '0') {
      const output = data.output || {};
      return {
        success: true,
        orderNo: output.KRX_FWDG_ORD_ORGNO || output.ODNO || output.ORD_NO || 'N/A',
        message: data.msg1
      };
    } else {
      return {
        success: false,
        message: `[${data.msg_cd}] ${data.msg1}`
      };
    }
  } catch (e) {
    Logger.log('주문 실패 예외: ' + e.toString());
    return {
      success: false,
      message: e.message
    };
  }
}