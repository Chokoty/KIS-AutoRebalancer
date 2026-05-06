// 설정 관리
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️ 설정');
  const props = PropertiesService.getUserProperties();
  
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // 보안 저장소(UserProperties) 우선 조회, 없으면 시트에서 조회
  // 시트별 격리를 위해 Spreadsheet ID를 접두어로 사용
  const secureAppKey = props.getProperty(ssId + '_KIS_APP_KEY');
  const secureAppSecret = props.getProperty(ssId + '_KIS_APP_SECRET');
  const secureAccount = props.getProperty(ssId + '_KIS_ACCOUNT');

  return {
    appKey: (secureAppKey || sheet.getRange('B2').getValue().toString()).trim(),
    appSecret: (secureAppSecret || sheet.getRange('B3').getValue().toString()).trim(),
    account: (secureAccount || sheet.getRange('B4').getValue().toString()).trim(),
    isMock: sheet.getRange('B5').getValue() === 'TRUE',
    baseUrl: sheet.getRange('B5').getValue() === 'TRUE'
      ? 'https://openapivts.koreainvestment.com:29443'
      : 'https://openapi.koreainvestment.com:9443',
    rebalanceTolerance: parseFloat(sheet.getRange('B6').getValue()) || 2.0,
    profitTakingThreshold: parseFloat(sheet.getRange('B7').getValue()) || 40.0,
    targetAmount: parseFloat(sheet.getRange('B8').getValue()) || 300000000,
    targetYield: parseFloat(sheet.getRange('B9').getValue()) || 10.0,
    isISA: sheet.getRange('B9').getValue() === 'TRUE',
    buyFeeRate: (parseFloat(sheet.getRange('B10').getValue()) || 0) / 100,
    sellFeeRate: (parseFloat(sheet.getRange('B11').getValue()) || 0) / 100
  };
}

// 액세스 토큰 관리
function getAccessToken() {
  const config = getConfig();
  const props = PropertiesService.getUserProperties();
  
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  // AppKey별로 고유한 키 생성 (계정별 토큰 분리) + 시트별 격리 추가
  const tokenKey = ssId + '_KIS_TOKEN_' + config.appKey.substring(0, 8);
  const expiryKey = ssId + '_KIS_EXPIRY_' + config.appKey.substring(0, 8);
  
  const token = props.getProperty(tokenKey);
  const expiry = props.getProperty(expiryKey);
  
  // 토큰이 있고 만료되지 않았으면 재사용
  if (token && expiry && new Date().getTime() < parseInt(expiry)) {
    return token;
  }
  
  // 새 토큰 발급
  return issueNewToken();
}

function issueNewToken() {
  const config = getConfig();
  const url = `${config.baseUrl}/oauth2/tokenP`;
  
  const payload = {
    grant_type: 'client_credentials',
    appkey: config.appKey,
    appsecret: config.appSecret
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.access_token) {
      const props = PropertiesService.getUserProperties();
      const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
      const tokenKey = ssId + '_KIS_TOKEN_' + config.appKey.substring(0, 8);
      const expiryKey = ssId + '_KIS_EXPIRY_' + config.appKey.substring(0, 8);
      
      props.setProperty(tokenKey, data.access_token);
      props.setProperty(expiryKey, (new Date().getTime() + (data.expires_in - 60) * 1000).toString());
      return data.access_token;
    } else {
      throw new Error('토큰 발급 실패: ' + JSON.stringify(data));
    }
  } catch (e) {
    throw new Error('토큰 발급 중 오류: ' + e.message);
  }
}

/**
 * 액세스 토큰 강제 만료 및 재발급
 */
function forceRefreshToken() {
  const config = getConfig();
  const props = PropertiesService.getUserProperties();
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const prefix = ssId + '_KIS_TOKEN_' + config.appKey.substring(0, 8);
  const expiryPrefix = ssId + '_KIS_EXPIRY_' + config.appKey.substring(0, 8);
  
  props.deleteProperty(prefix);
  props.deleteProperty(expiryPrefix);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('토큰을 초기화했습니다. 다음 작업 시 새로 발급됩니다.', '✅ 초기화 완료');
}