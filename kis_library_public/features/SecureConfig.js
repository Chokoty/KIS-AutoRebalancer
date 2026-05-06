/**
 * SecureConfig.js
 * UserProperties를 사용하여 민감한 API 키와 계좌 정보를 안전하게 관리합니다.
 */

/**
 * 보안 설정 다이얼로그 열기
 */
function openSecureConfigDialog() {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const props = PropertiesService.getUserProperties();
  const config = {
    appKey: props.getProperty(ssId + '_KIS_APP_KEY') || '',
    appSecret: props.getProperty(ssId + '_KIS_APP_SECRET') || '',
    account: props.getProperty(ssId + '_KIS_ACCOUNT') || '',
    geminiApiKey: props.getProperty(ssId + '_GEMINI_API_KEY') || '',
    geminiModelId: props.getProperty(ssId + '_GEMINI_MODEL_ID') || 'gemini-2.0-flash'
  };

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
      <style>
        body { padding: 20px; font-family: 'Noto Sans KR', sans-serif; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input[type="text"], input[type="password"] { width: 100%; box-sizing: border-box; padding: 8px; }
        .footer { margin-top: 20px; text-align: right; }
        .help-text { font-size: 11px; color: #666; margin-top: 4px; }
        .warning { color: #d93025; font-size: 12px; margin-top: 10px; border: 1px solid #d93025; padding: 8px; border-radius: 4px; background: #fce8e6; }
      </style>
    </head>
    <body>
      <div class="warning">
        이곳에 입력된 정보는 <b>본인의 개인 계정(UserProperties)</b>에만 암호화되어 저장됩니다. 
        시트를 다른 사람과 공유해도 이 정보는 공유되지 않습니다.
      </div>
      <br>
      <div class="form-group">
        <label>한국투자증권 APP KEY</label>
        <input type="text" id="appKey" value="${config.appKey}">
      </div>
      <div class="form-group">
        <label>한국투자증권 APP SECRET</label>
        <input type="password" id="appSecret" value="${config.appSecret}">
      </div>
      <div class="form-group">
        <label>계좌번호 (예: 12345678-01)</label>
        <input type="text" id="account" value="${config.account}">
      </div>
      <div class="form-group">
        <label>Gemini API Key</label>
        <input type="password" id="geminiApiKey" value="${config.geminiApiKey}">
      </div>

      <div class="form-group">
        <label>사용할 Gemini 모델</label>
        <select id="geminiModelId" style="width: 100%; padding: 8px;">
          <option value="gemini-2.5-flash" ${config.geminiModelId === 'gemini-2.5-flash' ? 'selected' : ''}>Gemini 2.5 Flash (최신/권장)</option>
          <option value="gemini-2.0-flash" ${config.geminiModelId === 'gemini-2.0-flash' ? 'selected' : ''}>Gemini 2.0 Flash (안정/대안)</option>
        </select>
        <p class="help-text">최신 모델인 2.5 Flash 사용을 권장합니다.</p>
      </div>
      
      <div class="footer">
        <button class="action" onclick="save()">저장하기</button>
        <button onclick="google.script.host.close()">취소</button>
      </div>

      <script>
        function save() {
          const data = {
            appKey: document.getElementById('appKey').value,
            appSecret: document.getElementById('appSecret').value,
            account: document.getElementById('account').value,
            geminiApiKey: document.getElementById('geminiApiKey').value,
            geminiModelId: document.getElementById('geminiModelId').value
          };
          
          google.script.run
            .withSuccessHandler(() => {
              alert('보안 설정이 안전하게 저장되었습니다.');
              google.script.host.close();
            })
            .saveSecureConfig(data);
        }
      </script>
    </body>
    </html>
  `;

  const output = HtmlService.createHtmlOutput(html)
    .setHeight(480)
    .setWidth(450)
    .setTitle('🛡️ API 키 보안 설정 (개인용)');
    
  SpreadsheetApp.getUi().showModalDialog(output, '🛡️ API 키 보안 설정');
}

/**
 * 보안 데이터 저장 (서버 측)
 */
function saveSecureConfig(data) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const props = PropertiesService.getUserProperties();
  
  const updates = {};
  updates[ssId + '_KIS_APP_KEY'] = data.appKey.trim();
  updates[ssId + '_KIS_APP_SECRET'] = data.appSecret.trim();
  updates[ssId + '_KIS_ACCOUNT'] = data.account.trim();
  updates[ssId + '_GEMINI_API_KEY'] = data.geminiApiKey.trim();
  updates[ssId + '_GEMINI_MODEL_ID'] = data.geminiModelId;
  
  props.setProperties(updates);
  
  // 시트에 남아있는 민감 정보 삭제 (선택 사항)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚙️ 설정');
  if (sheet) {
    sheet.getRange('B2:B4').setValue('🛡️ 보안 저장됨 (개인 설정)');
    sheet.getRange('B10').setValue('🛡️ 보안 저장됨 (개인 설정)');
  }
}
