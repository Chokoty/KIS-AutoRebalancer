# KIS Auto-Rebalancer

**한국투자증권(KIS) Open API**를 활용하여 구글 스프레드시트에서 주식 포트폴리오를 관리하고 자동 리밸런싱을 수행하는 **Google Apps Script(GAS)** 프로젝트입니다.

## 🚀 주요 기능

- **📊 통합 대시보드**: 실시간 자산 모니터링 (총 평가액, 예수금, 수익률). 목표 포트폴리오 대비 현재 비중 비교 및 매수/매도 신호 자동 계산.
- **⚡ 리밸런싱 실행**: 대시보드에서 계산된 매수/매도 주문을 버튼 한 번으로 일괄 실행. 매도 우선 순위 전략(현금 확보 후 매수)으로 안전하게 진행.
- **🛣️ 주간 자동 리밸런싱**: 매주 월요일 오전 10시, 포트폴리오 비중을 자동으로 점검하고 리밸런싱 실행.
- **💰 수익 실현**: 목표 수익률 기반 안전 인출 금액 계산. 포트폴리오 비율을 유지하며 수익을 실현하는 자동 매도 계획 생성.
- **🛡️ 보안 강화**: API Key 등 민감한 정보를 `UserProperties`에 안전하게 저장. 시트별 설정 격리로 여러 계좌 관리 가능.
- **📝 거래 내역 기록**: 실행된 모든 주문이 `📝 거래내역` 시트에 자동 저장.
- **🔢 수수료/세금 반영**: 매수/매도 제비용을 계산에 포함하여 정확한 가용 현금 산출. ISA 계좌 여부에 따른 수수료율 별도 설정 지원.

## 📂 파일 구조

```
kis_auto_rebalance/
├── container/
│   └── code.gs          # 구글 시트에 직접 붙여넣는 컨테이너 스크립트
└── kis_library_public/  # GAS 라이브러리 (Script ID로 추가)
    ├── core/
    │   ├── KISClient.js     # 한국투자증권 REST API 통신 모듈
    │   ├── config.js        # 환경 설정 및 토큰 관리
    │   └── code.js          # 포트폴리오 조회, 자동 리밸런싱 트리거
    ├── features/
    │   ├── Dashboard.js     # 리밸런싱 대시보드 UI 및 주문 실행
    │   ├── Withdraw.js      # 수익 실현 및 인출 계획
    │   └── SecureConfig.js  # API 키 보안 저장 관리
    └── utils/
        ├── Menu.js          # 구글 시트 상단 메뉴 구성
        ├── SheetManager.js  # 시트 초기화 및 렌더링 유틸리티
        └── UsageGuide.js    # 사용법 안내
```

## ⚙️ 설치 방법

### 1. 라이브러리 추가

구글 스프레드시트를 새로 만들고, `확장 프로그램 > Apps Script > 라이브러리(+)`에서 아래 Script ID를 입력해 라이브러리를 추가합니다.

```
Script ID: 1LXA06wO7XtQmqqZ4GdnFm6w4bwzl8nrG5dhcE2qc6h0WFcxtxj-OFoc6
```

- 버전: **HEAD (개발 모드)** 선택
- 식별자: `KIS` 로 설정

### 2. 컨테이너 스크립트 추가

Apps Script 편집기에서 `container/code.gs` 내용을 붙여넣고 저장합니다.

### 3. clasp으로 직접 배포하는 경우 (선택)

라이브러리 방식 대신 코드를 직접 시트에 배포할 수도 있습니다.

```bash
npm install -g @google/clasp
clasp login

git clone https://github.com/Chokoty/kis-auto-rebalance.git
cd kis-auto-rebalance/kis_library_public
clasp push -f
```

### 4. 초기 설정 실행

1. 구글 시트를 새로고침하면 상단에 `📊 KIS AutoTrader` 메뉴가 생깁니다
2. `⚙️ 설정 및 관리 > ⚙️ 초기 설정` 실행 (필요한 시트 자동 생성)
3. `🛡️ API 키 보안 설정`에서 한국투자증권 AppKey, AppSecret, 계좌번호 입력
4. 정보는 구글 서버의 `UserProperties`에 암호화되어 저장됩니다

### 5. 포트폴리오 설정

`📋 포트폴리오설정` 시트에서 종목코드, 종목명, 목표비율(%)을 수정합니다. 기본 예시 포트폴리오가 제공됩니다.

## 🛣️ 자동 리밸런싱 설정

메뉴에서 `🛣️ 주간 자동 리밸런싱 (차선 유지) 활성/비활성`을 클릭하면 **매주 월요일 오전 10시**에 자동으로 리밸런싱이 실행됩니다.

- 13일 이내 실행 이력이 있으면 중복 실행 방지
- 임계치(기본 2%) 이내 비중 차이는 주문 생략

## 🔧 기술 스택

- **플랫폼**: Google Apps Script (V8 런타임)
- **인터페이스**: Google Sheets (커스텀 메뉴, HTML 다이얼로그)
- **API**: 한국투자증권 Open API
- **배포**: clasp

## ⚠️ 주의사항

- **실전 투자**: 자동 리밸런싱은 **실제 돈으로 자동 매매**가 일어납니다. 반드시 모의투자로 충분히 테스트하고 본인의 책임하에 사용하세요.
- **모의투자 설정**: `⚙️ 설정` 시트의 `모의투자` 항목을 `TRUE`로 설정하면 실제 주문 없이 테스트할 수 있습니다.
- **KIS Open API**: [한국투자증권 오픈 API](https://apiportal.koreainvestment.com)에서 발급받은 AppKey/AppSecret이 필요합니다.

## 📋 라이선스

MIT License
