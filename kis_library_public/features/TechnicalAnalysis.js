/**
 * 기술적 지표 분석 엔진
 * RSI, MACD, 볼린저밴드, 스토캐스틱, 거래량 기반 Confluence Score 계산
 * Score: -1.0 (강한 매도) ~ +1.0 (강한 매수)
 */

// ─────────────────────────────────────────────────────────────
// 공개 API
// ─────────────────────────────────────────────────────────────

/**
 * 종목코드로 Confluence Score 반환 (CacheService 캐시 적용)
 * @param {string} stockCode
 * @returns {{ score: number, signals: object, summary: string }}
 */
function getConfluenceScore(stockCode) {
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  const cacheKey = 'TA_' + stockCode + '_' + today;
  const cache = CacheService.getScriptCache();

  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      Logger.log('[TA] ' + stockCode + ' 캐시 히트: score=' + parsed.score);
      return parsed;
    } catch (e) { /* 캐시 파싱 실패 시 재계산 */ }
  }

  const result = _computeScore(stockCode);
  cache.put(cacheKey, JSON.stringify(result), 21600); // 6시간 캐시
  Logger.log('[TA] ' + stockCode + ' score=' + result.score + ' | ' + result.summary);
  return result;
}

/**
 * FSD 모드 × Score → 매수 배율
 * @param {string} fsdMode
 * @param {number} score
 * @returns {number} multiplier
 */
function getBuyMultiplier(fsdMode, score) {
  const table = {
    'Chill':     [0.8, 0.6, 0.4, 0.0, 0.0],
    'Standard':  [1.0, 0.8, 0.6, 0.3, 0.0],
    'Hurry':     [1.2, 1.0, 0.8, 0.5, 0.3],
    'Assertive': [1.5, 1.2, 1.0, 0.7, 0.5],
    'Mad Max':   [2.0, 1.5, 1.2, 1.0, 0.8]
  };
  const row = table[fsdMode] || table['Standard'];
  if (score >= 0.6)  return row[0];
  if (score >= 0.3)  return row[1];
  if (score >= -0.3) return row[2];
  if (score >= -0.6) return row[3];
  return row[4];
}

/**
 * 매도 tolerance 축소 배율 (score < -0.3일 때 FSD 모드별 적용)
 * @param {string} fsdMode
 * @param {number} score
 * @returns {number} 0.0~1.0 (기존 tolerance에 곱함)
 */
function getSellToleranceMultiplier(fsdMode, score) {
  if (score >= -0.3) return 1.0; // 중립/긍정 신호면 tolerance 유지
  const multipliers = {
    'Chill':     1.0,
    'Standard':  0.8,
    'Hurry':     0.5,
    'Assertive': 0.5,
    'Mad Max':   0.3
  };
  return multipliers[fsdMode] || 0.8;
}

/**
 * 수익실현 시 TA score에 따른 매도 비율 (excess 대비)
 * score >= 0      → 0    (보유, 매도 안함 — 상승 중)
 * 0 > score≥-0.3  → 0.2  (약한 하락 신호 — excess 20% 매도)
 * -0.3>score≥-0.6 → 0.4  (중간 하락 — excess 40% 매도)
 * -0.6>score≥-0.8 → 0.7  (강한 하락 — excess 70% 매도)
 * score < -0.8    → 1.0  (매우 강한 하락 — excess 전량 매도)
 * FSD 공격성 모드는 추가 가산
 * Mad Max 는 score>=-0.1 도 트리거 (조기 차익실현)
 * @returns {number} 0.0~1.0
 */
function getProfitTakingSellRatio(fsdMode, score) {
  // Mad Max 모드: 약하게 꺾여도 차익실현
  if (fsdMode === 'Mad Max' && score < 0.1) {
    score = Math.min(score, -0.1); // 강제로 음수 영역 매핑
  }
  if (score >= 0) return 0;

  let baseRatio;
  if (score < -0.8)      baseRatio = 1.0;
  else if (score < -0.6) baseRatio = 0.7;
  else if (score < -0.3) baseRatio = 0.4;
  else                   baseRatio = 0.2;

  const fsdMult = {
    'Chill':     1.0,
    'Standard':  1.0,
    'Hurry':     1.2,
    'Assertive': 1.2,
    'Mad Max':   1.5
  };
  return Math.min(1.0, baseRatio * (fsdMult[fsdMode] || 1.0));
}

/**
 * 수익 재배분 매도 시 목표 대비 매도 비율
 * score < -0.5 AND profitRate >= 15% 트리거
 * @param {string} fsdMode
 * @returns {number} 0.70~0.95 (target ratio 대비 얼마까지 매도)
 */
function getRedistributeRatio(fsdMode) {
  const ratios = {
    'Chill':     0.95,
    'Standard':  0.90,
    'Hurry':     0.85,
    'Assertive': 0.80,
    'Mad Max':   0.70
  };
  return ratios[fsdMode] || 0.90;
}

// ─────────────────────────────────────────────────────────────
// 내부 계산
// ─────────────────────────────────────────────────────────────

function _computeScore(stockCode) {
  const NEUTRAL = { score: 0, signals: {}, summary: '데이터 부족 (중립)' };

  const history = _getPriceHistory(stockCode);
  if (!history || history.length < 20) return NEUTRAL;

  const closes  = history.map(d => d.close);
  const highs   = history.map(d => d.high);
  const lows    = history.map(d => d.low);
  const volumes = history.map(d => d.volume);

  const items = [];
  const signals = {};

  // RSI (25%)
  const rsiVal = _rsi(closes, 14);
  if (rsiVal !== null) {
    const s = rsiVal < 30 ? 1.0 : rsiVal < 45 ? 0.5 : rsiVal < 55 ? 0.0 : rsiVal < 70 ? -0.5 : -1.0;
    signals.rsi = { value: +rsiVal.toFixed(1), signal: s };
    items.push({ signal: s, weight: 0.25 });
  }

  // MACD (25%)
  const macdRes = _macd(closes, 12, 26, 9);
  if (macdRes !== null) {
    let s = macdRes.histogram > 0 ? 1.0 : -1.0;
    if (macdRes.prevHistogram !== null) {
      if (macdRes.prevHistogram < 0 && macdRes.histogram > 0) s = Math.min(1.0, s + 0.3);
      if (macdRes.prevHistogram > 0 && macdRes.histogram < 0) s = Math.max(-1.0, s - 0.3);
    }
    signals.macd = {
      macd: +macdRes.macd.toFixed(0),
      signalLine: +macdRes.signal.toFixed(0),
      histogram: +macdRes.histogram.toFixed(0),
      signal: s
    };
    items.push({ signal: s, weight: 0.25 });
  }

  // 볼린저밴드 (20%)
  const bbRes = _bollingerBands(closes, 20, 2);
  if (bbRes !== null) {
    const { upper, middle, lower, price } = bbRes;
    const midUpper = (middle + upper) / 2;
    const s = price <= lower ? 1.0 : price <= middle ? 0.5 : price <= midUpper ? -0.5 : -1.0;
    signals.bb = { upper: +upper.toFixed(0), middle: +middle.toFixed(0), lower: +lower.toFixed(0), price, signal: s };
    items.push({ signal: s, weight: 0.20 });
  }

  // 스토캐스틱 (20%)
  const stochRes = _stochastic(highs, lows, closes, 14, 3);
  if (stochRes !== null) {
    const { k } = stochRes;
    const s = k < 20 ? 1.0 : k < 40 ? 0.3 : k < 60 ? 0.0 : k < 80 ? -0.3 : -1.0;
    signals.stoch = { k: +k.toFixed(1), d: +stochRes.d.toFixed(1), signal: s };
    items.push({ signal: s, weight: 0.20 });
  }

  // 거래량 (10%)
  const volSignal = _volumeSignal(closes, volumes, 20);
  const avgVol = volumes.slice(-20).reduce((a, b) => a + b, 0) / Math.min(20, volumes.length);
  const volRatio = avgVol > 0 ? volumes[volumes.length - 1] / avgVol : 1;
  signals.volume = { ratio: +volRatio.toFixed(2), signal: volSignal };
  items.push({ signal: volSignal, weight: 0.10 });

  const totalWeight = items.reduce((s, i) => s + i.weight, 0);
  const weightedSum = items.reduce((s, i) => s + i.signal * i.weight, 0);
  const score = totalWeight > 0
    ? parseFloat(Math.max(-1.0, Math.min(1.0, weightedSum / totalWeight)).toFixed(2))
    : 0;

  return { score, signals, summary: _buildSummary(signals, score) };
}

function _getPriceHistory(stockCode) {
  try {
    const data = callKISAPI(
      '/uapi/domestic-stock/v1/quotations/inquire-daily-price',
      'FHKST01010400',
      {
        'FID_COND_MRKT_DIV_CODE': 'J',
        'FID_INPUT_ISCD': stockCode,
        'FID_INPUT_DATE_1': '',
        'FID_PERIOD_DIV_CODE': 'D',
        'FID_ORG_ADJ_PRC': '0'
      }
    );

    if (!data.output || !Array.isArray(data.output) || data.output.length === 0) {
      Logger.log('[TA] ' + stockCode + ' 일봉 데이터 없음');
      return null;
    }

    // KIS API는 최신순 → 역순 정렬 (오래된 것부터)
    return [...data.output].reverse().map(d => ({
      date:   d.stck_bsop_date,
      close:  parseInt(d.stck_clpr  || 0),
      open:   parseInt(d.stck_oprc  || 0),
      high:   parseInt(d.stck_hgpr  || 0),
      low:    parseInt(d.stck_lwpr  || 0),
      volume: parseInt(d.acml_vol   || 0)
    })).filter(d => d.close > 0);

  } catch (e) {
    Logger.log('[TA] ' + stockCode + ' 일봉 조회 실패: ' + e.toString());
    return null;
  }
}

// RSI (Wilder's smoothing)
function _rsi(closes, period) {
  if (closes.length < period + 1) return null;

  let gains = 0, losses = 0;
  for (let i = 1; i <= period; i++) {
    const d = closes[i] - closes[i - 1];
    if (d > 0) gains += d; else losses -= d;
  }
  let avgGain = gains / period;
  let avgLoss = losses / period;

  for (let i = period + 1; i < closes.length; i++) {
    const d = closes[i] - closes[i - 1];
    avgGain = (avgGain * (period - 1) + Math.max(0, d))  / period;
    avgLoss = (avgLoss * (period - 1) + Math.max(0, -d)) / period;
  }

  if (avgLoss === 0) return 100;
  return 100 - 100 / (1 + avgGain / avgLoss);
}

// MACD (12, 26, 9)
function _macd(closes, fast, slow, signal) {
  if (closes.length < slow) return null;

  const kFast   = 2 / (fast   + 1);
  const kSlow   = 2 / (slow   + 1);
  const kSignal = 2 / (signal + 1);

  // EMA fast: SMA 초기값 후 EMA 적용
  let emaFast = closes.slice(0, fast).reduce((a, b) => a + b, 0) / fast;
  for (let i = fast; i < slow; i++) emaFast = closes[i] * kFast + emaFast * (1 - kFast);

  // EMA slow: SMA 초기값
  let emaSlow = closes.slice(0, slow).reduce((a, b) => a + b, 0) / slow;

  // MACD 시리즈 (index slow-1 부터)
  const macdSeries = [emaFast - emaSlow];
  for (let i = slow; i < closes.length; i++) {
    emaFast = closes[i] * kFast + emaFast * (1 - kFast);
    emaSlow = closes[i] * kSlow + emaSlow * (1 - kSlow);
    macdSeries.push(emaFast - emaSlow);
  }

  // Signal line (9-period EMA of MACD series)
  const initLen = Math.min(signal, macdSeries.length);
  let signalLine = macdSeries.slice(0, initLen).reduce((a, b) => a + b, 0) / initLen;
  for (let i = initLen; i < macdSeries.length; i++) {
    signalLine = macdSeries[i] * kSignal + signalLine * (1 - kSignal);
  }

  const current  = macdSeries[macdSeries.length - 1];
  const histogram = current - signalLine;

  // 직전 히스토그램 (방향 전환 감지용)
  let prevHistogram = null;
  if (macdSeries.length >= 2) {
    prevHistogram = macdSeries[macdSeries.length - 2] - signalLine;
  }

  return { macd: current, signal: signalLine, histogram, prevHistogram };
}

// 볼린저밴드 (20일 SMA ± 2σ)
function _bollingerBands(closes, period, stdMult) {
  if (closes.length < period) return null;
  const recent = closes.slice(-period);
  const sma = recent.reduce((a, b) => a + b, 0) / period;
  const variance = recent.reduce((s, v) => s + Math.pow(v - sma, 2), 0) / period;
  const std = Math.sqrt(variance);
  return {
    upper:  sma + stdMult * std,
    middle: sma,
    lower:  sma - stdMult * std,
    price:  closes[closes.length - 1]
  };
}

// 스토캐스틱 %K(14), %D(3) 스무딩 3
function _stochastic(highs, lows, closes, kPeriod, dPeriod) {
  if (closes.length < kPeriod + dPeriod * 2) return null;

  // Raw %K
  const rawK = [];
  for (let i = kPeriod - 1; i < closes.length; i++) {
    const hh = Math.max(...highs.slice(i - kPeriod + 1, i + 1));
    const ll = Math.min(...lows.slice(i  - kPeriod + 1, i + 1));
    rawK.push(hh === ll ? 50 : ((closes[i] - ll) / (hh - ll)) * 100);
  }

  // %K smoothed (3-period SMA)
  const kSmoothed = [];
  for (let i = dPeriod - 1; i < rawK.length; i++) {
    kSmoothed.push(rawK.slice(i - dPeriod + 1, i + 1).reduce((a, b) => a + b, 0) / dPeriod);
  }
  if (kSmoothed.length === 0) return null;

  // %D (3-period SMA of %K smoothed)
  const dLine = [];
  for (let i = dPeriod - 1; i < kSmoothed.length; i++) {
    dLine.push(kSmoothed.slice(i - dPeriod + 1, i + 1).reduce((a, b) => a + b, 0) / dPeriod);
  }

  return {
    k: kSmoothed[kSmoothed.length - 1],
    d: dLine.length > 0 ? dLine[dLine.length - 1] : kSmoothed[kSmoothed.length - 1]
  };
}

// 거래량 신호 (20일 평균 대비 현재 거래량 × 가격 방향)
function _volumeSignal(closes, volumes, period) {
  if (volumes.length < 2) return 0;
  const avgVol = volumes.slice(-period).reduce((a, b) => a + b, 0) / Math.min(period, volumes.length);
  if (avgVol === 0) return 0;

  const ratio = volumes[volumes.length - 1] / avgVol;
  const priceDir = closes[closes.length - 1] > closes[closes.length - 2] ? 1 : -1;

  if (ratio > 1.5) return priceDir;
  if (ratio > 1.2) return priceDir * 0.5;
  return 0;
}

/**
 * Score → 판정 라벨 (시트 표시용)
 * 강한 매수/매수 우호/중립/매도 우호/강한 매도
 */
function getScoreVerdict(score) {
  if (score >= 0.6)  return '🟢🟢 강한 매수';
  if (score >= 0.3)  return '🟢 매수 우호';
  if (score >= -0.3) return '⚪ 중립';
  if (score >= -0.6) return '🔴 매도 우호';
  return '🔴🔴 강한 매도';
}

/**
 * 요약 텍스트 — 물리량 비유로
 *   💪 힘(RSI)   ➡️ 방향(MACD)   📏 범위(볼린저)
 *   ⚡ 단기(스토캐스틱)   🏋️ 무게(거래량)
 */
function _buildSummary(signals, score) {
  const parts = [];

  // 💪 힘 (RSI) — 가격 모멘텀의 세기
  if (signals.rsi) {
    const v = signals.rsi.value;
    let lbl;
    if (v < 30)      lbl = '과매도반등여지';
    else if (v < 45) lbl = '약함';
    else if (v < 55) lbl = '중립';
    else if (v < 70) lbl = '강함';
    else             lbl = '과열';
    parts.push('💪힘:' + lbl + '(' + v + ')');
  }

  // ➡️ 방향 (MACD) — 추세 방향과 전환
  if (signals.macd) {
    let lbl;
    if (signals.macd.signal >= 1.0)      lbl = '↑골든크로스';
    else if (signals.macd.signal <= -1.0) lbl = '↓데드크로스';
    else if (signals.macd.histogram > 0) lbl = '↑상승';
    else                                 lbl = '↓하락';
    parts.push('➡️방향:' + lbl);
  }

  // 📏 범위 (볼린저밴드) — 변동폭 위치
  if (signals.bb) {
    let lbl;
    if (signals.bb.signal >= 1.0)       lbl = '하단이탈(저점반등여지)';
    else if (signals.bb.signal >= 0.5)  lbl = '중간↓';
    else if (signals.bb.signal >= -0.5) lbl = '중간↑';
    else                                lbl = '상단이탈(고점위험)';
    parts.push('📏범위:' + lbl);
  }

  // ⚡ 단기 (스토캐스틱) — 단기 과열/과매도
  if (signals.stoch) {
    const k = signals.stoch.k;
    let lbl;
    if (k < 20)      lbl = '과매도(K ' + k + ')';
    else if (k < 40) lbl = '저점권';
    else if (k < 60) lbl = '중간';
    else if (k < 80) lbl = '고점권';
    else             lbl = '과매수(K ' + k + ')';
    parts.push('⚡단기:' + lbl);
  }

  // 🏋️ 무게 (거래량 비율) — 신호 신뢰도
  if (signals.volume) {
    const r = signals.volume.ratio;
    parts.push('🏋️거래량:' + r + 'x');
  }

  return parts.join(' | ');
}
