// ============================================================
// 레이스먼트 가로수길 대시보드 — Apps Script
// STEP 1: Sheets 세팅 + ERP 데이터 파이프라인
// ============================================================

// ── 상수 정의 ────────────────────────────────────────────────

const SHEET_NAMES = {
  RAW:      '판매데이터_RAW',
  CLEAN:    '판매데이터_CLEAN',
  GOAL:     '목표관리',
  ISSUE:    '제품이슈',
  NOREASON: '미구매사유',
  SAFE:     '세이프사이즈',
  SNS:      'SNS인증이벤트',
  NOTE:     '특이사항',
  FEEDBACK: '고객피드백',
  SETTING:  '대시보드설정',
};

// ERP 원본 컬럼명 → CLEAN 컬럼명 매핑
const COLUMN_MAP = {
  '거래명세서일':   '날짜',
  '거래명세서번호': '영수증번호',
  '담당자':        '직원명',
  '브랜드':        '브랜드',
  '카테고리2':     '카테고리',
  '실루엣':        '실루엣',
  '성별':          '성별',
  '품명':          '상품명',
  '품번':          '품번',
  '규격':          '사이즈',
  '수량':          '수량',
  '판매금액':      '매출액',
  '창고':          '창고',
};

const CLEAN_HEADERS = [
  '날짜', '영수증번호', '직원명', '브랜드', '카테고리',
  '실루엣', '성별', '상품명', '품번', '사이즈',
  '수량', '매출액', '창고', '거래유형',
];

const STORE_NAME = '매장 (신사동)'; // 가로수길 매장 필터값

const NOREASON_CODE_MAP = {
  1: '사이즈/재고 부족',
  2: '착화감 불만',
  3: '가격 부담',
  4: '디자인/컬러 불만',
  5: '원하는 상품/브랜드 없음',
  6: '비교 후 고민',
  7: '구매 시점 아님',
  8: '기타/직원 메모 필요',
};

// Google Form 응답 시트 컬럼 헤더
// Form 질문 문구가 바뀌면 여기만 수정하면 됨
// 헤더 뒤 공백은 indexOf 비교 시 trim()으로 처리
const FEEDBACK_HEADERS = {
  TIMESTAMP:      '타임스탬프',
  SATISFACTION:   '오늘 구매하신 제품에 만족하시나요?',
  RECOMMENDATION: '2. 추천받은 제품이 내 러닝 목적과 잘 맞는다고 느끼셨나요?',
  REVISIT:        '3. 오늘 경험 후 다시 방문하고 싶은 마음이 드셨나요?',
};


// ── 1. 초기 세팅: 전체 탭 생성 ──────────────────────────────

/**
 * 실행방법: Apps Script 편집기에서 initSheets() 선택 후 실행
 * 이미 존재하는 탭은 건너뜁니다.
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _createSheetIfNotExists(ss, SHEET_NAMES.RAW);
  _setupCleanSheet(ss);
  _setupGoalSheet(ss);
  _setupIssueSheet(ss);
  _setupNoReasonSheet(ss);
  _setupSafeSheet(ss);
  _setupSnsSheet(ss);
  _setupNoteSheet(ss);
  _setupFeedbackSheet(ss);
  _setupSettingSheet(ss);

  Logger.log('✅ 전체 탭 세팅 완료 — 판매데이터_RAW 탭에 ERP 엑셀을 붙여넣고 processRawData()를 실행하세요.');
}

function _createSheetIfNotExists(ss, name) {
  if (!ss.getSheetByName(name)) {
    ss.insertSheet(name);
  }
}

function _setupCleanSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.CLEAN);
  const sh = ss.getSheetByName(SHEET_NAMES.CLEAN);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, CLEAN_HEADERS.length).setValues([CLEAN_HEADERS]);
    sh.getRange(1, 1, 1, CLEAN_HEADERS.length).setFontWeight('bold');
  }
}

function _setupGoalSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.GOAL);
  const sh = ss.getSheetByName(SHEET_NAMES.GOAL);
  if (sh.getLastRow() === 0) {
    const headers = ['날짜', '일별목표', '주차', '주간목표', '월', '월간목표', '점장메모'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function _setupIssueSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.ISSUE);
  const sh = ss.getSheetByName(SHEET_NAMES.ISSUE);
  if (sh.getLastRow() === 0) {
    const headers = ['기록일', '작성자', '이슈유형', '브랜드', '상품명', '색상', '사이즈', 'ERP재고', '실재고', '이슈내용', '처리상태', '처리일', '점장메모'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function _setupNoReasonSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.NOREASON);
  const sh = ss.getSheetByName(SHEET_NAMES.NOREASON);
  if (sh.getLastRow() === 0) {
    const headers = ['기록일', '작성자', '카테고리', '브랜드', '상품명', '사이즈', '미구매사유코드', '미구매사유명', '메모'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function _setupSafeSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.SAFE);
  const sh = ss.getSheetByName(SHEET_NAMES.SAFE);
  if (sh.getLastRow() === 0) {
    const headers = ['기록일', '작성자', '세이프사이즈진행', '러닝목적', '구매여부', '구매상품', '고객반응메모'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function _setupSnsSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.SNS);
  const sh = ss.getSheetByName(SHEET_NAMES.SNS);
  if (sh.getLastRow() === 0) {
    const headers = ['기록일', '작성자', 'SNS인증참여', '인증채널', '구매여부', '혜택제공여부', '고객반응메모'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function _setupNoteSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.NOTE);
  const sh = ss.getSheetByName(SHEET_NAMES.NOTE);
  if (sh.getLastRow() === 0) {
    const headers = ['기록일', '작성자', '중요도', '내용', '후속조치필요', '점장메모'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function _setupFeedbackSheet(ss) {
  // Google Form 연동 시트는 이미 존재하므로 없을 때만 생성
  _createSheetIfNotExists(ss, SHEET_NAMES.FEEDBACK);
}

function _setupSettingSheet(ss) {
  _createSheetIfNotExists(ss, SHEET_NAMES.SETTING);
  const sh = ss.getSheetByName(SHEET_NAMES.SETTING);
  if (sh.getLastRow() === 0) {
    const headers = [
      '날짜', '매장상태문장',
      '오늘의액션1', '오늘의액션2', '오늘의액션3', '오늘의액션4', '오늘의액션5',
      '점장코멘트_매출', '점장코멘트_상품', '점장코멘트_미구매', '마지막수정시간',
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}


// ── 2. ERP 데이터 파이프라인: RAW → CLEAN ────────────────────

/**
 * 실행방법: 판매데이터_RAW에 ERP 엑셀 붙여넣기 후
 *           Apps Script 편집기에서 processRawData() 실행
 *
 * 처리 순서:
 *   1) RAW 탭에서 헤더 위치 자동 탐지 (1~4행 중 '거래명세서일' 있는 행)
 *   2) 헤더 다음 행(알파벳 코드행) 건너뜀
 *   3) 창고 = '매장 (신사동)' 필터
 *   4) 13개 컬럼 매핑
 *   5) 날짜 YYYY-MM-DD 표준화
 *   6) 수량 < 0 → 거래유형 = '반품', 나머지 = '정상판매'
 *   7) CLEAN 탭에 중복 없이 추가 (영수증번호+품번+사이즈 기준)
 *
 * 실제 ERP 파일 구조:
 *   행0 = 타이틀(판매데이터분석_ajl)
 *   행1 = 헤더(거래명세서일 등)  ← headerRowIdx
 *   행2 = 알파벳 코드행(A,B,C...)  ← 건너뜀
 *   행3 = TOTAL 합계행  ← 날짜 없음으로 자동 제외
 *   행4~ = 실제 데이터
 */
function processRawData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSh = ss.getSheetByName(SHEET_NAMES.RAW);
  const cleanSh = ss.getSheetByName(SHEET_NAMES.CLEAN);

  if (!rawSh) {
    Logger.log('❌ 판매데이터_RAW 탭을 찾을 수 없습니다.');
    return;
  }

  const rawData = rawSh.getDataRange().getValues();
  if (rawData.length < 5) {
    Logger.log('❌ RAW 탭에 데이터가 없거나 너무 적습니다. ERP 엑셀을 붙여넣었는지 확인하세요.');
    return;
  }

  // 1) 헤더 행 탐지 (최대 4행 안에서 '거래명세서일' 포함 행)
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(4, rawData.length); i++) {
    if (rawData[i].includes('거래명세서일')) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx === -1) {
    Logger.log('❌ 헤더 행을 찾을 수 없습니다. ERP 파일에 "거래명세서일" 컬럼이 있는지 확인하세요.');
    return;
  }

  const erpHeaders = rawData[headerRowIdx];

  // 2) ERP 컬럼 인덱스 매핑
  const erpColIdx = {};
  Object.keys(COLUMN_MAP).forEach(erpCol => {
    const idx = erpHeaders.indexOf(erpCol);
    if (idx !== -1) erpColIdx[erpCol] = idx;
  });

  // 필수 컬럼 누락 체크
  const required = ['거래명세서일', '수량', '판매금액', '창고'];
  const missing = required.filter(col => !(col in erpColIdx));
  if (missing.length > 0) {
    Logger.log('❌ 필수 컬럼 누락: ' + missing.join(', ') + ' — ERP 파일 형식을 확인하세요.');
    return;
  }

  // 3) 기존 CLEAN 데이터에서 중복 키 수집 (영수증번호+품번+사이즈)
  const existingKeys = new Set();
  const cleanData = cleanSh.getDataRange().getValues();
  const cleanHeaders = cleanData[0];
  const ciReceipt = cleanHeaders.indexOf('영수증번호');
  const ciPartNo  = cleanHeaders.indexOf('품번');
  const ciSize    = cleanHeaders.indexOf('사이즈');

  for (let i = 1; i < cleanData.length; i++) {
    const key = `${cleanData[i][ciReceipt]}_${cleanData[i][ciPartNo]}_${cleanData[i][ciSize]}`;
    existingKeys.add(key);
  }

  // 4) 데이터 행 처리
  // headerRowIdx+1 행이 알파벳 코드행(A,B,C...)인지 확인 후 건너뜀
  // 코드행이 없는 ERP 파일이 올 경우를 대비한 방어 처리
  // TOTAL 합계행은 날짜가 비어있으므로 아래 빈 행 제외 조건으로 자동 처리됨
  const codeRowIdx = headerRowIdx + 1;
  const isCodeRow = rawData[codeRowIdx] && String(rawData[codeRowIdx][0]).trim().toUpperCase() === 'A';
  const dataStartIdx = isCodeRow ? headerRowIdx + 2 : headerRowIdx + 1;
  const newRows = [];
  let skippedStore = 0;
  let skippedDup   = 0;

  for (let i = dataStartIdx; i < rawData.length; i++) {
    const row = rawData[i];

    // 날짜 없는 행 제외 (TOTAL 합계행 및 빈 행 모두 처리)
    if (!row[erpColIdx['거래명세서일']]) continue;

    // 창고 필터: 신사동만
    const store = String(row[erpColIdx['창고']] || '').trim();
    if (store !== STORE_NAME) {
      skippedStore++;
      continue;
    }

    // 중복 체크 (영수증번호+품번+사이즈)
    const receiptNo = String(row[erpColIdx['거래명세서번호']] || '');
    const partNo    = String(row[erpColIdx['품번']] || '');
    const size      = String(row[erpColIdx['규격']] || '');
    const dupKey    = `${receiptNo}_${partNo}_${size}`;
    if (existingKeys.has(dupKey)) {
      skippedDup++;
      continue;
    }
    existingKeys.add(dupKey);

    // 날짜 표준화
    const dateStr = _formatDate(row[erpColIdx['거래명세서일']]);

    // 수량 → 거래유형
    const qty = Number(row[erpColIdx['수량']] || 0);
    const txType = qty < 0 ? '반품' : '정상판매';

    // CLEAN 행 구성
    const cleanRow = [
      dateStr,                                            // 날짜
      receiptNo,                                          // 영수증번호
      String(row[erpColIdx['담당자']]   || '').trim(),   // 직원명
      String(row[erpColIdx['브랜드']]   || '').trim(),   // 브랜드
      String(row[erpColIdx['카테고리2']]|| '').trim(),   // 카테고리
      String(row[erpColIdx['실루엣']]   || '').trim(),   // 실루엣
      String(row[erpColIdx['성별']]     || '').trim(),   // 성별
      String(row[erpColIdx['품명']]     || '').trim(),   // 상품명
      partNo,                                             // 품번
      size,                                               // 사이즈
      qty,                                                // 수량
      Number(row[erpColIdx['판매금액']] || 0),           // 매출액 (부가세 제외)
      store,                                              // 창고
      txType,                                             // 거래유형
    ];
    newRows.push(cleanRow);
  }

  // 5) CLEAN 탭에 신규 행 추가
  if (newRows.length > 0) {
    const startRow = cleanSh.getLastRow() + 1;
    cleanSh.getRange(startRow, 1, newRows.length, CLEAN_HEADERS.length).setValues(newRows);
  }

  Logger.log('✅ 변환 완료 — 신규 추가: ' + newRows.length + '건 / 중복 건너뜀: ' + skippedDup + '건 / 타 매장 제외: ' + skippedStore + '건');
}

/**
 * 날짜값을 YYYY-MM-DD 문자열로 변환
 */
function _formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const str = String(val).trim();
  // 이미 YYYY-MM-DD 형태
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  // YYYYMMDD 형태
  if (/^\d{8}$/.test(str)) return `${str.slice(0,4)}-${str.slice(4,6)}-${str.slice(6,8)}`;
  return str;
}

/**
 * Google Form 타임스탬프 ("2026. 5. 12 오후 12:03:42" 형식) → YYYY-MM-DD 변환
 * new Date()로 한국어 오전/오후 형식을 파싱할 수 없어 별도 처리
 */
function _parseKoreanTimestamp(val) {
  if (!val) return '';
  // Apps Script에서 이미 Date 객체로 온 경우
  if (val instanceof Date) return _formatDate(val);
  // "2026. 5. 12 오후 12:03:42" 형식
  const str = String(val).trim();
  const match = str.match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (match) {
    const y = match[1];
    const m = String(match[2]).padStart(2, '0');
    const d = String(match[3]).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  // 그 외 형식은 기존 _formatDate로 fallback
  return _formatDate(new Date(val));
}


// ── 3. 집계 헬퍼 함수 (대시보드 API용) ──────────────────────

/**
 * 오늘 날짜 문자열 반환 (YYYY-MM-DD)
 */
function getTodayStr() {
  return _formatDate(new Date());
}

/**
 * 이번 주 시작일(화요일) ~ 종료일(일요일) 반환
 * 월요일 휴무 → 화~일 기준
 */
function getThisWeekRange() {
  const today = new Date();
  const dow = today.getDay(); // 0=일, 1=월, 2=화 ... 6=토

  // 가장 최근 화요일 계산
  const daysFromTue = (dow === 0) ? 5 : (dow === 1) ? 6 : dow - 2;
  const tue = new Date(today);
  tue.setDate(today.getDate() - daysFromTue);

  const sun = new Date(tue);
  sun.setDate(tue.getDate() + 5);

  return {
    start: _formatDate(tue),
    end:   _formatDate(sun),
  };
}

/**
 * 이번 달 범위 반환 (YYYY-MM-DD)
 */
function getThisMonthRange() {
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth();
  const start = _formatDate(new Date(y, m, 1));
  const end   = _formatDate(new Date(y, m + 1, 0));
  return { start, end };
}

/**
 * 주차 코드 반환 (YYYY-W## 형식, 화요일 기준)
 */
function getWeekCode(dateStr) {
  const d = new Date(dateStr);
  const dow = d.getDay();
  const daysFromTue = (dow === 0) ? 5 : (dow === 1) ? 6 : dow - 2;
  const tue = new Date(d);
  tue.setDate(d.getDate() - daysFromTue);

  const startOfYear = new Date(tue.getFullYear(), 0, 1);
  const weekNum = Math.ceil(((tue - startOfYear) / 86400000 + startOfYear.getDay() + 1) / 7);
  return `${tue.getFullYear()}-W${String(weekNum).padStart(2, '0')}`;
}

/**
 * CLEAN 탭에서 날짜 범위 + 거래유형 필터로 행 반환
 */
function getCleanRows(startDate, endDate, txType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.CLEAN);
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const ci = {};
  CLEAN_HEADERS.forEach(h => { ci[h] = headers.indexOf(h); });

  return data.slice(1).filter(row => {
    const d = row[ci['날짜']];
    if (!d) return false;
    if (d < startDate || d > endDate) return false;
    if (txType && row[ci['거래유형']] !== txType) return false;
    return true;
  }).map(row => {
    const obj = {};
    CLEAN_HEADERS.forEach(h => { obj[h] = row[ci[h]]; });
    return obj;
  });
}


// ── 4. 대시보드 데이터 API ───────────────────────────────────

/**
 * 웹앱에서 호출: 홈 화면 전체 데이터를 JSON으로 반환
 */
function getDashboardData(filterType) {
  // filterType: 'today' | 'week'
  const today = getTodayStr();
  const week  = getThisWeekRange();
  const month = getThisMonthRange();

  const startDate = filterType === 'today' ? today : week.start;
  const endDate   = filterType === 'today' ? today : week.end;

  // 판매 데이터 — today / week / month 각각 집계
  const todaySales  = getCleanRows(today,       today,       '정상판매');
  const todayReturn = getCleanRows(today,       today,       '반품');
  const salesRows   = getCleanRows(startDate,   endDate,     '정상판매');
  const returnRows  = getCleanRows(startDate,   endDate,     '반품');
  const salesWeek   = getCleanRows(week.start,  week.end,    '정상판매');
  const returnWeek  = getCleanRows(week.start,  week.end,    '반품');
  const salesMonth  = getCleanRows(month.start, month.end,   '정상판매');
  const returnMonth = getCleanRows(month.start, month.end,   '반품');

  return {
    filter:    filterType,
    today:     today,
    weekRange: week,

    // C01 오늘의 매장 상태
    storeStatus: _getStoreStatus(today),

    // C02 목표 매출 현황
    salesGoal: _getSalesGoal(today, week, month, todaySales, todayReturn, salesWeek, returnWeek, salesMonth, returnMonth),

    // C03 상품 흐름 — today/week 이중 구조로 반환 (토글 시 재로딩 불필요)
    productFlow: _getProductFlow(todaySales, salesWeek),

    // C04 제품 이슈
    issues: _getIssues(),

    // C05 미구매 사유 — today/week 이중 구조로 반환
    noReasons: _getNoReasons(week),

    // C06 고객 피드백
    feedback: _getFeedback(week),

    // C07 세이프사이즈
    safeSize: _getSafeSize(today),

    // C08 SNS 인증
    snsEvent: _getSnsEvent(today),

    // C09 특이사항 + 오늘의 액션
    notes:   _getNotes(today),
    actions: _getActions(today),
  };
}

// C01: 오늘의 매장 상태
function _getStoreStatus(today) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTING);
  if (!sh || sh.getLastRow() < 2) return { text: '' };

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const ciDate = headers.indexOf('날짜');
  const ciText = headers.indexOf('매장상태문장');

  // 오늘 날짜 행 우선, 없으면 가장 최근 행
  let found = null;
  for (let i = data.length - 1; i >= 1; i--) {
    const d = _formatDate(data[i][ciDate]);
    if (d === today) { found = data[i]; break; }
  }
  if (!found && data.length > 1) found = data[data.length - 1];
  return { text: found ? String(found[ciText] || '') : '' };
}

// C02: 목표 매출 현황
function _getSalesGoal(today, week, month, salesRows, returnRows, salesWeek, returnWeek, salesMonth, returnMonth) {
  const goalSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.GOAL);
  const goals  = { daily: 0, weekly: 0, monthly: 0 };

  if (goalSh && goalSh.getLastRow() > 1) {
    const data    = goalSh.getDataRange().getValues();
    const headers = data[0];
    const ciDate  = headers.indexOf('날짜');
    const ciDaily = headers.indexOf('일별목표');
    const ciWeek  = headers.indexOf('주간목표');
    const ciMonth = headers.indexOf('월간목표');
    const ciWCode = headers.indexOf('주차');
    const ciMon   = headers.indexOf('월');

    const thisWeekCode  = getWeekCode(today);
    const thisMonthCode = today.slice(0, 7); // YYYY-MM

    for (let i = 1; i < data.length; i++) {
      const d = _formatDate(data[i][ciDate]);
      if (d === today && data[i][ciDaily]) goals.daily = Number(data[i][ciDaily]);
      if (String(data[i][ciWCode]) === thisWeekCode && data[i][ciWeek]) goals.weekly = Number(data[i][ciWeek]);
      if (String(data[i][ciMon])  === thisMonthCode && data[i][ciMonth]) goals.monthly = Number(data[i][ciMonth]);
    }
  }

  const sumSales = arr => arr.reduce((s, r) => s + Number(r['매출액'] || 0), 0);

  const todaySales   = sumSales(salesRows)  - sumSales(returnRows);
  const weeklySales  = sumSales(salesWeek)  - sumSales(returnWeek);
  const monthlySales = sumSales(salesMonth) - sumSales(returnMonth);

  return {
    daily:   { goal: goals.daily,   actual: todaySales,   rate: _rate(todaySales, goals.daily),   status: _status(_rate(todaySales, goals.daily)) },
    weekly:  { goal: goals.weekly,  actual: weeklySales,  rate: _rate(weeklySales, goals.weekly),  status: _status(_rate(weeklySales, goals.weekly)) },
    monthly: { goal: goals.monthly, actual: monthlySales, rate: _rate(monthlySales, goals.monthly), status: _status(_rate(monthlySales, goals.monthly)) },
  };
}

function _rate(actual, goal) {
  if (!goal || goal === 0) return null;
  return Math.round((actual / goal) * 100);
}

function _status(rate) {
  if (rate === null) return 'unknown';
  if (rate >= 90) return '양호';
  if (rate >= 70) return '주의';
  return '위험';
}

// C03: 상품 흐름
// 프론트 기대 구조: { today: { topByQty, topBySales }, week: { topByQty, topBySales } }
// today/week 모두 계산해서 내려줌 → 프론트 토글 시 재로딩 없이 즉시 전환
function _getProductFlow(todaySalesRows, weekSalesRows) {
  function calcFlow(rows) {
    const qtyMap  = {};
    const salesMap = {};
    rows.forEach(r => {
      const name = r['상품명'] || '(미확인)';
      qtyMap[name]   = (qtyMap[name]   || 0) + Number(r['수량']   || 0);
      salesMap[name] = (salesMap[name] || 0) + Number(r['매출액'] || 0);
    });
    return {
      topByQty:   Object.entries(qtyMap).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([name,qty])=>({name,qty})),
      topBySales: Object.entries(salesMap).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([name,amount])=>({name,amount})),
    };
  }
  return {
    today: calcFlow(todaySalesRows),
    week:  calcFlow(weekSalesRows),
  };
}

// C04: 제품 이슈 (미처리 우선)
// 프론트 기대 구조: { pending: [], pendingTotal: N }
// 이슈유형 한글 → key 변환 (프론트 CSS 클래스용)
const ISSUE_TYPE_KEY_MAP = {
  'ERP재고불일치':    'erp_mismatch',
  'ERP 재고 불일치': 'erp_mismatch',
  '사이즈부족':       'size_shortage',
  '사이즈 부족':      'size_shortage',
  '제품파손오염불량':     'product_damage',
  '제품 파손/오염/불량':  'product_damage',
};

function _getIssues() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ISSUE);
  if (!sh || sh.getLastRow() < 2) return { pending: [], pendingTotal: 0 };

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const ci = {};
  headers.forEach((h, i) => { ci[h] = i; });

  const all = data.slice(1).filter(r => r[ci['기록일']]).map(r => {
    const typeLabel = String(r[ci['이슈유형']] || '').trim();
    return {
      date:      _formatDate(r[ci['기록일']]),
      author:    r[ci['작성자']],
      type:      ISSUE_TYPE_KEY_MAP[typeLabel] || typeLabel, // key로 변환, 없으면 원본
      typeLabel: typeLabel,                                   // 한글 표시용
      brand:     r[ci['브랜드']],
      product:   r[ci['상품명']],
      color:     r[ci['색상']],
      size:      r[ci['사이즈']],
      content:   r[ci['이슈내용']],
      status:    r[ci['처리상태']],
      memo:      r[ci['점장메모']],
    };
  });

  const pending = all.filter(r => r.status !== '처리완료');
  return { pending, pendingTotal: pending.length };
}

// C05: 미구매 사유
// 프론트 기대 구조: { today: { top3, total }, week: { top3, total } }
// saveNoReason()은 payload.reason(한글 문자열)을 받으므로
// 시트에는 미구매사유명(한글) 컬럼을 기준으로 집계
function _getNoReasons(week) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOREASON);
  if (!sh || sh.getLastRow() < 2) {
    return { today: { top3: [], total: 0 }, week: { top3: [], total: 0 } };
  }

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const ci = {};
  headers.forEach((h, i) => { ci[h] = i; });

  const today = getTodayStr();

  const todayMap = {}; let todayTotal = 0;
  const weekMap  = {}; let weekTotal  = 0;

  data.slice(1).forEach(r => {
    const d    = _formatDate(r[ci['기록일']]);
    const name = r[ci['미구매사유명']] || '기타';
    if (!d) return;

    if (d === today) {
      todayMap[name] = (todayMap[name] || 0) + 1;
      todayTotal++;
    }
    if (d >= week.start && d <= week.end) {
      weekMap[name] = (weekMap[name] || 0) + 1;
      weekTotal++;
    }
  });

  const toTop3 = (map) =>
    Object.entries(map)
      .sort((a, b) => b[1] - a[1]).slice(0, 3)
      .map(([reason, count]) => ({ reason, count }));

  return {
    today: { top3: toTop3(todayMap), total: todayTotal },
    week:  { top3: toTop3(weekMap),  total: weekTotal  },
  };
}

// C06: 고객 피드백 평균
function _getFeedback(week) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.FEEDBACK);
  if (!sh || sh.getLastRow() < 2) return { satisfaction: null, recommendation: null, revisit: null, count: 0 };

  const data    = sh.getDataRange().getValues();
  const headers = data[0];

  // trim()으로 헤더 뒤 공백 방어 (Form 응답 시트 헤더에 후행 공백 있을 수 있음)
  const ciTs  = headers.findIndex(h => String(h).trim() === FEEDBACK_HEADERS.TIMESTAMP);
  const ciSat = headers.findIndex(h => String(h).trim() === FEEDBACK_HEADERS.SATISFACTION);
  const ciRec = headers.findIndex(h => String(h).trim() === FEEDBACK_HEADERS.RECOMMENDATION);
  const ciRev = headers.findIndex(h => String(h).trim() === FEEDBACK_HEADERS.REVISIT);

  // 텍스트 응답 → 숫자 변환
  // 만족도는 숫자(1~5)로 오고, 추천 적합도·재방문 의향은 텍스트로 옴
  const SCALE_MAP = {
    '매우 그렇다': 5, '그렇다': 4, '보통이다': 3, '그렇지 않다': 2, '전혀 그렇지 않다': 1,
    '다시 방문 의향 있다': 5, '고민 중이다': 3, '잘 모르겠다': 2, '아마 안 올 것 같다': 1,
  };

  function toScore(val) {
    if (val === null || val === undefined || val === '') return null;
    const num = Number(val);
    if (!isNaN(num) && num > 0) return num;
    return SCALE_MAP[String(val).trim()] || null;
  }

  // 수정: new Date() 대신 _parseKoreanTimestamp() 사용
  const rows = data.slice(1).filter(r => {
    if (!r[ciTs]) return false;
    const d = _parseKoreanTimestamp(r[ciTs]);
    return d >= week.start && d <= week.end;
  });

  if (rows.length === 0) return { satisfaction: null, recommendation: null, revisit: null, count: 0 };

  const avg = (arr, ci) => {
    const vals = arr.map(r => toScore(r[ci])).filter(v => v !== null);
    return vals.length ? Math.round((vals.reduce((s, v) => s + v, 0) / vals.length) * 10) / 10 : null;
  };

  return {
    satisfaction:   avg(rows, ciSat),
    recommendation: avg(rows, ciRec),
    revisit:        avg(rows, ciRev),
    count:          rows.length,
  };
}

// C07: 세이프사이즈 성과
function _getSafeSize(today) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SAFE);
  if (!sh || sh.getLastRow() < 2) return { participants: 0, conversions: 0, rate: null };

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const ci = {};
  headers.forEach((h, i) => { ci[h] = i; });

  const todayRows    = data.slice(1).filter(r => _formatDate(r[ci['기록일']]) === today);
  const participants = todayRows.filter(r => String(r[ci['세이프사이즈진행']]).toUpperCase() === 'Y').length;
  const conversions  = todayRows.filter(r => String(r[ci['세이프사이즈진행']]).toUpperCase() === 'Y' && String(r[ci['구매여부']]).toUpperCase() === 'Y').length;
  const rate = participants > 0 ? Math.round((conversions / participants) * 100) : null;

  return { participants, conversions, rate };
}

// C08: SNS 인증 성과
function _getSnsEvent(today) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SNS);
  if (!sh || sh.getLastRow() < 2) return { participants: 0, conversions: 0, rate: null };

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const ci = {};
  headers.forEach((h, i) => { ci[h] = i; });

  const todayRows    = data.slice(1).filter(r => _formatDate(r[ci['기록일']]) === today);
  const participants = todayRows.filter(r => String(r[ci['SNS인증참여']]).toUpperCase() === 'Y').length;
  const conversions  = todayRows.filter(r => String(r[ci['SNS인증참여']]).toUpperCase() === 'Y' && String(r[ci['구매여부']]).toUpperCase() === 'Y').length;
  const rate = participants > 0 ? Math.round((conversions / participants) * 100) : null;

  return { participants, conversions, rate };
}

// C09: 오늘의 특이사항
function _getNotes(today) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOTE);
  if (!sh || sh.getLastRow() < 2) return [];

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const ci = {};
  headers.forEach((h, i) => { ci[h] = i; });

  return data.slice(1)
    .filter(r => _formatDate(r[ci['기록일']]) === today)
    .map(r => ({
      level:    r[ci['중요도']],
      content:  r[ci['내용']],
      followUp: r[ci['후속조치필요']],
      memo:     r[ci['점장메모']],
    }));
}

// C09: 오늘의 액션
function _getActions(today) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTING);
  if (!sh || sh.getLastRow() < 2) return [];

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const ci = {};
  headers.forEach((h, i) => { ci[h] = i; });

  let found = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (_formatDate(data[i][ci['날짜']]) === today) { found = data[i]; break; }
  }
  if (!found && data.length > 1) found = data[data.length - 1];
  if (!found) return [];

  return [1, 2, 3, 4, 5]
    .map(n => String(found[ci[`오늘의액션${n}`]] || '').trim())
    .filter(v => v !== '');
}


// ── 5. 직원 입력 폼 저장 API ─────────────────────────────────

/**
 * 웹앱 POST 요청 처리: 직원 입력 폼 데이터 저장
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const type   = params.type;
    let result;

    if      (type === 'issue')    result = saveIssue(params);
    else if (type === 'noreason') result = saveNoReason(params);
    else if (type === 'note')     result = saveNote(params);
    else if (type === 'safesize') result = saveSafeSize(params);
    else if (type === 'sns')      result = saveSns(params);
    else if (type === 'setting')  result = saveSetting(params);
    else result = { success: false, error: '알 수 없는 type: ' + type };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveIssue(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ISSUE);
  sh.appendRow([
    p.date || getTodayStr(),
    p.author || '',
    p.issueType || '',
    p.brand || '',
    p.product || '',
    p.color || '',
    p.size || '',
    p.erpStock || '',
    p.realStock || '',
    p.content || '',
    '미확인',   // 처리상태 자동
    '',          // 처리일
    '',          // 점장메모
  ]);
  return { success: true };
}

function saveNoReason(p) {
  // 프론트는 payload.reason 에 한글 사유명을 보냄 (예: '사이즈/재고 부족')
  // reasonCode는 역방향 매핑으로 복원
  const REASON_TO_CODE = Object.fromEntries(
    Object.entries(NOREASON_CODE_MAP).map(([k, v]) => [v, Number(k)])
  );
  const reasonName = p.reason || '';
  const code = p.reasonCode ? Number(p.reasonCode) : (REASON_TO_CODE[reasonName] || 0);
  const name = reasonName || NOREASON_CODE_MAP[code] || '';

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOREASON);
  sh.appendRow([
    p.date     || getTodayStr(),
    p.author   || '',
    p.category || '',
    p.brand    || '',
    p.product  || '',
    p.size     || '',
    code,
    name,
    p.memo     || '',
  ]);
  return { success: true };
}

function saveNote(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOTE);
  sh.appendRow([
    p.date || getTodayStr(),
    p.author || '',
    p.level || '보통',
    p.content || '',
    p.followUp || 'N',
    '',
  ]);
  return { success: true };
}

function saveSafeSize(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SAFE);
  sh.appendRow([
    p.date || getTodayStr(),
    p.author || '',
    p.proceeded || 'N',
    p.runningPurpose || '',
    p.purchased || 'N',
    p.purchasedProduct || '',
    p.memo || '',
  ]);
  return { success: true };
}

function saveSns(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SNS);
  sh.appendRow([
    p.date || getTodayStr(),
    p.author || '',
    p.participated || 'N',
    p.channel || '',
    p.purchased || 'N',
    p.benefitGiven || 'N',
    p.memo || '',
  ]);
  return { success: true };
}

function saveSetting(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTING);
  const today = getTodayStr();

  // 오늘 날짜 행이 있으면 업데이트, 없으면 추가
  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const ciDate  = headers.indexOf('날짜');
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (_formatDate(data[i][ciDate]) === today) { targetRow = i + 1; break; }
  }

  const row = [
    today,
    p.storeStatus || '',
    p.action1 || '', p.action2 || '', p.action3 || '', p.action4 || '', p.action5 || '',
    p.commentSales || '',
    p.commentProduct || '',
    p.commentNoReason || '',
    new Date().toISOString(),
  ];

  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return { success: true };
}


// ── 6. 직원 입력 폼 저장 프록시 (웹앱 내부 호출용) ──────────
//
// index_v2.html의 saveData()가 google.script.run.doPost_proxy(payload)를 호출함.
// doPost()는 HTTP POST 전용이라 웹앱 내부에서 직접 호출 불가.
// 이 함수가 내부 라우터 역할을 한다.

function doPost_proxy(payload) {
  const type = payload.type;
  if      (type === 'issue')    return saveIssue(payload);
  else if (type === 'noreason') return saveNoReason(payload);
  else if (type === 'note')     return saveNote(payload);
  else if (type === 'safesize') return saveSafeSize(payload);
  else if (type === 'sns')      return saveSns(payload);
  else if (type === 'setting')  return saveSetting(payload);
  return { success: false, error: '알 수 없는 type: ' + type };
}


// ── 7. 웹앱 진입점 ──────────────────────────────────────────

function doGet(e) {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('레이스먼트 가로수길 대시보드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTML 파일 include 헬퍼 (CSS/JS 분리용)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
