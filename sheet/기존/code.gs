/** 설정값 */
const SPREADSHEET_ID = '1UcrwW-6cC75IqGZL4rCQp8V55qH85HFABTMcfxPl7p4';
const RAW_SHEET_NAME = 'RAW';
const ORIGINAL_PHOTO_FOLDER_ID = '1hg-6ITs3Jj8auDStTnfuXCNwLefM8hGz';
const TZ = 'Asia/Seoul';
const DATE_FMT = 'yyyy-MM-dd';
/** RAW 시트 캐시: 키, TTL(초). 참조 시트 읽기 속도 개선용 */
const RAW_SHEET_CACHE_KEY = 'raw_sheet_data';
const RAW_SHEET_CACHE_TTL = 90;

/** 공통 유틸 */
function _ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function _raw() { return _ss().getSheetByName(RAW_SHEET_NAME); }
function _today() { return Utilities.formatDate(new Date(), TZ, DATE_FMT); }

/** joinDate를 yyyy-MM-dd 문자열로 변환. Date 객체, ISO UTC 문자열, 일반 문자열 모두 처리. */
function _toDateStr(joinDate) {
  if (!joinDate) return null;
  if (joinDate instanceof Date) {
    return Utilities.formatDate(joinDate, TZ, DATE_FMT);
  }
  if (typeof joinDate === 'string') {
    // ISO UTC 문자열 ("2026-03-08T15:00:00.000Z") → KST 기준 날짜로 변환
    if (joinDate.indexOf('T') > -1) {
      var d = new Date(joinDate);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, TZ, DATE_FMT);
      }
    }
    // 이미 yyyy-MM-dd 형식이면 그대로
    if (/^\d{4}-\d{2}-\d{2}/.test(joinDate)) {
      return joinDate.slice(0, 10);
    }
    return joinDate;
  }
  return String(joinDate);
}

/** RAW 시트 데이터 캐시 조회/저장. A:G만 캐싱(촬영 기록 H·I는 행별로 시트에서 읽음). 반환: { sheet, lastRow, values } */
function getRawSheetData() {
  const sheet = _raw();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { sheet: sheet, lastRow: lastRow, values: [] };
  }
  const cache = CacheService.getScriptCache();
  const cacheKey = RAW_SHEET_CACHE_KEY + '_' + lastRow;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      Logger.log('[getRawSheetData] 캐시 사용, 행 수: ' + (parsed.values ? parsed.values.length : 0));
      return { sheet: sheet, lastRow: lastRow, values: parsed.values };
    } catch (e) {
      Logger.log('[getRawSheetData] 캐시 파싱 실패: ' + e.toString());
    }
  }
  const values = sheet.getRange(2, 1, lastRow, 7).getValues(); // A:G
  try {
    cache.put(cacheKey, JSON.stringify({ values: values }), RAW_SHEET_CACHE_TTL);
    Logger.log('[getRawSheetData] 시트 읽어 캐시 저장, 행 수: ' + values.length);
  } catch (e) {
    Logger.log('[getRawSheetData] 캐시 저장 생략: ' + e.toString());
  }
  return { sheet: sheet, lastRow: lastRow, values: values };
}

/** RAW 시트 캐시 무효화 (시트 수정 후 호출) */
function invalidateRawSheetCache() {
  const sheet = _raw();
  const lastRow = sheet.getLastRow();
  const cache = CacheService.getScriptCache();
  cache.remove(RAW_SHEET_CACHE_KEY + '_' + lastRow);
  Logger.log('[invalidateRawSheetCache] 캐시 무효화');
}

/** 웹앱 진입점 */
function doGet(e) {
  // URL 파라미터로 페이지 구분
  const page = e?.parameter?.page || 'shooter';
  
  if (page === 'registration') {
    return HtmlService.createTemplateFromFile('Registration')
      .evaluate()
      .setTitle('사원증 촬영 등록')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // 기본: 촬영자 웹앱
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('우아한형제들 사원증 촬영 2.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 초기 데이터 */
function getInitData() {
  const today = _today();
  const email = Session.getActiveUser().getEmail();
  const userId = email ? email.split('@')[0] : '';
  const userIdDisplay = userId + '님';
  
  // 사용자별, 날짜별 캐시 확인 및 생성
  const cacheKey = getCacheKey(userId, today);
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  
  let todayUploads = [];
  
  if (cached) {
    // 캐시가 있으면 불러오기
    todayUploads = getTodayUploadsFromCache(cacheKey);
    
    // 성능 개선: 초기 로딩 시 드라이브 검색 완전 스킵
    // 링크가 있으면 그대로 사용, 없으면 나중에 필요할 때 검색
    // (초기 로딩 속도를 최우선으로 함)
    Logger.log('[getInitData] 캐시에서 ' + todayUploads.length + '개 항목 로드 완료 (드라이브 검색 스킵)');
  } else {
    // 캐시가 없었다면 시트에서 오늘 촬영 기록 확인
    Logger.log('[getInitData] 캐시 없음, 시트에서 오늘 촬영 기록 확인');
    const sheetUploads = getTodayUploadsFromSheet(today, userId);
    if (sheetUploads && sheetUploads.length > 0) {
      Logger.log('[getInitData] 시트에서 ' + sheetUploads.length + '개 항목 발견');
      todayUploads = sheetUploads;
      // 캐시에 저장
      const now = new Date();
      const midnight = new Date(now);
      midnight.setHours(24, 0, 0, 0);
      const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
      const cacheExpiration = Math.min(secondsUntilMidnight, 21600); // 최대 6시간
      cache.put(cacheKey, JSON.stringify(sheetUploads), cacheExpiration);
      Logger.log('[getInitData] 시트 데이터를 캐시에 저장 완료');
    } else {
      // 시트에도 없으면 빈 캐시 공간 생성
      initializeCache(cacheKey);
    }
  }
  
  return { 
    todayKST: today, 
    shooterCount: todayUploads.length, 
    userId: userIdDisplay,
    todayUploads: todayUploads
  };
}

/** 캐시 키 생성 (사용자별, 날짜별) */
function getCacheKey(userId, today) {
  return `uploads_${userId}_${today}`;
}

/** 빈 캐시 공간 초기화 */
function initializeCache(cacheKey) {
  const cache = CacheService.getScriptCache();
  const now = new Date();
  const midnight = new Date(now);
  midnight.setHours(24, 0, 0, 0);
  const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
  const cacheExpiration = Math.min(secondsUntilMidnight, 21600); // 최대 6시간
  
  cache.put(cacheKey, JSON.stringify([]), cacheExpiration);
}

/** 캐시 초기화 (오늘 날짜 캐시 삭제) */
function clearTodayCache() {
  try {
    const today = _today();
    const email = Session.getActiveUser().getEmail();
    const userId = email ? email.split('@')[0] : '';
    const cacheKey = getCacheKey(userId, today);
    const cache = CacheService.getScriptCache();
    
    Logger.log('[clearTodayCache] 캐시 삭제 시작: ' + cacheKey);
    cache.remove(cacheKey);
    Logger.log('[clearTodayCache] 캐시 삭제 완료');
    
    return { ok: true, message: '캐시가 초기화되었습니다.' };
  } catch (error) {
    Logger.log('[clearTodayCache] 오류: ' + error.toString());
    return { ok: false, message: '캐시 초기화 실패: ' + error.toString() };
  }
}

/** 캐시에서 개별 항목 삭제 */
function removeFromCache(empNo) {
  try {
    Logger.log('[removeFromCache] 시작 - 사번: ' + empNo);
    const today = _today();
    const email = Session.getActiveUser().getEmail();
    const userId = email ? email.split('@')[0] : '';
    const cacheKey = getCacheKey(userId, today);
    const cache = CacheService.getScriptCache();
    
    // 기존 캐시 가져오기
    const cached = cache.get(cacheKey);
    if (!cached) {
      Logger.log('[removeFromCache] 캐시 없음');
      return { ok: true, message: '캐시에 항목이 없습니다.' };
    }
    
    let uploads = [];
    try {
      uploads = JSON.parse(cached);
    } catch (e) {
      Logger.log('[removeFromCache] 캐시 파싱 실패: ' + e.toString());
      return { ok: false, message: '캐시 파싱 실패' };
    }
    
    // 해당 사번 제거
    const beforeLength = uploads.length;
    uploads = uploads.filter(u => u.empNo !== empNo);
    const afterLength = uploads.length;
    
    Logger.log('[removeFromCache] 삭제 전: ' + beforeLength + '개, 삭제 후: ' + afterLength + '개');
    
    // 캐시 저장 (자정까지 유효, 최대 6시간)
    const now = new Date();
    const midnight = new Date(now);
    midnight.setHours(24, 0, 0, 0);
    const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
    const cacheExpiration = Math.min(secondsUntilMidnight, 21600); // 최대 6시간
    
    cache.put(cacheKey, JSON.stringify(uploads), cacheExpiration);
    Logger.log('[removeFromCache] 캐시 업데이트 완료');
    
    return { ok: true, message: '삭제되었습니다.' };
  } catch (error) {
    Logger.log('[removeFromCache] 예외 발생: ' + error.toString());
    return { ok: false, message: '삭제 실패: ' + error.toString() };
  }
}

/** 캐시에서 오늘 날짜에 업로드된 정보들 가져오기 */
function getTodayUploadsFromCache(cacheKey) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  
  if (cached) {
    try {
      const uploads = JSON.parse(cached);
      // 순서대로 정렬 (order 필드 기준)
      uploads.sort((a, b) => (a.order || 0) - (b.order || 0));
      return uploads;
    } catch (e) {
      // 캐시 파싱 실패 시 빈 배열 반환
      return [];
    }
  }
  
  return [];
}

/** 캐시에 저장 (사번, 사우명 입력 시) */
function saveToCache(empNo, name, nameEn, joinDate) {
  try {
    Logger.log('[saveToCache 시작] 사번: ' + empNo);
    const today = _today();
    const email = Session.getActiveUser().getEmail();
    const userId = email ? email.split('@')[0] : '';
    const cacheKey = getCacheKey(userId, today);
    Logger.log('[saveToCache] 캐시 키: ' + cacheKey);
    const cache = CacheService.getScriptCache();
    
    // 기존 캐시 가져오기
    const cached = cache.get(cacheKey);
    let uploads = [];
    if (cached) {
      try {
        uploads = JSON.parse(cached);
        Logger.log('[saveToCache] 기존 캐시 불러옴: ' + uploads.length + '개 항목');
      } catch (e) {
        Logger.log('[saveToCache] 캐시 파싱 실패: ' + e.toString());
        uploads = [];
      }
    } else {
      Logger.log('[saveToCache] 기존 캐시 없음');
    }
    
    // 중복 체크 (같은 사번이 이미 있으면 업데이트)
    let existingIndex = -1;
    for (let i = 0; i < uploads.length; i++) {
      if (uploads[i].empNo === empNo) {
        existingIndex = i;
        Logger.log('[saveToCache] 기존 항목 발견: 인덱스 ' + i);
        break;
      }
    }
    
    const now = new Date();
    const timestamp = now.getTime();
    const order = uploads.length; // 현재 배열 길이를 순서로 사용
    
    const joinDateStr = _toDateStr(joinDate);
    
    const uploadData = {
      empNo: empNo,
      name: name,
      nameEn: nameEn || '',
      joinDate: joinDateStr,
      link: '', // 업로드 전이므로 빈 문자열
      order: existingIndex >= 0 ? uploads[existingIndex].order : order,
      timestamp: existingIndex >= 0 ? uploads[existingIndex].timestamp : timestamp
    };
    
    if (existingIndex >= 0) {
      uploads[existingIndex] = uploadData;
      Logger.log('[saveToCache] 기존 항목 업데이트');
    } else {
      uploads.push(uploadData);
      Logger.log('[saveToCache] 새 항목 추가');
    }
    
    // 캐시 저장 (자정까지 유효, 최대 6시간)
    const midnight = new Date(now);
    midnight.setHours(24, 0, 0, 0);
    const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
    const cacheExpiration = Math.min(secondsUntilMidnight, 21600); // 최대 6시간
    
    cache.put(cacheKey, JSON.stringify(uploads), cacheExpiration);
    Logger.log('[saveToCache] 캐시 저장 완료: ' + uploads.length + '개 항목');
  } catch (error) {
    Logger.log('[saveToCache] 예외 발생: ' + error.toString());
    throw error;
  }
}

/** 캐시에 업로드 정보 추가 */
function addUploadInfoToCache(empNo, link) {
  const today = _today();
  const email = Session.getActiveUser().getEmail();
  const userId = email ? email.split('@')[0] : '';
  const cacheKey = getCacheKey(userId, today);
  const cache = CacheService.getScriptCache();
  
  // 기존 캐시 가져오기
  const cached = cache.get(cacheKey);
  let uploads = [];
  if (cached) {
    try {
      uploads = JSON.parse(cached);
    } catch (e) {
      uploads = [];
    }
  }
  
  // 해당 사번 찾아서 업로드 정보 추가
  for (let i = 0; i < uploads.length; i++) {
    if (uploads[i].empNo === empNo) {
      uploads[i].link = link;
      break;
    }
  }
  
  // 캐시 저장 (자정까지 유효, 최대 6시간)
  const now = new Date();
  const midnight = new Date(now);
  midnight.setHours(24, 0, 0, 0);
  const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
  const cacheExpiration = Math.min(secondsUntilMidnight, 21600); // 최대 6시간
  
  cache.put(cacheKey, JSON.stringify(uploads), cacheExpiration);
}

/** 시트에서 오늘 날짜에 촬영 기록이 있는 항목들 가져오기 */
function getTodayUploadsFromSheet(today, userId) {
  try {
    Logger.log('[getTodayUploadsFromSheet] 시작 - 날짜: ' + today + ', 사용자: ' + userId);
    const data = getRawSheetData();
    const { sheet, lastRow, values } = data;
    if (lastRow < 2 || !values.length) {
      Logger.log('[getTodayUploadsFromSheet] 데이터 없음');
      return [];
    }
    
    const newEntry = `${today} ${userId}`;
    const uploads = [];
    
    Logger.log('[getTodayUploadsFromSheet] 검색 중... 총 ' + values.length + '개 행');
    
    for (let i = 0; i < values.length; i++) {
      const rowIndex = i + 2;
      const empNo = String(values[i][0] || '').trim();
      const name = String(values[i][1] || '').trim();
      const nameEn = String(values[i][2] || '').trim();
      const joinDate = values[i][4]; // E열
      
      // 사번과 이름이 없으면 스킵
      if (!empNo || !name) {
        continue;
      }
      
      // H열(8열) 촬영 기록은 캐시하지 않으므로 시트에서 읽기
      const logValue = String(sheet.getRange(rowIndex, 8).getValue() || '').trim();
      
      // 오늘 날짜와 사용자 ID가 포함되어 있는지 확인
      if (logValue && logValue.includes(newEntry)) {
        Logger.log('[getTodayUploadsFromSheet] 오늘 촬영 기록 발견: ' + empNo + ' / ' + name);
        
        // 항상 드라이브에서 최신 파일 찾기 (I열 링크는 참고용)
        Logger.log('[getTodayUploadsFromSheet] 드라이브에서 최신 파일 검색: ' + empNo);
        let previewLink = findLatestFileInDrive(empNo, name);
        if (!previewLink) {
          // 드라이브에서 찾지 못한 경우에만 I열 링크 사용
          const linkCell = sheet.getRange(rowIndex, 9);
          const rich = linkCell.getRichTextValue();
          if (rich) {
            const url = rich.getLinkUrl();
            if (url) {
              previewLink = url.replace(/\/view(\?.*)?$/, '/preview');
            }
          }
        }
        
        const joinDateStr = _toDateStr(joinDate);
        
        uploads.push({
          empNo: empNo,
          name: name,
          nameEn: nameEn || '',
          joinDate: joinDateStr,
          link: previewLink || '',
          order: uploads.length, // 순서는 발견 순서대로
          timestamp: new Date().getTime() // 타임스탬프는 현재 시간으로 설정
        });
      }
    }
    
    Logger.log('[getTodayUploadsFromSheet] 완료 - ' + uploads.length + '개 항목 발견');
    return uploads;
  } catch (error) {
    Logger.log('[getTodayUploadsFromSheet] 오류: ' + error.toString());
    return [];
  }
}

/** 드라이브에서 이름과 사번으로 최신 파일 찾기 */
function findLatestFileInDrive(empNo, name) {
  try {
    Logger.log('[findLatestFileInDrive] 시작 - 사번: ' + empNo + ', 이름: ' + name);
    if (!ORIGINAL_PHOTO_FOLDER_ID) {
      Logger.log('[findLatestFileInDrive] 폴더 ID 미설정');
      return null;
    }
    
    const folder = DriveApp.getFolderById(ORIGINAL_PHOTO_FOLDER_ID);
    // 파일명 패턴: ${name}_${empNo}로 시작하는 파일들 찾기
    // 형식: 사우명_사번_영문명(선택)_입사일yymmdd(선택)_버전번호(선택).확장자
    // 버전 번호는 _1, _2, _3 등 숫자만 가능
    const searchPattern = `${name}_${empNo}`;
    const files = folder.getFiles();
    
    let latestFile = null;
    let latestDate = new Date(0);
    let matchedCount = 0;
    
    // 정규식 패턴: 사우명_사번으로 시작하고, 그 다음에 언더스코어로 구분된 부분들이 올 수 있음
    // 허용되는 형식:
    // - 홍길동_12345678.jpg
    // - 홍길동_12345678_HongGildong.jpg
    // - 홍길동_12345678_110425.jpg
    // - 홍길동_12345678_HongGildong_110425.jpg
    // - 홍길동_12345678_HongGildong_110425_1.jpg (버전 번호)
    // - 홍길동_12345678_HongGildong_110425_2.jpg (버전 번호)
    // 패턴 설명: ^사우명_사번(_[^_]+)*(_\d+)?\.확장자$
    // - ^ : 시작
    // - 사우명_사번 : 정확히 일치
    // - (_[^_]+)* : 언더스코어로 시작하는 부분이 0개 이상 (영문명, 입사일 등)
    // - (_\d+)? : 언더스코어 + 숫자 (버전 번호, 선택적)
    // - \.확장자 : 점과 확장자로 끝남
    const basePattern = new RegExp('^' + escapeRegex(searchPattern) + '(_[^_]+)*(_\\d+)?\\.(jpg|jpeg|png|gif|webp|heic)$', 'i');
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // 파일명이 정확한 형식인지 확인
      // 1. ${name}_${empNo}로 시작 (startsWith로 빠른 필터링)
      // 2. 정규식으로 형식 검증 (영문명, 입사일, 중복번호 등이 올바른 형식인지)
      // 3. 확장자가 허용된 형식인지
      if (fileName.startsWith(searchPattern)) {
        if (basePattern.test(fileName)) {
          matchedCount++;
          const fileDate = file.getLastUpdated();
          if (fileDate > latestDate) {
            latestDate = fileDate;
            latestFile = file;
            Logger.log('[findLatestFileInDrive] 매칭된 파일: ' + fileName + ' (수정일: ' + fileDate + ')');
          }
        } else {
          Logger.log('[findLatestFileInDrive] 형식 불일치 (스킵): ' + fileName);
        }
      }
    }
    
    Logger.log('[findLatestFileInDrive] 총 ' + matchedCount + '개 파일 매칭됨');
    
    if (latestFile) {
      const previewLink = `https://drive.google.com/file/d/${latestFile.getId()}/preview`;
      Logger.log('[findLatestFileInDrive] 최신 파일 찾음: ' + latestFile.getName());
      return previewLink;
    }
    
    Logger.log('[findLatestFileInDrive] 파일을 찾을 수 없음');
    return null;
  } catch (error) {
    Logger.log('[findLatestFileInDrive] 오류: ' + error.toString());
    return null;
  }
}

/** 정규식 특수문자 이스케이프 */
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** 사용자 이미지 찾기 (백그라운드 검색용) */
function findImageForUser(empNo, name) {
  try {
    Logger.log('[findImageForUser] 시작 - 사번: ' + empNo + ', 이름: ' + name);
    
    // 1. I열 링크 먼저 확인 (캐시 사용)
    const data = getRawSheetData();
    const { sheet, values } = data;
    if (values && values.length > 0) {
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === String(empNo).trim() &&
            String(values[i][1]).trim() === String(name).trim()) {
          const rowIndex = i + 2;
          const cell = sheet.getRange(rowIndex, 9);
          const rich = cell.getRichTextValue();
          if (rich) {
            const url = rich.getLinkUrl();
            if (url) {
              const previewLink = url.replace(/\/view(\?.*)?$/, '/preview');
              Logger.log('[findImageForUser] I열 링크 찾음: ' + previewLink);
              return { link: previewLink };
            }
          }
          break;
        }
      }
    }
    
    // 2. I열 링크가 없으면 드라이브에서 검색
    const driveLink = findLatestFileInDrive(empNo, name);
    if (driveLink) {
      Logger.log('[findImageForUser] 드라이브에서 찾음: ' + driveLink);
      return { link: driveLink };
    }
    
    Logger.log('[findImageForUser] 이미지 없음');
    return { link: null };
  } catch (error) {
    Logger.log('[findImageForUser] 예외 발생: ' + error.toString());
    return { link: null };
  }
}

/** 사번/성명 검증 + I열 URL 반환 (드라이브에서도 찾기) */
function verifyUser(empNo, name) {
  try {
    Logger.log('[verifyUser 시작] 사번: ' + empNo + ', 이름: ' + name);
    const data = getRawSheetData();
    const { sheet, lastRow, values } = data;
    Logger.log('[verifyUser] lastRow: ' + lastRow);
    
    if (lastRow < 2 || !values.length) {
      Logger.log('[verifyUser] 데이터 없음');
      return { ok: false, message: '데이터가 없습니다.' };
    }
    
    Logger.log('[verifyUser] 검색 중... 총 ' + values.length + '개 행');
    
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(empNo).trim() &&
          String(values[i][1]).trim() === String(name).trim()) {
        Logger.log('[verifyUser] 일치하는 사용자 찾음: 행 ' + (i + 2));
        
        const rowIndex = i + 2;
        
        // H열(8열) 촬영 기록은 캐시 없음, 시트에서 읽기
        const today = _today();
        const email = Session.getActiveUser().getEmail();
        const userId = email ? email.split('@')[0] : '';
        const newEntry = `${today} ${userId}`;
        const logValue = String(sheet.getRange(rowIndex, 8).getValue() || '').trim();
        const hasTodayRecord = logValue && logValue.includes(newEntry);
        
        Logger.log('[verifyUser] H열 촬영 기록 확인: ' + (hasTodayRecord ? '있음' : '없음'));
        
        // 성능 개선: I열 링크를 먼저 확인 (빠름), 없으면 드라이브 검색 (느림)
        let previewLink = null;
        
        // 1. I열 링크 먼저 확인 (빠름)
        const cell = sheet.getRange(rowIndex, 9);
        const rich = cell.getRichTextValue();
        if (rich) {
          const url = rich.getLinkUrl();
          if (url) {
            previewLink = url.replace(/\/view(\?.*)?$/, '/preview');
            Logger.log('[verifyUser] I열 링크 사용: ' + previewLink);
          }
        }
        
        // 2. I열 링크가 없으면 드라이브에서 검색 (느리지만 정확)
        if (!previewLink) {
          Logger.log('[verifyUser] I열 링크 없음, 드라이브에서 최신 파일 검색 시작');
          previewLink = findLatestFileInDrive(empNo, name);
          if (previewLink) {
            Logger.log('[verifyUser] 드라이브에서 최신 파일 찾음: ' + previewLink);
          } else {
            Logger.log('[verifyUser] 드라이브에서도 파일을 찾을 수 없음');
          }
        }
        
        // C열: 영문명, E열: 입사일
        const nameEn = String(values[i][2] || '').trim();
        const joinDate = values[i][4]; // Date 객체 또는 문자열 (2011-04-25 형태)
        
        Logger.log('[verifyUser] 사용자 정보 - 영문명: ' + nameEn + ', 입사일: ' + joinDate);
        
        // 오늘 촬영 기록이 있으면 캐시에 링크 포함하여 저장, 없으면 링크 없이 저장
        try {
          Logger.log('[verifyUser] 캐시 저장 시작');
          if (hasTodayRecord && previewLink) {
            // 오늘 촬영 기록이 있고 링크가 있으면 캐시에 업로드 정보로 저장
            const today = _today();
            const cacheKey = getCacheKey(userId, today);
            const cache = CacheService.getScriptCache();
            const cached = cache.get(cacheKey);
            let uploads = [];
            if (cached) {
              try {
                uploads = JSON.parse(cached);
              } catch (e) {
                uploads = [];
              }
            }
            
            // 중복 체크
            let existingIndex = -1;
            for (let j = 0; j < uploads.length; j++) {
              if (uploads[j].empNo === empNo) {
                existingIndex = j;
                break;
              }
            }
            
            const joinDateStr = _toDateStr(joinDate);
            
            const uploadData = {
              empNo: empNo,
              name: name,
              nameEn: nameEn || '',
              joinDate: joinDateStr,
              link: previewLink,
              order: existingIndex >= 0 ? uploads[existingIndex].order : uploads.length,
              timestamp: existingIndex >= 0 ? uploads[existingIndex].timestamp : new Date().getTime()
            };
            
            if (existingIndex >= 0) {
              uploads[existingIndex] = uploadData;
            } else {
              uploads.push(uploadData);
            }
            
            // 캐시 저장
            const now = new Date();
            const midnight = new Date(now);
            midnight.setHours(24, 0, 0, 0);
            const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
            const cacheExpiration = Math.min(secondsUntilMidnight, 21600);
            cache.put(cacheKey, JSON.stringify(uploads), cacheExpiration);
            Logger.log('[verifyUser] 오늘 촬영 기록이 있어 캐시에 링크 포함하여 저장 완료');
          } else {
            // 오늘 촬영 기록이 없으면 기존대로 링크 없이 저장
            saveToCache(empNo, name, nameEn, joinDate);
            Logger.log('[verifyUser] 캐시 저장 완료 (링크 없이)');
          }
        } catch (cacheError) {
          Logger.log('[verifyUser] 캐시 저장 실패: ' + cacheError.toString());
        }
        
        const joinDateStr = _toDateStr(joinDate);
        
        const result = { ok: true, empNo, name, nameEn, joinDate: joinDateStr, link: previewLink };
        try {
          Logger.log('[verifyUser] 성공 - 반환값: ' + JSON.stringify(result));
        } catch (logError) {
          Logger.log('[verifyUser] 성공 - 로그 기록 실패: ' + logError.toString());
        }
        Logger.log('[verifyUser] 반환 직전');
        return result;
      }
    }
    Logger.log('[verifyUser] 일치하는 사용자를 찾을 수 없음');
    return { ok: false, message: '사번/성명을 찾을 수 없음' };
  } catch (error) {
    Logger.log('[verifyUser] 예외 발생: ' + error.toString());
    Logger.log('[verifyUser] 스택 트레이스: ' + (error.stack || '없음'));
    try {
      return { ok: false, message: '오류가 발생했습니다: ' + error.toString() };
    } catch (returnError) {
      Logger.log('[verifyUser] 반환값 생성 실패: ' + returnError.toString());
      // 최소한의 반환값이라도 보장
      return { ok: false, message: '오류가 발생했습니다.' };
    }
  }
}

/** 이미지 업로드 + H열/I열 기록 */
function uploadImageToPrinterFolder(base64, filename, empNo, name, nameEn, joinDate) {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4 * 60 * 1000; // 4분 (GAS 웹앱은 6분 제한이지만 안전 마진)
  
  try {
    Logger.log('[uploadImageToPrinterFolder] 시작 - 사번: ' + empNo + ', 이름: ' + name);
    
    // Base64 데이터 크기 체크 (더 엄격한 제한: 약 10MB)
    const base64Size = base64 ? base64.length : 0;
    const maxBase64Size = 10 * 1024 * 1024; // 약 10MB (더 보수적으로)
    if (base64Size > maxBase64Size) {
      Logger.log('[uploadImageToPrinterFolder] 파일 크기 초과: ' + (base64Size / 1024 / 1024).toFixed(2) + 'MB');
      return { ok: false, message: '파일 크기가 너무 큽니다. (최대 10MB)\n이미지를 압축하거나 크기를 줄여주세요.' };
    }
    
    if (!ORIGINAL_PHOTO_FOLDER_ID) {
      Logger.log('[uploadImageToPrinterFolder] ORIGINAL_PHOTO_FOLDER_ID 미설정');
      return { ok: false, message: '폴더 ID가 설정되지 않았습니다.' };
    }

    Logger.log('[uploadImageToPrinterFolder] Base64 크기: ' + (base64Size / 1024 / 1024).toFixed(2) + 'MB');
    
    // 실행 시간 체크
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      Logger.log('[uploadImageToPrinterFolder] 실행 시간 초과');
      return { ok: false, message: '처리 시간이 초과되었습니다. 다시 시도해주세요.' };
    }
    
    const folder = DriveApp.getFolderById(ORIGINAL_PHOTO_FOLDER_ID);
    const dot = filename.lastIndexOf('.');
    const ext = dot > -1 ? filename.slice(dot).toLowerCase() : '.jpg';

    // 입사일을 yymmdd 형식으로 변환 (한국 시간 기준)
    let joinDateStr = '';
    if (joinDate) {
      let joinDateFormatted = '';
      if (joinDate instanceof Date) {
        joinDateFormatted = Utilities.formatDate(joinDate, TZ, 'yyyy-MM-dd');
      } else if (typeof joinDate === 'string') {
        joinDateFormatted = joinDate;
      } else {
        joinDateFormatted = String(joinDate);
      }
      
      const dateMatch = joinDateFormatted.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (dateMatch) {
        const year = parseInt(dateMatch[1]) % 100;
        const month = dateMatch[2];
        const day = dateMatch[3];
        joinDateStr = String(year) + month + day;
      }
    }

    // 영문명 처리 (띄어쓰기 포함)
    const nameEnClean = nameEn ? nameEn.trim() : '';
    
    // 파일명 생성: 사우명_사번_영문명_입사일yymmdd.확장자
    let baseName = `${name}_${empNo}`;
    if (nameEnClean) {
      baseName += `_${nameEnClean}`;
    }
    if (joinDateStr) {
      baseName += `_${joinDateStr}`;
    }
    
    Logger.log('[uploadImageToPrinterFolder] 파일명 생성: ' + baseName + ext);
    
    let targetName = _dedupeNameUnderscore(folder, baseName + ext);

    const contentType = _guessContentType(ext);
    
    // 실행 시간 체크
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      Logger.log('[uploadImageToPrinterFolder] 실행 시간 초과 (디코딩 전)');
      return { ok: false, message: '처리 시간이 초과되었습니다. 다시 시도해주세요.' };
    }
    
    // Base64 디코딩 (메모리 효율적 처리)
    let data;
    try {
      const base64Data = _stripDataUrlPrefix(base64);
      Logger.log('[uploadImageToPrinterFolder] Base64 디코딩 시작, 데이터 길이: ' + base64Data.length);
      
      // Base64 디코딩 (GAS의 base64Decode는 내부적으로 최적화되어 있음)
      data = Utilities.base64Decode(base64Data);
      
      // 디코딩된 데이터 크기 확인
      const decodedSizeMB = data.length / 1024 / 1024;
      Logger.log('[uploadImageToPrinterFolder] Base64 디코딩 완료, 디코딩된 크기: ' + decodedSizeMB.toFixed(2) + 'MB');
      
      // 디코딩된 데이터가 10MB를 초과하면 에러 (더 보수적인 제한)
      if (decodedSizeMB > 10) {
        Logger.log('[uploadImageToPrinterFolder] 디코딩된 파일 크기 초과: ' + decodedSizeMB.toFixed(2) + 'MB');
        return { ok: false, message: '파일 크기가 너무 큽니다. (최대 10MB)\n이미지를 압축하거나 크기를 줄여주세요.' };
      }
    } catch (decodeError) {
      Logger.log('[uploadImageToPrinterFolder] Base64 디코딩 실패: ' + decodeError.toString());
      Logger.log('[uploadImageToPrinterFolder] 디코딩 에러 스택: ' + (decodeError.stack || '없음'));
      return { ok: false, message: '이미지 디코딩에 실패했습니다. 파일 형식을 확인해주세요.' };
    }
    
    // 실행 시간 체크
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      Logger.log('[uploadImageToPrinterFolder] 실행 시간 초과 (업로드 전)');
      return { ok: false, message: '처리 시간이 초과되었습니다. 다시 시도해주세요.' };
    }
    
    // Blob 생성 및 파일 업로드
    let file;
    try {
      const blob = Utilities.newBlob(data, contentType, targetName);
      Logger.log('[uploadImageToPrinterFolder] Blob 생성 완료, 파일 업로드 시작');
      
      // 메모리 해제를 위해 data 변수 null 처리 (가비지 컬렉션 힌트)
      data = null;
      
      file = folder.createFile(blob);
      Logger.log('[uploadImageToPrinterFolder] 파일 업로드 완료: ' + file.getName() + ', ID: ' + file.getId());
    } catch (uploadError) {
      Logger.log('[uploadImageToPrinterFolder] 파일 업로드 실패: ' + uploadError.toString());
      Logger.log('[uploadImageToPrinterFolder] 업로드 에러 스택: ' + (uploadError.stack || '없음'));
      return { ok: false, message: '파일 업로드에 실패했습니다. 파일 크기를 확인하거나 다시 시도해주세요.' };
    }

    // 실행 시간 체크
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      Logger.log('[uploadImageToPrinterFolder] 실행 시간 초과 (시트 업데이트 전)');
      // 파일은 업로드되었으므로 캐시만 업데이트하고 반환
      const previewLink = `https://drive.google.com/file/d/${file.getId()}/preview`;
      addUploadInfoToCache(empNo, previewLink);
      return { 
        ok: true, 
        fileId: file.getId(), 
        fileName: file.getName(), 
        webViewLink: previewLink,
        warning: '파일은 업로드되었지만 시트 업데이트는 시간 초과로 건너뛰었습니다.' 
      };
    }
    
    // 시트 업데이트 (A:G 캐시로 행 찾기)
    const sheetData = getRawSheetData();
    const { sheet, values } = sheetData;
    let rowIndex = -1;
    if (values && values.length > 0) {
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === String(empNo).trim() &&
            String(values[i][1]).trim() === String(name).trim()) {
          rowIndex = i + 2;
          break;
        }
      }
    }

    const previewLink = `https://drive.google.com/file/d/${file.getId()}/preview`;
    
    if (rowIndex > -1) {
      Logger.log('[uploadImageToPrinterFolder] 시트에서 행 찾음: ' + rowIndex);
      const today = _today();
      const email = Session.getActiveUser().getEmail();
      const userId = email ? email.split('@')[0] : '';
      const newEntry = `${today} ${userId}`;

      // H열(8열) 촬영 기록 업데이트 (중복 방지)
      const logCell = sheet.getRange(rowIndex, 8);
      const cur = String(logCell.getValue() || '').trim();
      const tokens = cur ? cur.split(',').map(s => s.trim()) : [];
      const hasToday = tokens.includes(newEntry);

      if (!hasToday) {
        const updated = cur ? newEntry + ', ' + cur : newEntry;
        logCell.setValue(updated);
        Logger.log('[uploadImageToPrinterFolder] H열 촬영 기록 업데이트 완료');
      }

      // I열(9열) 최신 파일명+하이퍼링크 (view 링크 기록)
      const linkCell = sheet.getRange(rowIndex, 9);
      const rich = SpreadsheetApp.newRichTextValue()
        .setText(file.getName())
        .setLinkUrl(`https://drive.google.com/file/d/${file.getId()}/view?usp=sharing`)
        .build();
      linkCell.setRichTextValue(rich);
      Logger.log('[uploadImageToPrinterFolder] I열 URL 업데이트 완료');
      
      // 캐시에 업로드 정보 추가
      addUploadInfoToCache(empNo, previewLink);
      Logger.log('[uploadImageToPrinterFolder] 캐시 업데이트 완료');
    } else {
      Logger.log('[uploadImageToPrinterFolder] 시트에서 행을 찾지 못함, 캐시만 업데이트');
      // 시트에서 찾지 못했어도 캐시에는 추가 (파일은 업로드됨)
      addUploadInfoToCache(empNo, previewLink);
    }

    // preview 주소 반환 (iframe 표시용)
    Logger.log('[uploadImageToPrinterFolder] 성공 - previewLink: ' + previewLink);
    return {
      ok: true,
      fileId: file.getId(),
      fileName: file.getName(),
      webViewLink: previewLink
    };
  } catch (error) {
    Logger.log('[uploadImageToPrinterFolder] 예외 발생: ' + error.toString());
    Logger.log('[uploadImageToPrinterFolder] 스택 트레이스: ' + (error.stack || '없음'));
    return {
      ok: false,
      message: '업로드 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/** 파일명 중복 시 -1, -2… 부여 (최적화: 타임아웃 방지, 입사일 제외하고 비교) */
function _dedupeNameUnderscore(folder, name) {
  const dot = name.lastIndexOf('.');
  const fullStem = dot > -1 ? name.slice(0, dot) : name;
  const ext = dot > -1 ? name.slice(dot) : '';
  
  // 입사일 부분 제거: 마지막 _ 뒤의 6자리 숫자 패턴 제거 (예: _251230, _260106)
  // 또한 이미 붙은 번호도 제거 (예: _1, _2)
  let baseStem = fullStem;
  // 마지막 _ 뒤가 숫자로만 이루어진 경우 제거 (입사일 또는 중복 번호)
  const lastUnderscoreIndex = baseStem.lastIndexOf('_');
  if (lastUnderscoreIndex > -1) {
    const afterLastUnderscore = baseStem.slice(lastUnderscoreIndex + 1);
    // 6자리 숫자(입사일) 또는 숫자만 있는 경우(중복 번호) 제거
    if (/^\d{1,6}$/.test(afterLastUnderscore)) {
      baseStem = baseStem.slice(0, lastUnderscoreIndex);
    }
  }
  
  let n = 0;
  let candidate = name;

  // 타임아웃 방지: 파일 검색에 시간 제한 설정
  const startTime = Date.now();
  const maxTime = 5000; // 최대 5초만 검색
  const existing = new Set();
  let fileCount = 0;
  const maxFiles = 200; // 최대 200개 파일까지만 체크

  try {
    const files = folder.getFiles();
    while (files.hasNext() && (Date.now() - startTime) < maxTime && fileCount < maxFiles) {
      const file = files.next();
      const fileName = file.getName();
      const fileDot = fileName.lastIndexOf('.');
      const fileStem = fileDot > -1 ? fileName.slice(0, fileDot) : fileName;
      
      // 파일명에서도 입사일/중복번호 제거 (확장자는 비교하지 않음)
      let fileBaseStem = fileStem;
      const fileLastUnderscoreIndex = fileBaseStem.lastIndexOf('_');
      if (fileLastUnderscoreIndex > -1) {
        const fileAfterLastUnderscore = fileBaseStem.slice(fileLastUnderscoreIndex + 1);
        if (/^\d{1,6}$/.test(fileAfterLastUnderscore)) {
          fileBaseStem = fileBaseStem.slice(0, fileLastUnderscoreIndex);
        }
      }
      
      // base 부분이 같으면 중복으로 간주 (확장자 무관)
      if (fileBaseStem === baseStem) {
        existing.add(fileName);
      }
      fileCount++;
    }
    
    // 시간 초과 또는 파일이 너무 많으면 경고 로그만 남기고 계속 진행
    if ((Date.now() - startTime) >= maxTime) {
      Logger.log('[_dedupeNameUnderscore] 검색 시간 초과, ' + fileCount + '개 파일 확인됨');
    }
    if (fileCount >= maxFiles) {
      Logger.log('[_dedupeNameUnderscore] 파일이 너무 많음, ' + fileCount + '개까지만 확인');
    }
  } catch (error) {
    Logger.log('[_dedupeNameUnderscore] 검색 중 오류: ' + error.toString());
    // 오류 발생 시 원본 이름 반환 (중복 가능하지만 에러 방지)
    return name;
  }

  // 중복 체크 및 번호 부여
  while (existing.has(candidate)) {
    n++;
    candidate = `${fullStem}_${n}${ext}`;
    // 무한 루프 방지
    if (n > 1000) {
      Logger.log('[_dedupeNameUnderscore] 중복 번호가 너무 많음, 원본 이름 사용');
      return name;
    }
  }
  
  return candidate;
}

/** dataURL 프리픽스 제거 */
function _stripDataUrlPrefix(s) {
  const i = s.indexOf('base64,');
  return i > -1 ? s.slice(i + 7) : s;
}

/** MIME 추정 */
function _guessContentType(ext) {
  switch (ext.toLowerCase()) {
    case '.png': return 'image/png';
    case '.gif': return 'image/gif';
    case '.webp': return 'image/webp';
    case '.heic': return 'image/heic';
    default: return 'image/jpeg';
  }
}

/** ============================================
 * 등록 웹앱 관련 함수들
 ============================================ */

/** 등록 캐시 키 생성 (날짜별, 모든 사용자 공유)
 * 
 * 저장 위치: Google Apps Script CacheService
 * 키 형식: registrations_YYYY-MM-DD
 * - 날짜별로 구분되므로 자정이 지나면 자동으로 새로운 키 생성
 * - 모든 사용자가 같은 키를 사용하므로 등록 목록이 공유됨
 * - 자정이 지나면 이전 날짜의 캐시는 접근 불가 (자동 초기화)
 */
function getRegistrationCacheKey(today) {
  return `registrations_${today}`;
}

/** 호출 상태 캐시 키 생성 (사번별) */
function getCallStatusCacheKey(empNo) {
  return `call_status_${empNo}`;
}

/** 등록 이벤트 캐시 키 생성 (날짜별) */
function getRegistrationEventKey(today) {
  return `registration_event_${today}`;
}

/** 등록 이벤트 확인 (촬영자 웹앱용) */
function checkRegistrationEvent(lastTimestamp) {
  try {
    const today = _today();
    const eventKey = getRegistrationEventKey(today);
    const cache = CacheService.getScriptCache();
    const eventCached = cache.get(eventKey);
    
    if (!eventCached) {
      return { hasNew: false };
    }
    
    try {
      const eventData = JSON.parse(eventCached);
      if (eventData.timestamp > (lastTimestamp || 0)) {
        Logger.log('[checkRegistrationEvent] 새로운 등록 이벤트 발견: ' + eventData.name + ' (' + eventData.empNo + ')');
        return {
          hasNew: true,
          timestamp: eventData.timestamp,
          empNo: eventData.empNo,
          name: eventData.name
        };
      }
    } catch (e) {
      Logger.log('[checkRegistrationEvent] 이벤트 파싱 실패: ' + e.toString());
    }
    
    return { hasNew: false };
  } catch (error) {
    Logger.log('[checkRegistrationEvent] 예외 발생: ' + error.toString());
    return { hasNew: false };
  }
}

/** 촬영 등록 (사번/이름 입력) */
function registerForShooting(empNo, name) {
  try {
    Logger.log('[registerForShooting] 시작 - 사번: ' + empNo + ', 이름: ' + name);
    
    // RAW 시트에서 검증
    const verifyResult = verifyUser(empNo, name);
    if (!verifyResult || !verifyResult.ok) {
      Logger.log('[registerForShooting] 검증 실패');
      return { ok: false, message: verifyResult?.message || '사번/성명을 찾을 수 없습니다.' };
    }
    
    const today = _today();
    const cacheKey = getRegistrationCacheKey(today);
    const cache = CacheService.getScriptCache();
    
    // 기존 등록 목록 가져오기
    const cached = cache.get(cacheKey);
    let registrations = [];
    if (cached) {
      try {
        registrations = JSON.parse(cached);
      } catch (e) {
        Logger.log('[registerForShooting] 캐시 파싱 실패: ' + e.toString());
        registrations = [];
      }
    }
    
    // 중복 체크 및 업데이트
    let existingIndex = -1;
    for (let i = 0; i < registrations.length; i++) {
      if (registrations[i].empNo === empNo) {
        existingIndex = i;
        break;
      }
    }
    
    const registrationData = {
      empNo: empNo,
      name: verifyResult.name,
      nameEn: verifyResult.nameEn || '',
      joinDate: verifyResult.joinDate || null,
      registeredAt: new Date().getTime(),
      called: false // 호출 상태
    };
    
    if (existingIndex >= 0) {
      registrations[existingIndex] = registrationData;
      Logger.log('[registerForShooting] 기존 등록 업데이트');
    } else {
      registrations.push(registrationData);
      Logger.log('[registerForShooting] 새 등록 추가');
    }
    
    // 캐시 저장 (자정까지 유효, 한국 시간 기준)
    // 자정이 지나면 날짜가 바뀌므로 새로운 캐시 키가 생성되어 자동으로 초기화됨
    const now = new Date();
    const midnight = new Date(now);
    midnight.setHours(24, 0, 0, 0);
    const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
    // 최대 6시간 제한 (GAS 캐시 제한), 하지만 자정까지가 더 짧으면 그 시간 사용
    const cacheExpiration = Math.min(secondsUntilMidnight, 21600);
    
    cache.put(cacheKey, JSON.stringify(registrations), cacheExpiration);
    Logger.log('[registerForShooting] 등록 완료: ' + registrations.length + '개 항목 (만료: ' + cacheExpiration + '초 후, ' + today + ')');
    
    // 등록 완료 이벤트 설정 (촬영자 웹앱 새로고침 신호)
    const registrationEventKey = `registration_event_${today}`;
    const eventData = {
      timestamp: new Date().getTime(),
      empNo: empNo,
      name: verifyResult.name
    };
    cache.put(registrationEventKey, JSON.stringify(eventData), cacheExpiration);
    Logger.log('[registerForShooting] 등록 이벤트 설정 완료');
    
    return { ok: true, message: '등록이 완료되었습니다.' };
  } catch (error) {
    Logger.log('[registerForShooting] 예외 발생: ' + error.toString());
    return { ok: false, message: '등록 중 오류가 발생했습니다: ' + error.toString() };
  }
}

/** 등록된 사람 목록 가져오기 (촬영자 웹앱용) 
 * 
 * 데이터 저장 위치: Google Apps Script CacheService
 * - 캐시 키: registrations_YYYY-MM-DD (날짜별)
 * - 모든 사용자가 같은 캐시 키를 사용하므로 공유됨
 * - 자정이 지나면 날짜가 바뀌어 새로운 캐시 키가 생성되므로 자동 초기화
 * - 캐시 만료 시간: 자정까지 (한국 시간 기준)
 */
function getRegisteredList() {
  try {
    const today = _today();
    const cacheKey = getRegistrationCacheKey(today);
    const cache = CacheService.getScriptCache();
    
    // 오늘 날짜의 캐시만 조회 (자정이 지나면 자동으로 새로운 날짜 키가 생성되어 이전 데이터는 접근 불가)
    const cached = cache.get(cacheKey);
    
    if (cached) {
      try {
        const registrations = JSON.parse(cached);
        // 등록 시간순 정렬
        registrations.sort((a, b) => (a.registeredAt || 0) - (b.registeredAt || 0));
        Logger.log('[getRegisteredList] 등록 목록 반환: ' + registrations.length + '개 (날짜: ' + today + ')');
        return { ok: true, list: registrations, today: today };
      } catch (e) {
        Logger.log('[getRegisteredList] 캐시 파싱 실패: ' + e.toString());
        return { ok: true, list: [], today: today };
      }
    }
    
    Logger.log('[getRegisteredList] 등록 목록 없음 (날짜: ' + today + ')');
    return { ok: true, list: [], today: today };
  } catch (error) {
    Logger.log('[getRegisteredList] 예외 발생: ' + error.toString());
    return { ok: false, message: '목록을 불러오는 중 오류가 발생했습니다.', list: [] };
  }
}

/** 특정인 호출 (등록 목록에 없어도 호출 가능 - 수동 추가한 사람 포함) */
function callPerson(empNo, name) {
  try {
    Logger.log('[callPerson] 시작 - 사번: ' + empNo + ', 이름: ' + (name || '없음'));
    
    const today = _today();
    const cacheKey = getRegistrationCacheKey(today);
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    
    let registrations = [];
    let found = null;
    let personName = name || '';
    
    if (cached) {
      try {
        registrations = JSON.parse(cached);
        // 등록 목록에서 해당 사번 찾기
        for (let i = 0; i < registrations.length; i++) {
          if (registrations[i].empNo === empNo) {
            found = registrations[i];
            personName = found.name; // 등록 목록에 있으면 그 이름 사용
            registrations[i].called = true;
            break;
          }
        }
      } catch (e) {
        Logger.log('[callPerson] 캐시 파싱 실패: ' + e.toString());
      }
    }
    
    // 등록 목록에 없어도 호출 가능 (수동 추가한 사람 포함). 이름이 없으면 RAW 시트(A:G 캐시)에서 조회
    if (!personName) {
      try {
        const data = getRawSheetData();
        const { values } = data;
        if (values && values.length > 0) {
          for (let i = 0; i < values.length; i++) {
            if (String(values[i][0]).trim() === String(empNo).trim()) {
              personName = String(values[i][1] || '').trim();
              break;
            }
          }
        }
      } catch (e) {
        Logger.log('[callPerson] RAW 시트 조회 실패: ' + e.toString());
      }
    }
    
    // 이름을 찾지 못한 경우
    if (!personName) {
      personName = empNo; // 사번으로 대체
    }
    
    // 등록 목록 업데이트 (등록 목록에 있던 경우)
    if (found && cached) {
      const now = new Date();
      const midnight = new Date(now);
      midnight.setHours(24, 0, 0, 0);
      const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
      const cacheExpiration = Math.min(secondsUntilMidnight, 21600);
      cache.put(cacheKey, JSON.stringify(registrations), cacheExpiration);
    }
    
    // 호출 상태 캐시에 저장 (등록 웹앱에서 확인용)
    const callStatusKey = getCallStatusCacheKey(empNo);
    const callStatus = {
      called: true,
      name: personName,
      empNo: empNo,
      calledAt: new Date().getTime()
    };
    cache.put(callStatusKey, JSON.stringify(callStatus), 300); // 5분간 유효
    
    // 최근 호출된 사람 목록에 추가 (등록 목록에 없어도 확인 가능하도록)
    // today 변수가 이미 위에서 선언되었으므로 재선언하지 않음
    const recentCallsKey = `recent_calls_${today}`;
    const recentCallsCached = cache.get(recentCallsKey);
    let recentCalls = [];
    if (recentCallsCached) {
      try {
        recentCalls = JSON.parse(recentCallsCached);
      } catch (e) {
        Logger.log('[callPerson] 최근 호출 목록 파싱 실패: ' + e.toString());
        recentCalls = [];
      }
    }
    // 중복 제거 후 추가
    if (!recentCalls.includes(empNo)) {
      recentCalls.push(empNo);
      // 최대 50개까지만 유지 (메모리 절약)
      if (recentCalls.length > 50) {
        recentCalls = recentCalls.slice(-50);
      }
      const now = new Date();
      const midnight = new Date(now);
      midnight.setHours(24, 0, 0, 0);
      const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
      const cacheExpiration = Math.min(secondsUntilMidnight, 21600);
      cache.put(recentCallsKey, JSON.stringify(recentCalls), cacheExpiration);
      Logger.log('[callPerson] 최근 호출 목록에 추가: ' + empNo + ' (총 ' + recentCalls.length + '개)');
    }
    
    Logger.log('[callPerson] 호출 완료: ' + personName);
    return { ok: true, message: personName + '님을 호출했습니다.', name: personName };
  } catch (error) {
    Logger.log('[callPerson] 예외 발생: ' + error.toString());
    return { ok: false, message: '호출 중 오류가 발생했습니다: ' + error.toString() };
  }
}

/** 호출 상태 확인 (등록 웹앱용) - 모든 호출 상태 확인 (등록 목록에 없어도 확인) */
function checkCallStatus() {
  try {
    const today = _today();
    const cache = CacheService.getScriptCache();
    
    let calledPerson = null;
    let latestCallTime = 0;
    
    // 최근 호출된 사람 목록 캐시 확인 (우선 확인 - 촬영자 대시보드에서 등록한 사람 포함)
    const recentCallsKey = `recent_calls_${today}`;
    const recentCallsCached = cache.get(recentCallsKey);
    
    if (recentCallsCached) {
      try {
        const recentCalls = JSON.parse(recentCallsCached);
        Logger.log('[checkCallStatus] 최근 호출 목록 확인: ' + recentCalls.length + '개');
        for (let i = 0; i < recentCalls.length; i++) {
          const empNo = recentCalls[i];
          const callStatusKey = getCallStatusCacheKey(empNo);
          const callStatusCached = cache.get(callStatusKey);
          if (callStatusCached) {
            try {
              const callStatus = JSON.parse(callStatusCached);
              if (callStatus.called && callStatus.calledAt > latestCallTime) {
                latestCallTime = callStatus.calledAt;
                calledPerson = callStatus;
                Logger.log('[checkCallStatus] 호출된 사람 발견 (최근 호출 목록): ' + callStatus.name + ' (' + callStatus.empNo + ')');
              }
            } catch (e) {
              Logger.log('[checkCallStatus] 호출 상태 파싱 실패: ' + e.toString());
            }
          }
        }
      } catch (e) {
        Logger.log('[checkCallStatus] 최근 호출 목록 파싱 실패: ' + e.toString());
      }
    } else {
      Logger.log('[checkCallStatus] 최근 호출 목록 없음');
    }
    
    // 등록 목록에서도 확인 (최근 호출 목록에 없을 경우 대비)
    const cacheKey = getRegistrationCacheKey(today);
    const cached = cache.get(cacheKey);
    
    if (cached) {
      try {
        const registrations = JSON.parse(cached);
        for (let i = 0; i < registrations.length; i++) {
          if (registrations[i].called) {
            const callStatusKey = getCallStatusCacheKey(registrations[i].empNo);
            const callStatusCached = cache.get(callStatusKey);
            if (callStatusCached) {
              try {
                const callStatus = JSON.parse(callStatusCached);
                if (callStatus.called && callStatus.calledAt > latestCallTime) {
                  latestCallTime = callStatus.calledAt;
                  calledPerson = callStatus;
                  Logger.log('[checkCallStatus] 호출된 사람 발견 (등록 목록): ' + callStatus.name + ' (' + callStatus.empNo + ')');
                }
              } catch (e) {
                // 파싱 실패 시 스킵
              }
            }
          }
        }
      } catch (e) {
        Logger.log('[checkCallStatus] 등록 목록 파싱 실패: ' + e.toString());
      }
    }
    
    if (calledPerson) {
      Logger.log('[checkCallStatus] 최종 호출된 사람: ' + calledPerson.name + ' (' + calledPerson.empNo + ')');
      return {
        called: true,
        name: calledPerson.name,
        empNo: calledPerson.empNo
      };
    }
    
    Logger.log('[checkCallStatus] 호출된 사람 없음');
    return { called: false };
  } catch (error) {
    Logger.log('[checkCallStatus] 예외 발생: ' + error.toString());
    return { called: false };
  }
}

/** 호출 상태 초기화 (등록 웹앱에서 알람 확인 후) */
function clearCallStatus(empNo) {
  try {
    Logger.log('[clearCallStatus] 시작 - 사번: ' + empNo);
    
    // 호출 상태 캐시 삭제
    const callStatusKey = getCallStatusCacheKey(empNo);
    const cache = CacheService.getScriptCache();
    cache.remove(callStatusKey);
    
    // 등록 목록에서 호출 상태 초기화
    const today = _today();
    const cacheKey = getRegistrationCacheKey(today);
    const cached = cache.get(cacheKey);
    
    if (cached) {
      try {
        let registrations = JSON.parse(cached);
        for (let i = 0; i < registrations.length; i++) {
          if (registrations[i].empNo === empNo) {
            registrations[i].called = false;
            break;
          }
        }
        
        const now = new Date();
        const midnight = new Date(now);
        midnight.setHours(24, 0, 0, 0);
        const secondsUntilMidnight = Math.floor((midnight - now) / 1000);
        const cacheExpiration = Math.min(secondsUntilMidnight, 21600);
        
        cache.put(cacheKey, JSON.stringify(registrations), cacheExpiration);
      } catch (e) {
        Logger.log('[clearCallStatus] 캐시 업데이트 실패: ' + e.toString());
      }
    }
    
    Logger.log('[clearCallStatus] 완료');
    return { ok: true };
  } catch (error) {
    Logger.log('[clearCallStatus] 예외 발생: ' + error.toString());
    return { ok: false };
  }
}

/** 호출 상태 확인 (등록 웹앱용) - 특정 사번 확인 */
function checkCallStatusForEmpNo(empNo) {
  try {
    if (!empNo) {
      return { called: false };
    }
    
    const callStatusKey = getCallStatusCacheKey(empNo);
    const cache = CacheService.getScriptCache();
    const callStatusCached = cache.get(callStatusKey);
    
    if (!callStatusCached) {
      return { called: false };
    }
    
    try {
      const callStatus = JSON.parse(callStatusCached);
      if (callStatus.called) {
        return {
          called: true,
          name: callStatus.name,
          empNo: callStatus.empNo
        };
      }
    } catch (e) {
      Logger.log('[checkCallStatusForEmpNo] 파싱 실패: ' + e.toString());
    }
    
    return { called: false };
  } catch (error) {
    Logger.log('[checkCallStatusForEmpNo] 예외 발생: ' + error.toString());
    return { called: false };
  }
}
