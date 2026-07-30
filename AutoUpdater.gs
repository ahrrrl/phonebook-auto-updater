/**
 * ERP 전화번호부 엑셀 자동 다운로드 및 공식 시트 동기화 스크립트
 */

const ERP_CONFIG = {
  // 1. 사용자 계정 정보는 구글 앱 스크립트의 '프로젝트 설정(톱니바퀴) > 스크립트 속성'에서 관리합니다.
  // 속성(Property): ERP_EMAIL / 값(Value): 본인 이메일
  // 속성(Property): ERP_PASSWORD / 값(Value): 본인 비밀번호
  // 속성(Property): ERP_BASE_URL / 값(Value): 서버 주소 (예: https://gs-core-...)
  
  // 2. API 주소 설정 (스크립트 속성 참조)
  get LOGIN_URL() { return (PropertiesService.getScriptProperties().getProperty('ERP_BASE_URL') || '') + '/api/auth/login'; },
  // 게시물 첨부파일 목록 API
  get FILES_LIST_URL() { return (PropertiesService.getScriptProperties().getProperty('ERP_BASE_URL') || '') + '/api/groupware/files/list/RESOURCE'; }, 
  get DOWNLOAD_BASE_URL() { return (PropertiesService.getScriptProperties().getProperty('ERP_BASE_URL') || '') + '/api/groupware/files/download/'; }
};

/**
 * 메인 실행 함수: 트리거(Trigger)에 연결하여 주기적으로 실행되게 설정합니다.
 */
function runAutoSync() {
  const email = PropertiesService.getScriptProperties().getProperty('ERP_EMAIL');
  const password = PropertiesService.getScriptProperties().getProperty('ERP_PASSWORD');
  const baseUrl = PropertiesService.getScriptProperties().getProperty('ERP_BASE_URL');

  if (!email || !password || !baseUrl) {
    throw new Error('이메일, 비밀번호 또는 서버 주소가 설정되지 않았습니다. 좌측 메뉴의 [프로젝트 설정(톱니바퀴) > 스크립트 속성]에서 ERP_EMAIL, ERP_PASSWORD, ERP_BASE_URL을 추가해주세요.');
  }

  Logger.log('1. ERP 로그인 시도 중...');
  const token = loginToErp_(email, password);
  
  Logger.log('2. 최신 첨부파일 정보 확인 중...');
  const fileId = getLatestAttachmentId_(token);
  if (!fileId) {
    Logger.log('첨부파일을 찾을 수 없습니다. 게시판 ID나 구조를 확인하세요.');
    return;
  }
  
  Logger.log('3. 엑셀 파일 다운로드 및 시트 변환 중... (파일 ID: ' + fileId + ')');
  const tempSheetId = downloadAndConvertExcel_(token, fileId);
  
  try {
    Logger.log('4. 공식 구글 시트 업데이트 중...');
    updateOfficialSheet_(tempSheetId);
    
    // 업데이트 성공 시간 기록 (Code.gs에서 읽어갈 수 있도록 저장)
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm');
    PropertiesService.getScriptProperties().setProperty('LAST_SYNC_TIME', now);
    
    Logger.log('동기화 완료!');
  } finally {
    // 임시로 생성된 구글 시트는 정리(휴지통으로 이동)합니다.
    try {
      DriveApp.getFileById(tempSheetId).setTrashed(true);
      Logger.log('임시 변환 파일 삭제 완료.');
    } catch(e) {
      Logger.log('임시 파일 삭제 실패: ' + e.message);
    }
  }
}

/**
 * 1. ERP 로그인 및 토큰 반환
 */
function loginToErp_(email, password) {
  const payload = {
    email: email,
    password: password,
    deviceType: "WEB"
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(ERP_CONFIG.LOGIN_URL, options);
  if (response.getResponseCode() !== 200 && response.getResponseCode() !== 201) {
    throw new Error('로그인 실패: ' + response.getContentText());
  }
  
  // 로그인 응답에 토큰이 어디있는지 확인 필요 (일반적으로 response JSON 안에 accessToken 등으로 존재)
  const data = JSON.parse(response.getContentText());
  const token = data.accessToken || data.token || data.data?.accessToken; 
  
  if (!token) {
    throw new Error('로그인은 성공했으나 토큰을 찾을 수 없습니다. 응답 구조 확인 필요: ' + response.getContentText());
  }
  return token;
}

/**
 * 2. 10번 게시글의 파일 ID 찾기
 * 스크린샷의 10번 게시물 응답을 참조합니다.
 */
function getLatestAttachmentId_(token) {
  // 게시글(10번)의 첨부파일 목록을 요청합니다.
  const articleUrl = ERP_CONFIG.FILES_LIST_URL + '/10';
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(articleUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('게시글 불러오기 실패: ' + response.getContentText());
  }
  
  const data = JSON.parse(response.getContentText());
  
  // 스크린샷 2에서 첨부파일 배열이 반환된 형태를 참조
  // 배열 형태일 수도 있고, 객체 내부에 files 배열이 있을 수도 있습니다.
  let files = [];
  if (Array.isArray(data)) {
    // 응답 자체가 배열인 경우 (스크린샷 2 형태)
    files = data;
  } else if (data.files && Array.isArray(data.files)) {
    // 객체 내부에 files가 있는 경우
    files = data.files;
  } else if (data.data && Array.isArray(data.data.files)) {
    // data 속성 내부에 있는 경우
    files = data.data.files;
  }
  
  if (files.length > 0) {
    // 가장 첫 번째(최신) 또는 특정 엑셀 파일의 ID를 반환
    const excelFile = files.find(f => f.fileName && f.fileName.includes('.xlsx'));
    if (excelFile) return excelFile.id;
    return files[0].id;
  }
  
  return null;
}

/**
 * 3. 엑셀 다운로드 및 구글 시트로 변환
 */
function downloadAndConvertExcel_(token, fileId) {
  const downloadUrl = ERP_CONFIG.DOWNLOAD_BASE_URL + fileId;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(downloadUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('엑셀 다운로드 실패: ' + response.getContentText());
  }
  
  const blob = response.getBlob();
  
  // Drive API(Advanced Service)를 활용하여 엑셀을 구글 시트로 변환하여 저장
  // **주의**: Apps Script 설정에서 'Drive API' 서비스가 활성화되어 있어야 합니다.
  const resource = {
    name: 'Temp_Phonebook_Sync_' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  
  const convertedFile = Drive.Files.create(resource, blob);
  return convertedFile.id;
}

/**
 * 4. 변환된 시트의 데이터를 기존 공식 시트에 덮어쓰기
 */
function updateOfficialSheet_(tempSheetId) {
  // 변환된 구글 시트 열기
  const tempSpreadsheet = SpreadsheetApp.openById(tempSheetId);
  const tempSheet = tempSpreadsheet.getSheets()[0];
  const newData = tempSheet.getDataRange().getValues();
  
  if (newData.length === 0) {
    throw new Error('다운로드한 엑셀 파일에 데이터가 없습니다.');
  }

  // 기존 공식 시트 열기
  const officialSpreadsheet = SpreadsheetApp.openById(OFFICIAL_SPREADSHEET_ID);
  
  let officialSheet = null;
  if (OFFICIAL_SHEET_NAME) {
    officialSheet = officialSpreadsheet.getSheetByName(OFFICIAL_SHEET_NAME);
  }
  
  // 이름이 '비상연락망'인 시트가 없으면 첫 번째 시트를 대체제로 사용 (사용자 제안 반영)
  if (!officialSheet) {
    officialSheet = officialSpreadsheet.getSheets()[0];
  }

  // 임시 시트를 공식 스프레드시트로 복사 (서식 및 병합 보존용)
  const tempCopiedSheet = tempSheet.copyTo(officialSpreadsheet);
  
  // 기존 공식 시트 초기화 (데이터, 서식, 병합 모두 해제)
  officialSheet.clear();
  
  // 복사해온 시트의 전체 범위를 공식 시트에 붙여넣기 (값, 서식, 병합 모두 포함)
  const sourceRange = tempCopiedSheet.getDataRange();
  sourceRange.copyTo(officialSheet.getRange(1, 1));
  
  // 열 너비 복사
  const numCols = sourceRange.getNumColumns();
  for (let i = 1; i <= numCols; i++) {
    officialSheet.setColumnWidth(i, tempCopiedSheet.getColumnWidth(i));
  }
  
  // 복사했던 임시 시트 삭제
  officialSpreadsheet.deleteSheet(tempCopiedSheet);
}
