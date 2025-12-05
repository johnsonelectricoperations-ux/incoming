/**
 * 협력사 검사성적서 관리 시스템
 * Code.gs - 메인 진입점 및 웹앱 라우팅
 */

// 웹앱 배포 URL (반드시 본인의 배포 URL로 변경)
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxpjukKc_G4SaPL52w47ANnQc8MvCmkP6IzUX8DaXBqqdz_L8Th3kLqv4_HdslfGnxIBw/exec';

// 스프레드시트 및 시트 이름 설정
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const USER_SHEET_NAME = 'Users';
const DATA_SHEET_NAME = 'Data';
const ITEM_LIST_SHEET_NAME = 'ItemList';
const INSPECTION_SPEC_SHEET_NAME = 'InspectionSpec';
const INSPECTION_RESULT_SHEET_NAME = '수입검사결과';

// Drive 폴더 설정
const DRIVE_FOLDER_NAME = '검사성적서_PDF파일';

// 세션 타임아웃 (30분 = 1800000ms)
const SESSION_TIMEOUT = 30 * 60 * 1000;

/**
 * 웹앱 진입점 (토큰 기반)
 */
function doGet(e) {
  // URL 파라미터에서 토큰 추출
  const token = e.parameter.token;

  // 디버깅용 로그 추가
  Logger.log('doGet - 토큰 존재: ' + (token ? 'yes' : 'no'));
  Logger.log('doGet - 페이지 파라미터: ' + e.parameter.page);

  // 토큰 검증
  let session = null;
  if (token) {
    session = getSessionByToken(token);
    Logger.log('doGet - 세션 검증 결과: ' + (session ? 'valid' : 'invalid'));
  }

  // 로그인 체크
  if (!session || !session.userId) {
    Logger.log('세션 없음 또는 만료 - 로그인 페이지로 이동');
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('검사성적서 관리 시스템')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 세션은 유효함 (getSessionByToken이 자동으로 갱신함)
  Logger.log('doGet - 유효한 세션: ' + session.userId);

  // 페이지 라우팅
  const page = e.parameter.page || 'dashboard';
  let template;

  switch(page) {
    case 'entry':
      template = HtmlService.createTemplateFromFile('DataEntry');
      break;
    case 'history':
      template = HtmlService.createTemplateFromFile('DataHistory');
      break;
    case 'certificate':
      template = HtmlService.createTemplateFromFile('Certificate');
      break;
    case 'inspection':
      template = HtmlService.createTemplateFromFile('InspectionSpec');
      break;
    case 'inspection-history':
      template = HtmlService.createTemplateFromFile('InspectionResultHistory');
      break;
    case 'users':
      // 관리자/JEO만 접근 가능
      if (session.role !== '관리자' && session.role !== 'JEO') {
        template = HtmlService.createTemplateFromFile('Dashboard');
      } else {
        template = HtmlService.createTemplateFromFile('UserManagement');
      }
      break;
    case 'dashboard':
    default:
      template = HtmlService.createTemplateFromFile('Dashboard');
  }

  // 템플릿에 토큰 전달
  template.token = token;

  return template.evaluate()
    .setTitle('검사성적서 관리 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 웹앱 URL 가져오기
 */
function getWebAppUrl() {
  return WEB_APP_URL;
}

/**
 * Drive 폴더 가져오기 또는 생성
 */
function getOrCreateDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
}

/**
 * 업체별 Drive 폴더 가져오기 또는 생성
 */
function getOrCreateCompanyFolder(companyName) {
  const mainFolder = getOrCreateDriveFolder();
  const folders = mainFolder.getFoldersByName(companyName);

  if (folders.hasNext()) {
    return folders.next();
  } else {
    return mainFolder.createFolder(companyName);
  }
}

/**
 * 업체별 년도/월 폴더 가져오기 또는 생성
 * @param {string} companyName - 업체명
 * @param {Date|string} date - 날짜 (Date 객체 또는 "YYYY-MM-DD" 문자열)
 * @returns {Folder} 월별 폴더
 */
function getOrCreateMonthlyFolder(companyName, date) {
  // 업체 폴더 가져오기
  const companyFolder = getOrCreateCompanyFolder(companyName);

  // 날짜에서 년도/월 추출
  let year, month;

  if (date instanceof Date) {
    year = Utilities.formatDate(date, 'Asia/Seoul', 'yyyy');
    month = Utilities.formatDate(date, 'Asia/Seoul', 'MM');
  } else if (typeof date === 'string') {
    // "YYYY-MM-DD" 또는 "YYYY/MM/DD" 형식
    const parts = date.split(/[-/]/);
    if (parts.length >= 2) {
      year = parts[0];
      month = parts[1];
    } else {
      // 날짜 형식이 올바르지 않으면 현재 날짜 사용
      const today = new Date();
      year = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy');
      month = Utilities.formatDate(today, 'Asia/Seoul', 'MM');
    }
  } else {
    // 기본값: 현재 날짜
    const today = new Date();
    year = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy');
    month = Utilities.formatDate(today, 'Asia/Seoul', 'MM');
  }

  // 년도 폴더 가져오기 또는 생성
  let yearFolder;
  const yearFolders = companyFolder.getFoldersByName(year);
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = companyFolder.createFolder(year);
  }

  // 월 폴더 가져오기 또는 생성
  let monthFolder;
  const monthFolders = yearFolder.getFoldersByName(month);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(month);
  }

  return monthFolder;
}

/**
 * 시트 초기화 함수 (최초 1회 실행)
 * 업체별 시트 구조로 업데이트됨 (2025)
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('=== 시트 초기화 시작 ===');

  // 1. Users 시트 생성
  let userSheet = ss.getSheetByName(USER_SHEET_NAME);
  if (!userSheet) {
    Logger.log('Users 시트 생성 중...');
    userSheet = ss.insertSheet(USER_SHEET_NAME);
    userSheet.getRange('A1:H1').setValues([[
      '업체CODE', '업체명', '성명', 'ID', 'PW', '권한', '상태', '생성일'
    ]]);
    userSheet.getRange('A1:H1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

    // 샘플 관리자 계정 추가
    userSheet.appendRow([
      'C01', 'JEO본사', '시스템관리자', 'admin', 'admin1234', '관리자', '활성', new Date()
    ]);

    Logger.log('Users 시트 생성 완료');
  } else {
    Logger.log('Users 시트가 이미 존재합니다.');
  }

  // 2. Data 시트 생성 (기존 호환성 유지용)
  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    Logger.log('Data 시트 생성 중...');
    dataSheet = ss.insertSheet(DATA_SHEET_NAME);
    dataSheet.getRange('A1:K1').setValues([[
      'ID', '업체명', '입고날짜', '입고시간', 'TM-NO', '제품명', '수량', 'PDF_URL', '등록일시', '등록자', '수정일시'
    ]]);
    dataSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

    // TM-NO 열(E열, 5번째 컬럼)을 텍스트 형식으로 미리 설정
    dataSheet.getRange('E:E').setNumberFormat('@STRING@');

    Logger.log('Data 시트 생성 완료');
  } else {
    Logger.log('Data 시트가 이미 존재합니다.');
  }

  // 3. JEO본사 업체 시트 생성
  Logger.log('JEO본사 업체 시트 생성 중...');
  const jeoSheetsResult = createCompanySheets('JEO본사', 'C01');
  if (jeoSheetsResult.success) {
    Logger.log('JEO본사 시트 생성 완료: ' + JSON.stringify(jeoSheetsResult.sheets));
  } else {
    Logger.log('JEO본사 시트 생성 실패: ' + jeoSheetsResult.message);
  }

  // 4. 업무연락 시트 생성
  let workNoticeSheet = ss.getSheetByName('업무연락');
  if (!workNoticeSheet) {
    Logger.log('업무연락 시트 생성 중...');
    workNoticeSheet = ss.insertSheet('업무연락');
    workNoticeSheet.getRange('A1:J1').setValues([[
      'ID', '제목', '내용', '대상업체', '작성자', '작성일시', '게시시작일', '게시종료일', '중요여부', '활성화여부'
    ]]);
    workNoticeSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#673ab7').setFontColor('#ffffff');

    // ID 열(A열)을 텍스트 형식으로 설정
    workNoticeSheet.getRange('A:A').setNumberFormat('@STRING@');

    Logger.log('업무연락 시트 생성 완료');
  } else {
    Logger.log('업무연락 시트가 이미 존재합니다.');
  }

  // 5. 공지사항 시트 생성
  let noticeSheet = ss.getSheetByName('공지사항');
  if (!noticeSheet) {
    Logger.log('공지사항 시트 생성 중...');
    noticeSheet = ss.insertSheet('공지사항');
    noticeSheet.getRange('A1:I1').setValues([[
      'ID', '제목', '내용', '작성자', '작성일시', '게시시작일', '게시종료일', '중요여부', '활성화여부'
    ]]);
    noticeSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#ff5722').setFontColor('#ffffff');

    // ID 열(A열)을 텍스트 형식으로 설정
    noticeSheet.getRange('A:A').setNumberFormat('@STRING@');

    Logger.log('공지사항 시트 생성 완료');
  } else {
    Logger.log('공지사항 시트가 이미 존재합니다.');
  }

  Logger.log('=== 시트 초기화 완료 ===');
  Logger.log('');
  Logger.log('초기 로그인 정보:');
  Logger.log('  ID: admin');
  Logger.log('  비밀번호: admin1234');
  Logger.log('  권한: 관리자');
  Logger.log('');
  Logger.log('웹앱 URL: ' + WEB_APP_URL);

  return {
    success: true,
    message: '시트 초기화가 완료되었습니다.',
    adminId: 'admin',
    adminPassword: 'admin1234'
  };
}

/**
 * 자동 마이그레이션 - 기존 시스템에 누락된 시트 추가
 * 로그인 시 자동 실행
 */
function autoMigrateSheetsIfNeeded() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let migrated = false;

    // 업무연락 시트 확인 및 생성
    let workNoticeSheet = ss.getSheetByName('업무연락');
    if (!workNoticeSheet) {
      Logger.log('[자동 마이그레이션] 업무연락 시트 생성 중...');
      workNoticeSheet = ss.insertSheet('업무연락');
      workNoticeSheet.getRange('A1:J1').setValues([[
        'ID', '제목', '내용', '대상업체', '작성자', '작성일시', '게시시작일', '게시종료일', '중요여부', '활성화여부'
      ]]);
      workNoticeSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#673ab7').setFontColor('#ffffff');
      workNoticeSheet.getRange('A:A').setNumberFormat('@STRING@');
      migrated = true;
      Logger.log('[자동 마이그레이션] 업무연락 시트 생성 완료');
    }

    // 공지사항 시트 확인 및 생성
    let noticeSheet = ss.getSheetByName('공지사항');
    if (!noticeSheet) {
      Logger.log('[자동 마이그레이션] 공지사항 시트 생성 중...');
      noticeSheet = ss.insertSheet('공지사항');
      noticeSheet.getRange('A1:I1').setValues([[
        'ID', '제목', '내용', '작성자', '작성일시', '게시시작일', '게시종료일', '중요여부', '활성화여부'
      ]]);
      noticeSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#ff5722').setFontColor('#ffffff');
      noticeSheet.getRange('A:A').setNumberFormat('@STRING@');
      migrated = true;
      Logger.log('[자동 마이그레이션] 공지사항 시트 생성 완료');
    }

    if (migrated) {
      Logger.log('[자동 마이그레이션] 시트 마이그레이션 완료');
    }

    return { success: true, migrated: migrated };
  } catch (error) {
    Logger.log('[자동 마이그레이션] 오류: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * 스프레드시트 가져오기
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * 시트 가져오기
 */
function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

/**
 * HTML 파일 Include 함수
 * 템플릿에서 <?!= include('파일명') ?>로 사용
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
