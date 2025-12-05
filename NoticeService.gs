/**
 * NoticeService.gs - 공지사항/업무연락 관리 서비스
 */

const NOTICE_SHEET_NAME = '공지사항';
const WORK_NOTICE_SHEET_NAME = '업무연락';

/**
 * 업무연락 추가
 */
function addWorkNotice(token, noticeData) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    // 필수 입력값 검증
    if (!noticeData.title || !noticeData.content || !noticeData.targetCompany) {
      return { success: false, message: '제목, 내용, 대상업체는 필수 입력 항목입니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(WORK_NOTICE_SHEET_NAME);

    // 시트가 없으면 생성
    if (!sheet) {
      sheet = ss.insertSheet(WORK_NOTICE_SHEET_NAME);
      sheet.getRange('A1:J1').setValues([[
        'ID', '제목', '내용', '대상업체', '작성자', '작성일시', '게시시작일', '게시종료일', '중요여부', '활성화여부'
      ]]);
      sheet.getRange('A1:J1').setFontWeight('bold').setBackground('#673ab7').setFontColor('#ffffff');
      sheet.getRange('A:A').setNumberFormat('@STRING@');
    }

    // ID 생성 (타임스탬프 기반)
    const noticeId = 'WN' + new Date().getTime();
    const now = formatDateTime(new Date());

    // 중요여부 기본값 설정
    const isImportant = noticeData.isImportant || 'N';
    const isActive = noticeData.isActive !== undefined ? noticeData.isActive : 'Y';

    // 데이터 추가
    sheet.appendRow([
      noticeId,
      noticeData.title,
      noticeData.content,
      noticeData.targetCompany,
      session.name,
      now,
      noticeData.startDate || '',
      noticeData.endDate || '',
      isImportant,
      isActive
    ]);

    Logger.log(`업무연락 추가 완료: ${noticeId} by ${session.name}`);

    return {
      success: true,
      message: '업무연락이 등록되었습니다.',
      noticeId: noticeId
    };
  } catch (error) {
    Logger.log('업무연락 추가 오류: ' + error.toString());
    return {
      success: false,
      message: '업무연락 추가 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 업무연락 조회 (관리자용 - 전체)
 */
function getAllWorkNotices(token) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(WORK_NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: true, data: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [] };
    }

    // 전체 데이터 조회
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const notices = values.map((row, index) => ({
      rowIndex: index + 2,
      id: row[0],
      title: row[1],
      content: row[2],
      targetCompany: row[3],
      author: row[4],
      createdAt: formatDateTime(row[5]),
      startDate: row[6] ? formatDate(row[6]) : '',
      endDate: row[7] ? formatDate(row[7]) : '',
      isImportant: row[8],
      isActive: row[9]
    })).filter(notice => notice.id); // ID가 있는 것만 필터링

    // 최신순 정렬
    notices.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    return {
      success: true,
      data: notices
    };
  } catch (error) {
    Logger.log('업무연락 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '업무연락 조회 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 활성 업무연락 조회 (사용자용 - 본인 업체 대상만)
 */
function getActiveWorkNotices(token) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(WORK_NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: true, data: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [] };
    }

    // 전체 데이터 조회
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const notices = values.map((row, index) => ({
      rowIndex: index + 2,
      id: row[0],
      title: row[1],
      content: row[2],
      targetCompany: row[3],
      author: row[4],
      createdAt: formatDateTime(row[5]),
      startDate: row[6] ? formatDate(row[6]) : '',
      endDate: row[7] ? formatDate(row[7]) : '',
      isImportant: row[8],
      isActive: row[9]
    })).filter(notice => {
      // ID가 있고 활성화된 것만
      if (!notice.id || notice.isActive !== 'Y') {
        return false;
      }

      // 대상업체 체크 (전체 또는 본인 업체)
      if (notice.targetCompany !== '전체' && notice.targetCompany !== session.companyName) {
        return false;
      }

      // 게시 기간 체크
      if (notice.startDate) {
        const startDate = new Date(notice.startDate);
        startDate.setHours(0, 0, 0, 0);
        if (today < startDate) {
          return false;
        }
      }

      if (notice.endDate) {
        const endDate = new Date(notice.endDate);
        endDate.setHours(23, 59, 59, 999);
        if (today > endDate) {
          return false;
        }
      }

      return true;
    });

    // 중요도순, 최신순 정렬
    notices.sort((a, b) => {
      // 중요 공지가 위로
      if (a.isImportant === 'Y' && b.isImportant !== 'Y') return -1;
      if (a.isImportant !== 'Y' && b.isImportant === 'Y') return 1;
      // 같은 중요도면 최신순
      return new Date(b.createdAt) - new Date(a.createdAt);
    });

    return {
      success: true,
      data: notices
    };
  } catch (error) {
    Logger.log('활성 업무연락 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '업무연락 조회 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 업무연락 수정
 */
function updateWorkNotice(token, rowIndex, noticeData) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    // 필수 입력값 검증
    if (!noticeData.title || !noticeData.content || !noticeData.targetCompany) {
      return { success: false, message: '제목, 내용, 대상업체는 필수 입력 항목입니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(WORK_NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: '업무연락 시트를 찾을 수 없습니다.' };
    }

    // 기존 ID 유지
    const noticeId = sheet.getRange(rowIndex, 1).getValue();

    // 데이터 업데이트
    sheet.getRange(rowIndex, 2, 1, 8).setValues([[
      noticeData.title,
      noticeData.content,
      noticeData.targetCompany,
      sheet.getRange(rowIndex, 5).getValue(), // 작성자 유지
      sheet.getRange(rowIndex, 6).getValue(), // 작성일시 유지
      noticeData.startDate || '',
      noticeData.endDate || '',
      noticeData.isImportant || 'N'
    ]]);

    // 활성화 여부 업데이트
    if (noticeData.isActive !== undefined) {
      sheet.getRange(rowIndex, 10).setValue(noticeData.isActive);
    }

    Logger.log(`업무연락 수정 완료: ${noticeId} by ${session.name}`);

    return {
      success: true,
      message: '업무연락이 수정되었습니다.'
    };
  } catch (error) {
    Logger.log('업무연락 수정 오류: ' + error.toString());
    return {
      success: false,
      message: '업무연락 수정 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 업무연락 삭제
 */
function deleteWorkNotice(token, rowIndex) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(WORK_NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: '업무연락 시트를 찾을 수 없습니다.' };
    }

    const noticeId = sheet.getRange(rowIndex, 1).getValue();

    // 행 삭제
    sheet.deleteRow(rowIndex);

    Logger.log(`업무연락 삭제 완료: ${noticeId} by ${session.name}`);

    return {
      success: true,
      message: '업무연락이 삭제되었습니다.'
    };
  } catch (error) {
    Logger.log('업무연락 삭제 오류: ' + error.toString());
    return {
      success: false,
      message: '업무연락 삭제 중 오류가 발생했습니다: ' + error.message
    };
  }
}

// ========== 공지사항 관리 함수들 ==========

/**
 * 공지사항 추가
 */
function addNotice(token, noticeData) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    // 필수 입력값 검증
    if (!noticeData.title || !noticeData.content) {
      return { success: false, message: '제목과 내용은 필수 입력 항목입니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(NOTICE_SHEET_NAME);

    // 시트가 없으면 생성
    if (!sheet) {
      sheet = ss.insertSheet(NOTICE_SHEET_NAME);
      sheet.getRange('A1:I1').setValues([[
        'ID', '제목', '내용', '작성자', '작성일시', '게시시작일', '게시종료일', '중요여부', '활성화여부'
      ]]);
      sheet.getRange('A1:I1').setFontWeight('bold').setBackground('#ff5722').setFontColor('#ffffff');
      sheet.getRange('A:A').setNumberFormat('@STRING@');
    }

    // ID 생성 (타임스탬프 기반)
    const noticeId = 'N' + new Date().getTime();
    const now = formatDateTime(new Date());

    // 중요여부 기본값 설정
    const isImportant = noticeData.isImportant || 'N';
    const isActive = noticeData.isActive !== undefined ? noticeData.isActive : 'Y';

    // 데이터 추가
    sheet.appendRow([
      noticeId,
      noticeData.title,
      noticeData.content,
      session.name,
      now,
      noticeData.startDate || '',
      noticeData.endDate || '',
      isImportant,
      isActive
    ]);

    Logger.log(`공지사항 추가 완료: ${noticeId} by ${session.name}`);

    return {
      success: true,
      message: '공지사항이 등록되었습니다.',
      noticeId: noticeId
    };
  } catch (error) {
    Logger.log('공지사항 추가 오류: ' + error.toString());
    return {
      success: false,
      message: '공지사항 추가 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 공지사항 조회 (관리자용 - 전체)
 */
function getAllNotices(token) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: true, data: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [] };
    }

    // 전체 데이터 조회
    const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const notices = values.map((row, index) => ({
      rowIndex: index + 2,
      id: row[0],
      title: row[1],
      content: row[2],
      author: row[3],
      createdAt: formatDateTime(row[4]),
      startDate: row[5] ? formatDate(row[5]) : '',
      endDate: row[6] ? formatDate(row[6]) : '',
      isImportant: row[7],
      isActive: row[8]
    })).filter(notice => notice.id); // ID가 있는 것만 필터링

    // 최신순 정렬
    notices.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    return {
      success: true,
      data: notices
    };
  } catch (error) {
    Logger.log('공지사항 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '공지사항 조회 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 활성 공지사항 조회 (사용자용 - 전체 업체)
 */
function getActiveNotices(token) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: true, data: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [] };
    }

    // 전체 데이터 조회
    const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const notices = values.map((row, index) => ({
      rowIndex: index + 2,
      id: row[0],
      title: row[1],
      content: row[2],
      author: row[3],
      createdAt: formatDateTime(row[4]),
      startDate: row[5] ? formatDate(row[5]) : '',
      endDate: row[6] ? formatDate(row[6]) : '',
      isImportant: row[7],
      isActive: row[8]
    })).filter(notice => {
      // ID가 있고 활성화된 것만
      if (!notice.id || notice.isActive !== 'Y') {
        return false;
      }

      // 게시 기간 체크
      if (notice.startDate) {
        const startDate = new Date(notice.startDate);
        startDate.setHours(0, 0, 0, 0);
        if (today < startDate) {
          return false;
        }
      }

      if (notice.endDate) {
        const endDate = new Date(notice.endDate);
        endDate.setHours(23, 59, 59, 999);
        if (today > endDate) {
          return false;
        }
      }

      return true;
    });

    // 중요도순, 최신순 정렬
    notices.sort((a, b) => {
      // 중요 공지가 위로
      if (a.isImportant === 'Y' && b.isImportant !== 'Y') return -1;
      if (a.isImportant !== 'Y' && b.isImportant === 'Y') return 1;
      // 같은 중요도면 최신순
      return new Date(b.createdAt) - new Date(a.createdAt);
    });

    return {
      success: true,
      data: notices
    };
  } catch (error) {
    Logger.log('활성 공지사항 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '공지사항 조회 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 공지사항 수정
 */
function updateNotice(token, rowIndex, noticeData) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    // 필수 입력값 검증
    if (!noticeData.title || !noticeData.content) {
      return { success: false, message: '제목과 내용은 필수 입력 항목입니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: '공지사항 시트를 찾을 수 없습니다.' };
    }

    // 기존 ID 유지
    const noticeId = sheet.getRange(rowIndex, 1).getValue();

    // 데이터 업데이트 (작성자, 작성일시는 유지)
    sheet.getRange(rowIndex, 2, 1, 7).setValues([[
      noticeData.title,
      noticeData.content,
      sheet.getRange(rowIndex, 4).getValue(), // 작성자 유지
      sheet.getRange(rowIndex, 5).getValue(), // 작성일시 유지
      noticeData.startDate || '',
      noticeData.endDate || '',
      noticeData.isImportant || 'N'
    ]]);

    // 활성화 여부 업데이트
    if (noticeData.isActive !== undefined) {
      sheet.getRange(rowIndex, 9).setValue(noticeData.isActive);
    }

    Logger.log(`공지사항 수정 완료: ${noticeId} by ${session.name}`);

    return {
      success: true,
      message: '공지사항이 수정되었습니다.'
    };
  } catch (error) {
    Logger.log('공지사항 수정 오류: ' + error.toString());
    return {
      success: false,
      message: '공지사항 수정 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 공지사항 삭제
 */
function deleteNotice(token, rowIndex) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 관리자 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return { success: false, message: '권한이 없습니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(NOTICE_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: '공지사항 시트를 찾을 수 없습니다.' };
    }

    const noticeId = sheet.getRange(rowIndex, 1).getValue();

    // 행 삭제
    sheet.deleteRow(rowIndex);

    Logger.log(`공지사항 삭제 완료: ${noticeId} by ${session.name}`);

    return {
      success: true,
      message: '공지사항이 삭제되었습니다.'
    };
  } catch (error) {
    Logger.log('공지사항 삭제 오류: ' + error.toString());
    return {
      success: false,
      message: '공지사항 삭제 중 오류가 발생했습니다: ' + error.message
    };
  }
}
