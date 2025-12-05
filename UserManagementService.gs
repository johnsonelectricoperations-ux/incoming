/**
 * UserManagementService.gs - 사용자 및 업체 관리 서비스
 */

/**
 * 모든 업체 목록 조회 (사용자 수 포함)
 * @param {string} token - 세션 토큰
 */
function getCompanies(token) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        data: []
      };
    }

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.',
        data: []
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.',
        data: []
      };
    }

    const data = userSheet.getDataRange().getDisplayValues();

    if (data.length <= 1) {
      return {
        success: true,
        data: []
      };
    }

    // 업체별 사용자 수 집계
    const companyMap = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const companyCode = String(row[0] || '').trim();
      const companyName = String(row[1] || '').trim();
      const status = String(row[6] || '').trim();

      if (!companyCode || !companyName) continue;

      if (!companyMap[companyCode]) {
        companyMap[companyCode] = {
          companyCode: companyCode,
          companyName: companyName,
          totalUsers: 0,
          activeUsers: 0
        };
      }

      companyMap[companyCode].totalUsers++;
      if (status === '활성') {
        companyMap[companyCode].activeUsers++;
      }
    }

    // Map을 배열로 변환하고 업체명으로 정렬
    const companies = Object.values(companyMap).sort((a, b) => {
      return a.companyName.localeCompare(b.companyName);
    });

    return {
      success: true,
      data: companies
    };

  } catch (error) {
    Logger.log('getCompanies 오류: ' + error.toString());
    return {
      success: false,
      message: '업체 목록 조회 중 오류가 발생했습니다.',
      data: []
    };
  }
}

/**
 * 특정 업체의 상세 정보 및 사용자 목록 조회
 * @param {string} token - 세션 토큰
 * @param {string} companyCode - 업체 코드
 */
function getCompanyDetails(token, companyCode) {
  try {
    Logger.log('getCompanyDetails 호출됨 - companyCode: ' + companyCode);

    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      Logger.log('getCompanyDetails - 세션 없음');
      return {
        success: false,
        message: '로그인이 필요합니다.',
        data: null
      };
    }

    Logger.log('getCompanyDetails - 세션 확인됨: ' + session.userId + ', 권한: ' + session.role);

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      Logger.log('getCompanyDetails - 권한 부족: ' + session.role);
      return {
        success: false,
        message: '관리자 권한이 필요합니다.',
        data: null
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      Logger.log('getCompanyDetails - Users 시트 없음');
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.',
        data: null
      };
    }

    const data = userSheet.getDataRange().getDisplayValues();
    Logger.log('getCompanyDetails - Users 시트 행 수: ' + data.length);

    if (data.length <= 1) {
      Logger.log('getCompanyDetails - 데이터 없음');
      return {
        success: false,
        message: '해당 업체를 찾을 수 없습니다.',
        data: null
      };
    }

    let companyName = '';
    const users = [];

    // 해당 업체의 사용자 목록 수집
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCompanyCode = String(row[0] || '').trim();

      if (rowCompanyCode === companyCode) {
        if (!companyName) {
          companyName = String(row[1] || '').trim();
        }

        // createdAt을 문자열로 변환 (Date 객체는 직렬화 문제 발생 가능)
        let createdAtStr = '';
        if (row[7]) {
          try {
            if (row[7] instanceof Date) {
              createdAtStr = Utilities.formatDate(row[7], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
            } else {
              createdAtStr = String(row[7]);
            }
          } catch (e) {
            Logger.log('createdAt 변환 오류: ' + e.toString());
            createdAtStr = '';
          }
        }

        users.push({
          rowIndex: i + 1,
          companyCode: rowCompanyCode,
          companyName: String(row[1] || ''),
          name: String(row[2] || ''),
          userId: String(row[3] || ''),
          role: String(row[5] || ''),
          status: String(row[6] || ''),
          createdAt: createdAtStr
        });
      }
    }

    Logger.log('getCompanyDetails - 사용자 수: ' + users.length);

    if (users.length === 0) {
      Logger.log('getCompanyDetails - 업체코드 ' + companyCode + ' 사용자 없음');
      return {
        success: false,
        message: '해당 업체를 찾을 수 없습니다. (업체코드: ' + companyCode + ')',
        data: null
      };
    }

    // 업체별 통계 조회
    Logger.log('getCompanyDetails - 통계 조회 시작: ' + companyName);
    const stats = getCompanyStatistics(companyName);
    Logger.log('getCompanyDetails - 통계 조회 완료');

    const result = {
      success: true,
      data: {
        companyCode: companyCode,
        companyName: companyName,
        users: users,
        stats: stats
      }
    };

    Logger.log('getCompanyDetails - 반환 직전, success: ' + result.success + ', users 수: ' + result.data.users.length);

    return result;

  } catch (error) {
    Logger.log('getCompanyDetails 오류: ' + error.toString());
    Logger.log('에러 스택: ' + error.stack);
    return {
      success: false,
      message: '업체 정보 조회 중 오류가 발생했습니다: ' + error.toString(),
      data: null
    };
  }
}

/**
 * 업체별 통계 조회 (Data, Item, InspectionSpec, Result 개수)
 * @param {string} companyName - 업체명
 */
function getCompanyStatistics(companyName) {
  try {
    Logger.log('getCompanyStatistics 시작 - companyName: ' + companyName);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stats = {
      dataCount: 0,
      itemCount: 0,
      specCount: 0,
      resultCount: 0
    };

    // Data 시트 카운트 (성적서 업로드)
    try {
      const dataSheetName = getDataSheetName(companyName);
      Logger.log('Data 시트명: ' + dataSheetName);
      const dataSheet = ss.getSheetByName(dataSheetName);
      if (dataSheet) {
        const dataData = dataSheet.getDataRange().getDisplayValues();
        stats.dataCount = dataData.length > 1 ? dataData.length - 1 : 0;
        Logger.log('Data 카운트: ' + stats.dataCount);
      }
    } catch (e) {
      Logger.log('Data 시트 조회 오류: ' + e.toString());
    }

    // ItemList 시트 카운트
    try {
      const itemSheetName = getItemListSheetName(companyName);
      Logger.log('Item 시트명: ' + itemSheetName);
      const itemSheet = ss.getSheetByName(itemSheetName);
      if (itemSheet) {
        const itemData = itemSheet.getDataRange().getDisplayValues();
        stats.itemCount = itemData.length > 1 ? itemData.length - 1 : 0;
        Logger.log('Item 카운트: ' + stats.itemCount);
      }
    } catch (e) {
      Logger.log('Item 시트 조회 오류: ' + e.toString());
    }

    // InspectionSpec 시트 카운트
    try {
      const specSheetName = getInspectionSpecSheetName(companyName);
      Logger.log('Spec 시트명: ' + specSheetName);
      const specSheet = ss.getSheetByName(specSheetName);
      if (specSheet) {
        const specData = specSheet.getDataRange().getDisplayValues();
        stats.specCount = specData.length > 1 ? specData.length - 1 : 0;
        Logger.log('Spec 카운트: ' + stats.specCount);
      }
    } catch (e) {
      Logger.log('Spec 시트 조회 오류: ' + e.toString());
    }

    // Result 시트 카운트
    try {
      const resultSheetName = getResultSheetName(companyName);
      Logger.log('Result 시트명: ' + resultSheetName);
      const resultSheet = ss.getSheetByName(resultSheetName);
      if (resultSheet) {
        const resultData = resultSheet.getDataRange().getDisplayValues();
        stats.resultCount = resultData.length > 1 ? resultData.length - 1 : 0;
        Logger.log('Result 카운트: ' + stats.resultCount);
      }
    } catch (e) {
      Logger.log('Result 시트 조회 오류: ' + e.toString());
    }

    Logger.log('getCompanyStatistics 완료 - stats: ' + JSON.stringify(stats));
    return stats;

  } catch (error) {
    Logger.log('getCompanyStatistics 전체 오류: ' + error.toString());
    Logger.log('에러 스택: ' + error.stack);
    return {
      dataCount: 0,
      itemCount: 0,
      specCount: 0,
      resultCount: 0
    };
  }
}

/**
 * 새 업체 등록 (첫 사용자와 함께)
 * @param {string} token - 세션 토큰
 * @param {Object} companyData - {companyName, userName, userId, password, role}
 */
function addCompany(token, companyData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.'
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.'
      };
    }

    // 입력값 정규화
    const companyName = (companyData.companyName || '').toString().trim();
    const userName = (companyData.userName || '').toString().trim();
    const userId = (companyData.userId || '').toString().trim();
    const password = (companyData.password || '').toString().trim();
    const role = (companyData.role || '일반').toString().trim();

    // 입력값 검증
    if (!companyName || !userName || !userId || !password) {
      return {
        success: false,
        message: '모든 항목을 입력해주세요.'
      };
    }

    // 권한 검증
    if (!['관리자', 'JEO', '일반'].includes(role)) {
      return {
        success: false,
        message: '올바른 권한을 선택해주세요.'
      };
    }

    // ID 길이 체크
    if (userId.length < 4) {
      return {
        success: false,
        message: 'ID는 최소 4자 이상이어야 합니다.'
      };
    }

    // 비밀번호 길이 체크
    if (password.length < 6) {
      return {
        success: false,
        message: '비밀번호는 최소 6자 이상이어야 합니다.'
      };
    }

    // ID 유효성 검증 (영문, 숫자만 허용)
    const idPattern = /^[a-zA-Z0-9]+$/;
    if (!idPattern.test(userId)) {
      return {
        success: false,
        message: 'ID는 영문과 숫자만 사용할 수 있습니다.'
      };
    }

    // 업체명 중복 체크
    const existingCompanyCode = findCompanyCodeByName(companyName);
    if (existingCompanyCode) {
      return {
        success: false,
        message: '이미 등록된 업체명입니다.'
      };
    }

    // ID 중복 체크
    const data = userSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const existingId = String(row[3] || '').trim();

      if (existingId.toLowerCase() === userId.toLowerCase()) {
        return {
          success: false,
          message: '이미 사용중인 ID입니다.'
        };
      }
    }

    // 새 업체 코드 생성
    const companyCode = generateNextCompanyCode();

    // 업체별 시트 생성 (관리자/JEO 권한은 제외)
    if (role !== '관리자' && role !== 'JEO') {
      const sheetResult = createCompanySheets(companyName, companyCode);
      if (!sheetResult.success) {
        return {
          success: false,
          message: '업체 시트 생성 중 오류가 발생했습니다.'
        };
      }
    } else {
      Logger.log(`관리자/JEO 계정은 시트 생성 생략: ${companyName} (${role})`);
    }

    // 사용자 추가
    const timestamp = new Date();
    userSheet.appendRow([
      companyCode,
      companyName,
      userName,
      userId,
      password,
      role,
      '활성',
      timestamp
    ]);

    Logger.log(`신규 업체 등록 완료: ${companyName} (${companyCode})`);

    return {
      success: true,
      message: `업체 '${companyName}'가 등록되었습니다.`,
      companyCode: companyCode
    };

  } catch (error) {
    Logger.log('addCompany 오류: ' + error.toString());
    return {
      success: false,
      message: '업체 등록 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 기존 업체에 사용자 추가
 * @param {string} token - 세션 토큰
 * @param {Object} userData - {companyCode, userName, userId, password, role}
 */
function addUserToCompany(token, userData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.'
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.'
      };
    }

    // 입력값 정규화
    const companyCode = (userData.companyCode || '').toString().trim();
    const userName = (userData.userName || '').toString().trim();
    const userId = (userData.userId || '').toString().trim();
    const password = (userData.password || '').toString().trim();
    const role = (userData.role || '일반').toString().trim();

    // 입력값 검증
    if (!companyCode || !userName || !userId || !password) {
      return {
        success: false,
        message: '모든 항목을 입력해주세요.'
      };
    }

    // 권한 검증
    if (!['관리자', 'JEO', '일반'].includes(role)) {
      return {
        success: false,
        message: '올바른 권한을 선택해주세요.'
      };
    }

    // ID 길이 체크
    if (userId.length < 4) {
      return {
        success: false,
        message: 'ID는 최소 4자 이상이어야 합니다.'
      };
    }

    // 비밀번호 길이 체크
    if (password.length < 6) {
      return {
        success: false,
        message: '비밀번호는 최소 6자 이상이어야 합니다.'
      };
    }

    // ID 유효성 검증
    const idPattern = /^[a-zA-Z0-9]+$/;
    if (!idPattern.test(userId)) {
      return {
        success: false,
        message: 'ID는 영문과 숫자만 사용할 수 있습니다.'
      };
    }

    // 업체 코드 존재 확인 및 업체명 가져오기
    const data = userSheet.getDataRange().getDisplayValues();
    let companyName = '';

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCompanyCode = String(row[0] || '').trim();

      if (rowCompanyCode === companyCode) {
        companyName = String(row[1] || '').trim();
        break;
      }
    }

    if (!companyName) {
      return {
        success: false,
        message: '존재하지 않는 업체 코드입니다.'
      };
    }

    // ID 중복 체크
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const existingId = String(row[3] || '').trim();

      if (existingId.toLowerCase() === userId.toLowerCase()) {
        return {
          success: false,
          message: '이미 사용중인 ID입니다.'
        };
      }
    }

    // 사용자 추가
    const timestamp = new Date();
    userSheet.appendRow([
      companyCode,
      companyName,
      userName,
      userId,
      password,
      role,
      '활성',
      timestamp
    ]);

    Logger.log(`사용자 추가 완료: ${userId} (${companyCode} - ${companyName})`);

    return {
      success: true,
      message: `사용자가 추가되었습니다.`
    };

  } catch (error) {
    Logger.log('addUserToCompany 오류: ' + error.toString());
    return {
      success: false,
      message: '사용자 추가 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 업체명 수정
 * @param {string} token - 세션 토큰
 * @param {string} companyCode - 업체 코드
 * @param {string} newCompanyName - 새 업체명
 */
function updateCompanyName(token, companyCode, newCompanyName) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.'
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.'
      };
    }

    // 입력값 정규화
    const trimmedCompanyCode = (companyCode || '').toString().trim();
    const trimmedNewCompanyName = (newCompanyName || '').toString().trim();

    // 입력값 검증
    if (!trimmedCompanyCode || !trimmedNewCompanyName) {
      return {
        success: false,
        message: '업체 코드와 업체명을 모두 입력해주세요.'
      };
    }

    // 업체 코드 존재 확인
    const data = userSheet.getDataRange().getDisplayValues();
    let oldCompanyName = '';
    let updatedCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCompanyCode = String(row[0] || '').trim();

      if (rowCompanyCode === trimmedCompanyCode) {
        if (!oldCompanyName) {
          oldCompanyName = String(row[1] || '').trim();
        }

        // 업체명 업데이트
        userSheet.getRange(i + 1, 2).setValue(trimmedNewCompanyName);
        updatedCount++;
      }
    }

    if (updatedCount === 0) {
      return {
        success: false,
        message: '존재하지 않는 업체 코드입니다.'
      };
    }

    // 업체 시트명도 변경 (기존 oldCompanyName → 새 newCompanyName)
    // 시트명 변경 함수 호출
    const renameResult = renameCompanySheets(oldCompanyName, trimmedNewCompanyName);
    if (!renameResult.success) {
      Logger.log(`시트명 변경 실패: ${renameResult.message}`);
      // 시트명 변경 실패해도 업체명은 수정됨
    }

    Logger.log(`업체명 수정 완료: ${oldCompanyName} → ${trimmedNewCompanyName} (${updatedCount}개 사용자)`);

    return {
      success: true,
      message: '업체명이 수정되었습니다.'
    };

  } catch (error) {
    Logger.log('updateCompanyName 오류: ' + error.toString());
    return {
      success: false,
      message: '업체명 수정 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 사용자 정보 수정
 * @param {string} token - 세션 토큰
 * @param {number} rowIndex - Users 시트의 행 번호
 * @param {Object} userData - {name, password, role, status}
 */
function updateUserInfo(token, rowIndex, userData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.'
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.'
      };
    }

    // 입력값 정규화
    const name = (userData.name || '').toString().trim();
    const password = userData.password ? userData.password.toString().trim() : null;
    const role = (userData.role || '').toString().trim();
    const status = (userData.status || '').toString().trim();

    // 입력값 검증
    if (!name || !role || !status) {
      return {
        success: false,
        message: '모든 항목을 입력해주세요.'
      };
    }

    // 권한 검증
    if (!['관리자', 'JEO', '일반'].includes(role)) {
      return {
        success: false,
        message: '올바른 권한을 선택해주세요.'
      };
    }

    // 상태 검증
    if (!['활성', '비활성'].includes(status)) {
      return {
        success: false,
        message: '올바른 상태를 선택해주세요.'
      };
    }

    // 비밀번호 변경 시 길이 체크
    if (password && password.length < 6) {
      return {
        success: false,
        message: '비밀번호는 최소 6자 이상이어야 합니다.'
      };
    }

    // 기존 데이터 가져오기
    const existingData = userSheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    const companyCode = existingData[0];
    const companyName = existingData[1];
    const userId = existingData[3];
    const currentPassword = existingData[4];
    const createdAt = existingData[7];

    // 데이터 수정 (비밀번호는 제공된 경우에만 변경)
    const newPassword = password ? password : currentPassword;

    userSheet.getRange(rowIndex, 1, 1, 8).setValues([[
      companyCode,
      companyName,
      name,
      userId,
      newPassword,
      role,
      status,
      createdAt
    ]]);

    Logger.log(`사용자 정보 수정 완료: ${userId}`);

    return {
      success: true,
      message: '사용자 정보가 수정되었습니다.'
    };

  } catch (error) {
    Logger.log('updateUserInfo 오류: ' + error.toString());
    return {
      success: false,
      message: '사용자 정보 수정 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 사용자 삭제
 * @param {string} token - 세션 토큰
 * @param {number} rowIndex - Users 시트의 행 번호
 */
function deleteUser(token, rowIndex) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.'
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.'
      };
    }

    // 기존 데이터 확인
    const existingData = userSheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    const userId = existingData[3];

    // 자기 자신은 삭제 불가
    if (userId === session.userId) {
      return {
        success: false,
        message: '현재 로그인한 계정은 삭제할 수 없습니다.'
      };
    }

    // 행 삭제
    userSheet.deleteRow(rowIndex);

    Logger.log(`사용자 삭제 완료: ${userId}`);

    return {
      success: true,
      message: '사용자가 삭제되었습니다.'
    };

  } catch (error) {
    Logger.log('deleteUser 오류: ' + error.toString());
    return {
      success: false,
      message: '사용자 삭제 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 업체명으로 검색
 * @param {string} token - 세션 토큰
 * @param {string} searchText - 검색어
 */
function searchCompanies(token, searchText) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        data: []
      };
    }

    // 관리자만 접근 가능
    if (session.role !== '관리자') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.',
        data: []
      };
    }

    const searchLower = (searchText || '').toString().toLowerCase();
    const allCompaniesResult = getCompanies(token);

    if (!allCompaniesResult.success) {
      return allCompaniesResult;
    }

    const filteredCompanies = allCompaniesResult.data.filter(company => {
      return company.companyName.toLowerCase().includes(searchLower) ||
             company.companyCode.toLowerCase().includes(searchLower);
    });

    return {
      success: true,
      data: filteredCompanies
    };

  } catch (error) {
    Logger.log('searchCompanies 오류: ' + error.toString());
    return {
      success: false,
      message: '업체 검색 중 오류가 발생했습니다.',
      data: []
    };
  }
}

/**
 * 비밀번호 변경 (일반 사용자용)
 * @param {string} token - 세션 토큰
 * @param {string} currentPassword - 현재 비밀번호
 * @param {string} newPassword - 새 비밀번호
 */
function changePassword(token, currentPassword, newPassword) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.'
      };
    }

    // 입력값 정규화
    const trimmedCurrentPassword = (currentPassword || '').toString().trim();
    const trimmedNewPassword = (newPassword || '').toString().trim();

    // 입력값 검증
    if (!trimmedCurrentPassword || !trimmedNewPassword) {
      return {
        success: false,
        message: '현재 비밀번호와 새 비밀번호를 모두 입력해주세요.'
      };
    }

    // 새 비밀번호 길이 체크
    if (trimmedNewPassword.length < 6) {
      return {
        success: false,
        message: '새 비밀번호는 최소 6자 이상이어야 합니다.'
      };
    }

    // 현재 비밀번호와 새 비밀번호가 같은지 확인
    if (trimmedCurrentPassword === trimmedNewPassword) {
      return {
        success: false,
        message: '현재 비밀번호와 새 비밀번호가 동일합니다.'
      };
    }

    // 사용자 정보 조회
    const data = userSheet.getDataRange().getDisplayValues();
    let userRowIndex = -1;
    let storedPassword = '';

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUserId = String(row[3] || '').trim();

      if (rowUserId === session.userId) {
        userRowIndex = i + 1; // 시트의 행 번호 (1-based)
        storedPassword = String(row[4] || '').trim();
        break;
      }
    }

    if (userRowIndex === -1) {
      return {
        success: false,
        message: '사용자 정보를 찾을 수 없습니다.'
      };
    }

    // 현재 비밀번호 확인
    if (storedPassword !== trimmedCurrentPassword) {
      return {
        success: false,
        message: '현재 비밀번호가 올바르지 않습니다.'
      };
    }

    // 비밀번호 업데이트 (5번째 컬럼 = 비밀번호)
    userSheet.getRange(userRowIndex, 5).setValue(trimmedNewPassword);

    Logger.log(`비밀번호 변경 완료: ${session.userId}`);

    return {
      success: true,
      message: '비밀번호가 성공적으로 변경되었습니다.'
    };

  } catch (error) {
    Logger.log('changePassword 오류: ' + error.toString());
    return {
      success: false,
      message: '비밀번호 변경 중 오류가 발생했습니다.'
    };
  }
}
