/**
 * Auth.gs - 로그인 및 세션 관리
 */

/**
 * 로그인 처리
 */
function loginUser(userId, password) {
  try {
    // 입력값 검증 및 정규화
    if (!userId || !password) {
      return {
        success: false,
        message: 'ID와 비밀번호를 입력해주세요.'
      };
    }

    const normalizedUserId = userId.toString().trim();
    const normalizedPassword = password.toString().trim();

    if (!normalizedUserId || !normalizedPassword) {
      return {
        success: false,
        message: 'ID와 비밀번호를 입력해주세요.'
      };
    }

    // Users 시트 존재 확인
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      logError('loginUser', 'Users 시트를 찾을 수 없습니다.');
      return {
        success: false,
        message: '시스템 오류: 사용자 데이터베이스를 찾을 수 없습니다. 관리자에게 문의하세요.'
      };
    }

    const data = userSheet.getDataRange().getDisplayValues();

    // 헤더만 있는 경우 확인
    if (data.length <= 1) {
      return {
        success: false,
        message: '등록된 사용자가 없습니다. 관리자에게 문의하세요.'
      };
    }

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 빈 행 건너뛰기
      if (!row || row.length === 0 || !row[2] || !row[3]) {
        continue;
      }

      const [companyName, name, id, pw, role, status] = row;

      // ID와 비밀번호 정규화 후 비교
      const normalizedId = (id || '').toString().trim();
      const normalizedPw = (pw || '').toString().trim();

      // ID와 비밀번호 확인 (대소문자 구분)
      if (normalizedId === normalizedUserId && normalizedPw === normalizedPassword) {
        // 상태 확인 (공백 제거 후 비교)
        const normalizedStatus = (status || '').toString().trim();
        if (normalizedStatus !== '활성') {
          return {
            success: false,
            message: '비활성화된 계정입니다. 관리자에게 문의하세요.'
          };
        }

        // 세션 생성
        const sessionData = {
          userId: normalizedUserId,
          companyName: companyName || '',
          name: name || '',
          role: role || '일반',
          loginTime: new Date().getTime()
        };

        try {
          setSession(sessionData);

          // 자동 마이그레이션 (누락된 시트 생성)
          autoMigrateSheetsIfNeeded();

          // 세션 저장 검증
          Utilities.sleep(200); // 200ms 대기
          const savedSession = getSession();

          if (!savedSession || savedSession.userId !== normalizedUserId) {
            throw new Error('세션 저장 실패');
          }

          Logger.log('로그인 성공: ' + normalizedUserId);

          return {
            success: true,
            message: '로그인 성공',
            user: {
              userId: normalizedUserId,
              companyName: companyName || '',
              name: name || '',
              role: role || '일반'
            }
          };
        } catch (sessionError) {
          logError('loginUser', '세션 저장 중 오류: ' + sessionError.toString());
          return {
            success: false,
            message: '로그인 처리 중 오류가 발생했습니다. 다시 시도해주세요.'
          };
        }
      }
    }

    return {
      success: false,
      message: 'ID 또는 비밀번호가 올바르지 않습니다.'
    };

  } catch (error) {
    logError('loginUser', error);

    let userMessage = '로그인 처리 중 오류가 발생했습니다.';

    // 에러 타입별 메시지
    const errorString = error.toString().toLowerCase();
    if (errorString.includes('permission') || errorString.includes('권한')) {
      userMessage = '시스템 접근 권한이 없습니다. 관리자에게 문의하세요.';
    } else if (errorString.includes('not found') || errorString.includes('찾을 수 없')) {
      userMessage = '시스템 구성 오류입니다. 관리자에게 문의하세요.';
    } else if (errorString.includes('network') || errorString.includes('timeout')) {
      userMessage = '네트워크 오류가 발생했습니다. 다시 시도해주세요.';
    }

    return {
      success: false,
      message: userMessage
    };
  }
}

/**
 * 로그아웃 처리
 */
function logoutUser() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteAllProperties();
    
    return {
      success: true,
      message: '로그아웃 되었습니다.'
    };
  } catch (error) {
    Logger.log('로그아웃 오류: ' + error.toString());
    return {
      success: false,
      message: '로그아웃 처리 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 세션 저장
 */
function setSession(sessionData) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('session', JSON.stringify(sessionData));
}

/**
 * 세션 가져오기
 */
function getSession() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const sessionString = userProperties.getProperty('session');
    
    if (!sessionString) {
      return null;
    }
    
    const session = JSON.parse(sessionString);
    
    // 세션 타임아웃 체크 (30분)
    const currentTime = new Date().getTime();
    const loginTime = session.loginTime || 0;
    
    if (currentTime - loginTime > SESSION_TIMEOUT) {
      // 세션 만료
      logoutUser();
      return null;
    }
    
    // 세션 시간 갱신
    session.loginTime = currentTime;
    setSession(session);
    
    return session;
    
  } catch (error) {
    Logger.log('세션 가져오기 오류: ' + error.toString());
    return null;
  }
}

/**
 * 현재 사용자 정보 가져오기
 */
function getCurrentUser() {
  const session = getSession();
  
  if (!session) {
    return {
      success: false,
      message: '로그인이 필요합니다.'
    };
  }
  
  return {
    success: true,
    user: {
      userId: session.userId,
      companyName: session.companyName,
      name: session.name,
      role: session.role
    }
  };
}

/**
 * 권한 체크
 * 권한 레벨: 관리자 > JEO > 일반
 */
function checkPermission(requiredRole) {
  const session = getSession();

  if (!session) {
    return false;
  }

  // 관리자와 JEO는 모든 권한 보유
  if (session.role === '관리자' || session.role === 'JEO') {
    return true;
  }

  // JEO 권한 체크 (추후 세부 권한 정의 예정)
  if (session.role === 'JEO' && requiredRole === 'JEO') {
    return true;
  }

  // 요청된 권한과 현재 권한 비교
  return session.role === requiredRole;
}

/**
 * 업체별 데이터 접근 권한 체크
 */
function checkCompanyAccess(companyName) {
  const session = getSession();

  if (!session) {
    return false;
  }

  // 관리자와 JEO는 모든 업체 데이터 접근 가능
  if (session.role === '관리자' || session.role === 'JEO') {
    return true;
  }

  // JEO는 현재 자기 업체만 접근 가능 (추후 변경 가능)
  if (session.role === 'JEO') {
    return session.companyName === companyName;
  }

  // 일반 사용자는 자신의 업체 데이터만 접근 가능
  return session.companyName === companyName;
}

/**
 * 사용자 등록 (회원가입)
 */
function registerUser(userData) {
  try {
    // Users 시트 존재 확인
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      logError('registerUser', 'Users 시트를 찾을 수 없습니다.');
      return {
        success: false,
        message: '시스템 오류: 사용자 데이터베이스를 찾을 수 없습니다. 관리자에게 문의하세요.'
      };
    }

    const data = userSheet.getDataRange().getDisplayValues();

    // 입력값 정규화
    const companyName = (userData.companyName || '').toString().trim();
    const name = (userData.name || '').toString().trim();
    const userId = (userData.userId || '').toString().trim();
    const password = (userData.password || '').toString().trim();

    // 입력값 검증
    if (!companyName || !name || !userId || !password) {
      return {
        success: false,
        message: '모든 항목을 입력해주세요.'
      };
    }

    // ID 길이 체크 (최소 4자)
    if (userId.length < 4) {
      return {
        success: false,
        message: 'ID는 최소 4자 이상이어야 합니다.'
      };
    }

    // 비밀번호 길이 체크 (최소 6자)
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

    // ID 중복 체크 (index 3 = ID 컬럼)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 빈 행 건너뛰기
      if (!row || row.length === 0 || !row[3]) {
        continue;
      }

      const existingId = (row[3] || '').toString().trim(); // ID 컬럼 (index 3)

      // 대소문자 구분 없이 중복 체크
      if (existingId.toLowerCase() === userId.toLowerCase()) {
        return {
          success: false,
          message: '이미 사용중인 ID입니다.'
        };
      }
    }

    // 업체 코드 결정
    let companyCode = findCompanyCodeByName(companyName);
    let isNewCompany = false;

    if (!companyCode) {
      // 새로운 업체 - 업체 코드 생성
      companyCode = generateNextCompanyCode();
      isNewCompany = true;
      Logger.log(`신규 업체 등록: ${companyName} (${companyCode})`);
    } else {
      Logger.log(`기존 업체에 사용자 추가: ${companyName} (${companyCode})`);
    }

    // 신규 업체인 경우 업체별 시트 생성
    if (isNewCompany) {
      const sheetResult = createCompanySheets(companyName, companyCode);
      if (!sheetResult.success) {
        return {
          success: false,
          message: '업체 시트 생성 중 오류가 발생했습니다. 관리자에게 문의하세요.'
        };
      }
      Logger.log(`업체 시트 생성 완료: ${JSON.stringify(sheetResult.sheets)}`);
    }

    // 사용자 추가 (새로운 Users 시트 구조)
    const timestamp = new Date();
    userSheet.appendRow([
      companyCode,   // A: 업체CODE
      companyName,   // B: 업체명
      name,          // C: 성명
      userId,        // D: ID
      password,      // E: PW
      '일반',        // F: 권한 (고정)
      '활성',        // G: 상태
      timestamp      // H: 생성일
    ]);

    Logger.log(`사용자 등록 완료: ${userId} (${companyCode} - ${companyName})`);

    return {
      success: true,
      message: '회원가입이 완료되었습니다. 로그인해주세요.',
      companyCode: companyCode,
      isNewCompany: isNewCompany
    };

  } catch (error) {
    logError('registerUser', error);

    let userMessage = '회원가입 처리 중 오류가 발생했습니다.';

    // 에러 타입별 메시지
    const errorString = error.toString().toLowerCase();
    if (errorString.includes('permission') || errorString.includes('권한')) {
      userMessage = '시스템 접근 권한이 없습니다. 관리자에게 문의하세요.';
    } else if (errorString.includes('not found') || errorString.includes('찾을 수 없')) {
      userMessage = '시스템 구성 오류입니다. 관리자에게 문의하세요.';
    }

    return {
      success: false,
      message: userMessage
    };
  }
}

/**
 * ===========================================
 * 토큰 기반 세션 관리 시스템
 * ===========================================
 */

/**
 * 세션 토큰 생성 (UUID v4)
 */
function generateSessionToken() {
  return Utilities.getUuid();
}

/**
 * 토큰 기반 세션 생성
 */
function createSession(userId, userData) {
  try {
    const token = generateSessionToken();
    const sessionData = {
      token: token,
      userId: userId,
      companyCode: userData.companyCode || '',
      companyName: userData.companyName || '',
      name: userData.name || '',
      role: userData.role || '일반',
      loginTime: new Date().getTime(),
      lastActivity: new Date().getTime()
    };

    // CacheService에 저장 (빠른 조회, 6시간 제한)
    const cache = CacheService.getScriptCache();
    cache.put('session_' + token, JSON.stringify(sessionData), 21600); // 6시간

    // ScriptProperties에도 저장 (백업, 장기 보관)
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('session_' + token, JSON.stringify(sessionData));

    Logger.log('세션 생성 완료: ' + token);

    return {
      success: true,
      token: token,
      sessionData: sessionData
    };

  } catch (error) {
    logError('createSession', error);
    return {
      success: false,
      message: '세션 생성 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 토큰으로 세션 조회
 */
function getSessionByToken(token) {
  try {
    if (!token || typeof token !== 'string') {
      return null;
    }

    const sessionKey = 'session_' + token;

    // 1. CacheService에서 먼저 조회 (빠름)
    const cache = CacheService.getScriptCache();
    let sessionString = cache.get(sessionKey);

    // 2. Cache에 없으면 ScriptProperties에서 조회
    if (!sessionString) {
      const scriptProps = PropertiesService.getScriptProperties();
      sessionString = scriptProps.getProperty(sessionKey);

      // ScriptProperties에 있으면 Cache에 복원
      if (sessionString) {
        cache.put(sessionKey, sessionString, 21600); // 6시간
      }
    }

    if (!sessionString) {
      return null;
    }

    const session = JSON.parse(sessionString);

    // 3. 세션 만료 체크 (30분)
    const currentTime = new Date().getTime();
    const lastActivity = session.lastActivity || session.loginTime;

    if (currentTime - lastActivity > SESSION_TIMEOUT) {
      // 세션 만료
      deleteSessionByToken(token);
      return null;
    }

    // 4. 마지막 활동 시간 갱신
    session.lastActivity = currentTime;
    cache.put(sessionKey, JSON.stringify(session), 21600);

    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty(sessionKey, JSON.stringify(session));

    return session;

  } catch (error) {
    logError('getSessionByToken', error);
    return null;
  }
}

/**
 * 토큰 유효성 검증
 */
function validateToken(token) {
  const session = getSessionByToken(token);
  return session !== null;
}

/**
 * 토큰 기반 로그아웃
 */
function logoutWithToken(token) {
  try {
    if (!token) {
      return {
        success: false,
        message: '토큰이 제공되지 않았습니다.'
      };
    }

    deleteSessionByToken(token);

    return {
      success: true,
      message: '로그아웃 되었습니다.'
    };
  } catch (error) {
    logError('logoutWithToken', error);
    return {
      success: false,
      message: '로그아웃 처리 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 세션 토큰 삭제
 */
function deleteSessionByToken(token) {
  try {
    if (!token) return;

    const sessionKey = 'session_' + token;

    // CacheService에서 삭제
    const cache = CacheService.getScriptCache();
    cache.remove(sessionKey);

    // ScriptProperties에서 삭제
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.deleteProperty(sessionKey);

    Logger.log('세션 삭제 완료: ' + token);

  } catch (error) {
    logError('deleteSessionByToken', error);
  }
}

/**
 * 세션 토큰 갱신
 */
function refreshSessionToken(token) {
  const session = getSessionByToken(token);

  if (!session) {
    return {
      success: false,
      message: '유효하지 않은 세션입니다.'
    };
  }

  return {
    success: true,
    session: session
  };
}

/**
 * 만료된 세션 정리 (정기적으로 실행)
 */
function cleanupExpiredSessions() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const allProps = scriptProps.getProperties();
    const currentTime = new Date().getTime();
    let cleanupCount = 0;

    for (const key in allProps) {
      if (key.startsWith('session_')) {
        try {
          const sessionData = JSON.parse(allProps[key]);
          const lastActivity = sessionData.lastActivity || sessionData.loginTime;

          // 30분 이상 지난 세션 삭제
          if (currentTime - lastActivity > SESSION_TIMEOUT) {
            const token = key.replace('session_', '');
            deleteSessionByToken(token);
            cleanupCount++;
          }
        } catch (e) {
          // 잘못된 형식의 데이터는 삭제
          scriptProps.deleteProperty(key);
          cleanupCount++;
        }
      }
    }

    Logger.log('만료된 세션 정리 완료: ' + cleanupCount + '개');

    return {
      success: true,
      cleanupCount: cleanupCount
    };

  } catch (error) {
    logError('cleanupExpiredSessions', error);
    return {
      success: false,
      message: '세션 정리 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 토큰으로 현재 사용자 정보 가져오기
 */
function getCurrentUserByToken(token) {
  const session = getSessionByToken(token);

  if (!session) {
    return {
      success: false,
      message: '로그인이 필요합니다.'
    };
  }

  return {
    success: true,
    user: {
      userId: session.userId,
      companyCode: session.companyCode,
      companyName: session.companyName,
      name: session.name,
      role: session.role
    }
  };
}

/**
 * 토큰 기반 로그인 (기존 loginUser 개선)
 */
function loginUserWithToken(userId, password) {
  try {
    // 입력값 검증 및 정규화
    if (!userId || !password) {
      return {
        success: false,
        message: 'ID와 비밀번호를 입력해주세요.'
      };
    }

    const normalizedUserId = userId.toString().trim();
    const normalizedPassword = password.toString().trim();

    if (!normalizedUserId || !normalizedPassword) {
      return {
        success: false,
        message: 'ID와 비밀번호를 입력해주세요.'
      };
    }

    // Users 시트 존재 확인
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      logError('loginUserWithToken', 'Users 시트를 찾을 수 없습니다.');
      return {
        success: false,
        message: '시스템 오류: 사용자 데이터베이스를 찾을 수 없습니다. 관리자에게 문의하세요.'
      };
    }

    const data = userSheet.getDataRange().getDisplayValues();

    // 헤더만 있는 경우 확인
    if (data.length <= 1) {
      return {
        success: false,
        message: '등록된 사용자가 없습니다. 관리자에게 문의하세요.'
      };
    }

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 빈 행 건너뛰기 (ID 컬럼은 이제 index 3)
      if (!row || row.length === 0 || !row[3] || !row[4]) {
        continue;
      }

      const [companyCode, companyName, name, id, pw, role, status] = row;

      // ID와 비밀번호 정규화 후 비교
      const normalizedId = (id || '').toString().trim();
      const normalizedPw = (pw || '').toString().trim();

      // ID와 비밀번호 확인
      if (normalizedId === normalizedUserId && normalizedPw === normalizedPassword) {
        // 상태 확인
        const normalizedStatus = (status || '').toString().trim();
        if (normalizedStatus !== '활성') {
          return {
            success: false,
            message: '비활성화된 계정입니다. 관리자에게 문의하세요.'
          };
        }

        // 토큰 기반 세션 생성
        const userData = {
          companyCode: companyCode || '',
          companyName: companyName || '',
          name: name || '',
          role: role || '일반'
        };

        const sessionResult = createSession(normalizedUserId, userData);

        if (!sessionResult.success) {
          return {
            success: false,
            message: '세션 생성 중 오류가 발생했습니다. 다시 시도해주세요.'
          };
        }

        // 자동 마이그레이션 (누락된 시트 생성)
        autoMigrateSheetsIfNeeded();

        Logger.log('토큰 로그인 성공: ' + normalizedUserId);

        return {
          success: true,
          message: '로그인 성공',
          token: sessionResult.token,
          user: {
            userId: normalizedUserId,
            companyCode: companyCode || '',
            companyName: companyName || '',
            name: name || '',
            role: role || '일반'
          }
        };
      }
    }

    return {
      success: false,
      message: 'ID 또는 비밀번호가 올바르지 않습니다.'
    };

  } catch (error) {
    logError('loginUserWithToken', error);

    let userMessage = '로그인 처리 중 오류가 발생했습니다.';

    const errorString = error.toString().toLowerCase();
    if (errorString.includes('permission') || errorString.includes('권한')) {
      userMessage = '시스템 접근 권한이 없습니다. 관리자에게 문의하세요.';
    } else if (errorString.includes('not found') || errorString.includes('찾을 수 없')) {
      userMessage = '시스템 구성 오류입니다. 관리자에게 문의하세요.';
    } else if (errorString.includes('network') || errorString.includes('timeout')) {
      userMessage = '네트워크 오류가 발생했습니다. 다시 시도해주세요.';
    }

    return {
      success: false,
      message: userMessage
    };
  }
}

/**
 * 업체명 목록 가져오기 (Users 시트에서)
 * @param {string} token - 세션 토큰
 */
function getCompanyList(token) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        companies: []
      };
    }

    // 관리자/JEO만 업체 목록 조회 가능
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.',
        companies: []
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.',
        companies: []
      };
    }

    const data = userSheet.getDataRange().getDisplayValues();

    if (data.length <= 1) {
      return {
        success: true,
        companies: []
      };
    }

    // 업체명 수집 (관리자/JEO 권한 업체 제외)
    const companySet = new Set();
    const adminJeoCompanies = new Set(); // 관리자/JEO 권한이 있는 업체

    // 먼저 관리자/JEO 권한을 가진 업체 찾기
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const companyName = (row[1] || '').toString().trim();
      const role = (row[5] || '').toString().trim(); // index 5 = 권한

      if (companyName && (role === '관리자' || role === 'JEO')) {
        adminJeoCompanies.add(companyName);
      }
    }

    // 모든 업체명 수집 (관리자/JEO 권한 업체 제외)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const companyName = (row[1] || '').toString().trim();

      if (companyName && !adminJeoCompanies.has(companyName)) {
        companySet.add(companyName);
      }
    }

    // Set을 배열로 변환하고 정렬
    const companies = Array.from(companySet).sort();

    return {
      success: true,
      companies: companies
    };

  } catch (error) {
    Logger.log('getCompanyList 오류: ' + error.toString());
    return {
      success: false,
      message: '업체 목록 조회 중 오류가 발생했습니다.',
      companies: []
    };
  }
}

/**
 * 일반 권한 업체 목록 조회 (관리자만 접근)
 */
function getNormalCompanyList(token) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        companies: []
      };
    }

    // 관리자/JEO만 업체 목록 조회 가능
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return {
        success: false,
        message: '관리자 권한이 필요합니다.',
        companies: []
      };
    }

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: 'Users 시트를 찾을 수 없습니다.',
        companies: []
      };
    }

    const data = userSheet.getDataRange().getDisplayValues();

    if (data.length <= 1) {
      return {
        success: true,
        companies: []
      };
    }

    // 일반 권한 업체명 수집 (중복 제거)
    const companySet = new Set();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const companyName = (row[1] || '').toString().trim(); // index 1 = 업체명
      const role = (row[5] || '').toString().trim(); // index 5 = 권한

      // 권한이 '일반'인 업체만 추가
      if (companyName && role === '일반') {
        companySet.add(companyName);
      }
    }

    // Set을 배열로 변환하고 정렬
    const companies = Array.from(companySet).sort();

    return {
      success: true,
      companies: companies
    };

  } catch (error) {
    Logger.log('getNormalCompanyList 오류: ' + error.toString());
    return {
      success: false,
      message: '업체 목록 조회 중 오류가 발생했습니다.',
      companies: []
    };
  }
}
