/**
 * Utils.gs - 공통 유틸리티 함수
 */

/**
 * 날짜 포맷 변환 (yyyy-MM-dd)
 */
function formatDate(date) {
  if (!date) return '';
  
  if (typeof date === 'string') {
    return date;
  }
  
  return Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd');
}

/**
 * 날짜시간 포맷 변환 (yyyy-MM-dd HH:mm:ss)
 */
function formatDateTime(datetime) {
  if (!datetime) return '';
  
  if (typeof datetime === 'string') {
    return datetime;
  }
  
  return Utilities.formatDate(datetime, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 시간 
 */
function getTimeOptions() {
  return ['오전', '오후'];
}

/**
 * 입력값 검증 - 필수값 체크
 */
function validateRequired(value, fieldName) {
  if (!value || value.toString().trim() === '') {
    return {
      valid: false,
      message: `${fieldName}은(는) 필수 입력 항목입니다.`
    };
  }
  return { valid: true };
}

/**
 * 입력값 검증 - 날짜 형식 체크
 */
function validateDate(dateString) {
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;
  
  if (!datePattern.test(dateString)) {
    return {
      valid: false,
      message: '날짜 형식이 올바르지 않습니다. (yyyy-MM-dd)'
    };
  }
  
  const date = new Date(dateString);
  if (isNaN(date.getTime())) {
    return {
      valid: false,
      message: '유효하지 않은 날짜입니다.'
    };
  }
  
  return { valid: true };
}

/**
 * 입력값 검증 - 숫자 체크
 */
function validateNumber(value, fieldName) {
  if (isNaN(value) || value === '') {
    return {
      valid: false,
      message: `${fieldName}은(는) 숫자만 입력 가능합니다.`
    };
  }
  
  if (Number(value) < 0) {
    return {
      valid: false,
      message: `${fieldName}은(는) 0 이상이어야 합니다.`
    };
  }
  
  return { valid: true };
}

/**
 * 데이터 입력 전체 검증
 */
function validateDataInput(dataObj) {
  let errors = [];
  
  // 날짜 검증
  let dateCheck = validateRequired(dataObj.date, '입고날짜');
  if (!dateCheck.valid) {
    errors.push(dateCheck.message);
  } else {
    dateCheck = validateDate(dataObj.date);
    if (!dateCheck.valid) {
      errors.push(dateCheck.message);
    }
  }
  
  // 시간 검증
  let timeCheck = validateRequired(dataObj.time, '입고시간');
  if (!timeCheck.valid) {
    errors.push(timeCheck.message);
  }
  
  // TM-NO 검증
  let tmNoCheck = validateRequired(dataObj.tmNo, 'TM-NO');
  if (!tmNoCheck.valid) {
    errors.push(tmNoCheck.message);
  }
  
  // 제품명 검증
  let productCheck = validateRequired(dataObj.productName, '제품명');
  if (!productCheck.valid) {
    errors.push(productCheck.message);
  }
  
  // 수량 검증
  let quantityCheck = validateRequired(dataObj.quantity, '수량');
  if (!quantityCheck.valid) {
    errors.push(quantityCheck.message);
  } else {
    quantityCheck = validateNumber(dataObj.quantity, '수량');
    if (!quantityCheck.valid) {
      errors.push(quantityCheck.message);
    }
  }
  
  if (errors.length > 0) {
    return {
      valid: false,
      errors: errors
    };
  }
  
  return { valid: true };
}

/**
 * HTML 특수문자 이스케이프
 */
function escapeHtml(text) {
  if (!text) return '';
  
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  
  return text.toString().replace(/[&<>"']/g, function(m) { return map[m]; });
}

/**
 * 에러 로깅
 */
function logError(functionName, error) {
  const timestamp = new Date();
  const errorMessage = `[${timestamp}] ${functionName}: ${error.toString()}`;
  Logger.log(errorMessage);
  
  // 필요시 에러 로그 시트에 기록
  try {
    const ss = getSpreadsheet();
    let errorSheet = ss.getSheetByName('ErrorLog');
    
    if (!errorSheet) {
      errorSheet = ss.insertSheet('ErrorLog');
      errorSheet.getRange('A1:C1').setValues([['날짜시간', '함수명', '에러메시지']]);
      errorSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#ea4335').setFontColor('#ffffff');
    }
    
    errorSheet.appendRow([timestamp, functionName, error.toString()]);
  } catch (e) {
    Logger.log('에러 로그 기록 실패: ' + e.toString());
  }
}

/**
 * 성공 응답 생성
 */
function createSuccessResponse(message, data = null) {
  const response = {
    success: true,
    message: message
  };
  
  if (data !== null) {
    response.data = data;
  }
  
  return response;
}

/**
 * 에러 응답 생성
 */
function createErrorResponse(message, errors = null) {
  const response = {
    success: false,
    message: message
  };
  
  if (errors !== null) {
    response.errors = errors;
  }
  
  return response;
}

/**
 * 배열을 페이지로 나누기
 */
function paginate(array, page, pageSize) {
  const startIndex = (page - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  
  return {
    data: array.slice(startIndex, endIndex),
    page: page,
    pageSize: pageSize,
    total: array.length,
    totalPages: Math.ceil(array.length / pageSize)
  };
}

/**
 * 문자열을 안전한 파일명으로 변환
 */
function sanitizeFileName(fileName) {
  if (!fileName) return 'file';
  
  // 특수문자 제거 및 공백을 언더스코어로 변경
  return fileName
    .replace(/[^a-zA-Z0-9가-힣._-]/g, '_')
    .replace(/\s+/g, '_')
    .substring(0, 100); // 길이 제한
}

/**
 * 객체 깊은 복사
 */
function deepCopy(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/**
 * 배열에서 중복 제거
 */
function removeDuplicates(array) {
  return [...new Set(array)];
}

/**
 * 날짜 범위 생성
 */
function getDateRange(startDate, endDate) {
  const dates = [];
  const current = new Date(startDate);
  const end = new Date(endDate);
  
  while (current <= end) {
    dates.push(formatDate(current));
    current.setDate(current.getDate() + 1);
  }
  
  return dates;
}

/**
 * 월의 첫날과 마지막날 구하기
 */
function getMonthRange(year, month) {
  const firstDay = new Date(year, month - 1, 1);
  const lastDay = new Date(year, month, 0);
  
  return {
    start: formatDate(firstDay),
    end: formatDate(lastDay)
  };
}

/**
 * Base64 인코딩
 */
function base64Encode(data) {
  return Utilities.base64Encode(data);
}

/**
 * Base64 디코딩
 */
function base64Decode(data) {
  return Utilities.base64Decode(data);
}

/**
 * 숫자를 천단위 콤마로 포맷
 */
function formatNumber(number) {
  if (isNaN(number)) return '0';
  return Number(number).toLocaleString('ko-KR');
}

/**
 * 파일 크기를 읽기 쉬운 형식으로 변환
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

/**
 * 랜덤 문자열 생성
 */
function generateRandomString(length = 10) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';

  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }

  return result;
}

/**
 * 업체별 시트 생성
 * @param {string} companyName - 업체명
 * @param {string} companyCode - 업체코드 (필수)
 */
function createCompanySheets(companyName, companyCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 업체코드가 없으면 업체명으로 조회
    if (!companyCode) {
      companyCode = findCompanyCodeByName(companyName);
    }

    // 업체코드가 없으면 에러
    if (!companyCode) {
      throw new Error('업체코드를 찾을 수 없습니다.');
    }

    // 마지막 시트 위치 가져오기 (맨 오른쪽에 생성하기 위함)
    const allSheets = ss.getSheets();
    const lastPosition = allSheets.length;

    // 1. Data 시트 생성 (성적서 업로드용)
    const dataSheetName = `Data_${companyCode}`;
    let dataSheet = ss.getSheetByName(dataSheetName);
    if (!dataSheet) {
      dataSheet = ss.insertSheet(dataSheetName, lastPosition);
      dataSheet.getRange('A1:L1').setValues([[
        '업체CODE', 'ID', '업체명', '입고날짜', '입고시간', 'TM-NO', '제품명', '수량', 'PDF_URL', '등록일시', '등록자', '수정일시'
      ]]);
      dataSheet.getRange('A1:L1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

      // TM-NO 열(F열, 6번째 컬럼)을 텍스트 형식으로 미리 설정
      dataSheet.getRange('F:F').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${dataSheetName}`);
    }

    // 2. List 시트 생성 (ItemList → List)
    const listSheetName = `List_${companyCode}`;
    let listSheet = ss.getSheetByName(listSheetName);
    if (!listSheet) {
      listSheet = ss.insertSheet(listSheetName, lastPosition + 1);
      listSheet.getRange('A1:F1').setValues([[
        '업체CODE', 'TM-NO', '제품명', '업체명', '검사형태', '검사기준서'
      ]]);
      listSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#6aa84f').setFontColor('#ffffff');

      // TM-NO 열(B열, 2번째 컬럼)을 텍스트 형식으로 미리 설정
      listSheet.getRange('B:B').setNumberFormat('@STRING@');
      // 검사기준서 URL 열(F열, 6번째 컬럼)을 텍스트 형식으로 미리 설정
      listSheet.getRange('F:F').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${listSheetName}`);
    }

    // 3. Spec 시트 생성 (InspectionSpec → Spec)
    const specSheetName = `Spec_${companyCode}`;
    let specSheet = ss.getSheetByName(specSheetName);
    if (!specSheet) {
      specSheet = ss.insertSheet(specSheetName, lastPosition + 2);
      specSheet.getRange('A1:I1').setValues([[
        '업체CODE', 'TM-NO', '제품명', '업체명', '검사항목', '측정방법', '규격하한', '규격상한', '시료수'
      ]]);
      specSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#e69138').setFontColor('#ffffff');

      // TM-NO 열(B열, 2번째 컬럼)을 텍스트 형식으로 미리 설정
      specSheet.getRange('B:B').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${specSheetName}`);
    }

    // 4. Result 시트 생성 (수입검사결과)
    const resultSheetName = `Result_${companyCode}`;
    let resultSheet = ss.getSheetByName(resultSheetName);
    if (!resultSheet) {
      resultSheet = ss.insertSheet(resultSheetName, lastPosition + 3);
      const headers = ['업체CODE', 'ID', '날짜', '업체명', 'TM-NO', '품명', '검사항목', '검사방법', '규격하한', '규격상한'];
      // 최대 10개 시료까지 지원
      for (let i = 1; i <= 10; i++) {
        headers.push('시료' + i);
      }
      headers.push('합부결과', '등록일시', '등록자');
      resultSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      resultSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');

      // TM-NO 열(E열, 5번째 컬럼)을 텍스트 형식으로 미리 설정
      resultSheet.getRange('E:E').setNumberFormat('@STRING@');

      Logger.log(`시트 생성 완료: ${resultSheetName}`);
    }

    return {
      success: true,
      message: `업체 "${companyName}" (${companyCode}) 시트 생성 완료`,
      sheets: {
        data: dataSheetName,
        list: listSheetName,
        spec: specSheetName,
        result: resultSheetName
      }
    };

  } catch (error) {
    logError('createCompanySheets', error);
    return {
      success: false,
      message: '시트 생성 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 시트 이름 정리 (Google Sheets 제한사항 고려)
 * @param {string} name - 원본 이름
 * @returns {string} 정리된 이름
 */
function sanitizeSheetName(name) {
  if (!name) return 'Company';

  // 특수문자 제거 (Google Sheets에서 금지된 문자: : / \ ? * [ ])
  let sanitized = name
    .replace(/[:\\/\?\*\[\]]/g, '_')
    .trim();

  // 최대 길이 30자로 제한 (Google Sheets 시트명 제한)
  if (sanitized.length > 30) {
    sanitized = sanitized.substring(0, 30);
  }

  return sanitized;
}

/**
 * 다음 업체 코드 생성
 * @returns {string} 새로운 업체 코드 (예: C01, C02, ...)
 */
function generateNextCompanyCode() {
  try {
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return 'C01'; // Users 시트가 없으면 첫 번째 코드 반환
    }

    const data = userSheet.getDataRange().getDisplayValues();

    // 헤더만 있는 경우
    if (data.length <= 1) {
      return 'C01';
    }

    let maxNumber = 0;

    // 모든 업체 코드에서 최대 번호 찾기
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const code = String(row[0] || '').trim(); // A컬럼 = 업체CODE

      // C01, C02 형식에서 숫자 부분 추출
      const match = code.match(/^C(\d+)$/);
      if (match) {
        const number = parseInt(match[1], 10);
        if (number > maxNumber) {
          maxNumber = number;
        }
      }
    }

    // 다음 번호 생성 (2자리 0 패딩)
    const nextNumber = maxNumber + 1;
    const nextCode = 'C' + String(nextNumber).padStart(2, '0');

    Logger.log(`다음 업체 코드 생성: ${nextCode}`);
    return nextCode;

  } catch (error) {
    logError('generateNextCompanyCode', error);
    return 'C01'; // 에러 시 기본값
  }
}

/**
 * 업체명으로 업체 코드 찾기
 * @param {string} companyName - 업체명
 * @returns {string|null} 업체 코드 또는 null
 */
function findCompanyCodeByName(companyName) {
  try {
    if (!companyName) return null;

    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) return null;

    const data = userSheet.getDataRange().getDisplayValues();

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const existingCompanyName = String(row[1] || '').trim(); // B컬럼 = 업체명

      if (existingCompanyName === companyName.trim()) {
        return String(row[0] || '').trim(); // A컬럼 = 업체CODE 반환
      }
    }

    return null; // 찾지 못함

  } catch (error) {
    logError('findCompanyCodeByName', error);
    return null;
  }
}

/**
 * 업체별 Data 시트 이름 생성 (성적서 업로드용)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getDataSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `Data_${sanitizeSheetName(companyName)}`;
  }
  return `Data_${companyCode}`;
}

/**
 * 업체별 List 시트 이름 생성 (구 ItemList)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getItemListSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `List_${sanitizeSheetName(companyName)}`;
  }
  return `List_${companyCode}`;
}

/**
 * 업체별 Spec 시트 이름 생성 (구 InspectionSpec)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getInspectionSpecSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `Spec_${sanitizeSheetName(companyName)}`;
  }
  return `Spec_${companyCode}`;
}

/**
 * 업체별 Result 시트 이름 생성 (수입검사결과)
 * @param {string} companyName - 업체명
 * @returns {string} 시트 이름
 */
function getResultSheetName(companyName) {
  const companyCode = findCompanyCodeByName(companyName);
  if (!companyCode) {
    Logger.log(`경고: ${companyName}의 업체코드를 찾을 수 없습니다.`);
    return `Result_${sanitizeSheetName(companyName)}`;
  }
  return `Result_${companyCode}`;
}

/**
 * 모든 업체명 목록 가져오기 (Users 시트에서)
 * @returns {Array<string>} 업체명 목록
 */
function getAllCompanyNames() {
  try {
    const userSheet = getSheet(USER_SHEET_NAME);
    if (!userSheet) {
      return [];
    }

    const data = userSheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      return [];
    }

    const companySet = new Set();

    // 헤더 제외하고 업체명 수집
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const companyName = String(row[1] || '').trim(); // B컬럼 = 업체명

      if (companyName) {
        companySet.add(companyName);
      }
    }

    return Array.from(companySet).sort();

  } catch (error) {
    logError('getAllCompanyNames', error);
    return [];
  }
}

/**
 * 업체 시트명 변경 (업체명이 변경되었을 때)
 * 주의: 현재 시스템은 업체코드 기반 시트명을 사용하므로 (Data_C01, List_C01 등)
 * 업체명이 변경되어도 시트명은 변경하지 않음
 *
 * @param {string} oldCompanyName - 기존 업체명
 * @param {string} newCompanyName - 새 업체명
 * @returns {Object} {success: boolean, message: string}
 */
function renameCompanySheets(oldCompanyName, newCompanyName) {
  try {
    // 현재 시스템은 업체코드 기반 시트명을 사용하므로
    // 업체명이 변경되어도 시트명은 변경하지 않음
    Logger.log(`업체명 변경: ${oldCompanyName} → ${newCompanyName} (시트명은 업체코드 기반이므로 변경 불필요)`);

    return {
      success: true,
      message: '시트명은 업체코드 기반이므로 변경이 불필요합니다.'
    };

  } catch (error) {
    logError('renameCompanySheets', error);
    return {
      success: false,
      message: '시트명 변경 중 오류가 발생했습니다.'
    };
  }
}
