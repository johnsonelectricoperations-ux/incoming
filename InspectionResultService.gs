/**
 * InspectionResultService.gs - 검사결과 CRUD 처리 (업체별 동적 시트)
 */

/**
 * 업체별 Result 시트 가져오기 또는 생성
 * @param {string} companyName - 업체명
 * @returns {Object} {success, sheet, message}
 */
function getOrCreateResultSheet(companyName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getResultSheetName(companyName);
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // JEO본사(관리자/JEO 권한)는 시트 생성 생략
      if (companyName === 'JEO본사') {
        return {
          success: false,
          message: '관리자/JEO 권한은 별도 시트가 필요하지 않습니다.'
        };
      }

      // 시트가 없으면 생성
      const result = createCompanySheets(companyName);
      if (!result.success) {
        return {
          success: false,
          message: '시트 생성에 실패했습니다.'
        };
      }
      sheet = ss.getSheetByName(sheetName);
    } else {
      // 기존 시트가 있는 경우, 헤더 확인 및 업데이트
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // '입고ID' 열이 없는지 확인 (이전 버전 시트 체크)
      if (headerRow[0] === '업체CODE' && headerRow[1] !== '입고ID') {
        Logger.log(`${sheetName}: 이전 버전 헤더 감지, '입고ID' 열 추가`);

        // B열(2번째 열)에 '입고ID' 컬럼 삽입
        sheet.insertColumnBefore(2);

        // 헤더 설정
        sheet.getRange(1, 2).setValue('입고ID');
        sheet.getRange(1, 2).setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');

        // 입고ID 열을 텍스트 형식으로 설정
        sheet.getRange('B:B').setNumberFormat('@STRING@');

        Logger.log(`${sheetName}: '입고ID' 열 추가 완료`);
      }
    }

    return {
      success: true,
      sheet: sheet
    };

  } catch (error) {
    logError('getOrCreateResultSheet', error);
    return {
      success: false,
      message: '시트 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 검사결과 저장 (배치)
 */
function saveInspectionResults(token, dataId, results) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 업체별 Data 시트에서 입고 정보 조회
    let dataInfo = null;
    let companiesToQuery = [];

    // 조회할 업체 목록 결정
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    // 각 업체별 Data 시트에서 ID 검색
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`saveInspectionResults: ${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const dataValues = dataSheet.getDataRange().getValues();

        for (let i = 1; i < dataValues.length; i++) {
          if (String(dataValues[i][1]) === String(dataId)) { // row[1]이 ID (업체CODE가 추가되어 인덱스 변경)
            let dateValue = dataValues[i][3]; // 날짜는 3번 인덱스
            if (dateValue instanceof Date) {
              dateValue = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
            } else if (dateValue) {
              dateValue = String(dateValue).trim();
            }

            dataInfo = {
              date: dateValue,
              companyName: String(dataValues[i][2]),
              tmNo: String(dataValues[i][5]),
              productName: String(dataValues[i][6])
            };
            break;
          }
        }

        if (dataInfo) {
          break; // 데이터를 찾았으면 루프 종료
        }
      } catch (e) {
        Logger.log(`saveInspectionResults: ${companyName} Data 시트 조회 오류 - ${e.message}`);
        continue;
      }
    }

    if (!dataInfo) {
      return { success: false, message: '입고 데이터를 찾을 수 없습니다.' };
    }

    // 일반 사용자는 자기 업체 데이터만 저장 가능
    if (session.role !== '관리자' && session.role !== 'JEO' && dataInfo.companyName !== session.companyName) {
      return {
        success: false,
        message: '다른 업체의 검사결과를 저장할 권한이 없습니다.'
      };
    }

    // 해당 업체의 Result 시트 가져오기
    const sheetResult = getOrCreateResultSheet(dataInfo.companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || '수입검사결과 시트를 찾을 수 없습니다.'
      };
    }

    const resultSheet = sheetResult.sheet;

    // 업체코드 조회
    const companyCode = findCompanyCodeByName(dataInfo.companyName);
    if (!companyCode) {
      return {
        success: false,
        message: '업체코드를 찾을 수 없습니다.'
      };
    }

    // 날짜 형식 변환
    let dateStr = dataInfo.date;
    if (dataInfo.date instanceof Date) {
      dateStr = Utilities.formatDate(dataInfo.date, 'Asia/Seoul', 'yyyy-MM-dd');
    } else if (dateStr) {
      dateStr = String(dateStr).trim();
    }

    const timestamp = new Date();

    // 먼저 해당 입고ID의 기존 검사결과가 있는지 확인하고 삭제
    const existingData = resultSheet.getDataRange().getValues();
    const rowsToDelete = [];

    for (let i = existingData.length - 1; i >= 1; i--) {
      // 입고ID(row[1])로 매칭
      if (String(existingData[i][1]) === String(dataId)) {
        rowsToDelete.push(i + 1); // 1-based index
      }
    }

    // 역순으로 삭제 (뒤에서부터 삭제해야 인덱스가 안 깨짐)
    if (rowsToDelete.length > 0) {
      Logger.log(`기존 검사결과 ${rowsToDelete.length}건 삭제 중...`);
      for (const rowIndex of rowsToDelete) {
        resultSheet.deleteRow(rowIndex);
      }
    }

    // 각 검사항목별로 행 추가
    results.forEach(function(result) {
      const id = 'IR' + timestamp.getTime() + '_' + Math.random().toString(36).substr(2, 9);

      // 기본 정보 (업체CODE, 입고ID 포함)
      const row = [
        companyCode,
        dataId,              // 입고ID 추가
        id,
        dateStr,
        dataInfo.companyName,
        dataInfo.tmNo,
        dataInfo.productName,
        result.inspectionItem,
        result.measurementMethod || '',
        result.lowerLimit || '',
        result.upperLimit || ''
      ];

      // 시료 측정값 (최대 10개)
      for (let i = 0; i < 10; i++) {
        row.push(result.samples[i] || '');
      }

      // 합부결과, 등록일시, 등록자
      row.push(result.passFailResult, timestamp, session.name);

      const lastRow = resultSheet.getLastRow() + 1;

      // 먼저 텍스트 형식으로 설정 (특히 입고ID, TM-NO 컬럼)
      resultSheet.getRange(lastRow, 1, 1, row.length).setNumberFormat('@STRING@');

      // 데이터 입력
      resultSheet.getRange(lastRow, 1, 1, row.length).setValues([row]);
    });

    Logger.log('검사결과 저장 완료 - 업체: ' + dataInfo.companyName);

    return {
      success: true,
      message: '검사결과가 저장되었습니다.'
    };

  } catch (error) {
    Logger.log('검사결과 저장 오류: ' + error.toString());
    return {
      success: false,
      message: '검사결과 저장 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 검사결과 조회 (dataId로)
 */
function getInspectionResultsByDataId(token, dataId) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.', data: [] };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 업체별 Data 시트에서 입고 정보 조회
    let dataInfo = null;
    let companiesToQuery = [];

    // 조회할 업체 목록 결정
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    // 각 업체별 Data 시트에서 ID 검색
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`getInspectionResultsByDataId: ${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const dataValues = dataSheet.getDataRange().getValues();

        for (let i = 1; i < dataValues.length; i++) {
          if (String(dataValues[i][1]) === String(dataId)) { // row[1]이 ID
            let dateValue = dataValues[i][3]; // 날짜는 3번 인덱스
            if (dateValue instanceof Date) {
              dateValue = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
            } else if (dateValue) {
              dateValue = String(dateValue).trim();
            }

            dataInfo = {
              date: dateValue,
              companyName: String(dataValues[i][2]),
              tmNo: String(dataValues[i][5]),
              productName: String(dataValues[i][6])
            };
            break;
          }
        }

        if (dataInfo) {
          break; // 데이터를 찾았으면 루프 종료
        }
      } catch (e) {
        Logger.log(`getInspectionResultsByDataId: ${companyName} Data 시트 조회 오류 - ${e.message}`);
        continue;
      }
    }

    if (!dataInfo) {
      return { success: false, message: '입고 데이터를 찾을 수 없습니다.', data: [] };
    }

    // 일반 사용자는 자기 업체 데이터만 조회 가능
    if (session.role !== '관리자' && session.role !== 'JEO' && dataInfo.companyName !== session.companyName) {
      return {
        success: false,
        message: '다른 업체의 검사결과를 조회할 권한이 없습니다.',
        data: []
      };
    }

    // 해당 업체의 Result 시트 가져오기
    const sheetResult = getOrCreateResultSheet(dataInfo.companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || '수입검사결과 시트를 찾을 수 없습니다.',
        data: []
      };
    }

    const resultSheet = sheetResult.sheet;
    const resultData = resultSheet.getDataRange().getValues();

    // 날짜 형식 변환
    let dateStr = dataInfo.date;
    if (dataInfo.date instanceof Date) {
      dateStr = Utilities.formatDate(dataInfo.date, 'Asia/Seoul', 'yyyy-MM-dd');
    } else if (dateStr) {
      dateStr = String(dateStr).trim();
    }

    // 입고ID로 직접 검색
    const results = [];

    Logger.log('=== getInspectionResultsByDataId 검색 시작 ===');
    Logger.log('dataId: ' + dataId);
    Logger.log('dataInfo: ' + JSON.stringify(dataInfo));
    Logger.log('resultData.length: ' + resultData.length);

    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];

      // 입고ID(row[1])로 직접 매칭
      if (String(row[1]) === String(dataId)) {
        Logger.log('매칭 성공! row: ' + i);

        // 날짜 형식 정규화
        let rowDateStr = row[3];
        if (row[3] instanceof Date) {
          rowDateStr = Utilities.formatDate(row[3], 'Asia/Seoul', 'yyyy-MM-dd');
        } else if (rowDateStr) {
          rowDateStr = String(rowDateStr).trim();
        }

        // 시료 데이터 추출 (컬럼 인덱스: 입고ID 추가로 +1)
        const samples = [];
        for (let j = 11; j < 21; j++) {
          // 값이 있으면 문자열로 변환, 없으면 빈 문자열
          const value = row[j];
          if (value !== null && value !== undefined && value !== '') {
            samples.push(String(value));
          } else {
            samples.push('');
          }
        }

        // registeredAt 날짜 변환
        let registeredAtStr = '';
        if (row[22] instanceof Date) {
          registeredAtStr = Utilities.formatDate(row[22], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        } else if (row[22]) {
          registeredAtStr = String(row[22]);
        }

        results.push({
          id: String(row[2] || ''),
          date: rowDateStr,
          companyName: String(row[4] || ''),
          tmNo: String(row[5] || ''),
          productName: String(row[6] || ''),
          inspectionItem: String(row[7] || ''),
          measurementMethod: String(row[8] || ''),
          lowerLimit: String(row[9] || ''),
          upperLimit: String(row[10] || ''),
          samples: samples,
          passFailResult: String(row[21] || ''),
          registeredAt: registeredAtStr,
          registeredBy: String(row[23] || '')
        });
      }
    }

    Logger.log('검사결과 조회 완료 - 건수: ' + results.length);

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('검사결과 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '검사결과 조회 중 오류가 발생했습니다: ' + error.message,
      data: []
    };
  }
}

/**
 * 모든 검사결과 키 조회 (date|companyName|tmNo 형식)
 */
function getAllInspectionResultKeys(token) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.', keys: [] };
    }

    const keysSet = new Set();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    // 각 업체 Result 시트에서 키 수집
    for (const companyName of companiesToQuery) {
      const sheetResult = getOrCreateResultSheet(companyName);
      if (!sheetResult.success) {
        continue;
      }

      const sheet = sheetResult.sheet;
      const data = sheet.getDataRange().getValues();

      // 헤더 제외하고 처리
      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // 날짜 형식 정규화 (컬럼 인덱스: 업체CODE(0), 입고ID(1), ID(2), 날짜(3), 업체명(4), TM-NO(5))
        let dateStr = row[3];
        if (row[3] instanceof Date) {
          dateStr = Utilities.formatDate(row[3], 'Asia/Seoul', 'yyyy-MM-dd');
        } else if (dateStr) {
          dateStr = String(dateStr).trim();
        }

        const rowCompanyName = String(row[4] || '');
        const tmNo = String(row[5] || '');

        if (dateStr && rowCompanyName && tmNo) {
          const key = dateStr + '|' + rowCompanyName + '|' + tmNo;
          keysSet.add(key);
        }
      }
    }

    const keys = Array.from(keysSet);

    Logger.log('검사결과 키 조회 완료 - 키 개수: ' + keys.length);

    return {
      success: true,
      keys: keys
    };

  } catch (error) {
    Logger.log('검사결과 키 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '검사결과 키 조회 중 오류가 발생했습니다.',
      keys: []
    };
  }
}

/**
 * 검사결과 이력 검색 (업체명/시작일자/종료일자/TM-NO)
 */
function searchInspectionResultHistory(token, filters) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.', data: [] };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];

    // 필터 파라미터 추출
    const filterCompanyName = filters.companyName || '';
    const filterDateFrom = filters.dateFrom || '';
    const filterDateTo = filters.dateTo || '';
    const filterTmNo = filters.tmNo || '';
    const filterInspectionType = filters.inspectionType || '';

    Logger.log('=== searchInspectionResultHistory 시작 ===');
    Logger.log('필터: ' + JSON.stringify(filters));

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    // 업체명 필터가 있으면 해당 업체만 조회
    if (filterCompanyName) {
      if (companiesToQuery.includes(filterCompanyName)) {
        companiesToQuery = [filterCompanyName];
      } else {
        // 권한이 없는 업체를 조회하려는 경우
        return {
          success: false,
          message: '해당 업체의 데이터를 조회할 권한이 없습니다.',
          data: []
        };
      }
    }

    Logger.log('조회 대상 업체: ' + companiesToQuery.join(', '));

    // 각 업체별로 Data 시트와 Result 시트 조회
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const dataValues = dataSheet.getDataRange().getValues();

        // 헤더 제외하고 처리
        for (let i = 1; i < dataValues.length; i++) {
          const row = dataValues[i];

          // 날짜 형식 정규화 (row[3]이 날짜)
          let dateStr = row[3];
          if (row[3] instanceof Date) {
            dateStr = Utilities.formatDate(row[3], 'Asia/Seoul', 'yyyy-MM-dd');
          } else if (dateStr) {
            dateStr = String(dateStr).trim();
          }

          const rowCompanyName = String(row[2] || '');
          const tmNo = String(row[5] || '');
          const productName = String(row[6] || '');
          const pdfUrl = String(row[8] || '');

          // 날짜 범위 필터 적용
          if (filterDateFrom && dateStr < filterDateFrom) {
            continue;
          }
          if (filterDateTo && dateStr > filterDateTo) {
            continue;
          }

          // TM-NO 필터 적용 (부분 일치)
          if (filterTmNo && tmNo.indexOf(filterTmNo) === -1) {
            continue;
          }

          // ItemList에서 검사형태 조회
          let inspectionType = '';
          const itemListSheetName = getItemListSheetName(rowCompanyName);
          const itemListSheet = ss.getSheetByName(itemListSheetName);

          if (itemListSheet) {
            try {
              const itemData = itemListSheet.getDataRange().getValues();
              for (let j = 1; j < itemData.length; j++) {
                if (String(itemData[j][1]) === tmNo) {
                  inspectionType = String(itemData[j][4] || '검사');
                  break;
                }
              }
            } catch (e) {
              Logger.log(`ItemList 조회 오류 (${rowCompanyName}, ${tmNo}): ${e.message}`);
            }
          }

          // 검사형태 필터 적용
          if (filterInspectionType && inspectionType !== filterInspectionType) {
            continue;
          }

          // 검사결과 존재 여부 및 합부판정 확인
          const resultKey = dateStr + '|' + rowCompanyName + '|' + tmNo;
          const inspectionResults = checkInspectionResults(companyName, resultKey);

          results.push({
            companyName: rowCompanyName,
            date: dateStr,
            tmNo: tmNo,
            productName: productName,
            pdfUrl: pdfUrl,
            hasInspectionResult: inspectionResults.exists,
            overallPassFail: inspectionResults.overallPassFail,
            inspectionType: inspectionType || '검사'
          });
        }

      } catch (e) {
        Logger.log(`${companyName} Data 시트 조회 오류 - ${e.message}`);
        continue;
      }
    }

    Logger.log('검색 완료 - 총 ' + results.length + '건');

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('검사결과 이력 검색 오류: ' + error.toString());
    return {
      success: false,
      message: '검사결과 이력 검색 중 오류가 발생했습니다: ' + error.message,
      data: []
    };
  }
}

/**
 * 검사결과 존재 여부 및 전체 합부판정 확인
 * @param {string} companyName - 업체명
 * @param {string} resultKey - 검색 키 (date|companyName|tmNo)
 * @returns {Object} {exists: boolean, overallPassFail: string}
 */
function checkInspectionResults(companyName, resultKey) {
  try {
    const sheetResult = getOrCreateResultSheet(companyName);
    if (!sheetResult.success) {
      return { exists: false, overallPassFail: '' };
    }

    const resultSheet = sheetResult.sheet;
    const resultData = resultSheet.getDataRange().getValues();

    let passCount = 0;
    let failCount = 0;
    let totalCount = 0;

    // 헤더 제외하고 처리
    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];

      // 날짜 형식 정규화 (row[3]이 날짜, 입고ID 추가로 +1)
      let rowDateStr = row[3];
      if (row[3] instanceof Date) {
        rowDateStr = Utilities.formatDate(row[3], 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (rowDateStr) {
        rowDateStr = String(rowDateStr).trim();
      }

      const rowKey = rowDateStr + '|' + String(row[4]) + '|' + String(row[5]);

      if (rowKey === resultKey) {
        totalCount++;
        const passFailResult = String(row[21] || '').trim();

        if (passFailResult === '합격') {
          passCount++;
        } else if (passFailResult === '불합격') {
          failCount++;
        }
      }
    }

    // 검사결과가 없으면
    if (totalCount === 0) {
      return { exists: false, overallPassFail: '' };
    }

    // 전체 합부판정 계산 (하나라도 불합격이면 전체 불합격)
    let overallPassFail = '';
    if (failCount > 0) {
      overallPassFail = '불합격';
    } else if (passCount === totalCount) {
      overallPassFail = '합격';
    }

    return {
      exists: true,
      overallPassFail: overallPassFail
    };

  } catch (error) {
    Logger.log('checkInspectionResults 오류: ' + error.toString());
    return { exists: false, overallPassFail: '' };
  }
}

/**
 * 검사결과 조회 (resultKey로 직접 조회)
 * @param {string} token - 세션 토큰
 * @param {string} resultKey - 검색 키 (date|companyName|tmNo)
 * @param {string} companyName - 업체명
 * @returns {Object} {success, data, message}
 */
function getInspectionResultsByKey(token, resultKey, companyName) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.', data: [] };
    }

    // 일반 사용자는 자기 업체 데이터만 조회 가능
    if (session.role !== '관리자' && session.role !== 'JEO' && companyName !== session.companyName) {
      return {
        success: false,
        message: '다른 업체의 검사결과를 조회할 권한이 없습니다.',
        data: []
      };
    }

    // 해당 업체의 Result 시트 가져오기
    const sheetResult = getOrCreateResultSheet(companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || '수입검사결과 시트를 찾을 수 없습니다.',
        data: []
      };
    }

    const resultSheet = sheetResult.sheet;
    const resultData = resultSheet.getDataRange().getValues();
    const results = [];

    Logger.log('=== getInspectionResultsByKey 검색 시작 ===');
    Logger.log('resultKey: ' + resultKey);
    Logger.log('companyName: ' + companyName);

    // 검색 키로 매칭
    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];

      // 날짜 형식 정규화 (row[3]이 날짜, 입고ID 추가로 +1)
      let rowDateStr = row[3];
      if (row[3] instanceof Date) {
        rowDateStr = Utilities.formatDate(row[3], 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (rowDateStr) {
        rowDateStr = String(rowDateStr).trim();
      }

      const rowKey = rowDateStr + '|' + String(row[4]) + '|' + String(row[5]);

      if (rowKey === resultKey) {
        // 시료 데이터 추출 (row[11-20], 입고ID 추가로 +1)
        const samples = [];
        for (let j = 11; j < 21; j++) {
          const value = row[j];
          if (value !== null && value !== undefined && value !== '') {
            samples.push(String(value));
          } else {
            samples.push('');
          }
        }

        // registeredAt 날짜 변환 (row[22], 입고ID 추가로 +1)
        let registeredAtStr = '';
        if (row[22] instanceof Date) {
          registeredAtStr = Utilities.formatDate(row[22], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        } else if (row[22]) {
          registeredAtStr = String(row[22]);
        }

        results.push({
          id: String(row[2] || ''),
          date: rowDateStr,
          companyName: String(row[4] || ''),
          tmNo: String(row[5] || ''),
          productName: String(row[6] || ''),
          inspectionItem: String(row[7] || ''),
          measurementMethod: String(row[8] || ''),
          lowerLimit: String(row[9] || ''),
          upperLimit: String(row[10] || ''),
          samples: samples,
          passFailResult: String(row[21] || ''),
          registeredAt: registeredAtStr,
          registeredBy: String(row[23] || '')
        });
      }
    }

    Logger.log('검사결과 조회 완료 - 건수: ' + results.length);

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('검사결과 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '검사결과 조회 중 오류가 발생했습니다: ' + error.message,
      data: []
    };
  }
}

/**
 * 이전 검사결과 조회 (자동입력용)
 * @param {string} token - 세션 토큰
 * @param {string} companyName - 업체명
 * @param {string} tmNo - TM-NO
 * @returns {Object} {success, data, message}
 */
function getPreviousInspectionData(token, companyName, tmNo) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.', data: [] };
    }

    // 권한 체크: 관리자 또는 JEO만 접근 가능
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return {
        success: false,
        message: '검사결과 조회 권한이 없습니다.',
        data: []
      };
    }

    // 해당 업체의 Result 시트 가져오기
    const sheetResult = getOrCreateResultSheet(companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || '수입검사결과 시트를 찾을 수 없습니다.',
        data: []
      };
    }

    const resultSheet = sheetResult.sheet;
    const resultData = resultSheet.getDataRange().getValues();

    // TM-NO가 일치하는 모든 데이터 수집 (날짜 내림차순)
    const matchingData = [];

    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];
      const rowCompanyName = String(row[4] || '');
      const rowTmNo = String(row[5] || '');

      if (rowCompanyName === companyName && rowTmNo === tmNo) {
        // 날짜 형식 정규화 (row[3]이 날짜, 입고ID 추가로 +1)
        let rowDateStr = row[3];
        if (row[3] instanceof Date) {
          rowDateStr = Utilities.formatDate(row[3], 'Asia/Seoul', 'yyyy-MM-dd');
        } else if (rowDateStr) {
          rowDateStr = String(rowDateStr).trim();
        }

        // 시료 데이터 추출 (row[11-20], 입고ID 추가로 +1)
        const samples = [];
        for (let j = 11; j < 21; j++) {
          const value = row[j];
          if (value !== null && value !== undefined && value !== '') {
            samples.push(parseFloat(value));
          }
        }

        // 유효한 시료가 있는 경우만 추가
        if (samples.length > 0) {
          matchingData.push({
            date: rowDateStr,
            inspectionItem: String(row[7] || ''),
            samples: samples
          });
        }
      }
    }

    if (matchingData.length === 0) {
      return {
        success: false,
        message: '이전 검사결과 이력이 없습니다.',
        data: []
      };
    }

    Logger.log('이전 검사결과 조회 완료 - 건수: ' + matchingData.length);

    return {
      success: true,
      data: matchingData
    };

  } catch (error) {
    Logger.log('이전 검사결과 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '이전 검사결과 조회 중 오류가 발생했습니다: ' + error.message,
      data: []
    };
  }
}
