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
        const dataValues = dataSheet.getDataRange().getDisplayValues();

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

    // 각 검사항목별로 행 추가
    results.forEach(function(result) {
      const id = 'IR' + timestamp.getTime() + '_' + Math.random().toString(36).substr(2, 9);

      // 기본 정보 (업체CODE 포함)
      const row = [
        companyCode,
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

      // 먼저 텍스트 형식으로 설정 (특히 TM-NO 컬럼)
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
        const dataValues = dataSheet.getDataRange().getDisplayValues();

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
    const resultData = resultSheet.getDataRange().getDisplayValues();

    // 날짜 형식 변환
    let dateStr = dataInfo.date;
    if (dataInfo.date instanceof Date) {
      dateStr = Utilities.formatDate(dataInfo.date, 'Asia/Seoul', 'yyyy-MM-dd');
    } else if (dateStr) {
      dateStr = String(dateStr).trim();
    }

    // dataId 대신 date|companyName|tmNo 조합으로 검색
    const searchKey = dateStr + '|' + dataInfo.companyName + '|' + dataInfo.tmNo;
    const results = [];

    Logger.log('=== getInspectionResultsByDataId 검색 시작 ===');
    Logger.log('dataId: ' + dataId);
    Logger.log('dataInfo: ' + JSON.stringify(dataInfo));
    Logger.log('searchKey: ' + searchKey);
    Logger.log('resultData.length: ' + resultData.length);

    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];

      // 날짜 형식 정규화 (컬럼 인덱스 수정: 업체CODE 추가로 +1)
      let rowDateStr = row[2];
      if (row[2] instanceof Date) {
        rowDateStr = Utilities.formatDate(row[2], 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (rowDateStr) {
        rowDateStr = String(rowDateStr).trim();
      }

      const rowKey = rowDateStr + '|' + String(row[3]) + '|' + String(row[4]);

      if (i <= 3) {
        Logger.log(`row[${i}] key: ${rowKey}`);
      }

      if (rowKey === searchKey) {
        Logger.log('매칭 성공! row: ' + i);
        // 시료 데이터 추출 (컬럼 인덱스 수정: 업체CODE 추가로 +1)
        const samples = [];
        for (let j = 10; j < 20; j++) {
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
        if (row[21] instanceof Date) {
          registeredAtStr = Utilities.formatDate(row[21], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        } else if (row[21]) {
          registeredAtStr = String(row[21]);
        }

        results.push({
          id: String(row[1] || ''),
          date: rowDateStr,
          companyName: String(row[3] || ''),
          tmNo: String(row[4] || ''),
          productName: String(row[5] || ''),
          inspectionItem: String(row[6] || ''),
          measurementMethod: String(row[7] || ''),
          lowerLimit: String(row[8] || ''),
          upperLimit: String(row[9] || ''),
          samples: samples,
          passFailResult: String(row[20] || ''),
          registeredAt: registeredAtStr,
          registeredBy: String(row[22] || '')
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
      const data = sheet.getDataRange().getDisplayValues();

      // 헤더 제외하고 처리
      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // 날짜 형식 정규화 (컬럼 인덱스 수정: 업체CODE 추가로 +1)
        let dateStr = row[2];
        if (row[2] instanceof Date) {
          dateStr = Utilities.formatDate(row[2], 'Asia/Seoul', 'yyyy-MM-dd');
        } else if (dateStr) {
          dateStr = String(dateStr).trim();
        }

        const rowCompanyName = String(row[3] || '');
        const tmNo = String(row[4] || '');

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
 * 최적화: ItemList와 Result 데이터를 사전에 캐싱하여 중복 조회 제거
 */
function searchInspectionResultHistory(token, filters) {
  const startTime = new Date().getTime();

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

    // 최적화: ItemList 데이터를 사전에 모두 로드하여 Map으로 캐싱
    const itemListStartTime = new Date().getTime();
    const itemInspectionTypeMap = {};
    for (const companyName of companiesToQuery) {
      const itemListSheetName = getItemListSheetName(companyName);
      const itemListSheet = ss.getSheetByName(itemListSheetName);

      if (itemListSheet) {
        try {
          const itemData = itemListSheet.getDataRange().getDisplayValues();
          for (let j = 1; j < itemData.length; j++) {
            const tmNo = String(itemData[j][1] || '');
            const key = companyName + '|' + tmNo;
            itemInspectionTypeMap[key] = String(itemData[j][4] || '검사');
          }
        } catch (e) {
          Logger.log(`ItemList 로드 오류 (${companyName}): ${e.message}`);
        }
      }
    }
    Logger.log(`ItemList 캐싱 완료 (${new Date().getTime() - itemListStartTime}ms)`);

    // 최적화: Result 데이터를 사전에 모두 로드하여 Map으로 캐싱
    const resultStartTime = new Date().getTime();
    const resultMap = {}; // key: resultKey, value: {exists, passCount, failCount}
    for (const companyName of companiesToQuery) {
      const sheetResult = getOrCreateResultSheet(companyName);
      if (!sheetResult.success) {
        continue;
      }

      try {
        const resultSheet = sheetResult.sheet;
        const resultData = resultSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < resultData.length; i++) {
          const row = resultData[i];

          // 날짜 형식 정규화 (row[2]가 날짜)
          let rowDateStr = row[2];
          if (row[2] instanceof Date) {
            rowDateStr = Utilities.formatDate(row[2], 'Asia/Seoul', 'yyyy-MM-dd');
          } else if (rowDateStr) {
            rowDateStr = String(rowDateStr).trim();
          }

          const rowKey = rowDateStr + '|' + String(row[3]) + '|' + String(row[4]);
          const passFailResult = String(row[20] || '').trim();

          if (!resultMap[rowKey]) {
            resultMap[rowKey] = { exists: true, passCount: 0, failCount: 0 };
          }

          if (passFailResult === '합격') {
            resultMap[rowKey].passCount++;
          } else if (passFailResult === '불합격') {
            resultMap[rowKey].failCount++;
          }
        }
      } catch (e) {
        Logger.log(`Result 로드 오류 (${companyName}): ${e.message}`);
      }
    }
    Logger.log(`Result 캐싱 완료 (${new Date().getTime() - resultStartTime}ms)`);

    // 각 업체별로 Data 시트 조회
    const dataStartTime = new Date().getTime();
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const dataValues = dataSheet.getDataRange().getDisplayValues();

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
          const quantity = Number(row[7]) || 0;
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

          // ItemList에서 검사형태 조회 (캐시 사용)
          const itemKey = rowCompanyName + '|' + tmNo;
          const inspectionType = itemInspectionTypeMap[itemKey] || '검사';

          // 검사형태 필터 적용
          if (filterInspectionType && inspectionType !== filterInspectionType) {
            continue;
          }

          // 검사결과 존재 여부 및 합부판정 확인 (캐시 사용)
          const resultKey = dateStr + '|' + rowCompanyName + '|' + tmNo;
          const resultInfo = resultMap[resultKey] || { exists: false, passCount: 0, failCount: 0 };

          let overallPassFail = '';
          if (resultInfo.exists) {
            if (resultInfo.failCount > 0) {
              overallPassFail = '불합격';
            } else if (resultInfo.passCount > 0) {
              overallPassFail = '합격';
            }
          }

          results.push({
            companyName: rowCompanyName,
            date: dateStr,
            tmNo: tmNo,
            productName: productName,
            quantity: quantity,
            pdfUrl: pdfUrl,
            hasInspectionResult: resultInfo.exists,
            overallPassFail: overallPassFail,
            inspectionType: inspectionType
          });
        }

      } catch (e) {
        Logger.log(`${companyName} Data 시트 조회 오류 - ${e.message}`);
        continue;
      }
    }
    Logger.log(`Data 조회 완료 (${new Date().getTime() - dataStartTime}ms)`);

    const totalTime = new Date().getTime() - startTime;
    Logger.log(`검색 완료 - 총 ${results.length}건 (${totalTime}ms)`);

    return {
      success: true,
      data: results
    };

  } catch (error) {
    const totalTime = new Date().getTime() - startTime;
    Logger.log(`검사결과 이력 검색 오류 (${totalTime}ms): ${error.toString()}`);
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
    const resultData = resultSheet.getDataRange().getDisplayValues();

    let passCount = 0;
    let failCount = 0;
    let totalCount = 0;

    // 헤더 제외하고 처리
    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];

      // 날짜 형식 정규화 (row[2]가 날짜)
      let rowDateStr = row[2];
      if (row[2] instanceof Date) {
        rowDateStr = Utilities.formatDate(row[2], 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (rowDateStr) {
        rowDateStr = String(rowDateStr).trim();
      }

      const rowKey = rowDateStr + '|' + String(row[3]) + '|' + String(row[4]);

      if (rowKey === resultKey) {
        totalCount++;
        const passFailResult = String(row[20] || '').trim();

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
    const resultData = resultSheet.getDataRange().getDisplayValues();
    const results = [];

    Logger.log('=== getInspectionResultsByKey 검색 시작 ===');
    Logger.log('resultKey: ' + resultKey);
    Logger.log('companyName: ' + companyName);

    // 검색 키로 매칭
    for (let i = 1; i < resultData.length; i++) {
      const row = resultData[i];

      // 날짜 형식 정규화 (row[2]가 날짜)
      let rowDateStr = row[2];
      if (row[2] instanceof Date) {
        rowDateStr = Utilities.formatDate(row[2], 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (rowDateStr) {
        rowDateStr = String(rowDateStr).trim();
      }

      const rowKey = rowDateStr + '|' + String(row[3]) + '|' + String(row[4]);

      if (rowKey === resultKey) {
        // 시료 데이터 추출
        const samples = [];
        for (let j = 10; j < 20; j++) {
          const value = row[j];
          if (value !== null && value !== undefined && value !== '') {
            samples.push(String(value));
          } else {
            samples.push('');
          }
        }

        // registeredAt 날짜 변환
        let registeredAtStr = '';
        if (row[21] instanceof Date) {
          registeredAtStr = Utilities.formatDate(row[21], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        } else if (row[21]) {
          registeredAtStr = String(row[21]);
        }

        results.push({
          id: String(row[1] || ''),
          date: rowDateStr,
          companyName: String(row[3] || ''),
          tmNo: String(row[4] || ''),
          productName: String(row[5] || ''),
          inspectionItem: String(row[6] || ''),
          measurementMethod: String(row[7] || ''),
          lowerLimit: String(row[8] || ''),
          upperLimit: String(row[9] || ''),
          samples: samples,
          passFailResult: String(row[20] || ''),
          registeredAt: registeredAtStr,
          registeredBy: String(row[22] || '')
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
