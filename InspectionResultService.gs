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

    // Data 시트에서 입고 정보 조회
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const dataValues = dataSheet.getDataRange().getValues();

    let dataInfo = null;
    for (let i = 1; i < dataValues.length; i++) {
      if (String(dataValues[i][0]) === String(dataId)) {
        dataInfo = {
          date: dataValues[i][2],
          companyName: dataValues[i][1],
          tmNo: dataValues[i][4],
          productName: dataValues[i][5]
        };
        break;
      }
    }

    if (!dataInfo) {
      return { success: false, message: '입고 데이터를 찾을 수 없습니다.' };
    }

    // 일반 사용자는 자기 업체 데이터만 저장 가능
    if (session.role !== '관리자' && dataInfo.companyName !== session.companyName) {
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

    // Data 시트에서 입고 정보 조회
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const dataValues = dataSheet.getDataRange().getValues();

    let dataInfo = null;
    for (let i = 1; i < dataValues.length; i++) {
      if (String(dataValues[i][0]) === String(dataId)) {
        dataInfo = {
          date: dataValues[i][2],
          companyName: dataValues[i][1],
          tmNo: dataValues[i][4],
          productName: dataValues[i][5]
        };
        break;
      }
    }

    if (!dataInfo) {
      return { success: false, message: '입고 데이터를 찾을 수 없습니다.', data: [] };
    }

    // 일반 사용자는 자기 업체 데이터만 조회 가능
    if (session.role !== '관리자' && dataInfo.companyName !== session.companyName) {
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

    // dataId 대신 date|companyName|tmNo 조합으로 검색
    const searchKey = dateStr + '|' + dataInfo.companyName + '|' + dataInfo.tmNo;
    const results = [];

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

      if (rowKey === searchKey) {
        // 시료 데이터 추출 (컬럼 인덱스 수정: 업체CODE 추가로 +1)
        const samples = [];
        for (let j = 10; j < 20; j++) {
          samples.push(row[j]);
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
          registeredAt: row[21],
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
    if (session.role === '관리자') {
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
