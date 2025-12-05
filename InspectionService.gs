/**
 * InspectionService.gs - InspectionSpec 시트 관리 (업체별 동적 시트)
 */

/**
 * 업체별 InspectionSpec 시트 가져오기 또는 생성
 * @param {string} companyName - 업체명
 * @returns {Object} {success, sheet, message}
 */
function getOrCreateInspectionSpecSheet(companyName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getInspectionSpecSheetName(companyName);
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
    logError('getOrCreateInspectionSpecSheet', error);
    return {
      success: false,
      message: '시트 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 검사 규격 조회
 * @param {string} token - 세션 토큰
 * @param {Object} options - {tmNo, companyName}
 */
function getInspectionSpecs(token, options) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        data: []
      };
    }

    const results = [];

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (options && options.companyName) {
      // 특정 업체 지정된 경우
      companiesToQuery = [options.companyName];
    } else if (session.role === '관리자' || session.role === 'JEO') {
      // 관리자는 모든 업체
      companiesToQuery = getAllCompanyNames();
    } else {
      // 일반 사용자는 자기 업체만
      companiesToQuery = [session.companyName];
    }

    // 각 업체 시트에서 데이터 조회
    for (const companyName of companiesToQuery) {
      const sheetResult = getOrCreateInspectionSpecSheet(companyName);
      if (!sheetResult.success) {
        continue;
      }

      const sheet = sheetResult.sheet;
      const data = sheet.getDataRange().getDisplayValues();

      if (data.length <= 1) {
        continue;
      }

      // 헤더 제외하고 처리
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowCompanyName = String(row[3] || '');

        // TM-NO 필터 적용
        if (options && options.tmNo && String(row[1]) !== options.tmNo) {
          continue;
        }

        results.push({
          rowIndex: i + 1,
          sheetName: sheet.getName(),
          tmNo: String(row[1] || ''),
          productName: String(row[2] || ''),
          companyName: rowCompanyName,
          inspectionItem: String(row[4] || ''),
          measurementMethod: String(row[5] || ''),
          lowerLimit: String(row[6] || ''),
          upperLimit: String(row[7] || ''),
          sampleSize: String(row[8] || '')
        });
      }
    }

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('getInspectionSpecs 오류: ' + error.toString());
    return {
      success: false,
      message: '검사 규격 조회 중 오류가 발생했습니다.',
      data: []
    };
  }
}

/**
 * 검사 규격 배치 추가 (여러 개 동시 추가)
 * @param {string} token - 세션 토큰
 * @param {Array} specsArray - 검사 규격 데이터 배열
 */
function addInspectionSpecsBatch(token, specsArray) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    if (!specsArray || specsArray.length === 0) {
      return {
        success: false,
        message: '추가할 데이터가 없습니다.'
      };
    }

    // 첫 번째 항목의 업체명으로 시트 결정 (모든 항목이 같은 업체라고 가정)
    const companyName = specsArray[0].companyName;

    // 일반 사용자는 자기 업체 항목만 추가 가능
    if (session.role !== '관리자' && session.role !== 'JEO' && companyName !== session.companyName) {
      return {
        success: false,
        message: '다른 업체의 항목을 추가할 권한이 없습니다.'
      };
    }

    // 해당 업체 시트 가져오기
    const sheetResult = getOrCreateInspectionSpecSheet(companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || 'InspectionSpec 시트를 찾을 수 없습니다.'
      };
    }

    const inspectionSheet = sheetResult.sheet;

    // 업체코드 조회
    const companyCode = findCompanyCodeByName(companyName);
    if (!companyCode) {
      return {
        success: false,
        message: '업체코드를 찾을 수 없습니다.'
      };
    }

    // 배치로 추가 (업체CODE 포함)
    const rowsToAdd = specsArray.map(spec => [
      companyCode,
      spec.tmNo,
      spec.productName,
      spec.companyName,
      spec.inspectionItem,
      spec.measurementMethod || '',
      spec.lowerLimit || '',
      spec.upperLimit || '',
      spec.sampleSize || ''
    ]);

    // 시트에 추가
    const startRow = inspectionSheet.getLastRow() + 1;
    const range = inspectionSheet.getRange(startRow, 1, rowsToAdd.length, 9);

    // 먼저 텍스트 형식으로 설정 (특히 TM-NO 컬럼)
    range.setNumberFormat('@STRING@');

    // 데이터 입력
    range.setValues(rowsToAdd);

    return {
      success: true,
      message: `${specsArray.length}개의 검사 규격이 추가되었습니다.`
    };

  } catch (error) {
    Logger.log('addInspectionSpecsBatch 오류: ' + error.toString());
    return {
      success: false,
      message: '검사 규격 배치 추가 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 검사 규격 추가
 * @param {string} token - 세션 토큰
 * @param {Object} specData - 검사 규격 데이터
 */
function addInspectionSpec(token, specData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 일반 사용자는 자기 업체 항목만 추가 가능
    if (session.role !== '관리자' && session.role !== 'JEO' && specData.companyName !== session.companyName) {
      return {
        success: false,
        message: '다른 업체의 항목을 추가할 권한이 없습니다.'
      };
    }

    // 해당 업체 시트 가져오기
    const sheetResult = getOrCreateInspectionSpecSheet(specData.companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || 'InspectionSpec 시트를 찾을 수 없습니다.'
      };
    }

    const inspectionSheet = sheetResult.sheet;

    // 업체코드 조회
    const companyCode = findCompanyCodeByName(specData.companyName);
    if (!companyCode) {
      return {
        success: false,
        message: '업체코드를 찾을 수 없습니다.'
      };
    }

    // 새 행 추가 (업체CODE 포함)
    const lastRow = inspectionSheet.getLastRow() + 1;

    // 먼저 텍스트 형식으로 설정 (특히 TM-NO 컬럼)
    inspectionSheet.getRange(lastRow, 1, 1, 9).setNumberFormat('@STRING@');

    // 데이터 입력
    inspectionSheet.getRange(lastRow, 1, 1, 9).setValues([[
      companyCode,
      specData.tmNo,
      specData.productName,
      specData.companyName,
      specData.inspectionItem,
      specData.measurementMethod || '',
      specData.lowerLimit || '',
      specData.upperLimit || '',
      specData.sampleSize || ''
    ]]);

    return {
      success: true,
      message: '검사 규격이 추가되었습니다.'
    };

  } catch (error) {
    Logger.log('addInspectionSpec 오류: ' + error.toString());
    return {
      success: false,
      message: '검사 규격 추가 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 검사 규격 수정
 * @param {string} token - 세션 토큰
 * @param {string} sheetName - 시트 이름
 * @param {number} rowIndex - 행 번호
 * @param {Object} specData - 검사 규격 데이터
 */
function updateInspectionSpec(token, sheetName, rowIndex, specData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inspectionSheet = ss.getSheetByName(sheetName);

    if (!inspectionSheet) {
      return {
        success: false,
        message: 'InspectionSpec 시트를 찾을 수 없습니다.'
      };
    }

    // 업체코드 조회
    const companyCode = findCompanyCodeByName(specData.companyName);
    if (!companyCode) {
      return {
        success: false,
        message: '업체코드를 찾을 수 없습니다.'
      };
    }

    // 기존 데이터 확인 (9개 컬럼)
    const existingData = inspectionSheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
    const existingCompany = String(existingData[3]);

    // 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO' && existingCompany !== session.companyName) {
      return {
        success: false,
        message: '해당 항목을 수정할 권한이 없습니다.'
      };
    }

    // 데이터 수정 (업체CODE 포함)
    const range = inspectionSheet.getRange(rowIndex, 1, 1, 9);

    // 먼저 텍스트 형식으로 설정
    range.setNumberFormat('@STRING@');

    // 데이터 입력
    range.setValues([[
      companyCode,
      specData.tmNo,
      specData.productName,
      specData.companyName,
      specData.inspectionItem,
      specData.measurementMethod || '',
      specData.lowerLimit || '',
      specData.upperLimit || '',
      specData.sampleSize || ''
    ]]);

    return {
      success: true,
      message: '검사 규격이 수정되었습니다.'
    };

  } catch (error) {
    Logger.log('updateInspectionSpec 오류: ' + error.toString());
    return {
      success: false,
      message: '검사 규격 수정 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 검사 규격 삭제
 * @param {string} token - 세션 토큰
 * @param {string} sheetName - 시트 이름
 * @param {number} rowIndex - 행 번호
 */
function deleteInspectionSpec(token, sheetName, rowIndex) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inspectionSheet = ss.getSheetByName(sheetName);

    if (!inspectionSheet) {
      return {
        success: false,
        message: 'InspectionSpec 시트를 찾을 수 없습니다.'
      };
    }

    // 기존 데이터 확인 (9개 컬럼)
    const existingData = inspectionSheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
    const existingCompany = String(existingData[3]);

    // 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO' && existingCompany !== session.companyName) {
      return {
        success: false,
        message: '해당 항목을 삭제할 권한이 없습니다.'
      };
    }

    // 행 삭제
    inspectionSheet.deleteRow(rowIndex);

    return {
      success: true,
      message: '검사 규격이 삭제되었습니다.'
    };

  } catch (error) {
    Logger.log('deleteInspectionSpec 오류: ' + error.toString());
    return {
      success: false,
      message: '검사 규격 삭제 중 오류가 발생했습니다.'
    };
  }
}

/**
 * TM-NO와 업체명으로 검사 규격 모두 삭제
 * @param {string} token - 세션 토큰
 * @param {string} tmNo - TM-NO
 * @param {string} companyName - 업체명
 */
function deleteInspectionSpecsByTmNoAndCompany(token, tmNo, companyName) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 일반 사용자는 자기 업체 항목만 삭제 가능
    if (session.role !== '관리자' && session.role !== 'JEO' && companyName !== session.companyName) {
      return {
        success: false,
        message: '다른 업체의 항목을 삭제할 권한이 없습니다.'
      };
    }

    // 해당 업체 시트 가져오기
    const sheetResult = getOrCreateInspectionSpecSheet(companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || 'InspectionSpec 시트를 찾을 수 없습니다.'
      };
    }

    const inspectionSheet = sheetResult.sheet;
    const data = inspectionSheet.getDataRange().getDisplayValues();

    let deletedCount = 0;

    // 뒤에서부터 순회하며 삭제 (인덱스 변경 문제 방지)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowTmNo = String(row[1] || '');
      const rowCompanyName = String(row[3] || '');

      if (rowTmNo === tmNo && rowCompanyName === companyName) {
        inspectionSheet.deleteRow(i + 1);
        deletedCount++;
      }
    }

    return {
      success: true,
      message: `${deletedCount}개의 검사 규격이 삭제되었습니다.`,
      deletedCount: deletedCount
    };

  } catch (error) {
    Logger.log('deleteInspectionSpecsByTmNoAndCompany 오류: ' + error.toString());
    return {
      success: false,
      message: '검사 규격 삭제 중 오류가 발생했습니다.'
    };
  }
}

/**
 * TM-NO로 검사 규격 조회
 * @param {string} token - 세션 토큰
 * @param {string} tmNo - TM-NO
 */
function getInspectionSpecsByTmNo(token, tmNo) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        data: []
      };
    }

    const results = [];

    // 검색할 업체 목록 결정
    let companiesToSearch = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToSearch = getAllCompanyNames();
    } else {
      companiesToSearch = [session.companyName];
    }

    // 각 업체 시트에서 검색
    for (const companyName of companiesToSearch) {
      const sheetResult = getOrCreateInspectionSpecSheet(companyName);
      if (!sheetResult.success) {
        continue;
      }

      const sheet = sheetResult.sheet;
      const data = sheet.getDataRange().getDisplayValues();

      // 헤더 제외하고 검색
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (String(row[1]) === String(tmNo)) {
          results.push({
            rowIndex: i + 1,
            sheetName: sheet.getName(),
            tmNo: String(row[1] || ''),
            productName: String(row[2] || ''),
            companyName: String(row[3] || ''),
            inspectionItem: String(row[4] || ''),
            measurementMethod: String(row[5] || ''),
            lowerLimit: String(row[6] || ''),
            upperLimit: String(row[7] || ''),
            sampleSize: String(row[8] || '')
          });
        }
      }
    }

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('getInspectionSpecsByTmNo 오류: ' + error.toString());
    return {
      success: false,
      message: 'TM-NO로 검사 규격 조회 중 오류가 발생했습니다.',
      data: []
    };
  }
}

/**
 * 검사결과등록을 위한 검사규격 조회
 * @param {string} token - 세션 토큰
 * @param {string} companyName - 업체명
 * @param {string} tmNo - TM-NO
 * @returns {Object} {success, specs: [{inspectionItem, measurementMethod, lowerLimit, upperLimit, sampleSize}]}
 */
function getInspectionSpecsForResult(token, companyName, tmNo) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        specs: []
      };
    }

    // 권한 체크: 관리자 또는 JEO만 접근 가능
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return {
        success: false,
        message: '검사결과등록 권한이 없습니다.',
        specs: []
      };
    }

    Logger.log(`getInspectionSpecsForResult: 업체=${companyName}, TM-NO=${tmNo}`);

    // 해당 업체의 Spec 시트 조회
    const sheetResult = getOrCreateInspectionSpecSheet(companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || '검사규격 시트를 찾을 수 없습니다.',
        specs: []
      };
    }

    const sheet = sheetResult.sheet;
    const data = sheet.getDataRange().getDisplayValues();
    const specs = [];

    // 헤더 제외하고 검색
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // TM-NO와 업체명이 모두 일치하는 행만 조회
      if (String(row[1]) === String(tmNo) && String(row[3]) === String(companyName)) {
        specs.push({
          inspectionItem: String(row[4] || ''),
          measurementMethod: String(row[5] || ''),
          lowerLimit: String(row[6] || ''),
          upperLimit: String(row[7] || ''),
          sampleSize: parseInt(row[8]) || 1
        });
      }
    }

    if (specs.length === 0) {
      return {
        success: false,
        message: '해당 업체/TM-NO의 검사규격이 등록되지 않았습니다.',
        specs: []
      };
    }

    Logger.log(`검사규격 조회 완료: ${specs.length}개 항목`);

    return {
      success: true,
      specs: specs
    };

  } catch (error) {
    Logger.log('getInspectionSpecsForResult 오류: ' + error.toString());
    return {
      success: false,
      message: '검사규격 조회 중 오류가 발생했습니다: ' + error.message,
      specs: []
    };
  }
}
