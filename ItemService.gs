/**
 * ItemService.gs - ItemList 시트 관리 (업체별 동적 시트)
 */

/**
 * 업체별 ItemList 시트 가져오기 또는 생성
 * @param {string} companyName - 업체명
 * @returns {Object} {success, sheet, message}
 */
function getOrCreateItemListSheet(companyName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getItemListSheetName(companyName);
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

      // '검사기준서' 열이 없는지 확인 (5개 열만 있는 이전 버전 시트 체크)
      if (headerRow.length === 5 && headerRow[4] === '검사형태') {
        Logger.log(`${sheetName}: 이전 버전 헤더 감지, '검사기준서' 열 추가`);
        sheet.getRange(1, 6).setValue('검사기준서');
        sheet.getRange(1, 6).setFontWeight('bold').setBackground('#6aa84f').setFontColor('#ffffff');
        sheet.getRange('F:F').setNumberFormat('@STRING@');
      }
    }

    return {
      success: true,
      sheet: sheet
    };

  } catch (error) {
    logError('getOrCreateItemListSheet', error);
    return {
      success: false,
      message: '시트 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * TM-NO 자동완성 검색
 * @param {string} token - 세션 토큰
 * @param {string} searchText - 검색어
 */
function searchTmNo(token, searchText) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        items: []
      };
    }

    const searchLower = (searchText || '').toString().toLowerCase();
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
      const sheetResult = getOrCreateItemListSheet(companyName);
      if (!sheetResult.success) {
        continue;
      }

      const sheet = sheetResult.sheet;
      const data = sheet.getDataRange().getDisplayValues();

      if (data.length <= 1) {
        continue;
      }

      // 헤더 제외하고 검색
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tmNo = String(row[1] || '');
        const productName = String(row[2] || '');
        const rowCompanyName = String(row[3] || '');

        // TM-NO에 검색어 포함 여부 확인
        if (tmNo.toLowerCase().includes(searchLower)) {
          results.push({
            tmNo: tmNo,
            productName: productName,
            companyName: rowCompanyName,
            inspectionType: String(row[4] || ''),
            inspectionStandardUrl: String(row[5] || '')
          });
        }

        // 최대 10개까지만
        if (results.length >= 10) {
          break;
        }
      }

      if (results.length >= 10) {
        break;
      }
    }

    return {
      success: true,
      items: results
    };

  } catch (error) {
    Logger.log('searchTmNo 오류: ' + error.toString());
    return {
      success: false,
      message: 'TM-NO 검색 중 오류가 발생했습니다.',
      items: []
    };
  }
}

/**
 * TM-NO로 제품 정보 가져오기
 * @param {string} token - 세션 토큰
 * @param {string} tmNo - TM-NO
 */
function getItemByTmNo(token, tmNo) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 검색할 업체 목록 결정
    let companiesToSearch = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToSearch = getAllCompanyNames();
    } else {
      companiesToSearch = [session.companyName];
    }

    // 각 업체 시트에서 검색
    for (const companyName of companiesToSearch) {
      const sheetResult = getOrCreateItemListSheet(companyName);
      if (!sheetResult.success) {
        continue;
      }

      const sheet = sheetResult.sheet;
      const data = sheet.getDataRange().getDisplayValues();

      // 헤더 제외하고 검색
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (String(row[1]) === String(tmNo)) {
          return {
            success: true,
            item: {
              tmNo: String(row[1] || ''),
              productName: String(row[2] || ''),
              companyName: String(row[3] || ''),
              inspectionType: String(row[4] || ''),
              inspectionStandardUrl: String(row[5] || '')
            }
          };
        }
      }
    }

    return {
      success: false,
      message: '해당 TM-NO를 찾을 수 없습니다.'
    };

  } catch (error) {
    Logger.log('getItemByTmNo 오류: ' + error.toString());
    return {
      success: false,
      message: '제품 정보 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * ItemList에 새 항목 추가
 * @param {string} token - 세션 토큰
 * @param {Object} itemData - {tmNo, productName, companyName, inspectionType}
 */
function addItem(token, itemData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 일반 사용자는 자기 업체 항목만 추가 가능
    if (session.role !== '관리자' && session.role !== 'JEO' && itemData.companyName !== session.companyName) {
      return {
        success: false,
        message: '다른 업체의 항목을 추가할 권한이 없습니다.'
      };
    }

    // 해당 업체 시트 가져오기
    const sheetResult = getOrCreateItemListSheet(itemData.companyName);
    if (!sheetResult.success) {
      return {
        success: false,
        message: sheetResult.message || 'ItemList 시트를 찾을 수 없습니다.'
      };
    }

    const itemSheet = sheetResult.sheet;

    // 업체코드 조회
    const companyCode = findCompanyCodeByName(itemData.companyName);
    if (!companyCode) {
      return {
        success: false,
        message: '업체코드를 찾을 수 없습니다.'
      };
    }

    // 중복 체크 (컬럼 인덱스 수정: TM-NO는 두번째 열(index 1))
    const data = itemSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(itemData.tmNo)) {
        return {
          success: false,
          message: '이미 등록된 TM-NO입니다.'
        };
      }
    }

    // 새 행 추가 (업체CODE 포함, 검사기준서 URL 포함)
    const lastRow = itemSheet.getLastRow() + 1;

    // 먼저 텍스트 형식으로 설정 (특히 TM-NO 컬럼)
    itemSheet.getRange(lastRow, 1, 1, 6).setNumberFormat('@STRING@');

    // 데이터 입력
    itemSheet.getRange(lastRow, 1, 1, 6).setValues([[
      companyCode,
      itemData.tmNo,
      itemData.productName,
      itemData.companyName,
      itemData.inspectionType || '',
      itemData.inspectionStandardUrl || ''
    ]]);

    return {
      success: true,
      message: 'TM-NO가 등록되었습니다.'
    };

  } catch (error) {
    Logger.log('addItem 오류: ' + error.toString());
    return {
      success: false,
      message: '항목 추가 중 오류가 발생했습니다.'
    };
  }
}

/**
 * Item 목록 조회
 * @param {string} token - 세션 토큰
 * @param {Object} options - {companyName, tmNo}
 */
function getItems(token, options) {
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
      const sheetResult = getOrCreateItemListSheet(companyName);
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
        if (options && options.tmNo) {
          const tmNo = String(row[1] || '');
          if (tmNo.toLowerCase().indexOf(options.tmNo.toLowerCase()) === -1) {
            continue;
          }
        }

        results.push({
          rowIndex: i + 1,
          sheetName: sheet.getName(),
          tmNo: String(row[1] || ''),
          productName: String(row[2] || ''),
          companyName: rowCompanyName,
          inspectionType: String(row[4] || ''),
          inspectionStandardUrl: String(row[5] || '')
        });
      }
    }

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('getItems 오류: ' + error.toString());
    return {
      success: false,
      message: 'Item 목록 조회 중 오류가 발생했습니다.',
      data: []
    };
  }
}

/**
 * Item 수정
 * @param {string} token - 세션 토큰
 * @param {string} sheetName - 시트 이름
 * @param {number} rowIndex - 행 번호
 * @param {Object} itemData - {tmNo, productName, companyName, inspectionType}
 */
function updateItem(token, sheetName, rowIndex, itemData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemSheet = ss.getSheetByName(sheetName);

    if (!itemSheet) {
      return {
        success: false,
        message: 'ItemList 시트를 찾을 수 없습니다.'
      };
    }

    // 업체코드 조회
    const companyCode = findCompanyCodeByName(itemData.companyName);
    if (!companyCode) {
      return {
        success: false,
        message: '업체코드를 찾을 수 없습니다.'
      };
    }

    // 기존 데이터 확인 (최소 5개 컬럼, 검사기준서 포함 시 6개)
    const lastCol = Math.max(itemSheet.getLastColumn(), 6);
    const existingData = itemSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const existingCompany = String(existingData[3]);

    // 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO' && existingCompany !== session.companyName) {
      return {
        success: false,
        message: '해당 항목을 수정할 권한이 없습니다.'
      };
    }

    // 데이터 수정 (업체CODE 포함, 검사기준서 URL 포함)
    const range = itemSheet.getRange(rowIndex, 1, 1, 6);

    // 먼저 텍스트 형식으로 설정
    range.setNumberFormat('@STRING@');

    // 데이터 입력
    range.setValues([[
      companyCode,
      itemData.tmNo,
      itemData.productName,
      itemData.companyName,
      itemData.inspectionType || '',
      itemData.inspectionStandardUrl || ''
    ]]);

    return {
      success: true,
      message: 'Item이 수정되었습니다.'
    };

  } catch (error) {
    Logger.log('updateItem 오류: ' + error.toString());
    return {
      success: false,
      message: 'Item 수정 중 오류가 발생했습니다.'
    };
  }
}

/**
 * Item 삭제
 * @param {string} token - 세션 토큰
 * @param {string} sheetName - 시트 이름
 * @param {number} rowIndex - 행 번호
 */
function deleteItem(token, sheetName, rowIndex) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemSheet = ss.getSheetByName(sheetName);

    if (!itemSheet) {
      return {
        success: false,
        message: 'ItemList 시트를 찾을 수 없습니다.'
      };
    }

    // 기존 데이터 확인 (최소 5개 컬럼, 검사기준서 포함 시 6개)
    const lastCol = Math.max(itemSheet.getLastColumn(), 6);
    const existingData = itemSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const existingCompany = String(existingData[3]);

    // 권한 체크
    if (session.role !== '관리자' && session.role !== 'JEO' && existingCompany !== session.companyName) {
      return {
        success: false,
        message: '해당 항목을 삭제할 권한이 없습니다.'
      };
    }

    // 행 삭제
    itemSheet.deleteRow(rowIndex);

    return {
      success: true,
      message: 'Item이 삭제되었습니다.'
    };

  } catch (error) {
    Logger.log('deleteItem 오류: ' + error.toString());
    return {
      success: false,
      message: 'Item 삭제 중 오류가 발생했습니다.'
    };
  }
}
