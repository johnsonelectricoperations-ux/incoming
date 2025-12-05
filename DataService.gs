/**
 * DataService.gs - 데이터 CRUD 처리 (토큰 기반)
 */

/**
 * 업체별 Data 시트 가져오기 또는 생성
 * @param {string} companyName - 업체명
 * @returns {Object} {success, sheet, message}
 */
function getOrCreateDataSheet(companyName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getDataSheetName(companyName);
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // JEO본사(관리자/JEO 권한)는 시트 생성 생략
      if (companyName === 'JEO본사') {
        return {
          success: false,
          sheet: null,
          message: '관리자/JEO 권한은 별도 시트가 필요하지 않습니다.'
        };
      }

      // 시트가 없으면 생성
      const result = createCompanySheets(companyName);
      if (!result.success) {
        return {
          success: false,
          sheet: null,
          message: `${companyName}의 Data 시트를 생성할 수 없습니다.`
        };
      }
      sheet = ss.getSheetByName(sheetName);
    }

    return {
      success: true,
      sheet: sheet,
      message: 'Data 시트를 찾았습니다.'
    };
  } catch (error) {
    logError('getOrCreateDataSheet', error);
    return {
      success: false,
      sheet: null,
      message: 'Data 시트 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 데이터 생성 (성적서 입력) - 업체별 Data 시트에 저장
 */
function createData(token, dataObj) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 업체별 Data 시트 가져오기 또는 생성
    const sheetResult = getOrCreateDataSheet(session.companyName);
    if (!sheetResult.success) {
      return { success: false, message: sheetResult.message };
    }

    const dataSheet = sheetResult.sheet;
    const timestamp = new Date();
    const id = 'D' + timestamp.getTime();

    // TM-NO를 명시적으로 문자열로 변환
    const tmNo = String(dataObj.tmNo);

    // 업체코드 조회
    const companyCode = findCompanyCodeByName(session.companyName);
    if (!companyCode) {
      return { success: false, message: '업체 코드를 찾을 수 없습니다.' };
    }

    // 데이터 추가
    const lastRow = dataSheet.getLastRow() + 1;

    // 먼저 텍스트 형식으로 설정 (날짜 자동 변환 방지)
    // companyCode, ID, companyName, time, tmNo, productName, pdfUrl, createdBy를 텍스트로 설정
    dataSheet.getRange(lastRow, 1).setNumberFormat('@STRING@'); // companyCode
    dataSheet.getRange(lastRow, 2).setNumberFormat('@STRING@'); // ID
    dataSheet.getRange(lastRow, 3).setNumberFormat('@STRING@'); // companyName
    dataSheet.getRange(lastRow, 5).setNumberFormat('@STRING@'); // time
    dataSheet.getRange(lastRow, 6).setNumberFormat('@STRING@'); // tmNo
    dataSheet.getRange(lastRow, 7).setNumberFormat('@STRING@'); // productName
    dataSheet.getRange(lastRow, 9).setNumberFormat('@STRING@'); // pdfUrl
    dataSheet.getRange(lastRow, 11).setNumberFormat('@STRING@'); // createdBy

    // 데이터 입력 (12컬럼: 업체CODE, ID, 업체명, 입고날짜, 입고시간, TM-NO, 제품명, 수량, PDF_URL, 등록일시, 등록자, 수정일시)
    dataSheet.getRange(lastRow, 1, 1, 12).setValues([[
      companyCode,
      id,
      session.companyName,
      dataObj.date,
      dataObj.time,
      tmNo,
      dataObj.productName,
      dataObj.quantity,
      dataObj.pdfUrl || '',
      timestamp,
      session.name,
      ''
    ]]);
    
    Logger.log('데이터 생성 완료: ' + id);
    
    return {
      success: true,
      message: '성적서가 등록되었습니다.',
      id: id
    };
    
  } catch (error) {
    Logger.log('데이터 생성 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 저장 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 데이터 조회 (업체별 필터링) - 업체별 Data 시트에서 조회
 */
function getData(token, options) {
  const startTime = new Date().getTime();
  Logger.log('=== getData 시작 ===');

  try {
    // options가 없으면 빈 객체로
    if (!options) {
      options = {};
    }

    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      Logger.log('getData: 세션 없음 또는 로그인 필요');
      return {
        success: false,
        message: '로그인이 필요합니다.',
        data: [],
        total: 0,
        page: 1,
        pageSize: 20,
        totalPages: 0
      };
    }

    Logger.log(`getData: 사용자=${session.name}, 권한=${session.role}, 업체=${session.companyName}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      const companyStartTime = new Date().getTime();
      companiesToQuery = getAllCompanyNames();
      Logger.log(`getData: 업체 목록 조회 완료 (${companiesToQuery.length}개 업체, ${new Date().getTime() - companyStartTime}ms)`);
    } else {
      companiesToQuery = [session.companyName];
      Logger.log(`getData: 일반 사용자 - 업체=${session.companyName}`);
    }

    // 각 업체 ItemList 시트에서 검사형태 조회 (시트가 있는 경우만)
    const itemListStartTime = new Date().getTime();
    const itemInspectionTypeMap = {};
    for (const companyName of companiesToQuery) {
      const sheetName = getItemListSheetName(companyName);
      const sheet = ss.getSheetByName(sheetName);

      // 시트가 없으면 스킵 (생성하지 않음)
      if (!sheet) {
        Logger.log(`${companyName}의 ItemList 시트가 없어서 기본 검사형태 사용`);
        continue;
      }

      try {
        const itemListData = sheet.getDataRange().getDisplayValues();

        if (itemListData.length > 1) {
          for (let i = 1; i < itemListData.length; i++) {
            const tmNo = String(itemListData[i][0] || '');
            const productName = String(itemListData[i][1] || '');
            const rowCompanyName = String(itemListData[i][2] || '');
            const inspectionType = String(itemListData[i][3] || '');
            const key = tmNo + '|' + rowCompanyName;
            itemInspectionTypeMap[key] = inspectionType;
          }
        }
      } catch (e) {
        Logger.log(`${companyName} ItemList 조회 오류: ${e.message}`);
        continue;
      }
    }
    Logger.log(`getData: ItemList 조회 완료 (${new Date().getTime() - itemListStartTime}ms)`);

    const dataQueryStartTime = new Date().getTime();
    let results = [];

    // 각 업체별 Data 시트에서 데이터 조회
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`${companyName}의 Data 시트(${dataSheetName})를 찾을 수 없습니다. 스킵.`);
        continue;
      }

      try {
        const allData = dataSheet.getDataRange().getDisplayValues();

        if (allData.length <= 1) {
          continue; // 헤더만 있거나 데이터 없음
        }

      // 헤더 제외하고 데이터 처리
      for (let i = 1; i < allData.length; i++) {
        const row = allData[i];

        if (!row[1] || !row[2]) continue; // ID와 업체명 체크

        let dateValue = row[3];
        if (dateValue instanceof Date) {
          dateValue = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
        } else if (typeof dateValue === 'string') {
          dateValue = dateValue.trim();
        }

        const tmNo = String(row[5] || '');
        const itemKey = tmNo + '|' + companyName;
        const inspectionType = itemInspectionTypeMap[itemKey] || '검사';

        const rowData = {
          rowIndex: i + 1,
          companyCode: String(row[0] || ''),
          id: String(row[1] || ''),
          companyName: String(row[2] || ''),
          date: String(dateValue || ''),
          time: String(row[4] || ''),
          tmNo: tmNo,
          productName: String(row[6] || ''),
          quantity: Number(row[7]) || 0,
          pdfUrl: String(row[8] || ''),
          createdAt: row[9] ? Utilities.formatDate(new Date(row[9]), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') : '',
          createdBy: String(row[10] || ''),
          updatedAt: row[11] ? Utilities.formatDate(new Date(row[11]), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') : '',
          inspectionType: inspectionType,
          dataSheetName: dataSheetName // 어느 시트에서 왔는지 추적
        };
      
      // 필터 적용
      // 기존 searchText (tmNo, productName 검색)
      if (options.searchText) {
        const searchLower = options.searchText.toLowerCase();
        if (
          !rowData.tmNo.toLowerCase().includes(searchLower) &&
          !rowData.productName.toLowerCase().includes(searchLower)
        ) {
          continue;
        }
      }

      // 개별 검색 필터
      if (options.searchCompany) {
        const searchLower = options.searchCompany.toLowerCase();
        if (!rowData.companyName.toLowerCase().includes(searchLower)) {
          continue;
        }
      }

      if (options.searchDate) {
        // 날짜 검색 (YYYY-MM-DD 형식)
        if (!rowData.date.includes(options.searchDate)) {
          continue;
        }
      }

      if (options.searchTmNo) {
        const searchLower = options.searchTmNo.toLowerCase();
        if (!rowData.tmNo.toLowerCase().includes(searchLower)) {
          continue;
        }
      }

      if (options.dateFrom && rowData.date < options.dateFrom) {
        continue;
      }

      if (options.dateTo && rowData.date > options.dateTo) {
        continue;
      }

        results.push(rowData);
      } // for loop (데이터 행)
      } catch (e) {
        Logger.log(`${companyName} Data 시트 조회 오류: ${e.message}`);
        continue;
      }
    } // for loop (업체별 시트)
    Logger.log(`getData: Data 시트 조회 완료 (총 ${results.length}건, ${new Date().getTime() - dataQueryStartTime}ms)`);

    // 정렬 (날짜 기준 내림차순)
    const sortStartTime = new Date().getTime();
    results.sort(function(a, b) {
      if (a.date === b.date) {
        if (a.time === b.time) return 0;
        if (a.time === '오후' && b.time === '오전') return -1;
        return 1;
      }
      return b.date > a.date ? 1 : -1;
    });
    
    Logger.log(`getData: 정렬 완료 (${new Date().getTime() - sortStartTime}ms)`);

    // 페이징
    const page = Number(options.page) || 1;
    const pageSize = Number(options.pageSize) || 20;
    const startIndex = (page - 1) * pageSize;
    const endIndex = startIndex + pageSize;

    const totalTime = new Date().getTime() - startTime;
    Logger.log(`=== getData 완료 (총 ${totalTime}ms) - 결과: ${results.length}건 중 ${Math.min(pageSize, results.length - startIndex)}건 반환 ===`);

    return {
      success: true,
      data: results.slice(startIndex, endIndex),
      total: results.length,
      page: page,
      pageSize: pageSize,
      totalPages: Math.ceil(results.length / pageSize)
    };

  } catch (error) {
    const totalTime = new Date().getTime() - startTime;
    Logger.log(`getData 오류 (${totalTime}ms): ${error.toString()}`);
    Logger.log('Stack trace: ' + error.stack);
    return {
      success: false,
      message: '데이터 조회 중 오류가 발생했습니다: ' + error.message,
      data: [],
      total: 0,
      page: 1,
      pageSize: 20,
      totalPages: 0
    };
  }
}

/**
 * 특정 데이터 조회 (ID로) - 업체별 Data 시트에서 조회
 */
function getDataById(token, id) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
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
        Logger.log(`getDataById: ${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const data = dataSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (String(row[1]) === String(id)) { // row[1]이 ID
          // 권한 체크
          if (session.role !== '관리자' && session.role !== 'JEO' && String(row[2]) !== session.companyName) {
            return { success: false, message: '접근 권한이 없습니다.' };
          }

          let dateValue = row[3];
          if (dateValue instanceof Date) {
            dateValue = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
          }

          return {
            success: true,
            data: {
              rowIndex: i + 1,
              companyCode: String(row[0]),
              id: String(row[1]),
              companyName: String(row[2]),
              date: String(dateValue),
              time: String(row[4]),
              tmNo: String(row[5]),
              productName: String(row[6]),
              quantity: Number(row[7]),
              pdfUrl: String(row[8] || ''),
              createdAt: String(row[9] || ''),
              createdBy: String(row[10] || ''),
              updatedAt: String(row[11] || ''),
              dataSheetName: dataSheetName
            }
          };
        }
      }
      } catch (e) {
        Logger.log(`getDataById: ${companyName} Data 시트 조회 오류 - ${e.message}`);
        continue;
      }
    }

    return { success: false, message: '데이터를 찾을 수 없습니다.' };

  } catch (error) {
    Logger.log('데이터 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 데이터 수정 - 업체별 Data 시트에서 수정
 */
function updateData(token, id, dataObj) {
  try {
    const session = getSessionByToken(token);
    if (!session) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    // 각 업체별 Data 시트에서 ID 검색 및 수정
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`updateData: ${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const data = dataSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (String(row[1]) === String(id)) { // row[1]이 ID
          // 권한 체크
          if (session.role !== '관리자' && session.role !== 'JEO' && String(row[2]) !== session.companyName) {
            return { success: false, message: '수정 권한이 없습니다.' };
          }

          const rowIndex = i + 1;
          const timestamp = new Date();

          // 먼저 텍스트 필드들을 텍스트 형식으로 설정 (날짜 자동 변환 방지)
          dataSheet.getRange(rowIndex, 5).setNumberFormat('@STRING@'); // time
          dataSheet.getRange(rowIndex, 6).setNumberFormat('@STRING@'); // tmNo
          dataSheet.getRange(rowIndex, 7).setNumberFormat('@STRING@'); // productName
          if (dataObj.pdfUrl) {
            dataSheet.getRange(rowIndex, 9).setNumberFormat('@STRING@'); // pdfUrl
          }

          // 데이터 업데이트 (업체CODE, ID는 변경하지 않음)
          dataSheet.getRange(rowIndex, 4).setValue(dataObj.date);
          dataSheet.getRange(rowIndex, 5).setValue(dataObj.time);
          dataSheet.getRange(rowIndex, 6).setValue(String(dataObj.tmNo));
          dataSheet.getRange(rowIndex, 7).setValue(dataObj.productName);
          dataSheet.getRange(rowIndex, 8).setValue(dataObj.quantity);

          if (dataObj.pdfUrl) {
            dataSheet.getRange(rowIndex, 9).setValue(dataObj.pdfUrl);
          }
          dataSheet.getRange(rowIndex, 12).setValue(timestamp);

          return {
            success: true,
            message: '성적서가 수정되었습니다.'
          };
        }
      }
      } catch (e) {
        Logger.log(`updateData: ${companyName} Data 시트 수정 오류 - ${e.message}`);
        continue;
      }
    }

    return { success: false, message: '데이터를 찾을 수 없습니다.' };

  } catch (error) {
    Logger.log('데이터 수정 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 수정 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 데이터 삭제 - 업체별 Data 시트에서 삭제
 */
function deleteData(token, id) {
  try {
    const session = getSessionByToken(token);
    if (!session) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    // 각 업체별 Data 시트에서 ID 검색 및 삭제
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`deleteData: ${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const data = dataSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (String(row[1]) === String(id)) { // row[1]이 ID
          // 권한 체크
          if (session.role !== '관리자' && session.role !== 'JEO' && String(row[2]) !== session.companyName) {
            return { success: false, message: '삭제 권한이 없습니다.' };
          }

          const rowIndex = i + 1;
          dataSheet.deleteRow(rowIndex);

          // PDF 파일 삭제 (옵션)
          if (row[8]) { // row[8]이 PDF URL
            try {
              const fileId = extractFileIdFromUrl(row[8]);
              if (fileId) {
                DriveApp.getFileById(fileId).setTrashed(true);
              }
            } catch (e) {
              Logger.log('PDF 파일 삭제 오류: ' + e.toString());
            }
          }

          return {
            success: true,
            message: '성적서가 삭제되었습니다.'
          };
        }
      }
      } catch (e) {
        Logger.log(`deleteData: ${companyName} Data 시트 삭제 오류 - ${e.message}`);
        continue;
      }
    }

    return { success: false, message: '데이터를 찾을 수 없습니다.' };

  } catch (error) {
    Logger.log('데이터 삭제 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 삭제 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 날짜별 데이터 조회 (확인증용) - 업체별 Data 시트에서 조회
 */
function getDataByDate(token, date) {
  try {
    const session = getSessionByToken(token);
    if (!session) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    let results = [];

    // 각 업체별 Data 시트에서 데이터 조회
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`getDataByDate: ${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const data = dataSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row[1] || !row[2]) continue; // ID와 업체명 체크

          let rowDate = row[3];
          if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, 'Asia/Seoul', 'yyyy-MM-dd');
          } else if (typeof rowDate === 'string') {
            rowDate = rowDate.trim();
          }

          if (rowDate === date) {
            results.push({
              companyCode: String(row[0] || ''),
              id: String(row[1]),
              companyName: String(row[2]),
              date: String(rowDate),
              time: String(row[4] || ''),
              tmNo: String(row[5] || ''),
              productName: String(row[6] || ''),
              quantity: Number(row[7]) || 0,
              pdfUrl: String(row[8] || '')
            });
          }
        }
      } catch (e) {
        Logger.log(`getDataByDate: ${companyName} Data 시트 조회 오류 - ${e.message}`);
        continue;
      }
    }

    // 시간순 정렬
    results.sort((a, b) => a.time.localeCompare(b.time));

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('날짜별 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 날짜 및 시간별 데이터 조회 (확인증용) - 업체별 Data 시트에서 조회
 */
function getDataByDateAndTime(token, date, time) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자' || session.role === 'JEO') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    let results = [];

    // 각 업체별 Data 시트에서 데이터 조회
    for (const companyName of companiesToQuery) {
      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`getDataByDateAndTime: ${companyName}의 Data 시트를 찾을 수 없음`);
        continue;
      }

      try {
        const data = dataSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row[1] || !row[2]) continue; // ID와 업체명 체크

          let rowDate = row[3];
          const rowTime = String(row[4] || '');

          // 날짜 형식 통일
          if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, 'Asia/Seoul', 'yyyy-MM-dd');
          } else if (typeof rowDate === 'string') {
            rowDate = rowDate.trim();
          }

          // 날짜 필터
          if (rowDate !== date) {
            continue;
          }

          // 시간 필터 (time이 빈 문자열이면 전체)
          if (time && rowTime !== time) {
            continue;
          }

          results.push({
            companyCode: String(row[0] || ''),
            id: String(row[1]),
            companyName: String(row[2]),
            date: String(rowDate),
            time: rowTime,
            tmNo: String(row[5] || ''),
            productName: String(row[6] || ''),
            quantity: Number(row[7]) || 0,
            pdfUrl: String(row[8] || '')
          });
        }
      } catch (e) {
        Logger.log(`getDataByDateAndTime: ${companyName} Data 시트 조회 오류 - ${e.message}`);
        continue;
      }
    }

    // 시간순 정렬 (오전 → 오후)
    results.sort(function(a, b) {
      if (a.time === b.time) return 0;
      if (a.time === '오전' && b.time === '오후') return -1;
      return 1;
    });

    return {
      success: true,
      data: results
    };

  } catch (error) {
    Logger.log('날짜/시간별 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 조회 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 검사결과등록을 위한 Data 검색
 * 최적화: ItemList 데이터를 사전에 캐싱하여 중복 조회 제거
 * @param {string} token - 세션 토큰
 * @param {Object} filters - 검색 필터 {companyName, date, time, tmNo}
 * @returns {Object} {success, data: [{id, companyName, date, time, tmNo, productName}]}
 */
function searchDataForInspectionResult(token, filters) {
  const startTime = new Date().getTime();

  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        data: []
      };
    }

    // 권한 체크: 관리자 또는 JEO만 접근 가능
    if (session.role !== '관리자' && session.role !== 'JEO') {
      return {
        success: false,
        message: '검사결과등록 권한이 없습니다.',
        data: []
      };
    }

    Logger.log('searchDataForInspectionResult: 검색 시작 - 사용자=' + session.name);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];

    // 일반 권한 업체 목록 조회 (관리자 제외)
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet) {
      return {
        success: false,
        message: '사용자 시트를 찾을 수 없습니다.',
        data: []
      };
    }

    const userData = userSheet.getDataRange().getDisplayValues();
    const normalCompanies = [];

    for (let i = 1; i < userData.length; i++) {
      const role = String(userData[i][3] || '');
      const companyName = String(userData[i][1] || '');

      if (role !== '관리자' && role !== 'JEO' && companyName && !normalCompanies.includes(companyName)) {
        normalCompanies.push(companyName);
      }
    }

    Logger.log(`검색 대상 업체: ${normalCompanies.length}개`);

    // 검색 필터 파싱
    const filterCompanyName = filters.companyName || '';
    const filterDate = filters.date || '';
    const filterTime = filters.time || '';
    const filterTmNo = filters.tmNo || '';

    // 최적화: ItemList 데이터를 사전에 모두 로드하여 Map으로 캐싱
    const itemListStartTime = new Date().getTime();
    const itemInspectionTypeMap = {};
    for (const companyName of normalCompanies) {
      // 업체명 필터가 있으면 해당 업체만 로드
      if (filterCompanyName && filterCompanyName !== companyName) {
        continue;
      }

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

    // 각 업체의 Data 시트에서 검색
    const dataStartTime = new Date().getTime();
    for (const companyName of normalCompanies) {
      // 업체명 필터 적용
      if (filterCompanyName && filterCompanyName !== companyName) {
        continue;
      }

      const dataSheetName = getDataSheetName(companyName);
      const dataSheet = ss.getSheetByName(dataSheetName);

      if (!dataSheet) {
        Logger.log(`${companyName}의 Data 시트가 없습니다.`);
        continue;
      }

      try {
        const dataValues = dataSheet.getDataRange().getDisplayValues();

        // 헤더 제외하고 검색
        for (let i = 1; i < dataValues.length; i++) {
          const row = dataValues[i];

          const id = String(row[1] || '');
          const rowCompanyName = String(row[2] || '');
          const tmNo = String(row[5] || '');
          const productName = String(row[6] || '');
          const quantity = Number(row[7]) || 0;

          // 날짜 형식 정규화
          let rowDate = row[3];
          if (row[3] instanceof Date) {
            rowDate = Utilities.formatDate(row[3], 'Asia/Seoul', 'yyyy-MM-dd');
          } else if (rowDate) {
            rowDate = String(rowDate).trim();
          }

          const rowTime = String(row[4] || '');

          // 날짜 필터
          if (filterDate && rowDate !== filterDate) {
            continue;
          }

          // 시간 필터 (부분 일치)
          if (filterTime && rowTime.indexOf(filterTime) === -1) {
            continue;
          }

          // TM-NO 필터 (부분 일치)
          if (filterTmNo && tmNo.indexOf(filterTmNo) === -1) {
            continue;
          }

          // ItemList에서 검사형태 조회 (캐시 사용)
          const itemKey = rowCompanyName + '|' + tmNo;
          const inspectionType = itemInspectionTypeMap[itemKey] || '검사';

          // 무검사 항목은 제외 (검사결과등록 대상이 아님)
          if (inspectionType === '무검사') {
            continue;
          }

          // 조건에 맞으면 결과에 추가
          results.push({
            id: id,
            companyName: rowCompanyName,
            date: rowDate,
            time: rowTime,
            tmNo: tmNo,
            productName: productName,
            quantity: quantity,
            inspectionType: inspectionType
          });
        }
      } catch (e) {
        Logger.log(`${companyName} Data 시트 조회 오류: ${e.message}`);
        continue;
      }
    }
    Logger.log(`Data 조회 완료 (${new Date().getTime() - dataStartTime}ms)`);

    const totalTime = new Date().getTime() - startTime;
    Logger.log(`검사결과등록용 검색 완료: ${results.length}건 (${totalTime}ms)`);

    return {
      success: true,
      data: results
    };

  } catch (error) {
    const totalTime = new Date().getTime() - startTime;
    Logger.log(`searchDataForInspectionResult 오류 (${totalTime}ms): ${error.toString()}`);
    return {
      success: false,
      message: '검색 중 오류가 발생했습니다: ' + error.message,
      data: []
    };
  }
}
