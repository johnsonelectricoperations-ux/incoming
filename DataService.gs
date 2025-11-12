/**
 * DataService.gs - 데이터 CRUD 처리 (토큰 기반)
 */

/**
 * 데이터 생성 (성적서 입력)
 */
function createData(token, dataObj) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    
    if (!dataSheet) {
      return { success: false, message: 'Data 시트를 찾을 수 없습니다.' };
    }
    
    const timestamp = new Date();
    const id = 'D' + timestamp.getTime();

    // TM-NO를 명시적으로 문자열로 변환
    const tmNo = String(dataObj.tmNo);

    // 데이터 추가
    const lastRow = dataSheet.getLastRow() + 1;

    // 먼저 텍스트 형식으로 설정 (날짜 자동 변환 방지)
    // ID, companyName, time, tmNo, productName, pdfUrl, createdBy를 텍스트로 설정
    dataSheet.getRange(lastRow, 1).setNumberFormat('@STRING@'); // ID
    dataSheet.getRange(lastRow, 2).setNumberFormat('@STRING@'); // companyName
    dataSheet.getRange(lastRow, 4).setNumberFormat('@STRING@'); // time
    dataSheet.getRange(lastRow, 5).setNumberFormat('@STRING@'); // tmNo
    dataSheet.getRange(lastRow, 6).setNumberFormat('@STRING@'); // productName
    dataSheet.getRange(lastRow, 8).setNumberFormat('@STRING@'); // pdfUrl
    dataSheet.getRange(lastRow, 10).setNumberFormat('@STRING@'); // createdBy

    // 데이터 입력
    dataSheet.getRange(lastRow, 1, 1, 11).setValues([[
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
 * 데이터 조회 (업체별 필터링)
 */
function getData(token, options) {
  try {
    // options가 없으면 빈 객체로
    if (!options) {
      options = {};
    }

    const session = getSessionByToken(token);
    if (!session || !session.userId) {
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
    
    const dataSheet = getSheet(DATA_SHEET_NAME);
    if (!dataSheet) {
      return {
        success: false,
        message: 'Data 시트를 찾을 수 없습니다.',
        data: [],
        total: 0,
        page: 1,
        pageSize: 20,
        totalPages: 0
      };
    }

    const allData = dataSheet.getDataRange().getValues();

    // 조회할 업체 목록 결정
    let companiesToQuery = [];
    if (session.role === '관리자') {
      companiesToQuery = getAllCompanyNames();
    } else {
      companiesToQuery = [session.companyName];
    }

    // 각 업체 ItemList 시트에서 검사형태 조회
    const itemInspectionTypeMap = {};
    for (const companyName of companiesToQuery) {
      const sheetResult = getOrCreateItemListSheet(companyName);
      if (!sheetResult.success) {
        continue;
      }

      const itemListData = sheetResult.sheet.getDataRange().getValues();

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
    }
    
    if (allData.length <= 1) {
      return {
        success: true,
        data: [],
        total: 0,
        page: 1,
        pageSize: 20,
        totalPages: 0
      };
    }
    
    let results = [];
    
    // 헤더 제외하고 데이터 처리
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      
      if (!row[0] || !row[1]) continue;
      
      const companyName = String(row[1]);
      
      // 권한 체크
      if (session.role !== '관리자' && companyName !== session.companyName) {
        continue;
      }
      
      let dateValue = row[2];
      if (dateValue instanceof Date) {
        dateValue = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (typeof dateValue === 'string') {
        dateValue = dateValue.trim();
      }
      
      const tmNo = String(row[4] || '');
      const itemKey = tmNo + '|' + companyName;
      const inspectionType = itemInspectionTypeMap[itemKey] || '검사';

      const rowData = {
        rowIndex: i + 1,
        id: String(row[0] || ''),
        companyName: String(row[1] || ''),
        date: String(dateValue || ''),
        time: String(row[3] || ''),
        tmNo: tmNo,
        productName: String(row[5] || ''),
        quantity: Number(row[6]) || 0,
        pdfUrl: String(row[7] || ''),
        createdAt: row[8] ? Utilities.formatDate(new Date(row[8]), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') : '',
        createdBy: String(row[9] || ''),
        updatedAt: row[10] ? Utilities.formatDate(new Date(row[10]), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') : '',
        inspectionType: inspectionType
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
    }
    
    // 정렬 (날짜 기준 내림차순)
    results.sort(function(a, b) {
      if (a.date === b.date) {
        if (a.time === b.time) return 0;
        if (a.time === '오후' && b.time === '오전') return -1;
        return 1;
      }
      return b.date > a.date ? 1 : -1;
    });
    
    // 페이징
    const page = Number(options.page) || 1;
    const pageSize = Number(options.pageSize) || 20;
    const startIndex = (page - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    
    return {
      success: true,
      data: results.slice(startIndex, endIndex),
      total: results.length,
      page: page,
      pageSize: pageSize,
      totalPages: Math.ceil(results.length / pageSize)
    };
    
  } catch (error) {
    Logger.log('데이터 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 조회 중 오류가 발생했습니다.',
      data: [],
      total: 0,
      page: 1,
      pageSize: 20,
      totalPages: 0
    };
  }
}

/**
 * 특정 데이터 조회 (ID로)
 */
function getDataById(token, id) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const dataSheet = getSheet(DATA_SHEET_NAME);
    if (!dataSheet) {
      return { success: false, message: 'Data 시트를 찾을 수 없습니다.' };
    }
    
    const data = dataSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]) === String(id)) {
        // 권한 체크
        if (session.role !== '관리자' && String(row[1]) !== session.companyName) {
          return { success: false, message: '접근 권한이 없습니다.' };
        }
        
        let dateValue = row[2];
        if (dateValue instanceof Date) {
          dateValue = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
        }
        
        return {
          success: true,
          data: {
            rowIndex: i + 1,
            id: String(row[0]),
            companyName: String(row[1]),
            date: String(dateValue),
            time: String(row[3]),
            tmNo: String(row[4]),
            productName: String(row[5]),
            quantity: Number(row[6]),
            pdfUrl: String(row[7] || ''),
            createdAt: String(row[8] || ''),
            createdBy: String(row[9] || ''),
            updatedAt: String(row[10] || '')
          }
        };
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
 * 데이터 수정
 */
function updateData(token, id, dataObj) {
  try {
    const session = getSessionByToken(token);
    if (!session) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const dataSheet = getSheet(DATA_SHEET_NAME);
    const data = dataSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]) === String(id)) {
        // 권한 체크
        if (session.role !== '관리자' && String(row[1]) !== session.companyName) {
          return { success: false, message: '수정 권한이 없습니다.' };
        }
        
        const rowIndex = i + 1;
        const timestamp = new Date();

        // 먼저 텍스트 필드들을 텍스트 형식으로 설정 (날짜 자동 변환 방지)
        dataSheet.getRange(rowIndex, 4).setNumberFormat('@STRING@'); // time
        dataSheet.getRange(rowIndex, 5).setNumberFormat('@STRING@'); // tmNo
        dataSheet.getRange(rowIndex, 6).setNumberFormat('@STRING@'); // productName
        if (dataObj.pdfUrl) {
          dataSheet.getRange(rowIndex, 8).setNumberFormat('@STRING@'); // pdfUrl
        }

        // 데이터 업데이트
        dataSheet.getRange(rowIndex, 3).setValue(dataObj.date);
        dataSheet.getRange(rowIndex, 4).setValue(dataObj.time);
        dataSheet.getRange(rowIndex, 5).setValue(String(dataObj.tmNo));
        dataSheet.getRange(rowIndex, 6).setValue(dataObj.productName);
        dataSheet.getRange(rowIndex, 7).setValue(dataObj.quantity);

        if (dataObj.pdfUrl) {
          dataSheet.getRange(rowIndex, 8).setValue(dataObj.pdfUrl);
        }
        dataSheet.getRange(rowIndex, 11).setValue(timestamp);
        
        return {
          success: true,
          message: '성적서가 수정되었습니다.'
        };
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
 * 데이터 삭제
 */
function deleteData(token, id) {
  try {
    const session = getSessionByToken(token);
    if (!session) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const dataSheet = getSheet(DATA_SHEET_NAME);
    const data = dataSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === id) {
        // 권한 체크
        if (session.role !== '관리자' && row[1] !== session.companyName) {
          return { success: false, message: '삭제 권한이 없습니다.' };
        }
        
        const rowIndex = i + 1;
        dataSheet.deleteRow(rowIndex);

        // PDF 파일 삭제 (옵션)
        if (row[7]) {
          try {
            const fileId = extractFileIdFromUrl(row[7]);
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
 * 날짜별 데이터 조회 (확인증용)
 */
function getDataByDate(token, date) {
  try {
    const session = getSessionByToken(token);
    if (!session) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const dataSheet = getSheet(DATA_SHEET_NAME);
    const data = dataSheet.getDataRange().getValues();
    
    let results = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const companyName = row[1];
      const rowDate = row[2];
      
      // 권한 체크 및 날짜 필터
      if (session.role !== '관리자' && companyName !== session.companyName) {
        continue;
      }
      
      if (rowDate === date) {
        results.push({
          id: row[0],
          companyName: row[1],
          date: row[2],
          time: row[3],
          tmNo: row[4],
          productName: row[5],
          quantity: row[6],
          pdfUrl: row[7]
        });
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
 * 날짜 및 시간별 데이터 조회 (확인증용)
 */
function getDataByDateAndTime(token, date, time) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const dataSheet = getSheet(DATA_SHEET_NAME);
    if (!dataSheet) {
      return { success: false, message: 'Data 시트를 찾을 수 없습니다.' };
    }
    
    const data = dataSheet.getDataRange().getValues();
    
    let results = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || !row[1]) continue;
      
      const companyName = String(row[1]);
      let rowDate = row[2];
      const rowTime = String(row[3] || '');
      
      // 날짜 형식 통일
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (typeof rowDate === 'string') {
        rowDate = rowDate.trim();
      }
      
      // 권한 체크
      if (session.role !== '관리자' && companyName !== session.companyName) {
        continue;
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
        id: String(row[0]),
        companyName: String(row[1]),
        date: String(rowDate),
        time: rowTime,
        tmNo: String(row[4] || ''),
        productName: String(row[5] || ''),
        quantity: Number(row[6]) || 0,
        pdfUrl: String(row[7] || '')
      });
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
