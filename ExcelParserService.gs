/**
 * ExcelParserService.gs - Excel 파일 파싱 및 데이터 추출
 *
 * 주요 기능:
 * 1. Excel 파일을 Google Sheets로 변환
 * 2. 유연한 헤더 인식으로 데이터 읽기
 * 3. 검사기준서(Standard) 및 검사성적서(Sample) 파싱
 */

/**
 * Excel 파일을 Google Sheets로 변환
 * @param {string} excelFileId - Drive에 저장된 Excel 파일 ID
 * @param {string} tempFolderName - 임시 변환 파일을 저장할 폴더명 (선택사항)
 * @returns {Object} {success, spreadsheetId, spreadsheet, message}
 */
function convertExcelToSheets(excelFileId, tempFolderName) {
  try {
    // Excel 파일 가져오기
    const excelFile = DriveApp.getFileById(excelFileId);

    if (!excelFile) {
      return {
        success: false,
        message: 'Excel 파일을 찾을 수 없습니다.'
      };
    }

    // Excel 파일을 Google Sheets로 변환
    const blob = excelFile.getBlob();

    // 변환할 폴더 찾기 또는 생성
    let targetFolder;
    if (tempFolderName) {
      const folders = DriveApp.getFoldersByName(tempFolderName);
      if (folders.hasNext()) {
        targetFolder = folders.next();
      } else {
        targetFolder = DriveApp.createFolder(tempFolderName);
      }
    } else {
      // 기본 임시 폴더
      const folders = DriveApp.getFoldersByName('JEO_TempConversions');
      if (folders.hasNext()) {
        targetFolder = folders.next();
      } else {
        targetFolder = DriveApp.createFolder('JEO_TempConversions');
      }
    }

    // Google Sheets로 변환
    const resource = {
      title: excelFile.getName() + '_converted',
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: targetFolder.getId() }]
    };

    const convertedFile = Drive.Files.insert(resource, blob, {
      convert: true
    });

    const spreadsheetId = convertedFile.id;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    Logger.log(`Excel 변환 완료: ${excelFile.getName()} → Spreadsheet ID: ${spreadsheetId}`);

    return {
      success: true,
      spreadsheetId: spreadsheetId,
      spreadsheet: spreadsheet,
      message: 'Excel 파일이 성공적으로 변환되었습니다.'
    };

  } catch (error) {
    Logger.log('convertExcelToSheets 오류: ' + error.toString());
    return {
      success: false,
      message: 'Excel 변환 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * Spreadsheet에서 데이터 읽기 (첫 번째 시트)
 * @param {string} spreadsheetId - Spreadsheet ID
 * @returns {Object} {success, data, headers, message}
 */
function readSpreadsheetData(spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheets()[0]; // 첫 번째 시트

    if (!sheet) {
      return {
        success: false,
        message: 'Spreadsheet에 시트가 없습니다.'
      };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      return {
        success: false,
        message: '시트에 데이터가 없습니다.'
      };
    }

    // 전체 데이터 읽기
    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    Logger.log(`데이터 읽기 완료: ${lastRow}행 x ${lastCol}열`);

    return {
      success: true,
      data: allData,
      sheetName: sheet.getName(),
      rowCount: lastRow,
      colCount: lastCol,
      message: '데이터를 성공적으로 읽었습니다.'
    };

  } catch (error) {
    Logger.log('readSpreadsheetData 오류: ' + error.toString());
    return {
      success: false,
      message: '데이터 읽기 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 헤더 행 찾기 (유연한 인식)
 * @param {Array<Array>} data - 2차원 배열 데이터
 * @param {Array<string>} expectedHeaders - 예상되는 헤더 키워드 목록 (예: ['항목', '규격', 'MIN', 'MAX'])
 * @returns {Object} {success, headerRowIndex, headerMap, message}
 */
function findHeaderRow(data, expectedHeaders) {
  try {
    // 처음 20행 내에서 헤더 검색 (대부분의 경우 헤더는 상단에 있음)
    const searchLimit = Math.min(20, data.length);

    for (let i = 0; i < searchLimit; i++) {
      const row = data[i];
      const foundHeaders = {};
      let matchCount = 0;

      // 각 셀을 검사하여 예상 헤더와 매칭
      for (let j = 0; j < row.length; j++) {
        const cellValue = String(row[j]).trim().toLowerCase();

        if (!cellValue) continue;

        // 예상 헤더 키워드와 비교
        for (const expectedHeader of expectedHeaders) {
          const expectedLower = expectedHeader.toLowerCase();

          // 부분 일치 또는 완전 일치 확인
          if (cellValue.includes(expectedLower) || expectedLower.includes(cellValue)) {
            foundHeaders[expectedHeader] = j; // 컬럼 인덱스 저장
            matchCount++;
            break; // 첫 번째 매칭만 사용
          }
        }
      }

      // 최소 2개 이상의 헤더가 매칭되면 헤더 행으로 간주
      if (matchCount >= 2) {
        Logger.log(`헤더 행 발견: ${i + 1}행, 매칭된 헤더: ${matchCount}개`);
        return {
          success: true,
          headerRowIndex: i,
          headerMap: foundHeaders, // {헤더명: 컬럼인덱스}
          matchCount: matchCount,
          message: '헤더 행을 찾았습니다.'
        };
      }
    }

    // 헤더를 찾지 못한 경우
    return {
      success: false,
      message: '헤더 행을 찾을 수 없습니다. 예상 헤더: ' + expectedHeaders.join(', ')
    };

  } catch (error) {
    Logger.log('findHeaderRow 오류: ' + error.toString());
    return {
      success: false,
      message: '헤더 행 검색 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 검사기준서(Standard) Excel 파싱
 * @param {string} excelFileId - Excel 파일 ID
 * @returns {Object} {success, items, message}
 */
function parseStandardExcel(excelFileId) {
  try {
    // Step 1: Excel을 Google Sheets로 변환
    const convertResult = convertExcelToSheets(excelFileId, 'JEO_TempConversions');

    if (!convertResult.success) {
      return convertResult;
    }

    const spreadsheetId = convertResult.spreadsheetId;

    // Step 2: 데이터 읽기
    const readResult = readSpreadsheetData(spreadsheetId);

    if (!readResult.success) {
      // 변환된 임시 파일 삭제
      DriveApp.getFileById(spreadsheetId).setTrashed(true);
      return readResult;
    }

    const allData = readResult.data;

    // Step 3: 헤더 행 찾기
    const expectedHeaders = ['항목', '규격', 'MIN', 'MAX', '최소', '최대'];
    const headerResult = findHeaderRow(allData, expectedHeaders);

    if (!headerResult.success) {
      // 변환된 임시 파일 삭제
      DriveApp.getFileById(spreadsheetId).setTrashed(true);
      return headerResult;
    }

    const headerRowIndex = headerResult.headerRowIndex;
    const headerMap = headerResult.headerMap;

    // Step 4: 데이터 추출
    const items = [];

    // 헤더 다음 행부터 데이터 읽기
    for (let i = headerRowIndex + 1; i < allData.length; i++) {
      const row = allData[i];

      // 항목명 추출
      let itemName = '';
      if (headerMap['항목'] !== undefined) {
        itemName = String(row[headerMap['항목']] || '').trim();
      }

      // 항목명이 비어있으면 건너뛰기
      if (!itemName) continue;

      // 규격 추출
      let spec = '';
      if (headerMap['규격'] !== undefined) {
        spec = String(row[headerMap['규격']] || '').trim();
      }

      // MIN 값 추출
      let minValue = null;
      if (headerMap['MIN'] !== undefined) {
        const minCell = row[headerMap['MIN']];
        if (minCell !== '' && minCell !== null && minCell !== undefined) {
          minValue = parseFloat(minCell);
        }
      } else if (headerMap['최소'] !== undefined) {
        const minCell = row[headerMap['최소']];
        if (minCell !== '' && minCell !== null && minCell !== undefined) {
          minValue = parseFloat(minCell);
        }
      }

      // MAX 값 추출
      let maxValue = null;
      if (headerMap['MAX'] !== undefined) {
        const maxCell = row[headerMap['MAX']];
        if (maxCell !== '' && maxCell !== null && maxCell !== undefined) {
          maxValue = parseFloat(maxCell);
        }
      } else if (headerMap['최대'] !== undefined) {
        const maxCell = row[headerMap['최대']];
        if (maxCell !== '' && maxCell !== null && maxCell !== undefined) {
          maxValue = parseFloat(maxCell);
        }
      }

      items.push({
        itemName: itemName,
        spec: spec,
        minValue: minValue,
        maxValue: maxValue
      });
    }

    Logger.log(`검사기준서 파싱 완료: ${items.length}개 항목`);

    // 변환된 임시 파일 삭제
    DriveApp.getFileById(spreadsheetId).setTrashed(true);

    return {
      success: true,
      items: items,
      itemCount: items.length,
      message: `${items.length}개의 검사 항목을 추출했습니다.`
    };

  } catch (error) {
    Logger.log('parseStandardExcel 오류: ' + error.toString());
    return {
      success: false,
      message: '검사기준서 파싱 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 검사성적서(Sample) Excel 파싱
 * @param {string} excelFileId - Excel 파일 ID
 * @returns {Object} {success, samples, message}
 */
function parseSampleExcel(excelFileId) {
  try {
    // Step 1: Excel을 Google Sheets로 변환
    const convertResult = convertExcelToSheets(excelFileId, 'JEO_TempConversions');

    if (!convertResult.success) {
      return convertResult;
    }

    const spreadsheetId = convertResult.spreadsheetId;

    // Step 2: 데이터 읽기
    const readResult = readSpreadsheetData(spreadsheetId);

    if (!readResult.success) {
      // 변환된 임시 파일 삭제
      DriveApp.getFileById(spreadsheetId).setTrashed(true);
      return readResult;
    }

    const allData = readResult.data;

    // Step 3: 헤더 행 찾기
    const expectedHeaders = ['항목', '측정값', '결과', '측정', '값'];
    const headerResult = findHeaderRow(allData, expectedHeaders);

    if (!headerResult.success) {
      // 변환된 임시 파일 삭제
      DriveApp.getFileById(spreadsheetId).setTrashed(true);
      return headerResult;
    }

    const headerRowIndex = headerResult.headerRowIndex;
    const headerMap = headerResult.headerMap;

    // Step 4: 데이터 추출
    const samples = [];

    // 헤더 다음 행부터 데이터 읽기
    for (let i = headerRowIndex + 1; i < allData.length; i++) {
      const row = allData[i];

      // 항목명 추출
      let itemName = '';
      if (headerMap['항목'] !== undefined) {
        itemName = String(row[headerMap['항목']] || '').trim();
      }

      // 항목명이 비어있으면 건너뛰기
      if (!itemName) continue;

      // 측정값 추출
      let measuredValue = null;
      if (headerMap['측정값'] !== undefined) {
        const valueCell = row[headerMap['측정값']];
        if (valueCell !== '' && valueCell !== null && valueCell !== undefined) {
          measuredValue = parseFloat(valueCell);
        }
      } else if (headerMap['측정'] !== undefined) {
        const valueCell = row[headerMap['측정']];
        if (valueCell !== '' && valueCell !== null && valueCell !== undefined) {
          measuredValue = parseFloat(valueCell);
        }
      } else if (headerMap['값'] !== undefined) {
        const valueCell = row[headerMap['값']];
        if (valueCell !== '' && valueCell !== null && valueCell !== undefined) {
          measuredValue = parseFloat(valueCell);
        }
      }

      // 측정값이 있는 경우만 추가
      if (measuredValue !== null && !isNaN(measuredValue)) {
        samples.push({
          itemName: itemName,
          measuredValue: measuredValue
        });
      }
    }

    Logger.log(`검사성적서 파싱 완료: ${samples.length}개 샘플`);

    // 변환된 임시 파일 삭제
    DriveApp.getFileById(spreadsheetId).setTrashed(true);

    return {
      success: true,
      samples: samples,
      sampleCount: samples.length,
      message: `${samples.length}개의 샘플 측정값을 추출했습니다.`
    };

  } catch (error) {
    Logger.log('parseSampleExcel 오류: ' + error.toString());
    return {
      success: false,
      message: '검사성적서 파싱 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}
