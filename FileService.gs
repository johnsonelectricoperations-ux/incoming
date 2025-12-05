/**
 * FileService.gs - PDF 파일 업로드 및 관리
 */

/**
 * PDF 파일 업로드
 * @param {string} token - 세션 토큰
 * @param {Object} fileData - { fileName, mimeType, bytes, tmNo, date }
 */
function uploadPdfFile(token, fileData) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    // 파일 크기 체크 (10MB 제한)
    const maxSize = 10 * 1024 * 1024; // 10MB
    if (fileData.bytes.length > maxSize) {
      return {
        success: false,
        message: '파일 크기는 10MB를 초과할 수 없습니다.'
      };
    }

    // MIME 타입 체크
    if (fileData.mimeType !== 'application/pdf') {
      return {
        success: false,
        message: 'PDF 파일만 업로드 가능합니다.'
      };
    }

    // 업체별 월별 폴더 가져오기 (년도/월 폴더 포함)
    const monthlyFolder = getOrCreateMonthlyFolder(session.companyName, fileData.date);
    
    // 파일명 생성 (날짜_TM-NO_원본파일명)
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
    const safeFileName = fileData.fileName.replace(/[^a-zA-Z0-9._-]/g, '_');
    const newFileName = `${timestamp}_${fileData.tmNo}_${safeFileName}`;

    // Blob 생성
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.bytes),
      fileData.mimeType,
      newFileName
    );

    // 파일 업로드
    const file = monthlyFolder.createFile(blob);
    
    // 파일 공유 설정 (링크가 있는 사용자)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      message: '파일이 업로드되었습니다.',
      fileId: file.getId(),
      fileName: newFileName,
      fileUrl: file.getUrl()
    };
    
  } catch (error) {
    Logger.log('파일 업로드 오류: ' + error.toString());
    return {
      success: false,
      message: '파일 업로드 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

/**
 * 파일 URL에서 파일 ID 추출
 */
function extractFileIdFromUrl(url) {
  try {
    if (!url) return null;
    
    // Google Drive URL 패턴들
    const patterns = [
      /\/file\/d\/([a-zA-Z0-9_-]+)/,
      /id=([a-zA-Z0-9_-]+)/,
      /^([a-zA-Z0-9_-]{25,})$/
    ];
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) {
        return match[1];
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('파일 ID 추출 오류: ' + error.toString());
    return null;
  }
}

/**
 * 파일 정보 가져오기
 * @param {string} token - 세션 토큰
 */
function getFileInfo(token, fileId) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const file = DriveApp.getFileById(fileId);
    
    // 파일이 속한 폴더 확인
    const folders = file.getParents();
    let isAuthorized = false;
    
    while (folders.hasNext()) {
      const folder = folders.next();
      // 관리자/JEO는 모든 파일 접근 가능
      if (session.role === '관리자' || session.role === 'JEO') {
        isAuthorized = true;
        break;
      }
      // 자신의 업체 폴더인지 확인
      if (folder.getName() === session.companyName) {
        isAuthorized = true;
        break;
      }
    }
    
    if (!isAuthorized) {
      return { success: false, message: '파일 접근 권한이 없습니다.' };
    }
    
    return {
      success: true,
      file: {
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        size: file.getSize(),
        mimeType: file.getMimeType(),
        createdDate: file.getDateCreated()
      }
    };
    
  } catch (error) {
    Logger.log('파일 정보 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '파일 정보를 가져올 수 없습니다.'
    };
  }
}

/**
 * 파일 삭제
 * @param {string} token - 세션 토큰
 */
function deleteFile(token, fileId) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const file = DriveApp.getFileById(fileId);
    
    // 파일이 속한 폴더 확인
    const folders = file.getParents();
    let isAuthorized = false;
    
    while (folders.hasNext()) {
      const folder = folders.next();
      // 관리자/JEO는 모든 파일 삭제 가능
      if (session.role === '관리자' || session.role === 'JEO') {
        isAuthorized = true;
        break;
      }
      // 자신의 업체 폴더인지 확인
      if (folder.getName() === session.companyName) {
        isAuthorized = true;
        break;
      }
    }
    
    if (!isAuthorized) {
      return { success: false, message: '파일 삭제 권한이 없습니다.' };
    }
    
    // 휴지통으로 이동
    file.setTrashed(true);
    
    return {
      success: true,
      message: '파일이 삭제되었습니다.'
    };
    
  } catch (error) {
    Logger.log('파일 삭제 오류: ' + error.toString());
    return {
      success: false,
      message: '파일 삭제 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 파일 다운로드 URL 생성
 * @param {string} token - 세션 토큰
 */
function getFileDownloadUrl(token, fileId) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }

    const fileInfo = getFileInfo(token, fileId);
    if (!fileInfo.success) {
      return fileInfo;
    }
    
    // 다운로드 URL 생성
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
    const viewUrl = `https://drive.google.com/file/d/${fileId}/view`;
    
    return {
      success: true,
      downloadUrl: downloadUrl,
      viewUrl: viewUrl
    };
    
  } catch (error) {
    Logger.log('다운로드 URL 생성 오류: ' + error.toString());
    return {
      success: false,
      message: 'URL 생성 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 업체 폴더의 모든 파일 목록 가져오기
 * @param {string} token - 세션 토큰
 */
function listCompanyFiles(token) {
  try {
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return { success: false, message: '로그인이 필요합니다.' };
    }
    
    const companyFolder = getOrCreateCompanyFolder(session.companyName);
    const files = companyFolder.getFiles();
    
    let fileList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        size: file.getSize(),
        createdDate: file.getDateCreated()
      });
    }
    
    // 날짜순 정렬 (최신순)
    fileList.sort((a, b) => b.createdDate - a.createdDate);
    
    return {
      success: true,
      files: fileList,
      total: fileList.length
    };
    
  } catch (error) {
    Logger.log('파일 목록 조회 오류: ' + error.toString());
    return {
      success: false,
      message: '파일 목록 조회 중 오류가 발생했습니다.'
    };
  }
}

// ===== 검사기준서 업로드 관련 함수 =====

// 검사기준서 루트 폴더 이름
const INSPECTION_STANDARD_FOLDER_NAME = '검사기준서';

/**
 * 검사기준서 루트 폴더 가져오기 또는 생성
 * @returns {Folder} 검사기준서 루트 폴더
 */
function getOrCreateInspectionStandardFolder() {
  try {
    const rootFolder = DriveApp.getRootFolder();
    const folders = rootFolder.getFoldersByName(INSPECTION_STANDARD_FOLDER_NAME);

    if (folders.hasNext()) {
      return folders.next();
    } else {
      // 폴더가 없으면 생성
      const newFolder = rootFolder.createFolder(INSPECTION_STANDARD_FOLDER_NAME);
      Logger.log('검사기준서 루트 폴더 생성: ' + newFolder.getId());
      return newFolder;
    }
  } catch (error) {
    Logger.log('검사기준서 루트 폴더 조회/생성 오류: ' + error.toString());
    throw error;
  }
}

/**
 * 업체별 검사기준서 폴더 가져오기 또는 생성
 * @param {string} companyName - 업체명
 * @returns {Folder} 업체별 검사기준서 폴더
 */
function getOrCreateCompanyInspectionStandardFolder(companyName) {
  try {
    const rootFolder = getOrCreateInspectionStandardFolder();
    const folders = rootFolder.getFoldersByName(companyName);

    if (folders.hasNext()) {
      return folders.next();
    } else {
      // 폴더가 없으면 생성
      const newFolder = rootFolder.createFolder(companyName);
      Logger.log(`${companyName} 검사기준서 폴더 생성: ${newFolder.getId()}`);
      return newFolder;
    }
  } catch (error) {
    Logger.log(`${companyName} 검사기준서 폴더 조회/생성 오류: ${error.toString()}`);
    throw error;
  }
}

/**
 * 검사기준서 파일 업로드
 * @param {string} token - 세션 토큰
 * @param {Object} fileData - { fileName, mimeType, bytes, companyName, tmNo }
 * @returns {Object} {success, fileUrl, message}
 */
function uploadInspectionStandardFile(token, fileData) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 권한 체크: 관리자/JEO 또는 해당 업체 사용자만
    if (session.role !== '관리자' && session.role !== 'JEO' && session.companyName !== fileData.companyName) {
      return {
        success: false,
        message: '파일 업로드 권한이 없습니다.'
      };
    }

    // 파일 크기 체크 (20MB 제한)
    const maxSize = 20 * 1024 * 1024; // 20MB
    if (fileData.bytes.length > maxSize) {
      return {
        success: false,
        message: '파일 크기는 20MB를 초과할 수 없습니다.'
      };
    }

    // 업체별 폴더 가져오기
    const companyFolder = getOrCreateCompanyInspectionStandardFolder(fileData.companyName);

    // 기존 파일 삭제 (같은 TM-NO로 시작하는 파일)
    const files = companyFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      if (fileName.startsWith(`${fileData.tmNo}_검사기준서`)) {
        file.setTrashed(true);
        Logger.log(`기존 검사기준서 삭제: ${fileName}`);
      }
    }

    // 파일 이름에서 확장자 추출
    const originalName = fileData.fileName;
    const extension = originalName.substring(originalName.lastIndexOf('.'));
    const newFileName = `${fileData.tmNo}_검사기준서${extension}`;

    // Blob 생성
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.bytes),
      fileData.mimeType,
      newFileName
    );

    // 새 파일 업로드
    const file = companyFolder.createFile(blob);

    // 파일 공유 설정 (링크가 있는 사용자는 볼 수 있음)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = file.getUrl();
    Logger.log(`검사기준서 업로드 완료: ${newFileName}, URL: ${fileUrl}`);

    return {
      success: true,
      fileUrl: fileUrl,
      fileName: newFileName,
      message: '검사기준서가 업로드되었습니다.'
    };

  } catch (error) {
    Logger.log('검사기준서 업로드 오류: ' + error.toString());
    return {
      success: false,
      message: '파일 업로드 중 오류가 발생했습니다: ' + error.message
    };
  }
}

/**
 * 검사기준서 파일 삭제
 * @param {string} token - 세션 토큰
 * @param {string} fileUrl - 삭제할 파일 URL
 * @param {string} companyName - 업체명
 * @returns {Object} {success, message}
 */
function deleteInspectionStandardFile(token, fileUrl, companyName) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);
    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.'
      };
    }

    // 권한 체크: 관리자/JEO 또는 해당 업체 사용자만
    if (session.role !== '관리자' && session.role !== 'JEO' && session.companyName !== companyName) {
      return {
        success: false,
        message: '파일 삭제 권한이 없습니다.'
      };
    }

    if (!fileUrl) {
      return {
        success: false,
        message: '삭제할 파일이 없습니다.'
      };
    }

    // URL에서 파일 ID 추출
    const fileId = extractFileIdFromUrl(fileUrl);
    if (!fileId) {
      return {
        success: false,
        message: '파일 ID를 찾을 수 없습니다.'
      };
    }

    // 파일 삭제
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    Logger.log(`검사기준서 삭제 완료: ${fileId}`);

    return {
      success: true,
      message: '검사기준서가 삭제되었습니다.'
    };

  } catch (error) {
    Logger.log('검사기준서 삭제 오류: ' + error.toString());
    return {
      success: false,
      message: '파일 삭제 중 오류가 발생했습니다: ' + error.message
    };
  }
}
