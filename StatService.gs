/**
 * StatService.gs - 통계 데이터 처리 (토큰 기반)
 */

/**
 * 대시보드 통계 데이터 가져오기 (토큰 기반)
 */
function getDashboardStats(token) {
  try {
    // 토큰 검증
    const session = getSessionByToken(token);

    if (!session || !session.userId) {
      return {
        success: false,
        message: '로그인이 필요합니다.',
        stats: {
          totalCount: 0,
          todayCount: 0,
          thisMonthCount: 0,
          recentData: []
        }
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    
    if (!dataSheet) {
      return {
        success: true,
        stats: {
          totalCount: 0,
          todayCount: 0,
          thisMonthCount: 0,
          recentData: []
        }
      };
    }
    
    const allData = dataSheet.getDataRange().getDisplayValues();
    
    if (allData.length <= 1) {
      return {
        success: true,
        stats: {
          totalCount: 0,
          todayCount: 0,
          thisMonthCount: 0,
          recentData: []
        }
      };
    }
    
    const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    const now = new Date();
    const thisMonthStart = Utilities.formatDate(
      new Date(now.getFullYear(), now.getMonth(), 1),
      'Asia/Seoul',
      'yyyy-MM-dd'
    );
    
    let totalCount = 0;
    let todayCount = 0;
    let thisMonthCount = 0;
    let recentData = [];
    
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      
      if (!row[0] || !row[1]) continue;
      
      const companyName = String(row[1]);
      let dateValue = row[2];
      
      if (dateValue instanceof Date) {
        dateValue = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
      } else if (typeof dateValue === 'string') {
        dateValue = dateValue.trim();
      } else {
        continue;
      }
      
      if (session.role !== '관리자' && session.role !== 'JEO' && companyName !== session.companyName) {
        continue;
      }
      
      totalCount++;
      
      if (dateValue === today) {
        todayCount++;
      }
      
      if (dateValue >= thisMonthStart) {
        thisMonthCount++;
      }
      
      // 최근 데이터 - 단순한 문자열/숫자만 저장
      recentData.push({
        id: String(row[0] || ''),
        companyName: companyName,
        date: String(dateValue),
        time: String(row[3] || ''),
        tmNo: String(row[4] || ''),
        productName: String(row[5] || ''),
        quantity: Number(row[6]) || 0
      });
    }
    
    // 정렬
    recentData.sort(function(a, b) {
      if (a.date === b.date) {
        if (a.time === b.time) return 0;
        if (a.time === '오후' && b.time === '오전') return -1;
        return 1;
      }
      return b.date > a.date ? 1 : -1;
    });
    
    recentData = recentData.slice(0, 5);
    
    return {
      success: true,
      stats: {
        totalCount: totalCount,
        todayCount: todayCount,
        thisMonthCount: thisMonthCount,
        recentData: recentData
      }
    };
    
  } catch (error) {
    Logger.log('getDashboardStats 오류: ' + error.toString());
    
    return {
      success: true,
      stats: {
        totalCount: 0,
        todayCount: 0,
        thisMonthCount: 0,
        recentData: []
      }
    };
  }
}
