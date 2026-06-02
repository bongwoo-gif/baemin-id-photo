function appendNewRowsFromExternal() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName("RAW");
  
    // 현재 RAW 시트의 A열 key 목록
    const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 1).getValues(); 
    const rawKeys = new Set(rawData.flat().filter(v => v)); // Set으로 빠르게 검색
  
    Logger.log("현재 RAW 시트 key 수: %s", rawKeys.size);
  
    // 외부 시트 로드
    const externalUrl = "https://docs.google.com/spreadsheets/d/1RmqRL7yCWmlnuNUYqRSbXlv_MH7MZ0epmDygiI65wo0/edit";
    const extSS = SpreadsheetApp.openByUrl(externalUrl);
    const extSheet = extSS.getSheetByName("[Master] 최종리스트업_2,255명_0919");
  
    // 외부 B~I열 로드
    const extRange = extSheet.getRange(2, 2, extSheet.getLastRow() - 1, 8).getValues(); 
    // B~I = 8개 (index 0~7), I index는 7
  
    Logger.log("외부 데이터 로드 완료: %s rows", extRange.length);
  
    let insert = [];
  
    extRange.forEach(row => {
      const key = row[0]; // B열
      const flag = row[7]; // I열 (true/false)
  
      if (flag !== true) return; // I열 TRUE만
  
      if (rawKeys.has(key)) {
        Logger.log("[SKIP] key=%s 이미 존재함", key);
        return;
      }
  
      // 외부 B~H = row[0]~row[6] → RAW A~G 로 삽입
      const newRow = row.slice(0, 7);
  
      insert.push(newRow);
  
      Logger.log("[ADD] key=%s 추가 예정", key);
    });
  
    // 신규 데이터가 있는 경우 RAW 하단에 추가
    if (insert.length > 0) {
      const startRow = rawSheet.getLastRow() + 1;
      rawSheet.getRange(startRow, 1, insert.length, 7).setValues(insert);
      Logger.log("총 %s건 추가 완료. startRow=%s", insert.length, startRow);
    } else {
      Logger.log("추가할 데이터 없음.");
    }
  }
  