/** 명단 시트에서 RAW 시트로 동기화 (A:G → A:G) */
function syncFromMyeongdan_appendAndUpdate() {
  const TARGET_SHEET_NAME = 'RAW';
  const SOURCE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1qZMfLaMym5dztu3kj4YSM4n78ouiQm0e-PRlugoxjso/edit';
  const SOURCE_SHEET_NAME = '명단';

  const targetSheet = SpreadsheetApp.getActive()
    .getSheetByName(TARGET_SHEET_NAME);
  const sourceSheet = SpreadsheetApp
    .openByUrl(SOURCE_SHEET_URL)
    .getSheetByName(SOURCE_SHEET_NAME);

  /** =============================
   * 1. RAW 시트 A:G 인덱싱
   ============================= */
  const targetLastRow = targetSheet.getLastRow();
  const targetValues = targetLastRow < 2
    ? []
    : targetSheet.getRange(2, 1, targetLastRow - 1, 7).getValues(); // A:G

  // key → rowIndex
  const targetMap = new Map();
  targetValues.forEach((row, i) => {
    const key = row[0];
    if (!key) return;
    targetMap.set(String(key).trim(), {
      rowIndex: i + 2
    });
  });

  /** =============================
   * 2. 명단 A:G 읽기
   ============================= */
  const sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;

  const sourceValues = sourceSheet
    .getRange(2, 1, sourceLastRow - 1, 7)
    .getValues();

  const rowsToAppend = [];
  const rowsToUpdate = [];

  /** =============================
   * 3. 비교 & 분기
   ============================= */
  sourceValues.forEach(row => {
    const key = row[0];
    if (!key) return;

    const keyStr = String(key).trim();
    const sourceF = row[5]; // ✅ 명단 F

    if (!targetMap.has(keyStr)) {
      // 신규 추가
      rowsToAppend.push(row);
    } else {
      // 기존 → F열만 업데이트
      rowsToUpdate.push({
        rowIndex: targetMap.get(keyStr).rowIndex,
        value: sourceF
      });
    }
  });

  /** =============================
   * 4. 신규 행 추가
   * A:G → A:G (그대로 복사)
   ============================= */
  if (rowsToAppend.length > 0) {
    const startRow = targetSheet.getLastRow() + 1;
    // 명단의 A:G를 그대로 사용 (별도 매핑 없이)
    const appendValues = rowsToAppend;

    // A열부터 7개 열에 쓰기
    targetSheet
      .getRange(startRow, 1, appendValues.length, 7)
      .setValues(appendValues);
  }

  /** =============================
   * 5. F열 업데이트 (기존 값)
   ============================= */
  rowsToUpdate.forEach(item => {
    targetSheet.getRange(item.rowIndex, 6).setValue(item.value);
  });
}

