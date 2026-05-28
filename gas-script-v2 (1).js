// ============================================================
// HINEMOS Shift App v2 — Google Apps Script
// Layout: 行 = 日付、列 = スタッフ名
// 左端: 日付 | 曜日 | スタッフ1 | スタッフ2 | ...
// ============================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const eventName = data.eventName || 'Responses';
    const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    let sheet = ss.getSheetByName(eventName);

    // ── シートが存在しない場合は新規作成 ──────────────────
    if (!sheet) {
      sheet = ss.insertSheet(eventName);

      // 1行目: ヘッダー（Date | Day | スタッフ名...）
      const staffNames = data.responses ? [] : [];
      const headerRow = ['Date', 'Day', data.name];
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

      // 2行目以降: 日付ごとに1行
      const dateRows = data.responses.map(r => {
        const d = new Date(r.date);
        const dow = DAYS[d.getDay()];
        const val = formatVal(r);
        return [r.date, dow, val];
      });
      if (dateRows.length > 0) {
        sheet.getRange(2, 1, dateRows.length, 3).setValues(dateRows);
      }

      applyHeaderStyle(sheet, 1, headerRow.length);
      applyDateStyles(sheet, data.responses, 3);
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(2);
      sheet.autoResizeColumns(1, headerRow.length);

    } else {
      // ── シートが既に存在する場合 ──────────────────────────

      // スタッフ名の列を探す or 追加する
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      let staffCol = headerRow.indexOf(data.name) + 1;

      if (staffCol === 0) {
        // 新しいスタッフ → 新しい列を追加
        staffCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, staffCol).setValue(data.name).setFontWeight('bold');
        applyHeaderStyle(sheet, 1, staffCol);
      }

      // 各日付の行にデータを書き込む
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
      data.responses.forEach(r => {
        // 日付が既に存在する行を探す
        let dateRow = -1;
        for (let i = 1; i < allData.length; i++) {
          if (allData[i][0] === r.date) { dateRow = i + 1; break; }
        }

        // 存在しない場合は新しい行を追加
        if (dateRow === -1) {
          dateRow = sheet.getLastRow() + 1;
          const d = new Date(r.date);
          const dow = DAYS[d.getDay()];
          sheet.getRange(dateRow, 1).setValue(r.date);
          sheet.getRange(dateRow, 2).setValue(dow);
          allData.push([r.date]); // ローカルキャッシュ更新
        }

        // スタッフの回答を書き込む
        const cell = sheet.getRange(dateRow, staffCol);
        const val  = formatVal(r);
        cell.setValue(val);

        // 色付け
        colorCell(cell, r);
      });

      sheet.autoResizeColumns(1, sheet.getLastColumn());
    }

    // 曜日列の色付け（毎回）
    styleDayColumn(sheet);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', eventName, name: data.name }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── ヘルパー関数 ──────────────────────────────────────────

function formatVal(r) {
  if (!r || r.type === 'off' || !r.type) return 'Off';
  return r.start + '-' + r.end;
}

function colorCell(cell, r) {
  if (!r || r.type === 'off' || !r.type) {
    cell.setBackground('#f4c7c3').setFontColor('#c0392b');
  } else if (r.type === 'full') {
    cell.setBackground('#b7e1cd').setFontColor('#1e6b44');
  } else {
    cell.setBackground('#fce8b2').setFontColor('#7d4c00');
  }
}

function applyHeaderStyle(sheet, row, lastCol) {
  const range = sheet.getRange(row, 1, 1, lastCol);
  range.setBackground('#2d4a2d')
       .setFontColor('#ffffff')
       .setFontWeight('bold');
}

function applyDateStyles(sheet, responses, staffCol) {
  responses.forEach((r, i) => {
    const rowNum = i + 2;
    const cell = sheet.getRange(rowNum, staffCol);
    colorCell(cell, r);
  });
}

function styleDayColumn(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  for (let row = 2; row <= lastRow; row++) {
    const dow = sheet.getRange(row, 2).getValue();
    if (dow === 'Sat') {
      sheet.getRange(row, 2).setBackground('#ebf3fc').setFontColor('#185fa5').setFontWeight('bold');
    } else if (dow === 'Sun') {
      sheet.getRange(row, 2).setBackground('#fcebeb').setFontColor('#a32d2d').setFontWeight('bold');
    } else {
      sheet.getRange(row, 2).setBackground('#f5f5f3').setFontColor('#555').setFontWeight('normal');
    }
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'HINEMOS Shift v2 API running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
